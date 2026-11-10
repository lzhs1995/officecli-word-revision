"""Staged macOS Word PDF export, serialized with all wordrev operations."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time
import uuid

from word_runtime import word_lock


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def apple_string(value):
    return '"' + str(value).replace('\\', '\\\\').replace('"', '\\"') + '"'


class StageError(RuntimeError):
    def __init__(self, stage, message, uncertain=False):
        super().__init__(message)
        self.stage = stage
        self.uncertain = uncertain


class ExportSession:
    def __init__(self, source, output, *, runner=subprocess.run, worker_id=None, worker_registry=None):
        from word_workers import resolve_worker, check_output
        self.worker = resolve_worker(worker_id, worker_registry)
        self.source = Path(source).resolve(strict=True)
        self.output = Path(output).absolute()
        check_output(self.worker, self.output)
        self.runner = runner
        self.run_id = uuid.uuid4().hex
        self.evidence = self.output.parent / (self.output.stem + '.word-run-' + self.run_id[:8])
        self.evidence.mkdir(parents=True, exist_ok=False)
        # Word can show an unanswerable sandbox bookmark prompt for an arbitrary
        # output directory, even with Full Disk Access. Render inside its own
        # container; Python publishes the validated file outside that container.
        self.staging = (Path.home() / 'Library/Containers/com.microsoft.Word/Data'
                        / 'Documents/wordrev-pdf' / self.run_id)
        if self.worker['backend'] == 'word-com':
            self.staging = Path(os.environ['LOCALAPPDATA']) / 'wordrev-pdf' / self.run_id
        self.owned = self.staging / ('wordrev_' + self.run_id[:12] + '.docx')
        self.temporary = self.staging / 'export.pdf'
        self.frozen_pdf = self.evidence / 'native-output.pdf'
        self.claim = self.output.with_name('.' + self.output.name + '.wordrev-claim')
        self.claimed = False
        self.events = []
        self.result = dict(success=False, backend=self.worker['backend'], run_id=self.run_id,
                           input=str(self.source), input_sha256=sha256(self.source),
                           pdfPath=str(self.output), evidence_dir=str(self.evidence),
                           staging_dir=str(self.staging), owned_document=str(self.owned),
                           worker=self.worker, native_closed=False,
                           broker_managed=False, logical_reservation_seconds=None,
                           service_sha256=sha256(__file__),
                           runtime_sha256=sha256(Path(__file__).with_name('word_runtime.py')))
        s = self.source.stat()
        self.source_identity = (s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns)

    def record(self, stage, status, **values):
        event = dict(stage=stage, status=status, at=time.time(), **values)
        self.events.append(event)
        with (self.evidence / 'stages.jsonl').open('a') as handle:
            handle.write(json.dumps(event, ensure_ascii=False) + '\n')

    def apple(self, stage, body, seconds=20):
        script = ('with timeout of ' + str(seconds) + ' seconds\n'
                  'tell application "Microsoft Word"\n' + body + '\n'
                  'end tell\nend timeout\n')
        script_path = self.evidence / (stage.lower() + '.applescript')
        script_path.write_text(script)
        self.record(stage, 'START', budget_seconds=seconds)
        start = time.monotonic()
        try:
            proc = self.runner(['/usr/bin/osascript', str(script_path)], capture_output=True,
                               text=True, timeout=seconds + 30)
        except subprocess.TimeoutExpired as exc:
            self.record(stage, 'TIMEOUT_UNRESOLVED', duration=time.monotonic()-start)
            raise StageError(stage, str(exc), True) from exc
        self.record(stage, 'PASS' if proc.returncode == 0 else 'FAIL',
                    duration=time.monotonic()-start, returncode=proc.returncode,
                    stdout=proc.stdout, stderr=proc.stderr)
        if proc.returncode:
            raise StageError(stage, proc.stderr, '-1712' in proc.stderr)
        return proc.stdout.strip()

    def binding(self):
        return ('set docRef to document ' + apple_string(self.owned.name) + '\n'
                'if (posix full name of docRef) is not ' + apple_string(self.owned) +
                ' then error "Owned document path mismatch"\n')

    def verify(self):
        self.record('VERIFY', 'START')
        if not self.result['native_closed']:
            raise StageError('VERIFY', 'Native closure is not established')
        if sha256(self.frozen_pdf) != self.result['frozen_pdf_sha256']:
            raise StageError('VERIFY', 'Frozen PDF changed after native closure')
        if not self.frozen_pdf.is_file() or self.frozen_pdf.read_bytes()[:5] != b'%PDF-':
            raise StageError('VERIFY', 'No valid PDF header')
        info = self.runner(['pdfinfo', str(self.frozen_pdf)], capture_output=True, text=True, timeout=30)
        txt = self.runner(['pdftotext', '-layout', str(self.frozen_pdf), '-'], capture_output=True, text=True, timeout=30)
        if info.returncode or txt.returncode or not txt.stdout.strip():
            raise StageError('VERIFY', 'PDF parsing/text extraction failed')
        metadata = dict(line.split(':', 1) for line in info.stdout.splitlines() if ':' in line)
        pages = int(metadata.get('Pages', '0').strip())
        if pages < 1:
            raise StageError('VERIFY', 'PDF has no pages')
        (self.evidence / 'pdfinfo.txt').write_text(info.stdout)
        (self.evidence / 'text.txt').write_text(txt.stdout)
        self.result.update(pages=pages, bytes=self.frozen_pdf.stat().st_size,
                           pdf_sha256=sha256(self.frozen_pdf),
                           text_sha256=hashlib.sha256(txt.stdout.encode()).hexdigest())
        if sha256(self.source) != self.result['input_sha256']:
            raise StageError('VERIFY', 'Source changed during export')
        # Copy across filesystems, then link for atomic no-clobber publication.
        publishing = self.output.parent / ('.' + self.output.name + '.' + self.run_id)
        try:
            shutil.copyfile(self.frozen_pdf, publishing)
            if sha256(publishing) != self.result['pdf_sha256']:
                raise StageError('VERIFY', 'Published PDF hash mismatch')
            os.link(publishing, self.output)
        finally:
            publishing.unlink(missing_ok=True)
        self.record('VERIFY', 'PASS', pages=pages, pdf_sha256=self.result['pdf_sha256'])

    def prepare(self):
        """Private file preparation; no native OS lock or application access."""
        self.claim.parent.mkdir(parents=True, exist_ok=True)
        try:
            with self.claim.open('x') as claim:
                json.dump(dict(run_id=self.run_id, evidence=str(self.evidence), worker=self.worker), claim)
            self.claimed = True
        except FileExistsError as exc:
            raise StageError('PREPARE', 'OUTPUT_CLAIMED: ' + str(self.claim)) from exc
        self.staging.mkdir(parents=True, exist_ok=False)
        shutil.copy2(self.source, self.owned)
        if sha256(self.owned) != self.result['input_sha256'] or sha256(self.source) != self.result['input_sha256']:
            raise StageError('STAGING', 'Source/staging changed during preparation')
        self.record('STAGING', 'PASS', path=str(self.owned), sha256=self.result['input_sha256'])

    def native_phase(self):
        """Return only after owned close and a fixed native result, or retain pending."""
        if self.worker['backend'] == 'word-com':
            return self.windows_native_phase()
        wait_start = time.monotonic()
        with word_lock('word-pdf:' + self.run_id):
            acquired = time.monotonic()
            self.result['os_lock_wait_seconds'] = acquired - wait_start
            self.record('NATIVE_LOCK', 'ACQUIRED')
            attempted_open = False
            try:
                # Detect a source edit during lock waiting, before native use.
                s = self.source.stat()
                if (s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns) != self.source_identity:
                    raise StageError('PREFLIGHT', 'Source changed while waiting for Word')
                self.result['word_probe'] = self.apple('PROBE', 'return {version, name of every document}')
                processes = self.runner(['/usr/bin/pgrep', '-x', 'Microsoft Word'],
                                        capture_output=True, text=True, timeout=10)
                pids = processes.stdout.split()
                if processes.returncode or len(pids) != 1 or not pids[0].isdigit():
                    raise StageError('PREFLIGHT', 'Word instance PID is not unique/readable')
                identity = self.runner(['/bin/ps', '-p', pids[0], '-o', 'uid=,lstart=,args='],
                                       capture_output=True, text=True, timeout=10)
                if identity.returncode or not identity.stdout.strip().startswith(str(os.getuid()) + ' '):
                    raise StageError('PREFLIGHT', 'Word instance user binding failed')
                self.result['word_instance_observed'] = dict(pid=int(pids[0]), process=identity.stdout.strip())
                self.apple('PREFLIGHT', 'if exists document ' + apple_string(self.owned.name) +
                           ' then error "TARGET_ALREADY_OPEN"\nreturn "READY"')
                for name in ['~$' + self.source.name, '~$' + self.source.name[2:]]:
                    if self.source.with_name(name).exists():
                        raise StageError('PREFLIGHT', 'TARGET_LOCK_PRESENT: ' + name)
                attempted_open = True
                self.apple('OPEN', 'open (POSIX file ' + apple_string(self.owned) +
                           ') read only true add to recent files false', 180)
                self.apple('RESOLVE_PATH', self.binding() + 'return posix full name of docRef')
                self.apple('EXPORT', self.binding() +
                           'set show revisions of docRef to false\nset print revisions of docRef to false\n'
                           'set outputFile to ' + apple_string(self.temporary) + '\n'
                           'save as docRef file name outputFile file format format PDF add to recent files false', 180)
                self.apple('CLOSE', self.binding() + 'close docRef saving no\n' +
                           'if exists document ' + apple_string(self.owned.name) +
                           ' then error "Owned document still open"\nreturn "WORDREV_CLOSED"', 30)
                self.result.update(native_closed=True, cleanup='PASS')
                # Materialize the exact native bytes while still owning the native boundary.
                if not self.temporary.is_file():
                    raise StageError('FREEZE', 'Native export produced no file')
                shutil.copyfile(self.temporary, self.frozen_pdf)
                digest = sha256(self.temporary)
                if sha256(self.frozen_pdf) != digest:
                    raise StageError('FREEZE', 'Native PDF copy mismatch')
                self.result.update(frozen_pdf=str(self.frozen_pdf), frozen_pdf_sha256=digest)
            except Exception as error:
                exc = error if isinstance(error, StageError) else StageError(
                    self.events[-1]['stage'] if self.events else 'NATIVE', str(error))
                if not self.result['native_closed'] and (attempted_open or exc.uncertain):
                    from word_runtime import mark_word_pending
                    mark_word_pending(dict(run_id=self.run_id, phase=exc.stage, source=str(self.source),
                        owned_document=str(self.owned), evidence_dir=str(self.evidence), worker=self.worker,
                        status='TIMEOUT_UNRESOLVED' if exc.uncertain else 'CLEANUP_PENDING'))
                    self.result['cleanup'] = 'PENDING'
                raise exc
            finally:
                self.result['os_lock_held_seconds'] = time.monotonic() - acquired
        self.record('NATIVE_LOCK', 'RELEASED', held_seconds=self.result['os_lock_held_seconds'])
        self.result['staged_docx_unchanged'] = sha256(self.owned) == self.result['input_sha256']
        self.record('NATIVE_CLOSED_FROZEN', 'PASS', sha256=self.result['frozen_pdf_sha256'])
        (self.evidence / 'native-closed.json').write_text(json.dumps(self.result, ensure_ascii=False, indent=2))
        return dict(self.result)

    def windows_native_phase(self):
        from word_windows_native import run_native
        from word_runtime import mark_word_pending
        wait_start = time.monotonic()
        with word_lock('word-pdf:' + self.run_id):
            acquired = time.monotonic()
            self.result['os_lock_wait_seconds'] = acquired - wait_start
            self.record('NATIVE_LOCK', 'ACQUIRED')
            try:
                s = self.source.stat()
                if (s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns) != self.source_identity:
                    raise StageError('PREFLIGHT', 'Source changed while waiting for Word')
                try:
                    state = run_native(self.owned, self.temporary, self.evidence, self.worker, self.runner)
                except Exception as exc:
                    # A dispatch boundary failure cannot establish whether the
                    # native host started or closed. Preserve recovery evidence.
                    state = dict(native_started=True, native_closed=False,
                                 phase='NATIVE_DISPATCH_UNKNOWN', uncertain=True, error=str(exc))
                self.result['word_native'] = state
                if not state.get('native_closed'):
                    if state.get('native_started'):
                        self.result['cleanup'] = 'PENDING'
                        mark_word_pending(dict(run_id=self.run_id, owned_document=str(self.owned),
                            evidence_dir=str(self.evidence), worker=self.worker, native_result=state))
                    raise StageError(state.get('phase', 'COM'), state.get('error', 'Native close not confirmed'),
                                     state.get('uncertain', False))
                self.result.update(native_closed=True, cleanup='PASS')
                shutil.copyfile(self.temporary, self.frozen_pdf)
                digest = sha256(self.temporary)
                if sha256(self.frozen_pdf) != digest:
                    raise StageError('FREEZE', 'Native PDF copy mismatch')
                self.result.update(frozen_pdf=str(self.frozen_pdf), frozen_pdf_sha256=digest)
            finally:
                self.result['os_lock_held_seconds'] = time.monotonic() - acquired
        self.record('NATIVE_LOCK', 'RELEASED', held_seconds=self.result['os_lock_held_seconds'])
        self.result['staged_docx_unchanged'] = sha256(self.owned) == self.result['input_sha256']
        self.record('NATIVE_CLOSED_FROZEN', 'PASS', sha256=self.result['frozen_pdf_sha256'])
        (self.evidence / 'native-closed.json').write_text(json.dumps(self.result, ensure_ascii=False, indent=2))
        return dict(self.result)

    def export(self, *, on_native_closed=None, native_only=False):
        if self.output.exists():
            self.result.update(code='OUTPUT_EXISTS', error='Existing PDF preserved; choose a new path')
            return self.finish()
        for command in ([] if native_only else ['pdfinfo', 'pdftotext']):
            if not shutil.which(command):
                self.result.update(code='PDF_VALIDATOR_MISSING', error=command)
                return self.finish()
        try:
            self.prepare()
            native = self.native_phase()
            # A caller may reconcile its external logical ticket here. This service
            # never claims that releasing the OS lock has released that reservation.
            if on_native_closed is not None:
                on_native_closed(native)
            if native_only:
                self.result.update(code='NATIVE_CLOSED_AWAITING_QA', qa_status='NOT_RUN')
            else:
                self.verify()
                self.result.update(success=True, code='EXPORT_PASS', qa_status='PASS')
                shutil.rmtree(self.staging)
        except Exception as error:
            exc = error if isinstance(error, StageError) else StageError(
                'VERIFY' if self.result['native_closed'] else 'RUNTIME', str(error))
            code = ('NATIVE_CLOSED_QA_FAILED' if self.result['native_closed'] else
                    'TIMEOUT_UNRESOLVED' if exc.uncertain else 'STAGE_FAILED')
            self.result.update(success=False, code=code, failed_stage=exc.stage,
                               error=str(exc), permissionRequired='-1743' in str(exc))
        finally:
            # Keep claims for unresolved native operations; never unlink another run's claim.
            if self.claimed and self.result.get('cleanup') != 'PENDING' and not native_only:
                if json.loads(self.claim.read_text()).get('run_id') == self.run_id:
                    self.claim.unlink()
        return self.finish()

    def finish(self):
        self.result['input_unchanged'] = sha256(self.source) == self.result['input_sha256']
        if not self.result['input_unchanged']:
            self.result.update(success=False, code='SOURCE_CHANGED')
        (self.evidence / 'result.json').write_text(json.dumps(self.result, ensure_ascii=False, indent=2)+'\n')
        return self.result


def export_pdf(source, output, *, worker_id=None, worker_registry=None, on_native_closed=None, native_only=False):
    return ExportSession(source, output, worker_id=worker_id, worker_registry=worker_registry).export(
        on_native_closed=on_native_closed, native_only=native_only)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('source', nargs='?')
    parser.add_argument('output', nargs='?')
    parser.add_argument('--worker-id')
    parser.add_argument('--worker-registry')
    parser.add_argument('--native-only', action='store_true', help='Return closed frozen output; no QA/PASS claim')
    parser.add_argument('--ack-pending', action='store_true',
                        help='clear an interrupted run only when Word has no open documents')
    args = parser.parse_args()
    if args.ack_pending:
        from word_runtime import clear_word_pending
        try:
            print(json.dumps(dict(success=True, code='PENDING_ACKNOWLEDGED',
                                  pending=clear_word_pending()), ensure_ascii=False, indent=2))
            return 0
        except Exception as exc:
            print(json.dumps(dict(success=False, code='RECOVERY_REQUIRED', error=str(exc)),
                             ensure_ascii=False, indent=2))
            return 1
    if not args.source or not args.output:
        parser.error('source and output are required unless --ack-pending is used')
    try:
        result = export_pdf(args.source, args.output, worker_id=args.worker_id, worker_registry=args.worker_registry,
                            native_only=args.native_only)
    except Exception as exc:
        result = dict(success=False, code='INVALID_INPUT', error=str(exc))
    print(json.dumps(result, ensure_ascii=False, indent=2))
    ready = (result.get('success') or (args.native_only and result.get('native_closed') and
                                      result.get('code') == 'NATIVE_CLOSED_AWAITING_QA'))
    return 0 if ready and result.get('cleanup') == 'PASS' and result.get('input_unchanged') else 1


if __name__ == '__main__':
    raise SystemExit(main())
