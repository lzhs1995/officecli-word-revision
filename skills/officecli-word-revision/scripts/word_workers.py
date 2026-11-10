"""Explicit local worker bindings; registration never creates another Word instance."""
from __future__ import annotations

import getpass
import hashlib
import json
import os
import platform
from pathlib import Path


def local_identity():
    system = platform.system()
    if system == 'Windows':
        import ctypes
        session = ctypes.c_ulong()
        if not ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(session)):
            raise RuntimeError('Cannot establish Windows session')
        session_id = str(session.value)
    else:
        session_id = 'gui/' + str(os.getuid())
    return dict(machine=platform.node(), user=getpass.getuser(), session=session_id, platform=system)


def resolve_worker(worker_id=None, registry=None):
    current = local_identity()
    backend = 'word-com' if current['platform'] == 'Windows' else 'word-applescript'
    default = dict(worker_id='local', **current, backend=backend,
                   word_instance='Microsoft Word/current-user-session', zotero_instance=None,
                   native_lock=str(Path.home() / '.cache/wordrev/word-automation.lock'))
    if worker_id is None or worker_id == 'local':
        return default
    if worker_id not in ('mac-test', 'windows-test') or registry is None:
        raise ValueError('UNKNOWN_WORKER: an explicit test-worker registration is required')
    raw = Path(registry).read_bytes()
    data = json.loads(raw)
    if data.get('schema') != 1 or worker_id not in data.get('workers', {}):
        raise ValueError('UNKNOWN_WORKER: ' + worker_id)
    binding = data['workers'][worker_id]
    for field, actual in current.items():
        if binding.get(field) != actual:
            raise ValueError('WORKER_IDENTITY_MISMATCH: ' + field)
    expected_platform = 'Darwin' if worker_id == 'mac-test' else 'Windows'
    if current['platform'] != expected_platform:
        raise ValueError('WORKER_PLATFORM_MISMATCH')
    if binding.get('word_instance') != default['word_instance']:
        raise ValueError('Only the existing per-user Word instance is supported')
    if binding.get('zotero_instance') is not None:
        raise ValueError('ZOTERO_ROUTE_UNVERIFIED: this worker is Word-only')
    if any(k in binding for k in ('native_lock', 'lock_path', 'pending_path')):
        raise ValueError('A worker registration cannot redirect the native lock or pending state')
    root = Path(binding['output_root']).resolve(strict=True)
    return dict(default, worker_id=worker_id, output_root=str(root),
                registry=str(Path(registry).resolve()), registry_sha256=hashlib.sha256(raw).hexdigest())


def check_output(worker, output):
    if 'output_root' in worker:
        Path(output).resolve().relative_to(Path(worker['output_root']))
