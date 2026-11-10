"""Bounded Word-only COM transaction; no Zotero routing or foreign process cleanup."""
from __future__ import annotations

import json
import hashlib
from pathlib import Path
import subprocess


def ps_literal(value):
    return "'" + str(value).replace("'", "''") + "'"


def native_script(owned, output, receipt, binding):
    return r'''$ErrorActionPreference = 'Stop'
$state = @{ native_started=$false; native_closed=$false; word_pid=$null; phase='PREFLIGHT'; binding_schema=3 }
$receipt = @RECEIPT@
$expectedMachine = @MACHINE@
$expectedUser = @USER@
$expectedSession = @SESSION@
try {
  $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
  if ($env:COMPUTERNAME -ne $expectedMachine -or $env:USERNAME -ne $expectedUser -or [string]$sid -ne $expectedSession) { throw 'WORKER_IDENTITY_MISMATCH' }
  $existing = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sid })
  $state.precreate_word_count = $existing.Count
  if ($existing.Count -ne 0) { throw 'WORD_INSTANCE_BUSY: existing Word process preserved' }
  Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class WordPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }'
  $state.phase = 'CREATE'
  $createdAfter = [DateTime]::UtcNow
  $state.native_started = $true
  $app = New-Object -ComObject Word.Application
  # Word exposes Hwnd on Window, not Application. Record the unique newborn
  # process first, then bind the owned document's Window.Hwnd before export.
  $created = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sid })
  $state.postcreate_word_count = $created.Count
  if ($created.Count -ne 1) { throw 'WORD_CREATION_AMBIGUOUS' }
  $nativeProcess = $created[0]
  $state.word_pid = $nativeProcess.Id
  $wordPidValue = $nativeProcess.Id
  if ($nativeProcess.StartTime.ToUniversalTime() -lt $createdAfter.AddSeconds(-1)) { throw 'WORD_PROCESS_PREDATES_CREATE' }
  $state.word_version = [string]$app.Version
  $state.word_start_time_utc = $nativeProcess.StartTime.ToUniversalTime().ToString('o')
  $state.identity_bound = $true
  if ($null -eq $app.Visible -or $null -eq $app.UserControl -or $app.Visible -or $app.UserControl) { throw 'WORD_NOT_PRIVATE_AUTOMATION' }
  foreach ($prop in @('Documents','Windows','ProtectedViewWindows')) {
    $collection = $app.$prop
    try { if ($null -eq $collection -or $null -eq $collection.Count -or $collection.Count -ne 0) { throw ('WORD_' + $prop + '_NOT_KNOWN_EMPTY') } }
    finally { if ($null -ne $collection) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($collection) } }
  }
  $state.phase = 'OPEN'
  $doc = $app.Documents.Open(@OWNED@, $false, $true, $false)
  if ($null -eq $doc) { throw 'OWNED_DOCUMENT_UNKNOWN' }
  if ([IO.Path]::GetFullPath($doc.FullName) -ne [IO.Path]::GetFullPath(@OWNED@)) { throw 'OWNED_PATH_MISMATCH' }
  if ($app.Documents.Count -ne 1 -or $doc.Windows.Count -ne 1 -or $app.ProtectedViewWindows.Count -ne 0) { throw 'DOCUMENT_WINDOW_OR_PV_CONFLICT' }
  $state.phase = 'BIND_OWNED_WINDOW'
  $ownedWindow = $doc.Windows.Item(1)
  $rawHandle = $ownedWindow.Hwnd
  if ($null -eq $rawHandle -or [long]$rawHandle -eq 0) { throw 'OWNED_WINDOW_HWND_UNKNOWN' }
  $windowHandle = ([long]$rawHandle -band 4294967295L)
  [uint32]$windowPidValue = 0
  [void][WordPid]::GetWindowThreadProcessId([IntPtr]$windowHandle, [ref]$windowPidValue)
  $state.owned_window_hwnd = $windowHandle
  $state.owned_window_pid = $windowPidValue
  $stillOwned = Get-Process -Id $wordPidValue
  if ($windowPidValue -ne $wordPidValue -or $stillOwned.SessionId -ne $sid -or $stillOwned.StartTime.ToUniversalTime().ToString('o') -ne $state.word_start_time_utc) { throw 'OWNED_WINDOW_PID_BINDING_FAILED' }
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ownedWindow)
  $ownedWindow = $null
  $state.phase = 'EXPORT'
  $doc.ShowRevisions = $false
  # Item=0 is document content, not markup; no shared Word option is changed.
  $doc.ExportAsFixedFormat(@OUTPUT@, 17, $false, 0, 0, 1, 1, 0)
  $state.phase = 'CLOSE'
  $doc.Close(0)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
  $doc = $null
  $state.document_closed = $true
  foreach ($prop in @('Documents','Windows','ProtectedViewWindows')) {
    $collection = $app.$prop
    try { if ($null -eq $collection -or $null -eq $collection.Count -or $collection.Count -ne 0) { throw ('FOREIGN_OR_UNKNOWN_' + $prop + ': instance not quit') } }
    finally { if ($null -ne $collection) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($collection) } }
  }
  foreach ($prop in @('BackgroundSavingStatus','BackgroundPrintingStatus')) {
    if ($null -eq $app.$prop -or $app.$prop -ne 0) { throw ('WORD_BUSY_OR_UNKNOWN_' + $prop) }
  }
  $state.phase = 'PREPARE_TYPED_QUIT'
  $piaPath = $app.GetType().Assembly.Location
  $state.interop_assembly = $app.GetType().Assembly.FullName
  $state.interop_sha256 = (Get-FileHash -LiteralPath $piaPath -Algorithm SHA256).Hash.ToLowerInvariant()
  $typedCode = @'
using System;
using Word = Microsoft.Office.Interop.Word;
public static class OriginalWordTypedQuit {
    public static void Quit(object existingApplication) {
        Word._Application app = (Word._Application)existingApplication;
        object saveChanges = 0;
        object originalFormat = Type.Missing;
        object routeDocument = Type.Missing;
        app.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
    }
    public static string[] ArgumentTypes() {
        object saveChanges = 0;
        object originalFormat = Type.Missing;
        object routeDocument = Type.Missing;
        return new string[] { saveChanges.GetType().FullName,
            originalFormat.GetType().FullName, routeDocument.GetType().FullName };
    }
}
'@
  Add-Type -TypeDefinition $typedCode -ReferencedAssemblies $piaPath
  $state.quit_argument_types = [OriginalWordTypedQuit]::ArgumentTypes()
  $state.phase = 'QUIT_OWNED_INSTANCE'
  [OriginalWordTypedQuit]::Quit($app)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  $app = $null
  $state.quit_returned = $true
  $state.phase = 'QUIT_REQUESTED'
  # Host must exit naturally before the parent can verify process disappearance.
} catch {
  $state.error = $_.Exception.Message
  $state.exception_type = $_.Exception.GetType().FullName
  $state.hresult = $_.Exception.HResult
  if ($null -ne $_.Exception.InnerException) { $state.inner_error = $_.Exception.InnerException.Message; $state.inner_hresult = $_.Exception.InnerException.HResult }
  # Preserve uncertain app/document. No retry, forced quit, kill or foreign close.
} finally {
  $state.finished_utc = [DateTime]::UtcNow.ToString('o')
  [IO.File]::WriteAllText($receipt, ($state | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding($false)))
}
if (-not $state.quit_returned -or -not $state.document_closed -or $state.error) { exit 1 }
'''.replace('@RECEIPT@', ps_literal(receipt)).replace('@MACHINE@', ps_literal(binding['machine'])).replace(
        '@USER@', ps_literal(binding['user'])).replace('@SESSION@', ps_literal(binding['session'])).replace(
        '@OWNED@', ps_literal(owned)).replace('@OUTPUT@', ps_literal(output))



def exit_probe_script(state, binding):
    """Only read processes in a fresh host; never attach COM or request Quit."""
    return r"""$ErrorActionPreference='Stop'
$targetId=@PID@
$targetBirth=@BIRTH@
$expectedSession=@SESSION@
$started=[DateTime]::UtcNow.ToString('o')
$clock=[Diagnostics.Stopwatch]::StartNew()
$observations=0
$status='UNKNOWN'
try {
 if($env:COMPUTERNAME -ne @MACHINE@ -or $env:USERNAME -ne @USER@ -or [string][Diagnostics.Process]::GetCurrentProcess().SessionId -ne $expectedSession){throw 'OBSERVER_IDENTITY_MISMATCH'}
 do {
  $observations++
  $all=@(Get-Process WINWORD -ErrorAction SilentlyContinue|ForEach-Object{@{pid=$_.Id;session=$_.SessionId;birth=$_.StartTime.ToUniversalTime().ToString('o')}})
  $target=Get-Process -Id $targetId -ErrorAction SilentlyContinue
  $targetData=if($null -ne $target){@{pid=$target.Id;session=$target.SessionId;birth=$target.StartTime.ToUniversalTime().ToString('o');name=$target.ProcessName}}else{$null}
  if($null -ne $targetData -and ([string]$targetData.session -ne $expectedSession -or $targetData.birth -ne $targetBirth -or $targetData.name -ne 'WINWORD')){$status='PID_REUSED_OR_IDENTITY_CHANGED';break}
  if(@($all|Where-Object{$_.pid -ne $targetId}).Count -ne 0){$status='FOREIGN_WORD_PRESENT';break}
  if($null -eq $targetData -and $all.Count -eq 0){$status='EXIT_CONFIRMED';break}
  $status='ORIGINAL_PROCESS_REMAINS'
  Start-Sleep -Milliseconds 250
 }while($clock.Elapsed.TotalSeconds -lt 15)
 $record=@{schema=1;status=$status;machine=$env:COMPUTERNAME;user=$env:USERNAME;observer_session=[Diagnostics.Process]::GetCurrentProcess().SessionId;expected_pid=$targetId;expected_birth=$targetBirth;expected_session=$expectedSession;target_process=$targetData;word_processes=$all;observations=$observations;started_utc=$started;finished_utc=[DateTime]::UtcNow.ToString('o');com_attached=$false;quit_called=$false}
}catch{$record=@{schema=1;status='OBSERVER_ERROR';error=$_.Exception.Message;started_utc=$started;finished_utc=[DateTime]::UtcNow.ToString('o');com_attached=$false;quit_called=$false}}
$record|ConvertTo-Json -Depth 6 -Compress
if($status -ne 'EXIT_CONFIRMED'){exit 2}
""".replace('@PID@', str(state['word_pid'])).replace('@BIRTH@', ps_literal(state['word_start_time_utc'])).replace(
        '@SESSION@', ps_literal(binding['session'])).replace('@MACHINE@', ps_literal(binding['machine'])).replace(
        '@USER@', ps_literal(binding['user']))


def child_binding_valid(state, returncode):
    pid = state.get('word_pid')
    return (returncode == 0 and state.get('binding_schema') == 3 and
            state.get('native_started') is True and state.get('native_closed') is False and
            state.get('precreate_word_count') == 0 and state.get('postcreate_word_count') == 1 and
            state.get('identity_bound') is True and type(pid) is int and pid > 0 and
            type(state.get('owned_window_hwnd')) is int and state['owned_window_hwnd'] > 0 and
            state.get('owned_window_pid') == pid and bool(state.get('word_start_time_utc')) and
            state.get('document_closed') is True and state.get('quit_returned') is True and
            state.get('phase') == 'QUIT_REQUESTED' and not state.get('error') and
            state.get('quit_argument_types') == ['System.Int32', 'System.Reflection.Missing', 'System.Reflection.Missing'])


def exit_proof_valid(proof, state, binding, returncode):
    return (returncode == 0 and proof.get('schema') == 1 and proof.get('status') == 'EXIT_CONFIRMED' and
            proof.get('machine') == binding['machine'] and proof.get('user') == binding['user'] and
            str(proof.get('observer_session')) == str(binding['session']) and
            proof.get('expected_pid') == state['word_pid'] and
            proof.get('expected_birth') == state['word_start_time_utc'] and
            str(proof.get('expected_session')) == str(binding['session']) and
            'target_process' in proof and proof['target_process'] is None and
            proof.get('word_processes') == [] and proof.get('com_attached') is False and
            proof.get('quit_called') is False and bool(proof.get('started_utc')) and bool(proof.get('finished_utc')))


def run_native(owned, output, evidence, binding, runner=subprocess.run, exit_runner=subprocess.run):
    evidence = Path(evidence)
    script, receipt = evidence / 'native.ps1', evidence / 'native-com.json'
    script.write_text(native_script(owned, output, receipt, binding), encoding='utf-8-sig')
    try:
        proc = runner(['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                       '-File', str(script)], capture_output=True, text=True, timeout=420)
    except subprocess.TimeoutExpired:
        return dict(native_started=True, native_closed=False, phase='COM_TIMEOUT', uncertain=True)
    if not receipt.exists():
        return dict(native_started=True, native_closed=False, phase='COM_RECEIPT_MISSING',
                    child_returncode=proc.returncode, returncode=proc.returncode, uncertain=True)
    try:
        raw = receipt.read_bytes()
        child = json.loads(raw.decode('utf-8-sig'))
        if not isinstance(child, dict):
            raise ValueError('Child receipt is not an object')
    except (ValueError, UnicodeError, OSError) as exc:
        return dict(native_started=True, native_closed=False, phase='COM_RECEIPT_INVALID',
                    child_returncode=proc.returncode, returncode=proc.returncode, uncertain=True, error=str(exc))
    state = dict(child, child_phase=child.get('phase'), child_returncode=proc.returncode,
                 child_receipt_sha256=hashlib.sha256(raw).hexdigest(),
                 returncode=proc.returncode, stdout=proc.stdout, stderr=proc.stderr)
    if not child_binding_valid(child, proc.returncode):
        state.update(native_closed=False, uncertain=True, phase='CHILD_EVIDENCE_INVALID')
        return state
    # subprocess.run has waited for the original PowerShell host to terminate.
    state['child_host_exited'] = True
    probe_script = evidence / 'exit-observation.ps1'
    try:
        probe_script.write_text(exit_probe_script(child, binding), encoding='utf-8-sig')
        observation = exit_runner(['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                                   '-File', str(probe_script)], capture_output=True, text=True, timeout=30)
        (evidence / 'exit-observation.stdout').write_text(observation.stdout)
        (evidence / 'exit-observation.stderr').write_text(observation.stderr)
        proof = json.loads(observation.stdout)
        if not isinstance(proof, dict):
            raise ValueError('Exit observation is not an object')
        state.update(exit_observer_returncode=observation.returncode, exit_observation=proof)
        accepted = exit_proof_valid(proof, child, binding, observation.returncode)
    except (subprocess.TimeoutExpired, ValueError, UnicodeError, OSError) as exc:
        accepted = False
        state['exit_observation_error'] = str(exc)
    if not accepted:
        state.update(native_closed=False, uncertain=True, phase='POST_HOST_EXIT_UNCONFIRMED')
    else:
        state.update(native_closed=True, process_exit_confirmed=True, phase='COMPLETE')
    # Preserve child bytes/returncode. The parent derives a separate terminal receipt.
    try:
        (evidence / 'native-parent.json').write_text(json.dumps(state, ensure_ascii=False, indent=2))
    except OSError as exc:
        state.update(native_closed=False, uncertain=True, phase='PARENT_RECEIPT_WRITE_FAILED', error=str(exc))
    return state
