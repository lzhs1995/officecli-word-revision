"""Synthetic exit-proof cases; no COM or native applications are invoked."""
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('candidate', Path(__file__).with_name('word_windows_native.py'))
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
BIND = dict(machine='TEST-WINDOWS', user='fixture-user', session='7')
CHILD = dict(binding_schema=3,native_started=True,native_closed=False,precreate_word_count=0,
             postcreate_word_count=1,identity_bound=True,word_pid=8157,owned_window_hwnd=12345,
             owned_window_pid=8157,word_start_time_utc='2020-01-02T09:00:00.0000000Z',
             document_closed=True,quit_returned=True,phase='QUIT_REQUESTED',
             quit_argument_types=['System.Int32','System.Reflection.Missing','System.Reflection.Missing'])
PROOF = dict(schema=1,status='EXIT_CONFIRMED',machine='TEST-WINDOWS',user='fixture-user',observer_session=7,
             expected_pid=8157,expected_birth=CHILD['word_start_time_utc'],expected_session='7',
             target_process=None,word_processes=[],com_attached=False,quit_called=False,
             started_utc='2020-01-02T09:00:02Z',finished_utc='2020-01-02T09:00:03Z')

class ParentExitTests(unittest.TestCase):
    def run_case(self, child=CHILD, proof=PROOF, child_rc=0, observer_rc=0, raw_override=None):
        with tempfile.TemporaryDirectory() as td:
            p=Path(td);calls=[];raw=None
            def child_runner(argv,**kwargs):
                nonlocal raw
                calls.append('child_complete')
                if child is not None:
                    raw=raw_override if raw_override is not None else json.dumps(child)
                    (p/'native-com.json').write_text(raw)
                return subprocess.CompletedProcess(argv,child_rc,'','')
            def observer(argv,**kwargs):
                self.assertEqual(calls,['child_complete'])
                calls.append('observe_after_child')
                return subprocess.CompletedProcess(argv,observer_rc,json.dumps(proof),'')
            if isinstance(proof, BaseException):
                def observer(argv, **kwargs):
                    calls.append('observe_after_child')
                    raise proof
            result=m.run_native('owned.docx','out.pdf',p,BIND,child_runner,observer)
            if raw is not None:self.assertEqual((p/'native-com.json').read_text(),raw)
            return result,calls

    def test_success_requires_post_host_independent_zero(self):
        r,calls=self.run_case();self.assertTrue(r['native_closed']);self.assertEqual(r['child_phase'],'QUIT_REQUESTED')
        self.assertEqual(r['child_returncode'],0);self.assertEqual(r['exit_observer_returncode'],0)
        self.assertEqual(calls,['child_complete','observe_after_child'])
    def test_original_process_still_running_remains_unknown(self):
        proof=dict(PROOF,status='ORIGINAL_PROCESS_REMAINS',target_process={'pid':8157},word_processes=[{'pid':8157}])
        r,_=self.run_case(proof=proof,observer_rc=2);self.assertFalse(r['native_closed']);self.assertTrue(r['uncertain'])
    def test_pid_reused_does_not_count_as_closed(self):
        proof=dict(PROOF,status='PID_REUSED_OR_IDENTITY_CHANGED',target_process={'pid':8157,'birth':'new'})
        r,_=self.run_case(proof=proof,observer_rc=2);self.assertFalse(r['native_closed'])
    def test_missing_or_invalid_receipt_never_observes_for_success(self):
        for child,raw in [(None,None),(CHILD,'{'),(CHILD,'null')]:
            with self.subTest(raw=raw):
                r,calls=self.run_case(child=child,raw_override=raw);self.assertFalse(r['native_closed']);self.assertEqual(calls,['child_complete'])
    def test_foreign_word_is_not_accepted(self):
        proof=dict(PROOF,status='FOREIGN_WORD_PRESENT',word_processes=[{'pid':999}])
        r,_=self.run_case(proof=proof,observer_rc=2);self.assertFalse(r['native_closed'])
    def test_child_error_or_rc1_stays_failed_even_with_empty_process_proof(self):
        for child,rc in [(dict(CHILD,error='COM failed'),0),(CHILD,1),(dict(CHILD,binding_schema=2),0)]:
            with self.subTest(rc=rc):
                r,calls=self.run_case(child=child,child_rc=rc);self.assertFalse(r['native_closed']);self.assertEqual(r['child_returncode'],rc);self.assertEqual(calls,['child_complete'])
    def test_incomplete_or_wrong_identity_observation_refused(self):
        for change in [{'machine':'other'},{'expected_birth':'other'},{'target_process':{}},{'word_processes':None},{'com_attached':True},{'observer_session':2},{'status':'UNKNOWN'}]:
            with self.subTest(change=change):
                r,_=self.run_case(proof=dict(PROOF,**change));self.assertFalse(r['native_closed'])
    def test_observer_launch_failure_or_timeout_remains_unknown(self):
        for exc in [OSError('observer unavailable'), subprocess.TimeoutExpired(['powershell'],30)]:
            r,_=self.run_case(proof=exc);self.assertFalse(r['native_closed']);self.assertTrue(r['uncertain'])
    def test_child_unknown_close_or_quit_is_refused(self):
        for change in [{'document_closed':None},{'quit_returned':None},{'owned_window_pid':99},{'native_closed':True}]:
            with self.subTest(change=change):
                r,calls=self.run_case(child=dict(CHILD,**change));self.assertFalse(r['native_closed']);self.assertEqual(calls,['child_complete'])

    def test_parent_receipt_write_failure_cannot_claim_closure(self):
        write_text = Path.write_text
        def fail_parent(path, *args, **kwargs):
            if path.name == 'native-parent.json':
                raise OSError('synthetic evidence storage failure')
            return write_text(path, *args, **kwargs)
        with patch.object(Path, 'write_text', fail_parent):
            result, calls = self.run_case()
        self.assertEqual(calls, ['child_complete', 'observe_after_child'])
        self.assertFalse(result['native_closed'])
        self.assertTrue(result['uncertain'])
        self.assertEqual(result['phase'], 'PARENT_RECEIPT_WRITE_FAILED')

if __name__=='__main__':unittest.main(verbosity=2)
