import unittest
from rendered_numeric_qa import check_decimal_inventory
class NumericAtomicityTests(unittest.TestCase):
 def test_complete(self):self.assertTrue(check_decimal_inventory(['0.274','-4.73 (-10.22, 0.77)'],['0.274  −4.73 (−10.22, 0.77)'])['all_pass'])
 def test_split_decimal(self):self.assertFalse(check_decimal_inventory(['0.274'],['0.2','74'])['all_pass'])
 def test_missing_duplicate(self):self.assertFalse(check_decimal_inventory(['0.274','0.274'],['0.274'])['all_pass'])
 def test_changed_sign(self):self.assertFalse(check_decimal_inventory(['-4.73'],['4.73'])['all_pass'])
 def test_empty_is_not_proof(self):self.assertFalse(check_decimal_inventory([],[])['all_pass'])
 def test_precision_equivalence(self):self.assertTrue(check_decimal_inventory(['.974'],['0.974'])['all_pass'])
if __name__=='__main__':unittest.main()
