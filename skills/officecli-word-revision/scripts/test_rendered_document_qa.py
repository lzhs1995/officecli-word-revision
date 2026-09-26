import unittest
from rendered_document_qa import check_page


class RenderedDocumentTests(unittest.TestCase):
    def char(self, text='A', font='ABCDEF+TimesNewRomanPSMT', x=20):
        return {'text':text, 'fontname':font, 'x0':x, 'x1':x+6, 'top':10}

    def test_real_matching_glyphs_and_subset_prefix(self):
        result=check_page([self.char()], {'font_regions':[{'box':[0,0,100,100], 'latin':['TimesNewRomanPSMT']}]})
        self.assertTrue(result['pass'])
        self.assertEqual(result['measurements'][0]['glyphs_checked'],1)

    def test_present_but_wrong_font_fails(self):
        result=check_page([self.char(font='ArialMT')], {'font_regions':[{'box':[0,0,100,100], 'latin':['TimesNewRomanPSMT']}]})
        self.assertFalse(result['pass'])

    def test_empty_rules_empty_regions_and_wrong_script_do_not_pass(self):
        for spec in ({}, {'font_regions':[{'box':[200,0,300,100], 'latin':['TimesNewRomanPSMT']}]},
                     {'font_regions':[{'box':[0,0,100,100], 'han':['SimSun']}]}):
            self.assertFalse(check_page([self.char()],spec)['pass'])

    def test_center_uses_declared_text_area(self):
        chars=[self.char(x=20),self.char('B',x=26)]
        spec={'stories':[{'text':'AB','box':[0,0,100,100],'text_area':[10,42],'alignment':'center','tolerance_pt':.1}]}
        self.assertTrue(check_page(chars,spec)['pass'])
        spec['stories'][0]['text_area']=[0,100]
        self.assertFalse(check_page(chars,spec)['pass'])


if __name__=='__main__':
    unittest.main()
