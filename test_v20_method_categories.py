"""Offline source-evidence fixtures, not proof of live collection."""
import csv
import tempfile
import unittest
from pathlib import Path
from vendor_methods_v20 import METHODS, document_facts, write_method_report
from deep_scraper import inventory_and_queue

T='2026-11-10T00:00:00+00:00'

class MethodCategoriesTests(unittest.TestCase):
    def test_all_29_inventory_vendors_have_method_category(self):
        root=Path(__file__).resolve().parents[1]
        candidates,byvendor=inventory_and_queue(root,root/'reference/vendor_product_source_worklist.csv',root/'reference/vendor_sources_v18.csv')
        self.assertEqual(len(byvendor),29)
        self.assertEqual(set(METHODS),set(byvendor))
        self.assertIn(('ASG','https://docs-be.rocketsoftware.com/bundle/ZekeSchedulingforzOS_InstallationGuide_V6.2.0/raw/resource/enus/ZekeSchedulingforzOS_InstallationGuide_V6.2.0.pdf'),{(x['vendor'],x['url']) for x in candidates})
    def test_partial_pdf_requirements_stay_evidence_only(self):
        pdftext='ASG-Zeke Scheduling for z/OS Release 6.2.0\n\nZeke 7.0 minimum release of z/OS is z/OS 2.2.\n\nZeke 7.0 and OASIS 4.0 are not supported on VSE.'
        facts=document_facts(pdftext,'ASG','https://docs-be.rocketsoftware.com/doc.pdf',T,is_pdf=True)
        self.assertTrue(facts)
        self.assertEqual(facts[0]['fact_type'],'Published OS requirement / support statement')
        self.assertEqual(facts[0]['zos_version'],'2.2')
        self.assertNotIn('compatibility',facts[0])
    def test_dynamic_ibm_is_not_claimed_extracted(self):
        self.assertEqual(METHODS['IBM'][0],'Dynamic')
        with tempfile.TemporaryDirectory() as tmp:
            rows=write_method_report(tmp,{'IBM':[{'product':'DB2','release':'12'}]},
              [{'vendor':'IBM','url':'https://www.ibm.com/software/reports/compatibility/clarity/index.html'}],
              [{'vendor':'IBM','url':'https://www.ibm.com/software/reports/compatibility/clarity/index.html','state':'Read'}],[],[])
            self.assertEqual(rows[0]['exact_inventory_claims'],0)
            self.assertEqual(rows[0]['result'],'Pages readable; no extracted facts')
            self.assertTrue(Path(tmp,'vendor_method_coverage_v20.csv').exists())
    def test_conditional_no_generic_os_certification(self):
        html='<p>bash version 5.2 available on z/OS through zopen package manager.</p>'
        facts=document_facts(html,'Open Source Software','https://zopen.community',T)
        self.assertEqual(facts,[])
    def test_partial_extract_requires_local_product_context(self):
        html='<p>General requirements include z/OS 3.2 certified.</p><p>Rocket Software Zeke release notes.</p>'
        self.assertEqual(document_facts(html,'Rocket','https://docs.rocketsoftware.com',T),[])
    def test_named_rocket_product_support_stays_evidence_only(self):
        html='<p>Zeke 7.0 minimum release of z/OS is z/OS 2.2; see release notes for supported versions.</p>'
        facts=document_facts(html,'Rocket','https://docs.rocketsoftware.com',T)
        self.assertEqual(len(facts),1)
        self.assertEqual(facts[0]['product'],'Zeke')
        self.assertEqual(facts[0]['zos_version'],'2.2')
        self.assertEqual(facts[0]['release'],'7.0')

if __name__=='__main__':unittest.main()
