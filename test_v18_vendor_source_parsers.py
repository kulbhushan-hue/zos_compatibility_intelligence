"""End-to-end offline fixtures modeled on vendor-published layouts.
These fixtures test parsers; they do not claim live vendor extraction success.
"""
import csv
import tempfile
import unittest
from pathlib import Path
from vendor_evidence_v18 import extract_public_facts
from deep_scraper import inventory_and_queue, DeepCrawler, source_quality
from multi_vendor import safe_public_url

STAMP='2026-11-10T00:00:00+00:00'

class V18VendorFactsTests(unittest.TestCase):
    def parse(self,v,html):
        return extract_public_facts(html.encode(),v,'https://www.example.com/official',STAMP)

    def test_action_release_support_and_eos_fields(self):
        r=self.parse('Action Software International','''<table><tr><th>Release</th><th>General Availability</th><th>Maintenance Level</th><th>End of Service</th><th>System Compatibility</th></tr>
        <tr><td>eventACTION 8.08</td><td>2025-02-06</td><td>8.08.24</td><td>2030-03-31</td><td>z/OS 2.1 – z/OS 3.2</td></tr></table>''')
        self.assertEqual({x['zos_version'] for x in r if x['fact_type']=='Published z/OS support range'}, {'2.5','3.1','3.2'})
        self.assertIn(('End of service','2030-03-31'),{(x['fact_type'],x['value']) for x in r})
        self.assertIn(('Published maintenance level','8.08.24'),{(x['fact_type'],x['value']) for x in r})
        self.assertTrue(all(x['verification'].startswith('Official page') for x in r))

    def test_action_support_range_does_not_infer_later_os(self):
        r=self.parse('Action Software International','''<table><tr><th>Release</th><th>General Availability</th><th>Maintenance Level</th><th>End of Service</th><th>System Compatibility</th></tr><tr><td>eventACTION 8.05</td><td>2018-12-31</td><td>Service Withdrawn</td><td>2024-03-31</td><td>z/OS 1.12 – z/OS 3.1</td></tr></table>''')
        self.assertNotIn('3.2',{x['zos_version'] for x in r})

    def test_vanguard_tested_release_exact(self):
        r=self.parse('VANGUARD Integrity Prof.','''<table><tr><th>Vanguard Product</th><th>Version</th><th>Minimum Tested z/OS</th><th>Maximum Tested z/OS</th><th>Discontinuance Date</th></tr><tr><td>Vanguard Security Solutions</td><td>3.1</td><td>2.5</td><td>3.2</td><td>None Announced</td></tr></table>''')
        self.assertEqual({x['zos_version'] for x in r if x['fact_type']=='Vendor-tested z/OS range'}, {'2.5','3.1','3.2'})
        self.assertTrue(all(x['release']=='3.1' for x in r))

    def test_cbl_stops_at_31_and_never_infers_32(self):
        r=self.parse('Compute (Bridgend) Ltd.','''<h2>SELCOPY - IBM Mainframe</h2><table><tr><th>Operating System</th><th>Version/Release</th><th>Maintenance</th></tr><tr><td>z/OS</td><td>3.1 and below</td><td>None</td></tr></table>''')
        self.assertEqual({x['zos_version'] for x in r},{'3.1'})
        self.assertEqual(r[0]['value'],'None')

    def test_confluent_certification_is_only_24(self):
        r=self.parse('Confluent, Inc','<h2>Certified z/OS Version</h2><p>Connect on z/OS is certified with z/OS version 2, release 4.</p>')
        self.assertEqual([x['zos_version'] for x in r],['2.4'])

    def test_bmc_minimum_is_not_exact_release(self):
        r=self.parse('BMC','''<table><tr><th>Operating systems/ Mainframe</th><th>BMC Performance Predictor for Mainframes</th><th>Comments</th></tr><tr><td>z/OS 3.2</td><td>X</td><td>Support started in version 2.5.00</td></tr></table>''')
        self.assertEqual(len(r),1)
        self.assertEqual((r[0]['release'],r[0]['zos_version'],r[0]['value']),('Not established','3.2','2.5.00'))

    def test_systemware_level_is_not_ptf_installation(self):
        r=self.parse('Systemware, Inc.','''<h3>IBM z/OS 3.2 compatibility was proven on the following Systemware maintenance levels:</h3><ul><li>Release 5.2.0: SW52D4</li><li>Release 5.1.0: SW51D4</li></ul>''')
        self.assertEqual({x['value'] for x in r},{'SW52D4','SW51D4'})
        self.assertTrue(all('not proof' in x['verification'] for x in r))

    def test_quality_report_counts_document_facts(self):
        rows=source_quality([{'vendor':'BMC','url':'https://docs.bmc.com/x','state':'Read'}],[],[],[],
          [{'vendor':'BMC','source_url':'https://docs.bmc.com/x','product':'ABC','release':'Not established','zos_version':'3.2','fact_type':'Product-family support matrix','value':'2.5.00'}])
        self.assertEqual(rows[0]['document_fact_rows'],1)
        self.assertEqual(rows[0]['result_class'],'Published document facts only')

    def test_29_vendor_inventory_report_and_approved_domain_filter(self):
        root=Path(__file__).resolve().parents[1]
        queue,inventory=inventory_and_queue(root,root/'reference/vendor_product_source_worklist.csv',root/'reference/vendor_sources_v18.csv')
        self.assertEqual(len(inventory),29)
        self.assertEqual(sum(map(len,inventory.values())),623)
        self.assertGreater(len({x['vendor'] for x in queue}),12)
        self.assertTrue(all(safe_public_url(x['url'],x['vendor']) for x in queue))
        self.assertFalse(safe_public_url('https://mainframemodernization.org/products/csf/','FIS'))

if __name__=='__main__':unittest.main()
