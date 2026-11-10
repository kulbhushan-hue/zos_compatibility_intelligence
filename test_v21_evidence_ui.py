import unittest
from pathlib import Path
import tempfile
import pandas as pd
from evidence_view_v21 import build_evidence_table,evidence_counts,source_run_status,readable_source_status
from deep_app_bridge_v13 import merge_verified_deep

class EvidenceViewV21Tests(unittest.TestCase):
    def setUp(self):
        self.frame=pd.DataFrame([
            {'vendor':'BMC','product':'Foo','vrm':'1.0','zos32':'Yes','provenance_3.2':'Excel AMXT','link':'https://example.org','baseline_status_3.2':'Yes'},
            {'vendor':'Precisely','product':'Bar','vrm':'2.1','zos32':'No','provenance_3.2':'Exact live vendor source (not installation verified)','verified_source_url_3.2':'https://vendor.example/bar','baseline_status_3.2':'Unverified'},
            {'vendor':'Broadcom','product':'Baz','vrm':'3.0','zos32':'Yes','provenance_3.2':'Vendor snapshot','verified_source_url_3.2':'https://vendor.example/baz','baseline_status_3.2':'No'},
        ])
    def test_baseline_does_not_count_as_vendor_verified(self):
        result=build_evidence_table(self.frame,'3.2')
        self.assertEqual(result.iloc[0]['Vendor-verified status'],'Not verified')
        self.assertEqual(result.iloc[0]['Official source'],'')
        self.assertEqual(result.iloc[1]['Baseline status'],'Unverified')
        self.assertEqual(result.iloc[1]['Vendor-verified status'],'No')
        self.assertEqual(evidence_counts(result),{'verified':2,'yes':1,'no':1,'not_verified':1})
    def test_provenance_without_source_url_not_verified(self):
        df=self.frame.copy();df.loc[1,'verified_source_url_3.2']=''
        self.assertEqual(build_evidence_table(df,'3.2').iloc[1]['Vendor-verified status'],'Not verified')
    def test_bridge_attaches_source_and_retains_baseline(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)/'claims.csv'
            p.write_text('vendor,product,release,zos_version,compatibility,evidence_level,source_url,source_timestamp,source_excerpt\nPrecisely,Bar,2.1,3.2,No,exact,https://vendor.example/bar,2026-11-10,evidence\n')
            original=self.frame.iloc[[1]].copy()
            original.loc[1,'provenance_3.2']='CSV baseline'
            result=merge_verified_deep(original,p)
            self.assertEqual(result.iloc[0]['verified_source_url_3.2'],'https://vendor.example/bar')
            self.assertEqual(result.iloc[0]['baseline_status_3.2'],'Unverified')
            self.assertEqual(result.iloc[0]['verified_excerpt_3.2'],'evidence')
    def test_missing_collector_status(self):
        self.assertEqual(source_run_status('/path/which/does/not/exist'),{})
        self.assertIn('No collector status',readable_source_status({}))

if __name__=='__main__':unittest.main()
