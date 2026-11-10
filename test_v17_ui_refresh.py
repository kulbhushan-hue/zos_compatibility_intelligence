"""Regression tests for colored layout and explicit all-vendor manual refresh."""
import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch
import auto_scrape_v13

class V17RefreshTests(unittest.TestCase):
    def test_main_source_monitor_refreshes_all_vendors(self):
        app=(Path(__file__).resolve().parents[1]/'app.py').read_text()
        page=app.split("elif page=='Source Monitor':",1)[1].split("elif page=='Extraction Coverage':",1)[0]
        self.assertIn('Refresh All Vendors',page)
        self.assertIn('_request_vendor_refresh(BASE)',page)
        self.assertIn('Advanced · Broadcom',page)

    def test_colored_css_and_config_are_bundled(self):
        root=Path(__file__).resolve().parents[1]
        self.assertIn('#0d9488',(root/'app.py').read_text())
        self.assertIn('base = "light"',(root/'.streamlit/config.toml').read_text())

    def test_manual_refresh_forces_new_source_cycle(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            folder=root/'exports'/'deep'
            folder.mkdir(parents=True)
            (root/'reference').mkdir()
            (root/'reference/vendor_sources_v13.csv').write_text('Vendor,URL\n')
            with patch('deep_scraper.main') as mock:
                auto_scrape_v13._run(root,False,200,force_fresh=True)
                self.assertIn('--fresh',mock.call_args.args[0])
                self.assertEqual(Path(mock.call_args.args[0][1]),folder)
            status=json.loads((folder/'auto_scrape_status.json').read_text())
            self.assertEqual(status['state'],'Completed')

    def test_requested_manual_refresh_is_nonblocking(self):
        with tempfile.TemporaryDirectory() as td:
            root=Path(td)
            with patch('auto_scrape_v13._run') as worker:
                self.assertTrue(auto_scrape_v13.request_refresh(root,max_pages=5))
                status=json.loads((root/'exports/deep/auto_scrape_status.json').read_text())
                self.assertEqual(status['state'],'Queued')
                self.assertFalse(auto_scrape_v13.request_refresh(root,max_pages=5))
                for _ in range(100):
                    if worker.called:break
                    time.sleep(.005)
                self.assertTrue(worker.called)
                self.assertEqual(worker.call_args.args[-1],True)

if __name__=='__main__':unittest.main()
