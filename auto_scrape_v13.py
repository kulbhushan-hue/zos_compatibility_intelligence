"""Launch the permitted vendor crawler on Streamlit startup without extra commands.

Does not block the UI or require Docker. Reports results to a JSON status file.
Playwright is optional on Windows/Linux and attempted only when installed.
On old Linux systems it can fail due to incompatible GLIBC; this is reported,
not treated as a successful IBM report extraction.
"""
from __future__ import annotations
import json
import os
import sys
import threading
import time
import traceback
from datetime import datetime, timezone
from pathlib import Path

STATUS = 'auto_scrape_status.json'
_STARTED = False
_MUTEX = threading.Lock()

def _save_status(root, state, **fields):
    output = root/'exports'/'deep'
    output.mkdir(parents=True,exist_ok=True)
    target=output/STATUS
    data={'state':state,'timestamp_utc':datetime.now(timezone.utc).isoformat(),**fields}
    tmp=target.with_suffix('.tmp')
    tmp.write_text(json.dumps(data,indent=2),encoding='utf-8')
    tmp.replace(target)

def _run(root:Path, browser_enabled:bool, max_pages:int, force_fresh:bool=False):
    lockfile=root/'exports'/'deep'/'.auto_scrape.lock'
    lockfile.parent.mkdir(parents=True,exist_ok=True)
    fd=None
    owned=False
    try:
        try:
            fd=os.open(str(lockfile), os.O_CREAT|os.O_EXCL|os.O_WRONLY,0o600)
        except FileExistsError:
            if time.time()-lockfile.stat().st_mtime>6*3600:
                lockfile.unlink(missing_ok=True)
                fd=os.open(str(lockfile),os.O_CREAT|os.O_EXCL|os.O_WRONLY,0o600)
            else:return
        owned=True
        os.write(fd,str(os.getpid()).encode())
        os.close(fd);fd=None
        from deep_scraper import main as scrape
        _save_status(root,'Scraping',message='Background source refresh started')
        # A completed checkpoint with zero pending URLs is not a refresh.
        # When startup is allowed (24-hour interval has elapsed), run a fresh
        # source cycle. Preserve partial queue runs so they can resume.
        checkpoint=root/'exports'/'deep'/'deep_scrape_checkpoint.json'
        fresh=force_fresh
        if checkpoint.exists() and not force_fresh:
            try:
                saved=json.loads(checkpoint.read_text(encoding='utf-8'))
                fresh=not bool(saved.get('queue'))
            except (ValueError,OSError):
                fresh=True
        args=['--output',str(root/'exports'/'deep'),'--max-pages',str(max_pages),'--depth','2',
              '--sources',str(root/'reference'/'vendor_sources_v18.csv')]
        if fresh:args.append('--fresh')
        scrape(args)
        if browser_enabled and (sys.platform=='win32' or sys.platform.startswith('linux')):
            # Windows Edge/Chrome or a pre-installed compatible browser is used.
            # Never silently install browsers or ignore robots restrictions.
            try:
                import importlib.util
                if importlib.util.find_spec('playwright'):
                    from dynamic_capture_v12 import capture
                    report=capture(root/'exports'/'deep',limit=30,delay=1.5)
                    if report.get('captured',0):
                        scrape(['--output',str(root/'exports'/'deep'),'--reparse-cache'])
            except Exception as exc:
                _save_status(root,'Partial',message='Static scraping completed. Optional browser capture unavailable',browser_error=str(exc)[:500])
                return
        # A finished URL queue is not proof of complete vendor extraction.
        coverage_file=root/'exports'/'deep'/'vendor_extraction_coverage.csv'
        coverage_summary={}
        if coverage_file.is_file():
            try:
                import csv
                with coverage_file.open(encoding='utf-8-sig',newline='') as fh:
                    items=list(csv.DictReader(fh))
                verified=sum(bool(int(x.get('urls_with_exact_claims') or 0)) for x in items)
                coverage_summary={'vendors_reported':len(items),'vendors_with_exact_inventory_evidence':verified}
            except (OSError,ValueError):
                pass
        incomplete=bool(coverage_summary and coverage_summary['vendors_with_exact_inventory_evidence'] < coverage_summary['vendors_reported'])
        _save_status(root,'Partial' if incomplete else 'Completed',
            message='Crawl finished; not all vendors have product/release evidence' if incomplete else 'URL crawl completed; inspect vendor coverage',
            coverage_report='vendor_extraction_coverage.csv',product_report='vendor_product_extraction_coverage.csv',**coverage_summary)
    except Exception as exc:
        _save_status(root,'Failed',error=str(exc)[:500],error_type=type(exc).__name__)
    finally:
        if fd is not None:os.close(fd)
        if owned and lockfile.exists():
            try:lockfile.unlink()
            except OSError:pass

def launch(root, *, max_pages=200, interval_seconds=24*3600, browser_enabled=True):
    """Nonblocking, process-local, once-per-refresh startup hook."""
    global _STARTED
    root=Path(root).resolve()
    if os.getenv('ZOS_AUTO_SCRAPE','1').strip().lower() in ('0','false','off','no'):return False
    with _MUTEX:
        if _STARTED:return False
        status=root/'exports'/'deep'/STATUS
        if status.is_file():
            try:
                data=json.loads(status.read_text(encoding='utf-8'))
                age=time.time()-datetime.fromisoformat(data['timestamp_utc']).timestamp()
                if 0<=age<interval_seconds and data.get('state') in ('Completed','Partial','Scraping'):
                    _STARTED=True;return False
            except (ValueError,KeyError,TypeError):pass
        _STARTED=True
    threading.Thread(target=_run,args=(root,browser_enabled,max(1,int(max_pages))),
                     daemon=True,name='zos-vendor-source-refresh').start()
    return True


def request_refresh(root, *, max_pages=200, browser_enabled=True):
    """Request a full vendor refresh without blocking a Streamlit rerun.

    Acquires the same on-disk lock as the automatic collector, preventing
    simultaneous refreshes across Streamlit sessions. The lock is released by
    _run on exit, and a stale lock is handled by the existing worker.
    """
    root=Path(root).resolve()
    folder=root/'exports'/'deep'
    folder.mkdir(parents=True,exist_ok=True)
    lock=folder/'.auto_scrape.lock'
    status_file=folder/STATUS
    if status_file.is_file():
        try:
            state=json.loads(status_file.read_text(encoding='utf-8'))
            age=time.time()-datetime.fromisoformat(state['timestamp_utc']).timestamp()
            if state.get('state') in ('Queued','Scraping') and 0<=age<60:
                return False
        except (OSError,KeyError,TypeError,ValueError):
            pass
    if lock.exists():
        try:
            if time.time()-lock.stat().st_mtime<=6*3600:
                return False
        except OSError:
            return False
    _save_status(root,'Queued',message='All-vendor refresh requested by user')
    threading.Thread(target=_run,args=(root,browser_enabled,max(1,int(max_pages)),True),
                     daemon=True,name='zos-manual-all-vendor-refresh').start()
    return True
