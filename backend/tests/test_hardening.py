import os
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

from fastapi import HTTPException

BACKEND_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BACKEND_DIR))

import db
import excel_engine
import main


class HardeningTests(unittest.TestCase):
    def test_financial_year_validation_rejects_bad_values(self):
        with self.assertRaises(HTTPException):
            main._validate_financial_year("2026")
        with self.assertRaises(HTTPException):
            main._validate_financial_year("2026-2028")
        self.assertEqual(main._validate_financial_year("2026-2027"), "2026-2027")

    def test_upload_validation_rejects_wrong_type_and_bad_xlsx(self):
        with self.assertRaises(HTTPException):
            main._validate_upload("report.csv", b"not an excel file")
        with self.assertRaises(HTTPException):
            main._validate_upload("report.xlsx", b"not a zip")

    def test_month_detection_does_not_default_to_current_month(self):
        with self.assertRaises(ValueError):
            excel_engine.infer_month("upload.xlsx", "no month here")

    def test_failed_upload_reconciliation_rolls_back_rows(self):
        previous_db_path = os.environ.get("AMFI_DB_PATH")
        old_parse_upload = excel_engine.parse_upload
        old_validate = excel_engine.validate_uploaded_month_against_generated

        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "amfi.db"
            os.environ["AMFI_DB_PATH"] = str(db_path)
            conn = sqlite3.connect(db_path)
            try:
                conn.execute(db.SCHEMA.format(table="amfi_metrics"))
                conn.commit()
            finally:
                conn.close()

            def fake_parse_upload(_upload_bytes, _filename):
                return (
                    [
                        {
                            "scheme_key": "sample|open ended schemes|equity",
                            "scheme": "Sample Scheme",
                            "asset_type": "Open Ended Schemes",
                            "scheme_structure": "Equity",
                            "debt_equity": "Equity",
                            "sales_prod_mis": "Retail",
                            "fund_type": "Active Equity",
                            "metrics": {
                                "no_schemes": 1,
                                "folios": 2,
                                "funds_mobilized": 3,
                                "redemption": 4,
                                "net_inflow": -1,
                                "net_aum": 5,
                                "avg_aum": 6,
                                "seg_portfolios": 0,
                                "seg_aum": 0,
                            },
                            "metric_displays": {},
                        }
                    ],
                    excel_engine.month_info(5, 2026),
                    [],
                )

            excel_engine.parse_upload = fake_parse_upload
            excel_engine.validate_uploaded_month_against_generated = lambda *args, **kwargs: ["forced mismatch"]

            try:
                with self.assertRaises(ValueError):
                    excel_engine.process_upload_db(b"ignored", "May'26.xlsx")
                conn = sqlite3.connect(db_path)
                try:
                    count = conn.execute("SELECT COUNT(*) FROM amfi_metrics").fetchone()[0]
                finally:
                    conn.close()
                self.assertEqual(count, 0)
            finally:
                excel_engine.parse_upload = old_parse_upload
                excel_engine.validate_uploaded_month_against_generated = old_validate
                if previous_db_path is None:
                    os.environ.pop("AMFI_DB_PATH", None)
                else:
                    os.environ["AMFI_DB_PATH"] = previous_db_path


if __name__ == "__main__":
    unittest.main()
