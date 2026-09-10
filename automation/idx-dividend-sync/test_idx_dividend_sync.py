#!/usr/bin/env python3
"""Behavior tests for the IDX cash-dividend CSV updater."""

import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from idx_dividend_sync import _event_from_dividend_schedule, execute, fetch_cash_events, sync_cash_events


class SyncCashEventsTests(unittest.TestCase):
    def test_new_cash_event_creates_the_matching_ticker_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            event = {
                "code": "TEST",
                "cashDividend": 12.5,
                "exDividend": "2026-08-31",
                "paymentDate": "2026-09-18",
                "fiscalYear": "2026",
                "note": "F",
            }

            result = sync_cash_events(repo, [event])

            target = repo / "jkse" / "dividends" / "TEST.csv"
            self.assertEqual(result, {"added": 1, "updated": 0, "skipped": 0, "files": ["TEST.csv"]})
            self.assertNotIn(b"\r\n", target.read_bytes())
            with target.open(newline="") as handle:
                self.assertEqual(
                    list(csv.DictReader(handle)),
                    [{
                        "ex_date": "2026-08-31",
                        "dividend": "12.5",
                        "payment_date": "2026-09-18",
                        "fiscal_year": "2026",
                        "dividend_type": "final",
                    }],
                )

    def test_matching_event_updates_dates_without_erasing_known_fiscal_year(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            target = repo / "jkse" / "dividends" / "TEST.csv"
            target.parent.mkdir(parents=True)
            target.write_text(
                "ex_date,dividend,payment_date,fiscal_year,dividend_type\n"
                "2026-08-31,10.0,2026-09-10,2025,final\n"
            )
            event = {
                "code": "TEST",
                "cashDividend": 12.5,
                "exDividend": "2026-08-31",
                "paymentDate": "2026-09-18",
                "note": "F",
            }

            result = sync_cash_events(repo, [event])

            self.assertEqual(result, {"added": 0, "updated": 1, "skipped": 0, "files": ["TEST.csv"]})
            with target.open(newline="") as handle:
                self.assertEqual(
                    list(csv.DictReader(handle)),
                    [{
                        "ex_date": "2026-08-31",
                        "dividend": "12.5",
                        "payment_date": "2026-09-18",
                        "fiscal_year": "2025",
                        "dividend_type": "final",
                    }],
                )

    def test_matching_event_backfills_a_missing_fiscal_year(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            target = repo / "jkse" / "dividends" / "TEST.csv"
            target.parent.mkdir(parents=True)
            target.write_text(
                "ex_date,dividend,payment_date,fiscal_year,dividend_type\n"
                "2026-08-31,10.0,2026-09-10,,final\n"
            )
            event = {
                "code": "TEST",
                "cashDividend": 12.5,
                "exDividend": "2026-08-31",
                "paymentDate": "2026-09-18",
                "fiscalYear": "2026",
                "note": "F",
            }

            result = sync_cash_events(repo, [event])

            self.assertEqual(result, {"added": 0, "updated": 1, "skipped": 0, "files": ["TEST.csv"]})
            with target.open(newline="") as handle:
                self.assertEqual(
                    list(csv.DictReader(handle)),
                    [{
                        "ex_date": "2026-08-31",
                        "dividend": "12.5",
                        "payment_date": "2026-09-18",
                        "fiscal_year": "2026",
                        "dividend_type": "final",
                    }],
                )

    def test_malformed_event_is_skipped_without_writing_a_path_outside_the_dividend_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            event = {
                "code": "../../escape",
                "cashDividend": 12.5,
                "exDividend": None,
                "paymentDate": "2026-09-18",
                "note": "F",
            }

            result = sync_cash_events(repo, [event])

            self.assertEqual(result, {"added": 0, "updated": 0, "skipped": 1, "files": []})
            self.assertFalse((repo / "escape.csv").exists())

    def test_new_event_is_ordered_newest_first_like_existing_jkse_csvs(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            target = repo / "jkse" / "dividends" / "TEST.csv"
            target.parent.mkdir(parents=True)
            target.write_text(
                "ex_date,dividend,payment_date,fiscal_year,dividend_type\n"
                "2026-01-10,5.0,2026-01-28,2025,interim\n"
            )
            event = {
                "code": "TEST",
                "cashDividend": 12.5,
                "exDividend": "2026-08-31",
                "paymentDate": "2026-09-18",
                "note": "F",
            }

            sync_cash_events(repo, [event])

            with target.open(newline="") as handle:
                self.assertEqual(
                    [row["ex_date"] for row in csv.DictReader(handle)],
                    ["2026-08-31", "2026-01-10"],
                )

    def test_execute_fetches_events_then_returns_source_and_write_counts(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            event = {
                "code": "TEST",
                "cashDividend": 12.5,
                "exDividend": "2026-08-31",
                "paymentDate": "2026-09-18",
                "note": "F",
            }

            result = execute(repo, lambda: [event])

            self.assertEqual(
                result,
                {"source_rows": 1, "added": 1, "updated": 0, "skipped": 0, "files": ["TEST.csv"]},
            )

    def test_dividend_schedule_uses_per_share_amount_not_total_payout(self):
        # PyMuPDF can return the two table values after the same per-share label.
        # The total payout precedes the intended per-share amount in that text order.
        schedule_text = """Dividen Tunai Interim
Perseroan menyampaikan rencana pembagian Dividen Interim untuk periode tahun buku 2026
Dividen per saham
IDR
IDR
169.147.398.000
30
Jadwal pembagian dividen:
Tanggal Daftar Pemegang Saham (DPS) yang berhak atas dividen tunai
Tanggal Cum Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Ex Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Cum Dividen di Pasar Tunai
Tanggal Ex Dividen di Pasar Tunai
Tanggal Pembayaran Dividen
15 September 2026
11 September 2026
14 September 2026
15 September 2026
16 September 2026
21 September 2026
"""

        event = _event_from_dividend_schedule("DKFT", schedule_text)

        self.assertEqual(
            event,
            {
                "code": "DKFT",
                "cashDividend": "30",
                "exDividend": "2026-09-14",
                "paymentDate": "2026-09-21",
                "fiscalYear": "2026",
                "note": "I",
            },
        )

    def test_dividend_schedule_prefers_confirmed_per_share_amount_over_total_payout(self):
        schedule_text = """Dividen Tunai Interim
Perseroan menyampaikan rencana pembagian Dividen Interim untuk periode tahun buku 2026
Dividen Per Saham (Jika sudah ada kepastian jumlah saham yang akan dibagi)
IDR
66
Dividen per saham sudah ditentukan
Total Nilai Dividen sekurang-kurangnya
Dividen per saham
6.159.999.999.912
IDR
Tidak
66
IDR
IDR
6.159.999.999.912
Total Nilai Dividen setinggi-tingginya
Jadwal pembagian dividen:
Tanggal Daftar Pemegang Saham (DPS) yang berhak atas dividen tunai
Tanggal Cum Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Ex Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Cum Dividen di Pasar Tunai
Tanggal Ex Dividen di Pasar Tunai
Tanggal Pembayaran Dividen
17 September 2026
15 September 2026
16 September 2026
17 September 2026
18 September 2026
02 October 2026
"""

        event = _event_from_dividend_schedule("BMRI", schedule_text)

        self.assertEqual(event["cashDividend"], "66")

    def test_fetch_discovers_interim_dividend_from_official_announcement_when_statistics_feed_is_empty(self):
        import fitz

        document = fitz.open()
        page = document.new_page()
        page.insert_text(
            (72, 72),
            """Dividen Tunai Interim
Perseroan menyampaikan rencana pembagian Dividen Interim untuk periode tahun buku 2026
Dividen Per Saham
IDR
100
Jadwal pembagian dividen:
Tanggal Daftar Pemegang Saham (DPS) yang berhak atas dividen tunai
Tanggal Cum Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Ex Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Cum Dividen di Pasar Tunai
Tanggal Ex Dividen di Pasar Tunai
Tanggal Pembayaran Dividen
07 September 2026
03 September 2026
04 September 2026
07 September 2026
08 September 2026
18 September 2026
""",
        )
        schedule_pdf = document.tobytes()
        document.close()

        class Response:
            def __init__(self, status_code, payload=None, content=b""):
                self.status_code = status_code
                self._payload = payload or {}
                self.content = content

            def json(self):
                return self._payload

        class Session:
            def get(self, url, *, impersonate, **_kwargs):
                if url.endswith("/en"):
                    return Response(200)
                if url.endswith("GetApiDataPaginated"):
                    return Response(200, {"data": []})
                if url.endswith("GetAllAnnouncement"):
                    return Response(
                        200,
                        {
                            "Items": [
                                {
                                    "Code": "CMRY",
                                    "Title": "Schedule of Corporate Action",
                                    "Attachments": [
                                        {
                                            "FullSavePath": "https://example.test/cmry-schedule.pdf",
                                            "OriginalFilename": "CMRY_Jadwal Dividen Interim.pdf",
                                        }
                                    ],
                                }
                            ]
                        },
                    )
                if url == "https://example.test/cmry-schedule.pdf":
                    return Response(200, content=schedule_pdf)
                raise AssertionError(f"unexpected URL: {url}")

        events = fetch_cash_events(session_factory=Session)

        self.assertEqual(
            events,
            [
                {
                    "code": "CMRY",
                    "cashDividend": "100",
                    "exDividend": "2026-09-04",
                    "paymentDate": "2026-09-18",
                    "fiscalYear": "2026",
                    "note": "I",
                }
            ],
        )

    def test_fetch_discovers_dividend_schedule_on_later_announcement_page(self):
        import fitz

        document = fitz.open()
        page = document.new_page()
        page.insert_text(
            (72, 72),
            """Dividen Tunai Interim
Perseroan menyampaikan rencana pembagian Dividen Interim untuk periode tahun buku 2026
Dividen Per Saham
IDR
66
Jadwal pembagian dividen:
Tanggal Daftar Pemegang Saham (DPS) yang berhak atas dividen tunai
Tanggal Cum Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Ex Dividen di Pasar Reguler dan Pasar Negosiasi
Tanggal Cum Dividen di Pasar Tunai
Tanggal Ex Dividen di Pasar Tunai
Tanggal Pembayaran Dividen
17 September 2026
15 September 2026
16 September 2026
17 September 2026
18 September 2026
02 October 2026
""",
        )
        schedule_pdf = document.tobytes()
        document.close()

        class Response:
            def __init__(self, status_code, payload=None, content=b""):
                self.status_code = status_code
                self._payload = payload or {}
                self.content = content

            def json(self):
                return self._payload

        class Session:
            def get(self, url, *, impersonate, params=None, **_kwargs):
                if url.endswith("/en"):
                    return Response(200)
                if url.endswith("GetApiDataPaginated"):
                    return Response(200, {"data": []})
                if url.endswith("GetAllAnnouncement"):
                    if params["pageNumber"] == 1:
                        return Response(200, {"Items": [], "PageCount": 2})
                    if params["pageNumber"] == 2:
                        return Response(
                            200,
                            {
                                "Items": [
                                    {
                                        "Code": "BMRI",
                                        "Title": "Schedule of Corporate Action",
                                        "Attachments": [
                                            {
                                                "FullSavePath": "https://example.test/bmri-schedule.pdf",
                                                "OriginalFilename": "BMRI_Dividend Schedule.pdf",
                                            }
                                        ],
                                    }
                                ],
                                "PageCount": 2,
                            },
                        )
                if url == "https://example.test/bmri-schedule.pdf":
                    return Response(200, content=schedule_pdf)
                raise AssertionError(f"unexpected URL: {url}")

        events = fetch_cash_events(session_factory=Session)

        self.assertEqual(
            events,
            [
                {
                    "code": "BMRI",
                    "cashDividend": "66",
                    "exDividend": "2026-09-16",
                    "paymentDate": "2026-10-02",
                    "fiscalYear": "2026",
                    "note": "I",
                }
            ],
        )

    def test_fetch_retries_with_safari_when_chrome_homepage_is_rejected(self):
        class Response:
            def __init__(self, status_code, payload=None):
                self.status_code = status_code
                self._payload = payload or {}

            def json(self):
                return self._payload

        class Session:
            def __init__(self):
                self.calls = []

            def get(self, url, *, impersonate, **_kwargs):
                self.calls.append((url, impersonate))
                if url.endswith("/en") and impersonate == "chrome124":
                    return Response(503)
                if url.endswith("/en") and impersonate == "safari15_5":
                    return Response(200)
                return Response(200, {"data": [{"code": "TEST"}]})

        session = Session()

        events = fetch_cash_events(session_factory=lambda: session)

        self.assertEqual(events, [{"code": "TEST"}])
        self.assertEqual(
            [impersonate for _url, impersonate in session.calls][:3],
            ["chrome124", "safari15_5", "safari15_5"],
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
