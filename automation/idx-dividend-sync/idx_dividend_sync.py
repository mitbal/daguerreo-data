#!/usr/bin/env python3
"""Synchronize IDX cash-dividend announcements into JKSE per-ticker CSVs.

The IDX ``LINK_DIVIDEND`` dataset provides cash dividends and their scheduled
cum/ex/payment dates. This updater writes only events with a valid ex-dividend
date and a positive cash amount; stock dividends are deliberately excluded
because the repository's CSV schema stores IDR per share, not share amounts.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import re
import sys
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Callable

CSV_FIELDS = ["ex_date", "dividend", "payment_date", "fiscal_year", "dividend_type"]
NOTE_TYPES = {"F": "final", "I": "interim", "SD": "stock", "BS": "bonus"}
TICKER_PATTERN = re.compile(r"[A-Z0-9]{1,10}")
IDX_HOME = "https://www.idx.co.id/en"
CASH_API_URL = "https://www.idx.co.id/primary/DigitalStatistic/GetApiDataPaginated"
ANNOUNCEMENT_API_URL = "https://www.idx.co.id/primary/NewsAnnouncement/GetAllAnnouncement"
ANNOUNCEMENT_LOOKBACK_DAYS = 7
DATE_PATTERN = re.compile(r"\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b")
FISCAL_YEAR_PATTERN = re.compile(
    r"(?:tahun\s+buku|fiscal(?:\s+year)?|financial\s+year)\s+(\d{4})",
    flags=re.IGNORECASE,
)
MONTHS = {
    "january": 1, "januari": 1, "february": 2, "februari": 2,
    "march": 3, "maret": 3, "april": 4, "may": 5, "mei": 5,
    "june": 6, "juni": 6, "july": 7, "juli": 7, "august": 8,
    "agustus": 8, "september": 9, "october": 10, "oktober": 10,
    "november": 11, "december": 12, "desember": 12,
}
DEFAULT_REPO = Path(__file__).resolve().parents[2]


def _event_row(event: dict[str, Any]) -> tuple[str, dict[str, str]] | None:
    raw_ticker = event.get("code")
    ex_date = event.get("exDividend")
    if not isinstance(raw_ticker, str) or not isinstance(ex_date, str):
        return None
    ticker = raw_ticker.upper()
    if not TICKER_PATTERN.fullmatch(ticker):
        return None
    try:
        date.fromisoformat(ex_date)
        amount = Decimal(str(event.get("cashDividend")))
    except (InvalidOperation, ValueError):
        return None
    if not amount.is_finite() or amount <= 0:
        return None

    raw_fiscal_year = str(event.get("fiscalYear") or "").strip()
    fiscal_year = raw_fiscal_year if re.fullmatch(r"\d{4}", raw_fiscal_year) else ""
    note = str(event.get("note") or "").upper()
    return ticker, {
        "ex_date": ex_date,
        "dividend": format(amount, "f"),
        "payment_date": str(event.get("paymentDate") or ""),
        "fiscal_year": fiscal_year,
        "dividend_type": NOTE_TYPES.get(note, "other"),
    }


def _read_rows(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="") as handle:
        return list(csv.DictReader(handle))


def _write_rows(path: Path, rows: list[dict[str, str]]) -> None:
    ordered = sorted(
        rows,
        key=lambda row: (row.get("ex_date") or "", row.get("dividend_type") or ""),
        reverse=True,
    )
    temporary = path.with_suffix(".tmp")
    with temporary.open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(ordered)
    temporary.replace(path)


def _date_after_label(text: str, labels: tuple[str, ...]) -> str | None:
    """Return the first IDX-formatted date after one of the supplied labels."""
    lower_text = text.lower()
    for label in labels:
        position = lower_text.find(label.lower())
        if position < 0:
            continue
        match = DATE_PATTERN.search(text, position, position + 500)
        if match is None:
            continue
        day, month_name, year = match.groups()
        month = MONTHS.get(month_name.lower())
        if month is None:
            continue
        try:
            return date(int(year), month, int(day)).isoformat()
        except ValueError:
            continue
    return None


def _schedule_distribution_dates(text: str) -> list[str] | None:
    """Read the six table values following an IDX dividend-schedule heading."""
    lower_text = text.lower()
    for heading in ("jadwal pembagian dividen", "dividend distribution schedule"):
        start = lower_text.find(heading)
        if start < 0:
            continue
        end = len(text)
        for marker in ("data keuangan", "the underlying financial data"):
            marker_position = lower_text.find(marker, start)
            if marker_position >= 0:
                end = min(end, marker_position)
        dates: list[str] = []
        for match in DATE_PATTERN.finditer(text, start, end):
            day, month_name, year = match.groups()
            month = MONTHS.get(month_name.lower())
            if month is None:
                continue
            try:
                dates.append(date(int(year), month, int(day)).isoformat())
            except ValueError:
                continue
        if len(dates) >= 6:
            return dates[:6]
    return None


def _decimal_from_idx_text(value: str) -> Decimal | None:
    compact = value.strip().replace(" ", "")
    if not compact:
        return None
    if "," in compact and "." in compact:
        decimal_separator = "," if compact.rfind(",") > compact.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        compact = compact.replace(thousands_separator, "").replace(decimal_separator, ".")
    elif "," in compact:
        suffix = compact.rsplit(",", 1)[1]
        compact = compact.replace(",", ".") if len(suffix) <= 2 else compact.replace(",", "")
    elif "." in compact and len(compact.rsplit(".", 1)[1]) == 3:
        compact = compact.replace(".", "")
    try:
        result = Decimal(compact)
    except InvalidOperation:
        return None
    return result if result.is_finite() and result > 0 else None


def _dividend_per_share_after_label(text: str) -> Decimal | None:
    """Extract the per-share figure despite PDF table reading-order artifacts."""
    labels = ("dividen per saham", "dividend per share")
    lower_text = text.lower()
    stop_markers = (
        "jadwal pembagian dividen",
        "dividend distribution schedule",
        "data keuangan",
        "the underlying financial data",
        "hormat kami",
        "respectfully",
    )
    for label in labels:
        position = lower_text.find(label)
        if position < 0:
            continue
        fragment = text[position:position + 500]
        fragment_lower = fragment.lower()
        stop_positions = [
            marker_position
            for marker in stop_markers
            if (marker_position := fragment_lower.find(marker)) > 0
        ]
        if stop_positions:
            fragment = fragment[:min(stop_positions)]

        # IDX's newer schedule template states the confirmed per-share amount
        # before a separate section containing the total payout. In that layout
        # the later total must never overwrite the explicitly confirmed amount.
        confirmation_positions = [
            position for marker in (
                "dividen per saham sudah ditentukan",
                "dividend per share has been determined",
            )
            if (position := fragment.lower().find(marker)) > 0
        ]
        if confirmation_positions:
            confirmed_amounts: list[Decimal] = []
            for line in fragment[:min(confirmation_positions)].splitlines():
                money_matches = re.findall(
                    r"(?:IDR|Rp\.?)\s*([0-9][0-9.,]*)",
                    line,
                    flags=re.IGNORECASE,
                )
                raw_values = money_matches or re.findall(r"^\s*([0-9][0-9.,]*)\s*$", line)
                for raw_value in raw_values:
                    amount = _decimal_from_idx_text(raw_value)
                    if amount is not None:
                        confirmed_amounts.append(amount)
            if confirmed_amounts:
                return confirmed_amounts[-1]

        # In some IDX PDFs, PyMuPDF emits a total payout and the per-share
        # value after the same label: IDR, IDR, total payout, per-share value.
        # Consume the nearby amount-bearing lines in source order and take the
        # final value, which is the per-share cell in that table layout.
        amounts: list[Decimal] = []
        for line in fragment.splitlines():
            money_matches = re.findall(
                r"(?:IDR|Rp\.?)\s*([0-9][0-9.,]*)",
                line,
                flags=re.IGNORECASE,
            )
            raw_values = money_matches or re.findall(r"^\s*([0-9][0-9.,]*)\s*$", line)
            for raw_value in raw_values:
                amount = _decimal_from_idx_text(raw_value)
                if amount is not None:
                    amounts.append(amount)
        if amounts:
            return amounts[-1]
    return None


def _event_from_dividend_schedule(ticker: str, schedule_text: str) -> dict[str, str] | None:
    normalized_text = schedule_text.lower()
    if "stock dividend" in normalized_text or "dividen saham" in normalized_text:
        return None
    if "interim" in normalized_text:
        note = "I"
    elif "final" in normalized_text:
        note = "F"
    else:
        return None

    amount = _dividend_per_share_after_label(schedule_text)
    schedule_dates = _schedule_distribution_dates(schedule_text)
    if schedule_dates is not None:
        ex_date = schedule_dates[2]
        payment_date = schedule_dates[5]
    else:
        ex_date = _date_after_label(
            schedule_text,
            (
                "tanggal ex dividen di pasar reguler dan pasar negosiasi",
                "dividend ex date in regular market and negotiation market",
            ),
        )
        payment_date = _date_after_label(
            schedule_text,
            ("tanggal pembayaran dividen", "dividend payment date"),
        )
    if amount is None or ex_date is None or payment_date is None:
        return None
    fiscal_year_match = FISCAL_YEAR_PATTERN.search(schedule_text)
    event = {
        "code": ticker,
        "cashDividend": format(amount, "f"),
        "exDividend": ex_date,
        "paymentDate": payment_date,
        "note": note,
    }
    if fiscal_year_match is not None:
        event["fiscalYear"] = fiscal_year_match.group(1)
    return event


def _announcement_dividend_events(session: Any, impersonate: str) -> list[dict[str, str]]:
    """Extract recent cash-dividend schedules from official IDX disclosures."""
    today = date.today()
    params = {
        "keywords": "",
        "pageSize": 1000,
        "dateFrom": (today - timedelta(days=ANNOUNCEMENT_LOOKBACK_DAYS)).strftime("%Y%m%d"),
        "dateTo": today.strftime("%Y%m%d"),
        "lang": "en",
    }
    response = session.get(
        ANNOUNCEMENT_API_URL,
        params={**params, "pageNumber": 1},
        impersonate=impersonate,
        timeout=30,
    )
    if response.status_code != 200:
        raise RuntimeError(f"announcement API HTTP {response.status_code}")
    first_page = response.json()
    announcements = first_page.get("Items", [])
    if not isinstance(announcements, list):
        raise RuntimeError("announcement API returned a non-list Items field")

    # The API caps each response at 1,000 rows. Dividend notices can land on a
    # later page during busy disclosure windows, so inspect every page in the
    # bounded rolling lookback instead of silently treating page one as complete.
    for page_number in range(2, int(first_page.get("PageCount") or 1) + 1):
        response = session.get(
            ANNOUNCEMENT_API_URL,
            params={**params, "pageNumber": page_number},
            impersonate=impersonate,
            timeout=30,
        )
        if response.status_code != 200:
            raise RuntimeError(f"announcement API page {page_number} HTTP {response.status_code}")
        page_items = response.json().get("Items", [])
        if not isinstance(page_items, list):
            raise RuntimeError(f"announcement API page {page_number} returned a non-list Items field")
        announcements.extend(page_items)

    import fitz

    events: list[dict[str, str]] = []
    for announcement in announcements:
        ticker = str(announcement.get("Code") or "").strip().upper()
        attachments = announcement.get("Attachments") or []
        attachment_names = " ".join(
            str(attachment.get("OriginalFilename") or "") for attachment in attachments
        ).lower()
        title = str(announcement.get("Title") or "").lower()
        if not TICKER_PATTERN.fullmatch(ticker) or ("dividen" not in title and "dividend" not in title
                                                  and "dividen" not in attachment_names and "dividend" not in attachment_names):
            continue
        for attachment in sorted(attachments, key=lambda item: item.get("IsAttachment", 0)):
            file_url = attachment.get("FullSavePath")
            if not isinstance(file_url, str) or not file_url:
                continue
            pdf_response = session.get(file_url, impersonate=impersonate, timeout=120)
            if pdf_response.status_code != 200 or not pdf_response.content:
                continue
            try:
                document = fitz.open(stream=pdf_response.content, filetype="pdf")
                schedule_text = "\n".join(page.get_text() for page in document)
                document.close()
            except Exception:
                continue
            event = _event_from_dividend_schedule(ticker, schedule_text)
            if event is not None:
                events.append(event)
                break
    return events


def _deduplicate_events(events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    unique_events: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for event in events:
        key = (
            str(event.get("code") or "").upper(),
            str(event.get("exDividend") or ""),
            str(event.get("note") or "").upper(),
        )
        if key in seen:
            continue
        seen.add(key)
        unique_events.append(event)
    return unique_events


def sync_cash_events(repo: Path, events: list[dict[str, Any]]) -> dict[str, Any]:
    """Upsert supplied IDX cash-dividend events into ticker CSVs."""
    dividend_dir = repo / "jkse" / "dividends"
    dividend_dir.mkdir(parents=True, exist_ok=True)
    files: list[str] = []
    added = updated = skipped = 0

    for event in events:
        normalized = _event_row(event)
        if normalized is None:
            skipped += 1
            continue
        ticker, incoming = normalized
        target = dividend_dir / f"{ticker}.csv"
        rows = _read_rows(target)
        matching_index = next(
            (
                index
                for index, row in enumerate(rows)
                if row.get("ex_date") == incoming["ex_date"]
                and row.get("dividend_type") == incoming["dividend_type"]
            ),
            None,
        )

        if matching_index is None:
            rows.append(incoming)
            _write_rows(target, rows)
            added += 1
            files.append(target.name)
            continue

        existing = rows[matching_index]
        candidate = incoming.copy()
        # Keep a previously sourced fiscal year, otherwise backfill one from a new event.
        candidate["fiscal_year"] = existing.get("fiscal_year") or incoming.get("fiscal_year") or ""
        if existing == candidate:
            skipped += 1
            continue

        rows[matching_index] = candidate
        _write_rows(target, rows)
        updated += 1
        files.append(target.name)

    return {"added": added, "updated": updated, "skipped": skipped, "files": files}


def fetch_cash_events(session_factory: Callable[[], Any] | None = None) -> list[dict[str, Any]]:
    """Fetch cash dividends from IDX statistics and recent official disclosures."""
    if session_factory is None:
        from curl_cffi import requests

        session_factory = requests.Session

    query = base64.b64encode(json.dumps({"year": date.today().year}).encode()).decode()
    failures: list[str] = []
    for impersonate in ("chrome124", "safari15_5"):
        session = session_factory()
        try:
            home = session.get(IDX_HOME, impersonate=impersonate, timeout=30)
            if home.status_code != 200:
                failures.append(f"{impersonate} homepage HTTP {home.status_code}")
                continue
            response = session.get(
                CASH_API_URL,
                params={
                    "urlName": "LINK_DIVIDEND",
                    "query": query,
                    "isPrint": False,
                    "cumulative": False,
                    "pageSize": 1500,
                },
                impersonate=impersonate,
                timeout=30,
            )
            if response.status_code != 200:
                failures.append(f"{impersonate} cash API HTTP {response.status_code}")
                continue
            records = response.json().get("data", [])
            if not isinstance(records, list):
                failures.append(f"{impersonate} cash API returned a non-list data field")
                continue
            announcement_events = _announcement_dividend_events(session, impersonate)
            return _deduplicate_events([*announcement_events, *records])
        except Exception as exc:
            failures.append(f"{impersonate} {type(exc).__name__}: {exc}")
    raise RuntimeError("IDX cash-dividend fetch failed; " + "; ".join(failures))


def execute(repo: Path, fetcher: Callable[[], list[dict[str, Any]]]) -> dict[str, Any]:
    """Fetch IDX records and synchronize them, returning audit counts."""
    events = list(fetcher())
    result = sync_cash_events(repo, events)
    return {"source_rows": len(events), **result}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=DEFAULT_REPO, help="daguerreo-data repository path")
    parser.add_argument("--verbose", action="store_true", help="print a successful no-change check")
    args = parser.parse_args(argv)

    try:
        result = execute(args.repo.resolve(), fetch_cash_events)
    except Exception as exc:
        print(f"IDX cash-dividend sync failed: {exc}", file=sys.stderr)
        return 1

    changed = result["added"] + result["updated"]
    if changed:
        print(
            "IDX cash-dividend sync: "
            f"{result['source_rows']} source rows checked; "
            f"{result['added']} added, {result['updated']} updated."
        )
        for filename in result["files"]:
            print(f"  - jkse/dividends/{filename}")
    elif args.verbose:
        print(
            "IDX cash-dividend sync: "
            f"{result['source_rows']} source rows checked; no CSV changes "
            f"({result['skipped']} unchanged or invalid)."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
