#!/usr/bin/env python3
"""Brokerage data integration module for Fidelity, Robinhood, SoFi, and SnapTrade.

Provides:
1. Automated CSV parsing for:
   - Fidelity (YTD Tax Information, Closed Positions, Realized Gain/Loss)
   - Robinhood (1099-B, Account Activities, Reports)
   - SoFi (Realized Gain/Loss, Account Transactions)
   - Generic 1099-B / 1099-DIV / 1099-INT Tax Summaries
2. SnapTrade API Client:
   - Retail brokerage OAuth integration (Fidelity, Robinhood, SoFi)
   - Normalized investment activities / tax data retrieval
3. Clean data structures for feeding short/long capital gains, qualified dividends,
   and interest directly into the NYC Tax Planner.
"""
from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import io
import json
import os
import re
import time
import urllib.parse
import urllib.request
import urllib.error
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, TextIO


@dataclass
class BrokerageSummary:
    short_gains: float = 0.0
    long_gains: float = 0.0
    qualified_dividends: float = 0.0
    ordinary_dividends: float = 0.0
    interest: float = 0.0
    source: str = "Unknown"
    record_count: int = 0
    accounts: list[str] = field(default_factory=list)
    details: list[str] = field(default_factory=list)

    def add(self, other: BrokerageSummary) -> BrokerageSummary:
        """Combine two brokerage summaries (e.g. Fidelity + Robinhood)."""
        combined = BrokerageSummary(
            short_gains=round(self.short_gains + other.short_gains, 2),
            long_gains=round(self.long_gains + other.long_gains, 2),
            qualified_dividends=round(self.qualified_dividends + other.qualified_dividends, 2),
            ordinary_dividends=round(self.ordinary_dividends + other.ordinary_dividends, 2),
            interest=round(self.interest + other.interest, 2),
            source=f"{self.source} + {other.source}" if self.source != other.source else self.source,
            record_count=self.record_count + other.record_count,
            accounts=list(set(self.accounts + other.accounts)),
            details=self.details + other.details
        )
        return combined

    def summary_text(self) -> str:
        lines = [
            f"Source: {self.source} ({self.record_count} items processed)",
            f"  • Short-Term Capital Gains: ${self.short_gains:,.2f}",
            f"  • Long-Term Capital Gains:  ${self.long_gains:,.2f}",
            f"  • Qualified Dividends:      ${self.qualified_dividends:,.2f}",
            f"  • Nonqualified / Ordinary:  ${self.ordinary_dividends:,.2f}",
            f"  • Taxable Interest:         ${self.interest:,.2f}",
        ]
        if self.accounts:
            lines.append(f"  • Accounts: {', '.join(self.accounts)}")
        return "\n".join(lines)


def parse_currency(val: Any) -> float:
    """Safely convert strings with $, commas, accounting parentheses, or trailing minus to float."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.lower() in ('--', 'n/a', 'none', '-', 'null', ''):
        return 0.0
    is_negative = False
    if s.startswith('(') and s.endswith(')'):
        is_negative = True
        s = s[1:-1].strip()
    elif s.endswith('-'):
        is_negative = True
        s = s[:-1].strip()
    elif s.startswith('-'):
        is_negative = True
        s = s[1:].strip()
    elif s.startswith('+$') or s.startswith('+'):
        s = s.replace('+', '').strip()

    s = s.replace('$', '').replace(',', '').replace('%', '').strip()
    if not s:
        return 0.0
    try:
        num = float(s)
        return -num if is_negative else num
    except ValueError:
        return 0.0


def parse_brokerage_csv(filepath_or_data: str, filename: str | None = None) -> BrokerageSummary:
    """Parse a CSV file or CSV string from Fidelity, Robinhood, SoFi, or generic 1099 format.
    
    Returns a populated BrokerageSummary dataclass.
    """
    if os.path.exists(filepath_or_data):
        with open(filepath_or_data, 'r', encoding='utf-8-sig', errors='replace') as f:
            raw_text = f.read()
        file_label = os.path.basename(filepath_or_data)
    else:
        raw_text = filepath_or_data
        file_label = filename or "CSV Data"

    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    if not lines:
        return BrokerageSummary(source=f"Empty file ({file_label})")

    # Detect broker brand hint from content or file name
    text_lower = raw_text.lower()
    broker_hint = "Generic"
    if "fidelity" in text_lower or "active trader" in text_lower:
        broker_hint = "Fidelity"
    elif "robinhood" in text_lower:
        broker_hint = "Robinhood"
    elif "sofi" in text_lower:
        broker_hint = "SoFi"

    # Check for Key-Value / Line-Summary 1099 Format
    summary_match = _try_parse_summary_lines(lines, broker_hint)
    if summary_match and summary_match.record_count > 0:
        summary_match.source = f"{broker_hint} Summary ({file_label})"
        return summary_match

    # Find the header row (skip top metadata/disclaimer lines if any)
    header_idx = -1
    best_score = 0
    keywords = [
        'symbol', 'description', 'proceeds', 'cost', 'gain', 'loss',
        'term', 'date', 'acquired', 'sold', 'amount', 'trans code', 'type', 'dividend', 'interest'
    ]

    reader = csv.reader(lines)
    all_rows = list(reader)

    for idx, row in enumerate(all_rows[:15]):
        row_str = " ".join(c.lower().strip() for c in row)
        score = sum(1 for kw in keywords if kw in row_str)
        if score > best_score:
            best_score = score
            header_idx = idx

    if header_idx == -1 or best_score < 2:
        # Fallback to first line if no header strongly detected
        header_idx = 0

    headers = [str(c).strip() for c in all_rows[header_idx]]
    data_rows = all_rows[header_idx + 1:]

    col_map = {h.lower().replace('_', ' ').replace('-', ' ').strip(): i for i, h in enumerate(headers)}

    # Helper to find column index from aliases
    def find_col(*aliases: str) -> int | None:
        for alias in aliases:
            a_clean = alias.lower().replace('_', ' ').replace('-', ' ').strip()
            if a_clean in col_map:
                return col_map[a_clean]
            for col_name, idx in col_map.items():
                if a_clean in col_name:
                    return idx
        return None

    short_gains = 0.0
    long_gains = 0.0
    qual_div = 0.0
    ord_div = 0.0
    interest = 0.0
    records = 0
    accounts = set()
    details = []

    # Column mappings
    term_col = find_col('term', 'holding period', 'holding', 'short/long')
    gain_loss_col = find_col('realized gain/loss', 'gain/loss', 'total gain/loss', 'gain loss ($)', 'gain', 'loss')
    proceeds_col = find_col('proceeds', 'sales proceeds', 'total proceeds')
    cost_col = find_col('cost basis', 'cost', 'total cost')
    wash_col = find_col('wash sale', 'wash sale adjustment', 'disallowed')
    account_col = find_col('account', 'account number', 'account name')
    type_col = find_col('trans code', 'type', 'action', 'activity type', 'transaction type')
    amount_col = find_col('amount', 'net amount', 'total ($)', 'total')
    desc_col = find_col('description', 'security description', 'name')
    date_sold_col = find_col('date sold', 'close date', 'disposal date', 'settle date', 'process date', 'activity date', 'date')
    date_acq_col = find_col('date acquired', 'open date', 'acquisition date')

    for row_idx, row in enumerate(data_rows, start=header_idx + 2):
        if not row or all(not str(c).strip() for c in row):
            continue

        # Check for footer / summary rows
        first_cell = str(row[0]).strip().lower()
        if first_cell.startswith('total') or first_cell.startswith('disclaimer') or 'the data provided' in first_cell:
            continue

        if account_col is not None and account_col < len(row):
            acc_val = str(row[account_col]).strip()
            if acc_val:
                accounts.add(acc_val)

        row_type = str(row[type_col]).strip().upper() if type_col is not None and type_col < len(row) else ""
        row_desc = str(row[desc_col]).strip() if desc_col is not None and desc_col < len(row) else ""
        row_term = str(row[term_col]).strip().lower() if term_col is not None and term_col < len(row) else ""

        # Check for Dividends
        is_div = False
        is_qual_div = False
        term_lower = row_term.lower()
        desc_lower = row_desc.lower()
        type_lower = row_type.lower()

        if any(code in row_type for code in ('DIV', 'CDIV', 'MDIV', 'QUAL')) or 'dividend' in desc_lower or 'dividend' in term_lower or term_lower.startswith('div'):
            is_div = True
            if 'non-qual' in type_lower or 'nonqual' in type_lower or 'ordinary' in type_lower or 'ordinary' in desc_lower:
                is_qual_div = False
            else:
                is_qual_div = True

        # Check for Interest
        is_int = False
        if any(code in row_type for code in ('INT', 'INTEREST')) or 'interest' in desc_lower or 'interest' in term_lower or term_lower.startswith('int'):
            is_int = True

        if is_div or is_int:
            amt = 0.0
            if amount_col is not None and amount_col < len(row):
                amt = parse_currency(row[amount_col])
            elif gain_loss_col is not None and gain_loss_col < len(row) and str(row[gain_loss_col]).strip():
                amt = parse_currency(row[gain_loss_col])
            elif proceeds_col is not None and proceeds_col < len(row):
                amt = parse_currency(row[proceeds_col])

            if amt != 0:
                records += 1
                if is_int:
                    interest += amt
                elif is_qual_div:
                    qual_div += amt
                else:
                    ord_div += amt
            continue

        # Otherwise treat as Trade / Capital Gain
        row_gain = None
        if gain_loss_col is not None and gain_loss_col < len(row):
            val = parse_currency(row[gain_loss_col])
            # Only use if not obviously empty
            if str(row[gain_loss_col]).strip():
                row_gain = val

        if row_gain is None and proceeds_col is not None and cost_col is not None:
            if proceeds_col < len(row) and cost_col < len(row):
                p_str = str(row[proceeds_col]).strip()
                c_str = str(row[cost_col]).strip()
                if p_str and c_str:
                    p_val = parse_currency(p_str)
                    c_val = parse_currency(c_str)
                    w_val = parse_currency(row[wash_col]) if wash_col is not None and wash_col < len(row) else 0.0
                    row_gain = (p_val - c_val) + w_val

        if row_gain is not None:
            records += 1
            # Determine Short vs Long Term
            is_long = False
            if 'long' in row_term or row_term in ('lt', 'l'):
                is_long = True
            elif 'short' in row_term or row_term in ('st', 's'):
                is_long = False
            elif date_sold_col is not None and date_acq_col is not None:
                # Try date diff
                ds = str(row[date_sold_col]).strip() if date_sold_col < len(row) else ""
                da = str(row[date_acq_col]).strip() if date_acq_col < len(row) else ""
                is_long = _is_holding_period_long(da, ds)

            if is_long:
                long_gains += row_gain
            else:
                short_gains += row_gain

    # Formulate detected source label
    source_label = f"{broker_hint} CSV ({file_label})"
    if records == 0:
        details.append("No realized capital gain or dividend/interest transactions detected in table.")
    else:
        details.append(f"Successfully processed {records} items from {file_label}.")

    return BrokerageSummary(
        short_gains=round(short_gains, 2),
        long_gains=round(long_gains, 2),
        qualified_dividends=round(qual_div, 2),
        ordinary_dividends=round(ord_div, 2),
        interest=round(interest, 2),
        source=source_label,
        record_count=records,
        accounts=sorted(list(accounts)),
        details=details
    )


def _is_holding_period_long(date_acq_str: str, date_sold_str: str) -> bool:
    """Check if holding period is more than 365 days."""
    formats = ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%d-%b-%Y', '%b %d, %Y')
    da = None
    ds = None
    for fmt in formats:
        if not da:
            try: da = datetime.strptime(date_acq_str.strip(), fmt)
            except Exception: pass
        if not ds:
            try: ds = datetime.strptime(date_sold_str.strip(), fmt)
            except Exception: pass
    if da and ds:
        return (ds - da).days > 365
    return False


def _try_parse_summary_lines(lines: list[str], broker_hint: str) -> BrokerageSummary | None:
    """Check for 2-column or free-form 1099 summary lines (e.g. 'Qualified dividends: $500.00')."""
    short_g = 0.0
    long_g = 0.0
    qual_d = 0.0
    ord_d = 0.0
    inte = 0.0
    found = 0

    for line in lines:
        # Check standard 1099 labels
        l_low = line.lower()
        if ',' in line:
            parts = [p.strip() for p in line.split(',') if p.strip()]
        elif ':' in line:
            parts = [p.strip() for p in line.split(':') if p.strip()]
        else:
            continue

        if len(parts) >= 2:
            label = parts[0].lower()
            val_str = parts[1]
            if any(term in label for term in ('short term', 'stcg', 'box 1a short', 'short gain')):
                short_g += parse_currency(val_str)
                found += 1
            elif any(term in label for term in ('long term', 'ltcg', 'box 1a long', 'long gain')):
                long_g += parse_currency(val_str)
                found += 1
            elif 'qualified dividend' in label or 'box 1b' in label:
                qual_d += parse_currency(val_str)
                found += 1
            elif any(term in label for term in ('ordinary dividend', 'total dividend', 'box 1a dividend')):
                ord_d += parse_currency(val_str)
                found += 1
            elif any(term in label for term in ('taxable interest', 'interest income', 'box 1 interest', '1099-int')):
                inte += parse_currency(val_str)
                found += 1

    if found >= 2:
        return BrokerageSummary(
            short_gains=round(short_g, 2),
            long_gains=round(long_g, 2),
            qualified_dividends=round(qual_d, 2),
            ordinary_dividends=round(ord_d, 2),
            interest=round(inte, 2),
            source=f"{broker_hint} Tax Summary",
            record_count=found
        )
    return None


class SnapTradeClient:
    """Pure Standard-Library SnapTrade API v1 Client.
    
    Supports OAuth connection to Fidelity, Robinhood, SoFi, Schwab, etc.
    Extracts activities, dividends, interest, and capital gains.
    """
    BASE_URL = "https://api.snaptrade.com/api/v1"

    def __init__(
        self,
        client_id: str,
        consumer_key: str,
        user_id: str | None = None,
        user_secret: str | None = None
    ):
        self.client_id = client_id.strip()
        self.consumer_key = consumer_key.strip()
        self.user_id = user_id.strip() if user_id else None
        self.user_secret = user_secret.strip() if user_secret else None

    def is_configured(self) -> bool:
        return bool(self.client_id and self.consumer_key)

    def _generate_signature(self, path: str, query_params: dict[str, Any], body: dict[str, Any] | None = None) -> str:
        """Construct HMAC-SHA256 Base64-encoded signature required by SnapTrade API."""
        signed_content = {
            "path": path,
            "query": query_params,
            "content": body if body is not None else None
        }
        content_str = json.dumps(signed_content, sort_keys=True, separators=(",", ":"))
        hmac_obj = hmac.new(
            self.consumer_key.encode("utf-8"),
            content_str.encode("utf-8"),
            hashlib.sha256
        )
        return base64.b64encode(hmac_obj.digest()).decode("utf-8")

    def request(
        self,
        method: str,
        endpoint: str,
        query_params: dict[str, Any] | None = None,
        body: dict[str, Any] | None = None,
        timeout: int = 20
    ) -> dict[str, Any] | list[Any]:
        """Send authenticated HTTP request to SnapTrade API using standard library urllib."""
        if not self.is_configured():
            raise ValueError("SnapTrade Client ID and Consumer Key are required.")

        params = dict(query_params or {})
        params["clientId"] = self.client_id
        params["timestamp"] = int(time.time())

        # If user credentials exist, append to query parameters for GET or query signing
        if self.user_id and "userId" not in params:
            params["userId"] = self.user_id
        if self.user_secret and "userSecret" not in params:
            params["userSecret"] = self.user_secret

        sig = self._generate_signature(endpoint, params, body)

        encoded_query = urllib.parse.urlencode(params)
        full_url = f"{self.BASE_URL}{endpoint}?{encoded_query}"

        data_bytes = None
        if body is not None:
            data_bytes = json.dumps(body).encode("utf-8")

        req = urllib.request.Request(
            full_url,
            data=data_bytes,
            method=method.upper()
        )
        req.add_header("Signature", sig)
        req.add_header("Content-Type", "application/json")
        req.add_header("Accept", "application/json")
        req.add_header("User-Agent", "NYC-Tax-Planner-2026/1.0")

        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                status_code = resp.getcode()
                raw_body = resp.read().decode("utf-8")
                if not raw_body:
                    return {}
                return json.loads(raw_body)
        except urllib.error.HTTPError as e:
            err_msg = e.read().decode("utf-8", errors="replace")
            try:
                err_json = json.loads(err_msg)
                detail = err_json.get("detail") or err_json.get("message") or err_msg
            except Exception:
                detail = err_msg
            raise RuntimeError(f"SnapTrade API Error ({e.code}): {detail}")
        except urllib.error.URLError as e:
            raise RuntimeError(f"SnapTrade Connection Error: {e.reason}")

    def check_status(self) -> dict[str, Any]:
        """Test API connectivity and credentials."""
        res = self.request("GET", "/status")
        return res if isinstance(res, dict) else {"status": str(res)}

    def register_user(self, user_id: str) -> str:
        """Register a new user in SnapTrade or retrieve secret."""
        body = {"userId": user_id}
        res = self.request("POST", "/snapTrade/registerUser", body=body)
        if isinstance(res, dict) and "userSecret" in res:
            self.user_id = user_id
            self.user_secret = res["userSecret"]
            return self.user_secret
        raise RuntimeError(f"Failed to register user in SnapTrade: {res}")

    def get_login_redirect_url(self, broker: str | None = None) -> str:
        """Generate an OAuth portal connection URL for the user to link their brokerage."""
        if not self.user_id or not self.user_secret:
            raise ValueError("SnapTrade user must be registered before generating a connection link.")

        body: dict[str, Any] = {"immediateRedirect": True}
        if broker:
            body["broker"] = broker.upper()

        res = self.request("POST", "/snapTrade/login", body=body)
        if isinstance(res, dict) and "redirectURI" in res:
            return res["redirectURI"]
        raise RuntimeError(f"Failed to get SnapTrade login URI: {res}")

    def list_accounts(self) -> list[dict[str, Any]]:
        """List all connected accounts (Fidelity, Robinhood, SoFi, etc.)."""
        if not self.user_id or not self.user_secret:
            return []
        res = self.request("GET", "/accounts")
        return res if isinstance(res, list) else []

    def get_activities(self, start_date: str | None = None, end_date: str | None = None) -> list[dict[str, Any]]:
        """Retrieve trade, dividend, and interest activities across connected accounts."""
        if not self.user_id or not self.user_secret:
            return []
        params = {}
        if start_date:
            params["startDate"] = start_date
        if end_date:
            params["endDate"] = end_date
        res = self.request("GET", "/activities", query_params=params)
        return res if isinstance(res, list) else []

    def fetch_tax_summary(self, target_year: int = 2026) -> BrokerageSummary:
        """Fetch all activities for the specified year and calculate tax totals."""
        start_date = f"{target_year}-01-01"
        end_date = f"{target_year}-12-31"
        activities = self.get_activities(start_date=start_date, end_date=end_date)
        accounts = self.list_accounts()

        account_names = []
        for acc in accounts:
            b_name = acc.get("brokerage_name") or acc.get("institution_name") or "Brokerage"
            a_num = acc.get("number") or acc.get("name") or ""
            account_names.append(f"{b_name} ({a_num})" if a_num else b_name)

        short_gains = 0.0
        long_gains = 0.0
        qual_div = 0.0
        ord_div = 0.0
        interest = 0.0
        count = 0

        for act in activities:
            a_type = str(act.get("type", "")).upper()
            amt = float(act.get("amount") or 0.0)
            trade_date = str(act.get("trade_date") or act.get("settlement_date") or "")
            if trade_date and not trade_date.startswith(str(target_year)):
                continue

            count += 1
            if a_type in ("DIVIDEND", "SUBSTITUTE_DIVIDEND"):
                # By default, brokerage dividends on common equities are qualified
                qual_div += amt
            elif a_type in ("INTEREST", "INTEREST_INCOME"):
                interest += amt
            elif a_type in ("REALIZED_GAIN_SHORT", "STCG"):
                short_gains += amt
            elif a_type in ("REALIZED_GAIN_LONG", "LTCG"):
                long_gains += amt

        source_str = "SnapTrade (" + ", ".join(account_names[:3]) + ")" if account_names else "SnapTrade API"
        return BrokerageSummary(
            short_gains=round(short_gains, 2),
            long_gains=round(long_gains, 2),
            qualified_dividends=round(qual_div, 2),
            ordinary_dividends=round(ord_div, 2),
            interest=round(interest, 2),
            source=source_str,
            record_count=count,
            accounts=account_names,
            details=[f"Fetched {count} activities across {len(accounts)} connected account(s) for tax year {target_year}."]
        )

    def save_config(self, filepath: str = "snaptrade_config.json") -> None:
        """Save credentials to local config file (git-ignored)."""
        data = {
            "client_id": self.client_id,
            "consumer_key": self.consumer_key,
            "user_id": self.user_id,
            "user_secret": self.user_secret
        }
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    @classmethod
    def load_config(cls, filepath: str = "snaptrade_config.json") -> SnapTradeClient | None:
        """Load credentials from local config file or environment variables."""
        # 1. Check environment variables
        c_id = os.environ.get("SNAPTRADE_CLIENT_ID")
        c_key = os.environ.get("SNAPTRADE_CONSUMER_KEY")
        u_id = os.environ.get("SNAPTRADE_USER_ID")
        u_sec = os.environ.get("SNAPTRADE_USER_SECRET")

        if c_id and c_key:
            return cls(c_id, c_key, u_id, u_sec)

        # 2. Check config file
        if os.path.exists(filepath):
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return cls(
                    client_id=data.get("client_id", ""),
                    consumer_key=data.get("consumer_key", ""),
                    user_id=data.get("user_id"),
                    user_secret=data.get("user_secret")
                )
            except Exception:
                pass
        return None
