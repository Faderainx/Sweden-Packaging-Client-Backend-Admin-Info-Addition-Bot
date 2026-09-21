from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from playwright.sync_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DEFAULT_CONFIG = APP_DIR / "config.json"
TARGET_ADMIN_EMAIL = "admin@example.com"
TARGET_INVOICE_EMAIL = "invoice@example.com"
MAIL_TIMESTAMP_TOLERANCE = timedelta(minutes=2)
DEFAULT_CODE_WAIT_SECONDS = 90


def normalize_email(value: str | None) -> str:
    return (value or "").strip().lower()


def _row_value(row: dict[str, Any], *names: str) -> str:
    normalized = {str(k).strip().lower(): str(v or "").strip() for k, v in row.items()}
    for name in names:
        value = normalized.get(name.lower(), "")
        if value:
            return value
    return ""


def load_rows(path: Path) -> list[dict[str, Any]]:
    """Read CSV or XLSX without logging credentials."""
    if not path.exists():
        raise FileNotFoundError(f"客户表格不存在：{path}")
    if path.suffix.lower() == ".xlsx":
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise RuntimeError("读取 XLSX 需要 openpyxl，请先运行启动脚本安装依赖。") from exc
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(x or "").strip() for x in rows[0]]
            return [dict(zip(headers, values)) for values in rows[1:] if any(values)]
        finally:
            workbook.close()
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def load_config(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    config.setdefault("portal_url", "https://portal-prod.npa.se/")
    config.setdefault("target_user", {"first_name": "Min", "last_name": "Liu", "email": TARGET_ADMIN_EMAIL, "permission": "Administrator"})
    config.setdefault("target_invoice_email", TARGET_INVOICE_EMAIL)
    config.setdefault("mail_routes", [])
    return config


def resolve_mail_route(email: str, routes: Iterable[dict[str, Any]]) -> dict[str, Any] | None:
    """Resolve exact domain patterns first, then conservative token aliases."""
    normalized = normalize_email(email)
    if "@" not in normalized:
        return None
    domain = normalized.rsplit("@", 1)[1].rstrip(".")
    exact_matches: list[tuple[int, dict[str, Any], str]] = []
    token_matches: list[tuple[dict[str, Any], str]] = []
    for route in routes:
        for raw_pattern in route.get("patterns", []):
            pattern = str(raw_pattern).strip().lower().lstrip("@").rstrip(".")
            if not pattern:
                continue
            if domain == pattern or domain.endswith("." + pattern):
                exact_matches.append((len(pattern), route, pattern))
        for raw_token in route.get("tokens", []):
            token = str(raw_token).strip().lower().lstrip("@").rstrip(".")
            if not token:
                continue
            labels = domain.split(".")
            if token in labels or any(label.endswith("-" + token) for label in labels):
                token_matches.append((route, token))
    if exact_matches:
        exact_matches.sort(key=lambda item: item[0], reverse=True)
        route = dict(exact_matches[0][1])
        route["matched_by"] = exact_matches[0][2]
        return route
    unique = {id(route): (route, token) for route, token in token_matches}
    if len(unique) == 1:
        route, token = next(iter(unique.values()))
        result = dict(route)
        result["matched_by"] = token
        result["ambiguous_alias"] = True
        return result
    return None


def extract_invoice_email(text: str) -> str | None:
    pattern = re.compile(r"(?:invoice\s*email|faktura\s*e-post|发票邮箱)\s*(?:[:：]|\s|\n)+([a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,})", re.I)
    match = pattern.search(text)
    return normalize_email(match.group(1)) if match else None


def extract_verification_code(text: str) -> str | None:
    """Extract a code only when it follows an explicit verification-code label."""
    patterns = [
        r"(?:账户验证码|帐户验证码|验证码|account\s+verification\s+code|verification\s+code)\s*[:：]?\s*(\d{6,8})",
        r"(?:one[- ]time\s+code|security\s+code)\s*[:：]?\s*(\d{6,8})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return match.group(1)
    return None


def _first_visible(locators: Iterable[Any]) -> Any | None:
    for locator in locators:
        try:
            count = locator.count()
            for index in range(count):
                candidate = locator.nth(index)
                if candidate.is_visible():
                    return candidate
        except Exception:
            continue
    return None


def _first_fillable(locators: Iterable[Any]) -> Any | None:
    for locator in locators:
        try:
            count = locator.count()
            for index in range(count):
                candidate = locator.nth(index)
                input_type = (candidate.get_attribute("type") or "").lower()
                if input_type in {"checkbox", "radio", "hidden", "submit", "button", "file"}:
                    continue
                if candidate.is_visible() and candidate.is_editable():
                    return candidate
        except Exception:
            continue
    return None


def _scopes(page: Page) -> list[Any]:
    """Return the page and currently loaded frames for provider pages such as Aliyun Mail."""
    return [page, *[frame for frame in page.frames if frame != page.main_frame]]


def _scoped_locators(page: Page, selectors: Iterable[str]) -> list[Any]:
    return [scope.locator(selector) for scope in _scopes(page) for selector in selectors]


def _wait_for_scoped(page: Page, selectors: Iterable[str], timeout_ms: int = 15000) -> Any | None:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        target = _first_visible(_scoped_locators(page, selectors))
        if target is not None:
            return target
        page.wait_for_timeout(250)
    return None


def _wait_for_fillable(page: Page, selectors: Iterable[str], timeout_ms: int = 15000) -> Any | None:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        target = _first_fillable(_scoped_locators(page, selectors))
        if target is not None:
            return target
        page.wait_for_timeout(250)
    return None


def click_first(page: Page, patterns: Iterable[str], *, exact: bool = False) -> bool:
    # Provider pages may have a navigation link such as Aliyun's
    # "个人邮箱登录" outside the embedded login form. Search visible
    # buttons across every scope first so a generic "登录" pattern cannot
    # accidentally click that navigation link instead of the submit button.
    for pattern in patterns:
        button_locators = []
        for scope in _scopes(page):
            button_locators.append(
                scope.get_by_role(
                    "button",
                    name=pattern if exact else re.compile(pattern, re.I),
                    exact=exact,
                )
            )
        target = _first_visible(button_locators)
        if target is not None:
            try:
                target.click()
                return True
            except Exception:
                continue
    for pattern in patterns:
        link_locators = []
        for scope in _scopes(page):
            link_locators.append(
                scope.get_by_role(
                    "link",
                    name=pattern if exact else re.compile(pattern, re.I),
                    exact=exact,
                )
            )
        target = _first_visible(link_locators)
        if target is not None:
            try:
                target.click()
                return True
            except Exception:
                continue
    for pattern in patterns:
        text_locators = []
        for scope in _scopes(page):
            text_locators.append(scope.get_by_text(pattern, exact=exact))
        target = _first_visible(text_locators)
        if target is not None:
            try:
                target.click()
                return True
            except Exception:
                continue
    return False


def click_first_until(
    page: Page,
    patterns: Iterable[str],
    *,
    exact: bool = False,
    timeout_ms: int = 15000,
) -> bool:
    """Wait for a late-rendered control, then click it once it is visible."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if click_first(page, patterns, exact=exact):
            return True
        page.wait_for_timeout(400)
    return False


def fill_first_until(
    page: Page,
    labels: Iterable[str],
    selectors: Iterable[str],
    value: str,
    *,
    timeout_ms: int = 10000,
) -> bool:
    """Wait for a modal form field that appears after an animated transition."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if fill_first(page, labels, selectors, value):
            return True
        page.wait_for_timeout(350)
    return False


def scroll_to_text(page: Page, patterns: Iterable[str], *, exact: bool = True) -> bool:
    """Bring a late-page section into view without relying on screen coordinates."""
    for pattern in patterns:
        for scope in _scopes(page):
            try:
                target = _first_visible([scope.get_by_text(pattern, exact=exact)])
                if target is None:
                    continue
                target.scroll_into_view_if_needed(timeout=3000)
                return True
            except Exception:
                continue
    return False


def fill_first(page: Page, labels: Iterable[str], selectors: Iterable[str], value: str) -> bool:
    # Provider login forms can redraw their iframe immediately after the
    # first control appears. Retry briefly so a visible locator that became
    # stale during the redraw is not treated as a missing field.
    for _ in range(3):
        locators: list[Any] = []
        for scope in _scopes(page):
            # Prefer semantic input selectors so broad labels do not resolve
            # to unrelated checkboxes (Aliyun's "remember username" is one
            # example).
            locators.extend(scope.locator(selector) for selector in selectors)
            for label in labels:
                locators.append(scope.get_by_label(label, exact=False))
        target = _first_fillable(locators)
        if target is None:
            page.wait_for_timeout(250)
            continue
        try:
            target.fill(value)
            return True
        except Exception:
            page.wait_for_timeout(250)
    return False


def _mail_login_form_visible(page: Page, email_selectors: Iterable[str]) -> bool:
    return (
        _first_fillable(_scoped_locators(page, email_selectors)) is not None
        or _first_visible(_scoped_locators(page, ['input[type="password"]', 'input[name*="pass" i]'])) is not None
    )


def _mailbox_logged_in(page: Page, route_key: str, email_selectors: Iterable[str]) -> bool:
    if _mail_login_form_visible(page, email_selectors):
        return False
    urls = [page.url.lower(), *[frame.url.lower() for frame in page.frames]]
    text = page_text(page).lower()
    if route_key == "163":
        return any("/js6/" in url or "main.jsp" in url for url in urls) and any(marker in text for marker in ("收件箱", "inbox"))
    if route_key == "aliyun":
        return any("auth/login" not in url and ("alimail" in url or "aliyun" in url) for url in urls) and any(marker in text for marker in ("收件箱", "inbox", "邮件"))
    return any("task=mail" in url for url in urls) or any(marker in text for marker in ("收件箱", "inbox"))


def _mail_login_error(page: Page) -> str | None:
    text = page_text(page).lower()
    for marker, description in (
        ("用户名或密码错误", "用户名或密码错误"),
        ("密码错误", "密码错误"),
        ("账号不存在", "账号不存在"),
        ("invalid login", "账号或密码错误"),
        ("login failed", "登录失败"),
        ("incorrect", "账号或密码错误"),
    ):
        if marker in text:
            return description
    return None


def _ensure_aliyun_consent(page: Page) -> None:
    # Aliyun keeps the login button inactive until the privacy/service
    # agreement checkbox is selected.
    for scope in _scopes(page):
        try:
            boxes = scope.locator('input[type="checkbox"]')
            for index in range(boxes.count()):
                box = boxes.nth(index)
                if not box.is_visible() or box.is_checked():
                    continue
                parent_text = ""
                try:
                    parent_text = box.locator("xpath=ancestor::*[self::label or @role='checkbox'][1]").inner_text(timeout=500)
                except Exception:
                    pass
                if any(marker in parent_text for marker in ("同意", "隐私", "服务协议")):
                    box.check()
                    return
        except Exception:
            continue
    if click_first(page, ["已阅读并同意"], exact=False):
        return
    raise RuntimeError("阿里云登录页未勾选隐私政策和产品服务协议")


def wait_for_stable(page: Page) -> None:
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except PlaywrightTimeoutError:
        pass


def page_text(page: Page) -> str:
    chunks: list[str] = []
    for scope in _scopes(page):
        try:
            chunks.append(scope.locator("body").inner_text(timeout=3000))
        except Exception:
            continue
    return "\n".join(chunks)


def _parse_mail_timestamp(text: str, reference: datetime) -> datetime | None:
    """Parse the short date labels used by the supported webmail UIs."""
    lines = [re.sub(r"\s+", " ", line.replace("\xa0", " ")).strip() for line in text.splitlines()]
    lines = [line for line in lines if line]
    for line in lines[:8]:
        match = re.search(r"(\d{4})[-/]\s*(\d{1,2})[-/]\s*(\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?", line)
        if match:
            year, month, day = (int(match.group(i)) for i in (1, 2, 3))
            hour, minute = int(match.group(4) or 0), int(match.group(5) or 0)
            try:
                return datetime(year, month, day, hour, minute)
            except ValueError:
                continue
        # NetEase exposes an absolute date in aria-labels and sometimes in
        # the opened-message header: ``2026年9月18日 09:36``.
        match = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日(?:\s*(\d{1,2}):(\d{2}))?", line)
        if match:
            year, month, day = (int(match.group(i)) for i in (1, 2, 3))
            hour, minute = int(match.group(4) or 0), int(match.group(5) or 0)
            try:
                return datetime(year, month, day, hour, minute)
            except ValueError:
                continue
        match = re.search(r"(\d{1,2})月(\d{1,2})日(?:\s*(\d{1,2}):(\d{2}))?", line)
        if match:
            month, day = int(match.group(1)), int(match.group(2))
            hour, minute = int(match.group(3) or 0), int(match.group(4) or 0)
            year = reference.year
            try:
                candidate = datetime(year, month, day, hour, minute)
                if candidate > reference + timedelta(days=1):
                    candidate = candidate.replace(year=year - 1)
                return candidate
            except ValueError:
                continue
        match = re.search(r"(?:今天|今日)\s*(\d{1,2}):(\d{2})", line)
        if match:
            return reference.replace(hour=int(match.group(1)), minute=int(match.group(2)), second=0, microsecond=0)
        match = re.search(r"\b(?:today|idag)\b\s*(\d{1,2}):(\d{2})", line, re.I)
        if match:
            return reference.replace(hour=int(match.group(1)), minute=int(match.group(2)), second=0, microsecond=0)
        match = re.search(r"昨天\s*(\d{1,2}):(\d{2})", line)
        if match:
            day = reference - timedelta(days=1)
            return day.replace(hour=int(match.group(1)), minute=int(match.group(2)), second=0, microsecond=0)
        match = re.search(r"\b(?:yesterday|i\s+g[åa]r)\b\s*(\d{1,2}):(\d{2})", line, re.I)
        if match:
            day = reference - timedelta(days=1)
            return day.replace(hour=int(match.group(1)), minute=int(match.group(2)), second=0, microsecond=0)
        match = re.search(r"(?:星期|周)([一二三四五六日天])\s*(\d{1,2}):(\d{2})", line)
        if match:
            weekday = "一二三四五六日天".index(match.group(1))
            weekday = min(weekday, 6)
            days_back = (reference.weekday() - weekday) % 7
            candidate = reference - timedelta(days=days_back)
            candidate = candidate.replace(hour=int(match.group(2)), minute=int(match.group(3)), second=0, microsecond=0)
            if candidate > reference:
                candidate -= timedelta(days=7)
            return candidate
        match = re.search(
            r"\b(mon(?:day)?|tue(?:sday)?|wed(?:nesday)?|thu(?:rsday)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?)\b"
            r"\s*(\d{1,2}):(\d{2})",
            line,
            re.I,
        )
        if match:
            weekday = {
                "mon": 0,
                "tue": 1,
                "wed": 2,
                "thu": 3,
                "fri": 4,
                "sat": 5,
                "sun": 6,
            }[match.group(1)[:3].lower()]
            days_back = (reference.weekday() - weekday) % 7
            candidate = reference - timedelta(days=days_back)
            candidate = candidate.replace(hour=int(match.group(2)), minute=int(match.group(3)), second=0, microsecond=0)
            if candidate > reference:
                candidate -= timedelta(days=7)
            return candidate
        match = re.fullmatch(r"(\d{1,2}):(\d{2})", line)
        if match:
            candidate = reference.replace(hour=int(match.group(1)), minute=int(match.group(2)), second=0, microsecond=0)
            if candidate > reference:
                candidate -= timedelta(days=1)
            return candidate
    return None


def _mail_row_key(row: Any, row_text: str, scope_url: str = "") -> str:
    """Return a key that remains stable when a provider redraws its inbox.

    NetEase (163) puts a generated page timestamp at the beginning of every
    row id.  Keeping that prefix would make every refresh look like a brand
    new message and would defeat the post-login baseline filter.  We keep the
    stable message-id suffix and use its ``aria-label`` only as a fallback.
    """
    data_id = row.get_attribute("data-mail-id")
    if data_id:
        return data_id
    sign = row.get_attribute("sign")
    row_id = row.get_attribute("id")
    if row_id:
        # 163 ids look like ``<render timestamp>_<message id>Dom``.
        if "_" in row_id and sign == "letter":
            row_id = row_id.split("_", 1)[1]
            return f"163|{row_id}"
        return row_id
    aria = row.get_attribute("aria-label")
    if aria and sign == "letter":
        return f"163|{aria}"
    return f"{scope_url}|{row_text[:300]}"


def _mail_rows(page: Page, route_key: str = "") -> list[tuple[Any, str, datetime | None, str]]:
    """Return visible NPA message rows and their displayed timestamps."""
    rows: list[tuple[Any, str, datetime | None, str]] = []
    seen: set[str] = set()
    reference = datetime.now()
    for scope in _scopes(page):
        for pattern in (
            "NPA kundportal",
            "账户验证码",
            "帐户验证码",
            "Account verification",
            "Verification code",
        ):
            try:
                matches = scope.get_by_text(pattern, exact=False)
                for index in range(matches.count()):
                    match = matches.nth(index)
                    if not match.is_visible():
                        continue
                    # Prefer the provider's clickable message container over
                    # an inner subject node.  Aliyun uses role=listitem with
                    # data-mail-id; Roundcube uses <tr>; NetEase uses a
                    # generated <div role=link sign=letter> around
                    # data-npa-row.
                    row = match.locator(
                        "xpath=ancestor::*[@role='link' and @sign='letter'][1]"
                    )
                    if row.count() == 0:
                        row = match.locator(
                            "xpath=ancestor::*[self::tr or @role='listitem' or @data-mail-id or @data-npa-row][1]"
                        )
                    if row.count() == 0:
                        row = match.locator("xpath=ancestor::*[self::article or self::li][1]")
                    if row.count() == 0 or not row.is_visible():
                        continue
                    row_text = row.inner_text(timeout=1000)
                    key = _mail_row_key(row, row_text, getattr(scope, "url", ""))
                    if key in seen:
                        continue
                    seen.add(key)
                    sent_at = _parse_mail_timestamp(row_text, reference)
                    if sent_at is None:
                        # NetEase displays only “昨日” in the visible row but
                        # keeps the exact date and time in aria-label.
                        sent_at = _parse_mail_timestamp(row.get_attribute("aria-label") or "", reference)
                    rows.append((row, row_text, sent_at, key))
            except Exception:
                continue
    return rows


# 通用行标记脚本。部分邮箱（网易 163 等）的邮件列表用纯 div + span 渲染，
# 没有 tr / role=listitem / data-mail-id / article / li 这类祖先节点，导致
# _mail_rows 一行都取不到。这里改为按"文本内容"识别行：同时含 NPA 关键词和
# 时间字样的最小元素即视为一行，并打上临时属性供 Playwright 定位。
_TAG_ROWS_JS = r"""
() => {
  const KEY = /NPA\s*kundportal|帐户验证码|账户验证码|Account\s+verification|Verification\s+code/i;
  const TIME = /(\d{1,2}\s*[:：]\s*\d{2})|今天|昨天|Today|Yesterday|Idag|I\s+g[åa]r|星期[一二三四五六日天]|周[一二三四五六日天]|Mon(?:day)?|Tue(?:sday)?|Wed(?:nesday)?|Thu(?:rsday)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?|(\d{1,2}\s*月\s*\d{1,2}\s*日)|(\d{4}[-/]\d{1,2}[-/]\d{1,2})/i;
  const MAXLEN = 500;
  document.querySelectorAll('[data-npa-row]').forEach((el) => el.removeAttribute('data-npa-row'));
  let count = 0;
  for (const node of document.body.querySelectorAll('*')) {
    const own = (node.innerText || '').replace(/\s+/g, ' ').trim();
    if (!own || own.length > MAXLEN) continue;
    if (!KEY.test(own) || !TIME.test(own)) continue;
    let hasSmaller = false;
    for (const child of node.children) {
      const text = (child.innerText || '').replace(/\s+/g, ' ').trim();
      if (text && text.length <= MAXLEN && KEY.test(text) && TIME.test(text)) {
        hasSmaller = true;
        break;
      }
    }
    if (hasSmaller) continue;
    node.setAttribute('data-npa-row', String(++count));
  }
  return count;
}
"""


def _mail_rows_generic(page: Page) -> list[tuple[Any, str, datetime | None, str]]:
    """按文本结构兜底识别邮件行，适用于 div/span 布局的邮箱列表。"""
    rows: list[tuple[Any, str, datetime | None, str]] = []
    seen: set[str] = set()
    reference = datetime.now()
    for scope in _scopes(page):
        try:
            scope.evaluate(_TAG_ROWS_JS)
        except Exception:
            continue
        try:
            locator = scope.locator("[data-npa-row]")
            for index in range(locator.count()):
                row = locator.nth(index)
                try:
                    if not row.is_visible():
                        continue
                    row_text = row.inner_text(timeout=1000)
                except Exception:
                    continue
                key = _mail_row_key(row, row_text, getattr(scope, "url", ""))
                if key in seen:
                    continue
                seen.add(key)
                sent_at = _parse_mail_timestamp(row_text, reference)
                if sent_at is None:
                    sent_at = _parse_mail_timestamp(row.get_attribute("aria-label") or "", reference)
                rows.append((row, row_text, sent_at, key))
        except Exception:
            continue
    return rows


def _all_mail_rows(page: Page) -> list[tuple[Any, str, datetime | None, str]]:
    """结构解析优先；一家都取不到时改用通用文本解析。"""
    rows = _mail_rows(page)
    if rows:
        return rows
    return _mail_rows_generic(page)


def _codes_in(text: str) -> set[str]:
    """列出文本里所有带验证码标签的数字，用于"页面上只有一个码"的兜底判断。"""
    return set(re.findall(
        r"(?:账户验证码|帐户验证码|验证码|account\s+verification\s+code|verification\s+code|"
        r"one[- ]time\s+code|security\s+code)\s*[:：]?\s*(\d{6,8})",
        text,
        re.I,
    ))


def _refresh_mailbox(page: Page) -> bool:
    """Ask the provider UI to reload the inbox without leaving the session."""
    if click_first(page, ["刷新", "Refresh", "Reload"], exact=False):
        page.wait_for_timeout(700)
        return True
    try:
        # Aliyun's current inbox toolbar exposes only an icon, so there is no
        # stable text or aria label to click. A normal reload preserves the
        # authenticated session and fetches the latest message list.
        page.reload(wait_until="domcontentloaded", timeout=15000)
        wait_for_stable(page)
        return True
    except Exception:
        pass
    return False


@dataclass
class Customer:
    row_number: int
    name: str
    portal_email: str
    portal_password: str
    mail_email: str
    mail_password: str
    enabled: bool = True

    @classmethod
    def from_row(cls, row_number: int, row: dict[str, Any]) -> "Customer":
        portal_email = _row_value(row, "portal_email", "npa_email", "customer_email", "email", "邮箱", "客户邮箱", "客户账号")
        portal_password = _row_value(row, "portal_password", "npa_password", "password", "邮箱密码", "客户密码", "密码")
        mail_email = _row_value(row, "mail_email", "mailbox_email", "邮箱账号", "邮箱登录账号") or portal_email
        mail_password = _row_value(row, "mail_password", "mailbox_password", "邮箱密码") or portal_password
        enabled_raw = _row_value(row, "enabled", "启用", "处理")
        enabled = enabled_raw.lower() not in {"0", "false", "no", "否", "跳过"}
        return cls(
            row_number=row_number,
            name=_row_value(row, "customer_name", "name", "客户名称", "客户中文名称", "客户英文名称") or portal_email,
            portal_email=portal_email,
            portal_password=portal_password,
            mail_email=mail_email,
            mail_password=mail_password,
            enabled=enabled,
        )


class Audit:
    """实时记录运行状态，并在每条客户完成后落盘检查点。

    运行目录下的日志、成功结果、失败结果和状态文件互相独立。这样即使
    浏览器崩溃或操作人员中途停止，已经完成的记录仍然可见，也不会依赖
    进程结束时才执行一次的批量写盘。
    """

    RESULT_FIELDS = [
        "row_number",
        "status",
        "status_label",
        "admin_action",
        "admin_result",
        "invoice_action",
        "invoice_result",
        "failed_step",
        "failure_reason",
        "error",
        "retryable",
        "updated_at",
        "run_id",
    ]
    STATUS_LABELS = {
        "pending": "未处理",
        "completed": "已完成",
        "failed": "失败",
        "manual_required": "待人工处理",
        "skipped": "已跳过",
    }
    ADMIN_ACTION_LABELS = {
        "added": "已添加并确认",
        "already_exists": "已存在，跳过添加",
        "": "未完成",
    }
    INVOICE_ACTION_LABELS = {
        "updated": "已修改并确认",
        "already_correct": "原值正确，无需修改",
        "manual_required": "待人工确认",
        "": "未完成",
    }
    EVENT_LABELS = {
        "run_started": "任务开始",
        "run_finished": "任务结束",
        "customer_started": "开始处理",
        "open_npa": "打开 NPA",
        "open_mailbox": "打开邮箱",
        "mailbox_login_succeeded": "邮箱登录成功",
        "mailbox_login_failed": "邮箱登录失败",
        "mailbox_login_retry": "邮箱登录重试",
        "npa_code_submit_attempt": "提交验证码尝试",
        "npa_code_submit_retry": "提交验证码重试",
        "verification_mail_selected": "选中验证码邮件",
        "verification_code_detected": "验证码已读取",
        "waiting_verification_mail": "等待验证码邮件",
        "verification_mail_waiting": "等待验证码邮件",
        "settings_opened": "已进入 Settings",
        "admin_added": "Add User 已完成",
        "admin_exists_skip": "管理员已存在，跳过添加",
        "invoice_email_read": "已读取发票邮箱",
        "invoice_email_updated": "发票邮箱已更新",
        "customer_failed": "处理失败",
        "customer_skipped": "已跳过",
        "result_output_write_failed": "结果表写入失败",
        "source_workbook_update_failed": "原表回写失败",
        "browser_recycled": "浏览器已重启",
    }

    def __init__(self, root: Path, input_path: Path, total: int) -> None:
        self.root = root
        self.root.mkdir(parents=True, exist_ok=True)
        self.logs_dir = root / "logs"
        self.results_dir = root / "results"
        self.failed_dir = root / "failed"
        self.state_dir = root / "state"
        for directory in (self.logs_dir, self.results_dir, self.failed_dir, self.state_dir):
            directory.mkdir(parents=True, exist_ok=True)
        self.events_path = self.logs_dir / "events.jsonl"
        self.operator_log_path = self.logs_dir / "operator.log"
        self.results_xlsx_path = self.results_dir / "results.xlsx"
        self.failed_xlsx_path = self.failed_dir / "failed.xlsx"
        self.summary_path = self.state_dir / "run_summary.json"
        self.input_path = input_path
        self.run_id = root.name
        self.total = total
        self.started_at = datetime.now().isoformat(timespec="seconds")
        self.status = "running"
        self.current_index = 0
        self.current_customer = ""
        self.last_event = ""
        self.results: list[dict[str, Any]] = []
        self._result_by_row: dict[int, dict[str, Any]] = {}
        self.source_headers, self.source_rows = self._read_source_snapshot(input_path)
        self._write_summary()
        self._write_result_workbooks()

    @staticmethod
    def _read_source_snapshot(path: Path) -> tuple[list[str], dict[int, dict[str, Any]]]:
        if path.suffix.lower() == ".xlsx":
            try:
                from openpyxl import load_workbook
            except ImportError as exc:
                raise RuntimeError("写入 XLSX 需要 openpyxl，请先运行启动脚本安装依赖。") from exc
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                rows = list(workbook.active.iter_rows(values_only=True))
            finally:
                workbook.close()
            if not rows:
                return [], {}
            headers = [str(value or "").strip() for value in rows[0]]
            values = {
                index + 2: {header: row[index] if index < len(row) else "" for index, header in enumerate(headers) if header}
                for index, row in enumerate(rows[1:])
                if any(row)
            }
            return headers, values
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            headers = list(reader.fieldnames or [])
            values = {index + 2: dict(row) for index, row in enumerate(reader) if any(row.values())}
        return headers, values

    @staticmethod
    def _atomic_write_text(path: Path, text: str) -> None:
        temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
        try:
            temporary.write_text(text, encoding="utf-8")
            os.replace(temporary, path)
        finally:
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass

    def _write_summary(self) -> None:
        completed = sum(row.get("status") == "completed" for row in self.results)
        failed = sum(row.get("status") == "failed" for row in self.results)
        manual_required = sum(row.get("status") == "manual_required" for row in self.results)
        skipped = sum(row.get("status") == "skipped" for row in self.results)
        pending = max(self.total - len(self.results), 0)
        summary = {
            "run_id": self.run_id,
            "status": self.status,
            "started_at": self.started_at,
            "finished_at": datetime.now().isoformat(timespec="seconds") if self.status == "completed" else "",
            "total": self.total,
            "processed": len(self.results),
            "completed": completed,
            "failed": failed,
            "manual_required": manual_required,
            "skipped": skipped,
            "pending": pending,
            "current_index": self.current_index,
            "current_customer": self.current_customer,
            "last_event": self.last_event,
        }
        self._atomic_write_text(self.summary_path, json.dumps(summary, ensure_ascii=False, indent=2))

    def _write_operator_line(self, customer: Customer | None, label: str, detail: str = "") -> None:
        timestamp = datetime.now().strftime("%m/%d %H:%M:%S")
        subject = customer.name if customer else "系统"
        clean_detail = re.sub(r"\s+", " ", str(detail or "")).strip()
        if len(clean_detail) > 240:
            clean_detail = clean_detail[:237] + "..."
        line = f"{timestamp} | {subject} | {label}"
        if clean_detail:
            line += f" | {clean_detail}"
        with self.operator_log_path.open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
            handle.flush()

    def set_current(self, index: int, customer: Customer) -> None:
        self.current_index = index
        self.current_customer = customer.name
        self._write_summary()

    def event(self, customer: Customer | None, event: str, detail: str = "", level: str = "info") -> None:
        payload = {
            "time": datetime.now().isoformat(timespec="seconds"),
            "level": level,
            "event": event,
            "customer": customer.name if customer else "",
            "customer_email": normalize_email(customer.portal_email) if customer else "",
            "detail": detail,
        }
        with self.events_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(payload, ensure_ascii=False) + "\n")
            handle.flush()
        self.last_event = event
        self._write_operator_line(customer, self.EVENT_LABELS.get(event, event), detail)
        print(f"[{level.upper()}] {payload['customer'] or '-'}: {event}{(' - ' + detail) if detail else ''}", flush=True)
        self._write_summary()

    def _combined_row(self, row_number: int, result: dict[str, Any]) -> dict[str, Any]:
        source = dict(self.source_rows.get(row_number, {}))
        source.update(result)
        return source

    def _pending_result(self, row_number: int) -> dict[str, Any]:
        """Return an explicit placeholder so results.xlsx never looks blank."""
        source = self.source_rows.get(row_number, {})
        return {
            "row_number": row_number,
            "customer_name": source.get("customer_name") or source.get("客户名称") or "",
            "portal_email": normalize_email(str(source.get("portal_email") or source.get("邮箱") or "")),
            "status": "pending",
            "status_label": self.STATUS_LABELS["pending"],
            "admin_action": "",
            "admin_result": self.ADMIN_ACTION_LABELS[""],
            "invoice_action": "",
            "invoice_result": self.INVOICE_ACTION_LABELS[""],
            "failed_step": "",
            "failure_reason": "",
            "error": "",
            "retryable": "",
            "updated_at": "",
            "run_id": self.run_id,
        }

    def _write_one_xlsx(self, path: Path, rows: list[dict[str, Any]]) -> None:
        try:
            from openpyxl import Workbook
            from openpyxl.utils import get_column_letter
            from openpyxl.styles import Alignment, Font, PatternFill
        except ImportError as exc:
            raise RuntimeError("写入 XLSX 需要 openpyxl，请先运行启动脚本安装依赖。") from exc
        headers = list(self.source_headers)
        for field in self.RESULT_FIELDS:
            if field not in headers:
                headers.append(field)
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "results"
        sheet.append(headers)
        header_fill = PatternFill("solid", fgColor="D9EAD3")
        header_font = Font(bold=True)
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in rows:
            values = [row.get(header, "") if row.get(header, "") is not None else "" for header in headers]
            sheet.append(values)
            for cell in sheet[sheet.max_row]:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            status = str(row.get("status", ""))
            if status == "failed":
                status_fill = PatternFill("solid", fgColor="F4CCCC")
            elif status == "manual_required":
                status_fill = PatternFill("solid", fgColor="FCE5CD")
            elif status == "completed":
                status_fill = PatternFill("solid", fgColor="D9EAD3")
            elif status == "pending":
                status_fill = PatternFill("solid", fgColor="EDEDED")
            else:
                status_fill = None
            if status_fill:
                status_index = headers.index("status_label") + 1 if "status_label" in headers else headers.index("status") + 1
                sheet.cell(sheet.max_row, status_index).fill = status_fill
        sheet.freeze_panes = "A2"
        if rows:
            sheet.auto_filter.ref = sheet.dimensions
        wide_fields = {"status_label", "admin_result", "invoice_result", "failed_step", "failure_reason", "error"}
        for index, header in enumerate(headers, 1):
            width = 34 if header in wide_fields else min(max(len(str(header)) + 2, 12), 28)
            sheet.column_dimensions[get_column_letter(index)].width = width
        temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
        try:
            workbook.save(temporary)
            os.replace(temporary, path)
        finally:
            workbook.close()
            try:
                temporary.unlink()
            except FileNotFoundError:
                pass

    def _write_result_workbooks(self) -> None:
        # Keep every source row in results.xlsx. Unprocessed rows are explicit
        # ``pending`` records instead of looking like a copied input sheet.
        ordered = []
        for row_number in sorted(self.source_rows):
            result = self._result_by_row.get(row_number) or self._pending_result(row_number)
            ordered.append(self._combined_row(row_number, result))
        failed = [row for row in ordered if row.get("status") in {"failed", "manual_required"}]
        self._write_one_xlsx(self.results_xlsx_path, ordered)
        self._write_one_xlsx(self.failed_xlsx_path, failed)

    def _update_source_workbook(self) -> None:
        if self.input_path.suffix.lower() != ".xlsx":
            return
        from openpyxl import load_workbook

        workbook = load_workbook(self.input_path)
        try:
            sheet = workbook.active
            headers = [str(sheet.cell(1, column).value or "").strip() for column in range(1, sheet.max_column + 1)]
            for field in self.RESULT_FIELDS:
                if field not in headers:
                    headers.append(field)
                    sheet.cell(1, len(headers), field)
            columns = {header: index + 1 for index, header in enumerate(headers) if header}
            for row_number, result in self._result_by_row.items():
                for field in self.RESULT_FIELDS:
                    sheet.cell(row_number, columns[field], result.get(field, ""))
            temporary = self.input_path.with_name(f".{self.input_path.stem}.robot_tmp_{os.getpid()}.xlsx")
            try:
                workbook.save(temporary)
                # Windows may keep the source ZIP open until close(). Close
                # before replacing the original so a WPS/Explorer lock is
                # not confused with an incomplete write.
                workbook.close()
                os.replace(temporary, self.input_path)
            finally:
                try:
                    temporary.unlink()
                except FileNotFoundError:
                    pass
        finally:
            try:
                workbook.close()
            except Exception:
                pass

    def result(self, customer: Customer, **values: Any) -> None:
        now = datetime.now().isoformat(timespec="seconds")
        status = str(values.get("status") or "failed")
        admin_action = str(values.get("admin_action") or "")
        invoice_action = str(values.get("invoice_action") or "")
        failed_step = str(values.get("failed_step") or "")
        error = str(values.get("error") or values.get("failure_reason") or "")
        retryable_value = values.get("retryable")
        if retryable_value is None:
            retryable_value = "是" if status in {"failed", "manual_required"} and failed_step not in {"输入校验", "邮箱入口匹配"} else "否"
        row = {
            "row_number": customer.row_number,
            "customer_name": customer.name,
            "portal_email": normalize_email(customer.portal_email),
            "status": status,
            "status_label": self.STATUS_LABELS.get(status, status),
            "admin_action": admin_action,
            "admin_result": self.ADMIN_ACTION_LABELS.get(admin_action, admin_action or "未完成"),
            "invoice_action": invoice_action,
            "invoice_result": self.INVOICE_ACTION_LABELS.get(invoice_action, invoice_action or "未完成"),
            "failed_step": failed_step,
            "failure_reason": error,
            "error": error,
            "retryable": str(retryable_value),
            "updated_at": now,
            "run_id": self.run_id,
        }
        self.results = [existing for existing in self.results if existing.get("row_number") != customer.row_number]
        self.results.append(row)
        self._result_by_row[customer.row_number] = row
        try:
            self._write_result_workbooks()
        except Exception as exc:
            self.event(customer, "result_output_write_failed", str(exc), "warning")
        try:
            self._update_source_workbook()
        except Exception as exc:
            self.event(customer, "source_workbook_update_failed", str(exc), "warning")
        status = row["status"]
        self._write_operator_line(customer, "本条已记录", f"状态={status}")
        completed = sum(item.get("status") == "completed" for item in self.results)
        failed = sum(item.get("status") == "failed" for item in self.results)
        manual_required = sum(item.get("status") == "manual_required" for item in self.results)
        skipped = sum(item.get("status") == "skipped" for item in self.results)
        progress = f"进度：{len(self.results)}/{self.total} | 成功 {completed} | 失败 {failed} | 待人工 {manual_required} | 跳过 {skipped}"
        self._write_operator_line(None, progress)
        print(f"[PROGRESS] {progress}", flush=True)
        self._write_summary()

    def write_results(self) -> None:
        self._write_result_workbooks()
        self._write_summary()


class NpaBot:
    def __init__(self, config: dict[str, Any], audit: Audit, *, auto_code: bool) -> None:
        self.config = config
        self.audit = audit
        self.auto_code = auto_code
        self.portal_url = config["portal_url"]
        self.target_user = config["target_user"]
        self.target_admin_email = normalize_email(self.target_user.get("email", TARGET_ADMIN_EMAIL))
        self.target_invoice_email = normalize_email(config.get("target_invoice_email", TARGET_INVOICE_EMAIL))
        # 等待验证码邮件出现的秒数。微软发码 + 邮箱刷新实测可达 30~60 秒。
        self.code_wait_seconds = int(config.get("code_wait_seconds", DEFAULT_CODE_WAIT_SECONDS) or DEFAULT_CODE_WAIT_SECONDS)
        self.current_step = ""

    def login_mailbox(
        self,
        page: Page,
        customer: Customer,
        route: dict[str, Any],
        *,
        attempt: int = 1,
    ) -> datetime:
        self.current_step = "登录邮箱"
        self.audit.event(customer, "open_mailbox", route.get("url", ""))
        page.goto(route["url"], wait_until="domcontentloaded")
        wait_for_stable(page)

        def retry_login(message: str) -> datetime:
            if attempt >= 3:
                raise RuntimeError(message)
            next_attempt = attempt + 1
            self.audit.event(customer, "mailbox_login_retry", f"{message}；准备第 {next_attempt} 次尝试", "warning")
            try:
                page.goto(route["url"], wait_until="domcontentloaded", timeout=15000)
                wait_for_stable(page)
            except Exception:
                # The recursive attempt will surface the actual page error.
                pass
            return self.login_mailbox(page, customer, route, attempt=next_attempt)

        email_selectors = [
            'input[type="email"]',
            'input[name*="user" i]',
            'input[name*="login" i]',
            'input[autocomplete^="username"]',
            'input[placeholder*="email" i]',
            'input[placeholder*="邮箱" i]',
            'input[placeholder*="账号" i]',
        ]
        # NetEase opens on QR-code login. Switch to its account/password tab
        # before looking for the (initially hidden) account fields.
        if route.get("key") == "163" and _wait_for_fillable(page, email_selectors, timeout_ms=3500) is None:
            click_first(page, ["密码登录"], exact=True)
            page.wait_for_timeout(500)
        # Some providers load the account and password controls at slightly
        # different times (NetEase uses a delayed iframe).
        if _wait_for_fillable(page, email_selectors) is None:
            return retry_login("邮箱登录页未找到账号输入框")
        password_input = _wait_for_scoped(page, ['input[type="password"]', 'input[name*="pass" i]'])
        if password_input is None:
            return retry_login("邮箱登录页未找到密码输入框；请确认入口地址、页面加载状态或是否需要人工接管")
        filled_email = fill_first(
            page,
            ["Email", "邮箱", "Username", "用户名", "账号"],
            email_selectors,
            customer.mail_email,
        )
        if not filled_email:
            return retry_login("邮箱登录页未找到账号输入框")
        password_input.fill(customer.mail_password)
        if route.get("key") == "aliyun":
            _ensure_aliyun_consent(page)
        clicked_login = False
        if route.get("key") == "163":
            # NetEase renders the submit control as an anchor with a stable
            # id, while the outer page also contains navigation links named
            # “登录”. Scope this click to the actual login form.
            target = _first_visible([
                scope.locator("#dologin, a[data-action='dologin']")
                for scope in _scopes(page)
            ])
            if target is not None:
                try:
                    target.click()
                    clicked_login = True
                except Exception:
                    clicked_login = False
        if not clicked_login and not click_first(page, ["Log in", "Login", "Sign in", r"登\s*录", "登录", "登陆"], exact=False):
            return retry_login("邮箱登录页未找到登录按钮")
        deadline = time.monotonic() + 15000 / 1000
        while time.monotonic() < deadline:
            if _mailbox_logged_in(page, str(route.get("key", "")), email_selectors):
                logged_in_at = datetime.now()
                self.current_step = "读取验证码"
                self.audit.event(customer, "mailbox_login_succeeded", logged_in_at.isoformat(timespec="seconds"))
                if route.get("key") == "163":
                    click_first(page, ["收件箱"], exact=True)
                # Let the initial inbox render, then remember the message IDs
                # already present. New-code selection can therefore reject an
                # older NPA message even when both messages show the same
                # minute-only timestamp.
                page.wait_for_timeout(1500)
                self._mailbox_baseline_keys = {
                    key for _, _, _, key in _all_mail_rows(page)
                }
                return logged_in_at
            page.wait_for_timeout(500)
        error = _mail_login_error(page)
        detail = f"：{error}" if error else "，提交后仍停留在邮箱登录页"
        self.audit.event(
            customer,
            "mailbox_login_failed",
            f"入口={route.get('key', '')}，页面={page.url}{detail}",
            "warning",
        )
        if error is None and attempt < 3:
            return retry_login(f"邮箱登录未成功{detail}")
        raise RuntimeError(f"邮箱登录未成功{detail}")

    def _read_code_from_page(self, page: Page) -> str | None:
        text = page_text(page)
        return extract_verification_code(text)

    def _read_code_from_selected_mail(self, page: Page, route_key: str) -> str | None:
        """Read only the opened message body, excluding older list previews."""
        selectors = [
            "#messagebody",
            "#message-htmlpart1",
            "[id*='messagebody' i]",
            "[id*='message-htmlpart' i]",
            "[class*='aym_scale_wrap']",
            "[class*='aym_table_wrap']",
            "[class*='message-content' i]",
        ]
        if route_key == "163":
            # NetEase places the opened HTML message in a frame named
            # ``frameBody``.  The frame is already included by _scopes, but
            # these selectors make the provider-specific body boundary
            # explicit and prevent list previews from being considered.
            selectors.extend(["iframe[id*='frameBody' i]", "body"])
        for scope in _scopes(page):
            for selector in selectors:
                try:
                    locators = scope.locator(selector)
                    for index in range(locators.count()):
                        candidate = locators.nth(index)
                        if selector != "body" and not candidate.is_visible():
                            continue
                        code = extract_verification_code(candidate.inner_text(timeout=1000))
                        if code:
                            return code
                except Exception:
                    continue
        return self._read_code_from_page(page)

    def get_code(
        self,
        mailbox_page: Page,
        customer: Customer,
        *,
        mailbox_login_at: datetime | None = None,
        route_key: str = "",
    ) -> str:
        self.current_step = "读取验证码"
        if self.auto_code:
            if route_key == "163":
                # NetEase may leave the user on its welcome page after login.
                # Open the inbox before looking for the verification message.
                click_first(mailbox_page, ["收件箱"], exact=True)
                mailbox_page.wait_for_timeout(700)
            # Mail providers often navigate to the inbox before the message
            # list has finished rendering (especially NetEase). Poll briefly
            # and open the newest NPA message once it becomes available.
            deadline = time.monotonic() + self.code_wait_seconds
            poll_count = 0
            self.audit.event(
                customer,
                "waiting_verification_mail",
                f"正在查找登录后收到的最新验证码邮件，最长等待 {self.code_wait_seconds} 秒",
            )
            while time.monotonic() < deadline:
                if poll_count and poll_count % 6 == 0:
                    _refresh_mailbox(mailbox_page)
                if mailbox_login_at is not None:
                    cutoff = mailbox_login_at - MAIL_TIMESTAMP_TOLERANCE
                    baseline = getattr(self, "_mailbox_baseline_keys", set())
                    candidates = []
                    mail_rows = _all_mail_rows(mailbox_page)
                    for row, row_text, sent_at, key in mail_rows:
                        if key in baseline or sent_at is None or sent_at < cutoff:
                            continue
                        candidates.append((row, row_text, sent_at))
                    candidates.sort(key=lambda item: item[2], reverse=True)
                    for row, row_text, sent_at in candidates:
                        code = extract_verification_code(row_text)
                        try:
                            row.click()
                        except Exception:
                            click_first(mailbox_page, ["NPA kundportal", "账户验证码", "Account verification", "Verification code"], exact=False)
                        mailbox_page.wait_for_timeout(700)
                        code = self._read_code_from_selected_mail(mailbox_page, route_key) or code
                        if code:
                            self.audit.event(customer, "verification_mail_selected", sent_at.isoformat(timespec="seconds"))
                            self.audit.event(customer, "verification_code_detected")
                            return code
                    if poll_count and poll_count % 20 == 0:
                        self.audit.event(
                            customer,
                            "verification_mail_waiting",
                            f"仍在等待新邮件；当前识别到 {len(mail_rows)} 封 NPA 邮件",
                        )
                else:
                    code = self._read_code_from_page(mailbox_page)
                    if code:
                        self.audit.event(customer, "verification_code_detected")
                        return code
                    click_first(mailbox_page, ["NPA kundportal", "账户验证码", "Account verification", "Verification code"], exact=False)
                mailbox_page.wait_for_timeout(500)
                poll_count += 1
            if mailbox_login_at is not None:
                self.audit.event(
                    customer,
                    "verification_mail_not_found_after_login",
                    f"登录时间={mailbox_login_at.isoformat(timespec='seconds')}，允许提前 {int(MAIL_TIMESTAMP_TOLERANCE.total_seconds())} 秒",
                    "warning",
                )
            # 兜底：列表行定位失败时，如果当前页面（列表预览 / 已打开的阅读区）
            # 只出现一个验证码，直接采用；同时出现多个（例如旧邮件还开着）则放弃，
            # 避免把过期验证码填进 NPA。
            page_codes = _codes_in(page_text(mailbox_page))
            if len(page_codes) == 1:
                self.audit.event(customer, "verification_code_from_page_fallback", "页面仅出现一个验证码")
                return next(iter(page_codes))
        raise RuntimeError("自动读取验证码失败，未找到登录后收到的有效验证码邮件")

    def login_npa(self, page: Page, mailbox_page: Page, customer: Customer, route: dict[str, Any]) -> None:
        self.current_step = "登录 NPA"
        self.audit.event(customer, "open_npa", self.portal_url)
        # NPA occasionally keeps a slow third-party resource open.  Use one
        # short DOM-load attempt followed by two lighter navigation retries;
        # a partially loaded page is still useful because the login-shell
        # wait below can finish rendering it.
        navigation_error: Exception | None = None
        for attempt in range(3):
            try:
                page.goto(
                    self.portal_url,
                    wait_until="domcontentloaded" if attempt == 0 else "commit",
                    timeout=15000 if attempt == 0 else 8000,
                )
                navigation_error = None
                break
            except PlaywrightTimeoutError as exc:
                navigation_error = exc
                self.audit.event(
                    customer,
                    "npa_navigation_retry",
                    f"首页加载超时，准备第 {attempt + 2} 次尝试",
                    "warning",
                )
                page.wait_for_timeout(400)
        if page.url in {"", "about:blank"}:
            raise RuntimeError("NPA 首页加载超时") from navigation_error
        wait_for_stable(page)
        # Log into the mailbox before requesting the NPA code. This records a
        # clean timestamp and a baseline of existing messages, so the code
        # reader can select only a newly arrived verification email.
        mailbox_login_at = self.login_mailbox(mailbox_page, customer, route)
        self.current_step = "登录 NPA"
        email_selectors = ['input[type="email"]', 'input[autocomplete^="username"]', 'input[name*="email" i]', 'input[name="username"]', 'input[placeholder*="email" i]', 'input[placeholder*="邮箱" i]', 'input[placeholder*="账号" i]']
        if _wait_for_scoped(page, email_selectors, timeout_ms=2500) is None:
            if not click_first_until(page, ["Logga in", "Log in", "Login", "登录", "登陆"], exact=False, timeout_ms=10000):
                raise RuntimeError("NPA 登录页未找到登录入口")
            if _wait_for_scoped(page, email_selectors, timeout_ms=30000) is None:
                # The portal can leave its public Swedish landing page open
                # while the CIAM redirect is still starting.  One reload is
                # safer than failing after the old 15-second window.
                try:
                    if "portal-prod.npa.se" in page.url.lower():
                        page.reload(wait_until="domcontentloaded", timeout=15000)
                        wait_for_stable(page)
                        click_first_until(page, ["Logga in", "Log in", "Login", "登录", "登陆"], exact=False, timeout_ms=10000)
                    if _wait_for_scoped(page, email_selectors, timeout_ms=15000) is None:
                        raise RuntimeError("NPA 登录页加载超时，未找到邮箱输入框")
                except PlaywrightTimeoutError:
                    raise RuntimeError("NPA 登录页加载超时，未找到邮箱输入框")
        if not fill_first(page, ["Email", "邮箱", "Username", "用户名"], ['input[type="email"]', 'input[autocomplete^="username"]', 'input[name*="email" i]', 'input[name="username"]', 'input[placeholder*="email" i]', 'input[placeholder*="邮箱" i]', 'input[placeholder*="账号" i]'], customer.portal_email):
            raise RuntimeError("NPA 登录页未找到邮箱输入框")
        if not click_first(page, ["Next", "Continue", "下一步", "Fortsätt"], exact=False):
            # Some login pages submit the email form with the login button itself.
            if not click_first(page, ["Logga in", "Log in", "Login", "登录", "登陆"], exact=False):
                raise RuntimeError("NPA 登录页未找到下一步按钮")
        wait_for_stable(page)
        password_after_email = _wait_for_scoped(page, ['input[type="password"]', 'input[name*="pass" i]'], timeout_ms=2000)
        if password_after_email is not None:
            if not customer.portal_password:
                raise RuntimeError("NPA 登录页要求密码，但客户记录未提供 portal_password")
            password_after_email.fill(customer.portal_password)
            if not click_first(page, ["Next", "Continue", "Sign in", "Log in", "登录"], exact=False):
                raise RuntimeError("NPA 密码页未找到提交按钮")
            wait_for_stable(page)
        _wait_for_scoped(page, ['input[name*="code" i]', 'input[name="npotc"]', 'input[autocomplete="one-time-code"]', 'input[inputmode="numeric"]'], timeout_ms=15000)
        code = self.get_code(
            mailbox_page,
            customer,
            mailbox_login_at=mailbox_login_at,
            route_key=str(route.get("key", "")),
        )
        self.current_step = "提交验证码"
        if not fill_first(page, ["Code", "Verification code", "验证码"], ['input[name*="code" i]', 'input[name="npotc"]', 'input[autocomplete="one-time-code"]', 'input[inputmode="numeric"]'], code):
            raise RuntimeError("NPA 页面未找到验证码输入框")
        # Microsoft CIAM sometimes renders the button before it is enabled,
        # and on other attempts adds its accessible name only after the OTP
        # input event.  Prefer the stable id and wait for an enabled control;
        # then fall back to the localized button names.
        submitted = False
        code_selectors = [
            'input[name*="code" i]',
            'input[name="npotc"]',
            'input[autocomplete="one-time-code"]',
            'input[inputmode="numeric"]',
        ]
        submit_name_pattern = re.compile(
            r"Continue|Verify|Submit|Sign in|Log in|登录|确认|Fortsätt|Verifiera|"
            r"Bekräfta|Nästa|Skicka|Gå vidare",
            re.I,
        )
        for attempt in range(1, 4):
            self.audit.event(customer, "npa_code_submit_attempt", f"第 {attempt}/3 次")
            submit_deadline = time.monotonic() + 10
            while time.monotonic() < submit_deadline and not submitted:
                candidates = []
                for scope in _scopes(page):
                    candidates.extend([
                        scope.locator("#oneTimeCodePrimaryButton"),
                        scope.get_by_role("button", name=submit_name_pattern),
                        scope.locator("button[type='submit'], input[type='submit']"),
                        scope.locator("button[id*='oneTimeCode' i], button[data-testid*='code' i]"),
                        scope.locator("button").filter(has_text=submit_name_pattern),
                    ])
                for target in candidates:
                    try:
                        if target.count() and target.first.is_visible() and target.first.is_enabled():
                            target.first.click(timeout=5000)
                            submitted = True
                            break
                    except Exception:
                        continue
                if not submitted:
                    page.wait_for_timeout(400)
            if not submitted:
                # Some NPA variants do not expose an accessible button name
                # but submit the one-time-code form when Enter is pressed.
                for scope in _scopes(page):
                    try:
                        code_input = _first_visible([scope.locator(selector) for selector in code_selectors])
                        if code_input is None:
                            continue
                        code_input.press("Enter")
                        page.wait_for_timeout(1200)
                        if _first_visible([scope.locator(selector) for selector in code_selectors]) is None:
                            submitted = True
                            break
                    except Exception:
                        continue
            if submitted:
                break
            if attempt < 3:
                self.audit.event(customer, "npa_code_submit_retry", "当前页面未找到可用提交控件，准备重新查找", "warning")
                page.wait_for_timeout(700)
        if not submitted:
            raise RuntimeError("NPA 页面未找到验证码提交按钮")
        wait_for_stable(page)
        # Do not proceed to Settings while the OAuth page is still showing
        # the code form.  This turns a later, misleading “Settings not found”
        # error into a precise login failure.
        portal_deadline = time.monotonic() + 30
        while time.monotonic() < portal_deadline:
            current_url = page.url.lower()
            if "portal-prod.npa.se" in current_url and "ciamlogin" not in current_url:
                text = page_text(page).lower()
                if any(marker in text for marker in ("settings", "welcome to the npa customer portal", "home")):
                    return
            page.wait_for_timeout(500)
        if "ciamlogin" in page.url.lower() or _first_visible(_scoped_locators(page, ['input[name*="code" i]', 'input[name="npotc"]', 'input[autocomplete="one-time-code"]', 'input[inputmode="numeric"]'])) is not None:
            raise RuntimeError("NPA 验证码提交后仍停留在登录页")

    def open_settings(self, page: Page, customer: Customer) -> None:
        # The NPA SPA shows a short “Laddar...” state after the OAuth
        # redirect. Wait for the navigation item instead of treating that
        # normal render delay as a failure.
        # The portal first renders the shell and then fetches the Settings
        # route.  On slower customer accounts the navigation can appear well
        # after the OAuth redirect has completed.
        deadline = time.monotonic() + 45
        while time.monotonic() < deadline:
            if click_first(page, ["Settings", "Inställningar", "设置"], exact=True):
                wait_for_stable(page)
                # The SOP screenshot shows Settings at the top-level sidebar,
                # while Users/Add User is much farther down the page.  Wait
                # for the actual Users section and scroll it into view so the
                # following operation follows the documented path.
                if _wait_for_scoped(page, ['text="Users"', 'text="Användare"', 'text="用户"'], timeout_ms=15000) is not None:
                    scroll_to_text(page, ["Users", "Användare", "用户"], exact=True)
                self.audit.event(customer, "settings_opened")
                return
            page.wait_for_timeout(500)
        raise RuntimeError("NPA 页面未找到 Settings")

    def target_user_exists(self, page: Page) -> bool:
        body = page_text(page)
        return self.target_admin_email in normalize_email(body)

    def add_user(self, page: Page, customer: Customer) -> str:
        if self.target_user_exists(page):
            self.audit.event(customer, "admin_exists_skip", self.target_admin_email)
            return "already_exists"
        # Settings is a React page: the text heading can be present before
        # the Users card and its button have finished rendering.  The former
        # one-shot click caused the formal run to fail even though Add User
        # was present a moment later.
        scroll_to_text(page, ["Users", "Användare", "用户"], exact=True)
        if not click_first_until(page, ["Add User", "Lägg till användare", "Bjud in användare", "添加用户"], exact=True, timeout_ms=20000):
            raise RuntimeError("Settings 页面未找到 Add User")
        if not fill_first_until(page, ["First Name", "Förnamn", "名"], ['#newUserFirstName', 'input[name*="first" i]'], str(self.target_user.get("first_name", "Min"))):
            raise RuntimeError("Add User 未找到 First Name")
        if not fill_first_until(page, ["Last Name", "Efternamn", "姓"], ['#newUserLastName', 'input[name*="last" i]'], str(self.target_user.get("last_name", "Liu"))):
            raise RuntimeError("Add User 未找到 Last Name")
        if not fill_first_until(page, ["Email", "E-post", "邮箱"], ['#newUserEmail', 'input[type="email"]', 'input[name*="email" i]'], self.target_admin_email):
            raise RuntimeError("Add User 未找到 Email")
        permission = str(self.target_user.get("permission", "Administrator"))
        select = _first_visible([page.locator("select[name*='permission' i]"), page.locator("select")])
        if select is not None:
            try:
                select.select_option(label=permission)
            except Exception:
                select.select_option(value=permission)
        else:
            # The current NPA form uses a button labelled “Permission” and
            # opens a Radix combobox; it is not a native <select>.  The old
            # code looked only for an already-visible “Administrator” item
            # and never opened this combobox.
            opened = False
            deadline = time.monotonic() + 10000 / 1000
            while time.monotonic() < deadline and not opened:
                combo_locators = []
                for scope in _scopes(page):
                    # The Swedish form exposes an unnamed Radix combobox
                    # whose visible text is “Behörighet”; English exposes
                    # “Permission”.  Use the role first and only fall back
                    # to the label when the role is not available.
                    combo_locators.extend([
                        scope.get_by_role("combobox"),
                        scope.locator('[role="combobox"]'),
                        scope.get_by_role("button", name=re.compile("Permission|Behörighet", re.I)),
                    ])
                combo = _first_visible(combo_locators)
                if combo is not None:
                    try:
                        combo.click()
                        opened = True
                        break
                    except Exception:
                        pass
                page.wait_for_timeout(300)
            if not opened:
                opened = click_first_until(page, ["Permission", "Behörighet"], exact=True, timeout_ms=2000)
            selected = False
            deadline = time.monotonic() + 5000 / 1000
            while time.monotonic() < deadline and not selected:
                option_locators = []
                for scope in _scopes(page):
                    option_locators.extend([
                        scope.get_by_role("option", name=re.compile(r"(?:" + re.escape(permission) + r"|Admin|Administratör)", re.I)),
                        scope.get_by_role("menuitem", name=re.compile(r"(?:" + re.escape(permission) + r"|Admin|Administratör)", re.I)),
                        scope.locator('[role="dialog"]').get_by_text(re.compile(r"(?:" + re.escape(permission) + r"|Admin|Administratör)", re.I)),
                    ])
                option = _first_visible(option_locators)
                if option is not None:
                    try:
                        option.click()
                        selected = True
                        break
                    except Exception:
                        pass
                page.wait_for_timeout(300)
            if not opened or not selected:
                raise RuntimeError("Add User 未找到权限选择控件")
        if not click_first(page, ["Add", "Bjud in", "Lägg till", "Skicka inbjudan", "添加"], exact=True):
            raise RuntimeError("Add User 表单未找到 Add 按钮")
        wait_for_stable(page)
        # The Add request succeeds before the Users table is refreshed.  Poll
        # the table, then reload the Settings route once if the SPA still has
        # its previous list in memory.  Without this confirmation window a
        # successful Add was incorrectly reported as failed.
        confirmed = False
        for attempt in range(2):
            deadline = time.monotonic() + 12
            while time.monotonic() < deadline:
                if self.target_user_exists(page):
                    confirmed = True
                    break
                page.wait_for_timeout(1000)
            if confirmed:
                break
            if attempt == 0:
                try:
                    page.reload(wait_until="domcontentloaded", timeout=15000)
                    wait_for_stable(page)
                    scroll_to_text(page, ["Users", "Användare", "用户"], exact=True)
                except Exception:
                    pass
        if not confirmed:
            raise RuntimeError("点击 Add 后未在 Users 列表发现目标管理员")
        self.audit.event(customer, "admin_added", self.target_admin_email)
        return "added"

    def invoice_action(self, page: Page, customer: Customer) -> str:
        body = page_text(page)
        current = extract_invoice_email(body)
        if current is None:
            self.audit.event(customer, "invoice_email_not_found", level="warning")
            return "manual_required"
        self.audit.event(customer, "invoice_email_read", current)
        if current == self.target_invoice_email:
            return "already_correct"
        scroll_to_text(page, ["Invoice email", "Faktura e-post", "发票邮箱"], exact=True)
        if not click_first(page, ["Edit invoice email", "Ändra faktura e-post", "编辑发票邮箱"], exact=True):
            raise RuntimeError("Settings 页面未找到 Edit invoice email")
        if not fill_first(page, ["Invoice email", "Faktura e-post", "发票邮箱"], ['input[type="email"]', 'input[name*="invoice" i]'], self.target_invoice_email):
            raise RuntimeError("Invoice email 编辑框未找到")
        if not click_first(page, ["Save", "Spara", "保存"], exact=True):
            raise RuntimeError("Invoice email 编辑框未找到 Save")
        wait_for_stable(page)
        # Saving the invoice field is asynchronous.  The Settings page can
        # still render the old value for several seconds after the API call
        # has succeeded, so confirm the exact target text before reporting a
        # mismatch and reload the route once if necessary.
        updated = None
        for attempt in range(2):
            deadline = time.monotonic() + 10
            while time.monotonic() < deadline:
                body = page_text(page)
                if self.target_invoice_email in normalize_email(body):
                    updated = self.target_invoice_email
                    break
                page.wait_for_timeout(800)
            if updated == self.target_invoice_email:
                break
            if attempt == 0:
                try:
                    page.reload(wait_until="domcontentloaded", timeout=15000)
                    wait_for_stable(page)
                    scroll_to_text(page, ["Invoice email", "Faktura e-post", "发票邮箱"], exact=True)
                except Exception:
                    pass
        if updated != self.target_invoice_email:
            raise RuntimeError("保存后 Invoice email 回读不一致")
        self.audit.event(customer, "invoice_email_updated", self.target_invoice_email)
        return "updated"

    def process(self, customer: Customer, browser_context: BrowserContext, route: dict[str, Any]) -> dict[str, str]:
        portal_page = browser_context.new_page()
        mailbox_page = browser_context.new_page()
        self.current_step = "登录 NPA"
        admin_action = ""
        invoice_action = ""
        try:
            self.login_npa(portal_page, mailbox_page, customer, route)
            self.current_step = "打开 Settings"
            self.open_settings(portal_page, customer)
            self.current_step = "Add User"
            admin_action = self.add_user(portal_page, customer)
            self.current_step = "Invoice email"
            invoice_action = self.invoice_action(portal_page, customer)
            status = "completed" if invoice_action not in {"manual_required"} else "manual_required"
            return {
                "status": status,
                "admin_action": admin_action,
                "invoice_action": invoice_action,
                "failed_step": "" if status == "completed" else self.current_step,
                "error": "" if status == "completed" else "发票邮箱需要人工确认",
            }
        except (RuntimeError, PlaywrightTimeoutError) as exc:
            failed_step = self.current_step or "未知步骤"
            self.audit.event(customer, "customer_failed", f"{failed_step}：{exc}", "error")
            return {
                "status": "failed",
                "admin_action": admin_action,
                "invoice_action": invoice_action,
                "failed_step": failed_step,
                "error": str(exc),
            }
        except Exception as exc:
            failed_step = self.current_step or "未知步骤"
            self.audit.event(customer, "customer_failed", f"{failed_step}：{exc}", "error")
            return {
                "status": "failed",
                "admin_action": admin_action,
                "invoice_action": invoice_action,
                "failed_step": failed_step,
                "error": f"{type(exc).__name__}: {exc}",
            }
        finally:
            portal_page.close()
            mailbox_page.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="瑞典 NPA 管理员与 Invoice email 自动化工具")
    parser.add_argument("--input", type=Path, default=Path("data/customers.csv"), help="客户 CSV/XLSX 表格")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="配置文件")
    parser.add_argument("--headed", action="store_true", help="显示浏览器窗口（默认行为，便于验证码人工接管）")
    parser.add_argument("--headless", action="store_true", help="后台运行；默认显示浏览器窗口")
    parser.add_argument("--auto-code", action="store_true", help="尝试从邮箱页面自动识别唯一数字验证码")
    parser.add_argument("--max-customers", type=int, default=0, help="最多处理多少条，0 表示全部")
    parser.add_argument("--recycle-every", type=int, default=20, help="每处理多少条后重启浏览器，0 表示不重启")
    return parser.parse_args()


def main() -> int:
    # The worker is launched without a console by the GUI. Force one stable
    # encoding so Chinese audit messages are decoded correctly on every
    # Windows system, independent of its active code page.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    args = parse_args()
    try:
        config = load_config(args.config)
        raw_rows = load_rows(args.input)
    except Exception as exc:
        print(f"启动失败：{exc}", file=sys.stderr)
        return 2
    customers = [Customer.from_row(i + 2, row) for i, row in enumerate(raw_rows)]
    customers = [customer for customer in customers if customer.enabled]
    if args.max_customers:
        customers = customers[: args.max_customers]
    if not customers:
        print("没有可处理的客户记录，请先填写客户表格。")
        return 0
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    audit = Audit(Path("runs") / run_id, args.input, len(customers))
    audit.event(None, "run_started", f"客户数={len(customers)}，模式=正式执行")
    bot = NpaBot(config, audit, auto_code=args.auto_code)
    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.launch(channel="msedge", headless=args.headless)
        except Exception as exc:
            print("启动失败：未找到 Microsoft Edge，请先安装 Edge 浏览器。", file=sys.stderr)
            return 2
        processed = 0
        for customer in customers:
            if args.recycle_every and processed and processed % args.recycle_every == 0:
                browser.close()
                browser = playwright.chromium.launch(channel="msedge", headless=args.headless)
                audit.event(None, "browser_recycled", f"已处理客户数={processed}")
            processed += 1
            audit.set_current(processed, customer)
            audit.event(customer, "customer_started", f"第 {processed}/{len(customers)} 条")
            if not customer.portal_email or not customer.mail_email or not customer.mail_password:
                audit.event(customer, "customer_skipped", "缺少 NPA 邮箱、邮箱账号或邮箱密码", "warning")
                audit.result(
                    customer,
                    status="manual_required",
                    admin_action="",
                    invoice_action="",
                    failed_step="输入校验",
                    error="缺少 NPA 邮箱、邮箱账号或邮箱密码",
                )
                continue
            route = resolve_mail_route(customer.mail_email, config["mail_routes"])
            if route is None:
                audit.event(customer, "customer_skipped", f"无法匹配邮箱入口：{customer.mail_email}", "warning")
                audit.result(
                    customer,
                    status="manual_required",
                    admin_action="",
                    invoice_action="",
                    failed_step="邮箱入口匹配",
                    error="无法匹配邮箱入口",
                )
                continue
            if route.get("ambiguous_alias"):
                audit.event(customer, "mail_route_alias_match", str(route.get("matched_by")), "warning")
            context = browser.new_context()
            try:
                result = bot.process(customer, context, route)
                audit.result(customer, **result)
            finally:
                context.close()
        browser.close()
    audit.status = "completed"
    audit.event(None, "run_finished", f"结果数={len(audit.results)}")
    audit.write_results()
    print(f"结果目录：{audit.root.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
