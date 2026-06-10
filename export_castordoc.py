"""
export_castordoc.py

CastorDoc / Coalesce documentation exporter.

Purpose
-------
This script opens CastorDoc with Playwright, lets the user log in manually,
then exports documentation pages as clean .txt files.

It is intentionally responsible only for:
- browser automation;
- CastorDoc navigation;
- ReadMe extraction;
- link crawling;
- TXT export;
- optional call to Excel generation after the TXT export is complete.

It does NOT build the Excel workbook directly. That logic lives in:
    excel_exporter.py

Main crawl rules
----------------
For each page:
1. Go to /home.
2. Save only the ReadMe content.
3. Follow only links found inside the ReadMe content.
4. If depth allows it, go to Subpages & Map and collect Knowledge tiles.
5. If a page has "0 Subpages & Map", keep only the ReadMe and do not crawl tiles.
6. Avoid breadcrumbs, top menu, sidebars, comments, history, etc.

Default depth behaviour
-----------------------
Depth 0, 1, 2:
    - save ReadMe;
    - follow ReadMe links;
    - explore Subpages & Map Knowledge tiles.

Depth 3:
    - save ReadMe only;
    - do not follow more ReadMe links;
    - do not explore Subpages.

Usage with uv
-------------
Single root page:

    uv run python export_castordoc.py \
      --root-url "https://app.castordoc.com/terms/internal/xxx/home" \
      --generate-excel

Multiple root pages:

    uv run python export_castordoc.py \
      --root-url "URL_1" \
      --root-url "URL_2" \
      --generate-excel

From a text file:

    uv run python export_castordoc.py \
      --root-urls-file root_urls.txt \
      --generate-excel
"""

from __future__ import annotations

import argparse
import re
import traceback
from collections import deque
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urldefrag, urljoin, urlparse

from playwright.sync_api import (
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)

from excel_exporter import build_excel_overview


# ============================================================
# ARGUMENTS
# ============================================================


def parse_args() -> argparse.Namespace:
    """
    Parse CLI arguments.

    The script can be used directly from terminal or indirectly from Streamlit.
    Streamlit writes root URLs into a temporary file and passes that file here.
    """
    parser = argparse.ArgumentParser(
        description="Export CastorDoc / Coalesce ReadMe documentation pages."
    )

    parser.add_argument(
        "--root-url",
        action="append",
        default=[],
        help=(
            "CastorDoc root documentation URL. Can be passed multiple times. "
            "Usually the /map or /home page containing the main requirements."
        ),
    )

    parser.add_argument(
        "--root-urls-file",
        default=None,
        help="Optional text file containing one CastorDoc root URL per line.",
    )

    parser.add_argument(
        "--base-url",
        default="https://app.castordoc.com",
        help="CastorDoc base URL.",
    )

    parser.add_argument(
        "--profile-dir",
        default="playwright_profile",
        help=(
            "Persistent Playwright browser profile directory. "
            "This keeps the login session between runs."
        ),
    )

    parser.add_argument(
        "--output-root",
        default="Output",
        help="Root output folder where timestamped exports are created.",
    )

    parser.add_argument(
        "--max-pages",
        type=int,
        default=250,
        help="Safety limit: maximum number of pages to visit.",
    )

    parser.add_argument(
        "--max-depth",
        type=int,
        default=3,
        help=(
            "Maximum crawl depth. Default 3. "
            "At max depth, the page ReadMe is saved but no further links/subpages are explored."
        ),
    )

    parser.add_argument(
        "--subpages-until-depth",
        type=int,
        default=2,
        help="Explore Subpages & Map only for pages with depth <= this value.",
    )

    parser.add_argument(
        "--generate-excel",
        action="store_true",
        help="Generate an Excel overview workbook after the TXT export is complete.",
    )

    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run browser in headless mode. Default is visible browser.",
    )

    return parser.parse_args()


# ============================================================
# GLOBAL CONFIG SET AT RUNTIME
# ============================================================

BASE_URL = "https://app.castordoc.com"
PROFILE_DIR = "playwright_profile"
OUTPUT_DIR = Path("Output")

MAX_PAGES = 250
MAX_DEPTH = 3
SUBPAGES_UNTIL_DEPTH = 2

GOTO_TIMEOUT_MS = 60_000
DEFAULT_TIMEOUT_MS = 10_000


# ============================================================
# DATA STRUCTURES
# ============================================================


@dataclass(frozen=True)
class CrawlItem:
    """
    One URL waiting to be visited.

    source_type:
        root            = URL pasted by user
        readme_link     = URL found inside a ReadMe
        knowledge_tile  = URL found under Subpages & Map / Knowledge
    """

    url: str
    source_type: str
    depth: int
    parent_url: str | None = None


@dataclass
class ReadMeExtraction:
    """
    Result of extracting a page's ReadMe.

    content:
        Clean business documentation text.

    readme_links:
        CastorDoc links found only inside the ReadMe content area.
    """

    title: str
    content: str
    readme_links: list[str]


# ============================================================
# LOGGING
# ============================================================


def log(message: str) -> None:
    """Print logs immediately so Streamlit can display them live."""
    print(message, flush=True)


# ============================================================
# URL + FILE HELPERS
# ============================================================


def clean_filename(text: str) -> str:
    """Convert a page title into a safe filename fragment."""
    text = text or "page"
    text = text.strip()
    text = re.sub(r"[^\w\-À-ÿ ]", "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text[:120].strip("_") or "page"


def normalize_url(href: str | None) -> str | None:
    """
    Normalize a CastorDoc URL.

    Important:
    - /map is useful for discovering subpages.
    - /home is useful for exporting the ReadMe.
    - The crawler stores canonical /home URLs to avoid duplicates.
    """
    if not href:
        return None

    full_url = urljoin(BASE_URL, href)
    full_url, _ = urldefrag(full_url)
    full_url = full_url.rstrip("/")
    full_url = re.sub(r"/map$", "/home", full_url)

    return full_url


def as_map_url(url: str) -> str:
    """Convert a canonical /home URL into /map for Subpages & Map exploration."""
    normalized = normalize_url(url) or url
    return re.sub(r"/home$", "/map", normalized)


def is_castordoc_internal_term_url(url: str | None) -> bool:
    """
    Keep only useful CastorDoc internal documentation pages.

    This avoids:
    - external links;
    - SAP generated technical links;
    - lineage/query/settings/comments/history pages.
    """
    if not url:
        return False

    if not url.startswith(BASE_URL):
        return False

    parsed = urlparse(url)
    path = parsed.path

    if "/terms/internal/" not in path:
        return False

    # Ignore SAP / technical generated internal pages.
    # Example: /terms/internal/-xxxx/...
    if re.search(r"/terms/internal/-[^/]+", path):
        return False

    # Keep only documentation page endpoints.
    if not (path.endswith("/home") or path.endswith("/map")):
        return False

    ignored_keywords = [
        "/lineage",
        "/query",
        "/columns",
        "/settings",
        "/comments",
        "/history",
    ]

    if any(keyword in path.lower() for keyword in ignored_keywords):
        return False

    return True


def read_root_urls_from_file(path: str | None) -> list[str]:
    """Read one root URL per line from a text file."""
    if not path:
        return []

    file_path = Path(path)

    if not file_path.exists():
        raise FileNotFoundError(f"Root URLs file not found: {file_path}")

    urls = []

    for line in file_path.read_text(encoding="utf-8").splitlines():
        clean_line = line.strip()

        if not clean_line:
            continue

        if clean_line.startswith("#"):
            continue

        urls.append(clean_line)

    return urls


def dedupe_keep_order(values: list[str]) -> list[str]:
    """Remove duplicates without changing order."""
    seen = set()
    result = []

    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)

    return result


def unique_output_path(title: str, counter: int, source_type: str, depth: int) -> Path:
    """Build a stable output path for one exported documentation page."""
    clean_title = clean_filename(title)
    clean_source_type = clean_filename(source_type)
    filename = f"{counter:03d}_d{depth}_{clean_source_type}_{clean_title}.txt"
    path = OUTPUT_DIR / filename

    suffix = 2

    while path.exists():
        filename = f"{counter:03d}_d{depth}_{clean_source_type}_{clean_title}_{suffix}.txt"
        path = OUTPUT_DIR / filename
        suffix += 1

    return path


# ============================================================
# PLAYWRIGHT HELPERS
# ============================================================


def safe_goto(page, url: str, label: str = "") -> bool:
    """
    Navigate to a URL without crashing the whole export on timeout.

    Returns True if navigation looked successful, False otherwise.
    """
    label_text = f" [{label}]" if label else ""

    try:
        log(f"Navigation{label_text}: {url}")
        page.goto(url, wait_until="domcontentloaded", timeout=GOTO_TIMEOUT_MS)
        page.wait_for_load_state("domcontentloaded", timeout=15_000)
        page.wait_for_timeout(2_000)
        return True

    except PlaywrightTimeoutError as error:
        log(f"Timeout navigation{label_text}: {url}")
        log(f"  Detail: {error}")
        page.wait_for_timeout(5_000)
        return False

    except PlaywrightError as error:
        log(f"Erreur Playwright navigation{label_text}: {url}")
        log(f"  Detail: {error}")
        page.wait_for_timeout(5_000)
        return False

    except Exception as error:
        log(f"Erreur inattendue navigation{label_text}: {url}")
        log(f"  Detail: {error}")
        page.wait_for_timeout(5_000)
        return False


def wait_until_logged_in(page, timeout_seconds: int = 240) -> None:
    """
    Wait for manual CastorDoc login.

    Browser is visible by default, so the user can authenticate manually.
    The persistent profile normally prevents needing to log in again later.
    """
    log("")
    log("Connexion CastorDoc")
    log("Si une page de login apparaît, connecte-toi manuellement dans Chromium.")
    log("Le script continuera automatiquement après détection de CastorDoc.")

    for second in range(timeout_seconds):
        current_url = page.url or ""

        try:
            body_text = page.locator("body").inner_text(timeout=2_000).lower()
        except Exception:
            body_text = ""

        looks_logged_in = (
            current_url.startswith(BASE_URL)
            and "/login" not in current_url.lower()
            and "sign in" not in body_text
            and "log in" not in body_text
            and "connexion" not in body_text
        )

        if looks_logged_in:
            log("Login détecté.")
            return

        if second % 15 == 0:
            log(f"Attente login... {second}s")

        page.wait_for_timeout(1_000)

    raise TimeoutError("Login non détecté dans le délai imparti.")


def safe_click_text(page, pattern: str, timeout: int = 5_000) -> bool:
    """Click the first visible element matching text regex. Return False if not found."""
    try:
        page.get_by_text(re.compile(pattern, re.IGNORECASE)).first.click(timeout=timeout)
        page.wait_for_timeout(1_200)
        return True
    except Exception:
        return False


def click_readme_tab(page) -> None:
    """
    Click the ReadMe tab if visible.

    If not found, the script still tries to extract from the current page.
    """
    clicked = (
        safe_click_text(page, r"^Read\s*Me$")
        or safe_click_text(page, r"ReadMe")
        or safe_click_text(page, r"README")
    )

    if clicked:
        log("Onglet ReadMe ouvert.")
    else:
        log("Onglet ReadMe non trouvé. Extraction depuis la page actuelle.")

    page.wait_for_timeout(1_000)


def get_subpages_count_from_page(page) -> int | None:
    """
    Read the "Subpages & Map" tab label.

    Examples:
    - "0 Subpages & Map"   -> 0
    - "30+ Subpages & Map" -> 30
    - "Subpages & Map"     -> None
    """
    try:
        labels = page.evaluate(
            """
            () => Array.from(document.querySelectorAll("button, a, div, span"))
                .map(el => (el.innerText || el.textContent || "").trim())
                .filter(Boolean)
                .filter(text =>
                    text.toLowerCase().includes("subpages") &&
                    text.toLowerCase().includes("map")
                );
            """
        )
    except Exception:
        return None

    for label in labels:
        match = re.search(r"(\d+)\+?\s*Subpages\s*&\s*Map", label, re.IGNORECASE)
        if match:
            return int(match.group(1))

    return None


def click_subpages_tab(page) -> bool:
    """
    Click Subpages & Map.

    Returns:
    - True if clicked.
    - False if not found.

    Note:
    The page may show "0 Subpages & Map". The caller checks the count before
    collecting Knowledge tiles.
    """
    clicked = (
        safe_click_text(page, r"\d+\+?\s*Subpages\s*&\s*Map")
        or safe_click_text(page, r"Subpages\s*&\s*Map")
        or safe_click_text(page, r"Subpages")
        or safe_click_text(page, r"Map")
    )

    if clicked:
        log("Onglet Subpages/Map ouvert.")
        page.wait_for_timeout(1_500)
        return True

    log("Onglet Subpages/Map non trouvé.")
    return False


def get_page_title(page) -> str:
    """Best-effort page title extraction."""
    selectors = [
        "h1",
        "[data-testid*='title']",
        "[class*='title']",
    ]

    for selector in selectors:
        try:
            value = page.locator(selector).first.inner_text(timeout=3_000).strip()
            if value and len(value) <= 200:
                return value
        except Exception:
            pass

    try:
        return page.title().strip() or "page"
    except Exception:
        return "page"


# ============================================================
# README-SCOPED EXTRACTION
# ============================================================


def extract_readme_content_and_links(page) -> ReadMeExtraction:
    """
    Extract only the ReadMe business content and links inside that content.

    This deliberately avoids extracting links from the full page DOM because
    that caused drift into sidebars, breadcrumbs, top navigation, comments, etc.
    """
    click_readme_tab(page)
    title = get_page_title(page)

    result = page.evaluate(
        """
        () => {
            function visible(el) {
                if (!el) return false;
                const rect = el.getBoundingClientRect();
                const style = window.getComputedStyle(el);
                return (
                    rect.width > 0 &&
                    rect.height > 0 &&
                    style.visibility !== "hidden" &&
                    style.display !== "none"
                );
            }

            function cleanText(text) {
                return (text || "")
                    .replace(/\\r/g, "")
                    .replace(/\\n{3,}/g, "\\n\\n")
                    .trim();
            }

            function scoreElement(el) {
                const text = cleanText(el.innerText || "");
                if (!text) return -999999;

                const lower = text.toLowerCase();
                let score = text.length;

                // Positive signs that this is the business documentation area.
                if (lower.includes("coalesce catalog description")) score += 2000;
                if (lower.includes("definition & purpose")) score += 1500;
                if (lower.includes("business concepts")) score += 1000;
                if (lower.includes("functional tests")) score += 1000;
                if (lower.includes("examples of values")) score += 1000;
                if (lower.includes("concept :")) score += 500;
                if (lower.includes("concepts :")) score += 500;

                // Negative signs that this is navigation or metadata.
                if (lower.includes("subpages & map")) score -= 2000;
                if (lower.includes("comments")) score -= 500;
                if (lower.includes("history")) score -= 500;
                if (lower.includes("knowledge >")) score -= 800;
                if (lower.includes("coalesce catalog >")) score -= 800;

                const links = Array.from(el.querySelectorAll("a[href]"));
                if (links.length > 40 && text.length < 4000) score -= 1500;

                return score;
            }

            const candidates = Array.from(
                document.querySelectorAll(
                    [
                        "main",
                        "article",
                        "[role='main']",
                        "section",
                        "div[class*='content']",
                        "div[class*='Content']",
                        "div[class*='description']",
                        "div[class*='Description']",
                        "div[class*='markdown']",
                        "div[class*='Markdown']",
                        "div"
                    ].join(",")
                )
            )
            .filter(visible)
            .map(el => ({ el, score: scoreElement(el) }))
            .filter(item => item.score > 0)
            .sort((a, b) => b.score - a.score);

            let container = candidates.length ? candidates[0].el : document.body;
            let text = cleanText(container.innerText || "");

            // Remove obvious top navigation prefix if present.
            const startMarkers = [
                "Coalesce Catalog Description",
                "🎯 Definition & purpose",
                "Definition & purpose",
                "Functional Tests Definition",
                "Business concepts & attributes",
                "Concept :",
                "Concepts :",
                "Read Me"
            ];

            let startIndex = -1;

            for (const marker of startMarkers) {
                const idx = text.indexOf(marker);
                if (idx !== -1 && (startIndex === -1 || idx < startIndex)) {
                    startIndex = idx;
                }
            }

            if (startIndex > 0) {
                text = text.slice(startIndex).trim();
            }

            // Cut obvious page metadata / right panel / bottom sections.
            const stopMarkers = [
                "\\nDetails",
                "\\nOwners",
                "\\nDomain",
                "\\nTags",
                "\\nComments",
                "\\nHistory",
                "\\nLineage",
                "\\nColumns"
            ];

            for (const marker of stopMarkers) {
                const idx = text.indexOf(marker);
                if (idx > 300) {
                    text = text.slice(0, idx).trim();
                    break;
                }
            }

            const links = Array.from(container.querySelectorAll("a[href]"))
                .filter(visible)
                .map(a => ({
                    href: a.href || a.getAttribute("href"),
                    text: cleanText(a.innerText || a.textContent || "")
                }))
                .filter(item => {
                    const lowerText = item.text.toLowerCase();

                    if (!item.href) return false;
                    if (!item.href.includes("/terms/internal/")) return false;

                    // Avoid tabs and common UI links.
                    if ([
                        "read me",
                        "comments",
                        "history",
                        "details",
                        "owners",
                        "domain",
                        "tags",
                        "lineage",
                        "columns"
                    ].includes(lowerText)) {
                        return false;
                    }

                    return true;
                });

            return {
                text,
                links
            };
        }
        """
    )

    raw_content = result.get("text", "") if isinstance(result, dict) else ""
    raw_links = result.get("links", []) if isinstance(result, dict) else []

    content = clean_extracted_readme_text(raw_content)

    links: list[str] = []

    for item in raw_links:
        if isinstance(item, dict):
            href = item.get("href")
        else:
            href = str(item)

        url = normalize_url(href)

        if is_castordoc_internal_term_url(url):
            links.append(url)

    return ReadMeExtraction(
        title=title,
        content=content,
        readme_links=dedupe_keep_order(links),
    )


def clean_extracted_readme_text(text: str) -> str:
    """Post-process ReadMe text to remove remaining UI fragments."""
    if not text:
        return ""

    text = text.replace("\r", "")
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = text.strip()

    text = re.sub(
        r"^Read Me\s+\d+\+?\s*Subpages\s*&\s*Map\s+Comments\s+History\s*",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()

    lines = text.splitlines()
    cleaned_lines = []

    for line in lines:
        clean_line = line.strip()

        if not clean_line:
            cleaned_lines.append(line)
            continue

        lower = clean_line.lower()

        if lower.startswith("knowledge >"):
            continue

        if "coalesce catalog >" in lower:
            continue

        cleaned_lines.append(line)

    text = "\n".join(cleaned_lines)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


# ============================================================
# SUBPAGES / KNOWLEDGE TILE EXTRACTION
# ============================================================


def collect_knowledge_tile_links(page, page_url: str) -> list[str]:
    """
    Open Subpages & Map and collect only visible CastorDoc Knowledge tiles/cards.

    Rules:
    - If the page has 0 subpages, return [] immediately.
    - Do not collect "Create a subpage".
    - Do not collect "Knowledge Map".
    - Do not collect breadcrumbs/topbar/sidebar links.
    """
    subpage_count = get_subpages_count_from_page(page)

    if subpage_count == 0:
        log("0 Subpages détecté. On garde uniquement le ReadMe.")
        return []

    map_url = as_map_url(page_url)

    ok = safe_goto(page, map_url, "subpages map")
    if not ok:
        return []

    subpage_count = get_subpages_count_from_page(page)

    if subpage_count == 0:
        log("0 Subpages détecté sur la page map. On garde uniquement le ReadMe.")
        return []

    clicked = click_subpages_tab(page)

    if not clicked:
        return []

    found_links: list[str] = []

    for scroll_round in range(30):
        page.wait_for_timeout(700)

        raw_links = page.evaluate(
            """
            () => {
                function visible(el) {
                    if (!el) return false;
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return (
                        rect.width > 0 &&
                        rect.height > 0 &&
                        style.visibility !== "hidden" &&
                        style.display !== "none"
                    );
                }

                function cleanText(text) {
                    return (text || "")
                        .replace(/\\r/g, "")
                        .replace(/\\n{2,}/g, "\\n")
                        .trim();
                }

                const links = Array.from(document.querySelectorAll("a[href]"))
                    .filter(visible)
                    .filter(a => (a.href || "").includes("/terms/internal/"))
                    .map(a => {
                        const text = cleanText(a.innerText || a.textContent || "");
                        const rect = a.getBoundingClientRect();

                        let container = a;
                        for (let i = 0; i < 5; i++) {
                            if (!container.parentElement) break;
                            container = container.parentElement;
                        }

                        const containerText = cleanText(container.innerText || "");

                        return {
                            href: a.href,
                            text,
                            containerText,
                            x: rect.x,
                            y: rect.y,
                            width: rect.width,
                            height: rect.height
                        };
                    });

                return links.filter(item => {
                    const text = item.text.toLowerCase();
                    const containerText = item.containerText.toLowerCase();

                    // Explicitly ignore CastorDoc utility cards.
                    if (containerText.includes("create a subpage")) return false;
                    if (containerText.includes("knowledge map")) return false;
                    if (containerText.includes("access to the lineage")) return false;

                    // Ignore tab/nav/common UI links.
                    const badText = [
                        "read me",
                        "comments",
                        "history",
                        "subpages & map",
                        "knowledge",
                        "coalesce catalog",
                        "create a subpage",
                        "knowledge map"
                    ].includes(text);

                    if (badText) return false;

                    // Real Knowledge tiles usually contain one of these markers.
                    const looksLikeKnowledgeTile =
                        containerText.includes("draft") ||
                        containerText.includes("domain") ||
                        containerText.includes("edited") ||
                        containerText.includes("no description") ||
                        containerText.includes("documentation_");

                    if (!looksLikeKnowledgeTile) return false;

                    return true;
                });
            }
            """
        )

        for item in raw_links:
            href = item.get("href") if isinstance(item, dict) else None
            url = normalize_url(href)

            if is_castordoc_internal_term_url(url) and url not in found_links:
                found_links.append(url)
                label = item.get("text", "") if isinstance(item, dict) else ""
                log(f"Knowledge tile trouvé: {label} -> {url}")

        page.mouse.wheel(0, 900)
        page.wait_for_timeout(700)

        log(f"Scan Knowledge tiles round {scroll_round + 1}/30 - total: {len(found_links)}")

    current_home = normalize_url(page_url)

    found_links = [
        url for url in found_links
        if url != current_home
    ]

    return dedupe_keep_order(found_links)


# ============================================================
# SAVE + SUMMARY
# ============================================================


def save_page(
    path: Path,
    title: str,
    url: str,
    source_type: str,
    depth: int,
    parent_url: str | None,
    content: str,
) -> None:
    """Save one ReadMe page as a .txt file with metadata header."""
    parent_line = parent_url or ""

    path.write_text(
        (
            f"TITLE: {title}\n"
            f"URL: {url}\n"
            f"SOURCE_TYPE: {source_type}\n"
            f"DEPTH: {depth}\n"
            f"PARENT_URL: {parent_line}\n"
            f"EXPORTED_AT: {datetime.now().isoformat(timespec='seconds')}\n\n"
            f"{content}\n"
        ),
        encoding="utf-8",
    )


def write_summary(
    root_urls: list[str],
    visited_urls: set[str],
    saved_urls: set[str],
    skipped_duplicate_urls: set[str],
    failed_urls: list[tuple[str, str]],
    counters: dict[str, int],
) -> None:
    """Write export_summary.txt for auditability and Excel generation."""
    summary_path = OUTPUT_DIR / "export_summary.txt"

    summary_lines = [
        "CASTORDOC EXPORT SUMMARY",
        f"OUTPUT_DIR: {OUTPUT_DIR}",
        f"ROOT_URLS: {len(root_urls)}",
        f"VISITED_URLS: {len(visited_urls)}",
        f"SAVED_URLS: {len(saved_urls)}",
        f"FAILED_URLS: {len(failed_urls)}",
        f"DUPLICATES_SKIPPED: {len(skipped_duplicate_urls)}",
        f"MAX_DEPTH: {MAX_DEPTH}",
        f"SUBPAGES_UNTIL_DEPTH: {SUBPAGES_UNTIL_DEPTH}",
        "",
        "COUNTERS:",
    ]

    for key in sorted(counters):
        summary_lines.append(f"- {key}: {counters[key]}")

    summary_lines.extend(["", "ROOT URLS:"])

    for url in root_urls:
        summary_lines.append(f"- {url}")

    summary_lines.extend(["", "FAILED DETAILS:"])

    for url, reason in failed_urls:
        summary_lines.append(f"- {url} | {reason}")

    summary_path.write_text("\n".join(summary_lines), encoding="utf-8")


# ============================================================
# CRAWL LOGIC
# ============================================================


def build_initial_queue(root_urls: list[str]) -> deque[CrawlItem]:
    """Build the initial queue from user-provided root URLs."""
    queue: deque[CrawlItem] = deque()

    for root_url in root_urls:
        normalized = normalize_url(root_url)

        if not is_castordoc_internal_term_url(normalized):
            log(f"URL racine ignorée car invalide: {root_url}")
            continue

        queue.append(
            CrawlItem(
                url=normalized,
                source_type="root",
                depth=0,
                parent_url=None,
            )
        )

    return queue


def add_to_queue(
    queue: deque[CrawlItem],
    queued_urls: set[str],
    visited_urls: set[str],
    saved_urls: set[str],
    skipped_duplicate_urls: set[str],
    url: str,
    source_type: str,
    depth: int,
    parent_url: str,
) -> bool:
    """Add a URL to the crawl queue if valid and not already processed."""
    normalized = normalize_url(url)

    if not is_castordoc_internal_term_url(normalized):
        return False

    if normalized in visited_urls or normalized in queued_urls or normalized in saved_urls:
        skipped_duplicate_urls.add(normalized)
        return False

    queue.append(
        CrawlItem(
            url=normalized,
            source_type=source_type,
            depth=depth,
            parent_url=parent_url,
        )
    )

    queued_urls.add(normalized)
    return True


def run_export(root_urls: list[str], headless: bool = False) -> None:
    """Main export loop."""
    visited_urls: set[str] = set()
    saved_urls: set[str] = set()
    queued_urls: set[str] = set()
    skipped_duplicate_urls: set[str] = set()
    failed_urls: list[tuple[str, str]] = []

    counters: dict[str, int] = {
        "root_saved": 0,
        "readme_link_saved": 0,
        "knowledge_tile_saved": 0,
        "empty_pages": 0,
        "navigation_failed": 0,
        "readme_links_discovered": 0,
        "readme_links_queued": 0,
        "knowledge_tiles_discovered": 0,
        "knowledge_tiles_queued": 0,
    }

    queue = build_initial_queue(root_urls)

    for item in queue:
        queued_urls.add(item.url)

    if not queue:
        raise ValueError("No valid CastorDoc root URL was provided.")

    first_url = queue[0].url

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=headless,
            viewport={"width": 1500, "height": 950},
            args=["--start-maximized"],
        )

        context.set_default_timeout(DEFAULT_TIMEOUT_MS)

        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT_MS)

        try:
            log("Ouverture CastorDoc...")
            safe_goto(page, first_url, "initial")
            wait_until_logged_in(page)

            log("")
            log(f"Queue initiale: {len(queue)} page(s).")

            while queue and len(visited_urls) < MAX_PAGES:
                item = queue.popleft()
                current_url = item.url

                if current_url in visited_urls:
                    skipped_duplicate_urls.add(current_url)
                    continue

                visited_urls.add(current_url)

                log("")
                log("=" * 80)
                log(f"Visite {len(visited_urls)} [{item.source_type}] depth={item.depth}")
                log(current_url)
                log("=" * 80)

                try:
                    ok = safe_goto(page, current_url, item.source_type)

                    if not ok:
                        counters["navigation_failed"] += 1
                        failed_urls.append((current_url, "navigation failed or timeout"))
                        continue

                    extraction = extract_readme_content_and_links(page)

                    if not extraction.content:
                        counters["empty_pages"] += 1
                        failed_urls.append((current_url, "empty ReadMe content"))
                        log(f"ReadMe vide ou contenu non détecté: {current_url}")

                    elif current_url not in saved_urls:
                        output_path = unique_output_path(
                            title=extraction.title,
                            counter=len(saved_urls) + 1,
                            source_type=item.source_type,
                            depth=item.depth,
                        )

                        save_page(
                            path=output_path,
                            title=extraction.title,
                            url=current_url,
                            source_type=item.source_type,
                            depth=item.depth,
                            parent_url=item.parent_url,
                            content=extraction.content,
                        )

                        saved_urls.add(current_url)

                        counter_key = f"{item.source_type}_saved"
                        if counter_key in counters:
                            counters[counter_key] += 1

                        log(f"Sauvegardé: {output_path}")

                    else:
                        skipped_duplicate_urls.add(current_url)
                        log(f"Déjà sauvegardé: {current_url}")

                    # Rule 1: follow only links found inside the ReadMe.
                    if item.depth < MAX_DEPTH:
                        readme_links = extraction.readme_links
                        counters["readme_links_discovered"] += len(readme_links)

                        added_readme_links = 0

                        for link_url in readme_links:
                            was_added = add_to_queue(
                                queue=queue,
                                queued_urls=queued_urls,
                                visited_urls=visited_urls,
                                saved_urls=saved_urls,
                                skipped_duplicate_urls=skipped_duplicate_urls,
                                url=link_url,
                                source_type="readme_link",
                                depth=item.depth + 1,
                                parent_url=current_url,
                            )

                            if was_added:
                                added_readme_links += 1

                        counters["readme_links_queued"] += added_readme_links
                        log(f"Liens ReadMe ajoutés: {added_readme_links}")

                    else:
                        log(
                            f"Depth max atteint ({MAX_DEPTH}). "
                            "Aucun lien ReadMe supplémentaire ajouté."
                        )

                    # Rule 2: explore Knowledge tiles only if depth allows it.
                    if item.depth <= SUBPAGES_UNTIL_DEPTH:
                        knowledge_links = collect_knowledge_tile_links(page, current_url)
                        counters["knowledge_tiles_discovered"] += len(knowledge_links)

                        added_knowledge_links = 0

                        for tile_url in knowledge_links:
                            was_added = add_to_queue(
                                queue=queue,
                                queued_urls=queued_urls,
                                visited_urls=visited_urls,
                                saved_urls=saved_urls,
                                skipped_duplicate_urls=skipped_duplicate_urls,
                                url=tile_url,
                                source_type="knowledge_tile",
                                depth=item.depth + 1,
                                parent_url=current_url,
                            )

                            if was_added:
                                added_knowledge_links += 1

                        counters["knowledge_tiles_queued"] += added_knowledge_links
                        log(f"Knowledge tiles ajoutés: {added_knowledge_links}")

                    else:
                        log(
                            f"Subpages non explorées car depth={item.depth} "
                            f"> {SUBPAGES_UNTIL_DEPTH}."
                        )

                    log(f"Liens restants en queue: {len(queue)}")

                except Exception as error:
                    error_message = f"{type(error).__name__}: {error}"
                    log(f"Erreur sur page: {current_url}")
                    log(error_message)
                    failed_urls.append((current_url, error_message))
                    traceback.print_exc()

            if len(visited_urls) >= MAX_PAGES:
                log("")
                log(f"Arrêt sécurité: MAX_PAGES atteint ({MAX_PAGES}).")

        finally:
            context.close()

    write_summary(
        root_urls=root_urls,
        visited_urls=visited_urls,
        saved_urls=saved_urls,
        skipped_duplicate_urls=skipped_duplicate_urls,
        failed_urls=failed_urls,
        counters=counters,
    )

    log("")
    log("=" * 80)
    log("Export terminé")
    log("=" * 80)
    log(f"Dossier: {OUTPUT_DIR}")
    log(f"Root pages: {counters['root_saved']}")
    log(f"ReadMe linked pages: {counters['readme_link_saved']}")
    log(f"Knowledge tile pages: {counters['knowledge_tile_saved']}")
    log(f"Pages visitées: {len(visited_urls)}")
    log(f"Pages sauvegardées: {len(saved_urls)}")
    log(f"Pages en erreur/vide: {len(failed_urls)}")
    log(f"Doublons ignorés: {len(skipped_duplicate_urls)}")
    log(f"Résumé: {OUTPUT_DIR / 'export_summary.txt'}")


def main() -> None:
    """Entry point."""
    global BASE_URL
    global PROFILE_DIR
    global OUTPUT_DIR
    global MAX_PAGES
    global MAX_DEPTH
    global SUBPAGES_UNTIL_DEPTH

    args = parse_args()

    BASE_URL = args.base_url.rstrip("/")
    PROFILE_DIR = args.profile_dir
    MAX_PAGES = args.max_pages
    MAX_DEPTH = args.max_depth
    SUBPAGES_UNTIL_DEPTH = args.subpages_until_depth

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    OUTPUT_DIR = Path(args.output_root) / f"castordoc_export_{timestamp}"
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    file_root_urls = read_root_urls_from_file(args.root_urls_file)
    all_root_urls = dedupe_keep_order(args.root_url + file_root_urls)

    if not all_root_urls:
        raise ValueError("No root URL provided. Use --root-url or --root-urls-file.")

    run_export(root_urls=all_root_urls, headless=args.headless)

    if args.generate_excel:
        excel_path = build_excel_overview(OUTPUT_DIR)
        log(f"Excel généré: {excel_path}")


if __name__ == "__main__":
    main()