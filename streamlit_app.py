"""
streamlit_app.py

Customer-friendly Streamlit UI for export_castordoc.py.

Usage with uv:

    uv run streamlit run streamlit_app.py

Dependencies:

    uv add streamlit playwright
    uv run playwright install chromium
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="CastorDoc Exporter",
    page_icon="📚",
    layout="wide",
)


# ============================================================
# CONSTANTS
# ============================================================

DEFAULT_OUTPUT_ROOT = "Output"
DEFAULT_PROFILE_DIR = "playwright_profile"
DEFAULT_MAX_PAGES = 200
DEFAULT_MAX_LINK_DEPTH = 2


# ============================================================
# HELPERS
# ============================================================


def get_latest_export_folder(output_root: Path) -> Path | None:
    if not output_root.exists():
        return None

    folders = [
        path
        for path in output_root.iterdir()
        if path.is_dir() and path.name.startswith("castordoc_export_")
    ]

    if not folders:
        return None

    return max(folders, key=lambda path: path.stat().st_mtime)


def list_export_folders(output_root: Path) -> list[Path]:
    if not output_root.exists():
        return []

    return sorted(
        [
            path
            for path in output_root.iterdir()
            if path.is_dir() and path.name.startswith("castordoc_export_")
        ],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def list_txt_files(folder_path: Path) -> list[Path]:
    if not folder_path or not folder_path.exists():
        return []

    return sorted(folder_path.glob("*.txt"))


def zip_folder(folder_path: Path) -> Path:
    zip_path = folder_path.with_suffix(".zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in folder_path.rglob("*"):
            if file_path.is_file():
                zip_file.write(file_path, file_path.relative_to(folder_path))

    return zip_path


def read_text_file(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1")
    except Exception as error:
        return f"Could not read file: {error}"


def parse_urls_from_text(text: str) -> list[str]:
    urls = []

    for line in text.splitlines():
        clean_line = line.strip()

        if not clean_line:
            continue

        if clean_line.startswith("#"):
            continue

        urls.append(clean_line)

    return list(dict.fromkeys(urls))


def build_command(
    python_runner: str,
    exporter_path: Path,
    root_urls_file: Path,
    output_root: str,
    profile_dir: str,
    max_pages: int,
    max_link_depth: int,
) -> list[str]:
    if python_runner == "uv run python":
        command_prefix = ["uv", "run", "python"]
    elif python_runner == "current python":
        command_prefix = [sys.executable]
    else:
        command_prefix = ["python"]

    return [
        *command_prefix,
        str(exporter_path),
        "--root-urls-file",
        str(root_urls_file),
        "--output-root",
        output_root,
        "--profile-dir",
        profile_dir,
        "--max-pages",
        str(max_pages),
        "--max-link-depth",
        str(max_link_depth),
    ]


def show_summary_cards(summary_text: str, fallback_file_count: int) -> None:
    metrics = {
        "Root pages": "0",
        "Subpages": "0",
        "ReadMe linked pages": "0",
        "Saved files": str(fallback_file_count),
    }

    for line in summary_text.splitlines():
        clean = line.strip()

        if clean.startswith("- root_saved:"):
            metrics["Root pages"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("- subpage_saved:"):
            metrics["Subpages"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("- readme_link_saved:"):
            metrics["ReadMe linked pages"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("SAVED_URLS:"):
            metrics["Saved files"] = clean.split(":", 1)[1].strip()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Root pages", metrics["Root pages"])

    with col2:
        st.metric("Subpages", metrics["Subpages"])

    with col3:
        st.metric("Linked pages", metrics["ReadMe linked pages"])

    with col4:
        st.metric("Saved files", metrics["Saved files"])


# ============================================================
# HEADER
# ============================================================

st.title("📚 CastorDoc Documentation Exporter")

st.markdown(
    """
Use this tool to export CastorDoc / Coalesce documentation into clean text files.

It saves:
- the **ReadMe of the page you provide**;
- the **ReadMe of its subpages**;
- the **ReadMe of useful CastorDoc links found inside those pages**.
"""
)


# ============================================================
# SIMPLE USER FLOW
# ============================================================

st.markdown("## 1. Log in to CastorDoc")

st.info(
    "When you start an export, a Chromium browser will open. "
    "If CastorDoc asks you to log in, log in manually. "
    "After that, the export continues automatically."
)

st.markdown("## 2. Choose the documentation page(s) to export")

st.markdown(
    """
Paste one or more CastorDoc URLs below.

Use the page that contains the main requirements and subpages, usually a `/map` page.
The exporter will also save that root page's own ReadMe.
"""
)

root_urls_text = st.text_area(
    "CastorDoc page URL(s)",
    value="https://app.castordoc.com/terms/internal/fact-production-work-e7cef424/map",
    height=150,
    placeholder=(
        "Paste one URL per line, for example:\n"
        "https://app.castordoc.com/terms/internal/fact-production-work-e7cef424/map"
    ),
)

with st.expander("What will be exported?", expanded=False):
    st.markdown(
        """
For each URL you paste, the tool will:

1. Open the page and save its **ReadMe**.
2. Open **Subpages & Map** and find direct subpages.
3. Save the **ReadMe** of each subpage.
4. Look inside each ReadMe for other CastorDoc documentation links.
5. Save the **ReadMe** of those linked pages too.
6. Skip duplicates automatically.
"""
    )


# ============================================================
# ADVANCED SETTINGS
# ============================================================

with st.expander("Advanced settings", expanded=False):
    output_root = st.text_input(
        "Export folder",
        value=DEFAULT_OUTPUT_ROOT,
        help="Folder where exports will be created.",
    )

    profile_dir = st.text_input(
        "Browser login profile",
        value=DEFAULT_PROFILE_DIR,
        help="Keeps your CastorDoc login session between runs.",
    )

    max_pages = st.number_input(
        "Safety limit: maximum pages to visit",
        min_value=1,
        max_value=1000,
        value=DEFAULT_MAX_PAGES,
        step=25,
    )

    crawl_mode = st.radio(
        "How far should linked pages be followed?",
        options=[
            "Root + subpages only",
            "Root + subpages + links inside ReadMe",
            "Deeper crawl",
        ],
        index=1,
    )

    if crawl_mode == "Root + subpages only":
        max_link_depth = 0
    elif crawl_mode == "Root + subpages + links inside ReadMe":
        max_link_depth = 1
    else:
        max_link_depth = DEFAULT_MAX_LINK_DEPTH

    python_runner = st.selectbox(
        "Python runner",
        options=["uv run python", "current python", "python"],
        index=0,
        help="Use `uv run python` if you run the project with Astral uv.",
    )

    show_logs = st.checkbox(
        "Show technical logs while exporting",
        value=True,
    )


# ============================================================
# EXPORT BUTTON
# ============================================================

st.markdown("## 3. Export")

run_export = st.button("🚀 Export documentation", type="primary")

status_box = st.empty()
progress_bar = st.progress(0)
log_box = st.empty()

if run_export:
    root_urls = parse_urls_from_text(root_urls_text)

    if not root_urls:
        st.error("Please paste at least one CastorDoc URL.")
        st.stop()

    exporter_path = Path("export_castordoc.py")

    if not exporter_path.exists():
        st.error("`export_castordoc.py` was not found in the current folder.")
        st.stop()

    output_root_path = Path(output_root)

    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".txt",
        delete=False,
        encoding="utf-8",
    ) as temp_file:
        for url in root_urls:
            temp_file.write(url + "\n")

        root_urls_file = Path(temp_file.name)

    command = build_command(
        python_runner=python_runner,
        exporter_path=exporter_path,
        root_urls_file=root_urls_file,
        output_root=output_root,
        profile_dir=profile_dir,
        max_pages=int(max_pages),
        max_link_depth=int(max_link_depth),
    )

    status_box.info(
        "Export started. A Chromium window should open. "
        "Log in manually if CastorDoc asks for authentication."
    )

    logs: list[str] = []
    start_time = time.time()

    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            env=os.environ.copy(),
        )

        while True:
            if process.stdout is None:
                break

            line = process.stdout.readline()

            if line:
                clean_line = line.rstrip()
                logs.append(clean_line)

                visited_count = sum(
                    1 for item in logs if item.startswith("Visite ")
                )

                progress = min(visited_count / int(max_pages), 1.0)
                progress_bar.progress(progress)

                if show_logs:
                    log_box.text_area(
                        "Export logs",
                        value="\n".join(logs[-400:]),
                        height=420,
                    )
                else:
                    log_box.info(f"Pages visited: {visited_count}")

            if process.poll() is not None:
                break

        return_code = process.returncode
        elapsed = round(time.time() - start_time, 1)

        try:
            root_urls_file.unlink(missing_ok=True)
        except Exception:
            pass

        latest_folder = get_latest_export_folder(output_root_path)

        if return_code == 0:
            status_box.success(f"Export completed in {elapsed}s.")
        else:
            status_box.error(f"Exporter finished with return code {return_code}.")

        if not latest_folder:
            st.warning("No export folder found.")
            st.stop()

        txt_files = list_txt_files(latest_folder)
        summary_file = latest_folder / "export_summary.txt"

        st.markdown("## Export result")

        summary_text = ""
        if summary_file.exists():
            summary_text = read_text_file(summary_file)

        show_summary_cards(summary_text, fallback_file_count=len(txt_files))

        st.markdown("### Export folder")
        st.code(str(latest_folder), language="text")

        if summary_file.exists():
            with st.expander("Show export summary", expanded=False):
                st.text_area(
                    "export_summary.txt",
                    value=summary_text,
                    height=260,
                )

        if txt_files:
            st.markdown("### Preview exported files")

            selected_file = st.selectbox(
                "Choose a file to preview",
                options=txt_files,
                format_func=lambda path: path.name,
            )

            if selected_file:
                st.text_area(
                    selected_file.name,
                    value=read_text_file(selected_file),
                    height=420,
                )

            zip_path = zip_folder(latest_folder)

            with open(zip_path, "rb") as file:
                st.download_button(
                    label="⬇️ Download ZIP",
                    data=file,
                    file_name=zip_path.name,
                    mime="application/zip",
                )
        else:
            st.warning(
                "No `.txt` files were exported. Check the summary and logs."
            )

    except FileNotFoundError:
        st.error(
            "Command failed. If you selected `uv run python`, make sure `uv` "
            "is installed and available in your terminal."
        )

    except Exception as error:
        st.exception(error)

        try:
            root_urls_file.unlink(missing_ok=True)
        except Exception:
            pass


# ============================================================
# PREVIOUS EXPORTS
# ============================================================

st.markdown("---")
st.markdown("## Previous exports")

output_root_path = Path(DEFAULT_OUTPUT_ROOT)
export_folders = list_export_folders(output_root_path)

if not export_folders:
    st.info("No previous export found yet.")
else:
    selected_export = st.selectbox(
        "Open a previous export",
        options=export_folders,
        format_func=lambda path: path.name,
    )

    if selected_export:
        previous_files = list_txt_files(selected_export)
        previous_summary = selected_export / "export_summary.txt"

        col1, col2 = st.columns(2)

        with col1:
            st.write(f"Folder: `{selected_export}`")

        with col2:
            st.write(f"Files: `{len(previous_files)}`")

        if previous_summary.exists():
            with st.expander("Show previous export summary", expanded=False):
                st.text_area(
                    "Previous summary",
                    value=read_text_file(previous_summary),
                    height=220,
                )

        if previous_files:
            selected_previous_file = st.selectbox(
                "Preview previous exported file",
                options=previous_files,
                format_func=lambda path: path.name,
                key="previous_file",
            )

            st.text_area(
                selected_previous_file.name,
                value=read_text_file(selected_previous_file),
                height=350,
            )

            previous_zip_path = zip_folder(selected_export)

            with open(previous_zip_path, "rb") as file:
                st.download_button(
                    label="⬇️ Download previous export ZIP",
                    data=file,
                    file_name=previous_zip_path.name,
                    mime="application/zip",
                )