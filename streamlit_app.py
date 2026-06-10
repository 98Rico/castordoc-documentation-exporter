# streamlit_app.py

"""
Customer-friendly Streamlit UI for the CastorDoc exporter.
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


st.set_page_config(
    page_title="CastorDoc Exporter",
    page_icon="📚",
    layout="wide",
)


DEFAULT_OUTPUT_ROOT = "Output"
DEFAULT_PROFILE_DIR = "playwright_profile"
DEFAULT_MAX_PAGES = 300
DEFAULT_MAX_DEPTH = 3
DEFAULT_SUBPAGES_UNTIL_DEPTH = 2

EXCEL_FILE_NAME = "castordoc_model_specification.xlsx"


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

    return sorted(
        path
        for path in folder_path.glob("*.txt")
        if path.name != "export_summary.txt"
    )


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
    max_depth: int,
    subpages_until_depth: int,
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
        "--max-depth",
        str(max_depth),
        "--subpages-until-depth",
        str(subpages_until_depth),
        "--generate-excel",
    ]


def parse_summary_metrics(summary_text: str, fallback_file_count: int) -> dict[str, str]:
    metrics = {
        "Root pages": "0",
        "ReadMe links": "0",
        "Knowledge tiles": "0",
        "Saved files": str(fallback_file_count),
        "Failed pages": "0",
    }

    for line in summary_text.splitlines():
        clean = line.strip()

        if clean.startswith("- root_saved:"):
            metrics["Root pages"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("- readme_link_saved:"):
            metrics["ReadMe links"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("- knowledge_tile_saved:"):
            metrics["Knowledge tiles"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("SAVED_URLS:"):
            metrics["Saved files"] = clean.split(":", 1)[1].strip()

        elif clean.startswith("FAILED_URLS:"):
            metrics["Failed pages"] = clean.split(":", 1)[1].strip()

    return metrics


def show_summary_cards(summary_text: str, fallback_file_count: int) -> None:
    metrics = parse_summary_metrics(summary_text, fallback_file_count)

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric("Root pages", metrics["Root pages"])

    with col2:
        st.metric("ReadMe links", metrics["ReadMe links"])

    with col3:
        st.metric("Knowledge tiles", metrics["Knowledge tiles"])

    with col4:
        st.metric("Saved files", metrics["Saved files"])

    with col5:
        st.metric("Failed pages", metrics["Failed pages"])


def safe_key(value: str) -> str:
    return (
        value.replace("\\", "_")
        .replace("/", "_")
        .replace(":", "_")
        .replace(" ", "_")
        .replace(".", "_")
    )


def show_download_buttons(export_folder: Path, key_prefix: str) -> None:
    zip_path = zip_folder(export_folder)
    excel_path = export_folder / EXCEL_FILE_NAME
    key = safe_key(key_prefix)

    download_col1, download_col2 = st.columns(2)

    with download_col1:
        with open(zip_path, "rb") as file:
            st.download_button(
                label="⬇️ Download TXT ZIP",
                data=file,
                file_name=zip_path.name,
                mime="application/zip",
                key=f"{key}_txt_zip_download",
            )

    with download_col2:
        if excel_path.exists():
            with open(excel_path, "rb") as file:
                st.download_button(
                    label="⬇️ Download Excel Overview",
                    data=file,
                    file_name=excel_path.name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    key=f"{key}_excel_download",
                )
        else:
            st.warning("Excel overview was not generated.")


st.title("📚 CastorDoc Documentation Exporter")

st.markdown(
    """
Export CastorDoc / Coalesce documentation into clean files.

The tool creates:

- `.txt` files for each exported ReadMe;
- one Excel model specification workbook;
- one ZIP file containing the full export.

It now tries metadata/API extraction first, then falls back to UI scrolling/clicking.
"""
)

st.markdown("## 1. Log in to CastorDoc")

st.info(
    "When you start an export, a Chromium browser will open. "
    "If CastorDoc asks you to log in, log in manually. "
    "After login, the export continues automatically."
)

st.markdown("## 2. Paste the page(s) to export")

root_urls_text = st.text_area(
    "CastorDoc page URL(s)",
    value="https://app.castordoc.com/terms/internal/fact-production-work-e7cef424/map",
    height=150,
    placeholder=(
        "Paste one URL per line, for example:\n"
        "https://app.castordoc.com/terms/internal/fact-production-work-e7cef424/map"
    ),
)

with st.expander("What exactly will be exported?", expanded=False):
    st.markdown(
        """
For each URL you paste, the tool will:

1. Open the page's **ReadMe**.
2. Save only the ReadMe content.
3. Follow links found inside the ReadMe.
4. Open **Subpages & Map** if the page depth allows it.
5. Try to extract Knowledge links from metadata/API/page state.
6. Fall back to scrolling/clicking visible Knowledge tiles.
7. Save the ReadMe of each Knowledge tile.
8. Stop deeper crawling after depth 3 by default.
9. Generate an Excel model specification workbook.
10. Skip duplicate pages automatically.
"""
    )

with st.expander("Advanced settings", expanded=False):
    output_root = st.text_input(
        "Export folder",
        value=DEFAULT_OUTPUT_ROOT,
    )

    profile_dir = st.text_input(
        "Browser login profile",
        value=DEFAULT_PROFILE_DIR,
    )

    max_pages = st.number_input(
        "Safety limit: maximum pages to visit",
        min_value=1,
        max_value=1000,
        value=DEFAULT_MAX_PAGES,
        step=25,
    )

    max_depth = st.number_input(
        "Maximum depth",
        min_value=0,
        max_value=10,
        value=DEFAULT_MAX_DEPTH,
        step=1,
    )

    subpages_until_depth = st.number_input(
        "Explore Subpages & Map until depth",
        min_value=0,
        max_value=10,
        value=DEFAULT_SUBPAGES_UNTIL_DEPTH,
        step=1,
    )

    python_runner = st.selectbox(
        "Python runner",
        options=["uv run python", "current python", "python"],
        index=0,
    )

    show_logs = st.checkbox(
        "Show technical logs while exporting",
        value=True,
    )

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
        max_depth=int(max_depth),
        subpages_until_depth=int(subpages_until_depth),
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
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            universal_newlines=True,
            env={
                **os.environ.copy(),
                "PYTHONIOENCODING": "utf-8",
            },
        )

        while True:
            if process.stdout is None:
                break

            line = process.stdout.readline()

            if line:
                clean_line = line.rstrip()
                logs.append(clean_line)

                visited_count = sum(1 for item in logs if item.startswith("Visite "))
                progress = min(visited_count / int(max_pages), 1.0)
                progress_bar.progress(progress)

                if show_logs:
                    # Avoid sending a huge websocket payload to Streamlit on every log line.
                    # Keep only the latest logs and refresh the UI periodically.
                    should_refresh_logs = (
                        len(logs) % 10 == 0
                        or "Export terminé" in clean_line
                        or "Erreur" in clean_line
                        or "Knowledge tiles finales" in clean_line
                    )

                    if should_refresh_logs:
                        visible_logs = "\n".join(logs[-120:])
                        visible_logs = visible_logs[-20_000:]

                        log_box.text_area(
                            "Export logs",
                            value=visible_logs,
                            height=420,
                        )
                else:
                    if len(logs) % 10 == 0:
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
                preview_text = read_text_file(selected_file)

                if len(preview_text) > 30_000:
                    preview_text = preview_text[:30_000] + "\n\n--- Preview truncated in UI. Download ZIP for full file. ---"

                st.text_area(
                    selected_file.name,
                    value=preview_text,
                    height=420,
                )

            show_download_buttons(
                export_folder=latest_folder,
                key_prefix=f"latest_{latest_folder.name}",
            )

        else:
            st.warning("No `.txt` files were exported. Check the summary and logs.")

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

            show_download_buttons(
                export_folder=selected_export,
                key_prefix=f"previous_{selected_export.name}",
            )