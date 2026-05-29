import os
import re
import time
import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook

# --- STYLING CONFIGURATIONS ---
st.set_page_config(page_title="Tofler Intelligence Scraper & Parser Dashboard", layout="wide")

# --- DATA EXTRACTION UTILITIES ---
def parse_table_data(table_text):
    """Extract company names from pasted table data (Excel, CSV, or tab-separated)"""
    print(f"[PARSE_TABLE] Received {len(table_text)} characters of table data")

    if not table_text.strip():
        print("[PARSE_TABLE] ERROR: Empty table data")
        return []

    lines = table_text.strip().split('\n')
    print(f"[PARSE_TABLE] Total lines: {len(lines)}")

    companies = []
    header_keywords = ['business name', 'company name', 'name', 'incorp', 'year', 'industry', 'status', 'filter']

    for i, line in enumerate(lines):
        print(f"[PARSE_TABLE] Processing line {i}: {line[:50]}")

        # Try tab-separated first, then comma-separated
        if '\t' in line:
            columns = [col.strip() for col in line.split('\t')]
        else:
            columns = [col.strip() for col in line.split(',')]

        if not columns or not columns[0]:
            print(f"[PARSE_TABLE] Line {i} is empty, skipping")
            continue

        company_name = columns[0]

        # Skip header rows
        is_header = any(keyword.lower() in company_name.lower() for keyword in header_keywords)
        if is_header:
            print(f"[PARSE_TABLE] Line {i} identified as header, skipping")
            continue

        # Skip empty or very short entries
        if len(company_name.strip()) < 2:
            print(f"[PARSE_TABLE] Line {i} too short, skipping")
            continue

        # Clean company name
        cleaned_name = company_name.strip()
        companies.append(cleaned_name)
        print(f"[PARSE_TABLE] Extracted: {cleaned_name}")

    # Remove duplicates while preserving order
    unique_companies = list(dict.fromkeys(companies))
    print(f"[PARSE_TABLE] Total extracted: {len(companies)}, Unique: {len(unique_companies)}")
    print(f"[PARSE_TABLE] Companies: {unique_companies}")

    return unique_companies

def clean_sheet_name(name: str) -> str:
    """Excel sheet names are strictly capped at 31 characters and cannot contain: \ / ? * : [ ]"""
    print(f"[CLEAN_SHEET] Input name: '{name}'")
    clean = re.sub(r"[\\/*?:\[\]]", "", name)
    result = clean[:30].strip()
    print(f"[CLEAN_SHEET] Output name: '{result}'")
    return result

def parse_markdown_table_from_html(html_path):
    """
    Scans the entire HTML file to safely locate the metadata header
    and the structured markdown metrics table grid.
    """
    print(f"[PARSE_HTML] Starting parsing for: {html_path}")
    if not os.path.exists(html_path):
        print(f"[PARSE_HTML] ERROR: File not found at {html_path}")
        return "", []

    try:
        with open(html_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        print(f"[PARSE_HTML] File opened successfully. Total lines: {len(lines)}")
    except Exception as e:
        print(f"[PARSE_HTML] ERROR reading file: {e}")
        return "", []

    metadata_title = ""
    table_rows = []

    for i, line in enumerate(lines):
        line_str = line.strip()

        # Capture the reporting metadata header summary strings
        if "Key metrics of" in line_str or "Based on March" in line_str:
            metadata_title = line_str.lstrip("#").strip()
            print(f"[PARSE_HTML] Found metadata at line {i}: {metadata_title[:50]}")
            continue

        # Target table structures starting with pipeline boundaries
        if line_str.startswith("|"):
            if "---" in line_str:
                continue
            cells = [c.strip() for c in line_str.split("|")[1:-1]]
            if cells:
                table_rows.append(cells)

    print(f"[PARSE_HTML] Parsing complete. Metadata: '{metadata_title[:30]}', Rows found: {len(table_rows)}")
    return metadata_title, table_rows

# --- SELENIUM AUTOMATION WORKER ---
def run_selenium_scraper(email, password, companies):
    """Logs into Tofler, iterates through company profiles, and dumps raw HTML payloads."""
    print(f"\n[SCRAPER] Starting scraper for {len(companies)} companies")
    print(f"[SCRAPER] Email: {email}")
    output_dir = os.path.join(os.getcwd(), "Raw extracted data")
    os.makedirs(output_dir, exist_ok=True)
    print(f"[SCRAPER] Output directory: {output_dir}")

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")

    status_box = st.empty()
    progress_bar = st.progress(0)

    driver = None
    try:
        print("[SCRAPER] Installing ChromeDriver...")
        status_box.info("🚀 Initiating Chrome Driver automation connection pipeline...")
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        print("[SCRAPER] ChromeDriver initialized")
        wait = WebDriverWait(driver, 15)

        print("[SCRAPER] Navigating to https://www.tofler.in/")
        status_box.info("🔐 Accessing Tofler authentication gateway portal...")
        driver.get("https://www.tofler.in/")

        try:
            login_link = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a[href='/login'].button.bg-main.text-white"))
            )
            print("[SCRAPER] Login link found, clicking...")
            login_link.click()
        except Exception as e:
            print(f"[SCRAPER] ERROR finding login link: {e}")
            raise

        try:
            email_input = wait.until(EC.presence_of_element_located((By.ID, "email")))
            password_input = wait.until(EC.presence_of_element_located((By.ID, "password")))
            print("[SCRAPER] Email and password fields found")
        except Exception as e:
            print(f"[SCRAPER] ERROR finding email/password fields: {e}")
            raise

        email_input.clear()
        email_input.send_keys(email)
        password_input.clear()
        password_input.send_keys(password)
        print("[SCRAPER] Credentials entered")

        try:
            submit_button = wait.until(EC.presence_of_element_located((By.ID, "submitBtn")))
            driver.execute_script("arguments[0].disabled = false; arguments[0].style.opacity = 1;", submit_button)
            wait.until(EC.element_to_be_clickable((By.ID, "submitBtn")))
            submit_button.click()
            print("[SCRAPER] Login button clicked")
        except Exception as e:
            print(f"[SCRAPER] ERROR clicking submit button: {e}")
            raise

        try:
            wait.until(EC.presence_of_element_located((By.ID, "searchbox")))
            print("[SCRAPER] Login successful, searchbox found")
            status_box.success("✅ Login session authorized successfully!")
        except Exception as e:
            print(f"[SCRAPER] ERROR: Login failed, searchbox not found: {e}")
            raise

        time.sleep(1)

        # Process company array stack
        for index, company in enumerate(companies):
            print(f"\n[SCRAPER] Processing company {index + 1}/{len(companies)}: {company}")
            status_box.info(f"🔍 Searching and navigating company file track: **{company}**")
            try:
                driver.get("https://www.tofler.in/")
                search_box = wait.until(EC.presence_of_element_located((By.ID, "searchbox")))
                search_box.clear()
                search_box.send_keys(company)
                print(f"[SCRAPER] Company name entered in search: {company}")

                time.sleep(3)

                dropdown_items_selector = ".ui-menu-item, .ui-autocomplete li a"
                suggestions = driver.find_elements(By.CSS_SELECTOR, dropdown_items_selector)
                print(f"[SCRAPER] Found {len(suggestions)} suggestions for '{company}'")

                if not suggestions:
                    print(f"[SCRAPER] WARNING: No suggestions found for '{company}', skipping")
                    st.warning(f"⚠️ No autocomplete indices matched for '{company}'. Skipping...")
                    continue

                first_suggestion = suggestions[0]
                driver.execute_script("arguments[0].click();", first_suggestion)
                print(f"[SCRAPER] Clicked first suggestion")

                wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(2)

                current_url = driver.current_url
                print(f"[SCRAPER] Current URL: {current_url}")

                if "tofler.in/" in current_url and current_url.strip("/") == "https://www.tofler.in":
                    print(f"[SCRAPER] WARNING: Still on home page for '{company}', skipping")
                    st.warning(f"⚠️ Anchor translation block encountered for '{company}'. Skipping...")
                    continue

                # Isolate target file naming rules
                sanitized_name = re.sub(r"[^A-Za-z0-9 _\-]", "_", company).strip().replace(" ", "_")
                raw_file_path = os.path.join(output_dir, f"{sanitized_name}.html")
                print(f"[SCRAPER] Saving to: {raw_file_path}")

                with open(raw_file_path, "w", encoding="utf-8") as raw_file:
                    raw_file.write(driver.page_source)
                print(f"[SCRAPER] HTML saved successfully ({len(driver.page_source)} bytes)")

            except Exception as loop_err:
                print(f"[SCRAPER] ERROR processing company '{company}': {type(loop_err).__name__} - {str(loop_err)}")
                st.error(f"❌ Automation runtime disruption matching company profile '{company}': {str(loop_err)}")

            progress_bar.progress((index + 1) / len(companies))

    except Exception as connection_error:
        print(f"[SCRAPER] CRITICAL ERROR: {type(connection_error).__name__} - {str(connection_error)}")
        st.error(f"💥 Critical driver architecture setup initialization failure: {str(connection_error)}")
    finally:
        print("[SCRAPER] Cleaning up - closing browser")
        status_box.success("-- Web automation execution thread terminated cleanly.")
        if driver:
            driver.quit()
        print("[SCRAPER] Browser closed successfully")

# --- WORKBOOK COMPILE ENGINE ---
def compile_data_to_excel(companies):
    """Processes downloaded HTML files and appends sheets. Always maintains a visible landing tab."""
    print(f"\n[COMPILE] Starting compilation for {len(companies)} companies")
    input_directory = os.path.join(os.getcwd(), "Raw extracted data")
    output_excel_file = os.path.join(os.getcwd(), "Companies_Clean_Financial_Report.xlsx")

    print(f"[COMPILE] Input directory: {input_directory}")
    print(f"[COMPILE] Output file: {output_excel_file}")

    if not os.path.exists(input_directory):
        print(f"[COMPILE] ERROR: Input directory does not exist: {input_directory}")
        st.error("Missing raw input data folder tracking locations.")
        return False

    file_exists = os.path.exists(output_excel_file)
    mode = 'a' if file_exists else 'w'
    if_sheet_exists = 'replace' if file_exists else None
    print(f"[COMPILE] File exists: {file_exists}, Mode: {mode}")

    # Open writer connection
    try:
        with pd.ExcelWriter(output_excel_file, engine="openpyxl", mode=mode, if_sheet_exists=if_sheet_exists) as excel_writer:

            # Fix the IndexError by initializing a visible dashboard index summary sheet first
            if not file_exists:
                print("[COMPILE] Creating Master Summary Index sheet")
                summary_df = pd.DataFrame([["Report Generated", time.strftime("%Y-%m-%d %H:%M:%S")]], columns=["System Log", "Timestamp"])
                summary_df.to_excel(excel_writer, sheet_name="Master Summary Index", index=False)

            for i, company in enumerate(companies):
                print(f"\n[COMPILE] Processing company {i + 1}/{len(companies)}: {company}")
                sanitized_name = re.sub(r"[^A-Za-z0-9 _\-]", "_", company).strip().replace(" ", "_")
                html_file_path = os.path.join(input_directory, f"{sanitized_name}.html")

                print(f"[COMPILE] Looking for HTML: {html_file_path}")

                if not os.path.exists(html_file_path):
                    print(f"[COMPILE] WARNING: HTML file not found for: {company}")
                    st.warning(f"⚠️ HTML file not found for: {company}")
                    continue

                try:
                    print(f"[COMPILE] Parsing HTML file...")
                    metadata, rows = parse_markdown_table_from_html(html_file_path)
                    print(f"[COMPILE] Parsed - Metadata length: {len(metadata)}, Rows: {len(rows)}")

                    if not rows:
                        print(f"[COMPILE] WARNING: No rows found in HTML for: {company}")
                        st.warning(f"⚠️ No markdown financial table found inside the HTML text for: {company}")
                        continue

                    headers = ["Metric", "Value", "Change"]
                    # Clean up header duplicate anomalies if present
                    if rows and rows[0][0].lower() == "metric":
                        df_data = rows[1:]
                        print(f"[COMPILE] Removed duplicate header row")
                    else:
                        df_data = rows

                    print(f"[COMPILE] Creating DataFrame with {len(df_data)} rows")
                    df = pd.DataFrame(df_data, columns=headers)
                    df["Value"] = df["Value"].str.replace("&gt;", ">").str.replace("&lt;", "<")

                    valid_sheet_title = clean_sheet_name(company)
                    print(f"[COMPILE] Writing sheet: '{valid_sheet_title}'")
                    df.to_excel(excel_writer, sheet_name=valid_sheet_title, index=False)

                    # Push the metadata summary context row back into cell A1
                    worksheet = excel_writer.sheets[valid_sheet_title]
                    if metadata:
                        worksheet.insert_rows(1, amount=2)
                        worksheet["A1"] = metadata
                        print(f"[COMPILE] Added metadata to sheet")

                    print(f"[COMPILE] Successfully processed: {company}")

                except Exception as parse_err:
                    print(f"[COMPILE] ERROR processing company '{company}': {type(parse_err).__name__} - {str(parse_err)}")
                    st.error(f"Failed cleaning markdown payload layout rules for '{company}': {str(parse_err)}")

        print(f"[COMPILE] All companies processed successfully")
        delete_raw_html_files(input_directory)
        return True

    except Exception as excel_err:
        print(f"[COMPILE] CRITICAL ERROR writing Excel: {type(excel_err).__name__} - {str(excel_err)}")
        return False


def delete_raw_html_files(input_directory):
    """Remove all HTML files from the raw data folder after pipeline completion."""
    print(f"[CLEANUP] Removing raw HTML files from: {input_directory}")
    if not os.path.isdir(input_directory):
        print(f"[CLEANUP] Directory not found: {input_directory}")
        return 0

    removed_count = 0
    for filename in os.listdir(input_directory):
        file_path = os.path.join(input_directory, filename)
        if os.path.isfile(file_path) and filename.lower().endswith(".html"):
            try:
                os.remove(file_path)
                removed_count += 1
                print(f"[CLEANUP] Deleted: {file_path}")
            except Exception as e:
                print(f"[CLEANUP] ERROR deleting {file_path}: {e}")

    if removed_count == 0:
        print("[CLEANUP] No HTML files found to delete")
    if os.path.isdir(input_directory) and not os.listdir(input_directory):
        try:
            os.rmdir(input_directory)
            print(f"[CLEANUP] Removed empty directory: {input_directory}")
        except Exception as e:
            print(f"[CLEANUP] ERROR removing directory: {e}")

    if removed_count:
        st.info(f"🧹 Cleaned up {removed_count} raw HTML file(s) from the pipeline folder.")
    return removed_count


def delete_excel_report(report_path):
    """Delete the master Excel report file from disk."""
    print(f"[DELETE] Attempting to remove Excel report: {report_path}")
    if not os.path.exists(report_path):
        print(f"[DELETE] Excel report not found: {report_path}")
        return False

    try:
        os.remove(report_path)
        print(f"[DELETE] Deleted Excel report: {report_path}")
        return True
    except Exception as e:
        print(f"[DELETE] ERROR deleting Excel report: {e}")
        st.error(f"Unable to delete the report file: {str(e)}")
        return False

# --- INTERACTIVE USER INTERFACE DESIGN ---
# Detect theme and use appropriate logo



st.title("📊 Tofler Financial Data Intelligence Harvester")

st.markdown("Automate data gathering loops, clean messy markdown structures, and monitor corporate profiles smoothly.")

# Persistent file check coordinates
excel_path = os.path.join(os.getcwd(), "Companies_Clean_Financial_Report.xlsx")

# Structural Split UI Columns layout configuration panels
setup_col, viewer_col = st.columns([1, 1.2])

with setup_col:
    st.subheader("⚙️ Scraping Parameters Panel")

    # Toggle between manual entry and table paste
    input_method = st.radio("📋 Choose Input Method:", ["Manual Entry (Comma-separated)", "Paste Table Data"], horizontal=True)

    with st.form("credentials_and_targets_form"):
        user_email = st.text_input(
            "Tofler Username Email ID",
            value="",
            placeholder="Enter your email"
        )
        user_pass = st.text_input(
            "Profile Password Security Code",
            value="",
            placeholder="Enter your password",
            type="password"
        )

        if input_method == "Manual Entry (Comma-separated)":
            company_input_raw = st.text_area(
                "Target Company Portfolio Profile Queue (Provide entries separated by commas)",
                placeholder="Enter company names : Tata,Zoho,Nestle",
                height=100
            )
            table_input = ""
        else:
            st.info("📌 Paste table data from Excel/CSV. Script will extract company names from the first column automatically.")
            table_input = st.text_area(
                "Paste your table data here (Excel, CSV, or Tab-separated):",
                placeholder="Paste the Raw Data(Companies name)",
                height=150
            )
            company_input_raw = ""

        submit_execution = st.form_submit_button("🔥 Fire Extraction Pipeline Sequence")

    if submit_execution:
        print("[MAIN] Form submitted")

        # Process input based on method
        if input_method == "Paste Table Data":
            if table_input.strip():
                target_companies_list = parse_table_data(table_input)
                print(f"[MAIN] Parsed table data, found {len(target_companies_list)} companies")
            else:
                print("[MAIN] ERROR: Table data is empty")
                st.error("Please paste table data")
                target_companies_list = []
        else:
            target_companies_list = [c.strip() for c in company_input_raw.split(",") if c.strip()]
            print(f"[MAIN] Manual entry: {target_companies_list}")

        print(f"[MAIN] Email: {user_email}")
        print(f"[MAIN] Target companies: {target_companies_list}")

        if not user_email or not user_pass:
            print("[MAIN] ERROR: Email or password is empty")
            st.error("Authentication validation constraints error: Fields cannot remain blank.")
        elif not target_companies_list:
            print("[MAIN] ERROR: No companies provided")
            st.error("Input targets parsing constraints error: Target configuration stack empty.")
        else:
            print(f"[MAIN] Starting scraper with {len(target_companies_list)} companies...")
            with st.spinner("Executing Selenium sequence tasks..."):
                run_selenium_scraper(user_email, user_pass, target_companies_list)

            print("[MAIN] Scraper completed, starting compilation...")
            with st.spinner("Extracting hidden table boundaries and building sheets..."):
                success = compile_data_to_excel(target_companies_list)
                if success:
                    print("[MAIN] Compilation successful!")
                    st.success("🎯 Compilation finalized! All targets synchronized inside the master report.")
                else:
                    print("[MAIN] Compilation failed!")


with viewer_col:
    st.subheader("📈 Embedded Spreadsheet Viewer Terminal")

    if os.path.exists(excel_path):
        try:
            print(f"\n[VIEWER] Opening Excel file: {excel_path}")
            # Look up and cache structural sheets present inside the report book layout
            with pd.ExcelFile(excel_path, engine="openpyxl") as xl_file:
                available_sheet_tabs = xl_file.sheet_names
                print(f"[VIEWER] Found sheets: {available_sheet_tabs}")

                selected_tab = st.selectbox("🎯 Select Active Corporate Entity Tracking Target Worksheet:", available_sheet_tabs)
                print(f"[VIEWER] Selected sheet: {selected_tab}")

            # Read structural metrics framework directly while accounting for metadata row spacing offsets
            raw_preview_df = pd.read_excel(excel_path, sheet_name=selected_tab)
            print(f"[VIEWER] DataFrame shape: {raw_preview_df.shape}, Columns: {list(raw_preview_df.columns)}")

            # Visual interface rendering logic separating metadata context titles from core metrics arrays
            if not raw_preview_df.empty and "Metric" not in raw_preview_df.columns:
                print("[VIEWER] Metadata detected in first row")
                # Capture metadata string header row safely
                metadata_header_string = str(raw_preview_df.columns[0])
                st.caption(f"ℹ️ **Reporting Metadata Context Header Description:** `{metadata_header_string}`")

                # Re-parse cleanly to present clean data tables to users without spacing block offsets
                clean_preview_df = pd.read_excel(excel_path, sheet_name=selected_tab, skiprows=2)
                print(f"[VIEWER] Clean DataFrame shape: {clean_preview_df.shape}")
                st.dataframe(clean_preview_df, use_container_width=True, hide_index=True)
            else:
                print("[VIEWER] No metadata detected, showing raw data")
                st.dataframe(raw_preview_df, use_container_width=True, hide_index=True)

            # Embed a native workbook downloader tool inside the UI layout column
            with open(excel_path, "rb") as file_bytes:
                st.download_button(
                    label="📥 Download Master Financial Report Workbook (.xlsx)",
                    data=file_bytes,
                    file_name="Companies_Clean_Financial_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if st.button("🗑️ Erase Master Excel Report"):
                if delete_excel_report(excel_path):
                    st.success("✅ Master workbook deleted successfully. Refresh the app if needed to update the viewer.")
                else:
                    st.warning("No workbook found to erase, or it could not be deleted.")

            print("[VIEWER] Download and delete buttons rendered")
        except Exception as viewer_err:
            print(f"[VIEWER] ERROR: {type(viewer_err).__name__} - {str(viewer_err)}")
            st.error(f"Could not open spreadsheet viewer tool layout grids: {str(viewer_err)}")
    else:
        print(f"[VIEWER] Excel file not found: {excel_path}")
        st.info("💡 No master financial workbook detected inside workspace yet. Fill out the form fields and fire the extraction pipeline to create one.")
