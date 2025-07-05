# Vaayuputra XCP Tool Automation
"""
Automates collation and processing of Excel/CSV files with ASINs using a modern GUI (customtkinter) and Playwright.
Distributable as a portable Windows .exe with bundled browser binaries.
"""

__version__ = "1.0.0"

import os
import sys
# Set PLAYWRIGHT_BROWSERS_PATH to the local ms-playwright folder next to the .exe
if getattr(sys, 'frozen', False):  # Running as compiled .exe
    exe_dir = os.path.dirname(sys.executable)
    browsers_path = os.path.join(exe_dir, 'ms-playwright')
    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = browsers_path  # Ensures portable Playwright

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import nest_asyncio
import logging
import datetime
import glob
import pyautogui
import time
import re

# Apply nest_asyncio to allow nested event loops
nest_asyncio.apply()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='xcp_tool.log'
)

class XCPToolGUI(ctk.CTk):
    # Mapping from marketplace_id to dropdown label
    MARKETPLACE_MAP = {
        'US': 'amazon.com',
        'CA': 'amazon.ca',
        'IN': 'amazon.in',
        'UK': 'amazon.co.uk',
        'DE': 'amazon.de',
        'FR': 'amazon.fr',
        'IT': 'amazon.it',
        'ES': 'amazon.es',
        'JP': 'amazon.co.jp',
        'AU': 'amazon.com.au',
        'SG': 'amazon.sg',
        'AE': 'amazon.ae',
        'SA': 'amazon.sa',
        'MX': 'amazon.com.mx',
        'BR': 'amazon.com.br',
        'NL': 'amazon.nl',
        'SE': 'amazon.se',
        'PL': 'amazon.pl',
        'TR': 'amazon.com.tr',
    }

    def __init__(self):
        super().__init__()
        
        # Initialize event loop
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        # Configure window
        self.title("XCP Tool Automation")
        self.geometry("1000x600")

        # Create main container
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Title
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Vaayuputra",
            font=ctk.CTkFont(family="Segoe UI", size=40, weight="bold", slant="italic"),
            text_color="#1e90ff"  # Dodger blue for a wind effect
        )
        self.title_label.grid(row=0, column=0, pady=20, padx=20)

        # File Selection Frame
        self.file_frame = ctk.CTkFrame(self.main_frame)
        self.file_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.file_frame.grid_columnconfigure(1, weight=1)

        self.file_label = ctk.CTkLabel(
            self.file_frame, 
            text="Input Excel File:",
            font=ctk.CTkFont(size=14)
        )
        self.file_label.grid(row=0, column=0, padx=10, pady=10)

        self.file_path = ctk.CTkEntry(self.file_frame)
        self.file_path.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.browse_button = ctk.CTkButton(
            self.file_frame,
            text="Browse",
            command=self.browse_file
        )
        self.browse_button.grid(row=0, column=2, padx=10, pady=10)

        # Progress Frame
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        self.progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            self.progress_frame,
            text="Status: Ready",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.grid(row=1, column=0, pady=5)

        # Log Frame
        self.log_frame = ctk.CTkFrame(self.main_frame)
        self.log_frame.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            self.log_frame,
            height=200,
            font=ctk.CTkFont(size=12)
        )
        self.log_text.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # Buttons Frame
        self.button_frame = ctk.CTkFrame(self.main_frame)
        self.button_frame.grid(row=5, column=0, padx=20, pady=10)

        self.start_button = ctk.CTkButton(
            self.button_frame,
            text="Start Processing",
            command=self.start_processing,
            width=150
        )
        self.start_button.grid(row=0, column=0, padx=10)

        self.stop_button = ctk.CTkButton(
            self.button_frame,
            text="Stop",
            command=self.stop_processing,
            width=150,
            state="disabled"
        )
        self.stop_button.grid(row=0, column=1, padx=10)

        # Suffix Management Frame
        self.suffix_frame = ctk.CTkFrame(self.main_frame)
        self.suffix_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.suffix_frame.grid_columnconfigure(1, weight=1)

        self.suffix_label = ctk.CTkLabel(
            self.suffix_frame,
            text="Add Suffix to Remove:",
            font=ctk.CTkFont(size=14)
        )
        self.suffix_label.grid(row=0, column=0, padx=10, pady=10)

        self.suffix_entry = ctk.CTkEntry(self.suffix_frame)
        self.suffix_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        self.add_suffix_button = ctk.CTkButton(
            self.suffix_frame,
            text="Add Suffix",
            command=self.add_suffix
        )
        self.add_suffix_button.grid(row=0, column=2, padx=10, pady=10)

        self.show_suffixes_button = ctk.CTkButton(
            self.suffix_frame,
            text="Show Current Suffixes",
            command=self.toggle_suffix_listbox
        )
        self.show_suffixes_button.grid(row=0, column=3, padx=10, pady=10)

        self.suffix_list_label = ctk.CTkLabel(
            self.suffix_frame,
            text="Current Suffixes:",
            font=ctk.CTkFont(size=12)
        )
        self.suffix_list_label.grid(row=1, column=0, padx=10, pady=5, sticky="nw")
        self.suffix_list_label.grid_remove()  # Hide initially

        self.suffix_listbox = Listbox(
            self.suffix_frame,
            height=5,
            selectmode=tk.MULTIPLE,
            exportselection=False,
            font=("Arial", 12)
        )
        self.suffix_listbox.grid(row=1, column=1, columnspan=3, padx=10, pady=5, sticky="ew")
        self.suffix_listbox.grid_remove()  # Hide initially

        self.remove_suffix_button = ctk.CTkButton(
            self.suffix_frame,
            text="Remove Selected Suffix",
            command=self.remove_selected_suffix
        )
        self.remove_suffix_button.grid(row=2, column=1, padx=10, pady=5)
        self.remove_suffix_button.grid_remove()  # Hide initially

        # Initialize suffixes
        self.suffixes = [
            '_UIL', '_IN', '_US', '_CA', '_SG', '_AU', '_IE', '_UK', '_CS2',
            '_Class_Consolidation', '_Paradigm', '_Mirage', '_100keyword', '_100_keyword'
        ]
        # Do not show suffixes at startup

        # Initialize processing flag
        self.is_processing = False

        # Keep the event loop running with Tkinter
        self.after(100, self._run_asyncio_loop)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_path.delete(0, tk.END)
            self.file_path.insert(0, filename)
            self.update_log(f"Selected file: {filename}")

    def update_log(self, message):
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        logging.info(message)

    def update_status(self, message):
        self.status_label.configure(text=f"Status: {message}")
        logging.info(f"Status update: {message}")

    def update_progress(self, value):
        self.progress_bar.set(value)

    async def process_asins(self):
        import datetime
        import time as pytime
        playwright = None
        browser = None
        context = None
        page = None
        script_dir = os.path.dirname(os.path.abspath(__file__))
        export_dir = os.path.join(script_dir, f"exports_{datetime.date.today()}")
        os.makedirs(export_dir, exist_ok=True)
        try:
            self.update_log("Maximizing window using PyAutoGUI...")
            pyautogui.hotkey('win', 'up')
            time.sleep(1)

            input_file = self.file_path.get()
            if not input_file:
                messagebox.showerror("Error", "Please select an input file")
                return

            self.update_status("Initializing...")
            self.update_progress(0.1)

            df = pd.read_excel(input_file)
            if 'Class' in df.columns:
                group_col = 'Class'
            elif 'rule_name' in df.columns:
                group_col = 'rule_name'
            else:
                messagebox.showerror("Error", "Input file must contain a 'Class' or 'rule_name' column.")
                self.update_log("Error: No 'Class' or 'rule_name' column found in input file.")
                return
            self.update_log(f"Successfully loaded {len(df)} rows from Excel. Grouping by '{group_col}' column.")
            self.update_progress(0.2)

            playwright = await async_playwright().start()
            browser = await playwright.chromium.launch(headless=False, args=['--start-maximized'])
            context = await browser.new_context(viewport=None)
            page = await context.new_page()
            self.update_log("Browser initialized successfully")
            self.update_progress(0.3)
            await page.goto('https://www.cp-central.catalog.amazon.dev/#/class/search')
            self.update_log("Navigated to CP Central")
            self.update_progress(0.4)
            # SSO Login Handling
            if "SSO/redirect" in page.url or "midway-auth.amazon.com" in page.url:
                self.update_log("SSO login required. Please complete the login in the opened browser window.")
                try:
                    await page.wait_for_selector('#awsui-input-0', timeout=0)
                    self.update_log("Login successful. Continuing automation.")
                except Exception as e:
                    self.update_log(f"Error waiting for login: {str(e)}")
                    return
            await page.wait_for_selector('#awsui-input-0', timeout=10000)
            await page.wait_for_timeout(500)

            class_search_url = 'https://www.cp-central.catalog.amazon.dev/#/class/search'
            class_counter = 0
            total_classes = len(df[group_col].unique())
            start_time = pytime.time()
            for class_name, group in df.groupby(group_col):
                if not self.is_processing:
                    self.update_log("Processing stopped by user")
                    break
                # Periodically close and reopen the page every 15 classes
                if class_counter > 0 and class_counter % 15 == 0:
                    try:
                        await page.close()
                        self.update_log(f"Page closed to free resources after {class_counter} classes.")
                    except Exception as e:
                        self.update_log(f"Error closing page: {str(e)}")
                    page = await context.new_page()
                    await page.goto(class_search_url)
                    await page.wait_for_timeout(1000)
                    self.update_log("New page opened in same context (SSO session preserved).")
                class_counter += 1
                class_start = pytime.time()
                clean_name = self.clean_class_name(class_name)
                self.update_log(f"Processing class {class_counter}/{total_classes}: {clean_name}")
                try:
                    await self.process_class(page, class_search_url, clean_name, group, export_dir)
                    elapsed = pytime.time() - class_start
                    self.update_log(f"Class '{clean_name}' processed in {elapsed:.2f} seconds.")
                except Exception as e:
                    self.update_log(f"Error processing class {clean_name}: {str(e)}")
                    continue
            total_elapsed = pytime.time() - start_time
            self.update_status("Processing complete")
            self.update_progress(1.0)
            self.update_log(f"All classes processed in {total_elapsed/60:.2f} minutes.")
        except Exception as e:
            self.update_log(f"Error: {str(e)}")
            logging.error(f"Error in process_asins: {str(e)}", exc_info=True)
            messagebox.showerror("Error", str(e))
        finally:
            if browser:
                try:
                    await browser.close()
                    self.update_log("Browser closed")
                except Exception as e:
                    logging.error(f"Error closing browser: {str(e)}")
            if playwright:
                await playwright.stop()
            self.is_processing = False
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            await self.collate_exports(export_dir)

    async def process_class(self, page, class_search_url, class_name, group, export_dir):
        # Retry logic: try up to 3 times if class input is not found
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            input_box = page.locator('input[placeholder*="class name"]')
            found = await self.wait_for_visible_enabled(input_box, page)
            if found:
                break
            self.update_log(f"Attempt {attempt}: Could not find class input for '{class_name}'. Reloading class search page...")
            await page.goto(class_search_url, wait_until="domcontentloaded")
            await page.wait_for_timeout(1500)
        else:
            self.update_log(f"Failed to find class input for '{class_name}' after {max_retries} attempts. Skipping class.")
            return
        await input_box.scroll_into_view_if_needed()
        await input_box.focus()
        await page.wait_for_timeout(100)
        await input_box.fill('')
        await input_box.type(class_name, delay=30)
        await page.keyboard.press('Enter')
        await page.wait_for_timeout(500)
        self.update_log(f"Class name '{class_name}' entered successfully.")
        # Click class link
        class_link = page.locator(f"a:text-is('{class_name}')")
        await class_link.first.wait_for(timeout=5000)
        await class_link.first.click()
        await page.wait_for_timeout(1000)
        # Click 'New sample ASINs test'
        await self.click_sample_test_btn(page)
        await page.wait_for_timeout(1000)
        # Uncheck box
        await self.uncheck_sample_asins_box(page)
        await page.wait_for_timeout(500)
        # Input ASINs in batches of 900
        asins = group['asin_id'].astype(str).tolist()
        batch_size = 900
        total_batches = (len(asins) + batch_size - 1) // batch_size
        # Get marketplace_id once for this group
        marketplace_id = group['marketplace_id'].iloc[0] if 'marketplace_id' in group.columns else None
        for batch_num, i in enumerate(range(0, len(asins), batch_size), 1):
            batch_asins = asins[i:i+batch_size]
            self.update_log(f"Processing batch {batch_num}/{total_batches} for class {class_name} with {len(batch_asins)} ASINs.")
            await self.input_asins(page, batch_asins)
            await self.click_test_sample_asins(page)
            # Marketplace selection will now happen inside export_results
            await self.export_results(page, f"{class_name}_batch{batch_num}", export_dir, class_search_url, marketplace_id)
            if batch_num < total_batches:
                await page.goto(class_search_url, wait_until="domcontentloaded")
                await page.wait_for_timeout(1000)
                # Re-enter class for next batch (with retry)
                for attempt in range(1, max_retries + 1):
                    input_box = page.locator('input[placeholder*="class name"]')
                    found = await self.wait_for_visible_enabled(input_box, page)
                    if found:
                        break
                    self.update_log(f"Attempt {attempt}: Could not find class input for '{class_name}' on batch {batch_num+1}. Reloading class search page...")
                    await page.goto(class_search_url, wait_until="domcontentloaded")
                    await page.wait_for_timeout(1500)
                else:
                    self.update_log(f"Failed to find class input for '{class_name}' on batch {batch_num+1} after {max_retries} attempts. Skipping remaining batches.")
                    break
                await input_box.scroll_into_view_if_needed()
                await input_box.focus()
                await page.wait_for_timeout(100)
                await input_box.fill('')
                await input_box.type(class_name, delay=30)
                await page.keyboard.press('Enter')
                await page.wait_for_timeout(500)
                class_link = page.locator(f"a:text-is('{class_name}')")
                await class_link.first.wait_for(timeout=5000)
                await class_link.first.click()
                await page.wait_for_timeout(1000)
                await self.click_sample_test_btn(page)
                await page.wait_for_timeout(1000)
                await self.uncheck_sample_asins_box(page)
                await page.wait_for_timeout(500)

    async def wait_for_visible_enabled(self, locator, page, retries=15, delay=200):
        for _ in range(retries):
            try:
                if await locator.is_visible() and await locator.is_enabled():
                    return True
            except Exception:
                pass
            await page.wait_for_timeout(delay)
        return False

    async def click_sample_test_btn(self, page):
        try:
            btn = page.locator('#app-content > div > div > div:nth-child(1) > div > div.awsui-util-action-stripe-large > div.awsui-util-action-stripe-group.awsui-util-pv-n > awsui-button:nth-child(1) > a')
            await btn.wait_for(timeout=5000)
            await btn.click()
        except Exception:
            try:
                btn = page.locator('xpath=//*[@id="app-content"]/div/div/div[1]/div/div[2]/div[2]/awsui-button[1]/a')
                await btn.wait_for(timeout=2000)
                await btn.click()
            except Exception as e:
                self.update_log(f"Could not click 'New sample ASINs test' button: {str(e)}")

    async def uncheck_sample_asins_box(self, page):
        try:
            checkbox = page.locator('input[type="checkbox"]')
            label = await checkbox.evaluate_handle('el => el.parentElement.textContent')
            if 'Include sample ASINs provided during the class authoring process' in await label.json_value():
                if await checkbox.is_checked():
                    await checkbox.click()
                    self.update_log("Unchecked the 'Include sample ASINs' box.")
        except Exception as e:
            self.update_log(f"Could not uncheck the box: {str(e)}")

    async def input_asins(self, page, asins):
        # Retry logic for filling ASINs
        for attempt in range(3):
            try:
                await page.wait_for_selector('textarea[placeholder^="Enter ASIN"]', timeout=10000)
                asin_inputs = page.locator('textarea[placeholder^="Enter ASIN"]')
                count = await asin_inputs.count()
                asin_input_area = None
                asin_input_index = 0
                if count > 1:
                    ids = []
                    for idx in range(count):
                        handle = asin_inputs.nth(idx)
                        id_val = await handle.get_attribute('id')
                        ids.append(id_val)
                    self.update_log(f"Found {count} ASIN textareas with ids: {ids}")
                    for idx in range(count):
                        handle = asin_inputs.nth(idx)
                        visible = await handle.is_visible()
                        enabled = await handle.is_enabled()
                        if visible and enabled:
                            asin_input_area = handle
                            asin_input_index = idx
                            break
                else:
                    asin_input_area = asin_inputs.first
                asin_text = '\n'.join(asins)
                await asin_input_area.fill(asin_text, timeout=20000)
                self.update_log(f"Filled ASINs textarea (index {asin_input_index}) with {len(asins)} ASINs.")
                await page.wait_for_timeout(500)
                return
            except Exception as e:
                self.update_log(f"Attempt {attempt+1}: Could not input ASINs: {str(e)}")
                await page.reload()
                await page.wait_for_timeout(1000)
        self.update_log("Failed to input ASINs after 3 attempts.")

    async def click_test_sample_asins(self, page):
        try:
            test_btn = page.locator('button:has(span:text("Test sample ASINs")), button:has-text("Test sample ASINs")')
            await test_btn.wait_for(timeout=5000)
            await test_btn.click()
            self.update_log("Clicked 'Test sample ASINs' button.")
        except Exception as e:
            self.update_log(f"Could not click 'Test sample ASINs' button: {str(e)}")

    async def export_results(self, page, class_name, export_dir, class_search_url, marketplace_id=None):
        try:
            export_btn = page.locator('#app-content > div > div:nth-child(3) > div.test-sample-asins-component > div:nth-child(4) > awsui-table > div > div.awsui-table-heading-container > div > div.awsui-table-header > span > div > div.awsui-util-action-stripe-group > awsui-button > button')
            # Wait for export button to be visible and enabled (ASIN test results loaded)
            await export_btn.wait_for(state="visible", timeout=120000)
            await export_btn.scroll_into_view_if_needed()
            while not await export_btn.is_enabled():
                await page.wait_for_timeout(500)
            self.update_log("ASINs tested, export button is now enabled.")
            # Now select marketplace (dropdown will be available)
            if marketplace_id:
                await self.select_marketplace_dropdown(page, marketplace_id)
            await export_btn.hover()
            await page.wait_for_timeout(200)
            async with page.expect_download() as download_info:
                await export_btn.click(force=True)
            download = await download_info.value
            class_export_name = os.path.join(export_dir, f"export_{class_name.replace('/', '_').replace(' ', '_')}.xlsx")
            await download.save_as(class_export_name)
            self.update_log(f"Downloaded export for class {class_name} as {class_export_name}")

            # Robust conversion to CSV (like FS Pre-filter Export)
            import shutil
            try:
                try:
                    df = pd.read_excel(class_export_name)
                except Exception:
                    # Try reading as CSV
                    try:
                        df = pd.read_csv(class_export_name)
                        class_export_csv = class_export_name.replace('.xlsx', '.csv')
                        df.to_csv(class_export_csv, index=False)
                        self.update_log(f"Downloaded file was CSV, saved as: {class_export_csv}")
                        try:
                            os.remove(class_export_name)
                        except Exception:
                            pass
                        class_export_name = class_export_csv
                    except Exception as e:
                        self.update_log(f"Downloaded file is not Excel or CSV: {str(e)}")
                        return
                else:
                    class_export_csv = class_export_name.replace('.xlsx', '.csv')
                    df.to_csv(class_export_csv, index=False)
                    os.remove(class_export_name)
                    self.update_log(f"Converted export for class {class_name} to CSV: {class_export_csv}")
            except Exception as e:
                self.update_log(f"Could not convert export for class {class_name} to CSV: {str(e)}")

            await page.goto(class_search_url, wait_until="domcontentloaded")
            self.update_log("Returned to fresh Class Search page for next class.")
        except Exception as e:
            self.update_log(f"Could not export results for class {class_name}: {str(e)}")

    async def collate_exports(self, export_dir):
        try:
            # Collate all CSV exports in the export_dir
            all_files = glob.glob(os.path.join(export_dir, 'export_*.csv'))
            if all_files:
                dfs = []
                for file in all_files:
                    try:
                        df = pd.read_csv(file)
                    except Exception:
                        self.update_log(f"Could not read file {file} as CSV. Skipping.")
                        continue
                    # Sanitize all column names
                    df.columns = [self.sanitize_excel_column(str(col)) for col in df.columns]
                    df['source_file'] = os.path.basename(file)
                    dfs.append(df)
                if dfs:
                    combined = pd.concat(dfs, ignore_index=True)
                    # Sanitize again after concat in case of new columns
                    combined.columns = [self.sanitize_excel_column(str(col)) for col in combined.columns]
                    # Save as CSV with present date in the export_dir
                    combined_file = os.path.join(export_dir, f"collated_exports_{datetime.date.today()}.csv")
                    try:
                        combined.to_csv(combined_file, index=False)
                        self.update_log(f"Collated all exports into {combined_file}")
                    except PermissionError:
                        self.update_log(f"Permission denied: Could not write to {combined_file}. Please close the file if it is open in Excel or another program and try again.")
                        messagebox.showerror("Permission Denied", f"Could not write to {combined_file}. Please close the file if it is open in Excel or another program and try again.")
                    except Exception as e:
                        self.update_log(f"Error saving collated exports: {str(e)}")
        except Exception as e:
            self.update_log(f"Error collating exports: {str(e)}")

    def start_processing(self):
        if not self.is_processing:
            self.is_processing = True
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            self.loop.create_task(self.process_asins())

    def _run_asyncio_loop(self):
        try:
            self.loop.call_soon(self.loop.stop)
            self.loop.run_forever()
        except Exception as e:
            logging.error(f"Asyncio loop error: {str(e)}", exc_info=True)
        self.after(100, self._run_asyncio_loop)

    def stop_processing(self):
        if self.is_processing:
            self.is_processing = False
            self.update_status("Stopping...")
            self.update_log("Stop requested by user")
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")

    def sanitize_excel_column(self, col_name):
        invalid_chars = ['/', '\\', '?', '*', '[', ']', ':', ';', '\n', '\r', '\t', '|']
        for ch in invalid_chars:
            col_name = col_name.replace(ch, '_')
        return col_name[:255]

    def add_suffix(self):
        new_suffixes = self.suffix_entry.get().strip()
        # Allow comma, semicolon, or whitespace separated suffixes
        if new_suffixes:
            # Split on comma, semicolon, or whitespace
            for new_suffix in [s.strip() for s in re.split(r'[;,\s]+', new_suffixes) if s.strip()]:
                if new_suffix and new_suffix not in self.suffixes:
                    self.suffixes.append(new_suffix)
                    self.update_log(f"Added new suffix: {new_suffix}")
            if self.suffix_listbox.winfo_ismapped():
                self.update_suffix_listbox()
        self.suffix_entry.delete(0, tk.END)

    def update_suffix_listbox(self):
        self.suffix_listbox.delete(0, tk.END)
        for suf in self.suffixes:
            self.suffix_listbox.insert(tk.END, suf)

    def toggle_suffix_listbox(self):
        if self.suffix_listbox.winfo_ismapped():
            self.suffix_listbox.grid_remove()
            self.suffix_list_label.grid_remove()
            self.remove_suffix_button.grid_remove()
            self.show_suffixes_button.configure(text="Show Current Suffixes")
        else:
            self.update_suffix_listbox()
            self.suffix_listbox.grid()
            self.suffix_list_label.grid()
            self.remove_suffix_button.grid()
            self.show_suffixes_button.configure(text="Hide Current Suffixes")

    def remove_selected_suffix(self):
        # Remove all selected suffixes in the Listbox
        try:
            selection = list(self.suffix_listbox.curselection())
            if selection:
                # Remove from highest index to lowest to avoid shifting
                for idx in sorted(selection, reverse=True):
                    removed = self.suffixes[idx]
                    del self.suffixes[idx]
                    self.update_log(f"Removed suffix: {removed}")
                self.update_suffix_listbox()
        except Exception:
            pass

    def clean_class_name(self, class_name):
        """
        Removes known suffix extensions from the end of a class name.
        Extensions are case-insensitive and only removed if at the end.
        """
        for ext in self.suffixes:
            if class_name.upper().endswith(ext.upper()):
                return class_name[: -len(ext)]
        return class_name
    async def select_marketplace_dropdown(self, page, marketplace_id):
        """
        Selects marketplace from dropdown by first getting label from marketplace_id,
        then searching for matching option in dropdown.
        """
        # Get marketplace label from id
        label = self.MARKETPLACE_MAP.get(str(marketplace_id).strip().upper())
        if not label:
            self.update_log(f"Unknown marketplace_id '{marketplace_id}', skipping dropdown selection.")
            return

        try:
            # Find and click dropdown trigger with text "All Marketplaces"
            self.update_log("Looking for marketplace dropdown trigger...")
            dropdown_trigger = page.locator('text="All marketplaces"')
            await dropdown_trigger.wait_for(state="visible", timeout=10000)
            self.update_log("Found dropdown trigger, scrolling into view...")
            await dropdown_trigger.scroll_into_view_if_needed()
            self.update_log("Clicking dropdown trigger...")
            await dropdown_trigger.click()
            self.update_log("Clicked marketplace dropdown, waiting for options to load...")
            await page.wait_for_timeout(3000)

            try:
                # Get all dropdown options
                self.update_log("Getting dropdown options...")
                options = await page.locator('.awsui-select-option').all_text_contents()
                self.update_log(f"Found {len(options)} dropdown options: {options}")

                # Check if our label exists in options
                if label not in options:
                    self.update_log(f"Warning: Label '{label}' not found in dropdown options")
                    # Try case-insensitive match
                    matching_options = [opt for opt in options if opt.lower() == label.lower()]
                    if matching_options:
                        self.update_log(f"Found case-insensitive match: {matching_options[0]}")
                        label = matching_options[0]
                    else:
                        self.update_log("No matching option found even with case-insensitive comparison")
                        return

                # Find and click the option matching our marketplace label using class selector
                self.update_log(f"Attempting to select option '{label}'...")
                marketplace_option = page.locator(f'.awsui-select-option:has-text("{label}")')
                await marketplace_option.wait_for(state="visible", timeout=5000)
                self.update_log("Found matching option, clicking...")
                await marketplace_option.click()
                self.update_log(f"Successfully selected marketplace '{label}' for id '{marketplace_id}'")
                
                # Verify selection
                await page.wait_for_timeout(1000)
                selected_text = await dropdown_trigger.text_content()
                self.update_log(f"Verification - Selected value shows as: {selected_text}")
                if label not in selected_text:
                    self.update_log(f"Warning: Selected value '{selected_text}' does not match expected '{label}'")
                return

            except Exception as e:
                self.update_log(f"Error in dropdown option selection: {str(e)}")
                self.update_log(f"Stack trace: {e.__traceback__}")
                return

        except Exception as e:
            self.update_log(f"Error in marketplace dropdown handling: {str(e)}")
            self.update_log(f"Stack trace: {e.__traceback__}")
   # Add a footer label for tool ownership
    def mainloop(self, *args, **kwargs):
        # Add footer before mainloop
        self.footer_label = ctk.CTkLabel(
            self,
            text="Created By Shubham Dayama",
            font=ctk.CTkFont(size=11, slant="italic"),
            text_color="#888888"  # Subtle gray for background blending
        )
        self.footer_label.place(relx=0.5, rely=0.98, anchor="s")
        super().mainloop(*args, **kwargs)

def main():
    os.environ['PYPPETEER_CHROMIUM_REVISION'] = ''
    try:
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        app = XCPToolGUI()
        app.mainloop()
    except Exception as e:
        logging.error(f"Application error: {str(e)}", exc_info=True)
        messagebox.showerror("Error", f"Application error: {str(e)}")

if __name__ == "__main__":
    main()
