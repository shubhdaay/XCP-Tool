import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import nest_asyncio
from typing import List, Dict
import logging
import threading
from PIL import Image
import os
import sys
import datetime
import glob

# Apply nest_asyncio to allow nested event loops
nest_asyncio.apply()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='xcp_tool.log'
)

class XCPToolGUI(ctk.CTk):
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
            text="XCP Tool Automation",
            font=ctk.CTkFont(size=24, weight="bold")
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
        playwright = None
        browser = None
        export_dir = f"exports_{datetime.date.today()}"
        os.makedirs(export_dir, exist_ok=True)
        try:
            input_file = self.file_path.get()
            if not input_file:
                messagebox.showerror("Error", "Please select an input file")
                return

            self.update_status("Initializing...")
            self.update_progress(0.1)

            # Read Excel file
            df = pd.read_excel(input_file)
            self.update_log(f"Successfully loaded {len(df)} rows from Excel")
            self.update_progress(0.2)

            playwright = await async_playwright().start()
            browser = await playwright.chromium.launch(headless=False, args=['--start-maximized'])
            context = await browser.new_context(viewport={"width": 1920, "height": 1080})
            page = await context.new_page()
            self.update_log("Browser initialized successfully")
            self.update_progress(0.3)

            # Navigate to CP Central
            await page.goto('https://www.cp-central.catalog.amazon.dev/#/class/search')
            self.update_log("Navigated to CP Central")
            self.update_progress(0.4)

            # --- SSO Login Handling ---
            if "SSO/redirect" in page.url or "midway-auth.amazon.com" in page.url:
                self.update_log("SSO login required. Please complete the login in the opened browser window.")
                try:
                    await page.wait_for_selector('#awsui-input-0', timeout=0)  # Wait indefinitely
                    self.update_log("Login successful. Continuing automation.")
                except Exception as e:
                    self.update_log(f"Error waiting for login: {str(e)}")
                    return

            # Check if already on Class Search page, else navigate
            class_search_url = 'https://www.cp-central.catalog.amazon.dev/#/class/search'
            if not page.url.startswith(class_search_url):
                await page.goto(class_search_url)
                self.update_log("Navigated to Class Search page.")
            else:
                self.update_log("Already on Class Search page.")
            await page.wait_for_selector('#awsui-input-0', timeout=10000)
            await page.wait_for_timeout(1000)

            # Process each class
            for class_name, group in df.groupby('Class'):
                if not self.is_processing:
                    self.update_log("Processing stopped by user")
                    break
                self.update_log(f"Processing class: {class_name}")
                try:
                    # Always reload the page before entering class name for stability
                    await page.reload(wait_until="domcontentloaded")
                    await page.wait_for_selector('#awsui-input-0', timeout=20000)
                    await page.wait_for_timeout(1000)
                    for attempt in range(2):
                        try:
                            input_box = page.locator('#awsui-input-0')
                            await input_box.wait_for(state="visible", timeout=10000)
                            await input_box.wait_for(state="attached", timeout=10000)
                            await input_box.scroll_into_view_if_needed()
                            await input_box.focus()
                            await page.wait_for_timeout(300)  # Give a short delay for focus
                            await input_box.fill('')
                            await input_box.type(class_name, delay=50)  # Slow typing to avoid race
                            await page.keyboard.press('Enter')
                            await page.wait_for_timeout(2000)
                            break
                        except Exception as e:
                            self.update_log(f"Attempt {attempt+1}: Could not enter class name, refreshing page... Error: {str(e)}")
                            await page.reload(wait_until="domcontentloaded")
                            await page.wait_for_selector('#awsui-input-0', timeout=20000)
                            await page.wait_for_timeout(1000)
                    # Wait up to 10 seconds for the input box after reload
                    input_found = False
                    for _ in range(20):  # 20 x 500ms = 10s
                        try:
                            await page.wait_for_selector('#awsui-input-0', timeout=500)
                            input_found = True
                            break
                        except Exception:
                            await page.wait_for_timeout(500)
                    if not input_found:
                        self.update_log("Warning: Input box still not found after waiting 10 seconds after reload.")
                    await page.wait_for_timeout(500)
                    # Try clicking the class name link directly (anchor tag)
                    class_link = page.locator(f"a:text-is('{class_name}')")
                    await class_link.first.wait_for(timeout=5000)
                    await class_link.first.click()
                    await page.wait_for_timeout(2000)
                    # Wait for the 'New sample ASINs test' button to be visible and click it
                    try:
                        # Try CSS selector first
                        sample_test_btn = page.locator('#app-content > div > div > div:nth-child(1) > div > div.awsui-util-action-stripe-large > div.awsui-util-action-stripe-group.awsui-util-pv-n > awsui-button:nth-child(1) > a')
                        await sample_test_btn.wait_for(timeout=5000)
                        await sample_test_btn.click()
                    except Exception:
                        try:
                            # Try XPath as fallback
                            sample_test_btn_xpath = page.locator('xpath=//*[@id="app-content"]/div/div/div[1]/div/div[2]/div[2]/awsui-button[1]/a')
                            await sample_test_btn_xpath.wait_for(timeout=2000)
                            await sample_test_btn_xpath.click()
                        except Exception as e:
                            self.update_log(f"Could not click 'New sample ASINs test' button: {str(e)}")
                    await page.wait_for_timeout(2000)
                    # Uncheck the box if it is checked
                    try:
                        # Use a more robust selector for the checkbox
                        checkbox = page.locator('input[type="checkbox"]')
                        # There may be more than one checkbox, so check the label text
                        label = await checkbox.evaluate_handle('el => el.parentElement.textContent')
                        if 'Include sample ASINs provided during the class authoring process' in await label.json_value():
                            if await checkbox.is_checked():
                                await checkbox.click()
                                self.update_log("Unchecked the 'Include sample ASINs' box.")
                    except Exception as e:
                        self.update_log(f"Could not uncheck the box: {str(e)}")
                    await page.wait_for_timeout(1000)
                    # Input ASINs for this class
                    try:
                        # Wait for all matching textareas and pick the correct one by position (usually 'Must match ASINs' is the first or second)
                        await page.wait_for_selector('textarea[placeholder^="Enter ASIN"]', timeout=10000)
                        asin_inputs = page.locator('textarea[placeholder^="Enter ASIN"]')
                        count = await asin_inputs.count()
                        asin_input_area = None
                        asin_input_index = 0
                        # Try to pick the textarea that is closest to the left (first in DOM)
                        if count > 1:
                            # Log all textarea ids for debugging
                            ids = []
                            for idx in range(count):
                                handle = asin_inputs.nth(idx)
                                id_val = await handle.get_attribute('id')
                                ids.append(id_val)
                            self.update_log(f"Found {count} ASIN textareas with ids: {ids}")
                            # Try the first one, if not visible, try the second
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
                        asins = group['asin_id'].astype(str).tolist()
                        asin_text = '\n'.join(asins)
                        await asin_input_area.fill(asin_text)
                        self.update_log(f"Filled ASINs textarea (index {asin_input_index}) with {len(asins)} ASINs.")
                        await page.wait_for_timeout(1000)
                    except Exception as e:
                        self.update_log(f"Could not input ASINs or click test button: {str(e)}")
                        continue
                    # Click the 'Test sample ASINs' button robustly
                    try:
                        test_btn = page.locator('button:has(span:text("Test sample ASINs")), button:has-text("Test sample ASINs")')
                        await test_btn.wait_for(timeout=5000)
                        await test_btn.click()
                        self.update_log("Clicked 'Test sample ASINs' button.")
                    except Exception as e:
                        self.update_log(f"Could not click 'Test sample ASINs' button: {str(e)}")

                    # Wait for the export button to become visible and enabled (indicating test completion)
                    try:
                        export_btn = page.locator('#app-content > div > div:nth-child(3) > div.test-sample-asins-component > div:nth-child(4) > awsui-table > div > div.awsui-table-heading-container > div > div.awsui-table-header > span > div > div.awsui-util-action-stripe-group > awsui-button > button')
                        await export_btn.wait_for(state="visible", timeout=120000)  # Wait up to 2 minutes
                        await export_btn.scroll_into_view_if_needed()
                        while not await export_btn.is_enabled():
                            await page.wait_for_timeout(1000)
                        self.update_log("ASINs tested, export button is now enabled.")
                    except Exception as e:
                        self.update_log(f"Export button did not become available: {str(e)}")

                    # Now click the export button to download the results
                    try:
                        self.update_log("Attempting to click the export button for download.")
                        await export_btn.hover()
                        await page.wait_for_timeout(500)  # Give a short delay after hover
                        async with page.expect_download() as download_info:
                            await export_btn.click(force=True)
                        download = await download_info.value
                        class_export_name = os.path.join(export_dir, f"export_{class_name.replace('/', '_').replace(' ', '_')}.xlsx")
                        await download.save_as(class_export_name)
                        self.update_log(f"Downloaded export for class {class_name} as {class_export_name}")
                        # Navigate back to Class Search page for the next class
                        class_search_url = 'https://www.cp-central.catalog.amazon.dev/#/class/search'
                        await page.goto(class_search_url, wait_until="domcontentloaded")
                        try:
                            await page.wait_for_selector('#awsui-input-0', timeout=20000)
                        except Exception as e:
                            self.update_log("Input box not found, refreshing page and retrying...")
                            await page.reload(wait_until="domcontentloaded")
                            await page.wait_for_selector('#awsui-input-0', timeout=20000)
                        self.update_log("Returned to Class Search page for next class.")
                    except Exception as e:
                        self.update_log(f"Could not export results for class {class_name}: {str(e)}")
                    progress = min(0.4 + 0.6 * (1.0), 1.0)
                    self.update_progress(progress)
                except Exception as e:
                    self.update_log(f"Error processing class {class_name}: {str(e)}")
                    continue
            self.update_status("Processing complete")
            self.update_progress(1.0)
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
            # Collate all exported files into one Excel sheet if any were downloaded
            try:
                all_files = glob.glob(os.path.join(export_dir, 'export_*.xlsx'))
                if all_files:
                    dfs = []
                    for file in all_files:
                        try:
                            # Try reading as Excel, if fails, try as CSV
                            df = pd.read_excel(file)
                        except Exception:
                            try:
                                df = pd.read_csv(file)
                            except Exception:
                                self.update_log(f"Could not read file {file} as Excel or CSV. Skipping.")
                                continue
                        df['source_file'] = os.path.basename(file)
                        # Sanitize column names for Excel compatibility
                        df.columns = [self.sanitize_excel_column(col) for col in df.columns]
                        dfs.append(df)
                    if dfs:
                        combined = pd.concat(dfs, ignore_index=True)
                        combined_file = os.path.join(export_dir, f"collated_exports_{datetime.date.today()}.xlsx")
                        # Always use a safe sheet name
                        with pd.ExcelWriter(combined_file, engine='openpyxl') as writer:
                            combined.to_excel(writer, index=False, sheet_name="Sheet1")
                        self.update_log(f"Collated all exports into {combined_file}")
            except Exception as e:
                self.update_log(f"Error collating exports: {str(e)}")

    def start_processing(self):
        if not self.is_processing:
            self.is_processing = True
            self.start_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
            # Run the async function in the main thread using create_task
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
        # Remove or replace invalid Excel characters
        invalid_chars = ['/', '\\', '?', '*', '[', ']', ':', ';', '\\n', '\\r', '\\t', '|']
        for ch in invalid_chars:
            col_name = col_name.replace(ch, '_')
        # Excel column name max length is 255
        return col_name[:255]

def main():
    # Set environment variable to prevent auto-download
    os.environ['PYPPETEER_CHROMIUM_REVISION'] = ''
    
    try:
        # Set appearance mode
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        app = XCPToolGUI()
        app.mainloop()
    except Exception as e:
        logging.error(f"Application error: {str(e)}", exc_info=True)
        messagebox.showerror("Error", f"Application error: {str(e)}")

if __name__ == "__main__":
    main()
