import customtkinter as ctk
import pandas as pd
import os
import time
import threading
import shutil
import glob
import json

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pathlib import Path

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

OUTPUT_SUBDIR = str(Path.home() / "Downloads" / "Download WHT")
TEMP_PDF_DIR = "TEMP_PDF"

def show_done_popup(message):
    popup = ctk.CTkToplevel()
    popup.title("Success")
    popup.geometry("400x200")
    popup.configure(fg_color="#537D5D")
    popup.lift()
    popup.attributes("-topmost", True)
    popup.after(100, lambda: popup.attributes("-topmost", False))

    label = ctk.CTkLabel(popup, text=message, font=("Tahoma", 16, "bold"),
                            text_color="white", wraplength=350)
    label.pack(pady=(40, 20))

    def close_all():
        popup.destroy()
        app.destroy()

    btn = ctk.CTkButton(popup, text="OK", command=close_all,
                        fg_color="#131D4F", hover_color="#39487A",
                        text_color="white", corner_radius=10, width=100, height=40)
    btn.pack(pady=10)

def get_csv_url_from_gsheet_url(gsheet_url):
    try:
        sheet_id = gsheet_url.split("/d/")[1].split("/")[0]
        gid = gsheet_url.split("gid=")[1].split("&")[0] if "gid=" in gsheet_url else "0"
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        return csv_url
    except Exception as e:
        raise ValueError(f"Error parsing Google Sheet URL: {e}")

def download_pdfs_from_gsheet(gsheet_url, status_label, progress_bar, count_label):
    try:
        csv_url = get_csv_url_from_gsheet_url(gsheet_url)
    except Exception as e:
        status_label.configure(text=f"❌ {e}")
        return

    try:
        df_raw = pd.read_csv(csv_url, header=7)
    except Exception as e:
        status_label.configure(text=f"❌ Failed to load sheet: {e}")
        return

    expected_cols = ["URL", "เอกสารอ้างอิง", "เลขที่ใบหัก ณ ที่จ่าย", "ผู้รับเงิน"]
    for col in expected_cols:
        if col not in df_raw.columns:
            status_label.configure(text=f"❌ Missing column '{col}' in sheet")
            return

    df = df_raw.loc[:, expected_cols].dropna(subset=["URL"]).reset_index(drop=True)

    output_dir = os.path.join(os.getcwd(), OUTPUT_SUBDIR)
    temp_dir = os.path.join(os.getcwd(), TEMP_PDF_DIR)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(temp_dir, exist_ok=True)

    chrome_prefs = {
        "printing.print_preview_sticky_settings.appState": json.dumps({
            "recentDestinations": [{"id": "Save as PDF", "origin": "local"}],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isHeaderFooterEnabled": False,
            "marginsType": 1,
            "mediaSize": {
                "name": "ISO_A4",
                "custom_display_name": "A4",
                "width_microns": 210000,
                "height_microns": 297000
            },
            "scalingType": 3,
            "scaling": "100"
        }),
        "savefile.default_directory": os.path.abspath(temp_dir),
        "download.default_directory": os.path.abspath(temp_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }

    options = Options()
    options.add_argument("--kiosk-printing")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1200,800")
    options.add_argument('--lang=en-GB')
    options.add_argument('--default-pdf-paper-size=A4')
    options.add_argument("--force-device-scale-factor=1")
    options.add_experimental_option("prefs", chrome_prefs)

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    total = len(df)
    completed = 0

    for idx, row in df.iterrows():
        link = row['URL']
        rr_no = str(row['เอกสารอ้างอิง'])
        wht_no = str(row['เลขที่ใบหัก ณ ที่จ่าย'])
        name = str(row['ผู้รับเงิน'])

        safe_name = name.replace("/", "_").replace("\\", "_")
        new_filename = f"{wht_no}_{rr_no}_{safe_name}.pdf"
        final_path = os.path.join(output_dir, new_filename)

        try:
            for f in glob.glob(os.path.join(temp_dir, "*.pdf")):
                os.remove(f)

            driver.get(link)

            print_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(text(), 'พิมพ์เอกสาร') or contains(text(), 'Print')]")
            ))
            print_btn.click()
            time.sleep(1)  # ให้เวลาหน้า print popup render

            timeout = 30
            start = time.time()
            downloaded_file = None
            while time.time() - start < timeout:
                pdf_files = glob.glob(os.path.join(temp_dir, "*.pdf"))
                if pdf_files:
                    downloaded_file = pdf_files[0]
                    break
                time.sleep(1)

            if downloaded_file:
                shutil.move(downloaded_file, final_path)
                print(f"[{idx+1}] ✅ Saved: {final_path}")
                completed += 1
            else:
                print(f"[{idx+1}] ❌ PDF not found for {link}")

        except Exception as e:
            print(f"[{idx+1}] ❌ Error on {link}: {e}")

        progress_bar.set((idx + 1) / total)
        count_label.configure(text=f"Completed {completed}/{total}")
        status_label.configure(text=f"Processing {idx + 1} / {total} ...")
        app.update()

    driver.quit()
    show_done_popup("✅ All files downloaded and renamed successfully!")
    status_label.configure(text="Completed.")

def start_download():
    url = url_entry.get().strip()
    if not url:
        status_label.configure(text="❌ Please enter Google Sheet URL")
        return

    progress_bar.set(0)
    count_label.configure(text="Completed 0/0")
    status_label.configure(text="Starting...")

    threading.Thread(target=download_pdfs_from_gsheet,
                        args=(url, status_label, progress_bar, count_label),
                        daemon=True).start()

# ==== GUI ====
app = ctk.CTk()
app.title("Download WHT PDFs from Google Sheet")
app.geometry("600x400")
app.configure(fg_color="#F7FCF4")

title = ctk.CTkLabel(app, text="Enter Google Sheet URL (WHT) to start downloading",
                        font=("Tahoma", 16, "bold"), text_color="#131D4F")
title.pack(pady=20)

url_entry = ctk.CTkEntry(app, placeholder_text="Paste your Google Sheet URL here",
                            width=500, height=35)
url_entry.pack(pady=10)

btn = ctk.CTkButton(app,
                    text="▶️ Start Download",
                    command=start_download,
                    corner_radius=10,
                    fg_color="#537D5D",
                    hover_color="#537D5D",
                    text_color="white",
                    width=200,
                    height=40)
btn.pack(pady=10)

progress_bar = ctk.CTkProgressBar(app,
                                    width=400,
                                    progress_color="#537D5D",
                                    fg_color="#BBDFC8")
progress_bar.set(0)
progress_bar.pack(pady=10)

count_label = ctk.CTkLabel(app, text="Completed 0/0", font=("Tahoma", 14), text_color="#333")
count_label.pack(pady=(5, 0))

status_label = ctk.CTkLabel(app, text="", font=("Tahoma", 14), text_color="#333")
status_label.pack(pady=(5, 20))

app.mainloop()
