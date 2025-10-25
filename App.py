import os
import re
import time
import zipfile
import random
import json
import unicodedata
from io import BytesIO
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd
import requests
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

import yt_dlp
from moviepy.editor import VideoFileClip

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

st.set_page_config(page_title="SocialScribe", layout="wide")
st.title("SocialScribe • YouTube Automation Suite")

OUTPUT_ROOT = Path("output_videos")
OUTPUT_ROOT.mkdir(exist_ok=True)
COOKIES_FILE = None  # optional for YouTube login cookies

# ---------------- Utility Functions ----------------

def sanitize_name(name):
    if not name:
        return "Unknown"
    name = unicodedata.normalize("NFKC", str(name))
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = name.rstrip(". ")
    return "".join(ch for ch in name if ch.isprintable()).strip() or "Unknown"

def parse_timecode(tc):
    tc = str(tc).strip().replace(".", ":")
    parts = [int(p) for p in tc.split(":")]
    if len(parts) == 1: return parts[0]
    if len(parts) == 2: return parts[0]*60 + parts[1]
    if len(parts) == 3: return parts[0]*3600 + parts[1]*60 + parts[2]
    return None

# ---------------- YouTube Report Logic ----------------

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ('top','left','bottom','right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'w:{edge}'
            border_element = tcBorders.find(qn(tag))
            if border_element is None:
                border_element = OxmlElement(tag)
                tcBorders.append(border_element)
            for key,val in edge_data.items():
                border_element.set(qn(f'w:{key}'), str(val))

def add_metadata_to_cell(cell, label, value, is_link=False):
    p = cell.add_paragraph()
    p_format = p.paragraph_format
    p_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p_format.space_before = Pt(0)
    p_format.space_after = Pt(6)

    run = p.add_run(label + " ")
    run.bold = True
    run.font.color.rgb = RGBColor(255,0,0)
    run_val = p.add_run(value)
    if is_link:
        run_val.font.color.rgb = RGBColor(5,99,193)
        run_val.font.underline = True
    return p

def create_youtube_report(youtube_url, doc=None):
    """Full logic: fetch metadata, screenshot, thumbnail, append to Word doc"""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--log-level=3")
    options.add_argument("window-size=1920,1280")
    options.add_argument("--mute-audio")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    thumbnail_card_path = "temp_thumbnail_card.png"
    main_screenshot_path = "temp_main_screenshot.png"
    channel_name = None

    try:
        driver.get(youtube_url)
        wait = WebDriverWait(driver, 15)
        script_element = wait.until(EC.presence_of_element_located((By.XPATH,"//script[@type='application/ld+json']")))
        data = json.loads(script_element.get_attribute('innerHTML'))
        video_title = data.get('name','Unknown Title')
        thumbnail_url = data.get('thumbnailUrl',[None])[0]
        upload_date_raw = data.get('uploadDate','')
        upload_date = ''
        if upload_date_raw:
            try:
                upload_date = datetime.strptime(upload_date_raw.split('T')[0], "%Y-%m-%d").strftime("%d-%m-%Y")
            except Exception:
                upload_date = upload_date_raw

        channel_name_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "yt-formatted-string.ytd-channel-name")))
        channel_name = channel_name_element.text.strip()
        channel_link = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"ytd-video-owner-renderer a.yt-simple-endpoint"))).get_attribute('href')

        # Expand description if exists
        try:
            more_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "tp-yt-paper-button#expand")))
            more_btn.click()
            time.sleep(1)
        except Exception:
            pass

        # Screenshot main player
        try:
            driver.find_element(By.ID,"primary-inner").screenshot(main_screenshot_path)
        except Exception:
            driver.save_screenshot(main_screenshot_path)

        # Thumbnail fallback
        try:
            if thumbnail_url:
                img_data = Image.open(BytesIO(requests.get(thumbnail_url).content))
                img_data.save(thumbnail_card_path)
            else:
                thumbnail_card_path = None
        except Exception:
            thumbnail_card_path = None

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    if doc:
        doc.add_paragraph()
        meta_table = doc.add_table(rows=1, cols=1)
        meta_table.autofit = False
        try:
            meta_table.columns[0].width = Inches(7.7)
        except Exception:
            pass
        cell = meta_table.cell(0,0)
        cell.text=""
        add_metadata_to_cell(cell,"Title:",video_title)
        add_metadata_to_cell(cell,"Post Date:",upload_date)
        add_metadata_to_cell(cell,"Link:",youtube_url,is_link=True)
        add_metadata_to_cell(cell,"Channel:",channel_link,is_link=True)
        add_metadata_to_cell(cell,"Channel name:",channel_name)
        border = {"sz":12,"val":"single","color":"000000"}
        full_border = {"top":border,"bottom":border,"left":border,"right":border}
        set_cell_border(cell,**full_border)
        # Insert images
        for path in [thumbnail_card_path, main_screenshot_path]:
            if path and os.path.exists(path):
                try:
                    p = doc.add_paragraph()
                    p.add_run().add_picture(path,height=Inches(2.5))
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception:
                    continue
        # Cleanup
        for path in [thumbnail_card_path, main_screenshot_path]:
            if path and os.path.exists(path):
                os.remove(path)

    return channel_name

# ---------------- Video Trimmer Logic ----------------

def download_video(url, output_dir, cookies_file=None):
    opts = {"format":"best[height<=720]","outtmpl":os.path.join(output_dir,"temp_video.%(ext)s"),"quiet":True,"noplaylist":True}
    if cookies_file:
        opts["cookiefile"]=cookies_file
    with yt_dlp.YoutubeDL(opts) as ydl:
        info=ydl.extract_info(url,download=True)
        path=ydl.prepare_filename(info)
    return path

def trim_clip(input_file, start_s, end_s, output_file):
    with VideoFileClip(input_file) as video:
        start_s = max(0,float(start_s))
        end_s = min(float(video.duration), float(end_s))
        if start_s>=end_s: return False
        clip = video.subclip(start_s,end_s)
        clip.write_videofile(output_file, codec="libx264", audio_codec="aac", logger=None)
        clip.close()
    return True

# ---------------- Streamlit UI ----------------

tabs = st.tabs(["YouTube Report Generator","Video Trimmer"])

# ---------------- Tab 1: YouTube Reports ----------------
with tabs[0]:
    st.subheader("Generate YouTube Report")
    urls_input = st.text_area("Paste YouTube URLs (one per line):")
    if st.button("Generate Report"):
        urls = [u.strip() for u in urls_input.splitlines() if u.strip()]
        if not urls:
            st.warning("Please enter at least one URL")
        else:
            doc = Document()
            channel_name = None
            for i,url in enumerate(urls,1):
                st.write(f"Processing [{i}/{len(urls)}]: {url}")
                try:
                    ch = create_youtube_report(url, doc)
                    if channel_name is None: channel_name=ch
                except Exception as e:
                    st.error(f"Failed: {e}")
            fname = f"{sanitize_name(channel_name or 'YouTube_Report')}.docx"
            doc.save(fname)
            with open(fname,"rb") as f:
                st.download_button("📄 Download Report", data=f, file_name=fname, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.success("Report ready!")

# ---------------- Tab 2: Video Trimmer ----------------
with tabs[1]:
    st.subheader("Trim Videos from Excel")
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("Preview:")
        st.dataframe(df.head())
        if st.button("Process Videos"):
            zip_path = OUTPUT_ROOT / "trimmed_videos.zip"
            with zipfile.ZipFile(zip_path,"w") as zipf:
                for idx,row in df.iterrows():
                    video_url = str(row[0]).strip()
                    st.write(f"Processing video {idx+1}: {video_url}")
                    try:
                        temp_file = download_video(video_url, OUTPUT_ROOT)
                        # Example: first 5 seconds clip
                        out_file = OUTPUT_ROOT / f"{Path(temp_file).stem}_clip.mp4"
                        trim_clip(temp_file,0,5,out_file)
                        zipf.write(out_file)
                        os.remove(temp_file)
                        os.remove(out_file)
                    except Exception as e:
                        st.error(f"Failed {video_url}: {e}")
            with open(zip_path,"rb") as f:
                st.download_button("📦 Download All Trimmed Videos", data=f, file_name="trimmed_videos.zip", mime="application/zip")
            st.success("All videos processed!")
