#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import urllib.request
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

def download_image(url, filename):
    """Download an image from URL"""
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        req = urllib.request.Request(url, headers=headers)
        response = urllib.request.urlopen(req, timeout=10)

        with open(filename, 'wb') as f:
            f.write(response.read())
        print(f"✓ Downloaded: {filename}")
        return True
    except Exception as e:
        print(f"✗ Failed to download from {url}: {e}")
        return False

# Try multiple potential logo URLs
potential_urls = [
    "https://ctp-llp.com/wp-content/uploads/2021/03/CTP-Logo.png",
    "https://ctp-llp.com/wp-content/uploads/2021/03/CTP-Logo.jpg",
    "https://ctp-llp.com/wp-content/themes/ctp/images/logo.png",
    "https://ctp-llp.com/images/logo.png",
    "https://ctp-llp.com/logo.png",
    "https://www.c-t-p.co.za/wp-content/uploads/2020/01/CTP-Logo.png",
]

logo_file = None
for url in potential_urls:
    filename = f"ctp_logo_{potential_urls.index(url)}.png"
    if download_image(url, filename):
        logo_file = filename
        break

if logo_file:
    print(f"\n✓ Logo downloaded successfully: {logo_file}")

    # Try to insert into Excel
    try:
        wb = load_workbook("CTP_Document_Submittal_Response_DSR_Template.xlsx")
        ws = wb.active

        # Add image to the header
        img = XLImage(logo_file)
        img.width = 200
        img.height = 60
        ws.add_image(img, 'A1')

        wb.save("CTP_Document_Submittal_Response_DSR_Template.xlsx")
        print("✓ Logo inserted into Excel template!")

    except Exception as e:
        print(f"✗ Could not insert logo into Excel: {e}")
else:
    print("\n✗ Could not download logo from any source")
    print("You can manually insert the logo from: https://ctp-llp.com/")
