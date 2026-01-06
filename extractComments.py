import os
import time
import base64
import io
import argparse
import sys
from ctypes import windll
from datetime import datetime

import win32com.client as win32
import win32gui
import win32ui
from PIL import Image
import xlsxwriter

# --- DPI Awareness for Multi-Monitor Support ---
try:
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# --- Core Extraction Logic ---

def capture_window_background(hwnd):
    """
    Captures window content directly from memory (works if window is obscured).
    """
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        wDC = win32gui.GetWindowDC(hwnd)
        dcObj = win32ui.CreateDCFromHandle(wDC)
        cDC = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, width, height)
        cDC.SelectObject(dataBitMap)

        # PrintWindow with PW_RENDERFULLCONTENT (flag 2)
        result = windll.user32.PrintWindow(hwnd, cDC.GetSafeHdc(), 2)
        if result == 0:
            result = windll.user32.PrintWindow(hwnd, cDC.GetSafeHdc(), 0)

        bmpinfo = dataBitMap.GetInfo()
        bmpstr = dataBitMap.GetBitmapBits(True)
        
        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())

        return img
    except:
        return None

def extract_comments_data(docx_path):
    """
    Orchestrates Word to extract text, meta-data, and screenshots.
    Returns a list of dictionaries containing raw PIL images.
    """
    abs_path = os.path.abspath(docx_path)
    if not os.path.exists(abs_path):
        print(f"Error: File '{abs_path}' not found.")
        sys.exit(1)

    print(f"Initializing Word for '{os.path.basename(docx_path)}'...")
    word = win32.Dispatch("Word.Application")
    word.Visible = True 
    
    try:
        doc = word.Documents.Open(abs_path)
        if word.WindowState == 2: # Restore if minimized
            word.WindowState = 0
        word.Activate()
        time.sleep(2)
        
        hwnd = word.ActiveWindow.Hwnd
        total_comments = doc.Comments.Count
        results = []

        print(f"Processing {total_comments} comments...")

        for i, comment in enumerate(doc.Comments):
            idx = i + 1
            
            # Metadata
            author = comment.Author
            text = comment.Range.Text.strip()
            try:
                c_date = comment.Date
                date_str = c_date.strftime("%Y-%m-%d %H:%M")
            except:
                date_str = "Unknown"

            page_num = comment.Reference.Information(3) # wdActiveEndPageNumber

            # Snapshot Action
            comment.Reference.Select()
            word.ActiveWindow.ScrollIntoView(comment.Reference)
            time.sleep(0.3) # Render buffer

            img = capture_window_background(hwnd)
            
            results.append({
                "id": idx,
                "page": page_num,
                "author": author,
                "date": date_str,
                "text": text,
                "image_obj": img  # Store raw PIL object
            })
            
            #sys.stdout.write(f"\rCaptured [{idx}/{total_comments}]")
            #sys.stdout.flush()
            print(f"[{idx}/{total_comments}] Processed Page {page_num} - {author} - {date_str}")
        
        print("\nExtraction complete.")
        return results

    except Exception as e:
        print(f"Error during extraction: {e}")
        return []
    finally:
        try:
            doc.Close(SaveChanges=False)
            word.Quit()
        except:
            pass

# --- Report Generators ---

def save_as_html(data, doc_name, output_path):
    print("Generating HTML Report...")
    
    html_rows = ""
    for c in data:
        # Convert PIL image to Base64 for HTML
        img_str = ""
        if c['image_obj']:
            buffered = io.BytesIO()
            c['image_obj'].save(buffered, format="PNG")
            b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
            img_str = f"data:image/png;base64,{b64}"

        html_rows += f"""
            <tr>
                <td style="text-align:center; font-weight:bold; color:#555;">{c['page']}</td>
                <td class="content">{c['text']}</td>
                <td class="screenshot"><img src="{img_str}" alt="Context"></td>
                <td class="meta"><strong>{c['author']}</strong><br>{c['date']}</td>
            </tr>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Report: {doc_name}</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f4f4f9; padding: 20px; }}
            h1 {{ text-align: center; color: #333; }}
            table {{ width: 100%; border-collapse: collapse; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            th {{ background: #0078d4; color: white; padding: 12px; text-align: left; }}
            td {{ border-bottom: 1px solid #ddd; padding: 12px; vertical-align: top; }}
            .screenshot img {{ width: 100%; max-width: 400px; border: 1px solid #ccc; transition: transform 0.2s; }}
            .screenshot img:hover {{ transform: scale(1.8); position:relative; z-index:10; border-color: #0078d4; cursor: zoom-in; }}
            .content {{ width: 35%; }}
            .meta {{ width: 15%; font-size: 0.9em; color: #666; }}
        </style>
    </head>
    <body>
        <h1>Comment Report: {doc_name}</h1>
        <table>
            <thead><tr><th>Pg</th><th>Comment</th><th>Context</th><th>Meta</th></tr></thead>
            <tbody>{html_rows}</tbody>
        </table>
    </body>
    </html>
    """
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"HTML Report saved to: {output_path}")

def save_as_xlsx(data, doc_name, output_path):
    print("Generating Excel Report...")
    
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet("Comments")
    
    # Formats
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#0078d4', 'font_color': 'white', 'border': 1})
    text_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
    center_fmt = workbook.add_format({'align': 'center', 'valign': 'top', 'border': 1})
    
    # Set Column Widths
    worksheet.set_column('A:A', 5)   # Page
    worksheet.set_column('B:B', 40)  # Comment Text
    worksheet.set_column('C:C', 50)  # Screenshot Column
    worksheet.set_column('D:D', 20)  # Author
    worksheet.set_column('E:E', 15)  # Date

    # Write Header
    headers = ["Pg", "Comment Text", "Screenshot", "Author", "Date"]
    for col, h in enumerate(headers):
        worksheet.write(0, col, h, header_fmt)
        
    # Write Data
    for i, item in enumerate(data):
        row = i + 1
        
        # 1. Write Text Data
        worksheet.write(row, 0, item['page'], center_fmt)
        worksheet.write(row, 1, item['text'], text_fmt)
        # (Skip col 2 for image)
        worksheet.write(row, 3, item['author'], text_fmt)
        worksheet.write(row, 4, item['date'], center_fmt)
        
        # 2. Handle Image
        if item['image_obj']:
            # We must resize image to act as a thumbnail, otherwise Excel file gets huge
            # and rows become massive. Let's resize to width 350px.
            orig_w, orig_h = item['image_obj'].size
            aspect_ratio = orig_h / orig_w
            new_w = 350
            new_h = int(new_w * aspect_ratio)
            
            thumb = item['image_obj'].resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            # Save to BytesIO stream
            image_stream = io.BytesIO()
            thumb.save(image_stream, format="PNG")
            
            # Insert into Excel
            worksheet.insert_image(row, 2, "img.png", {
                'image_data': image_stream,
                'x_offset': 5,
                'y_offset': 5,
                'object_position': 1 # Move and size with cells
            })
            
            # Set row height to match image height (pixels to points approx conversion)
            # 1 pixel approx 0.75 points
            worksheet.set_row(row, new_h * 0.75 + 10)
        else:
            worksheet.set_row(row, 50) # Default height if no image

    workbook.close()
    print(f"Excel Report saved to: {output_path}")

# --- Main CLI ---

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract Word comments to HTML or XLSX.")
    parser.add_argument("filename", help="Path to .docx file")
    parser.add_argument("-f", "--format", choices=['html', 'xlsx', 'both'], default='html', help="Output format (default: html)")
    parser.add_argument("-o", "--output", help="Custom output filename (optional)")
    
    args = parser.parse_args()
    
    # 1. Extract Data
    raw_data = extract_comments_data(args.filename)
    
    if not raw_data:
        print("No data extracted.")
        sys.exit()

    base_name = os.path.splitext(args.filename)[0]
    
    # 2. Generate Reports based on choice
    if args.format in ['html', 'both']:
        out_name = args.output if (args.output and args.output.endswith('.html')) else f"{base_name}_report.html"
        save_as_html(raw_data, os.path.basename(args.filename), out_name)
        
    if args.format in ['xlsx', 'both']:
        out_name = args.output if (args.output and args.output.endswith('.xlsx')) else f"{base_name}_report.xlsx"
        # If both selected, ensure extensions don't conflict
        if args.format == 'both' and args.output:
            # strip extension and re-add
            clean_name = os.path.splitext(args.output)[0]
            out_name = f"{clean_name}.xlsx"
            
        save_as_xlsx(raw_data, os.path.basename(args.filename), out_name)