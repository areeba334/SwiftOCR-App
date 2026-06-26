import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
import ctypes  # <--- Ye Naya Library hai Windows fix ke liye
import PyPDF2
import docx

# OCR libraries import 
try:
    import pytesseract
    from PIL import Image
    import pdf2image
except ImportError:
    pass

# --- DYNAMIC PATH SETUP ---
def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

base_path = get_base_path()

# Tesseract Path
tesseract_cmd_path = os.path.join(base_path, 'Tesseract-OCR', 'tesseract.exe')
pytesseract.pytesseract.tesseract_cmd = tesseract_cmd_path

# Poppler Path
path_option_1 = os.path.join(base_path, 'poppler', 'Library', 'bin')
path_option_2 = os.path.join(base_path, 'poppler', 'bin')

if os.path.exists(path_option_1):
    poppler_path = path_option_1
else:
    poppler_path = path_option_2

# --- SEARCH FUNCTIONS ---
def search_text_file(file_path, word):
    results = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for i, line in enumerate(file):
                count = line.lower().count(word.lower())
                if count > 0: results[i + 1] = count 
    except Exception as e: print(f"Error reading {file_path}: {e}")
    return results

def search_image_file(file_path, word):    
    results = {}
    try:
        img = Image.open(file_path)
        text_lines = pytesseract.image_to_string(img).splitlines()
        for i, line in enumerate(text_lines):
            count = line.lower().count(word.lower())
            if count > 0: results[i + 1] = count
    except Exception as e: print(f"Error OCR-ing {file_path}: {e}")
    return results

def search_pdf_file(file_path, word):
    results = {}
    text_found_digitally = False
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if text and text.strip():
                    text_found_digitally = True
                    count = text.lower().count(word.lower())
                    if count > 0: results[i + 1] = count 
    except Exception as e: print(f"PyPDF2 error on {file_path}: {e}")

    if not text_found_digitally:
        try:
            pages = pdf2image.convert_from_path(file_path, poppler_path=poppler_path)
            for i, page_image in enumerate(pages):
                text = pytesseract.image_to_string(page_image)
                count = text.lower().count(word.lower())
                if count > 0: results[i + 1] = count 
        except Exception as e: print(f"PDF-OCR error on {file_path}: {e}")
    return results

def search_word_file(file_path, word):
    results = {}
    try:
        doc = docx.Document(file_path)
        for i, para in enumerate(doc.paragraphs):
            count = para.text.lower().count(word.lower())
            if count > 0: results[i + 1] = count 
    except Exception as e: print(f"Error reading {file_path}: {e}")
    return results


# --- MAIN SEARCH LOGIC ---
def start_search():
    word = word_entry.get()
    folder_or_file = folder_entry.get()

    if not word or not folder_or_file:
        messagebox.showwarning("Input Error", "Please enter both the word and the path.")
        return

    if not os.path.exists(tesseract_cmd_path):
        messagebox.showerror("System Error", f"Tesseract not found at:\n{tesseract_cmd_path}")
        return

    search_results = {}
    files_to_search = []

    if os.path.isdir(folder_or_file):
        for root, dirs, files in os.walk(folder_or_file):
            for file in files: files_to_search.append(os.path.join(root, file))
    elif os.path.isfile(folder_or_file):
        files_to_search.append(folder_or_file) 
    else:
        messagebox.showerror("Path Error", "Invalid path selected.")
        return

    results_text.config(state=tk.NORMAL)
    results_text.delete(1.0, tk.END)
    results_text.insert(tk.END, "Searching... Please wait.\n\n")
    window.update_idletasks()

    for file_path in files_to_search:
        findings = {}
        ext = os.path.splitext(file_path)[1].lower() 
        try:
            if ext == '.txt' and var_txt.get() == 1: findings = search_text_file(file_path, word)
            elif ext == '.pdf' and var_pdf.get() == 1: findings = search_pdf_file(file_path, word)
            elif ext == '.docx' and var_docx.get() == 1: findings = search_word_file(file_path, word)
            elif ext in ['.jpg', '.jpeg', '.png'] and (var_jpg.get() == 1 or var_png.get() == 1):
                findings = search_image_file(file_path, word)
        except Exception as e: print(f"Error processing {file_path}: {e}")

        if findings: search_results[file_path] = findings

    results_text.delete(1.0, tk.END)
    if search_results:
        grand_total_count = 0
        for file_path, findings in search_results.items():
            results_text.insert(tk.END, f"File: {file_path}\n")
            unit = "Line"
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.pdf': unit = "Page"
            if ext == '.docx': unit = "Paragraph"
            if ext in ['.jpg', '.jpeg', '.png']: unit = "Image"
            
            file_total = 0
            for location, count in findings.items():
                if unit == "Image": results_text.insert(tk.END, f"  - Found: {count} time(s)\n")
                else: results_text.insert(tk.END, f"  - On {unit} {location}: {count} time(s)\n")
                file_total += count
            results_text.insert(tk.END, f"  Total in this file: {file_total}\n\n")
            grand_total_count += file_total
    else:
        results_text.insert(tk.END, "No Word Found.\n")
    results_text.config(state=tk.DISABLED)


# --- FIXED BROWSE FUNCTION (THE FIX IS HERE) ---
# --- BROWSE FUNCTION (Sirf Folder/Drive ke liye) ---
def browse_folder_or_file():
    # 1. Windows Crash Fix (Zaroori hai)
    try:
        ctypes.windll.ole32.CoInitialize(None)
    except Exception:
        pass
    
    # 2. Window ko refresh karein taaki hang na ho
    window.update_idletasks()
    
    try:
        # parent=window lagane se dialog box humesha app ke upar khulega
        folder_selected = filedialog.askdirectory(parent=window, title="Select Folder or Drive")
        
        # Agar user ne folder select kiya (Cancel nahi kiya)
        if folder_selected:
            folder_entry.delete(0, tk.END)
            folder_entry.insert(0, folder_selected)
            
            # App window ko wapis focus mein laayein
            window.lift()
            window.focus_force()
            
    except Exception as e:
        messagebox.showerror("Error", f"Folder select karne mein error aaya:\n{e}")

def clear_results():
    results_text.config(state=tk.NORMAL)
    results_text.delete(1.0, tk.END)
    results_text.config(state=tk.DISABLED)
    folder_entry.delete(0, tk.END)
    word_entry.delete(0, tk.END)

def export_results():
    content = results_text.get(1.0, 'end-1c') 
    if not content.strip():
        messagebox.showwarning("Empty", "Nothing to Export.")
        return
    try:
        ctypes.windll.ole32.CoInitialize(None) # Fix for Save Dialog too
    except: pass
    
    file_path = filedialog.asksaveasfilename(title="Save results as", defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as f: f.write(content)
            messagebox.showinfo("Success", f"Results saved to {file_path}")
        except Exception: pass

# GUI Setup
window = tk.Tk()
window.title("Portable Word Search Tool")

tk.Label(window, text="Enter word to search:").grid(row=0, column=0, pady=5, padx=10, sticky="w")
word_entry = tk.Entry(window, width=40)
word_entry.grid(row=0, column=1, pady=5, padx=10)

tk.Label(window, text="Select folder to search:").grid(row=1, column=0, pady=5, padx=10, sticky="w")
folder_entry = tk.Entry(window, width=40)
folder_entry.grid(row=1, column=1, pady=5, padx=10)
tk.Button(window, text="Browse Folder", command=browse_folder_or_file).grid(row=1, column=2, pady=5, padx=10)

chk_frame = tk.Frame(window)
chk_frame.grid(row=2, column=0, columnspan=3, pady=5)
tk.Label(chk_frame, text="Search in:").pack(side='left')

var_txt = tk.IntVar(value=0) 
var_pdf = tk.IntVar(value=0)
var_docx = tk.IntVar(value=0)
var_jpg = tk.IntVar(value=0)
var_png = tk.IntVar(value=0)

for text, var in [("TXT", var_txt), ("PDF", var_pdf), ("DOCX", var_docx), ("JPG", var_jpg), ("PNG", var_png)]:
    tk.Checkbutton(chk_frame, text=text, variable=var).pack(side='left')

tk.Button(window, text="Start Search", command=start_search, font=('Arial', 10, 'bold')).grid(row=3, column=1, pady=10)

results_text = scrolledtext.ScrolledText(window, height=15, width=70, wrap=tk.WORD)
results_text.grid(row=4, column=0, columnspan=3, pady=10, padx=10)
results_text.config(state=tk.DISABLED)

bottom_frame = tk.Frame(window)
bottom_frame.grid(row=5, column=0, columnspan=3, pady=(0, 10))
tk.Button(bottom_frame, text="Clear Results", command=clear_results).pack(side='left', padx=10)
tk.Button(bottom_frame, text="Export Results", command=export_results).pack(side='left', padx=10)

window.mainloop()