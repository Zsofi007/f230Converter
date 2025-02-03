import json
import os
import sys
import threading
import tkinter as tk
import xml.etree.ElementTree as ET
import zipfile
from collections import defaultdict
from datetime import datetime
from io import BytesIO
from tkinter import *
from tkinter import filedialog, messagebox

import pdfplumber
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
from ttkbootstrap.tooltip import ToolTip

adresa_arnis = "strada CAISULUI nr. 45 bl. CORP C2 sc. 2 et. 2 apt. 17 loc. FUNDENI jud.  IF--Ilfov cod postal 077085"
den_arnis = "ASOCIATIA ROMANA PENTRU NOU-NASCUTII INDELUNG SPITALIZATI - ARNIS"
den_i = "NAN RUXANDRA-NICOLETA"
cif_arnis = "2871109134160"
cui_arnis = "32224197"
telefon_i = "0721664814"
email_i="RUXANDRA.NAN@ARNIS.ONG"

iban_to_family = {
    "RO88INGB0000999906692521": "MATHE",
    "RO29INGB0000999903935097": "ARNIS",
}

max_declarations_per_xml = 100
extracted_data = {}

def get_resource_path(relative_path):
    """ Get the absolute path to a resource, works for dev and for PyInstaller """
    if getattr(sys, 'frozen', False):  # Running as an executable
        base_path = sys._MEIPASS
    else:  # Running in a normal Python environment
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

coords_path = get_resource_path('coordinates.json')

# Load coordinates from the JSON file
with open(coords_path, 'r') as f:
    coordinates = json.load(f)

def pixels_to_points(pixels, dpi=144):
    return pixels * (72 / dpi)

def new_flow():
    remove_file_entries(gui_elements)
    hide_download_gui()
    hide_generate_gui()
    

# --- XML Generation ---

def generate_xml(gui_elements):
    value = gui_elements['document_number_entry'].get().strip()
    if not value:  # If the entry is empty, default to 1
        document_number = 1
    elif value.isdigit():  # If it's a valid number
        document_number = int(value)
    else:  # If it's not a valid number, show an error message
        messagebox.showerror("Invalid Input", f"'{value}' is not a number. Please enter a valid number.")
        return None  # Exit the function without returning a number

    pdf_paths = gui_elements.get('pdf_paths', [])
    if not pdf_paths:
        messagebox.showwarning("Warning", "No PDFs selected.")
        return
    
    extracted_data_list = [pdf['data'] for pdf in pdf_paths]
    grouped_data = defaultdict(list)
    
    for data in extracted_data_list:
        iban = data.get("iban", "UNKNOWN_IBAN")
        grouped_data[iban].append(data)
    
    xml_files = []
    for iban, data_list in grouped_data.items():
        family_name = iban_to_family.get(iban, iban)

        # Split into chunks of max_declarations_per_xml
        for i in range(0, len(data_list), max_declarations_per_xml):
            chunk = data_list[i:i + max_declarations_per_xml]
            file_index = (i // max_declarations_per_xml) + 1
            xml_tree = create_xml_structure(chunk, iban, document_number)
            document_number += 1
            xml_filename = f"b230_{family_name}_{file_index}.xml"
            xml_files.append({'name': xml_filename, 'xml_tree': xml_tree})
    
    gui_elements['xml_files'] = xml_files
    messagebox.showinfo("Success", "PDFs processed successfully. XML files created.")
    show_download_gui()


    
def format_address(data):
    address_parts = [
        f"strada {data.get('strada', '')}",
        f"nr. {data.get('nr', '')}",
        f"bl. {data.get('bloc', '')}" if data.get('bloc') else "",
        f"sc. {data.get('scara', '')}" if data.get('scara') else "",
        f"et. {data.get('etaj', '')}" if data.get('etaj') else "",
        f"apt. {data.get('ap', '')}" if data.get('ap') else "",
        f"loc. {data.get('localitate', '')}",
        f"jud. {data.get('judet', '')}",
        f"cod postal {data.get('cod_postal', '')}"
    ]
    return " ".join(filter(None, address_parts))

def generate_borderou_element(data, document_number):
    data_borderou=datetime.now().strftime("%d.%m.%Y")
    return ET.Element("borderou230", nr_borderou=f"{document_number}", data_borderou=data_borderou, luna="12", an=data[0].get("an", "").replace(" ", ""),
                         totalPlata_A=str(len(data)),
                         den=den_arnis,
                         cui=cui_arnis, den_i=den_i,
                         cif_i=cif_arnis,
                         adresa_i=adresa_arnis,
                         telefon_i=telefon_i,
                         email_i=email_i,
                         xmlns="mfp:anaf:dgti:b230:declaratie:v1",
                         attrib={"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
                                 "xsi:schemaLocation": "mfp:anaf:dgti:b230:declaratie:v1 B230.xsd"})

def generate_declaration_element(root, data, index):
    address = format_address(data)
    attributes = {
        "nume_c": data.get("nume_c", ""),
        "prenume_c": data.get("prenume_c", ""),
        "adresa_c": address,
        "telefon_c": data.get("telefon_c", ""),
        "email_c": data.get("email_c", ""),
        "cif_c": data.get("cif_c", "").replace(" ", ""),
        "nr_poz": str(index)
    }
    if data.get("initiala_c"):
        attributes["initiala_c"] = data["initiala_c"]
    
    return ET.SubElement(root, "declaratie230", **attributes)

def generate_bursa_entit_element(declaratie, extracted_data, iban):
    return ET.SubElement(declaratie, "bursa_entit", 
                         bifa_entitate="1" if extracted_data.get("bifa_entitate", "") == "X" else "0",
                         den_entitate=extracted_data.get("den_entitate", ""),
                         cif_entitate=extracted_data.get("cif_entitate", ""), 
                         cont_entitate=iban,
                         procent=extracted_data.get("procent", "").replace(",", ".").replace("%", ""), 
                         valabilitate_distribuire="2" if extracted_data.get("doi_ani", "") == "X" else "1")


def create_xml_structure(extracted_data_list, iban, document_number):
    root = generate_borderou_element(extracted_data_list, document_number)
    for index, extracted_data in enumerate(extracted_data_list, start=1):
        declaratie = generate_declaration_element(root, extracted_data, index)
        generate_bursa_entit_element(declaratie, extracted_data, iban)
    return ET.ElementTree(root)

# def save_xml(tree, filename):
#     with open(filename, "wb") as f:
#         f.write(b'<?xml version="1.0"?>\n')
#         tree.write(f, encoding="utf-8")

# --- PDF Processing ---

def count_non_mac_files(zip_path):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # Filter files that do not start with '__MACOSX'
        non_macos_files = [file_name for file_name in zip_ref.namelist() if not file_name.startswith('__MACOSX')]
        return len(non_macos_files)

def extract_pdf_data_from_zip(zip_path):
    # Start the progress bar in a separate thread
    gui_elements['processing_label'].pack(fill='x', padx=10, pady=10)
    gui_elements['process_pdf_progress'].pack(fill='x', padx=10, pady=10)
    nr_of_files = count_non_mac_files(zip_path)
    disable_select_button()
    def worker():
        extracted_data = {}
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # Iterate over the PDF files in the zip archive
            for file_name in zip_ref.namelist():
                if file_name.lower().endswith('.pdf') and not file_name.startswith('__MACOSX'):
                    name_to_add = file_name
                    gui_elements['file_list'].after(0, lambda name=name_to_add: add_file_entry(gui_elements, name))
                    gui_elements['process_pdf_progress'].after(0, lambda step_lengt=100/nr_of_files: gui_elements['process_pdf_progress'].step(step_lengt))   
                    with zip_ref.open(file_name) as file:
                        # Open the PDF in memory
                        with pdfplumber.open(BytesIO(file.read())) as pdf:
                            page = pdf.pages[0]  # assuming the relevant data is on the first page
                            
                            # Extract the data based on coordinates for each field
                            file_data = {}
                            for field, coords in coordinates.items():
                                x = pixels_to_points(coords['x'])
                                y = pixels_to_points(coords['y'])
                                w = pixels_to_points(coords['width'])
                                h = pixels_to_points(coords['height'])

                                # Extract text from the bounding box (x, y, x + w, y + h)
                                cropped_region = page.within_bbox((x, y, x + w, y + h))
                                text = cropped_region.extract_text()
                                file_data[field] = text.strip() if text else ""
                            
                            gui_elements['pdf_paths'].append({
                                'name': file_name,
                                'data': file_data  # Store the extracted data
                            })
                            extracted_data[file_name] = file_data     
        gui_elements['process_pdf_progress'].after(0, lambda: gui_elements['process_pdf_progress'].pack_forget())
        gui_elements['processing_label'].after(0, lambda: gui_elements['processing_label'].pack_forget())
        show_generate_gui()
        enable_select_button()
    threading.Thread(target=worker, daemon=True).start()

def select_pdfs(gui_elements):
    new_flow()
    zip_path = filedialog.askopenfilename(filetypes=[("ZIP Files", "*.zip")])
    if not zip_path:
        return
    
    max_size = 1 * 1024 * 1024 * 1024  # 1 GB
    file_size = os.path.getsize(zip_path)

    if file_size > max_size:
        messagebox.showwarning("Warning", "File size exceeds 1 GB.")
        return
    
    extract_pdf_data_from_zip(zip_path)
    gui_elements['file_list'].yview(tk.END)  # Scroll to the bottom

# --- ZIP Download ---

def save_xml(tree):
    memory_file = BytesIO()
    tree.write(memory_file, encoding="utf-8", xml_declaration=True)
    memory_file.seek(0)
    return memory_file

def download_zip():
    zip_filename = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
    if zip_filename:
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for xml_file in gui_elements.get('xml_files', []):
                memory_file = save_xml(xml_file['xml_tree'])
                zipf.writestr(xml_file['name'], memory_file.read())
        messagebox.showinfo("Success", "ZIP file saved successfully.")

# --- Quit App ---

def quit_app():
    root.quit()


# --- GUI Setup ---

def set_root_geometry():
    app_width = 1000
    app_height = 400
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.geometry(f"{app_width}x{app_height}+{screen_width//2 - app_width//2}+{screen_height//2 - app_height//2}")
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=2)
    root.grid_rowconfigure(0, weight=1)  # Main content resizes
    root.grid_rowconfigure(1, weight=0)  # Bottom row for quit button


def add_file_entry(gui_elements, file_name):
    """Dynamically add a file entry with name and icon."""
    file_frame = ttkb.Frame(gui_elements['file_list'],padding=5, style="custom.TFrame")
    file_frame.pack(fill='x', padx=5, pady=0.5)

    # File name label
    name_label = ttkb.Label(file_frame, text=file_name, padding=5,bootstyle=("dark", "inverse"))
    name_label.pack(side="left", fill="x", expand=True)

    # Icon button (Placeholder action)
    icon_button = ttkb.Button(file_frame, text="X", bootstyle="outline", width=3, command=lambda: print(f"Action on {file_name}"))
    icon_button.pack(side="right", padx=5)

    gui_elements['file_entries'].append(file_frame)  # Store references

def remove_file_entries(gui_elements):
    gui_elements['pdf_paths'] = []
    for file_entry in gui_elements['file_entries']:
        file_entry.destroy()
    gui_elements['file_entries'] = []

def remove_file_entry(gui_elements, file_name):
    print("Removing file entry")


def create_gui_elements():
    gui_elements = {
        'pdf_paths': [],
        'xml_files': [],
        'file_entries': []
    }
    # Left panel (33% width)
    select_button = ttkb.Button(frame_left, text="Import zipped PDFs", bootstyle=PRIMARY, command=lambda: select_pdfs(gui_elements))
    select_button.pack(fill='x', padx=10, pady=15)
    ToolTip(select_button, text="File size should NOT exceed 1 GB.")

    #generate
    document_number_label = ttkb.Label(frame_left, text="Enter the number of the first document.")
    document_number_entry = ttkb.Entry(frame_left, placeholder="Document Number")
    document_number_after_label = ttkb.Label(frame_left, text="e.g. if 5 files are generated and the first\ndocument number is 100, the documents \nwill be numbered 100, 101, 102, 103, 104.\nIf left empty, the default value is 1.")
    generate_button = ttkb.Button(frame_left, text="Generate XMLs", bootstyle=PRIMARY, command=lambda: generate_xml(gui_elements))

    #download
    download_button = ttkb.Button(frame_left, text="Download ZIP", bootstyle=SUCCESS, command=download_zip)

    # Right panel (67% width) - File List

    file_list_frame = ScrolledFrame(frame_right, style="custom.TFrame", padding=10, autohide=True)
    file_list_frame.pack(fill=BOTH, expand=YES, padx=10, pady=10)

    processing_label = ttkb.Label(frame_left, text="Processing PDFs...")

    process_pdf_progress = ttkb.Progressbar(frame_left, orient="horizontal", maximum=100)

    gui_elements['file_list'] = file_list_frame
    gui_elements['select_button'] = select_button
    gui_elements['document_number_label'] = document_number_label
    gui_elements['document_number_entry'] = document_number_entry
    gui_elements['document_number_after_label'] = document_number_after_label
    gui_elements['generate_button'] = generate_button
    gui_elements['download_button'] = download_button
    gui_elements['processing_label'] = processing_label
    gui_elements['process_pdf_progress'] = process_pdf_progress
    gui_elements['pdf_paths'] = []
    gui_elements['xml_files'] = []

    # Quit button
    quit_button = ttkb.Button(frame_right, text="Quit", bootstyle="danger", command=root.quit)
    quit_button.pack(side="bottom", anchor="se", padx=10, pady=10)

    return gui_elements

def enable_generate_button():
    gui_elements['generate_button'].config(state="normal")

def disable_generate_button():
    gui_elements['generate_button'].config(state="disabled")

def enable_select_button():
    gui_elements['select_button'].config(state="normal")

def disable_select_button():
    gui_elements['select_button'].config(state="disabled")

def show_generate_gui():
    gui_elements['document_number_label'].pack(fill='x', padx=10)
    gui_elements['document_number_entry'].pack(anchor="w",padx=10, pady=3)
    gui_elements['document_number_after_label'].pack(fill='x', padx=10)
    gui_elements['generate_button'].pack(fill='x', padx=10, pady=5)

def show_download_gui():
    gui_elements['download_button'].pack(fill='x', padx=10, pady=5)

def hide_generate_gui():
    gui_elements['document_number_label'].pack_forget()
    gui_elements['document_number_entry'].pack_forget()
    gui_elements['document_number_after_label'].pack_forget()
    gui_elements['generate_button'].pack_forget()

def hide_download_gui():
    gui_elements['download_button'].pack_forget()



# Initialize root window
root = ttkb.Window(themename="superhero")
root.title("PDF to XML-B230 Converter ARNIS")

logo_path = get_resource_path('logo-ARNIS.png')
icon = PhotoImage(file=logo_path)
root.iconphoto(False, icon)

style = ttkb.Style(theme="superhero")
style.configure("custom.TFrame",background="#20374C", foreground="#ffffff")
# Create main frames (parent container)
frame_left = ttkb.Frame(root, padding=10)
frame_left.grid(row=0, column=0, sticky='nsew')

frame_right = ttkb.Frame(root, padding=10)
frame_right.grid(row=0, column=1, sticky='nsew')

# Create GUI elements
gui_elements = create_gui_elements()

# Set window geometry
set_root_geometry()

# Start the Tkinter main loop
root.mainloop()