import os
import shutil
import cv2
import subprocess
import tkinter
import pytesseract
import pdf2image

from PyPDF2 import PdfReader, PdfMerger

def create_folder(folder_name):
    if not os.path.exists(folder_name):
        try:
            os.makedirs(folder_name)
        except Exception as e:
            print("Error creating folder:", e)

def create_data_folders():
    create_folder(not_OCRed_directory)
    create_folder(ocr_data_directory)

program_file = os.path.dirname(os.path.abspath(__file__))
not_OCRed_directory = os.path.join(program_file, "Original Not OCRed")
ocr_data_directory = os.path.join(program_file, "OCR-DATA")

root = tkinter.Tk()
root.title("Folder OCR")
root.geometry("350x150")

def trigger_OCR():
    messagebox = tkinter.Toplevel(root)
    messagebox.title("Working")
    messagebox.geometry("500x100")
    
    message_label = tkinter.Label(messagebox, text="Working on those files, this window will close when the process is finished.\nThere will be several pop ups, just ignore them.")
    message_label.pack(pady=20)
    messagebox.protocol("WM_DELETE_WINDOW", lambda: None)
    create_data_folders()
    user_path = path_entry.get()
    all_files = fetch_all_files(user_path)
    ocred_pdfs = []
    error_pdfs = []
    for i, file in enumerate(all_files, start=1):
        root_directory = os.path.dirname(file)
        file_w_ext = os.path.basename(file).split('/')[-1]
        split_tup = os.path.splitext(file_w_ext)
        file_name = split_tup[0]
        file_extension = split_tup[1]
        if file_extension.lower() == ".pdf":
            try:
                shutil.copy2(file, not_OCRed_directory)
            except Exception as e:
                print(e)
            copied_file = os.path.join(not_OCRed_directory, file_w_ext)
            needs_OCRed = check_OCR(copied_file)
            try:
                if needs_OCRed:
                    OCR_proccess(file, root_directory, file_w_ext, file_name)
                    ocred_pdfs.append(file)
                else:
                    pass
            except Exception as e:
                error = f"{e} - {file}"
                error_pdfs.append(error)
        root.update()
    if len(error_pdfs) > 0:
        show_messages(error_pdfs, message_type="errors")
    if len(ocred_pdfs) > 0:
        show_messages(ocred_pdfs, message_type="OCRed")
    if len(ocred_pdfs) == 0 and len(error_pdfs) == 0:
        show_messages(messages=[], message_type="none")
    root.after(2000, messagebox.destroy)
    root.after(2000, path_entry.delete(0, tkinter.END))

def show_messages(messages, message_type):
    error_message_box = tkinter.Toplevel(root)
    error_message_box.title(message_type.capitalize())

    canvas = tkinter.Canvas(error_message_box, height=500)
    canvas.pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=True)

    scrollbar = tkinter.Scrollbar(error_message_box, command=canvas.yview)
    scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    frame = tkinter.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor="nw")

    
    total_height = 0

    if messages:
        max_width = 500
        title_label = tkinter.Label(frame, text=f"{message_type.capitalize()}:")
        title_label.grid(sticky="w", padx=20, pady=20)
        total_height += title_label.winfo_reqheight() + 20

        for i, file in enumerate(messages):
            message_label = tkinter.Label(frame, text=f"{file}")
            message_label.grid(sticky="w", padx=20, pady=10)

            label_width = message_label.winfo_reqwidth()
            if label_width > max_width:
                max_width = label_width

            label_height = message_label.winfo_reqheight()
            total_height += label_height + 10

    else:
        max_width = 200
        title_label = tkinter.Label(frame, text="No files needed to be OCRed, nothing done.")
        title_label.grid(sticky="w", padx=20, pady=20)
        label_width = title_label.winfo_reqwidth()
        if label_width > max_width:
                max_width = label_width
        total_height = title_label.winfo_reqheight() + 20

    frame.grid_rowconfigure(0, weight=1)

    close_button = tkinter.Button(frame, text="Ok", command=lambda: close_error(error_message_box))
    close_button.grid(row=len(messages) + 1, column=0, pady=20)

    window_height = total_height + close_button.winfo_reqheight()
    window_height = min(window_height, 500)
    error_message_box.geometry(f"{max_width + 40}x{window_height + 50}")

    window_width = error_message_box.winfo_width()
    window_height = error_message_box.winfo_height()
    screen_width = error_message_box.winfo_screenwidth()
    screen_height = error_message_box.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    error_message_box.geometry(f"+{x}+{y}")

def close_error(box):
    box.destroy()
    
def fetch_all_files(folder_save_point):
    target_files = []
    for file_name in os.listdir(folder_save_point):
        if os.path.isfile(os.path.join(folder_save_point, file_name)):
            target_files.append(os.path.join(folder_save_point, file_name))
    return target_files

def OCR_proccess(file_path, directory, file_name_ext, file_name):
    ocr_data_folder = os.path.join(ocr_data_directory, file_name)
    create_folder(ocr_data_folder)
    image_conversion(file_path, ocr_data_folder, file_name)
    jpgs = [f for f in os.listdir(ocr_data_folder) if os.path.isfile(os.path.join(ocr_data_folder, f))]
    for page, pic in enumerate(jpgs):
        page += 1
        OCR(page, pic, ocr_data_folder)
    extracted_files = fetch_all_files(ocr_data_folder)
    merge_pdf(extracted_files, file_name, directory)

def image_conversion(inpath, folder_save_point, file_name):
        OUTPUT_FOLDER = folder_save_point
        FIRST_PAGE = None
        LAST_PAGE = None
        FORMAT = 'jpg'
        USERPWD = None
        USE_CROPBOX = False
        STRICT = False
        page = 0
        pdf2image.convert_from_path(inpath,
                                    output_folder = OUTPUT_FOLDER,
                                    first_page = FIRST_PAGE,
                                    last_page = LAST_PAGE,
                                    fmt = FORMAT,
                                    userpw = USERPWD,
                                    use_cropbox = USE_CROPBOX,
                                    strict = STRICT,
                                    dpi=300, output_file = f"{file_name}"
                                    )

def get_tesseract_path():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    tesseract_exe = os.path.join(current_dir, "tesseract.exe")
    return tesseract_exe

def OCR(pages, pic, ocr_data_folder):
        pytesseract.pytesseract.tesseract_cmd = get_tesseract_path()
        input_dir = os.path.join(ocr_data_folder, pic)
        img = cv2.imread(input_dir, 1)
        img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        result = pytesseract.image_to_pdf_or_hocr(img_rgb, lang = "eng",
                                                  config = "--dpi 300")
        input_dir = input_dir.replace('.jpg', '.pdf')
        f = open(input_dir, "w+b")
        f.write(bytearray(result))
        f.close()

def merge_pdf(extracted_files: list [str], file_name, base_directory):
        output_PDF = file_name + ".pdf"
        merger = PdfMerger()
        pdfs = 0
        for pdf in extracted_files:
            if pdf.endswith('.pdf'):
                merger.append(pdf)
                pdfs += 1
        merger.write(output_PDF)
        merger.close()
        compress_pdf(output_PDF)
        shutil.copy2(output_PDF, base_directory)
        os.remove(output_PDF)

def get_pdfsizeopt_path():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    pdfsize_exe = os.path.join(current_dir, "pdfsizeopt.exe")
    return pdfsize_exe

def compress_pdf(pdf_path):
    try:
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.run(
            [get_pdfsizeopt_path(), '--use-pngout=no', pdf_path, pdf_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW,
            startupinfo=si
        )
        
    except subprocess.CalledProcessError as e:
        print(f"An error occurred during compression: {e}")

def check_OCR(file_path):
    try:
        with open(file_path, "rb") as pdf_file:
            pdf_reader = PdfReader(pdf_file)
            for page_num, page in enumerate(pdf_reader.pages, start=1):
                if page.extract_text().strip():
                    return False
        return True
    except PermissionError:
        print("Permission Error, unable to open the file")
        return False
    except Exception as e:
        print(f"{e}")
        return False

def clear_entry(event):
    path_entry.delete(0, tkinter.END)
    path_entry["fg"] = "black"

frame = tkinter.Frame(master=root)
frame.pack(pady=10, padx=20, fill="both", expand=True)

label = tkinter.Label(master=frame, text="OCR all PDFs in a folder")
label.grid(column=0, columnspan=2, row=0)

example_label = tkinter.Label(master=frame, text="Enter the Folder path below:", justify="left")
example_label.grid(column=0, columnspan=2, row=1)

path_entry = tkinter.Entry(master=frame, width=50, fg="gray")
path_entry.insert(string=r"C:\Users\dee.HFMLEGAL\hfmlegal.com\hfmlegal.com - Documents\General", index=0)
path_entry.grid(column=0, columnspan=2, row=2, sticky="w", pady=(10, 5))
path_entry.bind("<FocusIn>", clear_entry)

OCR_button = tkinter.Button(master=frame, text="OCR", command=trigger_OCR)
OCR_button.grid(column=0, row=3, pady=10, padx=10, sticky="ew")

root.mainloop()