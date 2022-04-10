from pdf2docx import Converter  # install it using pip install pdf2docx or else code will not work
import PyPDF2                   # install it using pip install PyPDF2 or else code will not work
import fitz
from tkinter import *
from tkinter import filedialog
import time


# ---------------- WINDOW -------------------

window = Tk()
window.title('PDF Converter')
window.iconbitmap('pdf.ico')
window.geometry('500x500')
window.resizable(False, False)
window.config(bg='cyan')

# ---------------- WINDOW -------------------


# ---------------- SELECT FILE -------------------

def open_file():
    open_file.filename = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
    upload_file.delete(0, END)
    upload_file.insert(0, open_file.filename)
    if saved_file.get() != '':
        saved_file.delete(0, END)
    else:
        pass

# ---------------- SELECT FILE -------------------


# ---------------- SAVE FILE -------------------

def save_file():
    txt.delete('1.0', END)
    try:
        if choice.get() == 'PDF2Docx':  # If pdf to word is selected
            save_file.foldername = filedialog.asksaveasfilename(defaultextension='.docx',
                                                                filetypes=[('Word file', '*.docx')])
        elif choice.get() == 'PDF2Txt':  # If pdf to text is selected
            save_file.foldername = filedialog.asksaveasfilename(defaultextension='.txt',
                                                                filetypes=[('Text file', '*.txt')])
        elif choice.get() == 'PDF2Image':  # If pdf to image is selected
            save_file.foldername = filedialog.asksaveasfilename(defaultextension='.jpg',
                                                                filetypes=[('JPEG', '*.jpg'),  # For JPEG
                                                                           ('PNG', '*.png')])  # For PNG
            path_open = r"%s" % upload_file.get()
            with fitz.open(path_open) as doc:
                x = 0
                for page in doc:
                    x += 1

                save_file.page_nums_int = []
                for number in range(1, x+1):
                    save_file.page_nums_int.append(number)

            page_nums_str = [str(x) for x in save_file.page_nums_int]
            save_file.num_choice = StringVar()
            save_file.num_choice.set('Select Page number')
            save_file.drop_num = OptionMenu(window, save_file.num_choice, *page_nums_str)   # Widget for selecting page number
            save_file.drop_num.place(x=210, y=120)                                          # Only applicable for PDf to image conversion

        saved_file.delete(0, END)
        saved_file.insert(0, save_file.foldername)
    except TypeError:
        txt.delete('1.0', END)
        txt.insert('1.0', '#######   Select a PDF File First!!!!!!   ########')           # Error handling if no pdf file is selected
    except AttributeError:
        txt.delete('1.0', END)
        txt.insert('1.0', '#######   Select a conversion first!!!!!!   ########')         # Error handling if any conversion isn't selected

# ---------------- SAVE FILE -------------------


# ---------------- CONVERTING PROCESS -------------------

def progress():
    try:
        # -------- Loading Screen -------------
        txt.delete('1.0', END)

        txt.insert('1.0', f'{choice.get()} is Selected')
        window.update_idletasks()
        time.sleep(2)

        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverting')

        window.update_idletasks()
        time.sleep(1)
        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverting.')

        window.update_idletasks()
        time.sleep(1)
        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverting..')

        window.update_idletasks()
        time.sleep(1)
        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverted\nSaving')

        window.update_idletasks()
        time.sleep(1)
        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverted\nSaving.')

        window.update_idletasks()
        time.sleep(1)
        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverted\nSaving..')

        window.update_idletasks()
        time.sleep(1)
        txt.delete('1.0', END)
        txt.insert('1.0', f'{choice.get()} is Selected\nConverted\nSaved\n\nProcess Completed Successfully')
        # --------------- Loading Screen -------------------

        if choice.get() == 'PDF2Docx':  # PDF to doc Conversion
            pdf_file = open_file.filename
            docx_file = save_file.foldername
            cv = Converter(pdf_file)
            cv.convert(docx_file)      # all pages by default
            cv.close()
        elif choice.get() == 'PDF2Txt':  # PDF to text conversion
            path_open = r"%s" % open_file.filename
            pdffileobj = open(path_open, 'rb')
            pdfreader = PyPDF2.PdfFileReader(pdffileobj)
            x = pdfreader.numPages
            pageobj = pdfreader.getPage(x-1)
            text = pageobj.extractText()

            path_save = r"%s" % save_file.foldername
            file1 = open(path_save, 'a')
            file1.writelines(text)
            file1.close()
        elif choice.get() == 'PDF2Image':  # PDF to Image conversion
            path_open = r"%s" % open_file.filename
            path_save = r"%s" % save_file.foldername

            pdf_file = path_open
            doc = fitz.open(pdf_file)
            page = doc.load_page(int(save_file.num_choice.get()) - 1)  # Page number to convert (Only one page converted at a time)
            pix = page.get_pixmap()
            output = path_save
            pix.save(output)

            save_file.drop_num.destroy()  # Remove the select page number widget
    except ValueError:
        txt.delete('1.0', END)
        txt.insert('1.0', 'Failed :/\nPlease select which page to convert')  # Error handling if no page is selected
    except Exception as es:
        print(es)
        txt.delete('1.0', END)
        er = f'Failed\n---FATAL ERROR---\n\n{es}'
        txt.insert('1.0', er)

# ---------------- CONVERTING PROCESS -------------------


# ---------------- WIDGETS -------------------

Label(window, text='PDF Converter', font=('Algerian', 30), bg='cyan', fg='black').pack()

Label(window, text="Select pdf file").place(x=20, y=80)
upload_file = Entry(window, width=50,)
upload_file.place(x=100, y=80)
Button(window, text='Browse', command=open_file).place(x=415, y=79)

conversions = [
    'PDF2Docx',
    'PDF2Txt',
    'PDF2Image'
]
choice = StringVar()
choice.set('Select Conversion')
drop = OptionMenu(window, choice, *conversions)
drop.place(x=100, y=120)


Label(window, text="Save File in").place(x=20, y=180)
saved_file = Entry(window, width=50)
saved_file.place(x=100, y=180)
Button(window, text='Browse', command=save_file).place(x=415, y=179)

txt = Text(window, bg='black', fg='white', height=10, width=59)
txt.place(x=10, y=250)
Button(window, text="Convert", width=15, height=4, font=(None, 10, 'bold'), command=progress).place(x=180, y=420)

# ---------------- WIDGETS -------------------

window.mainloop()
