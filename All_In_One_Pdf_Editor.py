from tkinter import *
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
from tkinter.filedialog import askopenfile, askopenfilename, askopenfilenames, asksaveasfilename
from tkinter.messagebox import showinfo, showerror
from PIL import Image, ImageTk
from pdf2image import convert_from_path
from docx2pdf import convert
import img2pdf
import os

root = Tk()
root.title("All in one PDF editor")
root.geometry("900x558")
root.resizable(0, 0)
root.configure(bg='#A7C7E7')

photo = PhotoImage(file = 'lcons\icon.png')
root.iconphoto(False, photo)

Label(root, text="ALL IN ONE PDF EDITOR", fg="NAVY BLUE", font="Georgia 20 bold", bg='#A7C7E7').place(x=260, y=7)


# Pdf Viewer function
def view_pdf():
    try:
        filename = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))
        pages = convert_from_path(filename, poppler_path=r'C:\Users\anike\Poppler\poppler-0.68.0_x86\poppler-0.68.0\bin', size=(800, 900))

        newWindow = Toplevel()
        newWindow.title("Pdf Viewer")
        newWindow.geometry("810x700")

        scrol_y = Scrollbar(newWindow, orient=VERTICAL)

        pdf = Text(newWindow, yscrollcommand=scrol_y.set, bg="grey")

        scrol_y.pack(side=RIGHT, fill=Y)
        scrol_y.config(command=pdf.yview)

        pdf.pack(fill=BOTH, expand=1)

        photos = []

        for i in range(len(pages)):
            photos.append(ImageTk.PhotoImage(pages[i]))

        for photo in photos:
            pdf.image_create(END, image=photo)
            pdf.insert(END, '\n\n')
        newWindow.mainloop()
    except:
        showerror("Warning", "No pdf selected")

image = Image.open('lcons\icon1.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img = ImageTk.PhotoImage(image)

Button(root, text="\nVIEW PDF", font="Helvetica 12 bold", width=120, height=110, command=view_pdf, bg='silver', image=img, compound=TOP).place(x=130, y=48)


# Pdf to image
def pdf_to_img():
    try:
        filename = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))
        images = convert_from_path(filename, poppler_path=r'C:\Users\anike\Poppler\poppler-0.68.0_x86\poppler-0.68.0\bin', size=(800, 900))
        
        i = 0
        file = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor/new_folder", initialfile = "pdf2img",title = "name file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
        
        for img in images:
            img.save(f'{file}{i}.jpg', 'JPEG')
            i = i + 1
    except:
        showerror("Warning", "No pdf selected")

image = Image.open('lcons\icon2.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img2 = ImageTk.PhotoImage(image)

Button(root, text="\nPDF TO IMAGE", font="Helvetica 12 bold", width=120, height=110, command=pdf_to_img, bg='silver', image=img2, compound=TOP).place(x=380, y=48)


# IMAGES TO PDF
def image_to_pdf():
    try:
        file_names = askopenfilenames(initialdir="/", title="Select File")
        file = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor",title = "name file",filetypes = (("pdf files","*.pdf"),("all files","*.*")))
        with open(f"{file}.pdf", "wb") as f:
            f.write(img2pdf.convert(file_names))

        pages = convert_from_path(f"{file}.pdf", poppler_path=r'C:\Users\anike\Poppler\poppler-0.68.0_x86\poppler-0.68.0\bin', size=(800, 900))

        newWindow = Toplevel()
        newWindow.title("Image Pdf Viewer")
        newWindow.geometry("810x700")

        scrol_y = Scrollbar(newWindow, orient=VERTICAL)

        pdf = Text(newWindow, yscrollcommand=scrol_y.set, bg="grey")

        scrol_y.pack(side=RIGHT, fill=Y)
        scrol_y.config(command=pdf.yview)

        pdf.pack(fill=BOTH, expand=1)

        photos = []

        for i in range(len(pages)):
            photos.append(ImageTk.PhotoImage(pages[i]))

        for photo in photos:
            pdf.image_create(END, image=photo)
            pdf.insert(END, '\n\n')
        newWindow.mainloop()
    except:
        showerror("Warning", "No Images selected")

image = Image.open('lcons\icon3.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img3 = ImageTk.PhotoImage(image)

Button(root, text="\nIMAGE TO PDF", font="Helvetica 12 bold", width=120, height=110, command=image_to_pdf, bg='silver', image=img3, compound=TOP).place(x=630, y=48)


# Pdf merger function for 2 pdf
def pdf_merger():
    try:
        filename1 = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))
        pdf_obj = open(filename1, "rb")
        pdf_reader1 = PdfFileReader(pdf_obj)

        filename2 = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))
        pdf_obj = open(filename2, "rb")
        pdf_reader2 = PdfFileReader(pdf_obj)

        output = PdfFileMerger()

        output.append(pdf_reader1)
        output.append(pdf_reader2)

        file = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor",title = "name file",filetypes = (("pdf files","*.pdf"),("all files","*.*")))

        output.write(f"{file}.pdf")

        newWindow = Toplevel()
        newWindow.title("Merged pdf viewer")
        newWindow.geometry("810x700")

        scrol_y = Scrollbar(newWindow, orient=VERTICAL)

        pdf = Text(newWindow, yscrollcommand=scrol_y.set, bg="grey")

        scrol_y.pack(side=RIGHT, fill=Y)
        scrol_y.config(command=pdf.yview)

        pdf.pack(fill=BOTH, expand=1)

        pages = convert_from_path(f"{file}.pdf", poppler_path=r'C:\Users\anike\Poppler\poppler-0.68.0_x86\poppler-0.68.0\bin', size=(800, 900))

        photos = []

        for i in range(len(pages)):
            photos.append(ImageTk.PhotoImage(pages[i]))

        for photo in photos:
            pdf.image_create(END, image=photo)
            pdf.insert(END, '\n\n')
        newWindow.mainloop()
    except:
        showerror("Warning", "No pdf selected")

image = Image.open('lcons\icon4.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img4 = ImageTk.PhotoImage(image)

Button(root, text="\nMERGE PDF", font="Helvetica 12 bold", width=120, height=110, command=pdf_merger, bg='silver', image=img4, compound=TOP).place(x=130, y=174)


# password protection
def password():
    newWindow = Toplevel()
    newWindow.title("Pdf Viewer")
    newWindow.geometry("500x400")

    instructions = Label(newWindow, text="Enter a password and select a pdf\n")
    instructions.pack(anchor='center', pady=5)

    password = Entry(newWindow, show="*", width=15)
    password.pack(anchor='center', pady=5)

    def openfile():
        try:
            pdf_file = askopenfile(parent=newWindow, mode="rb", title="choose a file", filetypes=[("PDF Files", " *.pdf")])
            FileName = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor",title = "name file",filetypes = (("pdf files","*.pdf"),("all files","*.*")))
            if pdf_file is not None:
                pdf_reader = PdfFileReader(pdf_file)
                pdf_writer = PdfFileWriter()
                for page_num in range(pdf_reader.numPages):
                    pdf_writer.addPage(pdf_reader.getPage(page_num))
                pdf_writer.encrypt(password.get())
                PasswordFile = f"{FileName}.pdf"
                result_pdf = open(PasswordFile, 'wb')

                pdf_writer.write(result_pdf)
                result_pdf.close()
                password.delete(0, 'end')

                showinfo("Success", "File is successfully Password Protected")
            else:
                showerror("Failed", "Unable to Password Protect the file")
        except:
            showerror("Warning", "No pdf selected")

    browse_btn = Button(newWindow, command=openfile, text="Browse file", width=9)
    browse_btn.pack(anchor='center', pady=5)
    newWindow.mainloop()

image = Image.open('lcons\icon5.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img5 = ImageTk.PhotoImage(image)

Button(root, text="PASSWORD\nPROTECTION", font="Helvetica 12 bold", width=120, height=110, command=password, bg='silver', image=img5, compound=TOP).place(x=380, y=174)


# pdf to text function
def pdf_to_text():
    try:
        filename = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))
        pdf_obj = open(filename, "rb")
        pdf_reader = PdfFileReader(pdf_obj)

        pageObj = pdf_reader.getPage(0)
        pageText = pageObj.extractText()

        newWindow = Toplevel()
        newWindow.title("Pdf Viewer")
        newWindow.geometry("500x400")
        T = Text(newWindow, height=25, width=45)
        T.pack(pady=5)
        T.insert(INSERT, pageText)
        newWindow.mainloop()
    except:
        showerror("Warning", "No pdf selected")

image = Image.open('lcons\icon6.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img6 = ImageTk.PhotoImage(image)

Button(root, text="\nPDF TO TEXT", font="Helvetica 12 bold", width=120, height=110, command=pdf_to_text, bg='silver', image=img6, compound=TOP).place(x=630, y=174)


# rotate pdf
def rotate_selection():
    rotate_selection.num = radio.get()

def rotate_pdf():
    try:
        filename = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))

        file = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor",title = "name file",filetypes = (("pdf files","*.pdf"),("all files","*.*")))
        new_File_Name = f"{file}.pdf"

        rotate_selection()
        rotation_1 = rotate_selection.num

        pdf_File_Object = open(filename, 'rb')

        pdf_Reader = PdfFileReader(pdf_File_Object)

        pdf_Writer = PdfFileWriter()

        for page in range(pdf_Reader.numPages):
            page_Object = pdf_Reader.getPage(page)
            page_Object.rotateClockwise(rotation_1)
            pdf_Writer.addPage(page_Object)

        new_File = open(new_File_Name, 'wb')

        pdf_Writer.write(new_File)

        pdf_File_Object.close()

        new_File.close()
        newWindow = Toplevel()
        newWindow.title("Rotated Pdf Viewer")
        newWindow.geometry("935x600")

        scrol_y = Scrollbar(newWindow, orient=VERTICAL)

        pdf = Text(newWindow, yscrollcommand=scrol_y.set, bg="grey")

        scrol_y.pack(side=RIGHT, fill=Y)
        scrol_y.config(command=pdf.yview)

        pdf.pack(fill=BOTH, expand=1)

        photos = []

        pages = convert_from_path(new_File_Name, poppler_path=r'C:\Users\anike\Poppler\poppler-0.68.0_x86\poppler-0.68.0\bin', size=(800, 900))

        for i in range(len(pages)):
            photos.append(ImageTk.PhotoImage(pages[i]))

        for photo in photos:
            pdf.image_create(END, image=photo)
            pdf.insert(END, '\n\n')
        newWindow.mainloop()
    except:
        showerror("Warning", "No pdf selected")

radio = IntVar()

def new_window_for_rotate():
    Window = Toplevel()
    Window.title("Pdf Viewer")
    Window.geometry("600x500")

    R1 = Radiobutton(Window, text="90", variable=radio, value=90, command=rotate_selection)
    R1.pack(anchor=W)

    R2 = Radiobutton(Window, text="180", variable=radio, value=180, command=rotate_selection)
    R2.pack(anchor=W)

    R3 = Radiobutton(Window, text="270", variable=radio, value=270, command=rotate_selection)
    R3.pack(anchor=W)

    Button(Window, text="Select PDF", width=9, command=rotate_pdf).pack(anchor='center', pady=5)

    Window.mainloop()

image = Image.open('lcons\icon7.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img7 = ImageTk.PhotoImage(image)

Button(root, text="\nROTATE PDF", font="Helvetica 12 bold", width=120, height=110,
       command=new_window_for_rotate, bg='silver', image=img7, compound=TOP).place(x=130, y=302)


# split pdf
def split_pdf():
    try:
        path = askopenfile(parent=root, mode="rb", title="choose a file", filetypes=[
            ("PDF Files", " *.pdf")])
        pdf = PdfFileReader(path)
        
        file = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor/split", title = "name file", filetypes = (("PDF Files","*.pdf"), ("all files","*.*")))
    
        for page in range(pdf.getNumPages()):
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf.getPage(page))
            
            output = f'{file}{page}.pdf'
            with open(output, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)
        
    except:
        showerror("Warning", "No pdf selected")

image = Image.open('lcons\icon8.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img8 = ImageTk.PhotoImage(image)

Button(root, text="\nSPLIT PDF", font="Helvetica 12 bold", width=120, height=110, command=split_pdf, bg='silver', image=img8, compound=TOP).place(x=380, y=302)


# Pdf Info
def pdf_info():
    try:
        pdf_path = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Text files", "*.pdf*"), ("all files", "*.*")))

        with open(pdf_path, 'rb') as f:
            pdf = PdfFileReader(f)
            information = pdf.getDocumentInfo()
            number_of_pages = pdf.getNumPages()

        file_size = (os.path.getsize(pdf_path))/1024

        newWindow = Toplevel()
        newWindow.title("PDF Info Viewer")
        newWindow.geometry("600x400")
        txt = f"""
        Information about : {pdf_path}

        PDF Size: {'%0.2f' %file_size} KB
        Author Name: {information.author}
        Creator: {information.creator}
        Producer: {information.producer}
        Subject: {information.subject}
        Title: {information.title}
        Number of pages: {number_of_pages}
        """

        text = Text(newWindow, width=80, height=15)
        text.insert(INSERT, txt)
        text.pack()
        newWindow.mainloop()
    except:
        showerror("Warning", "No PDF selected")

image = Image.open('lcons\icon9.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img9 = ImageTk.PhotoImage(image)

Button(root, text="PDF\nINFORMATION", font="Helvetica 12 bold", width=120, height=110,
       command=pdf_info, bg='silver', image=img9, compound=TOP).place(x=630, y=302)


# Word to PDF
def word_to_pdf():
    try:
        filename = askopenfilename(initialdir="/", title="Select a File", filetypes=(("Word Files", "*.docx*"), ("all files", "*.*")))
        
        file = asksaveasfilename(initialdir = "C:/Learn Programming/python/college projects/All in one pdf editor", title = "name file", filetypes = (("PDF Files","*.pdf"), ("all files","*.*")))
        new_file = f"{file}.pdf"
        
        convert(filename, new_file)
       
        pages = convert_from_path(new_file, poppler_path=r'C:\Users\anike\Poppler\poppler-0.68.0_x86\poppler-0.68.0\bin', size=(800, 900))

        newWindow = Toplevel()
        newWindow.title("Image to Pdf Viewer")
        newWindow.geometry("810x700")

        scrol_y = Scrollbar(newWindow, orient=VERTICAL)

        pdf = Text(newWindow, yscrollcommand=scrol_y.set, bg="grey")

        scrol_y.pack(side=RIGHT, fill=Y)
        scrol_y.config(command=pdf.yview)

        pdf.pack(fill=BOTH, expand=1)

        photos = []

        for i in range(len(pages)):
            photos.append(ImageTk.PhotoImage(pages[i]))

        for photo in photos:
            pdf.image_create(END, image=photo)
            pdf.insert(END, '\n\n')
        newWindow.mainloop()
    except:
        showerror("Warning", "No pdf selected")

image = Image.open('lcons\icon10.png')
image = image.resize((60, 60), Image.ANTIALIAS)
img10 = ImageTk.PhotoImage(image)

Button(root, text="\nWORD TO PDF", font="Helvetica 12 bold", width=120, height=110, command=word_to_pdf, bg='silver', image=img10, compound=TOP).place(x=380, y=430)

root.mainloop()