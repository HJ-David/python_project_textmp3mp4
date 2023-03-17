from tkinter import *
from tkinter import filedialog
import pandas as pd
import openpyxl
import xlrd
import PyPDF2
import docx

from gtts import gTTS
import os


def open_file():
    global contents
    global file_path
    # Ask the user to select a file
    file_path = filedialog.askopenfilename(title="Open File", filetypes=[("Text Files", "*.txt"), 
                                                                         ("Excel Files", "*.xlsx *.xls"), 
                                                                         ("Numbers Files", "*.numbers"), 
                                                                         ("Word Files", "*.docx"), 
                                                                         ("PDF Files", "*.pdf"),
                                                                         ("All", "*.*")
                                                                         ])
    # Check the file type and read its contents accordingly
    if file_path.endswith('.txt'):
        # Open the selected file and read its contents
        with open(file_path, 'r', encoding='utf-8') as file:
            contents = file.readlines()
        # Convert the contents list to a string
        contents_str = "".join(contents)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith(('.xlsx', '.xls')):
        # Read the Excel file using pandas
        df = pd.read_excel(file_path, 'Sheet1')
        contents = df[df.columns[0]].values.tolist()
        # Convert the dataframe to a string
        contents_str = df.to_string(index=False)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith('.numbers'):
        # Read the Numbers file using pandas
        df = pd.read_excel(file_path, engine='odf')
        # Convert the dataframe to a string
        contents_str = df.to_string(index=False)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith('.docx'):
        # Read the Word file using docx
        doc = docx.Document(file_path)
        # Extract the text from the document
        contents = [p.text for p in doc.paragraphs]
        # Convert the contents list to a string
        contents_str = "".join(contents)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith('.pdf'):
        # Read the PDF file using PyPDF2
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfFileReader(file)
            contents = []
            for i in range(reader.numPages):
                page = reader.getPage(i)
                contents.append(page.extractText())
        # Convert the contents list to a string
        contents_str = "".join(contents)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)

def convert():
    initialdir = os.path.expanduser(file_path)
    directory_path = filedialog.askdirectory(initialdir=initialdir)
    # do something with the directory_path
    for index, line in enumerate(contents, 1):
        speak = gTTS(line, lang= "en")
        speak.save(str(directory_path) + "//text{0}.mp3".format(index))
        # speak.save("./Tkinter_app/audio//text{0}.mp3".format(index))
    successful_label.config(text="Successful!")
    path_label.config(text="MP3 location: " + str(directory_path))


# Clear the textbox
def clear_text_box():
    my_text.delete(1.0, END)
    successful_label.config(text="")
    path_label.config(text="") 
 
# Create a window with a button to open the file
root = Tk()
root.title('Text to MP3')
root.geometry("500x500")

# Create a textbox
my_text = Text(root, height=30, width=60)
my_text.pack(pady=10)

#Create A Menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Add some dropdown menus
file_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open", command=open_file)
file_menu.add_command(label="Clear", command=clear_text_box)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

open_button = Button(root, text="Convert to MP3", command=convert)
open_button.pack()
# Create the label
successful_label = Label(root, text="")
successful_label.pack()
path_label = Label(root, text="")
path_label.pack()


root.mainloop()
