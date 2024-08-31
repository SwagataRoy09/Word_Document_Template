import tkinter as tk
from tkinter import filedialog
from docx import Document
def upload_image():
    file_path = filedialog.askopenfilename(filetypes=[("Image File","*.png;*.jpg;*.jpeg")])
    root.quit()
    return file_path

def create_doc(image_path):
    document = Document()
    try:
        document.add_picture(image_path)
    except FileNotFoundError:
        print("Error: Image file not found. Skipping image insertion.")
    document.save('Output.docx')

root=tk.Tk()
upload_button = tk.Button(root,text="Upload Image",command=upload_image)
upload_button.pack()
root.mainloop(1)
image_path=upload_image()
if image_path:
    create_doc(image_path)

