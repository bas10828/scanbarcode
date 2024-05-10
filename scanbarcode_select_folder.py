from tkinter import Tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.filedialog import askdirectory

import cv2
from pyzbar.pyzbar import decode

import glob #globbing utility.
import os
import shutil

import xlsxwriter
from tkinter import Entry
from tkinter import Label

root = Tk()
root.title('Scan Barcode')
root.resizable(False,False)
root.geometry('300x150')

# เพิ่ม Label สำหรับชื่อไฟล์ Excel
excel_name_label = Label(root, text="Excel Name:")
excel_name_label.pack()

# เพิ่มคอมโพเนนท์ Entry สำหรับกรอกชื่อไฟล์ Excel
file_name_entry = Entry(root)
file_name_entry.pack(expand=True)

def Select_folder():

    # ดึงชื่อไฟล์ Excel ที่ผู้ใช้กรอก
    excel_file_name = file_name_entry.get()
    
    # ตรวจสอบว่าผู้ใช้กรอกชื่อไฟล์หรือไม่
    if not excel_file_name:
        messagebox.showerror("Error", "Please enter a file name.")
        return

    path=askdirectory(title='Select your folder')
    print('Path:',path + '/*')

    scores = ([])
    count = 1
    A = path + '/*'
    for file in glob.glob(A):
        img = cv2.imread(file)
        detectedBarcodes = decode(img)

        if not detectedBarcodes:
            print("Barcode Not Detected or your barcode is blank/corrupted!")

        else:
            for barcode in detectedBarcodes:
                (x,y,w,h) = barcode.rect

                cv2.rectangle(img, (x-10, y-10),
                            (x + w+10, y + h+10),
                            (255,0,0), 2)

                if barcode.data != "" :
                    print(barcode.data)
                    name = os.path.basename(file)
                    data = barcode.data
                    string_data = data.decode()
                    # data_result = string_data[1:]
                    scores.append([name,string_data])                    

        # cv2.imshow("image", img)
        # cv2.waitKey(1)
        # cv2.destroyAllWindows()


    #create excel
    workbook = xlsxwriter.Workbook(f'{excel_file_name}.xlsx')
    worksheet = workbook.add_worksheet("My sheet")

    print("scores",scores)

    row = 0
    col = 0

    for file_name, data_file  in (scores):
        worksheet.write(row, col, file_name)
        worksheet.write(row, col + 1, data_file)
        row += 1

    workbook.close()

open_button = ttk.Button(root,text='Select your folder' ,command=Select_folder)
open_button.pack(expand=True)
root.mainloop()