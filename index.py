from tkinter import *
from openpyxl import Workbook
from openpyxl.styles import Font,Alignment,Border,Side
from tkinter import font as tkfont

root = Tk()
root.title("Absensi Perkuliahan")
root.resizable(width=False,height=False)
workbook = Workbook()
sheet = workbook.active

styling = tkfont.Font(family='Helvetica',weight='bold', size=15)
styling2 = tkfont.Font(family='Helvetica', size=9)

font = Font(bold=True)
border = Border(left=Side(border_style='thin',color='00000000'),
                right=Side(border_style='thin',color='00000000'),
                top=Side(border_style='thin',color='00000000'),
                bottom=Side(border_style='thin',color='00000000'))

alignment = Alignment(horizontal='center', vertical='center')



HEIGHT = 500
WIDTH = 600
canvas = Canvas(root, height=HEIGHT, width=WIDTH, bg='lightblue')
canvas.pack()

sheet['A1'] = "Mata Kuliah\t:"
A1 = sheet['A1']
A1.font = font
sheet['A2'] = "Tanggal Perkuliahan\t:"
A2 = sheet['A2']
A2.font = font

sheet['A3'] = "No"
A3 = sheet['A3']
A3.font = font
A3.border = border
A3.alignment = alignment

sheet['B3'] = "Nama"
B3 = sheet['B3']
B3.font = font
B3.border = border
B3.alignment = alignment

sheet['C3'] = "NIM"
C3 = sheet['C3']
C3.font = font
C3.border = border
C3.alignment = alignment

sheet['D3'] = "Jurusan"
D3 = sheet['D3']
D3.font = font
D3.border = border
D3.alignment = alignment

num = 0


def InsertData():
    global num
    num = num + 1
    sheetnum = num + 3

    sheet['A'+str(sheetnum)] = num
    DataNo = sheet['A'+str(sheetnum)]
    DataNo.border = border
    DataNo.alignment = alignment

    sheet['B'+str(sheetnum)] = namaEntry.get()
    DataNama = sheet['B'+str(sheetnum)]
    DataNama.border = border
    DataNama.alignment = alignment

    sheet['C' + str(sheetnum)] = NIMEntry.get()
    DataNIM = sheet['C' + str(sheetnum)]
    DataNIM.border = border
    DataNIM.alignment = alignment

    sheet['D' + str(sheetnum)] = jurusanEntry.get()
    DataJurusan = sheet['D' + str(sheetnum)]
    DataJurusan.border = border
    DataJurusan.alignment = alignment

    sheet['B1'] = matkulEntry.get()
    sheet['B2'] = tanggalEntry.get()

    namaEntry.delete(0, END)
    NIMEntry.delete(0, END)
    jurusanEntry.delete(0, END)

def SaveData():
    global informasi
    workbook.save(filename=str(matkulEntry.get())+"_"+str(tanggalEntry.get())+".xlsx")
    informasi['text'] = "Data absen telah di save!\nNama file: "+str(matkulEntry.get())+"_"+str(tanggalEntry.get())+".xlsx"

def CreateNewData():
    global informasi, num
    informasi['text'] = 'Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.'
    namaEntry.delete(0, END)
    NIMEntry.delete(0, END)
    jurusanEntry.delete(0, END)
    matkulEntry.delete(0, END)
    tanggalEntry.delete(0, END)
    num = 0
