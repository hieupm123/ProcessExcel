import tkinter
from tkinter import ttk
from tkinter import *
from readFileExcel import *
from writeFileExcel import *
from tkinter import messagebox

window = tkinter.Tk()
window.title("Data Entry Form")

frame = tkinter.Frame(window)
frame.pack()
writeExcel = WriteFileExcel(); #d
readExcel = ReadFileExcel(); #a

mayVaCongTy = readExcel.docMayVaCongTy("../myexcel/source.xlsx"); #b
dulieu, thang = readExcel.tinhTien(mayVaCongTy, "../myexcel/follow.xlsx"); # c, e

def bangthongtin():
    messagebox.showinfo("Thông báo", "Đã sinh bảng thông tin các máy")
    writeExcel.danhSachDVCNT(dulieu, mayVaCongTy);

def luubang():
    messagebox.showinfo("Thông báo", "Đã lưu các bản sinh vào file " + "output_excel.xlsx")
    writeExcel.luuFile("../myexcel/output_excel.xlsx");



def baocaochung():
    top = Toplevel()
    top.title("Báo cáo chung")
    user_info_frame =tkinter.LabelFrame(top, text="User Information")
    user_info_frame.grid(row= 0, column=0, padx=20, pady=10)

    first_name_label = tkinter.Label(user_info_frame, text="Chọn tháng mong muốn sinh báo cáo")
    first_name_label.grid(row=0, column=0)
    title_combobox = ttk.Combobox(user_info_frame, values=thang)
    title_combobox.grid(row=0, column=1)
    def sinhbaocaochung():
        thangbao = title_combobox.get();
        writeExcel.baoCaoTongHop(thang, dulieu, thangbao);
        messagebox.showinfo("Thông báo", "Đã sinh báo cáo chung")

    button = Button(top, text="Sinh báo cáo", command= sinhbaocaochung)
    button.grid(row=1, column=0, sticky="news", padx=20, pady=10)
    top.mainloop

def baocaochitiet():
    top = Toplevel()
    top.title("Báo cáo chi tiết")
    user_info_frame =tkinter.LabelFrame(top, text="User Information")
    user_info_frame.grid(row= 0, column=0, padx=20, pady=10)

    first_name_label = tkinter.Label(user_info_frame, text="Chọn tháng mong muốn sinh báo cáo")
    first_name_label.grid(row=0, column=0)
    title_combobox = ttk.Combobox(user_info_frame, values=thang)
    title_combobox.grid(row=0, column=1)
    first_name_label = tkinter.Label(user_info_frame, text="Thêm vào mã công ty mong muốn")
    first_name_label.grid(row=1, column=0)

    sll = Entry(user_info_frame);
    sll.grid(row=1, column=1)

    tmp = [];
    def xacnhan():
        sl = sll.get();
        messagebox.showinfo("Thông báo", "Đã thêm công ty " + sl)
        tmp.append(sl);
        sll.delete(0, END)
        

    btnSll = Button(user_info_frame, text="Xác nhận công ty", command= xacnhan)
    btnSll.grid(row=1, column=2, sticky="news", padx=10, pady=8)

    def sinhbaocao():
        thangbao = title_combobox.get();
        messagebox.showinfo("Thông báo", "Đã sinh báo cáo chi tiết")
        writeExcel.baoCaoChiTiet(thang, dulieu, thangbao, tmp);
    
    button = Button(top, text="Sinh báo cáo", command= sinhbaocao)
    button.grid(row=3, column=0, sticky="news", padx=20, pady=10)
    top.mainloop

button1 = Button(frame, text="Sinh bảng thông tin", command= bangthongtin)
button1.grid(row=1, column=0, sticky="news", padx=20, pady=10)

button2 = Button(frame, text="Báo cáo chung", command= baocaochung)
button2.grid(row=2, column=0, sticky="news", padx=20, pady=10)

button3 = Button(frame, text="Báo cáo chi tiết", command= baocaochitiet)
button3.grid(row=3, column=0, sticky="news", padx=20, pady=10)

button4 = Button(frame, text="Lưu các bảng về file excel", command= luubang)
button4.grid(row=4, column=0, sticky="news", padx=20, pady=10)


window.mainloop()