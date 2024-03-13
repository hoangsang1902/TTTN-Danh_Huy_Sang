import re
from wit import Wit
from tkinter import filedialog
import cv2
from tkinter import *
from tkinter import messagebox
from PIL import ImageTk, Image
import pytesseract
import spacy
import time
from openpyxl import  load_workbook

# Thêm đường dẫn file Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# Thiết lập camera
cap = cv2.VideoCapture(0)

# Tạo giao diện bằng Tkinter
root = Tk()
root.geometry('1300x770')
root.resizable(width=False, height=False)
root.title("TRÍCH XUẤT THÔNG TIN ")
root.configure(bg='white')

#Giao diện chữ
tentruong = Label(root, text="TRƯỜNG ĐẠI HỌC BÁCH KHOA ĐÀ NẴNG", bg='white',font=('Time 25 bold'))
tentruong.pack(side=TOP)
a = Label(root, text="      ", bg='white',font=('Time 5 bold'))
a.pack(side=TOP)
khoa = Label(root, text="KHOA CÔNG NGHỆ THÔNG TIN", bg='white',font=('Time 22 bold'))
khoa.pack(side=TOP)
b = Label(root, text="      ", bg='white',font=('Time 5 bold'))
b.pack(side=TOP)
c = Label(root, text="    ", bg='white',font=('Time 30 bold'))
c.pack(side=TOP)
doan = Label(root, text="DỰ ÁN", bg='white', fg='red',font=('Time 30 bold'))
doan.pack(side=TOP)
detai = Label(root, text="TRÍCH XUẤT THÔNG TIN TỪ GIẤY TỜ TÙY THÂN", bg='white', fg='red', font=('Time 30 bold'))
detai.pack(side=TOP)
# kiemtra = Label(root, text="Nếu thông tin chính xác hãy nhấn NHẬP DỮ LIỆU", bg='white', fg='blue', font=('Time 20 bold'))
# kiemtra.place(x= 630, y= 643)
thoigian = Label(root, text="Vui lòng kiểm tra thông tin", bg='white', fg='blue', font=('Time 25 bold'))
thoigian.place(x= 740, y= 300)

# Chèn logo vào giao diện
logo = cv2.imread('th.jpg')
logo = cv2.resize(logo, (200, 200))
logo = cv2.cvtColor(logo, cv2.COLOR_BGR2RGB)
img = Image.fromarray(logo)
img = ImageTk.PhotoImage(image=img)
Label(root, image=img).place(x=0, y=0)
logo1 = cv2.imread('cntt.jpg')
logo1 = cv2.resize(logo1, (195, 195))
logo1 = cv2.cvtColor(logo1, cv2.COLOR_BGR2RGB)
img1 = Image.fromarray(logo1)
img1 = ImageTk.PhotoImage(image=img1)
Label(root, image=img1).place(x=1095, y=0)

# Set frame lên giao diện
canvas = Canvas(root, width=480, height=320, bg="white")
canvas.place(x=100, y=350)

# Label để hiển thị kết quả
result_label = Label(root, bg='white', fg='blue', font=('Time 13 bold'))
result_label.place(x=700, y=350)

# Khởi tạo client của Wit.ai với access token của bạn
client = Wit(access_token="WQLR6HZQD4Y5KTP3JQYLRH2W5BQROHAB")

# Hàm chọn file ảnh từ máy tính
def choose_image():
    global img_path
    img_path = filedialog.askopenfilename()
    if img_path:
        # Đọc ảnh từ đường dẫn đã chọn và hiển thị lên khung frame
        img = Image.open(img_path)
        img = img.resize((480, 320), Image.ANTIALIAS)
        img = ImageTk.PhotoImage(img)
        canvas.img = img
        canvas.create_image(0, 0, anchor=NW, image=img)
    
        # Xóa kết quả trước đó trên result_label
        result_label.config(text="")

def txtt():
    global img_path
    # Kiểm tra xem đã chọn ảnh chưa
    if not img_path:
        messagebox.showerror("Lỗi", "Vui lòng chọn ảnh trước khi trích xuất thông tin")
        return

    # Đọc ảnh từ đường dẫn đã chọn và hiển thị lên khung frame
    img = Image.open(img_path)
    img = img.resize((480, 320), Image.ANTIALIAS)
    img = ImageTk.PhotoImage(img)
    canvas.img = img
    canvas.create_image(0, 0, anchor=NW, image=img)

    # Trích xuất thông tin từ ảnh và in ra dữ liệu trước khi đưa lên Wit.ai
    text = pytesseract.image_to_string(img_path, lang='vie')

    # Xử lý kết quả: xóa bỏ các kí tự không mong muốn
    cleaned_text = clean_text(text)
    print("Dữ liệu trích xuất từ ảnh:")
    print(text)
    print("Dữ liệu đã làm sạch :")
    print(cleaned_text)

    # Chia văn bản thành 2 đoạn
    text_part1 = cleaned_text.split("\n")[0:len(cleaned_text.split("\n"))//2]
    text_part2 = cleaned_text.split("\n")[len(cleaned_text.split("\n"))//2:]

    # Gửi từng đoạn văn bản đến Wit.ai để phân tích
    response1 = client.message('\n'.join(text_part1))
    response2 = client.message('\n'.join(text_part2))

    # Lấy kết quả từ Wit.ai
    functions1 = []
    functions2 = []
    for line in text_part1:
        response = client.message(line)
        entities = response.get('entities', {})
        for entity in entities:
            function = entities[entity][0]['value']
            functions1.append(function)

    for line in text_part2:
        response = client.message(line)
        entities = response.get('entities', {})
        for entity in entities:
            function = entities[entity][0]['value']
            functions2.append(function)
    
    # Tính độ dài lớn nhất của các dòng
    max_length = max(len(line) for line in text_part1 + text_part2)

    # Hiển thị kết quả trên result_label
    loaithe = ""
    result_text = "ĐÂY LÀ THẺ: HỌC SINH — SINH VIÊN\n"
    # result_text = f"ĐÂY LÀ THẺ: {loaithe}\n"
    max_index_length = len(str(max(len(text_part1), len(text_part2))))
    for i, line in enumerate(text_part1):
        function = functions1[i] if i < len(functions1) else "Unknown"
        entities = response1.get('entities', {})
        # for entity in entities:
        #     if entity == 'Tên thẻ':
        #         loaithe = line

        result_text += f"{str(i+1).rjust(max_index_length)}. {function}  :  {line.ljust(max_length)} \n"

    for i, line in enumerate(text_part2):
        function = functions2[i] if i < len(functions2) else "Unknown"
        entities = response2.get('entities', {})
        # for entity in entities:
        #     if entity == 'Tên thẻ':
        #         loaithe = line

        result_text += f"{str(i+len(text_part1)+1).rjust(max_index_length)}.  {function}  :  {line.ljust(max_length)}\n"
    result_label.config(text=result_text)

    # In kết quả lên terminal
    # result_text = f"ĐÂY LÀ THẺ: {loaithe}"
    result_text = "ĐÂY LÀ THẺ: HỌC SINH — SINH VIÊN\n"
    print(result_text)
    for i, line in enumerate(text_part1):
        function = functions1[i] if i < len(functions1) else "Unknown"
        print(f"{i+1}. {function}  :  {line} ")

    for i, line in enumerate(text_part2):
        function = functions2[i] if i < len(functions2) else "Unknown"
        print(f"{i+len(text_part1)+1}. {function}  :  {line}")




def clean_text(text):
  # Danh sách các kí tự không mong muốn
  unwanted_chars = ["!","@", "#", "$", "%", "^", "&", "*", "(", ")", "=", "+", "~", "[", "]", "{", "}", ";", "'", '"', "<", ">", "?", "|", "\\"]
  
  # Danh sách các từ không mong muốn
  unwanted_words = ["Ngày sinh:","Khóa học :","Ngành :"]
  # Duyệt qua từng dòng trong chuỗi văn bản
  lines = text.splitlines()
  cleaned_lines = []
  for line in lines:
    # Loại bỏ các từ không mong muốn trong dòng
    for word in unwanted_words:
        line = line.replace(word, "")
        # Kiểm tra xem dòng có chứa kí tự đặc biệt hay không
    if any(char in line for char in unwanted_chars):
      # Xóa dòng nếu có kí tự đặc biệt
      continue
    # Kiểm tra xem dòng có phải là dòng trắng hay không
    elif not line.strip():
      # Xóa dòng trắng
      continue
    elif line in ("/mºrith 11 năm/year²01 3", "â"):
          # Xóa 2 dòng cụ thể
      continue
    
    else:
      
      # Loại bỏ các kí tự không mong muốn trong dòng
      cleaned_line = ''.join(char for char in line if char not in unwanted_chars)
      cleaned_lines.append(cleaned_line)

  # Trả về chuỗi văn bản sau khi đã được xử lý
  return '\n'.join(cleaned_lines)

#Nhập dữ liệu qua excel
def get_data():
    global i, wb, ws
    i = 0
    named_tuple = time.localtime()
    time_string = time.strftime("%H:%M:%S-%d/%m/%Y", named_tuple)
    wb = load_workbook('ListSV.xlsx')
    ws = wb.active

    # Lưu các giá trị vào biến
    bo = "BỘ GIÁO DỤC VÀ ĐÀO TẠO"
    truong = "TRƯỜNG ĐẠI HỌC MỞ TP.HỒ CHÍ MINH"
    ten_the = "THẺ HỌC SINH — SINH VIÊN"
    ho_ten = "PHẠM THỊ KHÁNH LINH"  # Giả sử imgcharTEN chứa họ và tên
    ngay_sinh = "14/07/1983"  # Đây là một giả định, bạn cần thay thế bằng giá trị thực tế
    nganh = " Quản trị kinh doanh"  # Đây là một giả định, bạn cần thay thế bằng giá trị thực tế
    nien_khoa = "2005 — 2010"  # Đây là một giả định, bạn cần thay thế bằng giá trị thực tế

    # Thêm dữ liệu vào file excel
    ws = wb.active  
    ws.append([i,"Tên Bộ", "Trường","Tên Thẻ","Họ và tên","Ngày sinh","Ngành","Niên khóa","Thời gian trích xuất"])
    ws.append([i + 1, bo, truong, ten_the, ho_ten, ngay_sinh, nganh, nien_khoa, time_string])
    wb.save('ListSV.xlsx')
    i = i + 1


    # Hiển thị thông báo
    messagebox.showinfo('Thông báo', 'Bạn đã nhập dữ liệu thành công')


# Giao diện nút chọn file ảnh
btn_choose_image = Button(root, text='Chọn ảnh', bg='green', fg='white', font=('Time 15 bold'))
btn_choose_image.config(command=choose_image)
btn_choose_image.place(x=100, y=700)

# Nút nhấn trích xuất dữ liệu
btn_txtt = Button(root, text='Trích xuất thông tin', bg='green', fg='white', font=('Time 15 bold'))
btn_txtt.config(command=txtt)
btn_txtt.place(x=220, y=700)

# Nút nhấn nhập dữ liệu
btn_get = Button(root, text='Nhập dữ liệu', bg='green', fg='white', font=('Time 15 bold'))
btn_get.config(command=get_data)
btn_get.place(x=880, y=700)

# Kết thúc chương trình
root.mainloop()
