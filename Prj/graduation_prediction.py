import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

# Khai báo biến toàn cục
tree = None

def process_file():
    global tree  # Khai báo rằng biến tree là toàn cục

    # Mở hộp thoại để người dùng chọn file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        return  # Nếu không chọn file, thoát khỏi hàm

    # Đọc dữ liệu từ file Excel
    df = pd.read_excel(file_path)

    # Tách biến đầu vào (X) và biến mục tiêu (y)
    X = df[['age', 'gpa', 'attendance', 'financial_aid']]
    y = df['on_time_graduation']

    # Chuẩn hóa dữ liệu
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # Khởi tạo và huấn luyện mô hình
    model = LogisticRegression()
    model.fit(X_scaled, y)

    # Dự đoán cho toàn bộ dữ liệu
    y_pred = model.predict(X_scaled)

    # Thêm dự đoán vào DataFrame
    df['predicted_graduation'] = y_pred
    df['actual_graduation'] = y

    # Xóa dữ liệu cũ trong bảng Treeview nếu đã có
    if tree:
        for item in tree.get_children():
            tree.delete(item)

    # Hiển thị bảng và tiêu đề cột nếu có dữ liệu
    if not tree:
        # Tạo bảng Treeview để hiển thị kết quả
        tree = ttk.Treeview(root, columns=('Student', 'Age', 'GPA', 'Attendance', 'Predicted', 'Actual', 'Gender', 'Major'), show='headings')

        # Đặt tiêu đề cho các cột
        tree.heading('Student', text='Sinh viên')
        tree.heading('Age', text='Tuổi')
        tree.heading('GPA', text='GPA')
        tree.heading('Attendance', text='Thời gian học (%)')
        tree.heading('Predicted', text='Dự đoán')
        tree.heading('Actual', text='Thực tế')
        tree.heading('Gender', text='Giới tính')
        tree.heading('Major', text='Chuyên ngành')

        # Đặt kích thước cho các cột
        tree.column('Student', width=120)
        tree.column('Age', width=50)
        tree.column('GPA', width=50)
        tree.column('Attendance', width=100)
        tree.column('Predicted', width=150)
        tree.column('Actual', width=150)
        tree.column('Gender', width=80)
        tree.column('Major', width=150)

        # Đặt bảng trong cửa sổ giao diện
        tree.pack(expand=True, fill='both')

    # Ẩn label thông báo khi dữ liệu được tải
    label_info.pack_forget()

    # Thêm dữ liệu vào bảng
    for index, row in df.iterrows():
        tree.insert('', 'end', values=(
            row['name'],
            row['age'],
            row['gpa'],
            row['attendance'],
            'Tốt nghiệp đúng hạn' if row['predicted_graduation'] == 1 else 'Không tốt nghiệp đúng hạn',
            'Tốt nghiệp đúng hạn' if row['actual_graduation'] == 1 else 'Không tốt nghiệp đúng hạn',
            row['gender'],
            row['major']
        ))

# Tạo giao diện đồ họa Tkinter
root = tk.Tk()
root.title("Dự đoán tốt nghiệp sinh viên")
root.geometry("1000x600")

# Thêm nút để người dùng chọn file
btn_load = tk.Button(root, text="Chọn File Excel", command=process_file)
btn_load.pack(pady=10)

# Thêm label thông báo yêu cầu người dùng chọn file
label_info = tk.Label(root, text="Vui lòng chọn file Excel để dự đoán", font=("Arial", 12))
label_info.pack(pady=20)

# Khởi tạo Treeview (bảng) là None để kiểm tra khi dữ liệu được tải
tree = None

# Bắt đầu vòng lặp giao diện đồ họa
root.mainloop()
