import pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Khai báo biến toàn cục để lưu trữ các đối tượng Treeview và khung biểu đồ
tree = None
pie_chart_frame = None
line_chart_frame = None

# Hàm để tải về file Excel mẫu
def download_sample_file():
    # Tạo dữ liệu mẫu với các cột cần thiết
    sample_data = {
        'name': ['Nguyen Van A', 'Tran Thi B'],
        'age': [21, 22],
        'gpa': [3.5, 3.7],
        'attendance': [85, 90],
        'financial_aid': [1, 0],
        'on_time_graduation': [1, 0],
        'gender': ['Nam', 'Nữ'],
        'major': ['Khoa học máy tính', 'Kinh tế']
    }
    sample_df = pd.DataFrame(sample_data)
    
    # Mở hộp thoại lưu file để người dùng chọn nơi lưu file Excel mẫu
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    
    # Nếu người dùng chọn nơi lưu, thì lưu file
    if file_path:
        sample_df.to_excel(file_path, index=False)
        # Thông báo cho người dùng là file đã được lưu thành công
        messagebox.showinfo("Tải về", "File Excel mẫu đã được lưu thành công.")

# Hàm để tạo biểu đồ tròn (pie chart) so sánh tỉ lệ tốt nghiệp đúng hạn và không đúng hạn
def create_pie_chart(df):
    global pie_chart_frame

    # Xóa biểu đồ tròn cũ nếu đã có
    if pie_chart_frame:
        pie_chart_frame.destroy()

    # Tạo khung mới để chứa biểu đồ
    pie_chart_frame = tk.Frame(root, bg='lightgray')
    pie_chart_frame.pack(side='left', fill='both', expand=True)

    # Tính toán số lượng sinh viên tốt nghiệp đúng hạn và không đúng hạn từ dự đoán
    on_time_count = df['predicted_graduation'].sum()
    not_on_time_count = len(df) - on_time_count

    # Tạo biểu đồ tròn từ số liệu
    fig, ax = plt.subplots()
    ax.pie([on_time_count, not_on_time_count], labels=['Đúng hạn', 'Không đúng hạn'], autopct='%1.1f%%', colors=['#5cb85c', '#d9534f'])
    ax.set_title('Tỉ lệ tốt nghiệp đúng hạn')

    # Hiển thị biểu đồ tròn lên giao diện Tkinter
    canvas = FigureCanvasTkAgg(fig, master=pie_chart_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

# Hàm tạo biểu đồ thanh so sánh giữa dự đoán và thực tế
def create_bar_chart(df):
    global line_chart_frame

    # Xóa biểu đồ thanh cũ nếu đã có
    if line_chart_frame:
        line_chart_frame.destroy()

    # Tạo khung mới để chứa biểu đồ bar
    line_chart_frame = tk.Frame(root, bg='lightgray')
    line_chart_frame.pack(side='right', fill='both', expand=True)

    # Tính toán số lượng sinh viên tốt nghiệp đúng hạn và không đúng hạn theo dự đoán và thực tế
    predicted_counts = df['predicted_graduation'].value_counts()
    actual_counts = df['actual_graduation'].value_counts()

    # Tạo danh sách tên các loại (đúng hạn, không đúng hạn)
    categories = ['Không đúng hạn', 'Đúng hạn']
    
    # Số lượng sinh viên dự đoán
    predicted_values = [
        predicted_counts.get(0, 0),  # Không đúng hạn dự đoán
        predicted_counts.get(1, 0)   # Đúng hạn dự đoán
    ]
    # Số lượng sinh viên thực tế
    actual_values = [
        actual_counts.get(0, 0),  # Không đúng hạn thực tế
        actual_counts.get(1, 0)   # Đúng hạn thực tế
    ]

    # Tạo biểu đồ thanh để so sánh giữa dự đoán và thực tế
    fig, ax = plt.subplots()
    bar_width = 0.35
    index = range(len(categories))

    # Vẽ các thanh cho dữ liệu dự đoán và thực tế
    ax.bar(index, predicted_values, bar_width, label='Dự đoán', color='#5bc0de')
    ax.bar([i + bar_width for i in index], actual_values, bar_width, label='Thực tế', color='#5cb85c')

    # Đặt tên cho trục x và y
    ax.set_xlabel('Loại')
    ax.set_ylabel('Số lượng sinh viên')
    ax.set_title('So sánh kết quả dự đoán và thực tế về việc tốt nghiệp đúng hạn')

    # Đặt tên cho từng loại thanh trên trục x
    ax.set_xticks([i + bar_width / 2 for i in index])
    ax.set_xticklabels(categories)
    ax.legend()

    # Hiển thị biểu đồ thanh lên giao diện Tkinter
    canvas = FigureCanvasTkAgg(fig, master=line_chart_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

# Hàm xử lý file Excel và dự đoán kết quả tốt nghiệp
def process_file():
    global tree

    # Mở hộp thoại để người dùng chọn file Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    if not file_path:
        return  # Nếu không chọn file, thoát khỏi hàm

    # Đọc dữ liệu từ file Excel
    df = pd.read_excel(file_path)

    # Lấy dữ liệu đầu vào (X) và mục tiêu (y)
    X = df[['age', 'gpa', 'attendance', 'financial_aid']]
    y = df['on_time_graduation']

    # Chuẩn hóa dữ liệu để đưa vào mô hình
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # Tạo và huấn luyện mô hình Logistic Regression
    model = LogisticRegression()
    model.fit(X_scaled, y)

    # Dự đoán kết quả cho toàn bộ dữ liệu
    y_pred = model.predict(X_scaled)

    # Thêm kết quả dự đoán vào DataFrame
    df['predicted_graduation'] = y_pred
    df['actual_graduation'] = y

    # Xóa dữ liệu cũ trong bảng Treeview nếu có
    if tree:
        for item in tree.get_children():
            tree.delete(item)

    # Tạo Treeview mới nếu chưa có
    if not tree:
        tree = ttk.Treeview(root, columns=('Student', 'Age', 'GPA', 'Attendance', 'Financial_aid', 'Predicted', 'Actual', 'Gender', 'Major'), show='headings')
        tree.heading('Student', text='Sinh viên')
        tree.heading('Age', text='Tuổi')
        tree.heading('GPA', text='GPA')
        tree.heading('Attendance', text='Thời gian học (%)')
        tree.heading('Financial_aid', text='Hỗ trợ tài chính')
        tree.heading('Predicted', text='Dự đoán')
        tree.heading('Actual', text='Thực tế')
        tree.heading('Gender', text='Giới tính')
        tree.heading('Major', text='Chuyên ngành')

        # Đặt kích thước cột cho bảng Treeview
        tree.column('Student', width=120)
        tree.column('Age', width=50)
        tree.column('GPA', width=50)
        tree.column('Attendance', width=100)
        tree.column('Financial_aid', width=100)
        tree.column('Predicted', width=150)
        tree.column('Actual', width=150)
        tree.column('Gender', width=80)
        tree.column('Major', width=150)

        # Hiển thị Treeview lên giao diện
        tree.pack(expand=True, fill='both')

    # Ẩn nhãn thông báo khi dữ liệu đã được chọn
    label_info.pack_forget()

    # Thêm dữ liệu vào bảng Treeview
    for index, row in df.iterrows():
        tree.insert('', 'end', values=(
            row['name'],
            row['age'],
            row['gpa'],
            row['attendance'],
            row['financial_aid'],
            'Tốt nghiệp đúng hạn' if row['predicted_graduation'] == 1 else 'Không tốt nghiệp đúng hạn',
            'Tốt nghiệp đúng hạn' if row['actual_graduation'] == 1 else 'Không tốt nghiệp đúng hạn',
            row['gender'],
            row['major']
        ))

    # Tạo và hiển thị biểu đồ tròn
    create_pie_chart(df)
    # Tạo và hiển thị biểu đồ thanh so sánh
    create_bar_chart(df)

# Khởi tạo giao diện Tkinter
root = tk.Tk()
root.title("AI Dự đoán sinh viên tốt nghiệp")
root.geometry("1200x900")
root.configure(bg='lightblue')

# Thêm tiêu đề lớn cho giao diện
title_label = tk.Label(root, text="AI Dự đoán sinh viên tốt nghiệp", font=("Arial", 20, "bold"), bg='lightblue', fg='darkblue')
title_label.pack(pady=20)

# Tạo khung cho các nút
button_frame = tk.Frame(root, bg='lightblue')
button_frame.pack(pady=10)

# Thêm nút chọn file Excel để dự đoán
btn_load = tk.Button(button_frame, text="Chọn dữ liệu để dự đoán", command=process_file, bg='#0275d8', fg='white', font=("Arial", 12), width=25)
btn_load.pack(side='left', padx=10)

# Thêm nút tải về file Excel mẫu
btn_download = tk.Button(button_frame, text="Tải về file Excel mẫu", command=download_sample_file, bg='#5cb85c', fg='white', font=("Arial", 12), width=25)
btn_download.pack(side='left', padx=10)

# Thêm nhãn thông báo hướng dẫn người dùng
label_info = tk.Label(root, text="Hãy thiết lập dữ liệu theo file Excel mẫu", font=("Arial", 10), bg='lightblue', fg='darkblue')
label_info.pack(pady=10)

# Bắt đầu vòng lặp giao diện Tkinter
root.mainloop()
