# Tool Import hàng loạt vào `Phôi DS.docx` (Project con)

Ứng dụng Streamlit đọc dữ liệu từ nhiều file Word (`.docx`) hoặc file nén (`.zip`) danh sách Chi đoàn, làm sạch dữ liệu và **import trực tiếp vào bản sao sạch của `Phôi DS.docx`** mà **không thay đổi bất kỳ định dạng gốc nào của phôi**.

## Cách chạy

Nháy đúp `run_app.bat` trong thư mục `Project con/`, hoặc chạy từ PowerShell:

```powershell
cd "D:\DS project\Project con"
..\.venv\Scripts\streamlit.exe run app.py
```

## Chạy kiểm thử tự động

```powershell
cd "D:\DS project"
.\.venv\Scripts\python.exe "Project con\tests\test_import.py"
```
