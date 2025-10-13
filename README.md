# PASDV Streamlit App

Ứng dụng Streamlit để thẩm định phương án sử dụng vốn/kinh doanh từ file **pasdv.docx**.

## Tính năng
- Upload `.docx`, tự động trích xuất thông tin KH & phương án (regex)
- Sidebar cho phép chỉnh sửa thủ công tất cả trường
- Lập kế hoạch trả nợ theo **annuity**, xuất **Excel**
- Tính các **chỉ tiêu CADAP**: DSR, LTV, E/C, Debt/Income, ROI, CFR, Coverage, kiểm tra hợp lý tổng nhu cầu vốn
- Tích hợp **Gemini** (chọn model, nhập API key) để phân tích và khuyến nghị
- Nút **đóng gói** ZIP mã nguồn

## Chạy cục bộ
```bash
pip install -r requirements.txt
streamlit run python.py
```

## Cấu hình Gemini
- Tạo API key cho Google Generative AI
- Khai báo trong Sidebar **hoặc** set secret `GENAI_API_KEY` trên Streamlit Cloud

## Deploy Streamlit Community Cloud
- Push 3 file: `python.py`, `requirements.txt`, `README.md` lên GitHub
- Tạo app mới, chọn file chạy: `python.py`
