import streamlit as st
import pandas as pd
import base64
from io import BytesIO
import openpyxl

def main():
    st.title("Chương trình đọc tệp Excel và cho phép tải xuống")

    # Widget để người dùng tải lên tệp Excel
    uploaded_file = st.file_uploader("Chọn tệp Excel", type=["xls", "xlsx"])

    if uploaded_file is not None:
        # Đọc dữ liệu từ tệp Excel vào DataFrame
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # Hiển thị dữ liệu trong DataFrame
        st.write("Dữ liệu từ tệp Excel:")
        st.dataframe(df)

        # Tạo liên kết để tải xuống tệp Excel
        tmp_download_link = download_link(df)
        st.markdown(tmp_download_link, unsafe_allow_html=True)

def download_link(df):
    # Tạo một bộ đệm tạm thời để lưu trữ dữ liệu Excel
    buffer = BytesIO()

    # Ghi DataFrame vào bộ đệm tạm thời
    df.to_excel(buffer, index=False)

    # Đặt con trỏ của bộ đệm về đầu tệp
    buffer.seek(0)

    # Tạo đường dẫn tải xuống và trả về mã HTML cho liên kết tải xuống
    href = f'<a href="data:application/octet-stream;base64,{base64.b64encode(buffer.read()).decode("utf-8")}" download="data.xlsx">Tải xuống tệp Excel</a>'

    # Giải phóng bộ đệm tạm thời
    buffer.close()

    return href

if __name__ == "__main__":
    main()
