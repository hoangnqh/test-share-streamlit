import streamlit as st
import pandas as pd
import base64
from io import BytesIO
import pandas as pd
import openpyxl
from datetime import datetime

def main():
    st.title("Tạo/Cập nhật bảng \"Theo dõi\"")

    # Widget để người dùng tải lên tệp Excel
    uploaded_file = st.file_uploader("Chọn tệp Excel đã trích xuất từ Base", type=["xls", "xlsx"])

    if uploaded_file is not None:
        # Đọc dữ liệu từ tệp Excel vào DataFrame
        df = solve(uploaded_file)
        # df = pd.read_excel(uploaded_file, engine='openpyxl')

        # # Hiển thị dữ liệu trong DataFrame
        # st.write("Dữ liệu từ tệp Excel:")
        # st.dataframe(df)

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
    href = f'<a href="data:application/octet-stream;base64,{base64.b64encode(buffer.read()).decode("utf-8")}" download="Theo_doi.xlsx">Tải xuống tệp Excel kết quả</a>'

    # Giải phóng bộ đệm tạm thời
    buffer.close()

    return href


def check_days(day1_str, day2_str):
    day1 = datetime.strptime(day1_str, "%Y/%m/%d")
    day2 = datetime.strptime(day2_str, "%Y/%m/%d")
    # Tính số ngày đi (lấy giá trị tuyệt đối để đảm bảo kết quả dương)
    cnt = abs((day1 - day2).days)
    if cnt <= 31:
        return True
    return False

def convert_to_datetime(date_str):
    if not isinstance(date_str, str):
        return date_str  # Trả về nguyên bản nếu không phải chuỗi
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d 00:00:00")
        formatted_date = date_obj.strftime("%Y/%m/%d")
        return formatted_date
    except ValueError:
        return date_str

def solve(file_path):
    # Mở tệp Excel
    wb = openpyxl.load_workbook(file_path, data_only=True)

    # Lấy danh sách tên các bảng
    sheet_names = wb.sheetnames

    dict = {}

    # Duyệt qua từng bảng và đọc dưới dạng DataFrame
    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        # print(f'Bảng: {sheet_name}')
        
        data = sheet.values
        columns = next(data)  # Lấy hàng đầu tiên làm tên cột
        df = pd.DataFrame(data, columns=columns, dtype=str)
        column_names = df.columns.tolist()
        for index, row in df.iterrows():
            for col in column_names:
                df.loc[index,col] = convert_to_datetime(df.loc[index,col])
        dict[sheet_name] = df
        # print(df)

        # print('\n')  # In ra một dòng trống giữa các bảng

    # Đóng tệp Excel sau khi hoàn thành
    wb.close()
    ####################################################################################
        
    gio_hoc = {}
    gio_hoc["7h45 - 9h45"] = "07:45 - 09:45"
    gio_hoc["10h - 12h"] = "10:00 - 12:00"
    gio_hoc["16h - 18h"] = "16:00 - 18:00"
    gio_hoc["18h - 20h"] = "18:00 - 20:00"

    lopHoc_khoi = {}
    lopHoc_ngayBatDauHoc = {}
    khoi = {}
    khoi['4'] = []
    khoi['5'] = []
    for index, row in dict["🤡 Lớp học"].iterrows():
        if row["Trạng thái"] != "Đang học":
            continue
        lopHoc_khoi[row["Lớp học"]] = row["Khối"]
        khoi[row["Khối"]].append(row["Lớp học"])
        lopHoc_ngayBatDauHoc[row["Lớp học"]] = row["Ngày bắt đầu học"]

    requiredBuoiHoc = ["Lớp", "Khu vực", "Buổi học thứ", "Môn học", "# môn học", "Giáo viên", "Ngày học", "Giờ học"]

        

    columns = [
        "Khu vực", # (Type: Một lựa chọn)
        "Khối", # (Type: Số)
        "Lớp của buổi học", # (Type: Một lựa chọn)
        "Lớp của học viên", # (Type: Một lựa chọn)
        "Môn học", # (Type: Một lựa chọn)
        "Giáo viên", # (Type: Nhiều dòng)
        "Trợ giảng", # (Type: Nhiều dòng)
        "Buổi học (môn học)", # (Type: Số)
        "Buổi học (lớp học)", # (Type: Số)
        "Ngày học", # (Type: Ngày)
        "Giờ học", # (Type: Một lựa chọn)
        "Ca học", # (Type: Một lựa chọn)
        "Ngày, giờ, lớp học", # ="Ngày học"+" "+"Giờ học"+" "+"Lớp của buổi học" (Type: Nhiều dòng)
        "Thứ trong tuần", # (Type: Một lựa chọn)
        "Mã học viên", # (Type: Liên kết 1 chiều tới Mã học viên)
        "Loại", # 0 là "Học bù", 1 là "Học chính" (Type: Một lựa chọn)
        ################################################################################
        "Đã chọn học bù", # Là học chính hoặc nếu học bù thì cùng ngày với ngày học bù của ngày học chính. (Type: Công thức)
        "Điểm danh", # "1" là đi học, "0" là nghỉ (Type: Một lựa chọn)
        "Kỷ luật", #
        "Hiểu bài", #
        "Tương tác", #
        "BTVN", #
        "Note", #
        "Nhận xét của GV", # (Type: Nhiều dòng)
        "Mini test", 
        "Ngày học bù", # Để trống hoặc ghi rõ ngày (Type: Nhiều dòng)
        "Ngày có thể học bù", # Tất cả những ngày mà có thể bù nếu có thể. (Type: Nhiều lựa chọn)
        "Ngày có thể học bù hiện tại" # Tất cả những ngày học bù mà có thể cho đến hiện tại (Type: Công thức)
        ]
    df = pd.DataFrame(columns=columns)


    for index1, buoiHoc in dict["😏 Buổi học"].iterrows():
        checkRequire = True
        for temp in requiredBuoiHoc:
            if buoiHoc[temp] == None:
                checkRequire = False
        if checkRequire == False:
            continue

        # if check_days(lopHoc_ngayBatDauHoc[buoiHoc["Lớp"]], lopHoc_ngayBatDauHoc[tenLopHoc]):
        for index2, hocVien in dict["Học viên"].iterrows():
            if hocVien["Lớp học"] != buoiHoc["Lớp"]:
                continue

            # Thêm ngày học bù
            list_ngay_co_the_hoc_bu = []
            for index3, buoiHocBu in dict["😏 Buổi học"].iterrows():
                checkRequire = True
                for temp in requiredBuoiHoc:
                    if buoiHoc[temp] == None:
                        checkRequire = False
                if checkRequire == False:
                    continue
                
                if (buoiHoc["Lớp"] != buoiHocBu["Lớp"] and
                    lopHoc_khoi[buoiHoc["Lớp"]] == lopHoc_khoi[buoiHocBu["Lớp"]] and 
                    buoiHoc["Môn học"] ==  buoiHocBu["Môn học"] and 
                    buoiHoc["# môn học"] == buoiHocBu["# môn học"] and  
                    check_days(buoiHoc["Ngày học"], buoiHocBu["Ngày học"])):
                    # Ngày có thể học bù
                    list_ngay_co_the_hoc_bu.append(buoiHocBu["Ngày học"]+" "+gio_hoc[buoiHocBu["Giờ học"]]+" "+buoiHocBu["Lớp"]+",")
                    
                    new_row2 = {
                        "Khu vực": buoiHocBu["Khu vực"], 
                        "Khối": lopHoc_khoi[buoiHocBu["Lớp"]], 
                        "Lớp của buổi học": buoiHocBu["Lớp"], 
                        "Lớp của học viên": hocVien["Lớp học"], 
                        "Môn học": buoiHocBu["Môn học"], 
                        "Giáo viên": buoiHocBu["Giáo viên"], 
                        "Trợ giảng": buoiHocBu["Trợ giảng"], 
                        "Buổi học (môn học)": buoiHocBu["# môn học"], 
                        "Buổi học (lớp học)": buoiHocBu["Buổi học thứ"], 
                        "Ngày học": buoiHocBu["Ngày học"], 
                        "Giờ học": gio_hoc[buoiHocBu["Giờ học"]], 
                        "Ca học": buoiHocBu["Buổi"], 
                        "Ngày, giờ, lớp học": buoiHocBu["Ngày học"]+" "+gio_hoc[buoiHocBu["Giờ học"]]+" "+buoiHocBu["Lớp"], # Viết code
                        "Thứ trong tuần": buoiHocBu["Thứ"], 
                        "Mã học viên": hocVien["Mã học viên"],
                        "Loại": "0", # Viết code
                        ######################################
                        "Đã chọn học bù": None, # Công thức                       
                        "Điểm danh": None, 
                        "Kỷ luật": None, 
                        "Hiểu bài": None,
                        "Tương tác": None, 
                        "BTVN": None, 
                        "Note": None, 
                        "Nhận xét của GV": None, 
                        "Mini test": None, 
                        "Ngày học bù": None, 
                        "Ngày có thể học bù": None, # Viết code
                        "Ngày có thể học bù hiện tại": None # Công thức
                    } 
                    new_df = pd.DataFrame([new_row2])
                    df = pd.concat([df, new_df], ignore_index=True)
            list_ngay_co_the_hoc_bu.sort()
            ngay_co_the_hoc_bu = ""
            for ngayhocbu in list_ngay_co_the_hoc_bu:
                ngay_co_the_hoc_bu += ngayhocbu
            ngay_co_the_hoc_bu=ngay_co_the_hoc_bu[:-1]
            new_row = {
                "Khu vực": buoiHoc["Khu vực"], 
                "Khối": lopHoc_khoi[buoiHoc["Lớp"]], 
                "Lớp của buổi học": buoiHoc["Lớp"], 
                "Lớp của học viên": hocVien["Lớp học"], 
                "Môn học": buoiHoc["Môn học"], 
                "Giáo viên": buoiHoc["Giáo viên"], 
                "Trợ giảng": buoiHoc["Trợ giảng"], 
                "Buổi học (môn học)": buoiHoc["# môn học"], 
                "Buổi học (lớp học)": buoiHoc["Buổi học thứ"], 
                "Ngày học": buoiHoc["Ngày học"], 
                "Giờ học": gio_hoc[buoiHoc["Giờ học"]], 
                "Ca học": buoiHoc["Buổi"], 
                "Ngày, giờ, lớp học": buoiHoc["Ngày học"]+" "+gio_hoc[buoiHoc["Giờ học"]]+" "+buoiHoc["Lớp"], # Viết code
                "Thứ trong tuần": buoiHoc["Thứ"], 
                "Mã học viên": hocVien["Mã học viên"],
                "Loại": "1", # Viết code
                ######################################
                "Đã chọn học bù": None, # Công thức                       
                "Điểm danh": None, 
                "Kỷ luật": None, 
                "Hiểu bài": None,
                "Tương tác": None, 
                "BTVN": None, 
                "Note": None, 
                "Nhận xét của GV": None, 
                "Mini test": None, 
                "Ngày học bù": None, 
                "Ngày có thể học bù": ngay_co_the_hoc_bu, # Viết code
                "Ngày có thể học bù hiện tại": None # Công thức
            } 
            # Thêm ngày học chính
            new_df = pd.DataFrame([new_row])
            df = pd.concat([df, new_df], ignore_index=True)
    fixed_cols = [
        "Khu vực", # (Type: Một lựa chọn)
        "Khối", # (Type: Số)
        "Lớp của buổi học", # (Type: Một lựa chọn)
        "Lớp của học viên", # (Type: Một lựa chọn)
        "Môn học", # (Type: Một lựa chọn)
        "Giáo viên", # (Type: Nhiều dòng)
        "Trợ giảng", # (Type: Nhiều dòng)
        "Buổi học (môn học)", # (Type: Số)
        "Buổi học (lớp học)", # (Type: Số)
        "Ngày học", # (Type: Ngày)
        "Giờ học", # (Type: Một lựa chọn)
        "Ca học", # (Type: Một lựa chọn)
        "Ngày, giờ, lớp học", # ="Ngày học"+" "+"Giờ học" (Type: Nhiều dòng)
        "Thứ trong tuần", # (Type: Một lựa chọn)
        "Mã học viên", # (Type: Liên kết 1 chiều tới Mã học viên)
        "Loại", # 0 là "Học bù", 1 là "Học chính" (Type: Một lựa chọn)
    ]


    # Sau này bổ xung thêm tính năng nếu lớp học đã kết thúc thì chỉ lấy ngày điểm danh và ngày học chính mà thôi cho file nhẹ hơn

    final_df = df.copy()
    if "Theo dõi" in dict:
        # Thêm thông tin từ file cũ
        for index1, theoDoiCu in dict["Theo dõi"].iterrows():
            cnt_theoDoi = 0
            max_cnt = 0
            for index2, theoDoi in df.iterrows():
                cnt = 0
                for col in fixed_cols:
                    if col == "Ngày học":
                        if convert_to_datetime(theoDoiCu[col]) == theoDoi[col]:
                            cnt += 1
                    elif theoDoiCu[col] == theoDoi[col]:
                        cnt += 1
                max_cnt = max(cnt, max_cnt)
                if cnt == len(fixed_cols):
                    # print("--------------------Nếu đã có thì thay đổi trường thông tin")
                    # print(theoDoiCu[col] )
                    # Nếu đã có thì thay đổi trường thông tin
                    for col in columns:
                        if final_df.loc[index2, col] == None:
                            final_df.loc[index2, col] = theoDoiCu[col] 
                else:
                    cnt_theoDoi += 1
                    
            if cnt_theoDoi ==  df.shape[0]:
                # Nếu chưa có thì thêm vào cuối
                    new_row = {}
                    for col in columns:
                        new_row[col] = theoDoiCu[col] 
                    new_df = pd.DataFrame([new_row])
                    final_df = pd.concat([final_df, new_df], ignore_index=True)
                    # print("--------------------Nếu chưa có thì thêm vào cuối")
                    # print(theoDoiCu)
            # print(max_cnt)
    return final_df

if __name__ == "__main__":
    main()