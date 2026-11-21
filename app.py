import streamlit as st
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
import pandas as pd
import io

st.set_page_config(page_title="Tool Chuyển Dữ Liệu Excel", page_icon="📊", layout="wide")

st.title("📊 Tool Chuyển Dữ Liệu Excel Theo Yêu Cầu")
st.markdown("Tải lên **File Nguồn (File 1)** và **File Đích (File 2)**, chọn các cột cần đối chiếu và điền dữ liệu tự động.")

# Tạo 2 cột giao diện để tải file
col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ File Nguồn (Lấy dữ liệu từ đây)")
    file_1 = st.file_uploader("Chọn File 1 (.xlsx)", type=["xlsx"], key="f1")
    wb1 = None
    sheet_1 = None
    if file_1:
        wb1 = openpyxl.load_workbook(file_1, data_only=True)
        sheet_1 = st.selectbox("Chọn Sheet ở File 1:", wb1.sheetnames, key="s1")

with col2:
    st.subheader("2️⃣ File Đích (Điền dữ liệu vào đây)")
    file_2 = st.file_uploader("Chọn File 2 (.xlsx)", type=["xlsx"], key="f2")
    wb2 = None
    sheet_2 = None
    if file_2:
        wb2 = openpyxl.load_workbook(file_2)
        sheet_2 = st.selectbox("Chọn Sheet ở File 2:", wb2.sheetnames, key="s2")

st.divider()

# Phần cấu hình linh hoạt
if file_1 and file_2 and sheet_1 and sheet_2:
    st.subheader("⚙️ Cấu hình cột đối chiếu & điền dữ liệu")
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        dong_bat_dau_1 = st.number_input("Dòng bắt đầu dữ liệu (File 1):", min_value=1, value=2, step=1)
        cot_ma_1 = st.text_input("Cột Mã/Khóa đối chiếu (File 1):", value="A").strip().upper()
    with c2:
        dong_bat_dau_2 = st.number_input("Dòng bắt đầu dữ liệu (File 2):", min_value=1, value=2, step=1)
        cot_ma_2 = st.text_input("Cột Mã/Khóa đối chiếu (File 2):", value="A").strip().upper()
    with c3:
        cot_nguon_list = st.text_input(
            "Các cột CẦN LẤY ở File 1 (cách nhau dấu phẩy):", 
            value="B, C",
            help="Ví dụ: B hoặc B, D, F"
        ).strip().upper()
    with c4:
        cot_dich_list = st.text_input(
            "Các cột SẼ ĐIỀN ở File 2 (tương ứng):", 
            value="D, E",
            help="Số lượng cột phải khớp với bên File 1. Ví dụ: D hoặc D, G, H"
        ).strip().upper()

    # Xem trước dữ liệu 5 dòng đầu để người dùng dễ kiểm tra tên cột
    with st.expander("👀 Xem trước 5 dòng đầu của 2 file để kiểm tra tên cột"):
        prev1, prev2 = st.columns(2)
        with prev1:
            st.caption(f"File 1 - {sheet_1}")
            ws1_prev = wb1[sheet_1]
            data1 = list(ws1_prev.iter_rows(values_only=True))[:6]
            if data1:
                cols1 = [f"Cột {get_column_letter(i+1)}" for i in range(len(data1[0]))]
                st.dataframe(pd.DataFrame(data1, columns=cols1), use_container_width=True)
        with prev2:
            st.caption(f"File 2 - {sheet_2}")
            ws2_prev = wb2[sheet_2]
            data2 = list(ws2_prev.iter_rows(values_only=True))[:6]
            if data2:
                cols2 = [f"Cột {get_column_letter(i+1)}" for i in range(len(data2[0]))]
                st.dataframe(pd.DataFrame(data2, columns=cols2), use_container_width=True)

    # Tùy chọn nâng cao
    cong_gop = st.checkbox(
        "Cộng gộp giá trị nếu File 1 có nhiều dòng trùng mã (dành cho số liệu)", 
        value=False
    )

    if st.button("🚀 Bắt đầu chuyển dữ liệu", type="primary", use_container_width=True):
        try:
            list_nguon = [c.strip() for c in cot_nguon_list.split(",") if c.strip()]
            list_dich = [c.strip() for c in cot_dich_list.split(",") if c.strip()]

            if len(list_nguon) != len(list_dich):
                st.error(f"❌ Lỗi: Số lượng cột lấy ở File 1 ({len(list_nguon)} cột) phải bằng số lượng cột điền ở File 2 ({len(list_dich)} cột)!")
            else:
                idx_ma_1 = column_index_from_string(cot_ma_1)
                idx_ma_2 = column_index_from_string(cot_ma_2)
                idx_nguon_list = [column_index_from_string(c) for c in list_nguon]
                idx_dich_list = [column_index_from_string(c) for c in list_dich]

                ws1 = wb1[sheet_1]
                ws2 = wb2[sheet_2]

                # Đọc dữ liệu từ File 1
                du_lieu_goc = {}
                for r in range(dong_bat_dau_1, ws1.max_row + 1):
                    ma = ws1.cell(row=r, column=idx_ma_1).value
                    if ma is not None:
                        ma_sach = str(ma).strip()
                        vals = [ws1.cell(row=r, column=c_idx).value for c_idx in idx_nguon_list]
                        
                        if ma_sach not in du_lieu_goc:
                            du_lieu_goc[ma_sach] = vals
                        elif cong_gop:
                            # Cộng dồn nếu là số
                            for i in range(len(vals)):
                                v_cu = du_lieu_goc[ma_sach][i] or 0
                                v_moi = vals[i] or 0
                                if isinstance(v_cu, (int, float)) and isinstance(v_moi, (int, float)):
                                    du_lieu_goc[ma_sach][i] = v_cu + v_moi
                        else:
                            du_lieu_goc[ma_sach] = vals

                # Điền dữ liệu sang File 2
                so_dong_khop = 0
                ma_thieu = []
                for r in range(dong_bat_dau_2, ws2.max_row + 1):
                    ma2 = ws2.cell(row=r, column=idx_ma_2).value
                    if ma2 is not None:
                        ma2_sach = str(ma2).strip()
                        if ma2_sach in du_lieu_goc:
                            vals_to_write = du_lieu_goc[ma2_sach]
                            for i, c_dich_idx in enumerate(idx_dich_list):
                                ws2.cell(row=r, column=c_dich_idx).value = vals_to_write[i]
                            so_dong_khop += 1
                        else:
                            ma_thieu.append(ma2_sach)

                # Xuất file kết quả ra bộ nhớ đệm để tải về
                output = io.BytesIO()
                wb2.save(output)
                output.seek(0)

                st.success(f"✅ Hoàn tất! Đã đối chiếu và điền dữ liệu thành công cho **{so_dong_khop}** dòng.")
                
                if ma_thieu:
                    with st.expander(f"⚠️ Có {len(ma_thieu)} mã ở File 2 không tìm thấy trong File 1 (Bấm để xem)"):
                        st.write(", ".join(ma_thieu))

                st.download_button(
                    label="📥 Tải File Excel Kết Quả Về Máy",
                    data=output,
                    file_name=f"Ket_Qua_{file_2.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

        except Exception as e:
            st.error(f"❌ Có lỗi xảy ra (vui lòng kiểm tra lại tên chữ cái cột): {e}")
