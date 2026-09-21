import traceback
import streamlit as st
from datetime import datetime
from attendance_engine import generate_reconciled_excel, parse_department_pdf

st.set_page_config(page_title="HR Attendance Reconciler", page_icon="📊", layout="wide")

st.title("📊 Sistem Penyelarasan Kehadiran Bulanan (HR Matrix)")
st.write("Sistem ini akan mencantumkan kehadiran staf individu, rekod cuti syarikat, dan time-off jabatan ke dalam satu format matriks Excel.")

# Pilihan Bulan & Tahun
c_year, c_month = st.columns(2)
with c_year:
    selected_year = st.selectbox("Tahun", [2025, 2026, 2027], index=1)
with c_month:
    selected_month = st.selectbox(
        "Bulan",
        list(range(1, 13)),
        index= datetime.now().month - 1,
        format_func=lambda m: datetime(2000, m, 1).strftime('%B')
    )

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. Kehadiran (PDF Jabatan)")
    att_files = st.file_uploader(
        "Muat naik PDF Clock In/Out mengikut Jabatan (boleh pilih banyak fail: BD.pdf, CS.pdf, dll)", 
        type=["pdf"], 
        accept_multiple_files=True
    )

with col2:
    st.subheader("2. Laporan Cuti (Semua Staf)")
    leave_file = st.file_uploader(
        "Muat naik fail Cuti Syarikat (cth: ALL DEPARTMENT.pdf)", 
        type=["pdf"]
    )

with col3:
    st.subheader("3. Laporan Time-Off (Mengikut Jabatan)")
    timeoff_files = st.file_uploader(
        "Muat naik fail Timeslip/Time-Off (boleh pilih banyak jabatan)", 
        type=["pdf"], 
        accept_multiple_files=True
    )

st.divider()

if st.button("🚀 Jana Laporan Final Excel", type="primary", use_container_width=True):
    if not att_files:
        st.warning("⚠️ Sila muat naik sekurang-kurangnya satu fail kehadiran individu.")
    elif not leave_file:
        st.warning("⚠️ Sila muat naik fail Laporan Cuti (Leave Report).")
    else:
        error_placeholder = st.empty()
        success_placeholder = st.empty()
        download_placeholder = st.empty()

        with st.spinner("Sedang memproses & menyelaraskan data kehadiran merentas jabatan..."):
            try:
                # Read all file bytes upfront
                att_bytes_list = [f.read() for f in att_files]
                leave_bytes = leave_file.read()
                timeoff_bytes_list = [f.read() for f in timeoff_files] if timeoff_files else []

                # Validate files are not empty
                for i, b in enumerate(att_bytes_list):
                    if len(b) == 0:
                        raise ValueError(f"Fail kehadiran #{i+1} kosong (0 bytes). Sila muat naik semula.")
                if len(leave_bytes) == 0:
                    raise ValueError("Fail cuti kosong (0 bytes). Sila muat naik semula.")

                # Count employees BEFORE generating (uses same bytes)
                total_emp = sum(len(parse_department_pdf(b)) for b in att_bytes_list)

                # Generate Excel
                excel_output = generate_reconciled_excel(
                    attendance_pdf_list=att_bytes_list,
                    leave_pdf_bytes=leave_bytes,
                    timeoff_pdf_list=timeoff_bytes_list,
                    year=selected_year,
                    month=selected_month
                )

                month_name = datetime(selected_year, selected_month, 1).strftime('%B')
                file_name = f"Final_Attendance_{month_name}_{selected_year}.xlsx"

            except Exception as e:
                tb = traceback.format_exc()
                error_placeholder.error(
                    f"❌ Ralat semasa memproses fail:\n\n**{type(e).__name__}**: {e}\n\n"
                    f"```\n{tb}\n```"
                )
                st.stop()

        # Show results OUTSIDE the spinner (renders properly over network)
        success_placeholder.success(
            f"✅ Berjaya menjana laporan untuk **{total_emp}** pekerja "
            f"daripada **{len(att_files)}** fail jabatan!"
        )
        download_placeholder.download_button(
            label=f"📥 Muat Turun {file_name}",
            data=excel_output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )