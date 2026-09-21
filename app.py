import traceback
import io
import streamlit as st
from datetime import datetime
import pypdf
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
        index=datetime.now().month - 1,
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
        # --- State variables ---
        result_excel  = None
        result_name   = None
        result_count  = 0
        error_msg     = None
        debug_lines   = []

        with st.spinner("Sedang memproses & menyelaraskan data kehadiran merentas jabatan..."):

            try:
                # 1. Read bytes
                att_bytes_list     = [f.read() for f in att_files]
                leave_bytes        = leave_file.read()
                timeoff_bytes_list = [f.read() for f in timeoff_files] if timeoff_files else []

                # 2. Validate not empty
                for i, b in enumerate(att_bytes_list):
                    if len(b) == 0:
                        raise ValueError(f"Fail kehadiran #{i+1} kosong (0 bytes). Sila muat naik semula.")
                if len(leave_bytes) == 0:
                    raise ValueError("Fail cuti kosong (0 bytes). Sila muat naik semula.")

                # 3. Debug: preview raw PDF text from first attendance file
                try:
                    reader     = pypdf.PdfReader(io.BytesIO(att_bytes_list[0]))
                    first_page = reader.pages[0].extract_text() or ""
                    debug_lines = first_page.split("\n")[:40]
                except Exception as de:
                    debug_lines = [f"[Gagal baca PDF preview: {de}]"]

                # 4. Parse employees
                all_employees = []
                for b in att_bytes_list:
                    emps = parse_department_pdf(b)
                    all_employees.extend(emps)
                result_count = len(all_employees)

                # 5. Generate Excel
                excel_bytes = generate_reconciled_excel(
                    attendance_pdf_list=att_bytes_list,
                    leave_pdf_bytes=leave_bytes,
                    timeoff_pdf_list=timeoff_bytes_list,
                    year=selected_year,
                    month=selected_month
                )

                month_name  = datetime(selected_year, selected_month, 1).strftime('%B')
                result_name = f"Final_Attendance_{month_name}_{selected_year}.xlsx"
                result_excel = excel_bytes

            except Exception as e:
                error_msg = f"**{type(e).__name__}**: {e}\n\n```\n{traceback.format_exc()}\n```"

        # --- Render results AFTER spinner exits ---
        if error_msg:
            st.error(f"❌ Ralat semasa memproses fail:\n\n{error_msg}")

        elif result_excel is None:
            st.error("❌ Sesuatu yang tidak dijangka berlaku. Tiada output dihasilkan.")

        else:
            if result_count == 0:
                st.warning(
                    "⚠️ Laporan dijana tetapi **tiada pekerja dikesan** daripada PDF kehadiran. "
                    "Kemungkinan format PDF berbeza. Semak bahagian 🔍 Debug di bawah."
                )
            else:
                st.success(
                    f"✅ Berjaya menjana laporan untuk **{result_count}** pekerja "
                    f"daripada **{len(att_files)}** fail jabatan!"
                )

            st.download_button(
                label=f"📥 Muat Turun {result_name}",
                data=result_excel,
                file_name=result_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # --- Debug expander (always shown after processing) ---
        if debug_lines:
            with st.expander("🔍 Debug: 40 baris pertama PDF Kehadiran (untuk semak format)"):
                st.code("\n".join(debug_lines), language=None)