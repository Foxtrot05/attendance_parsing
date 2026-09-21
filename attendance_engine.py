import io
import re
import calendar
from datetime import datetime, time
import pypdf
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def normalize_name(name):
    if not name:
        return ""
    name = name.upper()
    name = re.sub(r"\b(TS|GS|SR)\b\.?", "", name)
    name = re.sub(r"\b(BIN|BINTI|BTE|BT|A/L|A/P)\b", "", name)
    name = re.sub(r"[^A-Z\s]", " ", name)
    return " ".join(name.split())


# Mapping postcode → nama bandar ringkas
POSTCODE_CITY_MAP = {
    "71800": "NILAI",    "71700": "MANTIN",   "71450": "REMBAU",
    "43500": "SEMENYIH", "43000": "KAJANG",   "43650": "BANGI",
    "50000": "KL",       "50050": "KL",       "50460": "KL",
    "50480": "KL",       "50088": "KL",
    "47810": "PJ",       "40150": "SHAH ALAM",
    "68000": "AMPANG",   "68100": "BATU CAVES","41000": "KLANG",
    "71000": "PORT DICKSON", "72000": "KUALA PILAH",
    "43700": "BERANANG", "63000": "CYBERJAYA","62000": "PUTRAJAYA",
    "62502": "PUTRAJAYA","62300": "PUTRAJAYA",
}

def clean_location(loc_str):
    if not loc_str or loc_str.strip().lower() in ["no record", "", "look good", "unknown"]:
        return "UNKNOWN"
    loc_lower = loc_str.lower()
    # MAP2U office — Jalan BBN / BBN
    if "jalan bbn" in loc_lower or "jln bbn" in loc_lower or "bbn" in loc_lower:
        return "MAP2U"
    # Buang Google Plus Code (cth: RQGX+2J) dari awal string
    loc_clean = re.sub(r'^[A-Z0-9]{4,8}\+[A-Z0-9]{2,4}\s*,?\s*', '', loc_str, flags=re.IGNORECASE).strip()
    # Cuba cari postcode dan map ke nama bandar
    postcode_match = re.search(r'\b(\d{5})\b', loc_clean)
    if postcode_match:
        pc = postcode_match.group(1)
        if pc in POSTCODE_CITY_MAP:
            return POSTCODE_CITY_MAP[pc]
        # Ekstrak nama bandar selepas postcode
        after_pc = loc_clean[postcode_match.end():].strip().lstrip(',').strip()
        city_match = re.match(r'([A-Za-z][A-Za-z\s]+?)(?:,|$)', after_pc)
        if city_match:
            city = city_match.group(1).strip().split()[0].upper()
            skip = {"MALAYSIA", "SELANGOR", "NEGERI", "MELAKA", "PAHANG",
                    "JOHOR", "PERAK", "SEMBILAN", "PERSEKUTUAN", "WILAYAH"}
            if city and city not in skip:
                return city
    # Fallback: ambil bahagian sebelum negeri/Malaysia
    parts = [p.strip() for p in loc_clean.split(',') if p.strip()]
    skip_words = {"malaysia", "selangor", "negeri sembilan", "melaka", "pahang",
                  "johor", "perak", "kuala lumpur", "wilayah persekutuan"}
    for part in reversed(parts):
        if part.lower() not in skip_words and not re.match(r'^\d+$', part):
            tok = part.strip().split()[0].upper()
            if len(tok) > 2:
                return tok
    return parts[0].strip().split()[0].upper() if parts else "UNKNOWN"


def parse_time(time_str):
    try:
        return datetime.strptime(time_str.strip(), "%I:%M %p").time()
    except Exception:
        return None


def fmt_time(ts):
    try:
        t_obj = datetime.strptime(ts.strip(), "%I:%M %p")
        return t_obj.strftime("%-I:%M %p")   # Linux: %-I buang leading zero
    except Exception:
        return ts.strip()


def determine_bracket_and_early(t_in, t_out):
    if not t_in or not t_out:
        return False
    if t_in <= time(8, 0):
        return t_out < time(16, 30)
    elif t_in <= time(8, 30):
        return t_out < time(17, 30)
    elif t_in <= time(9, 0):
        return t_out < time(18, 0)
    else:
        return t_out < time(18, 0)


def get_status_label(count):
    if not count or count == 0:
        return ""
    elif 1 <= count <= 2:
        return "GOOD"
    elif 3 <= count <= 6:
        return "MIDDLE"
    else:
        return "CRITICAL"


# ---------------------------------------------------------------------------
# PDF Leave Parser
# ---------------------------------------------------------------------------

def parse_all_department_leaves(pdf_bytes):
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    text = "\n".join([page.extract_text() for page in reader.pages])
    leave_records = []
    pattern = re.compile(
        r'(\d+)\s+([A-Z\s\.\@\?]+?)\s+(Annual Leave|Replacement Leave|Medical Leave|'
        r'Compassionate Leave[^\n]*|Special Leave[^\n]*|Cuti [^\n]*)\s+'
        r'(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+([\d\.]+)\s+.*?(Approved|Cancelled)',
        re.IGNORECASE
    )
    for match in pattern.finditer(text):
        _, name, leave_type, start_d, end_d, days, status = match.groups()
        if status.strip().lower() == "approved":
            leave_records.append({
                "norm_name":  normalize_name(name),
                "type":       leave_type.strip(),
                "start_date": datetime.strptime(start_d, "%d/%m/%Y"),
                "end_date":   datetime.strptime(end_d,   "%d/%m/%Y"),
                "days":       float(days)
            })
    return leave_records


# ---------------------------------------------------------------------------
# PDF Time-Off Parser
# ---------------------------------------------------------------------------

def parse_multiple_timeoff_records(timeoff_bytes_list):
    timeoff_records = []
    pattern = re.compile(
        r'([A-Z\s]+?)\s+M2U\d+\s+[A-Z\s]+\s+(\d{2}/\d{2}/\d{4})\s+'
        r'(\d{1,2}:\d{2}\s*[AP]M)\s+(\d{1,2}:\d{2}\s*[AP]M)\s+.*?Approved',
        re.IGNORECASE
    )
    for pdf_b in timeoff_bytes_list:
        reader = pypdf.PdfReader(io.BytesIO(pdf_b))
        text = "\n".join([page.extract_text() for page in reader.pages])
        for match in pattern.finditer(text):
            name, date_str, s_time, e_time = match.groups()
            timeoff_records.append({
                "norm_name": normalize_name(name),
                "date":      datetime.strptime(date_str, "%d/%m/%Y").date(),
                "time_str":  f"TIME SLIP {s_time.replace(' ', '')}-{e_time.replace(' ', '')}".replace(":00", "")
            })
    return timeoff_records


# ---------------------------------------------------------------------------
# Core: parse ONE department PDF → list of employee dicts
# ---------------------------------------------------------------------------

# Regex pattern tetap (compile sekali)
_RE_DATE        = re.compile(r'^(\d{2}/\d{2}/\d{4})$')
_RE_TIME        = re.compile(r'^(\d{2}:\d{2}\s*[AP]M)')
_RE_STATUS_FULL = re.compile(r'^\(.*\)$', re.IGNORECASE)
_RE_STATUS_CONT = re.compile(r'^(Good|Time|Early)\)$', re.IGNORECASE)
_RE_HOURS       = re.compile(r'^\d{2}:\d{2}:\d{2}\s+\d+:\d{2}:\d{2}\s+\d+:\d{2}:\d{2}$')
_RE_ROW_DEFAULT = re.compile(r'^(\d{1,2})\s+Default$')
_RE_ROW_FULL    = re.compile(r'^(\d{1,2})\s+Default\s+Shift$')
_RE_EMP_NAME    = re.compile(r'Employee Name\s*[:\s|]\s*(.+)', re.IGNORECASE)
_RE_SKIP        = re.compile(
    r'^(NO\s+SHIFT|ATTENDANCE\s+TIMESHEET|SHIFT\s+DATETIME|BREAK\s+HOUR|'
    r'Date\s+Generated|Name\s*:|Division$)',
    re.IGNORECASE
)


def _flush_record(cur, days_record):
    """Simpan rekod attendance semasa ke dict days_record."""
    try:
        d_in  = datetime.strptime(cur["d_in"],  "%d/%m/%Y").date()
        d_out = datetime.strptime(cur["d_out"], "%d/%m/%Y").date()
        t_in  = parse_time(cur["t_in"])
        t_out = parse_time(cur["t_out"])
        days_record[d_in.day] = {
            "t_in_str":    fmt_time(cur["t_in"]),
            "loc_in":      clean_location(cur["loc_in"]),
            "t_out_str":   fmt_time(cur["t_out"]),
            "loc_out":     clean_location(cur["loc_out"]),
            "is_late":     t_in > time(9, 0) if t_in else False,
            "is_overnight": d_out > d_in,
            "is_early":    determine_bracket_and_early(t_in, t_out),
            "d_in":  d_in,
            "d_out": d_out,
        }
    except Exception:
        pass


def _run_state_machine(lines, raw_name):
    """
    Jalankan state machine pada senarai baris untuk satu pekerja.
    Return dict: {norm_name, raw_name, records, notes, timeoff_notes}
    """
    days_record   = {}
    notes_dict    = {}
    timeoff_notes = {}

    state = "IDLE"
    cur   = {}

    for line in lines:
        if not line:
            continue
        if _RE_SKIP.match(line):
            continue
        # Skip sambungan status yang terputus
        if _RE_STATUS_CONT.match(line):
            continue

        # ---- IDLE: tunggu baris nombor baris attendance ----
        if state == "IDLE":
            m = _RE_ROW_FULL.match(line)
            if m:
                cur   = {"no": m.group(1)}
                state = "WAIT_DATE_IN"
                continue
            m2 = _RE_ROW_DEFAULT.match(line)
            if m2:
                cur   = {"no": m2.group(1)}
                state = "WAIT_SHIFT"
                continue

        # ---- Tunggu "Shift" di baris berikutnya ----
        elif state == "WAIT_SHIFT":
            if re.match(r'^Shift$', line, re.IGNORECASE):
                state = "WAIT_DATE_IN"
            else:
                state = "IDLE"
            continue

        # ---- Tunggu tarikh check-in ----
        elif state == "WAIT_DATE_IN":
            m = _RE_DATE.match(line)
            if m:
                cur["d_in"] = m.group(1)
                state = "WAIT_TIME_IN"
            continue

        # ---- Tunggu masa check-in ----
        elif state == "WAIT_TIME_IN":
            m = _RE_TIME.match(line)
            if m:
                cur["t_in"]  = m.group(1).strip()
                cur["loc_in"] = ""
                state = "WAIT_STATUS_IN"
            continue

        # ---- Skip status check-in, mula collect lokasi ----
        elif state == "WAIT_STATUS_IN":
            if _RE_STATUS_FULL.match(line):
                state = "COLLECT_LOC_IN"
            elif _RE_DATE.match(line):
                # Tiada status, terus ke tarikh checkout
                cur["d_out"]  = line
                cur["loc_out"] = ""
                state = "WAIT_TIME_OUT"
            else:
                cur["loc_in"] = line
                state = "COLLECT_LOC_IN"
            continue

        # ---- Kumpul lokasi check-in (multi-line) ----
        elif state == "COLLECT_LOC_IN":
            m = _RE_DATE.match(line)
            if m:
                cur["d_out"]  = m.group(1)
                cur["loc_out"] = ""
                state = "WAIT_TIME_OUT"
            else:
                cur["loc_in"] = (cur["loc_in"] + " " + line).strip()
            continue

        # ---- Tunggu masa check-out ----
        elif state == "WAIT_TIME_OUT":
            m = _RE_TIME.match(line)
            if m:
                cur["t_out"]  = m.group(1).strip()
                cur["loc_out"] = ""
                state = "WAIT_STATUS_OUT"
            continue

        # ---- Skip status check-out, mula collect lokasi ----
        elif state == "WAIT_STATUS_OUT":
            if _RE_STATUS_FULL.match(line):
                state = "COLLECT_LOC_OUT"
            elif _RE_HOURS.match(line):
                _flush_record(cur, days_record)
                state = "IDLE"
            else:
                # Baris ni mungkin lokasi, atau separuh status "(Look" — skip yang separuh
                if not line.endswith("(Look") and not line.endswith("(On"):
                    cur["loc_out"] = line
                state = "COLLECT_LOC_OUT"
            continue

        # ---- Kumpul lokasi check-out (multi-line) ----
        elif state == "COLLECT_LOC_OUT":
            if _RE_HOURS.match(line):
                _flush_record(cur, days_record)
                state = "IDLE"
            else:
                cur["loc_out"] = (cur["loc_out"] + " " + line).strip()
            continue

    # Flush rekod terakhir jika ada
    if state == "COLLECT_LOC_OUT" and "d_in" in cur and "d_out" in cur:
        _flush_record(cur, days_record)

    return {
        "norm_name":    normalize_name(raw_name),
        "raw_name":     raw_name,
        "records":      days_record,
        "notes":        notes_dict,
        "timeoff_notes": timeoff_notes,
    }


def parse_department_pdf(pdf_bytes):
    """
    Parse satu PDF jabatan yang mengandungi rekod SATU atau RAMAI pekerja.
    Return: list of employee dicts
    """
    reader    = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    full_text = "\n".join([page.extract_text() for page in reader.pages])
    all_lines = [l.strip() for l in full_text.split("\n")]

    # Split baris kepada segmen per pekerja menggunakan penanda "Employee Name"
    segments = []   # list of (raw_name, [lines])
    cur_name  = None
    cur_lines = []

    for line in all_lines:
        m = _RE_EMP_NAME.match(line)
        if m:
            # Simpan segmen pekerja sebelumnya
            if cur_name is not None:
                segments.append((cur_name, cur_lines))
            raw = m.group(1).strip()
            if "department" in raw.lower():
                raw = raw.split("Department")[0].strip()
            cur_name  = raw
            cur_lines = []
        else:
            if cur_name is not None:
                cur_lines.append(line)

    # Simpan segmen terakhir
    if cur_name is not None:
        segments.append((cur_name, cur_lines))

    # Parse setiap segmen
    employees = []
    for raw_name, seg_lines in segments:
        emp = _run_state_machine(seg_lines, raw_name)
        employees.append(emp)

    return employees


# Alias untuk keserasian ke belakang (jika kod lain guna parse_single)
def parse_single_individual_pdf(pdf_bytes):
    """Wrapper — kembalikan rekod pekerja pertama dalam PDF."""
    emps = parse_department_pdf(pdf_bytes)
    return emps[0] if emps else {
        "norm_name": "UNKNOWN", "raw_name": "UNKNOWN",
        "records": {}, "notes": {}, "timeoff_notes": {}
    }


# ---------------------------------------------------------------------------
# Excel Generator
# ---------------------------------------------------------------------------

def generate_reconciled_excel(attendance_pdf_list, leave_pdf_bytes,
                               timeoff_pdf_list=None, year=2025, month=12):
    # 1. Parse Data
    employees_list = []
    for pdf_b in attendance_pdf_list:
        employees_list.extend(parse_department_pdf(pdf_b))

    leave_list   = parse_all_department_leaves(leave_pdf_bytes)
    timeoff_list = parse_multiple_timeoff_records(timeoff_pdf_list) if timeoff_pdf_list else []

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reconciliation"

    num_days = calendar.monthrange(year, month)[1]
    headers  = ["DEPARTMENT/ NAME"] + list(range(1, num_days + 1))
    headers += [
        'WFH', 'LEAVE', 'MC', 'TIMEOFF', 'UNPAID LEAVE', 'PUBLIC LEAVE',
        'NOT CLOCK IN', 'STATUS', 'NOT CLOCKOUT', 'STATUS',
        'EARLY CLOCK OUT', 'STATUS', 'LATE (DAYS)', 'STATUS', 'NOTES'
    ]
    ws.append(headers)

    fill_weekend = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
    fill_cuti    = PatternFill(start_color="FF34A853", end_color="FF34A853", fill_type="solid")
    fill_normal  = PatternFill(start_color="FF00FFFF", end_color="FF00FFFF", fill_type="solid")
    fill_red     = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    fill_orange  = PatternFill(start_color="FFFF6D01", end_color="FFFF6D01", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin',   color='D3D3D3'),
        right=Side(style='thin',  color='D3D3D3'),
        top=Side(style='thin',    color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    for col_num in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_num)
        c.font      = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 2
    for emp_info in employees_list:
        norm_emp  = emp_info["norm_name"]
        row_cells = [None] * len(headers)
        row_cells[0] = emp_info["raw_name"]

        cnt_wfh, cnt_leave, cnt_mc, cnt_timeoff = 0, 0.0, 0.0, 0
        cnt_not_in, cnt_not_out, cnt_early, cnt_late = 0, 0, 0, 0

        staf_leaves   = [l for l in leave_list
                         if l["norm_name"] in norm_emp or norm_emp in l["norm_name"]]
        staf_timeoffs = {to["date"]: to["time_str"] for to in timeoff_list
                         if to["norm_name"] in norm_emp or norm_emp in to["norm_name"]}

        for day in range(1, num_days + 1):
            curr_date = datetime(year, month, day).date()
            weekday   = curr_date.weekday()
            col_idx   = day + 1

            if weekday in [5, 6] or (month == 12 and day == 25):
                ws.cell(row=row_idx, column=col_idx).fill = fill_weekend
                continue

            # --- Semak cuti ---
            is_on_leave = False
            for lv in staf_leaves:
                if lv["start_date"].date() <= curr_date <= lv["end_date"].date():
                    is_on_leave = True
                    if "medical" in lv["type"].lower():
                        cnt_mc += 1
                        row_cells[day] = "MC"
                    else:
                        cnt_leave += 1 if lv["days"] >= 1 else 0.5
                        row_cells[day] = "CUTI"
                    ws.cell(row=row_idx, column=col_idx).fill = fill_cuti
                    break
            if is_on_leave:
                continue

            rec       = emp_info["records"].get(day)
            note      = emp_info["notes"].get(curr_date, "")
            timeoff_str = staf_timeoffs.get(curr_date) or emp_info["timeoff_notes"].get(curr_date)

            if not rec:
                row_cells[day] = "TIADA REKOD"
                cnt_not_in    += 1
                ws.cell(row=row_idx, column=col_idx).fill = fill_red
                continue

            cell_text_parts = []
            if "wfh" in note.lower():
                cell_text_parts.append("WFH")
                cnt_wfh += 1
            if timeoff_str:
                cell_text_parts.append(timeoff_str)
                cnt_timeoff += 1

            cell_text_parts.append(f"C/I: {rec['t_in_str']} ({rec['loc_in']})")

            if "lupa" in note.lower() and "out" in note.lower():
                cell_text_parts.append(f"LUPA C/O: {rec['t_out_str']} ({rec['loc_out']})")
                cnt_not_out += 1
            elif rec["is_overnight"]:
                cell_text_parts.append(
                    f"C/O: {rec['t_out_str']} {rec['d_out'].day}HB ({rec['loc_out']})"
                )
            else:
                cell_text_parts.append(f"C/O: {rec['t_out_str']} ({rec['loc_out']})")

            row_cells[day] = "\n".join(cell_text_parts)

            target_cell = ws.cell(row=row_idx, column=col_idx)
            if rec["is_late"] or rec["is_overnight"] or rec["is_early"]:
                target_cell.fill = fill_red
                if rec["is_late"]:
                    cnt_late += 1
                if rec["is_early"]:
                    cnt_early += 1
            elif timeoff_str:
                target_cell.fill = fill_orange
            else:
                target_cell.fill = fill_normal

        # --- Metrik Rumusan ---
        m_idx = num_days + 1
        row_cells[m_idx]      = cnt_wfh     if cnt_wfh     > 0 else None
        row_cells[m_idx + 1]  = cnt_leave   if cnt_leave   > 0 else None
        row_cells[m_idx + 2]  = cnt_mc      if cnt_mc      > 0 else None
        row_cells[m_idx + 3]  = cnt_timeoff if cnt_timeoff > 0 else None
        row_cells[m_idx + 4]  = None
        row_cells[m_idx + 5]  = None
        row_cells[m_idx + 6]  = cnt_not_in  if cnt_not_in  > 0 else None
        row_cells[m_idx + 7]  = get_status_label(cnt_not_in)
        row_cells[m_idx + 8]  = cnt_not_out if cnt_not_out > 0 else None
        row_cells[m_idx + 9]  = get_status_label(cnt_not_out)
        row_cells[m_idx + 10] = cnt_early   if cnt_early   > 0 else None
        row_cells[m_idx + 11] = get_status_label(cnt_early)
        row_cells[m_idx + 12] = cnt_late    if cnt_late    > 0 else None
        row_cells[m_idx + 13] = get_status_label(cnt_late)

        for c_idx, val in enumerate(row_cells, 1):
            cell           = ws.cell(row=row_idx, column=c_idx)
            cell.value     = val
            cell.border    = thin_border
            cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

        row_idx += 1

    ws.column_dimensions['A'].width = 30
    for d in range(1, num_days + 1):
        col_letter = openpyxl.utils.get_column_letter(d + 1)
        ws.column_dimensions[col_letter].width = 24

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()