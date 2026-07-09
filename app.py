import streamlit as st
import sqlite3
import hashlib
from datetime import date
from io import BytesIO
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DB_PATH = "laptops.db"

ASSET_TYPES = ["Laptop", "i Pad", "Desktop", "Phone"]

LOCATIONS = [
    "Bird Automotive Pvt.Ltd 4IDC",
    "Bird Automotive Pvt.Ltd  Baani",
    "Bird Automotive Pvt.Ltd Service Factory ",
    "Bird Automotive Pvt.Ltd Mini Delhi",
    "Bird Automotive Pvt.Ltd Dehradun",
    "Bird Automobile Porsche Kolkata WS",
    "Bird Automobile Porsche Kolkata SR",
    "Bird Premium Selection Bentley SR",
    "Bird Premium Selection Bentley WS",
    "E-Next Mobility Citroen 48",
    "E-Next Mobility Sector-29 GGN",
]

# ---------------------------------------------------------
# DATABASE SETUP
# ---------------------------------------------------------

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'head_admin'))
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS laptops (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asset_type TEXT,
            model_name TEXT NOT NULL,
            serial_number TEXT,
            vendor_name TEXT,
            purchase_date TEXT,
            department TEXT,
            location TEXT,
            designation TEXT,
            current_holder TEXT,
            employee_code TEXT,
            employee_mobile TEXT,
            date_issued TEXT,
            reallocation_date TEXT,
            status TEXT DEFAULT 'Active',
            remarks TEXT
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            laptop_id INTEGER NOT NULL,
            person_name TEXT,
            department TEXT,
            location TEXT,
            designation TEXT,
            date_issued TEXT,
            date_ended TEXT,
            remarks TEXT,
            FOREIGN KEY(laptop_id) REFERENCES laptops(id)
        )
    """)

    conn.commit()

    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        c.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
            ("admin", hash_pw("admin123"), "admin"),
        )
        c.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
            ("headadmin", hash_pw("head123"), "head_admin"),
        )
        conn.commit()

    conn.close()


def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()


# ---------------------------------------------------------
# AUTH
# ---------------------------------------------------------

def check_login(username, password):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT password_hash, role FROM users WHERE username = ?", (username,))
    row = c.fetchone()
    conn.close()
    if row and row[0] == hash_pw(password):
        return row[1]
    return None


def login_screen():
    st.title("🗂️ Asset Management")
    st.subheader("Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        role = check_login(username, password)
        if role:
            st.session_state["logged_in"] = True
            st.session_state["username"] = username
            st.session_state["role"] = role
            st.rerun()
        else:
            st.error("Invalid username or password.")


# ---------------------------------------------------------
# DATA HELPERS
# ---------------------------------------------------------

def add_asset(asset_type, model_name, serial_number, vendor_name, purchase_date,
              department, location, designation, holder, employee_code, employee_mobile, date_issued, remarks):
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        INSERT INTO laptops (asset_type, model_name, serial_number, vendor_name, purchase_date,
                              department, location, designation, current_holder, employee_code, employee_mobile, date_issued, status, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active', ?)
    """, (asset_type, model_name, serial_number, vendor_name, str(purchase_date),
          department, location, designation, holder, employee_code, employee_mobile, str(date_issued), remarks))
    asset_id = c.lastrowid
    c.execute("""
        INSERT INTO history (laptop_id, person_name, department, location, designation, date_issued, date_ended, remarks)
        VALUES (?, ?, ?, ?, ?, ?, NULL, ?)
    """, (asset_id, holder, department, location, designation, str(date_issued), remarks))
    conn.commit()
    conn.close()


def reallocate_asset(asset_id, new_holder, new_department, new_location, new_designation,
                      reallocation_date, remarks):
    conn = get_conn()
    c = conn.cursor()

    c.execute("""
        UPDATE history SET date_ended = ?
        WHERE laptop_id = ? AND date_ended IS NULL
    """, (str(reallocation_date), asset_id))

    c.execute("""
        UPDATE laptops
        SET current_holder = ?, department = ?, location = ?, designation = ?,
            date_issued = ?, reallocation_date = ?, remarks = ?
        WHERE id = ?
    """, (new_holder, new_department, new_location, new_designation,
          str(reallocation_date), str(reallocation_date), remarks, asset_id))

    c.execute("""
        INSERT INTO history (laptop_id, person_name, department, location, designation, date_issued, date_ended, remarks)
        VALUES (?, ?, ?, ?, ?, ?, NULL, ?)
    """, (asset_id, new_holder, new_department, new_location, new_designation, str(reallocation_date), remarks))

    conn.commit()
    conn.close()


def get_assets():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM laptops ORDER BY id DESC")
    rows = c.fetchall()
    cols = [d[0] for d in c.description]
    conn.close()
    return cols, rows


def get_history(asset_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute("""SELECT person_name, department, location, designation, date_issued, date_ended, remarks
                 FROM history WHERE laptop_id = ? ORDER BY id DESC""", (asset_id,))
    rows = c.fetchall()
    conn.close()
    return rows


def delete_asset(asset_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM history WHERE laptop_id = ?", (asset_id,))
    c.execute("DELETE FROM laptops WHERE id = ?", (asset_id,))
    conn.commit()
    conn.close()


def add_user(username, password, role):
    conn = get_conn()
    c = conn.cursor()
    try:
        c.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                  (username, hash_pw(password), role))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def change_password(username, old_password, new_password):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT password_hash FROM users WHERE username = ?", (username,))
    row = c.fetchone()
    if not row or row[0] != hash_pw(old_password):
        conn.close()
        return False
    c.execute("UPDATE users SET password_hash = ? WHERE username = ?",
              (hash_pw(new_password), username))
    conn.commit()
    conn.close()
    return True


# ---------------------------------------------------------
# EXPORT FUNCTIONS
# ---------------------------------------------------------

def export_to_pdf(filtered_data, columns):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    elements = []
    style = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=style['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#1f2a4a'),
        spaceAfter=12,
    )
    elements.append(Paragraph("Asset Management Report", title_style))
    elements.append(Spacer(1, 0.3*inch))
    
    table_data = [[col.upper() for col in columns]]
    for row in filtered_data:
        table_data.append([str(val) if val else "-" for val in row])
    
    table = Table(table_data, colWidths=[1.2*inch]*len(columns))
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3b6ef0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4ff')]),
    ]))
    
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


def export_to_excel(filtered_data, columns):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Assets"
    
    header_fill = PatternFill(start_color="3b6ef0", end_color="3b6ef0", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for col_num, col in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = col.upper()
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    
    for row_num, row_data in enumerate(filtered_data, 2):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.value = value if value else "-"
            cell.border = border
            cell.alignment = Alignment(horizontal="left", vertical="center")
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width
    
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def export_to_word(filtered_data, columns):
    doc = Document()
    doc.add_heading("Asset Management Report", 0)
    doc.add_paragraph()
    
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = 'Light Grid Accent 1'
    
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(columns):
        hdr_cells[i].text = col.upper()
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for row_data in filtered_data:
        row_cells = table.add_row().cells
        for i, value in enumerate(row_data):
            row_cells[i].text = str(value) if value else "-"
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# MAIN APP
# ---------------------------------------------------------

def main_app():
    role = st.session_state["role"]
    username = st.session_state["username"]

    st.sidebar.title("🗂️ Asset Management")
    st.sidebar.write(f"Logged in as **{username}** ({role.replace('_', ' ').title()})")
    if st.sidebar.button("Logout"):
        for k in ["logged_in", "username", "role"]:
            st.session_state.pop(k, None)
        st.rerun()

    pages = ["Dashboard", "Add Asset", "Reallocate Asset", "Change Password", "Reports"]
    if role == "head_admin":
        pages.append("Manage Users")

    page = st.sidebar.radio("Navigate", pages)

    st.title("🗂️ Asset Management")

    if page == "Dashboard":
        show_dashboard(role)
    elif page == "Add Asset":
        show_add_asset()
    elif page == "Reallocate Asset":
        show_reallocate()
    elif page == "Change Password":
        show_change_password()
    elif page == "Reports":
        show_reports(role)
    elif page == "Manage Users":
        show_manage_users()


def show_dashboard(role):
    st.header("📋 Asset Inventory")

    cols, rows = get_assets()
    if not rows:
        st.info("No assets added yet. Go to 'Add Asset' to get started.")
        return

    search_query = st.text_input("🔍 Search (model name, serial number, department, holder name, vendor name, issued date or location)")

    col1, col2, col3 = st.columns(3)
    dept_idx = cols.index("department")
    loc_idx = cols.index("location")
    type_idx = cols.index("asset_type")

    departments = sorted(set(r[dept_idx] for r in rows if r[dept_idx]))
    locations = sorted(set(r[loc_idx] for r in rows if r[loc_idx]))
    types = sorted(set(r[type_idx] for r in rows if r[type_idx]))

    with col1:
        dept_filter = st.selectbox("Filter by Department", ["All"] + departments)
    with col2:
        loc_filter = st.selectbox("Filter by Location", ["All"] + locations)
    with col3:
        type_filter = st.selectbox("Filter by Asset Type", ["All"] + types)

    filtered = rows
    if dept_filter != "All":
        filtered = [r for r in filtered if r[dept_idx] == dept_filter]
    if loc_filter != "All":
        filtered = [r for r in filtered if r[loc_idx] == loc_filter]
    if type_filter != "All":
        filtered = [r for r in filtered if r[type_idx] == type_filter]

    if search_query.strip():
        q = search_query.strip().lower()
        search_fields = ["model_name", "serial_number", "department", "current_holder", "vendor_name", "date_issued", "location"]
        search_idxs = [cols.index(f) for f in search_fields]
        filtered = [
            r for r in filtered
            if any(q in str(r[i]).lower() for i in search_idxs if r[i])
        ]

    for r in filtered:
        row = dict(zip(cols, r))
        label = f"🗂️ [{row['asset_type'] or 'Asset'}] {row['model_name']} — held by {row['current_holder'] or 'Unassigned'}"
        with st.expander(label):
            c1, c2 = st.columns(2)
            with c1:
                st.write(f"**Asset Type:** {row['asset_type']}")
                st.write(f"**Serial Number:** {row['serial_number']}")
                st.write(f"**Vendor:** {row['vendor_name']}")
                st.write(f"**Purchase Date:** {row['purchase_date']}")
                st.write(f"**Department:** {row['department']}")
                st.write(f"**Location:** {row['location']}")
            with c2:
                st.write(f"**Designation:** {row['designation']}")
                st.write(f"**Current Holder:** {row['current_holder']}")
                st.write(f"**Employee Code:** {row['employee_code']}")
                st.write(f"**Employee Mobile:** {row['employee_mobile']}")
                st.write(f"**Date Issued:** {row['date_issued']}")
                st.write(f"**Status:** {row['status']}")
                if role == "head_admin":
                    st.write(f"**Reallocation Date:** {row['reallocation_date'] or '-'}")

            if row.get("remarks"):
                st.write(f"**Remarks:** {row['remarks']}")

            if role == "head_admin":
                st.markdown("**📜 Assignment History**")
                hist = get_history(row["id"])
                if hist:
                    for h in hist:
                        person, department, location, designation, issued, ended, remarks = h
                        status = f"until {ended}" if ended else "→ current"
                        line = f"- {person} ({designation}, {department}, {location}) — {issued} {status}"
                        if remarks:
                            line += f" | remark: {remarks}"
                        st.write(line)
                else:
                    st.write("No history yet.")

                if st.button("🗑️ Delete this asset record", key=f"del_{row['id']}"):
                    delete_asset(row["id"])
                    st.success("Deleted.")
                    st.rerun()


def show_add_asset():
    st.header("➕ Add New Asset")
    with st.form("add_asset_form"):
        asset_type = st.selectbox("Asset Type", ASSET_TYPES)
        model_name = st.text_input("Model Name")
        serial_number = st.text_input("Asset Serial Number")
        vendor_name = st.text_input("Vendor Name (where purchased)")
        purchase_date = st.date_input("Purchase Date", value=date.today())

        department = st.text_input("Department")
        location = st.selectbox("Location", LOCATIONS)
        designation = st.text_input("Designation of Holder")
        holder = st.text_input("Name of Person")
        employee_code = st.text_input("Employee Code")
        employee_mobile = st.text_input("Employee Mobile Number")
        date_issued = st.date_input("Date Issued", value=date.today())
        remarks = st.text_area("Remarks (optional)")

        submitted = st.form_submit_button("Add Asset")
        if submitted:
            if not model_name:
                st.error("Model name is required.")
            else:
                add_asset(asset_type, model_name, serial_number, vendor_name, purchase_date,
                          department, location, designation, holder, employee_code, employee_mobile, date_issued, remarks)
                st.success(f"Asset '{model_name}' added successfully.")


def show_reallocate():
    st.header("🔄 Reallocate Asset")
    cols, rows = get_assets()
    if not rows:
        st.info("No assets available. Add one first.")
        return

    options = {}
    for r in rows:
        row = dict(zip(cols, r))
        label = f"[{row['asset_type']}] {row['model_name']} (ID {row['id']}) — currently: {row['current_holder']} @ {row['location']}"
        options[label] = row

    selected_label = st.selectbox("Select Asset", list(options.keys()))
    selected_row = options[selected_label]
    asset_id = selected_row["id"]
    current_location = selected_row["location"]

    asset_action = st.radio("What do you want to do?", ["Reallocate to Person", "Mark as Scrap"])

    if asset_action == "Mark as Scrap":
        if st.button("Mark as Scrap"):
            conn = get_conn()
            c = conn.cursor()
            c.execute("UPDATE laptops SET status = 'Scrap', current_holder = NULL WHERE id = ?", (asset_id,))
            conn.commit()
            conn.close()
            st.success("Asset marked as Scrap.")
            st.rerun()
    else:
        with st.form("reallocate_form"):
            new_holder = st.text_input("New Holder's Name")
            new_department = st.text_input("New Department")
            new_employee_code = st.text_input("New Employee Code")
            new_employee_mobile = st.text_input("New Employee Mobile Number")

            location_options = LOCATIONS.copy()
            if current_location and current_location not in location_options:
                location_options.append(current_location)
            default_index = location_options.index(current_location) if current_location in location_options else 0
            new_location = st.selectbox("New Location", location_options, index=default_index)

            new_designation = st.text_input("New Designation")
            reallocation_date = st.date_input("Reallocation Date", value=date.today())
            remarks = st.text_area("Remarks (optional)")

            submitted = st.form_submit_button("Reallocate")
            if submitted:
                if not new_holder:
                    st.error("New holder's name is required.")
                else:
                    conn = get_conn()
                    c = conn.cursor()
                    c.execute("""
                        UPDATE history SET date_ended = ?
                        WHERE laptop_id = ? AND date_ended IS NULL
                    """, (str(reallocation_date), asset_id))
                    c.execute("""
                        UPDATE laptops
                        SET current_holder = ?, department = ?, location = ?, designation = ?,
                            employee_code = ?, employee_mobile = ?, date_issued = ?, reallocation_date = ?, remarks = ?
                        WHERE id = ?
                    """, (new_holder, new_department, new_location, new_designation,
                          new_employee_code, new_employee_mobile, str(reallocation_date), str(reallocation_date), remarks, asset_id))
                    c.execute("""
                        INSERT INTO history (laptop_id, person_name, department, location, designation, date_issued, date_ended, remarks)
                        VALUES (?, ?, ?, ?, ?, ?, NULL, ?)
                    """, (asset_id, new_holder, new_department, new_location, new_designation, str(reallocation_date), remarks))
                    conn.commit()
                    conn.close()
                    st.success("Asset reallocated successfully. Previous assignment saved to history.")
                    st.rerun()


def show_change_password():
    st.header("🔑 Change My Password")
    username = st.session_state["username"]

    with st.form("change_password_form"):
        old_password = st.text_input("Current Password", type="password")
        new_password = st.text_input("New Password", type="password")
        confirm_password = st.text_input("Confirm New Password", type="password")

        submitted = st.form_submit_button("Update Password")
        if submitted:
            if not old_password or not new_password:
                st.error("All fields are required.")
            elif new_password != confirm_password:
                st.error("New password and confirmation do not match.")
            else:
                ok = change_password(username, old_password, new_password)
                if ok:
                    st.success("Password updated successfully. Use your new password next time you log in.")
                else:
                    st.error("Current password is incorrect.")


def show_reports(role):
    st.header("📊 Asset Reports")
    
    cols, rows = get_assets()
    if not rows:
        st.info("No assets to report on yet.")
        return

    st.subheader("Filter & Search")
    
    search_query = st.text_input("🔍 Search by vendor, holder, department, location, serial number, or date")
    
    col1, col2, col3 = st.columns(3)
    dept_idx = cols.index("department")
    loc_idx = cols.index("location")
    type_idx = cols.index("asset_type")
    vendor_idx = cols.index("vendor_name")
    holder_idx = cols.index("current_holder")

    departments = sorted(set(r[dept_idx] for r in rows if r[dept_idx]))
    locations = sorted(set(r[loc_idx] for r in rows if r[loc_idx]))
    vendors = sorted(set(r[vendor_idx] for r in rows if r[vendor_idx]))
    types = sorted(set(r[type_idx] for r in rows if r[type_idx]))

    with col1:
        dept_filter = st.selectbox("Filter by Department", ["All"] + departments, key="report_dept")
    with col2:
        type_filter = st.selectbox("Filter by Asset Type", ["All"] + types, key="report_type")
    with col3:
        vendor_filter = st.selectbox("Filter by Vendor", ["All"] + vendors, key="report_vendor")

    col4, col5 = st.columns(2)
    with col4:
        loc_filter = st.selectbox("Filter by Location", ["All"] + locations, key="report_loc")
    with col5:
        status_filter = st.selectbox("Filter by Status", ["All", "Active", "Inactive"], key="report_status")

    filtered = rows
    if dept_filter != "All":
        filtered = [r for r in filtered if r[dept_idx] == dept_filter]
    if type_filter != "All":
        filtered = [r for r in filtered if r[type_idx] == type_filter]
    if vendor_filter != "All":
        filtered = [r for r in filtered if r[vendor_idx] == vendor_filter]
    if loc_filter != "All":
        filtered = [r for r in filtered if r[loc_idx] == loc_filter]
    if status_filter != "All":
        status_idx = cols.index("status")
        filtered = [r for r in filtered if r[status_idx] == status_filter]

    if search_query.strip():
        q = search_query.strip().lower()
        search_fields = ["vendor_name", "current_holder", "department", "location", "serial_number", "date_issued", "reallocation_date", "status", "employee_code", "employee_mobile"]
        search_idxs = [cols.index(f) for f in search_fields if f in cols]
        filtered = [
            r for r in filtered
            if any(q in str(r[i]).lower() for i in search_idxs if i < len(r) and r[i])
        ]

    if filtered:
        st.write(f"**Total Assets in Report:** {len(filtered)}")
        st.markdown("---")

        st.subheader("Export Options")
        export_cols = [
            "id", "asset_type", "model_name", "serial_number", "vendor_name",
            "purchase_date", "department", "location", "designation", "current_holder",
            "employee_code", "employee_mobile", "date_issued", "reallocation_date", "status", "remarks"
        ]
        export_col_idxs = [cols.index(c) for c in export_cols if c in cols]
        export_data = [[r[i] for i in export_col_idxs] for r in filtered]
        export_cols_filtered = [export_cols[i] for i in range(len(export_cols)) if export_cols[i] in cols]

        col_pdf, col_excel, col_word = st.columns(3)

        with col_pdf:
            pdf_buffer = export_to_pdf(export_data, export_cols_filtered)
            st.download_button(
                label="📄 Download as PDF",
                data=pdf_buffer,
                file_name="Asset_Report.pdf",
                mime="application/pdf"
            )

        with col_excel:
            excel_buffer = export_to_excel(export_data, export_cols_filtered)
            st.download_button(
                label="📊 Download as Excel",
                data=excel_buffer,
                file_name="Asset_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_word:
            word_buffer = export_to_word(export_data, export_cols_filtered)
            st.download_button(
                label="📝 Download as Word",
                data=word_buffer,
                file_name="Asset_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        st.markdown("---")
        st.subheader("Preview")
        st.dataframe(export_data, column_config={
            col: st.column_config.TextColumn(col.upper()) for col in export_cols_filtered
        }, use_container_width=True)
    else:
        st.warning("No assets match your filters.")


def show_manage_users():
    st.header("👤 Manage Users (Head Admin only)")
    st.write("Create new Admin or Head Admin login accounts.")

    with st.form("add_user_form"):
        new_username = st.text_input("New Username")
        new_password = st.text_input("New Password", type="password")
        new_role = st.selectbox("Role", ["admin", "head_admin"])

        submitted = st.form_submit_button("Create User")
        if submitted:
            if not new_username or not new_password:
                st.error("Username and password are required.")
            else:
                ok = add_user(new_username, new_password, new_role)
                if ok:
                    st.success(f"User '{new_username}' created as {new_role}.")
                else:
                    st.error("Username already exists.")


# ---------------------------------------------------------
# THEME
# ---------------------------------------------------------

def apply_theme():
    st.markdown("""
        <style>
        .stApp {
            background: linear-gradient(160deg, #dbe9ff 0%, #eef2ff 45%, #fef6e8 100%);
        }
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #2b3a67 0%, #1f2a4a 100%);
        }
        section[data-testid="stSidebar"] * {
            color: #f5f7fa !important;
        }
        h1, h2, h3 {
            color: #1f2a4a;
        }
        div.stButton > button {
            background-color: #3b6ef0;
            color: white;
            border-radius: 8px;
            border: none;
        }
        div.stButton > button:hover {
            background-color: #2b53c4;
            color: white;
        }
        div[data-testid="stExpander"] {
            background-color: #ffffff;
            border-radius: 10px;
            border: 1px solid #dde3ee;
        }
        div[data-testid="stForm"] {
            background-color: #ffffff;
            padding: 1.2rem;
            border-radius: 10px;
            border: 1px solid #dde3ee;
        }
        </style>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------

def main():
    st.set_page_config(page_title="Asset Management", page_icon="🗂️", layout="wide")
    apply_theme()
    init_db()

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if not st.session_state["logged_in"]:
        login_screen()
    else:
        main_app()


if __name__ == "__main__":
    main()
