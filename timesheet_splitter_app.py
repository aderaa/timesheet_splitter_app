import os
import io
import re
import shutil
import zipfile
from datetime import datetime

import pandas as pd
from docx import Document

import streamlit as st

# Try to import reportlab for PDF generation
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# ==== Configuration (column names, etc.) ====

# Columns in "1.Full Timesheet Report.xlsx"
FULL_EMP_ID_COL = "EMP ID"
FULL_EMP_NAME_COL = "User Name"
FULL_DATE_COL = "Date"
FULL_HOURS_COL = "Regular Time (Hours)"
FULL_PROJECT_TYPE_COL = "Project Type"

# Columns in "3.Databse.xlsx"
DB_EMP_ID_COL = "Emp ID"
DB_EMP_NAME_COL = "Employee Name"
DB_VENDOR_COL = "Organization"

SUMMARY_DOC_NAME = "Timesheet_Split_Summary.docx"

# Default vendor for P* employees not in DB
DEFAULT_VENDOR_NAME = "UnAssigned Vendor"

# Default ignore list for Project Type
DEFAULT_IGNORE_LIST = ["On Bench"]


# ==== Simple i18n helper ====

TEXT = {
    "en": {
        "title": "Timesheet Splitter for Outsourced Staff",
        "subtitle": "Upload the full Clarity timesheet export and the vendor database. The app will create one Excel + PDF file per employee, grouped under folders per vendor, plus a DOCX summary report.",
        "settings": "Settings",
        "output_mode": "Output mode",
        "output_mode_folder": "Save files to a folder on this machine",
        "output_mode_zip": "Download everything as a ZIP file",
        "output_folder": "Target output folder (relative or absolute path)",
        "full_timesheet": "1) Full Timesheet Report (from Clarity)",
        "vendor_db": "2) Vendor / Employee Database",
        "sample_file": "Optional: Example of single-user timesheet (for reference only)",
        "folder_not_empty": "⚠️ The folder '{path}' is not empty. Its contents will be deleted before export.",
        "confirm_clear": "I understand, clear this folder before export",
        "start": "🚀 Start splitting",
        "fatal_error": "Unexpected error",
        "progress_init": "Preparing data...",
        "progress_emp": "Processing employee {idx}/{total}: {emp_id}",
        "done": "✅ Done! Timesheets have been split successfully.",
        "download_summary": "Download summary (DOCX)",
        "download_zip": "Download ZIP file",
        "metrics_title": "Run summary",
        "metric_total_emps": "Unique employees in source",
        "metric_exported_emps": "Employees exported",
        "metric_ignored_emps": "Ignored (not in DB & not P*)",
        "metric_failed_emps": "Failed splits",
        "metric_unassigned_emps": "Unassigned 'P*' employees",
        "metric_project_flagged_emps": "Employees with ignored Project Type",
        "vendor_summary_title": "Hours per vendor",
        "exported_table_title": "Exported employees",
        "ignored_table_title": "Ignored employees (not in DB & not P*)",
        "failed_table_title": "Failed splits",
        "unassigned_table_title": "Unassigned 'P*' employees (default vendor)",
        "project_ignored_table_title": "Employees with Project Types in ignore list",
        "no_failed": "No failed splits 🎉",
        "no_ignored": "No ignored employees 🎉",
        "no_unassigned": "No unassigned 'P*' employees 🎉",
        "no_project_flagged": "No employees with ignored Project Types 🎉",
        "run_timestamp": "Run timestamp",
        "period": "Timesheet period",
        "no_dates": "Not available",
        "summary_doc_title": "Timesheet Split Summary",
        "doc_section_overview": "1. Overview",
        "doc_section_vendor_summary": "2. Hours per vendor",
        "doc_section_exported_emps": "3. Exported employees",
        "doc_section_ignored": "4. Ignored employees (not in vendor database and not starting with 'P')",
        "doc_section_failed": "5. Failed splits",
        "doc_section_unassigned": "6. 'P*' employees not in vendor DB (assigned to default vendor)",
        "doc_section_project_ignored": "7. Employees with Project Types in ignore list",
        "doc_overview_bullet_1": "Run timestamp: {ts}",
        "doc_overview_bullet_2": "Timesheet period: {period}",
        "doc_overview_bullet_3": "Total unique employees in source: {total}",
        "doc_overview_bullet_4": "Employees exported: {exported}",
        "doc_overview_bullet_5": "Ignored employees (not in DB & not P*): {ignored}",
        "doc_overview_bullet_6": "Failed splits: {failed}",
        "doc_overview_bullet_7": "Unassigned 'P*' employees (default vendor): {unassigned}",
        "doc_overview_bullet_8": "Employees with project types in ignore list: {project_flagged}",
        "table_vendor": "Vendor",
        "table_vendor_hours": "Total hours",
        "table_vendor_emp_count": "Employees",
        "table_emp_id": "Emp ID",
        "table_emp_name": "Employee Name",
        "table_emp_hours": "Total hours",
        "table_reason": "Reason",
        "table_unassigned_reason": "Reason",
        "table_project_types": "Ignored Project Types",
        "reason_not_in_db": "Not in vendor database",
        "reason_unassigned": "Emp ID starts with 'P' and not found in vendor database (assigned to default vendor)",
        "reason_exception": "Exception while writing file",
        "ui_language": "Language / اللغة",
        "ui_language_en": "English",
        "ui_language_ar": "العربية",
        "ignore_list": "Project Types to flag (comma-separated)",
        "pdf_not_available": "PDF generation library (reportlab) is not installed. Only Excel files will be generated. Please run: pip install reportlab",
    },
    "ar": {
        "title": "أداة تقسيم الجداول الزمنية للموظفين الخارجيين",
        "subtitle": "قم برفع ملف الجداول الزمنية الكامل من Clarity وملف قاعدة بيانات الموردين، وستقوم الأداة بإنشاء ملف Excel وملف PDF لكل موظف داخل مجلد باسم المورد مع تقرير ملخص بصيغة DOCX.",
        "settings": "الإعدادات",
        "output_mode": "طريقة الإخراج",
        "output_mode_folder": "حفظ الملفات في مجلد على هذا الجهاز",
        "output_mode_zip": "تحميل كل الملفات كملف ZIP واحد",
        "output_folder": "مسار مجلد الإخراج (نسبي أو مطلق)",
        "full_timesheet": "١) ملف الجداول الزمنية الكامل (من Clarity)",
        "vendor_db": "٢) ملف قاعدة بيانات الموردين / الموظفين",
        "sample_file": "اختياري: مثال لجدول زمني لموظف واحد (للمرجع فقط)",
        "folder_not_empty": "⚠️ المجلد '{path}' غير فارغ. سيتم حذف محتوياته قبل التصدير.",
        "confirm_clear": "أقرّ بذلك، قم بإفراغ هذا المجلد قبل التصدير",
        "start": "🚀 ابدأ عملية التقسيم",
        "fatal_error": "خطأ غير متوقع",
        "progress_init": "جاري تجهيز البيانات...",
        "progress_emp": "جاري معالجة الموظف {idx} من {total}: {emp_id}",
        "done": "✅ تم التنفيذ! تم تقسيم الجداول الزمنية بنجاح.",
        "download_summary": "تحميل تقرير الملخص (DOCX)",
        "download_zip": "تحميل ملف ZIP",
        "metrics_title": "ملخص العملية",
        "metric_total_emps": "عدد الموظفين في المصدر",
        "metric_exported_emps": "عدد الموظفين الذين تم تصديرهم",
        "metric_ignored_emps": "الموظفون المتجاهلون (غير موجودين بقاعدة البيانات ولا يبدأ رقمهم بـ P)",
        "metric_failed_emps": "عدد المحاولات الفاشلة",
        "metric_unassigned_emps": "موظفو P غير المربوطين بمورد",
        "metric_project_flagged_emps": "موظفون لديهم نوع مشروع من قائمة التجاهل",
        "vendor_summary_title": "ساعات العمل لكل مورد",
        "exported_table_title": "الموظفون الذين تم تصديرهم",
        "ignored_table_title": "الموظفون المتجاهلون (غير موجودين بقاعدة البيانات ولا يبدأ رقمهم بـ P)",
        "failed_table_title": "المحاولات الفاشلة",
        "unassigned_table_title": "الموظفون الذين يبدأ رقمهم بـ P وغير موجودين بقاعدة البيانات (المورد الافتراضي)",
        "project_ignored_table_title": "الموظفون الذين لديهم نوع مشروع ضمن قائمة التجاهل",
        "no_failed": "لا توجد محاولات فاشلة 🎉",
        "no_ignored": "لا يوجد موظفون متجاهلون 🎉",
        "no_unassigned": "لا يوجد موظفو P غير مربوطين بمورد 🎉",
        "no_project_flagged": "لا يوجد موظفون لديهم أنواع مشروع من قائمة التجاهل 🎉",
        "run_timestamp": "تاريخ ووقت التشغيل",
        "period": "فترة الجداول الزمنية",
        "no_dates": "غير متوفر",
        "summary_doc_title": "تقرير ملخص تقسيم الجداول الزمنية",
        "doc_section_overview": "١. نظرة عامة",
        "doc_section_vendor_summary": "٢. ساعات العمل لكل مورد",
        "doc_section_exported_emps": "٣. الموظفون الذين تم تصديرهم",
        "doc_section_ignored": "٤. الموظفون المتجاهلون (غير موجودين في قاعدة بيانات الموردين ولا يبدأ رقمهم بـ P)",
        "doc_section_failed": "٥. المحاولات الفاشلة",
        "doc_section_unassigned": "٦. الموظفون الذين يبدأ رقمهم بـ P وغير موجودين في قاعدة بيانات الموردين (المورد الافتراضي)",
        "doc_section_project_ignored": "٧. الموظفون الذين لديهم أنواع مشروع ضمن قائمة التجاهل",
        "doc_overview_bullet_1": "تاريخ ووقت التشغيل: {ts}",
        "doc_overview_bullet_2": "فترة الجداول الزمنية: {period}",
        "doc_overview_bullet_3": "إجمالي عدد الموظفين في المصدر: {total}",
        "doc_overview_bullet_4": "عدد الموظفين الذين تم تصديرهم: {exported}",
        "doc_overview_bullet_5": "عدد الموظفين المتجاهلين (غير موجودين في قاعدة البيانات ولا يبدأ رقمهم بـ P): {ignored}",
        "doc_overview_bullet_6": "عدد المحاولات الفاشلة: {failed}",
        "doc_overview_bullet_7": "عدد موظفي P غير المربوطين بمورد (المورد الافتراضي): {unassigned}",
        "doc_overview_bullet_8": "عدد الموظفين الذين لديهم نوع مشروع ضمن قائمة التجاهل: {project_flagged}",
        "table_vendor": "المورد",
        "table_vendor_hours": "إجمالي الساعات",
        "table_vendor_emp_count": "عدد الموظفين",
        "table_emp_id": "رقم الموظف",
        "table_emp_name": "اسم الموظف",
        "table_emp_hours": "إجمالي الساعات",
        "table_reason": "السبب",
        "table_unassigned_reason": "السبب",
        "table_project_types": "أنواع المشروع من قائمة التجاهل",
        "reason_not_in_db": "غير موجود في قاعدة بيانات الموردين",
        "reason_unassigned": "رقم الموظف يبدأ بـ P وغير موجود في قاعدة بيانات الموردين (تم إسناده للمورد الافتراضي)",
        "reason_exception": "خطأ أثناء حفظ الملف",
        "ui_language": "Language / اللغة",
        "ui_language_en": "English",
        "ui_language_ar": "العربية",
        "ignore_list": "قائمة أنواع المشروع المطلوب تمييزها (مفصولة بفواصل)",
        "pdf_not_available": "مكتبة إنشاء ملفات PDF (reportlab) غير مثبتة، سيتم إنشاء ملفات Excel فقط. برجاء تنفيذ الأمر: pip install reportlab",
    },
}


def t(key: str, lang: str) -> str:
    """Translation helper."""
    return TEXT.get(lang, TEXT["en"]).get(key, TEXT["en"].get(key, key))


# ==== Styling ====
CUSTOM_CSS = """
<style>
.app-title {
    text-align: center;
    color: #2c3e50;
    font-size: 2.3rem;
    margin-bottom: 0.2rem;
}
.app-subtitle {
    text-align: center;
    color: #555;
    margin-bottom: 1.5rem;
}
.summary-card {
    padding: 0.9rem 1.2rem;
    border-radius: 12px;
    background: linear-gradient(135deg, #1abc9c, #3498db);
    color: white;
    margin-bottom: 1.2rem;
}
</style>
"""


# ==== Core helpers ====

def safe_name(name: str) -> str:
    """Make a string safe for use as a Windows file/folder name."""
    if pd.isna(name):
        name = "Unknown"
    name = str(name)
    # Replace characters not allowed in Windows filenames
    return re.sub(r'[<>:"/\\\\|?*]', "_", name)


def dataframe_to_pdf_bytes(df: pd.DataFrame, title: str = "") -> bytes:
    """Render a pandas DataFrame to a simple table PDF and return bytes."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("reportlab is not installed")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    elements = []
    styles = getSampleStyleSheet()

    if title:
        elements.append(Paragraph(title, styles["Heading2"]))

    # Build table data
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)

    style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ]
    )
    table.setStyle(style)
    elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


def prepare_employee_data(full_df: pd.DataFrame, db_df: pd.DataFrame, ignore_list):
    """
    Build in-memory structures for employees:
    - employees: list of dicts with vendor, emp_id, emp_name, hours, df, flags (only those to be exported)
    - ignored: employees present in full_df but not in db_df and not starting with 'P'
    - unassigned: P* employees not in db but assigned to default vendor
    - project_flagged: employees whose rows contain project types from ignore_list (NOT exported)
    """

    # Handle merged cells: forward-fill object columns
    obj_cols = full_df.select_dtypes(include=["object"]).columns
    full_df[obj_cols] = full_df[obj_cols].ffill()

    # Normalize date & hours columns
    if FULL_DATE_COL in full_df.columns:
        full_df[FULL_DATE_COL] = pd.to_datetime(full_df[FULL_DATE_COL], errors="coerce")
    if FULL_HOURS_COL in full_df.columns:
        full_df[FULL_HOURS_COL] = pd.to_numeric(full_df[FULL_HOURS_COL], errors="coerce").fillna(0.0)
    else:
        full_df[FULL_HOURS_COL] = 0.0

    # Vendor mapping
    db_df = db_df.copy()
    emp_id_to_vendor = {}
    emp_id_to_name = {}

    if DB_EMP_ID_COL not in db_df.columns or DB_VENDOR_COL not in db_df.columns:
        missing = []
        if DB_EMP_ID_COL not in db_df.columns:
            missing.append(DB_EMP_ID_COL)
        if DB_VENDOR_COL not in db_df.columns:
            missing.append(DB_VENDOR_COL)
        raise ValueError(
            f"Vendor DB is missing required columns: {', '.join(missing)}. "
            f"Found columns: {list(db_df.columns)}"
        )

    for _, row in db_df.iterrows():
        emp_id = str(row[DB_EMP_ID_COL]).strip()
        vendor = row[DB_VENDOR_COL]
        emp_name = row.get(DB_EMP_NAME_COL, None)
        if emp_id:
            emp_id_to_vendor[emp_id] = vendor
            if emp_name is not None and not pd.isna(emp_name):
                emp_id_to_name[emp_id] = emp_name

    # Prepare ignore list for Project Type
    ignore_set = {v.strip().lower() for v in ignore_list if v and isinstance(v, str)}

    employees = []          # WILL be exported
    ignored = []            # not in DB & not P*
    unassigned = []         # P* not in DB → default vendor
    project_flagged = []    # hit ignore list → REPORTED ONLY, NOT EXPORTED

    if FULL_EMP_ID_COL not in full_df.columns or FULL_EMP_NAME_COL not in full_df.columns:
        raise ValueError("Full timesheet file is missing required columns.")

    unique_emp_ids = sorted({str(x).strip() for x in full_df[FULL_EMP_ID_COL].dropna().unique()})
    has_project_type_col = FULL_PROJECT_TYPE_COL in full_df.columns

    for emp_id in unique_emp_ids:
        emp_rows = full_df[full_df[FULL_EMP_ID_COL].astype(str).str.strip() == emp_id]
        if emp_rows.empty:
            continue

        emp_name_series = emp_rows[FULL_EMP_NAME_COL].dropna()
        emp_name = str(emp_name_series.iloc[0]) if not emp_name_series.empty else ""
        total_hours = float(emp_rows[FULL_HOURS_COL].sum())

        # --- Determine vendor / ignored / unassigned ---
        is_unassigned = False

        if emp_id not in emp_id_to_vendor:
            if emp_id.startswith("P"):  # assign to default vendor
                vendor = DEFAULT_VENDOR_NAME
                is_unassigned = True
                unassigned.append(
                    {
                        "Vendor": vendor,
                        "Emp ID": emp_id,
                        "Employee Name": emp_name,
                        "Total Hours": total_hours,
                    }
                )
            else:
                # completely ignored (not exported)
                ignored.append(
                    {
                        "Emp ID": emp_id,
                        "Employee Name": emp_name,
                        "Total Hours": total_hours,
                        "Reason": "not_in_db",
                    }
                )
                continue
        else:
            vendor = emp_id_to_vendor[emp_id]

        # --- Check for ignored Project Types ---
        has_ignored_pt = False
        matched_pts = []
        if has_project_type_col and ignore_set:
            pts = emp_rows[FULL_PROJECT_TYPE_COL].dropna().astype(str).str.strip()
            matched = sorted({v for v in pts if v.strip().lower() in ignore_set})
            if matched:
                has_ignored_pt = True
                matched_pts = matched
                # This employee is REPORTED ONLY, NOT exported
                project_flagged.append(
                    {
                        "Vendor": vendor,
                        "Emp ID": emp_id,
                        "Employee Name": emp_name,
                        "Total Hours": total_hours,
                        "IgnoredProjectTypes": matched_pts,
                    }
                )
                # 🚫 DO NOT export them
                continue

        # --- If we reach here, employee WILL be exported ---
        employees.append(
            {
                "Vendor": vendor,
                "Emp ID": emp_id,
                "Employee Name": emp_name,
                "Total Hours": total_hours,
                "df": emp_rows,
                "IsUnassigned": is_unassigned,
                "HasIgnoredProjectType": has_ignored_pt,
                "IgnoredProjectTypes": matched_pts,
            }
        )

    return employees, ignored, unassigned, project_flagged, unique_emp_ids

def build_summary_structures(employees, ignored, failed, unassigned, project_flagged, full_df, lang: str):
    """Prepare data for UI + DOCX summary."""
    now = datetime.now()
    if FULL_DATE_COL in full_df.columns:
        dates = full_df[FULL_DATE_COL].dropna()
        if not dates.empty:
            period_str = f"{dates.min().date()} → {dates.max().date()}"
        else:
            period_str = t("no_dates", lang)
    else:
        period_str = t("no_dates", lang)

    total_emps = len({str(x).strip() for x in full_df[FULL_EMP_ID_COL].dropna().unique()})
    exported_emps = len({e["Emp ID"] for e in employees})
    ignored_emps = len(ignored)
    failed_emps = len(failed)
    unassigned_emps = len({u["Emp ID"] for u in unassigned})
    project_flagged_emps = len({p["Emp ID"] for p in project_flagged})

    summary_stats = {
        "run_timestamp": now,
        "period": period_str,
        "total_emps": total_emps,
        "exported_emps": exported_emps,
        "ignored_emps": ignored_emps,
        "failed_emps": failed_emps,
        "unassigned_emps": unassigned_emps,
        "project_flagged_emps": project_flagged_emps,
    }

    # Vendor summary
    vendor_summary_rows = []
    vendor_group = {}
    for e in employees:
        vendor = e["Vendor"]
        vendor_group.setdefault(vendor, {"hours": 0.0, "count": 0})
        vendor_group[vendor]["hours"] += e["Total Hours"]
        vendor_group[vendor]["count"] += 1

    for vendor, agg in vendor_group.items():
        vendor_summary_rows.append(
            {
                t("table_vendor", lang): vendor,
                t("table_vendor_hours", lang): round(agg["hours"], 2),
                t("table_vendor_emp_count", lang): agg["count"],
            }
        )
    vendor_summary_df = pd.DataFrame(vendor_summary_rows)

    # Exported employees table
    exported_rows = []
    for e in employees:
        exported_rows.append(
            {
                t("table_vendor", lang): e["Vendor"],
                t("table_emp_id", lang): e["Emp ID"],
                t("table_emp_name", lang): e["Employee Name"],
                t("table_emp_hours", lang): round(e["Total Hours"], 2),
            }
        )
    exported_df = pd.DataFrame(exported_rows)

    # Ignored table (not in DB & not P*)
    ignored_rows = []
    for ig in ignored:
        ignored_rows.append(
            {
                t("table_emp_id", lang): ig["Emp ID"],
                t("table_emp_name", lang): ig["Employee Name"],
                t("table_emp_hours", lang): round(ig["Total Hours"], 2),
                t("table_reason", lang): t("reason_not_in_db", lang),
            }
        )
    ignored_df = pd.DataFrame(ignored_rows)

    # Failed table
    failed_rows = []
    for fl in failed:
        failed_rows.append(
            {
                t("table_emp_id", lang): fl["Emp ID"],
                t("table_emp_name", lang): fl["Employee Name"],
                t("table_emp_hours", lang): round(fl.get("Total Hours", 0.0), 2),
                t("table_reason", lang): f"{t('reason_exception', lang)}: {fl['Error']}",
            }
        )
    failed_df = pd.DataFrame(failed_rows)

    # Unassigned P* employees table
    unassigned_rows = []
    for u in unassigned:
        unassigned_rows.append(
            {
                t("table_vendor", lang): u["Vendor"],
                t("table_emp_id", lang): u["Emp ID"],
                t("table_emp_name", lang): u["Employee Name"],
                t("table_emp_hours", lang): round(u["Total Hours"], 2),
                t("table_unassigned_reason", lang): t("reason_unassigned", lang),
            }
        )
    unassigned_df = pd.DataFrame(unassigned_rows)

    # Employees with ignored Project Types
    proj_rows = []
    for p in project_flagged:
        proj_rows.append(
            {
                t("table_vendor", lang): p["Vendor"],
                t("table_emp_id", lang): p["Emp ID"],
                t("table_emp_name", lang): p["Employee Name"],
                t("table_emp_hours", lang): round(p["Total Hours"], 2),
                t("table_project_types", lang): ", ".join(p["IgnoredProjectTypes"]),
            }
        )
    project_flagged_df = pd.DataFrame(proj_rows)

    return summary_stats, vendor_summary_df, exported_df, ignored_df, failed_df, unassigned_df, project_flagged_df


def build_docx_summary(
    summary_stats,
    vendor_summary_df,
    exported_df,
    ignored_df,
    failed_df,
    unassigned_df,
    project_flagged_df,
    lang: str,
) -> Document:
    """Generate the DOCX summary report."""
    doc = Document()

    doc.add_heading(t("summary_doc_title", lang), level=0)

    # Overview
    doc.add_heading(t("doc_section_overview", lang), level=1)
    ts_str = summary_stats["run_timestamp"].strftime("%Y-%m-%d %H:%M:%S")
    period_str = summary_stats["period"]
    total = summary_stats["total_emps"]
    exported = summary_stats["exported_emps"]
    ignored = summary_stats["ignored_emps"]
    failed = summary_stats["failed_emps"]
    unassigned = summary_stats["unassigned_emps"]
    project_flagged = summary_stats["project_flagged_emps"]

    bullets = [
        t("doc_overview_bullet_1", lang).format(ts=ts_str),
        t("doc_overview_bullet_2", lang).format(period=period_str),
        t("doc_overview_bullet_3", lang).format(total=total),
        t("doc_overview_bullet_4", lang).format(exported=exported),
        t("doc_overview_bullet_5", lang).format(ignored=ignored),
        t("doc_overview_bullet_6", lang).format(failed=failed),
        t("doc_overview_bullet_7", lang).format(unassigned=unassigned),
        t("doc_overview_bullet_8", lang).format(project_flagged=project_flagged),
    ]
    for b in bullets:
        doc.add_paragraph(b, style="List Bullet")

    # Vendor summary
    doc.add_heading(t("doc_section_vendor_summary", lang), level=1)
    if vendor_summary_df is not None and not vendor_summary_df.empty:
        cols = list(vendor_summary_df.columns)
        table = doc.add_table(rows=1 + len(vendor_summary_df), cols=len(cols))
        hdr_cells = table.rows[0].cells
        for j, c in enumerate(cols):
            hdr_cells[j].text = str(c)

        for i, (_, row) in enumerate(vendor_summary_df.iterrows(), start=1):
            row_cells = table.rows[i].cells
            for j, c in enumerate(cols):
                row_cells[j].text = str(row[c])
    else:
        doc.add_paragraph("—")

    # Exported employees
    doc.add_heading(t("doc_section_exported_emps", lang), level=1)
    if exported_df is not None and not exported_df.empty:
        cols = list(exported_df.columns)
        table = doc.add_table(rows=1 + len(exported_df), cols=len(cols))
        hdr_cells = table.rows[0].cells
        for j, c in enumerate(cols):
            hdr_cells[j].text = str(c)

        for i, (_, row) in enumerate(exported_df.iterrows(), start=1):
            row_cells = table.rows[i].cells
            for j, c in enumerate(cols):
                row_cells[j].text = str(row[c])
    else:
        doc.add_paragraph("—")

    # Ignored employees
    doc.add_heading(t("doc_section_ignored", lang), level=1)
    if ignored_df is not None and not ignored_df.empty:
        cols = list(ignored_df.columns)
        table = doc.add_table(rows=1 + len(ignored_df), cols=len(cols))
        hdr_cells = table.rows[0].cells
        for j, c in enumerate(cols):
            hdr_cells[j].text = str(c)

        for i, (_, row) in enumerate(ignored_df.iterrows(), start=1):
            row_cells = table.rows[i].cells
            for j, c in enumerate(cols):
                row_cells[j].text = str(row[c])
    else:
        doc.add_paragraph("—")

    # Failed splits
    doc.add_heading(t("doc_section_failed", lang), level=1)
    if failed_df is not None and not failed_df.empty:
        cols = list(failed_df.columns)
        table = doc.add_table(rows=1 + len(failed_df), cols=len(cols))
        hdr_cells = table.rows[0].cells
        for j, c in enumerate(cols):
            hdr_cells[j].text = str(c)

        for i, (_, row) in enumerate(failed_df.iterrows(), start=1):
            row_cells = table.rows[i].cells
            for j, c in enumerate(cols):
                row_cells[j].text = str(row[c])
    else:
        doc.add_paragraph("—")

    # Unassigned P* employees
    doc.add_heading(t("doc_section_unassigned", lang), level=1)
    if unassigned_df is not None and not unassigned_df.empty:
        cols = list(unassigned_df.columns)
        table = doc.add_table(rows=1 + len(unassigned_df), cols=len(cols))
        hdr_cells = table.rows[0].cells
        for j, c in enumerate(cols):
            hdr_cells[j].text = str(c)

        for i, (_, row) in enumerate(unassigned_df.iterrows(), start=1):
            row_cells = table.rows[i].cells
            for j, c in enumerate(cols):
                row_cells[j].text = str(row[c])
    else:
        doc.add_paragraph("—")

    # Employees with ignored Project Types
    doc.add_heading(t("doc_section_project_ignored", lang), level=1)
    if project_flagged_df is not None and not project_flagged_df.empty:
        cols = list(project_flagged_df.columns)
        table = doc.add_table(rows=1 + len(project_flagged_df), cols=len(cols))
        hdr_cells = table.rows[0].cells
        for j, c in enumerate(cols):
            hdr_cells[j].text = str(c)

        for i, (_, row) in enumerate(project_flagged_df.iterrows(), start=1):
            row_cells = table.rows[i].cells
            for j, c in enumerate(cols):
                row_cells[j].text = str(row[c])
    else:
        doc.add_paragraph("—")

    return doc


# ==== Streamlit app ====

def main():
    st.set_page_config(page_title="Timesheet Splitter", page_icon="⏱️", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Language selector
    if "lang" not in st.session_state:
        st.session_state["lang"] = "en"

    lang = st.sidebar.radio(
        t("ui_language", st.session_state["lang"]),
        ("en", "ar"),
        index=0,
        format_func=lambda x: t(f"ui_language_{x}", st.session_state["lang"]),
    )
    st.session_state["lang"] = lang
    lang = st.session_state["lang"]

    st.markdown(f"<h1 class='app-title'>⏱️ {t('title', lang)}</h1>", unsafe_allow_html=True)
    st.markdown(f"<p class='app-subtitle'>{t('subtitle', lang)}</p>", unsafe_allow_html=True)

    with st.sidebar:
        st.markdown(f"### ⚙️ {t('settings', lang)}")
        output_mode = st.radio(
            t("output_mode", lang),
            ("folder", "zip"),
            format_func=lambda x: t("output_mode_folder", lang) if x == "folder" else t("output_mode_zip", lang),
        )

        output_folder = None
        confirm_clear = False
        if output_mode == "folder":
            output_folder = st.text_input(t("output_folder", lang), value="output")
            if output_folder:
                if os.path.exists(output_folder) and os.listdir(output_folder):
                    st.warning(t("folder_not_empty", lang).format(path=output_folder))
                    confirm_clear = st.checkbox(t("confirm_clear", lang))

        ignore_list_input = st.text_input(t("ignore_list", lang), value=", ".join(DEFAULT_IGNORE_LIST))

    col1, col2, col3 = st.columns(3)
    with col1:
        full_file = st.file_uploader(t("full_timesheet", lang), type=["xlsx"])
    with col2:
        db_file = st.file_uploader(t("vendor_db", lang), type=["xlsx"])
    with col3:
        sample_file = st.file_uploader(t("sample_file", lang), type=["xlsx"])  # not used, just for reference

    # Determine if start button should be disabled
    disable_start = full_file is None or db_file is None
    if output_mode == "folder":
        if not output_folder:
            disable_start = True
        elif os.path.exists(output_folder) and os.listdir(output_folder) and not confirm_clear:
            disable_start = True

    start = st.button(t("start", lang), type="primary", disabled=disable_start)

    if not start:
        return

    try:
        progress = st.progress(0.0, text=t("progress_init", lang))
    except TypeError:
        # Older Streamlit without text argument
        progress = st.progress(0.0)
        st.info(t("progress_init", lang))

    status_placeholder = st.empty()

    try:
        # Read input files
        full_df = pd.read_excel(full_file)
        db_df = pd.read_excel(db_file)

        # Normalize column names (strip spaces)
        full_df.columns = full_df.columns.str.strip()
        db_df.columns = db_df.columns.str.strip()

        # Parse ignore list from UI
        ignore_list = [s.strip() for s in ignore_list_input.split(",") if s.strip()]
        if not ignore_list:
            ignore_list = DEFAULT_IGNORE_LIST

        if not REPORTLAB_AVAILABLE:
            st.warning(t("pdf_not_available", lang))

        # Prepare in-memory employee data
        employees, ignored, unassigned, project_flagged, unique_emp_ids = prepare_employee_data(
            full_df, db_df, ignore_list
        )

        total_to_process = len(employees)
        failed = []
        successful_employees = []

        # Prepare output on disk if needed
        if output_mode == "folder":
            if os.path.exists(output_folder):
                # Clear folder if user confirmed
                if os.listdir(output_folder):
                    shutil.rmtree(output_folder)
            os.makedirs(output_folder, exist_ok=True)

        # For ZIP mode, we'll build the archive in memory
        zip_buffer = io.BytesIO() if output_mode == "zip" else None
        if zip_buffer is not None:
            zip_file = zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED)
        else:
            zip_file = None

        for idx, emp in enumerate(employees, start=1):
            emp_id = emp["Emp ID"]
            emp_name = emp["Employee Name"]
            vendor = emp["Vendor"]
            emp_df = emp["df"]
            total_hours = emp["Total Hours"]

            try:
                if total_to_process:
                    progress_val = idx / total_to_process
                else:
                    progress_val = 1.0
                progress.progress(progress_val)
                status_placeholder.write(
                    t("progress_emp", lang).format(idx=idx, total=total_to_process, emp_id=emp_id)
                )

                safe_vendor_folder = safe_name(vendor)
                file_base_name = f"{vendor}-{emp_id}-{emp_name}.xlsx"
                safe_file_name = safe_name(file_base_name)

                if output_mode == "folder":
                    vendor_folder_path = os.path.join(output_folder, safe_vendor_folder)
                    os.makedirs(vendor_folder_path, exist_ok=True)
                    file_path = os.path.join(vendor_folder_path, safe_file_name)
                    emp_df.to_excel(file_path, index=False)

                    # PDF generation (same filename, .pdf extension)
                    if REPORTLAB_AVAILABLE:
                        pdf_bytes = dataframe_to_pdf_bytes(emp_df, title=file_base_name)
                        pdf_path = (
                            file_path[:-5] + ".pdf"
                            if file_path.lower().endswith(".xlsx")
                            else file_path + ".pdf"
                        )
                        with open(pdf_path, "wb") as pf:
                            pf.write(pdf_bytes)

                else:  # zip
                    # Write Excel to bytes & add to zip under vendor folder
                    xls_buffer = io.BytesIO()
                    emp_df.to_excel(xls_buffer, index=False)
                    xls_buffer.seek(0)
                    arcname = f"{safe_vendor_folder}/{safe_file_name}"
                    zip_file.writestr(arcname, xls_buffer.getvalue())

                    # PDF generation for ZIP
                    if REPORTLAB_AVAILABLE:
                        pdf_bytes = dataframe_to_pdf_bytes(emp_df, title=file_base_name)
                        if safe_file_name.lower().endswith(".xlsx"):
                            pdf_name = safe_file_name[:-5] + ".pdf"
                        else:
                            pdf_name = safe_file_name + ".pdf"
                        pdf_arcname = f"{safe_vendor_folder}/{pdf_name}"
                        zip_file.writestr(pdf_arcname, pdf_bytes)

                # If we reached here, it is successful
                successful_employees.append(
                    {
                        "Vendor": vendor,
                        "Emp ID": emp_id,
                        "Employee Name": emp_name,
                        "Total Hours": total_hours,
                    }
                )

            except Exception as e:  # noqa: BLE001
                failed.append(
                    {
                        "Vendor": vendor,
                        "Emp ID": emp_id,
                        "Employee Name": emp_name,
                        "Total Hours": total_hours,
                        "Error": str(e),
                    }
                )

        # Close ZIP if used
        if zip_file is not None:
            zip_file.close()

        # Use successful_employees for summary of exported staff only
        employees_for_summary = []
        success_keys = {(e["Emp ID"], e["Vendor"]) for e in successful_employees}
        for emp in employees:
            if (emp["Emp ID"], emp["Vendor"]) in success_keys:
                employees_for_summary.append(emp)

        # Unassigned P* staff: keep only those that were actually exported
        unassigned_for_summary = [
            u for u in unassigned if (u["Emp ID"], u["Vendor"]) in success_keys
        ]

        # 🚩 Project-ignored staff are by definition NOT exported,
        # so we keep the full list for reporting:
        project_flagged_for_summary = project_flagged


        # Build summary dataframes
        (
            summary_stats,
            vendor_summary_df,
            exported_df,
            ignored_df,
            failed_df,
            unassigned_df,
            project_flagged_df,
        ) = build_summary_structures(
            employees_for_summary,
            ignored,
            failed,
            unassigned_for_summary,
            project_flagged_for_summary,
            full_df,
            lang,
        )

        # Build DOCX summary
        doc = build_docx_summary(
            summary_stats,
            vendor_summary_df,
            exported_df,
            ignored_df,
            failed_df,
            unassigned_df,
            project_flagged_df,
            lang,
        )
        doc_buffer = io.BytesIO()
        doc.save(doc_buffer)
        doc_buffer.seek(0)

        # Place DOCX summary in output
        if output_mode == "folder":
            summary_path = os.path.join(output_folder, SUMMARY_DOC_NAME)
            with open(summary_path, "wb") as f:
                f.write(doc_buffer.getvalue())
        else:
            # Add DOCX at root of ZIP
            zip_buffer.seek(0)
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file_append:
                zip_file_append.writestr(SUMMARY_DOC_NAME, doc_buffer.getvalue())
            zip_buffer.seek(0)

        progress.progress(1.0)
        status_placeholder.empty()
        st.success(t("done", lang))

        # Show high-level metrics
        st.markdown(f"### 📊 {t('metrics_title', lang)}")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric(t("metric_total_emps", lang), summary_stats["total_emps"])
        c2.metric(t("metric_exported_emps", lang), summary_stats["exported_emps"])
        c3.metric(t("metric_ignored_emps", lang), summary_stats["ignored_emps"])
        c4.metric(t("metric_failed_emps", lang), summary_stats["failed_emps"])
        c5.metric(t("metric_unassigned_emps", lang), summary_stats["unassigned_emps"])
        c6.metric(
            t("metric_project_flagged_emps", lang),
            summary_stats["project_flagged_emps"],
        )

        # Show run details
        st.markdown("----")
        c7, c8 = st.columns(2)
        c7.write(
            f"**{t('run_timestamp', lang)}:** {summary_stats['run_timestamp'].strftime('%Y-%m-%d %H:%M:%S')}"
        )
        c8.write(f"**{t('period', lang)}:** {summary_stats['period']}")

        # Show tables
        with st.expander("📦 " + t("vendor_summary_title", lang), expanded=True):
            if vendor_summary_df is not None and not vendor_summary_df.empty:
                st.dataframe(vendor_summary_df, use_container_width=True)
            else:
                st.write("—")

        with st.expander("👤 " + t("exported_table_title", lang), expanded=False):
            if exported_df is not None and not exported_df.empty:
                st.dataframe(exported_df, use_container_width=True)
            else:
                st.write("—")

        with st.expander("👀 " + t("ignored_table_title", lang), expanded=False):
            if ignored_df is not None and not ignored_df.empty:
                st.dataframe(ignored_df, use_container_width=True)
            else:
                st.write(t("no_ignored", lang))

        with st.expander("🧩 " + t("unassigned_table_title", lang), expanded=False):
            if unassigned_df is not None and not unassigned_df.empty:
                st.dataframe(unassigned_df, use_container_width=True)
            else:
                st.write(t("no_unassigned", lang))

        with st.expander("🚩 " + t("project_ignored_table_title", lang), expanded=False):
            if project_flagged_df is not None and not project_flagged_df.empty:
                st.dataframe(project_flagged_df, use_container_width=True)
            else:
                st.write(t("no_project_flagged", lang))

        with st.expander("⚠️ " + t("failed_table_title", lang), expanded=False):
            if failed_df is not None and not failed_df.empty:
                st.dataframe(failed_df, use_container_width=True)
            else:
                st.write(t("no_failed", lang))

        # Download buttons
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📄 " + t("download_summary", lang),
                data=doc_buffer,
                file_name=SUMMARY_DOC_NAME,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        with col_dl2:
            if output_mode == "zip" and zip_buffer is not None:
                st.download_button(
                    label="🗂️ " + t("download_zip", lang),
                    data=zip_buffer,
                    file_name=f"timesheet_split_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                )

    except Exception as e:  # noqa: BLE001
        st.error(f"{t('fatal_error', lang)}: {e}")


if __name__ == "__main__":
    main()
