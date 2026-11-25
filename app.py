import os
import io
import calendar
from datetime import datetime, timezone, timedelta
from flask import Flask, render_template, request, jsonify, send_file
from sqlalchemy import create_engine, Column, Integer, String, Date, UniqueConstraint
from sqlalchemy.orm import declarative_base, sessionmaker
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ==================== PIN-KODLAR SOZLAMALARI ====================
ROLE_PINS = {
    "ustoz": {
        "pin": os.environ.get("USTOZ_PIN", "1111"),
        "label": "Ustoz"
    },
    "oshpaz": {
        "pin": os.environ.get("OSHPAZ_PIN", "2222"),
        "label": "Oshpaz"
    },
    "admin": {
        "pin": os.environ.get("ADMIN_PIN", "7777"),
        "label": "Super Admin"
    }
}

# ==================== POSTGRESQL BAZA SOZLAMALARI ====================
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///local_davomat.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg://", 1)

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

TASHKENT_TZ = timezone(timedelta(hours=5))


def get_tashkent_today():
    return datetime.now(TASHKENT_TZ).date()


# ==================== BAZA JADVALLARI ====================
class Student(Base):
    __tablename__ = "students"
    id = Column(Integer, primary_key=True, index=True)
    full_name = Column(String(200), nullable=False)
    class_name = Column(String(50), nullable=False)


class Attendance(Base):
    __tablename__ = "attendance"
    id = Column(Integer, primary_key=True, index=True)
    student_id = Column(Integer, nullable=False, index=True)
    att_date = Column(Date, nullable=False, index=True)
    status = Column(String(20), nullable=True)  # 'present', 'absent'
    reason = Column(String(300), default="")

    __table_args__ = (
        UniqueConstraint("student_id", "att_date", name="uq_student_date"),
    )


Base.metadata.create_all(bind=engine)


# ==================== ASOSIY SAHIFA ====================
@app.route("/")
def index():
    return render_template("index.html")


# ==================== API: LOGIN TEKSHIRISH ====================
@app.route("/api/login", methods=["POST"])
def login():
    data = request.get_json() or {}
    role = data.get("role", "ustoz")
    pin = str(data.get("pin", "")).strip()

    role_info = ROLE_PINS.get(role)
    if not role_info or pin != role_info["pin"]:
        return jsonify({"success": False, "error": "Tanlangan rol uchun PIN-kod noto'g'ri!"}), 401

    return jsonify({
        "success": True,
        "role": role,
        "label": role_info["label"]
    })


# ==================== API: OSHPAZ REJIMI (HAR 1 SONIYADA) ====================
@app.route("/api/chef-stats", methods=["GET"])
def chef_stats():
    today = get_tashkent_today()
    db = SessionLocal()
    try:
        present_count = db.query(Attendance).filter(
            Attendance.att_date == today,
            Attendance.status == "present"
        ).count()
        return jsonify({
            "date": today.strftime("%d.%m.%Y"),
            "present_count": present_count
        })
    finally:
        db.close()


# ==================== API: DAVOMAT VA O'QUVCHILAR ====================
@app.route("/api/data", methods=["GET"])
def get_data():
    date_str = request.args.get("date")
    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else get_tashkent_today()
    except ValueError:
        target_date = get_tashkent_today()

    db = SessionLocal()
    try:
        students = db.query(Student).order_by(Student.class_name, Student.full_name).all()
        records = db.query(Attendance).filter(Attendance.att_date == target_date).all()

        att_map = {
            r.student_id: {
                "status": r.status,
                "reason": r.reason or ""
            }
            for r in records
        }

        students_list = [
            {"id": s.id, "full_name": s.full_name, "class_name": s.class_name}
            for s in students
        ]

        return jsonify({"students": students_list, "attendance": att_map})
    finally:
        db.close()


@app.route("/api/students", methods=["POST"])
def add_student():
    data = request.get_json() or {}
    full_name = (data.get("full_name") or "").strip()
    class_name = (data.get("class_name") or "").strip()

    if not full_name or not class_name:
        return jsonify({"error": "Ism va sinf kiritilishi shart"}), 400

    db = SessionLocal()
    try:
        new_student = Student(full_name=full_name, class_name=class_name)
        db.add(new_student)
        db.commit()
        db.refresh(new_student)
        return jsonify({
            "id": new_student.id,
            "full_name": new_student.full_name,
            "class_name": new_student.class_name
        })
    finally:
        db.close()


@app.route("/api/students/<int:student_id>", methods=["DELETE"])
def delete_student(student_id):
    db = SessionLocal()
    try:
        student = db.query(Student).filter(Student.id == student_id).first()
        if not student:
            return jsonify({"error": "O'quvchi topilmadi"}), 404

        db.query(Attendance).filter(Attendance.student_id == student_id).delete()
        db.delete(student)
        db.commit()
        return jsonify({"status": "deleted"})
    finally:
        db.close()


@app.route("/api/attendance/single", methods=["POST"])
def save_single_attendance():
    data = request.get_json() or {}
    target_date = datetime.strptime(data.get("date"), "%Y-%m-%d").date()
    student_id = data.get("student_id")
    status = data.get("status")
    reason = (data.get("reason") or "").strip()

    db = SessionLocal()
    try:
        record = db.query(Attendance).filter(
            Attendance.student_id == student_id,
            Attendance.att_date == target_date
        ).first()

        if record:
            record.status = status
            record.reason = reason
        else:
            db.add(Attendance(
                student_id=student_id,
                att_date=target_date,
                status=status,
                reason=reason
            ))
        db.commit()
        return jsonify({"status": "ok"})
    finally:
        db.close()


@app.route("/api/attendance/bulk", methods=["POST"])
def save_bulk_attendance():
    data = request.get_json() or {}
    target_date = datetime.strptime(data.get("date"), "%Y-%m-%d").date()
    records_dict = data.get("records", {})

    db = SessionLocal()
    try:
        for sid_str, info in records_dict.items():
            sid = int(sid_str)
            status = info.get("status")
            reason = (info.get("reason") or "").strip()

            record = db.query(Attendance).filter(
                Attendance.student_id == sid,
                Attendance.att_date == target_date
            ).first()

            if record:
                record.status = status
                record.reason = reason
            else:
                db.add(Attendance(
                    student_id=sid,
                    att_date=target_date,
                    status=status,
                    reason=reason
                ))
        db.commit()
        return jsonify({"status": "saved"})
    finally:
        db.close()


# ==================== EXCEL DIZAYN VA EKSPORT ====================
def get_excel_styles():
    thin_side = Side(border_style="thin", color="CBD5E1")
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    return {
        "title_font": Font(name="Calibri", size=14, bold=True, color="FFFFFF"),
        "title_fill": PatternFill(start_color="1D63ED", end_color="1D63ED", fill_type="solid"),
        "header_font": Font(name="Calibri", size=11, bold=True, color="0F172A"),
        "header_fill": PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid"),
        "regular_font": Font(name="Calibri", size=11, color="0F172A"),
        "bold_font": Font(name="Calibri", size=11, bold=True, color="0F172A"),
        "present_fill": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
        "present_font": Font(name="Calibri", size=11, bold=True, color="15803D"),
        "absent_fill": PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
        "absent_font": Font(name="Calibri", size=11, bold=True, color="B91C1C"),
        "unmarked_fill": PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"),
        "unmarked_font": Font(name="Calibri", size=10, italic=True, color="64748B"),
        "border": border,
        "center": Alignment(horizontal="center", vertical="center"),
        "left": Alignment(horizontal="left", vertical="center"),
    }


@app.route("/api/export/daily", methods=["GET"])
def export_daily_excel():
    date_str = request.args.get("date")
    class_filter = request.args.get("class_name", "ALL")
    target_date = datetime.strptime(date_str, "%Y-%m-%d").date()

    db = SessionLocal()
    try:
        query = db.query(Student).order_by(Student.class_name, Student.full_name)
        if class_filter and class_filter != "ALL":
            query = query.filter(Student.class_name == class_filter)
        students = query.all()

        records = db.query(Attendance).filter(Attendance.att_date == target_date).all()
        att_map = {r.student_id: r for r in records}

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Davomat {date_str}"
        st = get_excel_styles()

        ws.merge_cells("A1:E1")
        ws["A1"] = f"KUNLIK DAVOMAT HISOBOTI — {date_str} ({class_filter if class_filter != 'ALL' else 'Barcha sinflar'})"
        ws["A1"].font = st["title_font"]
        ws["A1"].fill = st["title_fill"]
        ws["A1"].alignment = st["center"]
        ws.row_dimensions[1].height = 32

        ws.append(["#", "O'quvchi F.I.Sh", "Sinf", "Holati", "Sabab / Izoh"])
        ws.row_dimensions[2].height = 25
        for col_idx in range(1, 6):
            cell = ws.cell(row=2, column=col_idx)
            cell.font = st["header_font"]
            cell.fill = st["header_fill"]
            cell.alignment = st["center"] if col_idx != 2 else st["left"]
            cell.border = st["border"]

        present_count, absent_count, unmarked_count = 0, 0, 0

        for idx, s in enumerate(students, start=1):
            rec = att_map.get(s.id)
            status_val = rec.status if rec else None
            reason_val = rec.reason if rec and rec.reason else ""

            if status_val == "present":
                holat_text = "Bor (+)"
                present_count += 1
            elif status_val == "absent":
                holat_text = "Yo'q (-)"
                absent_count += 1
            else:
                holat_text = "Belgilanmagan"
                unmarked_count += 1

            row_num = idx + 2
            ws.append([idx, s.full_name, s.class_name, holat_text, reason_val])
            ws.row_dimensions[row_num].height = 22

            for col_idx in range(1, 6):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.border = st["border"]
                cell.font = st["regular_font"]
                cell.alignment = st["center"] if col_idx in (1, 3, 4) else st["left"]

                if col_idx == 2:
                    cell.font = st["bold_font"]
                elif col_idx == 4:
                    if status_val == "present":
                        cell.fill = st["present_fill"]
                        cell.font = st["present_font"]
                    elif status_val == "absent":
                        cell.fill = st["absent_fill"]
                        cell.font = st["absent_font"]
                    else:
                        cell.fill = st["unmarked_fill"]
                        cell.font = st["unmarked_font"]

        summary_row = len(students) + 4
        ws.cell(
            row=summary_row, column=2,
            value=f"Jami: {len(students)} | Bor (+): {present_count} | Yo'q (-): {absent_count} | Belgilanmagan: {unmarked_count}"
        ).font = st["bold_font"]

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 34
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 35

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"Davomat_{date_str}_{class_filter}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    finally:
        db.close()


@app.route("/api/export/monthly", methods=["GET"])
def export_monthly_excel():
    month_str = request.args.get("month")
    class_filter = request.args.get("class_name", "ALL")

    year, month = map(int, month_str.split("-"))
    num_days = calendar.monthrange(year, month)[1]
    start_date = datetime(year, month, 1).date()
    end_date = datetime(year, month, num_days).date()

    db = SessionLocal()
    try:
        query = db.query(Student).order_by(Student.class_name, Student.full_name)
        if class_filter and class_filter != "ALL":
            query = query.filter(Student.class_name == class_filter)
        students = query.all()

        records = db.query(Attendance).filter(
            Attendance.att_date >= start_date,
            Attendance.att_date <= end_date
        ).all()

        att_lookup = {(r.student_id, r.att_date.day): r.status for r in records}

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Oylik {month_str}"
        st = get_excel_styles()

        total_cols = 3 + num_days + 2
        last_col_letter = get_column_letter(total_cols)

        ws.merge_cells(f"A1:{last_col_letter}1")
        ws["A1"] = f"1 OYLIK DAVOMAT JADVALI — {month_str} ({class_filter if class_filter != 'ALL' else 'Barcha sinflar'})"
        ws["A1"].font = st["title_font"]
        ws["A1"].fill = st["title_fill"]
        ws["A1"].alignment = st["center"]
        ws.row_dimensions[1].height = 32

        headers = ["#", "O'quvchi F.I.Sh", "Sinf"] + [str(d) for d in range(1, num_days + 1)] + ["Bor (+)", "Yo'q (-)"]
        ws.append(headers)
        ws.row_dimensions[2].height = 24

        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=2, column=col_idx)
            cell.font = st["header_font"]
            cell.fill = st["header_fill"]
            cell.alignment = st["center"] if col_idx != 2 else st["left"]
            cell.border = st["border"]

        for idx, s in enumerate(students, start=1):
            row_num = idx + 2
            p_total, a_total = 0, 0
            row_data = [idx, s.full_name, s.class_name]
            day_statuses = []

            for d in range(1, num_days + 1):
                status = att_lookup.get((s.id, d))
                day_statuses.append(status)
                if status == "present":
                    row_data.append("+")
                    p_total += 1
                elif status == "absent":
                    row_data.append("-")
                    a_total += 1
                else:
                    row_data.append("")

            row_data.extend([p_total, a_total])
            ws.append(row_data)
            ws.row_dimensions[row_num].height = 21

            for col_idx in range(1, total_cols + 1):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.border = st["border"]
                cell.font = st["regular_font"]
                cell.alignment = st["center"] if col_idx != 2 else st["left"]

                if col_idx == 2:
                    cell.font = st["bold_font"]
                elif 4 <= col_idx <= 3 + num_days:
                    d_status = day_statuses[col_idx - 4]
                    if d_status == "present":
                        cell.fill = st["present_fill"]
                        cell.font = st["present_font"]
                    elif d_status == "absent":
                        cell.fill = st["absent_fill"]
                        cell.font = st["absent_font"]
                elif col_idx == total_cols - 1:
                    cell.font = st["present_font"]
                elif col_idx == total_cols:
                    cell.font = st["absent_font"]

        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 32
        ws.column_dimensions["C"].width = 10
        for d in range(1, num_days + 1):
            ws.column_dimensions[get_column_letter(3 + d)].width = 4.2
        ws.column_dimensions[get_column_letter(total_cols - 1)].width = 9
        ws.column_dimensions[get_column_letter(total_cols)].width = 9
        ws.freeze_panes = "D3"

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"Oylik_Davomat_{month_str}_{class_filter}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    finally:
        db.close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)