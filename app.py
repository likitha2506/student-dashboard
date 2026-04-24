import calendar
import io
import os
import sqlite3
from datetime import date, datetime, timedelta
from functools import wraps

from flask import (
    Flask,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, "database.db")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "static", "images")
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif"}

app = Flask(__name__)
app.config["SECRET_KEY"] = "change-this-secret-key"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_error):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'student'))
        );

        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT UNIQUE NOT NULL,
            user_id INTEGER,
            name TEXT NOT NULL,
            age INTEGER,
            department TEXT,
            email TEXT,
            phone TEXT,
            profile_photo TEXT,
            FOREIGN KEY (user_id) REFERENCES users (id)
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            status TEXT NOT NULL CHECK(status IN ('Present', 'Absent')),
            FOREIGN KEY (student_id) REFERENCES students (id)
        );

        CREATE TABLE IF NOT EXISTS marks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            subject TEXT NOT NULL,
            marks REAL NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students (id)
        );
        """
    )

    admin = db.execute("SELECT id FROM users WHERE username = ?", ("admin",)).fetchone()
    if not admin:
        db.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            ("admin", generate_password_hash("admin123"), "admin"),
        )

    student_user = db.execute("SELECT id FROM users WHERE username = ?", ("student",)).fetchone()
    if not student_user:
        cursor = db.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            ("student", generate_password_hash("student123"), "student"),
        )
        user_id = cursor.lastrowid
        db.execute(
            """
            INSERT INTO students (student_id, user_id, name, age, department, email, phone)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "S1001",
                user_id,
                "Demo Student",
                20,
                "Computer Science",
                "student@example.com",
                "1234567890",
            ),
        )

    db.commit()


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped


def role_required(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if session.get("role") not in roles:
                flash("You do not have permission to access this page.", "danger")
                return redirect(url_for("index"))
            return view(*args, **kwargs)

        return wrapped

    return decorator


def authenticate_and_login(username, password, role):
    db = get_db()
    user = db.execute(
        "SELECT * FROM users WHERE username = ? AND role = ?", (username, role)
    ).fetchone()

    if user and check_password_hash(user["password"], password):
        session.clear()
        session["user_id"] = user["id"]
        session["username"] = user["username"]
        session["role"] = user["role"]
        flash("Login successful.", "success")
        if role == "admin":
            return redirect(url_for("admin_dashboard"))
        return redirect(url_for("student_dashboard"))

    flash("Invalid username, password, or role.", "danger")
    return None


def get_student_metrics(student_pk):
    db = get_db()
    total_classes = db.execute(
        "SELECT COUNT(*) AS total FROM attendance WHERE student_id = ?", (student_pk,)
    ).fetchone()["total"]
    present_classes = db.execute(
        "SELECT COUNT(*) AS present FROM attendance WHERE student_id = ? AND status = 'Present'",
        (student_pk,),
    ).fetchone()["present"]
    avg_marks = db.execute(
        "SELECT AVG(marks) AS avg_marks FROM marks WHERE student_id = ?", (student_pk,)
    ).fetchone()["avg_marks"]

    attendance_percentage = (present_classes / total_classes * 100) if total_classes else 0
    avg_marks = float(avg_marks) if avg_marks is not None else 0.0

    if attendance_percentage < 70 or avg_marks < 60:
        status = "At Risk"
        suggestion = "Increase class attendance and focus on weaker subjects."
    else:
        status = "Safe"
        suggestion = "Keep up the consistent performance."

    predicted_gpa = round((avg_marks / 20), 2)

    return {
        "attendance_percentage": round(attendance_percentage, 2),
        "avg_marks": round(avg_marks, 2),
        "status": status,
        "suggestion": suggestion,
        "predicted_gpa": predicted_gpa,
    }


def get_notifications():
    db = get_db()

    low_attendance = db.execute(
        """
        SELECT s.name, s.student_id,
               ROUND(100.0 * SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) / COUNT(a.id), 2) AS attendance_percent
        FROM students s
        JOIN attendance a ON a.student_id = s.id
        GROUP BY s.id
        HAVING attendance_percent < 75
        """
    ).fetchall()

    low_marks = db.execute(
        """
        SELECT s.name, s.student_id, ROUND(AVG(m.marks), 2) AS avg_marks
        FROM students s
        JOIN marks m ON m.student_id = s.id
        GROUP BY s.id
        HAVING avg_marks < 60
        """
    ).fetchall()

    upcoming_exams = [
        {
            "subject": "Mathematics",
            "date": (date.today() + timedelta(days=7)).isoformat(),
        },
        {
            "subject": "Physics",
            "date": (date.today() + timedelta(days=12)).isoformat(),
        },
    ]

    return {
        "low_attendance": low_attendance,
        "low_marks": low_marks,
        "upcoming_exams": upcoming_exams,
    }


@app.route("/")
def index():
    if "user_id" not in session:
        return redirect(url_for("login"))

    if session.get("role") == "admin":
        return redirect(url_for("admin_dashboard"))

    return redirect(url_for("student_dashboard"))


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]
        role = request.form["role"]

        if role not in {"admin", "student"}:
            flash("Invalid role selected.", "danger")
            return redirect(url_for("register"))

        db = get_db()
        existing_user = db.execute(
            "SELECT id FROM users WHERE username = ?", (username,)
        ).fetchone()

        if existing_user:
            flash("Username already exists.", "warning")
            return redirect(url_for("register"))

        cursor = db.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            (username, generate_password_hash(password), role),
        )
        user_id = cursor.lastrowid

        if role == "student":
            student_id = request.form.get("student_id", "").strip() or f"S{1000 + user_id}"
            name = request.form.get("name", "").strip() or username
            db.execute(
                """
                INSERT INTO students (student_id, user_id, name, age, department, email, phone)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    student_id,
                    user_id,
                    name,
                    request.form.get("age") or None,
                    request.form.get("department"),
                    request.form.get("email"),
                    request.form.get("phone"),
                ),
            )

        db.commit()
        flash("Registration successful. Please login.", "success")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    selected_role = request.args.get("role", "")

    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]
        role = request.form["role"]

        response = authenticate_and_login(username, password, role)
        if response:
            return response

    return render_template("login.html", selected_role=selected_role)


@app.route("/login/admin", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]
        response = authenticate_and_login(username, password, "admin")
        if response:
            return response
    return render_template("login.html", selected_role="admin", locked_role=True)


@app.route("/login/student", methods=["GET", "POST"])
def student_login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]
        response = authenticate_and_login(username, password, "student")
        if response:
            return response
    return render_template("login.html", selected_role="student", locked_role=True)


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.route("/dashboard/admin")
@login_required
@role_required("admin")
def admin_dashboard():
    db = get_db()
    student_count = db.execute("SELECT COUNT(*) AS c FROM students").fetchone()["c"]
    attendance_count = db.execute("SELECT COUNT(*) AS c FROM attendance").fetchone()["c"]
    marks_count = db.execute("SELECT COUNT(*) AS c FROM marks").fetchone()["c"]

    notifications = get_notifications()

    return render_template(
        "admin_dashboard.html",
        student_count=student_count,
        attendance_count=attendance_count,
        marks_count=marks_count,
        notifications=notifications,
    )


@app.route("/dashboard/student")
@login_required
@role_required("student")
def student_dashboard():
    db = get_db()
    student = db.execute(
        "SELECT * FROM students WHERE user_id = ?", (session["user_id"],)
    ).fetchone()

    if not student:
        flash("Student profile not found.", "warning")
        return redirect(url_for("logout"))

    metrics = get_student_metrics(student["id"])
    return render_template("student_dashboard.html", student=student, metrics=metrics)


@app.route("/students")
@login_required
@role_required("admin")
def students():
    search = request.args.get("search", "").strip()
    department = request.args.get("department", "").strip()
    gpa_min = request.args.get("gpa_min", "").strip()
    gpa_max = request.args.get("gpa_max", "").strip()

    query = """
        SELECT s.*,
               ROUND(AVG(m.marks)/20.0, 2) AS gpa
        FROM students s
        LEFT JOIN marks m ON m.student_id = s.id
        WHERE 1=1
    """
    params = []

    if search:
        query += " AND (s.name LIKE ? OR s.student_id LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%"])

    if department:
        query += " AND s.department = ?"
        params.append(department)

    query += " GROUP BY s.id"

    students_list = get_db().execute(query, params).fetchall()

    def in_range(value):
        if value is None:
            value = 0
        if gpa_min:
            try:
                if float(value) < float(gpa_min):
                    return False
            except ValueError:
                pass
        if gpa_max:
            try:
                if float(value) > float(gpa_max):
                    return False
            except ValueError:
                pass
        return True

    filtered_students = [row for row in students_list if in_range(row["gpa"])]

    departments = get_db().execute(
        "SELECT DISTINCT department FROM students WHERE department IS NOT NULL AND department != ''"
    ).fetchall()

    return render_template(
        "students.html",
        students=filtered_students,
        departments=departments,
        search=search,
        department=department,
        gpa_min=gpa_min,
        gpa_max=gpa_max,
    )


@app.route("/students/add", methods=["GET", "POST"])
@login_required
@role_required("admin")
def add_student():
    if request.method == "POST":
        db = get_db()
        db.execute(
            """
            INSERT INTO students (student_id, name, age, department, email, phone)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                request.form["student_id"],
                request.form["name"],
                request.form.get("age") or None,
                request.form.get("department"),
                request.form.get("email"),
                request.form.get("phone"),
            ),
        )
        db.commit()
        flash("Student added successfully.", "success")
        return redirect(url_for("students"))

    return render_template("student_form.html", student=None)


@app.route("/students/edit/<int:student_id>", methods=["GET", "POST"])
@login_required
@role_required("admin")
def edit_student(student_id):
    db = get_db()
    student = db.execute("SELECT * FROM students WHERE id = ?", (student_id,)).fetchone()

    if not student:
        flash("Student not found.", "warning")
        return redirect(url_for("students"))

    if request.method == "POST":
        db.execute(
            """
            UPDATE students
            SET student_id = ?, name = ?, age = ?, department = ?, email = ?, phone = ?
            WHERE id = ?
            """,
            (
                request.form["student_id"],
                request.form["name"],
                request.form.get("age") or None,
                request.form.get("department"),
                request.form.get("email"),
                request.form.get("phone"),
                student_id,
            ),
        )
        db.commit()
        flash("Student updated successfully.", "success")
        return redirect(url_for("students"))

    return render_template("student_form.html", student=student)


@app.route("/students/delete/<int:student_id>", methods=["POST"])
@login_required
@role_required("admin")
def delete_student(student_id):
    db = get_db()
    db.execute("DELETE FROM attendance WHERE student_id = ?", (student_id,))
    db.execute("DELETE FROM marks WHERE student_id = ?", (student_id,))
    db.execute("DELETE FROM students WHERE id = ?", (student_id,))
    db.commit()
    flash("Student deleted successfully.", "info")
    return redirect(url_for("students"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    db = get_db()

    if session.get("role") == "admin":
        student = None
    else:
        student = db.execute(
            "SELECT * FROM students WHERE user_id = ?", (session["user_id"],)
        ).fetchone()

    if request.method == "POST" and student:
        profile_photo = student["profile_photo"]
        uploaded_file = request.files.get("profile_photo")

        if uploaded_file and uploaded_file.filename:
            if allowed_file(uploaded_file.filename):
                filename = secure_filename(
                    f"{student['student_id']}_{int(datetime.now().timestamp())}_{uploaded_file.filename}"
                )
                os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
                save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                uploaded_file.save(save_path)
                profile_photo = filename
            else:
                flash("Invalid image file type.", "danger")
                return redirect(url_for("profile"))

        db.execute(
            """
            UPDATE students
            SET name = ?, age = ?, department = ?, email = ?, phone = ?, profile_photo = ?
            WHERE id = ?
            """,
            (
                request.form.get("name"),
                request.form.get("age") or None,
                request.form.get("department"),
                request.form.get("email"),
                request.form.get("phone"),
                profile_photo,
                student["id"],
            ),
        )
        db.commit()
        flash("Profile updated successfully.", "success")
        return redirect(url_for("profile"))

    return render_template("profile.html", student=student)


@app.route("/attendance", methods=["GET", "POST"])
@login_required
@role_required("admin")
def attendance():
    db = get_db()
    students_list = db.execute("SELECT id, name, student_id FROM students ORDER BY name").fetchall()

    if request.method == "POST":
        attendance_date = request.form["date"]
        for student in students_list:
            key = f"status_{student['id']}"
            status = request.form.get(key)
            if status in {"Present", "Absent"}:
                existing = db.execute(
                    "SELECT id FROM attendance WHERE student_id = ? AND date = ?",
                    (student["id"], attendance_date),
                ).fetchone()
                if existing:
                    db.execute(
                        "UPDATE attendance SET status = ? WHERE id = ?",
                        (status, existing["id"]),
                    )
                else:
                    db.execute(
                        "INSERT INTO attendance (student_id, date, status) VALUES (?, ?, ?)",
                        (student["id"], attendance_date, status),
                    )
        db.commit()
        flash("Attendance saved successfully.", "success")
        return redirect(url_for("attendance"))

    return render_template("attendance.html", students=students_list, today=date.today().isoformat())


@app.route("/attendance/summary")
@login_required
def attendance_summary():
    db = get_db()

    if session.get("role") == "student":
        student = db.execute(
            "SELECT id, name, student_id FROM students WHERE user_id = ?", (session["user_id"],)
        ).fetchone()
        if not student:
            flash("Student profile not found.", "warning")
            return redirect(url_for("student_dashboard"))
        summary = db.execute(
            """
            SELECT s.name,
                   s.student_id,
                   COUNT(a.id) AS total_classes,
                   SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) AS presents,
                   ROUND(100.0 * SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) / COUNT(a.id), 2) AS percent
            FROM students s
            LEFT JOIN attendance a ON a.student_id = s.id
            WHERE s.id = ?
            GROUP BY s.id
            """,
            (student["id"],),
        ).fetchall()
    else:
        summary = db.execute(
            """
            SELECT s.name,
                   s.student_id,
                   COUNT(a.id) AS total_classes,
                   SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) AS presents,
                   ROUND(100.0 * SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) / COUNT(a.id), 2) AS percent
            FROM students s
            LEFT JOIN attendance a ON a.student_id = s.id
            GROUP BY s.id
            """
        ).fetchall()

    month = request.args.get("month")
    monthly_rows = []
    if month:
        monthly_rows = db.execute(
            """
            SELECT s.student_id, s.name,
                   SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) AS presents,
                   COUNT(a.id) AS total
            FROM students s
            LEFT JOIN attendance a ON a.student_id = s.id AND substr(a.date, 1, 7) = ?
            GROUP BY s.id
            ORDER BY s.name
            """,
            (month,),
        ).fetchall()

    return render_template("attendance_summary.html", summary=summary, monthly_rows=monthly_rows, month=month)


@app.route("/marks", methods=["GET", "POST"])
@login_required
@role_required("admin")
def marks():
    db = get_db()

    if request.method == "POST":
        mark_id = request.form.get("mark_id")
        student_id = request.form["student_id"]
        subject = request.form["subject"]
        score = request.form["marks"]

        if mark_id:
            db.execute(
                "UPDATE marks SET student_id = ?, subject = ?, marks = ? WHERE id = ?",
                (student_id, subject, score, mark_id),
            )
            flash("Marks updated successfully.", "success")
        else:
            db.execute(
                "INSERT INTO marks (student_id, subject, marks) VALUES (?, ?, ?)",
                (student_id, subject, score),
            )
            flash("Marks added successfully.", "success")

        db.commit()
        return redirect(url_for("marks"))

    marks_list = db.execute(
        """
        SELECT m.id, s.name, s.student_id, m.student_id AS student_pk, m.subject, m.marks
        FROM marks m
        JOIN students s ON s.id = m.student_id
        ORDER BY m.id DESC
        """
    ).fetchall()
    students_list = db.execute("SELECT id, name, student_id FROM students ORDER BY name").fetchall()

    return render_template("marks.html", marks_list=marks_list, students=students_list)


@app.route("/marks/delete/<int:mark_id>", methods=["POST"])
@login_required
@role_required("admin")
def delete_mark(mark_id):
    db = get_db()
    db.execute("DELETE FROM marks WHERE id = ?", (mark_id,))
    db.commit()
    flash("Marks record deleted.", "info")
    return redirect(url_for("marks"))


@app.route("/marks/report")
@login_required
def marks_report():
    db = get_db()

    if session.get("role") == "student":
        student = db.execute(
            "SELECT id FROM students WHERE user_id = ?", (session["user_id"],)
        ).fetchone()
        if not student:
            flash("Student profile not found.", "warning")
            return redirect(url_for("student_dashboard"))

        rows = db.execute(
            "SELECT subject, marks FROM marks WHERE student_id = ? ORDER BY subject",
            (student["id"],),
        ).fetchall()
        grouped = {"My Report": rows}
    else:
        rows = db.execute(
            """
            SELECT s.name, s.student_id, m.subject, m.marks
            FROM marks m
            JOIN students s ON s.id = m.student_id
            ORDER BY s.name, m.subject
            """
        ).fetchall()
        grouped = {}
        for row in rows:
            key = f"{row['name']} ({row['student_id']})"
            grouped.setdefault(key, []).append(row)

    averages = db.execute(
        "SELECT subject, ROUND(AVG(marks), 2) AS avg_marks FROM marks GROUP BY subject ORDER BY subject"
    ).fetchall()

    semester_avg = db.execute("SELECT ROUND(AVG(marks), 2) AS avg FROM marks").fetchone()["avg"]
    semester_avg = semester_avg if semester_avg is not None else 0

    return render_template(
        "marks_report.html",
        grouped=grouped,
        averages=averages,
        semester_avg=semester_avg,
    )


@app.route("/analytics")
@login_required
@role_required("admin")
def analytics():
    return render_template("analytics.html")


@app.route("/api/charts")
@login_required
def api_charts():
    db = get_db()

    marks_by_subject = db.execute(
        "SELECT subject, ROUND(AVG(marks), 2) AS avg_marks FROM marks GROUP BY subject ORDER BY subject"
    ).fetchall()

    attendance_data = db.execute(
        """
        SELECT s.name,
               ROUND(100.0 * SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) / COUNT(a.id), 2) AS attendance_percent
        FROM students s
        JOIN attendance a ON a.student_id = s.id
        GROUP BY s.id
        ORDER BY s.name
        """
    ).fetchall()

    gpa_trend = db.execute(
        """
        SELECT substr(created_at, 1, 7) AS month,
               ROUND(AVG(marks)/20.0, 2) AS gpa
        FROM marks
        GROUP BY substr(created_at, 1, 7)
        ORDER BY month
        """
    ).fetchall()

    dept_perf = db.execute(
        """
        SELECT s.department,
               ROUND(AVG(m.marks), 2) AS avg_marks
        FROM students s
        JOIN marks m ON m.student_id = s.id
        WHERE s.department IS NOT NULL AND s.department != ''
        GROUP BY s.department
        ORDER BY s.department
        """
    ).fetchall()

    return jsonify(
        {
            "marksBySubject": [dict(row) for row in marks_by_subject],
            "attendancePercent": [dict(row) for row in attendance_data],
            "gpaTrend": [dict(row) for row in gpa_trend],
            "departmentPerformance": [dict(row) for row in dept_perf],
        }
    )


@app.route("/export/students/excel")
@login_required
@role_required("admin")
def export_students_excel():
    try:
        from openpyxl import Workbook
    except ImportError:
        flash("Install openpyxl to export Excel files: pip install openpyxl", "danger")
        return redirect(url_for("students"))

    db = get_db()
    rows = db.execute(
        "SELECT student_id, name, age, department, email, phone FROM students ORDER BY name"
    ).fetchall()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Students"
    headers = ["Student ID", "Name", "Age", "Department", "Email", "Phone"]
    sheet.append(headers)

    for row in rows:
        sheet.append([row["student_id"], row["name"], row["age"], row["department"], row["email"], row["phone"]])

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="students_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def build_simple_pdf(title, headers, rows):
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
    except ImportError:
        return None

    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=letter)
    width, height = letter
    y = height - 40

    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(40, y, title)
    y -= 30

    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(40, y, " | ".join(headers))
    y -= 20

    pdf.setFont("Helvetica", 10)
    for row in rows:
        line = " | ".join([str(item) for item in row])
        pdf.drawString(40, y, line[:120])
        y -= 16
        if y < 40:
            pdf.showPage()
            y = height - 40
            pdf.setFont("Helvetica", 10)

    pdf.save()
    output.seek(0)
    return output


@app.route("/export/attendance/pdf")
@login_required
@role_required("admin")
def export_attendance_pdf():
    db = get_db()
    rows = db.execute(
        """
        SELECT s.student_id, s.name,
               COUNT(a.id) AS total,
               SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) AS present
        FROM students s
        LEFT JOIN attendance a ON a.student_id = s.id
        GROUP BY s.id
        ORDER BY s.name
        """
    ).fetchall()

    table_rows = []
    for row in rows:
        total = row["total"] or 0
        present = row["present"] or 0
        percent = round((present / total * 100), 2) if total else 0
        table_rows.append([row["student_id"], row["name"], total, present, f"{percent}%"])

    pdf_buffer = build_simple_pdf(
        "Attendance Report",
        ["Student ID", "Name", "Total", "Present", "Percent"],
        table_rows,
    )

    if pdf_buffer is None:
        flash("Install reportlab to export PDF files: pip install reportlab", "danger")
        return redirect(url_for("attendance_summary"))

    return send_file(pdf_buffer, as_attachment=True, download_name="attendance_report.pdf", mimetype="application/pdf")


@app.route("/export/marks/pdf")
@login_required
@role_required("admin")
def export_marks_pdf():
    db = get_db()
    rows = db.execute(
        """
        SELECT s.student_id, s.name, m.subject, m.marks
        FROM marks m
        JOIN students s ON s.id = m.student_id
        ORDER BY s.name, m.subject
        """
    ).fetchall()

    table_rows = [[row["student_id"], row["name"], row["subject"], row["marks"]] for row in rows]
    pdf_buffer = build_simple_pdf(
        "Marks Report",
        ["Student ID", "Name", "Subject", "Marks"],
        table_rows,
    )

    if pdf_buffer is None:
        flash("Install reportlab to export PDF files: pip install reportlab", "danger")
        return redirect(url_for("marks_report"))

    return send_file(pdf_buffer, as_attachment=True, download_name="marks_report.pdf", mimetype="application/pdf")


@app.template_filter("month_name")
def month_name(value):
    try:
        year, month = value.split("-")
        return f"{calendar.month_name[int(month)]} {year}"
    except Exception:
        return value


if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    with app.app_context():
        init_db()
    app.run(debug=True)
