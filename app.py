
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
import sqlite3, os, math, io
from datetime import datetime
from functools import wraps

app = Flask(__name__)
app.secret_key = os.environ.get("SAARIC_SECRET", "change-this-local-secret")
DB = os.path.join(os.path.dirname(__file__), "saaric.db")

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    with db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS items(
            id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT UNIQUE, name TEXT NOT NULL,
            description TEXT, price REAL NOT NULL DEFAULT 0)""")
        con.execute("""CREATE TABLE IF NOT EXISTS bills(
            id INTEGER PRIMARY KEY AUTOINCREMENT, bill_no TEXT UNIQUE, customer_name TEXT NOT NULL,
            phone TEXT, address TEXT, created_at TEXT, subtotal REAL, total REAL)""")
        con.execute("""CREATE TABLE IF NOT EXISTS bill_items(
            id INTEGER PRIMARY KEY AUTOINCREMENT, bill_id INTEGER, item_name TEXT,
            code TEXT, price REAL, qty INTEGER, amount REAL,
            FOREIGN KEY(bill_id) REFERENCES bills(id))""")

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("admin"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

@app.route("/")
def home():
    return redirect(url_for("billing") if session.get("admin") else url_for("login"))

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username","").strip()
        password = request.form.get("password","")
        # Change these credentials before using the app beyond a local demo.
        if username == "admin" and password == "saaric123":
            session["admin"] = True
            return redirect(url_for("dashboard"))
        flash("Username ya password galat hai.")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/dashboard")
@login_required
def dashboard():
    with db() as con:
        items = con.execute("SELECT * FROM items ORDER BY name").fetchall()
        bills = con.execute("SELECT * FROM bills ORDER BY id DESC LIMIT 8").fetchall()
    return render_template("dashboard.html", items=items, bills=bills)

@app.route("/items", methods=["GET","POST"])
@login_required
def items():
    with db() as con:
        if request.method == "POST":
            code = request.form.get("code","").strip().upper()
            name = request.form.get("name","").strip()
            description = request.form.get("description","").strip()
            try:
                price = float(request.form.get("price",0))
                if not name or price < 0: raise ValueError()
                con.execute("INSERT INTO items(code,name,description,price) VALUES(?,?,?,?)",
                            (code or None,name,description,price))
                flash("Item add ho gaya.")
            except sqlite3.IntegrityError:
                flash("Ye product code pehle se hai.")
            except ValueError:
                flash("Valid item name aur price enter karein.")
            return redirect(url_for("items"))
        rows = con.execute("SELECT * FROM items ORDER BY name").fetchall()
    return render_template("items.html", items=rows)

@app.route("/items/delete/<int:item_id>", methods=["POST"])
@login_required
def delete_item(item_id):
    with db() as con:
        con.execute("DELETE FROM items WHERE id=?", (item_id,))
    flash("Item delete kar diya gaya.")
    return redirect(url_for("items"))

@app.route("/billing", methods=["GET","POST"])
@login_required
def billing():
    with db() as con:
        items = con.execute("SELECT * FROM items ORDER BY name").fetchall()
    if request.method == "POST":
        customer = request.form.get("customer_name","").strip()
        phone = request.form.get("phone","").strip()
        address = request.form.get("address","").strip()
        ids = request.form.getlist("item_id[]")
        qtys = request.form.getlist("qty[]")
        bill_lines = []
        subtotal = 0
        with db() as con:
            for item_id, qty in zip(ids, qtys):
                try:
                    qty = int(qty)
                    item = con.execute("SELECT * FROM items WHERE id=?", (int(item_id),)).fetchone()
                    if item and qty > 0:
                        amount = float(item["price"]) * qty
                        subtotal += amount
                        bill_lines.append((item["name"], item["code"] or "", float(item["price"]), qty, amount))
                except (ValueError, TypeError):
                    continue
            if not customer or not bill_lines:
                flash("Customer name aur kam se kam ek valid item zaroor add karein.")
                return redirect(url_for("billing"))
            now = datetime.now().strftime("%d-%m-%Y %I:%M %p")
            next_id = con.execute("SELECT COALESCE(MAX(id),0)+1 FROM bills").fetchone()[0]
            bill_no = f"SE-{datetime.now().strftime('%y%m')}-{next_id:04d}"
            cur = con.execute("INSERT INTO bills(bill_no,customer_name,phone,address,created_at,subtotal,total) VALUES(?,?,?,?,?,?,?)",
                              (bill_no,customer,phone,address,now,subtotal,subtotal))
            bill_id = cur.lastrowid
            for line in bill_lines:
                con.execute("INSERT INTO bill_items(bill_id,item_name,code,price,qty,amount) VALUES(?,?,?,?,?,?)",
                            (bill_id,*line))
        return redirect(url_for("view_bill", bill_id=bill_id))
    return render_template("billing.html", items=items)

@app.route("/bill/<int:bill_id>")
@login_required
def view_bill(bill_id):
    with db() as con:
        bill = con.execute("SELECT * FROM bills WHERE id=?", (bill_id,)).fetchone()
        lines = con.execute("SELECT * FROM bill_items WHERE bill_id=?", (bill_id,)).fetchall()
    if not bill:
        return redirect(url_for("dashboard"))
    return render_template("bill.html", bill=bill, lines=lines)

@app.route("/bills")
@login_required
def bills():
    q = request.args.get("q", "").strip()
    with db() as con:
        if q:
            pattern = f"%{q}%"
            rows = con.execute("""SELECT DISTINCT b.* FROM bills b
                LEFT JOIN bill_items bi ON bi.bill_id=b.id
                WHERE b.customer_name LIKE ? OR b.phone LIKE ? OR b.bill_no LIKE ?
                   OR bi.code LIKE ? OR bi.item_name LIKE ?
                ORDER BY b.id DESC""", (pattern, pattern, pattern, pattern, pattern)).fetchall()
        else:
            rows = con.execute("SELECT * FROM bills ORDER BY id DESC").fetchall()
    return render_template("bills.html", bills=rows, q=q)

@app.route("/bills/export")
@login_required
def export_bills():
    # Export one row per billed product, preserving bill/customer/product details.
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    with db() as con:
        rows = con.execute("""SELECT b.bill_no, b.created_at, b.customer_name, b.phone, b.address,
            bi.code, bi.item_name, bi.price, bi.qty, bi.amount, b.subtotal, b.total
            FROM bills b JOIN bill_items bi ON bi.bill_id=b.id
            ORDER BY b.id DESC, bi.id ASC""").fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "All Bill Details"
    headers = ["Bill Number", "Bill Date", "Customer Name", "Customer Phone", "Customer Address",
               "Product ID / Code", "Product Name", "Unit Price (INR)", "Quantity", "Line Amount (INR)",
               "Bill Subtotal (INR)", "Bill Total (INR)"]
    ws.append(headers)
    for row in rows:
        ws.append([row[k] for k in ("bill_no", "created_at", "customer_name", "phone", "address", "code",
            "item_name", "price", "qty", "amount", "subtotal", "total")])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="25211D")
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 14), 36)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Saaric_Bills_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    init_db()
    app.run(host="127.0.0.1", port=5000, debug=False)
