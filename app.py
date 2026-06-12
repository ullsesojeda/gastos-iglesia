from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
DB_NAME = "gastos.db"

# ------------------ BASE DE DATOS ------------------

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS gastos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            pagado_a TEXT,
            concepto TEXT,
            observaciones TEXT,
            responsable TEXT,
            importe REAL
        )
    """)
    conn.commit()
    conn.close()

def get_connection():
    return sqlite3.connect(DB_NAME)

# ------------------ RUTAS PRINCIPALES ------------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        fecha = request.form.get("fecha")
        pagado_a = request.form.get("pagado_a")
        concepto = request.form.get("concepto")
        observaciones = request.form.get("observaciones")
        responsable = request.form.get("responsable")
        importe = request.form.get("importe")

        conn = get_connection()
        c = conn.cursor()
        c.execute("""
            INSERT INTO gastos (fecha, pagado_a, concepto, observaciones, responsable, importe)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (fecha, pagado_a, concepto, observaciones, responsable, float(importe)))
        conn.commit()
        conn.close()

        return redirect(url_for("list_gastos"))

    return render_template("index.html")

@app.route("/gastos", methods=["GET"])
def list_gastos():
    # Filtros
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")
    concepto = request.args.get("concepto")
    responsable = request.args.get("responsable")

    query = "SELECT * FROM gastos WHERE 1=1"
    params = []

    if fecha_inicio:
        query += " AND fecha >= ?"
        params.append(fecha_inicio)
    if fecha_fin:
        query += " AND fecha <= ?"
        params.append(fecha_fin)
    if concepto:
        query += " AND concepto LIKE ?"
        params.append(f"%{concepto}%")
    if responsable:
        query += " AND responsable = ?"
        params.append(responsable)

    conn = get_connection()
    c = conn.cursor()
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    # Totales
    total_registros = len(rows)
    total_importe = sum(r[6] for r in rows) if rows else 0

    return render_template(
        "list.html",
        gastos=rows,
        total_registros=total_registros,
        total_importe=total_importe,
        filtros={
            "fecha_inicio": fecha_inicio or "",
            "fecha_fin": fecha_fin or "",
            "concepto": concepto or "",
            "responsable": responsable or ""
        }
    )

# ------------------ EDITAR / ELIMINAR ------------------

@app.route("/editar/<int:gasto_id>", methods=["GET", "POST"])
def editar_gasto(gasto_id):
    conn = get_connection()
    c = conn.cursor()

    if request.method == "POST":
        fecha = request.form.get("fecha")
        pagado_a = request.form.get("pagado_a")
        concepto = request.form.get("concepto")
        observaciones = request.form.get("observaciones")
        responsable = request.form.get("responsable")
        importe = request.form.get("importe")

        c.execute("""
            UPDATE gastos
            SET fecha = ?, pagado_a = ?, concepto = ?, observaciones = ?, responsable = ?, importe = ?
            WHERE id = ?
        """, (fecha, pagado_a, concepto, observaciones, responsable, float(importe), gasto_id))
        conn.commit()
        conn.close()
        return redirect(url_for("list_gastos"))

    c.execute("SELECT * FROM gastos WHERE id = ?", (gasto_id,))
    gasto = c.fetchone()
    conn.close()
    return render_template("index.html", gasto=gasto)

@app.route("/eliminar/<int:gasto_id>", methods=["POST"])
def eliminar_gasto(gasto_id):
    conn = get_connection()
    c = conn.cursor()
    c.execute("DELETE FROM gastos WHERE id = ?", (gasto_id,))
    conn.commit()
    conn.close()
    return redirect(url_for("list_gastos"))

# ------------------ IMPORTAR / EXPORTAR EXCEL ------------------

@app.route("/importar_excel", methods=["POST"])
def importar_excel():
    file = request.files.get("archivo_excel")

    if not file:
        return redirect(url_for("list_gastos"))

    df = pd.read_excel(file)

    # Normalizar encabezados
    df.columns = df.columns.str.strip().str.upper()

    conn = sqlite3.connect("gastos.db")
    cursor = conn.cursor()

    for _, row in df.iterrows():

        # Convertir fecha DD/MM/YYYY → YYYY-MM-DD
        fecha_excel = str(row["FECHA"]).strip().replace("//", "")
        dia, mes, anio = fecha_excel.split("/")
        fecha_sql = f"{anio}-{mes}-{dia}"

        cursor.execute("""
            INSERT INTO gastos (fecha, pagado_a, concepto, observaciones, responsable, importe)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            fecha_sql,
            str(row["PAGADO_A"]),
            str(row["CONCEPTO"]),
            str(row["OBSERVACIONES"]),
            str(row["RESPONSABLE"]),
            float(row["IMPORTE"])
        ))

    conn.commit()
    conn.close()

    return redirect(url_for("list_gastos"))

@app.route("/exportar_excel", methods=["GET"])
def exportar_excel():
    # Usamos los mismos filtros que en la lista
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")
    concepto = request.args.get("concepto")
    responsable = request.args.get("responsable")

    query = "SELECT * FROM gastos WHERE 1=1"
    params = []

    if fecha_inicio:
        query += " AND fecha >= ?"
        params.append(fecha_inicio)
    if fecha_fin:
        query += " AND fecha <= ?"
        params.append(fecha_fin)
    if concepto:
        query += " AND concepto LIKE ?"
        params.append(f"%{concepto}%")
    if responsable:
        query += " AND responsable = ?"
        params.append(responsable)

    conn = get_connection()
    c = conn.cursor()
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    df = pd.DataFrame(rows, columns=["ID", "FECHA", "PAGADO_A", "CONCEPTO", "OBSERVACIONES", "RESPONSABLE", "IMPORTE"])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Gastos")
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="gastos_filtrados.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------ PDF IMPRESIÓN ------------------

@app.route("/exportar_pdf", methods=["GET"])
def exportar_pdf():
    # mismos filtros
    fecha_inicio = request.args.get("fecha_inicio")
    fecha_fin = request.args.get("fecha_fin")
    concepto = request.args.get("concepto")
    responsable = request.args.get("responsable")

    query = "SELECT * FROM gastos WHERE 1=1"
    params = []

    if fecha_inicio:
        query += " AND fecha >= ?"
        params.append(fecha_inicio)
    if fecha_fin:
        query += " AND fecha <= ?"
        params.append(fecha_fin)
    if concepto:
        query += " AND concepto LIKE ?"
        params.append(f"%{concepto}%")
    if responsable:
        query += " AND responsable = ?"
        params.append(responsable)

    conn = get_connection()
    c = conn.cursor()
    c.execute(query, params)
    rows = c.fetchall()
    conn.close()

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    y = height - 50
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(50, y, "Gastos filtrados")
    y -= 30
    pdf.setFont("Helvetica", 10)

    for r in rows:
        linea = f"{r[0]} | {r[1]} | {r[2]} | {r[3]} | {r[4]} | {r[5]} | {r[6]}"
        pdf.drawString(50, y, linea)
        y -= 15
        if y < 50:
            pdf.showPage()
            y = height - 50
            pdf.setFont("Helvetica", 10)

    pdf.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="gastos_filtrados.pdf",
        mimetype="application/pdf"
    )

# ------------------ MAIN ------------------

if __name__ == "__main__":
    init_db()
    app.run(debug=True)
