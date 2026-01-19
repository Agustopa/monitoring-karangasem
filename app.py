from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import sqlite3
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
app.secret_key = "karangasem_super_app"

def get_db_connection():
    # Menggunakan path yang aman untuk database
    conn = sqlite3.connect("data.db")
    conn.row_factory = sqlite3.Row
    return conn

# Inisialisasi Database saat pertama dijalankan
with get_db_connection() as conn:
    conn.execute("CREATE TABLE IF NOT EXISTS data_tm (id INTEGER PRIMARY KEY AUTOINCREMENT, nama_tm TEXT, nama_konsumen TEXT, tipe_bayar TEXT, nominal INTEGER, tanggal TEXT)")
    conn.execute("CREATE TABLE IF NOT EXISTS data_ao (id INTEGER PRIMARY KEY AUTOINCREMENT, nama_nasabah TEXT, nama_ao TEXT, status_tagihan TEXT, tanggal TEXT)")
    conn.execute("CREATE TABLE IF NOT EXISTS targets (nama_user TEXT UNIQUE, target_crm INTEGER, target_sales INTEGER)")
    conn.commit()

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/tm", methods=["GET", "POST"])
def tm():
    if request.method == "POST":
        conn = get_db_connection()
        conn.execute("INSERT INTO data_tm (nama_tm, nama_konsumen, tipe_bayar, nominal, tanggal) VALUES (?,?,?,?,?)",
                     (request.form['nama_tm'], request.form['nama_konsumen'], request.form['tipe'], request.form['nominal'], datetime.now().strftime('%Y-%m-%d')))
        conn.commit()
        conn.close()
        flash("Data Marketing Berhasil Dikirim!")
        return redirect(url_for('tm'))
    return render_template("tm.html")

@app.route("/ao", methods=["GET", "POST"])
def ao():
    if request.method == "POST":
        conn = get_db_connection()
        conn.execute("INSERT INTO data_ao (nama_nasabah, nama_ao, status_tagihan, tanggal) VALUES (?,?,?,?)",
                     (request.form['nama_nasabah'], request.form['nama_ao'], request.form['status'], datetime.now().strftime('%Y-%m-%d')))
        conn.commit()
        conn.close()
        flash("Laporan Tagihan AO Berhasil Disimpan!")
        return redirect(url_for('ao'))
    return render_template("ao.html")

@app.route("/bm")
def bm():
    conn = get_db_connection()
    # Menghitung realisasi vs target untuk Dashboard
    list_tm = conn.execute("""
        SELECT t.nama_user, t.target_crm, t.target_sales, 
        COUNT(d.id) as realisasi_crm, SUM(d.nominal) as realisasi_sales
        FROM targets t LEFT JOIN data_tm d ON t.nama_user = d.nama_tm
        GROUP BY t.nama_user
    """).fetchall()
    list_ao = conn.execute("SELECT nama_ao, status_tagihan, COUNT(*) as jml FROM data_ao GROUP BY nama_ao, status_tagihan").fetchall()
    conn.close()
    return render_template("bm.html", list_tm=list_tm, list_ao=list_ao)

@app.route("/set_target", methods=["GET", "POST"])
def set_target():
    if request.method == "POST":
        conn = get_db_connection()
        conn.execute("INSERT INTO targets (nama_user, target_crm, target_sales) VALUES (?,?,?) ON CONFLICT(nama_user) DO UPDATE SET target_crm=excluded.target_crm, target_sales=excluded.target_sales",
                     (request.form['nama'].lower(), request.form['t_crm'], request.form['t_sales']))
        conn.commit()
        conn.close()
        return redirect(url_for('bm'))
    return render_template("set_target.html")

@app.route("/export/<tipe>")
def export(tipe):
    conn = get_db_connection()
    table = "data_tm" if tipe == "tm" else "data_ao"
    df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
    conn.close()
    filename = f"Laporan_{tipe}.xlsx"
    df.to_excel(filename, index=False)
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)