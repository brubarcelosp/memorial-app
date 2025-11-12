pip install Flask Flask-SQLAlchemy Flask-Login

from flask import Blueprint, render_template, request, send_file, redirect, url_for, flash
import os
from .builders import generate_docx_from_form, generate_excel_from_form

bp = Blueprint("main", __name__, template_folder="templates", static_folder="../static")

@bp.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@bp.route("/generate", methods=["POST"])
def generate():
    form = request.form.to_dict()
    uploaded = request.files.getlist("files")
    files = [(f.filename, f.read()) for f in uploaded if f and f.filename]
    try:
        out_path = generate_docx_from_form(form, files)
    except Exception as e:
        flash(str(e))
        return redirect(url_for("main.index"))
    return send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))

@bp.route("/download-excel", methods=["POST"])
def download_excel():
    form = request.form.to_dict()
    uploaded = request.files.getlist("files")
    files = [(f.filename, f.read()) for f in uploaded if f and f.filename]
    try:
        out_path = generate_excel_from_form(form, files)
    except Exception as e:
        flash(str(e))
        return redirect(url_for("main.index"))
    return send_file(out_path, as_attachment=True, download_name=os.path.basename(out_path))