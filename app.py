#!/usr/bin/env python3
"""
Excel Report Automator - Flask Backend API
------------------------------------------
Provides web API endpoints for uploading CSV/Excel datasets,
cleaning and validating records, and downloading formatted Excel reports.
"""

import os
import uuid
import sys
import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS

# Import core business logic from excel_report_automator
from excel_report_automator import (
    read_data_file,
    standardize_columns,
    detect_category_column,
    clean_and_validate,
    generate_summary_dataframes,
    create_formatted_excel
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp_uploads")
REPORTS_DIR = os.path.join(BASE_DIR, "reports_output")

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static")
)
CORS(app)

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)


@app.route("/")
def index():
    """Serve the React front-end application directly."""
    return send_file(os.path.join(BASE_DIR, "templates", "index.html"))


@app.route("/api/process", methods=["POST"])
def process_file():
    """
    Handle CSV/Excel file upload, process data, generate Excel report,
    and return JSON stats + previews.
    """
    if "file" not in request.files:
        return jsonify({"success": False, "error": "No file uploaded in request."}), 400

    file = request.files["file"]
    if not file or file.filename.strip() == "":
        return jsonify({"success": False, "error": "Selected file is empty."}), 400

    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()
    if ext not in [".csv", ".xlsx", ".xls"]:
        return jsonify({
            "success": False,
            "error": f"Invalid file format '{ext}'. Only .csv, .xlsx, and .xls files are supported."
        }), 400

    unique_id = uuid.uuid4().hex[:8]
    temp_input_path = os.path.join(TEMP_DIR, f"upload_{unique_id}{ext}")

    try:
        file.save(temp_input_path)

        if os.path.getsize(temp_input_path) == 0:
            os.remove(temp_input_path)
            return jsonify({"success": False, "error": "Uploaded file is 0 bytes."}), 400

        # Read and process dataset
        raw_df = read_data_file(temp_input_path)
        df = standardize_columns(raw_df.copy())
        
        # User options if passed
        user_cat = request.form.get("category_col", None)
        category_col = detect_category_column(df, user_category_col=user_cat)

        req_cols_str = request.form.get("required_cols", "")
        req_cols_list = [c.strip() for c in req_cols_str.split(",") if c.strip()] if req_cols_str else None

        # Clean & validate
        cleaned_df, errors_df, stats = clean_and_validate(df, required_cols_input=req_cols_list)
        kpi_df, category_summary_df = generate_summary_dataframes(cleaned_df, stats, category_col)

        # Generate formatted Excel file
        report_filename = f"report_{os.path.splitext(filename)[0]}_{unique_id}.xlsx"
        report_output_path = os.path.join(REPORTS_DIR, report_filename)
        
        create_formatted_excel(report_output_path, cleaned_df, errors_df, kpi_df, category_summary_df, category_col)

        # Remove temporary uploaded file
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)

        # Format previews for React JSON response
        # Clean null values for JSON serialization
        cleaned_json = cleaned_df.head(100).fillna("").to_dict(orient="records")
        errors_json = errors_df.head(100).fillna("").to_dict(orient="records")
        kpi_json = kpi_df.fillna("").to_dict(orient="records")
        cat_summary_json = category_summary_df.fillna("").to_dict(orient="records")

        return jsonify({
            "success": True,
            "report_filename": report_filename,
            "filename": filename,
            "stats": stats,
            "category_col": category_col,
            "columns": list(cleaned_df.columns),
            "error_columns": list(errors_df.columns),
            "cleaned_data": cleaned_json,
            "errors_data": errors_json,
            "kpi_data": kpi_json,
            "category_summary_data": cat_summary_json
        })

    except Exception as e:
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)
        return jsonify({"success": False, "error": f"Processing error: {str(e)}"}), 500


@app.route("/api/download/<filename>", methods=["GET"])
def download_report(filename):
    """Serve generated Excel report for download."""
    # Security check: prevent path traversal
    safe_filename = os.path.basename(filename)
    file_path = os.path.join(REPORTS_DIR, safe_filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "Report file not found."}), 404

    return send_file(
        file_path,
        as_attachment=True,
        download_name=safe_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    print("Starting Excel Report Automator Web App on http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=True)
