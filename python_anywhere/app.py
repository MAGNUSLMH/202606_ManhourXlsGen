import io
import os
import shutil
import zipfile
from datetime import datetime
from zoneinfo import ZoneInfo

from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

from .code01 import build_workbook
from .code02 import highlight_workbook

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "/tmp"  # PythonAnywhere temporary directory
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # Limit upload size to 16MB


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if "file" not in request.files:
            return "No file part in the request", 400

        file = request.files["file"]

        if file.filename == "":
            return "No selected file", 400

        action_manhour = request.form.get("manhour")
        action_highlight = request.form.get("highlight")

        if not (action_manhour or action_highlight):
            return "No action selected", 400

        if file and file.filename.endswith((".xlsx", ".xlsm")):
            raw_filename = secure_filename(file.filename)

            # Save the master upload file
            master_input_path = os.path.join(
                app.config["UPLOAD_FOLDER"], f"master_{raw_filename}"
            )
            file.save(master_input_path)

            timestamp = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime(
                "%Y%m%d%H%M%S"
            )

            # Hardcoded configurations
            staffID = ["488704", "429465", "676139"]
            jCol1, jCol2 = "A", "I"
            nCol1, nCol2 = ["C", "E", "G"], ["J", "L", "N"]
            oCol1, oCol2 = ["D", "F", "H"], ["K", "M", "O"]
            dates = list(range(4, 949, 59))

            files_to_cleanup = [master_input_path]

            try:
                # BOTH CHECKBOXES TICKED -> Generate ZIP
                if action_manhour and action_highlight:
                    # 1. Process Manhour file
                    work_manhour_input = os.path.join(
                        app.config["UPLOAD_FOLDER"], f"work_mh_{timestamp}_{raw_filename}"
                    )
                    shutil.copyfile(master_input_path, work_manhour_input)
                    files_to_cleanup.append(work_manhour_input)

                    manhour_filename = f"{timestamp}_Manhour.xlsx"
                    manhour_output_path = os.path.join(
                        app.config["UPLOAD_FOLDER"], manhour_filename
                    )
                    files_to_cleanup.append(manhour_output_path)

                    build_workbook(
                        filename=manhour_output_path,
                        staffID=staffID,
                        dates=dates,
                        timesheet_path=work_manhour_input,
                        jCol1=jCol1,
                        jCol2=jCol2,
                        nCol1=nCol1,
                        nCol2=nCol2,
                        oCol1=oCol1,
                        oCol2=oCol2,
                    )

                    # 2. Process Highlight file (in-place on its own working copy)
                    work_highlight_input = os.path.join(
                        app.config["UPLOAD_FOLDER"], f"work_hl_{timestamp}_{raw_filename}"
                    )
                    shutil.copyfile(master_input_path, work_highlight_input)
                    files_to_cleanup.append(work_highlight_input)

                    highlight_filename = f"{timestamp}_Highlight.xlsx"

                    highlight_workbook(
                        filename=highlight_filename,
                        staffID=staffID,
                        dates=dates,
                        timesheet_path=work_highlight_input,
                        jCol1=jCol1,
                        jCol2=jCol2,
                        nCol1=nCol1,
                        nCol2=nCol2,
                        oCol1=oCol1,
                        oCol2=oCol2,
                    )

                    # 3. Create ZIP archive in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        zf.write(manhour_output_path, arcname=manhour_filename)
                        zf.write(work_highlight_input, arcname=highlight_filename)

                    zip_buffer.seek(0)
                    download_name = f"{timestamp}_Processed_Timesheets.zip"

                    return send_file(
                        zip_buffer,
                        as_attachment=True,
                        download_name=download_name,
                        mimetype="application/zip",
                    )

                # ONLY MANHOUR CHECKED
                elif action_manhour:
                    work_input_path = os.path.join(
                        app.config["UPLOAD_FOLDER"], f"work_{timestamp}_{raw_filename}"
                    )
                    shutil.copyfile(master_input_path, work_input_path)
                    files_to_cleanup.append(work_input_path)

                    download_name = f"{timestamp}_Manhour.xlsx"
                    output_path = os.path.join(
                        app.config["UPLOAD_FOLDER"], download_name
                    )
                    files_to_cleanup.append(output_path)

                    build_workbook(
                        filename=output_path,
                        staffID=staffID,
                        dates=dates,
                        timesheet_path=work_input_path,
                        jCol1=jCol1,
                        jCol2=jCol2,
                        nCol1=nCol1,
                        nCol2=nCol2,
                        oCol1=oCol1,
                        oCol2=oCol2,
                    )

                    with open(output_path, "rb") as f:
                        file_data = io.BytesIO(f.read())

                    return send_file(
                        file_data,
                        as_attachment=True,
                        download_name=download_name,
                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                # ONLY HIGHLIGHT CHECKED
                elif action_highlight:
                    work_input_path = os.path.join(
                        app.config["UPLOAD_FOLDER"], f"work_{timestamp}_{raw_filename}"
                    )
                    shutil.copyfile(master_input_path, work_input_path)
                    files_to_cleanup.append(work_input_path)

                    download_name = f"{timestamp}_Highlight.xlsx"

                    highlight_workbook(
                        filename=download_name,
                        staffID=staffID,
                        dates=dates,
                        timesheet_path=work_input_path,
                        jCol1=jCol1,
                        jCol2=jCol2,
                        nCol1=nCol1,
                        nCol2=nCol2,
                        oCol1=oCol1,
                        oCol2=oCol2,
                    )

                    with open(work_input_path, "rb") as f:
                        file_data = io.BytesIO(f.read())

                    return send_file(
                        file_data,
                        as_attachment=True,
                        download_name=download_name,
                        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            except Exception as e:
                return f"An error occurred during processing: {e!s}", 500
            finally:
                # Clean up all registered temporary files
                for path in files_to_cleanup:
                    if path and os.path.exists(path):
                        os.remove(path)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)