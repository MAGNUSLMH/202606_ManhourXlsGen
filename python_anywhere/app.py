import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

# Import your existing function from code01.py
from .code01 import build_workbook

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "/tmp"  # PythonAnywhere temporary directory
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # Limit upload size to 16MB


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        # Check if the post request has the file part
        if "file" not in request.files:
            return "No file part in the request", 400

        file = request.files["file"]

        if file.filename == "":
            return "No selected file", 400

        if file and file.filename.endswith((".xlsx", ".xlsm")):
            # Secure the filename and save it temporarily
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(input_path)

            # Define output path
            output_filename = "processed_timesheet.xlsx"
            output_path = os.path.join(
                app.config["UPLOAD_FOLDER"], output_filename
            )

            # Hardcoded configurations from your original main()
            staffID = ["488704", "429465", "676139"]
            jCol1, jCol2 = "A", "I"
            nCol1, nCol2 = ["C", "E", "G"], ["J", "L", "N"]
            oCol1, oCol2 = ["D", "F", "H"], ["K", "M", "O"]
            dates = list(range(4, 949, 59))

            try:
                # Run your processing logic
                build_workbook(
                    filename=output_path,
                    staffID=staffID,
                    dates=dates,
                    timesheet_path=input_path,
                    jCol1=jCol1,
                    jCol2=jCol2,
                    nCol1=nCol1,
                    nCol2=nCol2,
                    oCol1=oCol1,
                    oCol2=oCol2,
                )

                # Send the generated file back to the user as an automatic download
                return send_file(
                    output_path,
                    as_attachment=True,
                    download_name="sample.xlsx",
                )

            except Exception as e:
                return f"An error occurred during processing: {str(e)}", 500
            finally:
                # Clean up uploaded files after processing to save space
                if os.path.exists(input_path):
                    os.remove(input_path)
                if os.path.exists(output_path):
                    os.remove(output_path)

    # Simple HTML interface returned on GET request
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)