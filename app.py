import io
from flask import Flask, render_template, request, send_file
from reports.goh_plans import generate

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate_report():
    f = request.files.get("csv_file")
    if not f or not f.filename:
        return "No file uploaded.", 400
    if not f.filename.lower().endswith(".csv"):
        return "Please upload a CSV file.", 400

    try:
        zip_bytes = generate(f.read())
    except ValueError as e:
        return str(e), 400
    except Exception as e:
        return f"Error generating report: {e}", 500

    return send_file(
        io.BytesIO(zip_bytes),
        mimetype="application/zip",
        as_attachment=True,
        download_name="reports.zip",
    )


if __name__ == "__main__":
    app.run(debug=True)
