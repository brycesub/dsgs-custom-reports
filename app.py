import io
import os
import zipfile
from datetime import date

from flask import Flask, abort, render_template, request, send_file

from reports.pipeline import generate as generate_pipeline
from reports.pipeline_pdf import generate as generate_pipeline_pdf
from reports.plans import generate as generate_plans
from reports.plans_pdf import generate as generate_plans_pdf
from research import list_studies, resolve_artifact

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit
app.config["RESEARCH_DIR"] = os.environ.get("RESEARCH_DIR")  # arts-funder-research checkout


def _get_csv_file():
    f = request.files.get("csv_file")
    if not f or not f.filename:
        return None, ("No file uploaded.", 400)
    if not f.filename.lower().endswith(".csv"):
        return None, ("Please upload a CSV file.", 400)
    return f, None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/research")
def research_index():
    return render_template("research.html", studies=list_studies(app.config["RESEARCH_DIR"]))


@app.route("/research/<slug>/<path:artifact>")
def research_artifact(slug, artifact):
    path = resolve_artifact(app.config["RESEARCH_DIR"], slug, artifact)
    if path is None:
        abort(404)
    return send_file(path, mimetype="text/html")


@app.route("/generate/plans", methods=["POST"])
def generate_plans_report():
    f, err = _get_csv_file()
    if err:
        return err

    try:
        results = generate_plans(f.read())
    except ValueError as e:
        return str(e), 400
    except Exception:
        app.logger.exception("Unexpected error generating plans report")
        return "An unexpected error occurred. Please try again.", 500

    if len(results) == 1:
        client, xlsx_bytes = results[0]
        return send_file(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{client}_report.xlsx",
        )

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for client, xlsx_bytes in results:
            zf.writestr(f"{client}_report.xlsx", xlsx_bytes)
    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="reports.zip",
    )


@app.route("/generate/plans/pdf", methods=["POST"])
def generate_plans_pdf_report():
    f, err = _get_csv_file()
    if err:
        return err

    try:
        results = generate_plans_pdf(f.read())
    except ValueError as e:
        return str(e), 400
    except Exception:
        app.logger.exception("Unexpected error generating plans PDF report")
        return "An unexpected error occurred. Please try again.", 500

    if len(results) == 1:
        client, pdf_bytes = results[0]
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"{client}_report.pdf",
        )

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for client, pdf_bytes in results:
            zf.writestr(f"{client}_report.pdf", pdf_bytes)
    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name="reports.zip",
    )


@app.route("/generate/pipeline/pdf", methods=["POST"])
def generate_pipeline_pdf_report():
    f, err = _get_csv_file()
    if err:
        return err

    try:
        client, pdf_bytes = generate_pipeline_pdf(f.read(), f.filename)
    except ValueError as e:
        return str(e), 400
    except Exception:
        app.logger.exception("Unexpected error generating pipeline PDF report")
        return "An unexpected error occurred. Please try again.", 500

    today = date.today().isoformat()
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"Pipeline_{client}_{today}.pdf",
    )


@app.route("/generate/pipeline", methods=["POST"])
def generate_pipeline_report():
    f, err = _get_csv_file()
    if err:
        return err

    try:
        client, xlsx_bytes = generate_pipeline(f.read(), f.filename)
    except ValueError as e:
        return str(e), 400
    except Exception:
        app.logger.exception("Unexpected error generating pipeline report")
        return "An unexpected error occurred. Please try again.", 500

    today = date.today().isoformat()
    return send_file(
        io.BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Pipeline_{client}_{today}.xlsx",
    )


if __name__ == "__main__":
    app.run(debug=True)
