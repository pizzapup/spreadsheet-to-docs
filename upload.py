from flask import Blueprint, request, render_template
from werkzeug.utils import secure_filename
import logging
from process import process_uploaded_file

ALLOWED_EXTENSIONS = {"xlsx", "csv", "xls"}

upload_blueprint = Blueprint("upload", __name__)


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@upload_blueprint.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if "file" not in request.files:
            return "No file part in the request."
        file = request.files["file"]
        if file.filename == "":
            return "No file selected."
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            try:
                (
                    df,
                    table_html,
                    column_names,
                    column_feedback,
                    default_filename_template,
                    has_null_values,  # Capture this value
                ) = process_uploaded_file(file)
                return render_template(
                    "process_form.html",
                    table_html=table_html,
                    data=df.to_json(orient="records"),
                    column_names=column_names,
                    column_feedback=column_feedback,
                    default_filename_template=default_filename_template,
                    has_null_values=has_null_values,  # Pass to template
                )
            except ValueError as e:
                logging.error(f"Error processing the file {filename}: {e}")
                return str(e)
            except Exception as e:
                logging.error(f"Unexpected error: {e}")
                return "An unexpected error occurred while processing the file."
        else:
            logging.warning("File type not allowed.")
            return "Invalid file type."
    return render_template("upload_form.html")
