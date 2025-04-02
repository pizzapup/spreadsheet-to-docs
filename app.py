from flask import (
    Flask,
    request,
    redirect,
    url_for,
    Response,
    render_template,
    send_file,
)
import pandas as pd
from werkzeug.utils import secure_filename
from docx import Document
from io import BytesIO
import logging
import zipfile

# Set up logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)

# Configure allowed file extensions
ALLOWED_EXTENSIONS = {"xlsx", "csv"}


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if "file" not in request.files:
            logging.warning("No file part in the request.")
            return "No file part in the request."
        file = request.files["file"]
        if file.filename == "":
            logging.warning("No file selected.")
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
                ) = process_uploaded_file(file)
                return render_template(
                    "file_preview.html",
                    table_html=table_html,
                    data=df.to_json(orient="records"),
                    column_names=column_names,
                    column_feedback=column_feedback,
                    default_filename_template=default_filename_template,
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


@app.route("/generate_docs", methods=["POST"])
def generate_docs():
    """Generate Word documents from the uploaded data."""
    try:
        data = request.form.get("data")
        filename_template = request.form.get("filename_template", "")
        zip_filename = request.form.get("zip_filename", "").strip()
        null_handling = request.form.get("null_handling", "omit")
        null_value = (
            request.form.get("null_value", "N/A") if null_handling == "fill" else None
        )

        # Default to "Documents.zip" if no name is provided
        if not zip_filename:
            zip_filename = "Documents.zip"
        elif not zip_filename.lower().endswith(".zip"):
            zip_filename += ".zip"

        if not data:
            raise ValueError("No data provided for document generation.")

        df = pd.read_json(data)

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for i, (_, row) in enumerate(df.iterrows()):
                doc = Document()
                doc.add_heading("Generated Document", level=1)
                for col in df.columns:
                    value = row[col]
                    if pd.isnull(value) or value == "":
                        if null_handling == "omit":
                            continue
                        else:
                            value = null_value
                    doc.add_paragraph(f"{col}: {value}")
                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                # Generate filename using the template
                filename = filename_template
                for col in df.columns:
                    placeholder = f"{{{col}}}"
                    filename = filename.replace(placeholder, str(row[col]))
                filename = filename.replace("{index}", str(i))
                if not filename.strip():
                    filename = f"Document_{i}.docx"
                else:
                    filename += ".docx"

                zip_file.writestr(filename, doc_io.read())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name=zip_filename,
        )
    except Exception as e:
        logging.error(f"Error generating documents: {e}")
        return "An error occurred while generating the documents."


def process_uploaded_file(file):
    """Process the uploaded file and return the DataFrame and metadata."""
    df = init_file(file)
    table_html = df.head().to_html(classes="table table-striped", index=False)
    column_names = list(df.columns)
    column_feedback = analyze_columns_for_filenames(df)

    # Only include feedback if there are issues
    column_feedback = {
        col: feedback for col, feedback in column_feedback.items() if feedback
    }

    # Determine default filename template
    if "First and Middle Name" in df.columns and "Last Name" in df.columns:
        default_filename_template = "{First and Middle Name}-{Last Name}"
    else:
        default_filename_template = "Document-{index}"

    return df, table_html, column_names, column_feedback, default_filename_template


def analyze_columns_for_filenames(df):
    """Analyze columns for filename usability and provide feedback."""
    feedback = {}
    for col in df.columns:
        long_values = (
            df[col].apply(lambda x: len(str(x)) > 80 if pd.notnull(x) else False).sum()
        )
        invalid_chars = (
            df[col]
            .apply(
                lambda x: (
                    any(c in str(x) for c in r'\/:*?"<>|') if pd.notnull(x) else False
                )
            )
            .sum()
        )

        if long_values > 0:
            feedback[col] = (
                f"Column '{col}' contains {long_values} value(s) longer than 80 characters. These will be truncated to 60 characters."
            )
            df[col] = df[col].apply(lambda x: str(x)[:60] if pd.notnull(x) else x)

        if invalid_chars > 0:
            feedback[col] = (
                f"Column '{col}' contains {invalid_chars} value(s) with invalid characters. These will be replaced with underscores (_)."
            )
            df[col] = df[col].apply(
                lambda x: (
                    "".join(c if c not in r'\/:*?"<>|' else "_" for c in str(x))
                    if pd.notnull(x)
                    else x
                )
            )

    return feedback


def init_file(file):
    """Initialize the uploaded file and return a DataFrame."""
    try:
        logging.info("Initializing file in memory.")
        if file.filename.endswith(".xlsx"):
            df = pd.read_excel(file)
        elif file.filename.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            raise ValueError("Invalid file type. Only .xlsx and .csv are supported.")

        df.columns = df.columns.str.strip()

        # Check if the DataFrame is empty
        if df.empty:
            raise ValueError("The uploaded file is empty.")

        # Log a warning if required columns are missing, but do not throw an error
        required_columns = {"First and Middle Name", "Last Name"}
        missing_columns = required_columns - set(df.columns)
        if missing_columns:
            logging.warning(
                f"Missing required columns: {', '.join(missing_columns)}. Proceeding without them."
            )

        return df

    except Exception as e:
        logging.error(f"Error initializing file: {e}")
        raise ValueError(
            "Error initializing the file. Please check the file format and content."
        )


if __name__ == "__main__":
    app.run(debug=True)
