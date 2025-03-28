from flask import Flask, request, redirect, url_for
import pandas as pd
import os
from docx import Document

app = Flask(__name__)

# Ensure the output directory exists
if not os.path.exists("output"):
    os.makedirs("output")


@app.route("/")
def index():
    return """
    <!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <title>Upload File</title>
      </head>
      <body>
        <div class="container">
          <h1 class="mt-5">Upload Excel File</h1>
          <form method="post" action="/upload" enctype="multipart/form-data">
            <div class="form-group">
              <label for="file">Choose Excel file</label>
              <input type="file" class="form-control-file" id="file" name="file">
            </div>
            <button type="submit" class="btn btn-primary">Upload</button>
          </form>
        </div>
      </body>
    </html>
    """


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return redirect(request.url)

    file = request.files["file"]
    if file.filename == "":
        return redirect(request.url)

    if file:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()

        # Iterate through each row in the dataframe
        for index, row in df.iterrows():
            # Create a new word document
            doc = Document()
            # Iterate through each column in the row
            for col in df.columns:
                # If the answer is not empty
                if not pd.isnull(row[col]):
                    # Add the question in bold
                    question = doc.add_paragraph()
                    question.add_run(f"{col}: ").bold = True
                    # Add the answer in normal text
                    doc.add_paragraph(str(row[col]))
            # Save the document to the output directory with the filename '[Last_Name] ['First and Middle Name].docx'
            doc.save(f'output/{row["Last Name"]} {row["First and Middle Name"]}.docx')

        return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
