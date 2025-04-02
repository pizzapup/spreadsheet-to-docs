import pandas as pd
import logging


def process_uploaded_file(file):
    """Process the uploaded file and return the DataFrame and metadata."""
    df = init_file(file)
    table_html = df.head().to_html(classes="table table-striped", index=False)
    column_names = list(df.columns)
    column_feedback = init_column_feedback(df)

    column_feedback = {
        col: feedback for col, feedback in column_feedback.items() if feedback
    }

    default_filename_template = get_default_filename_template(df)
    has_null_values = df.isnull().any().any()  # Check if any null values exist

    return (
        df,
        table_html,
        column_names,
        column_feedback,
        default_filename_template,
        has_null_values,  # Include this in the return
    )


def init_column_feedback(df):
    feedback = {}
    # feedback: column_name -> info
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
        null_values = df[col].isnull().sum()
        if long_values > 0:
            feedback[col] = (
                f"Values longer than 80 characters will be truncated to 60 characters. (Effects {long_values} cell{'s' if long_values > 1 else ''})"
            )
            df[col] = df[col].apply(lambda x: str(x)[:60] if pd.notnull(x) else x)

        if invalid_chars > 0:
            feedback[col] = (
                f"** This column contains invalid characters. These characters will be replaced with underscores. (Effects {invalid_chars} cell{'s' if invalid_chars > 1 else ''})"
            )
            df[col] = df[col].apply(
                lambda x: (
                    "".join(c if c not in r'\/:*?"<>|' else "_" for c in str(x))
                    if pd.notnull(x)
                    else x
                )
            )

        if null_values > 0:
            feedback[col] = (
                f"** This column contains null or empty values. This could lead to unexpected results in the filenames. (Effects {null_values} cell{'s' if null_values > 1 else ''})"
            )
    return feedback


def get_default_filename_template(df):
    if "First and Middle Name" in df.columns and "Last Name" in df.columns:
        return "{First and Middle Name}-{Last Name}"
    else:
        return "Document-{index}"


def init_file(file):
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

        return df

    except Exception as e:
        logging.error(f"Error initializing file: {e}")
        raise ValueError(
            "Error initializing the file. Please check the file format and content."
        )
