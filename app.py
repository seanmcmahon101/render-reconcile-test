import os
import logging
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from werkzeug.utils import secure_filename
import openpyxl
import google.generativeai as genai
import markdown
from docx import Document  # For creating Word documents
from io import BytesIO  # For handling DOCX in memory for download

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'your_default_secret_key')  # For session management, ensure FLASK_SECRET_KEY is set in Render

# --- Configuration ---
API_KEY = os.getenv("GEMINI_API_KEY")  # Load API key from environment variable
DEFAULT_MODEL_NAME = 'gemini-2.0-flash'  # Default model
THINKING_MODEL_NAME = 'gemini-2.0-flash-thinking-exp'  # Thinking model
SYSTEM_PROMPT = """You are an expert in analyzing Excel spreadsheets and explaining their purpose, structure, and formulas in a clear and concise manner.
Your goal is to provide a detailed explanation of the spreadsheet's content, including the meaning of each column, the data structure, and the purpose of any formulas.
You should format your explanations in Markdown, using headings, bullet points, and code blocks where appropriate to enhance readability.
Focus on providing insights that would be helpful to someone unfamiliar with the spreadsheet."""
PROMPT_PREFIX = "The Excel sheet contains the following information in a structured way:\n"
PROMPT_SUFFIX = "\nPlease provide a detailed explanation in Markdown format. Explain what this sheet is about, what each column/section represents, and how the data is structured and what its purpose might be. If there are formulas, explain their logic in simple terms. Structure your answer with headings, bullet points, and code blocks where appropriate for formulas or data examples to enhance readability."
UPLOAD_FOLDER = 'uploads'  # Folder to temporarily store uploads
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}  # Allowed Excel file extensions
DEFAULT_DOCX_FILENAME = "excel_explanation.docx"  # Define here for consistency


os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Ensure upload folder exists

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def configure_api():
    """Configures the Gemini API with the API key."""
    if not API_KEY:
        logging.error("API_KEY environment variable not set.")
        return False
    try:
        genai.configure(api_key=API_KEY)
        return True
    except Exception as e:
        logging.error(f"Error configuring Gemini API: {e}")
        return False


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def load_excel_data(file_path):
    """Loads data from the Excel file using openpyxl."""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=False)
        sheet = wb.active
        return sheet
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        return None


def build_prompt(sheet):
    """Builds the prompt for the Gemini API based on the Excel sheet data."""
    prompt_content = ""
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            if cell.value is not None or cell.comment is not None:
                cell_info = ""
                if cell.data_type == 'f':
                    cell_info = f"formula '{cell.value}'"
                elif cell.value is not None:
                    cell_info = f"value '{cell.value}'"
                else:
                    cell_info = "no value"

                comment_text = ""
                if cell.comment:
                    comment_text = f" with comment '{cell.comment.text.strip().replace('\n', ' ')}'"

                prompt_content += f"- Cell {cell.coordinate} has {cell_info}{comment_text}.\n"

    full_prompt = PROMPT_PREFIX + prompt_content + PROMPT_SUFFIX
    logging.info("Prompt built successfully.")
    return full_prompt


def get_explanation_from_gemini(prompt, model_name):  # Added model_name parameter
    """Gets explanation from Gemini API."""
    model = genai.GenerativeModel(model_name)  # Use the selected model_name
    try:
        response = model.generate_content(prompt)
        explanation = response.text
        logging.info(f"Explanation received from Gemini API using model: {model_name}")  # Log model name
        return explanation
    except Exception as e:
        logging.error(f"Error communicating with Gemini API: {e}")
        return None


def export_to_docx(explanation):
    """Exports the explanation to a DOCX file in memory and returns BytesIO object."""
    doc = Document()
    for line in explanation.splitlines():
        doc.add_paragraph(line)

    docx_stream = BytesIO()
    try:
        doc.save(docx_stream)
        docx_stream.seek(0)  # Rewind to the beginning of the stream
        logging.info("Explanation exported to DOCX in memory.")
        return docx_stream
    except Exception as e:
        logging.error(f"Error exporting to DOCX: {e}")
        return None


@app.route('/', methods=['GET', 'POST'])
def index():
    """Handles the main application logic."""
    explanation_html = None
    docx_stream = None
    error = None
    model_name = DEFAULT_MODEL_NAME  # Default model

    if request.method == 'POST':
        if 'excel_file' not in request.files:
            error = 'No file part'
            return render_template('index.html', error=error, model_name=model_name)  # Pass model_name to template

        file = request.files['excel_file']

        if file.filename == '':
            error = 'No selected file'
            return render_template('index.html', error=error, model_name=model_name)  # Pass model_name to template

        if file and allowed_file(file.filename):
            try:
                filename = secure_filename(file.filename)
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(file_path)  # Save temporarily

                # Determine model name from form selection
                selected_model = request.form.get('model_select')
                if selected_model == 'thinking':
                    model_name = THINKING_MODEL_NAME
                else:
                    model_name = DEFAULT_MODEL_NAME  # Default if no selection or standard is selected

                sheet = load_excel_data(file_path)
                if sheet:
                    prompt = build_prompt(sheet)
                    explanation_markdown = get_explanation_from_gemini(prompt, model_name)  # Pass model_name

                    if explanation_markdown:
                        explanation_html = markdown.markdown(explanation_markdown)
                        session['explanation_markdown'] = explanation_markdown  # Store for DOCX export
                    else:
                        error = "Failed to get explanation from Gemini API."
                else:
                    error = "Failed to load Excel data."
            except Exception as e:
                error = f"An error occurred: {e}"
            finally:
                os.remove(file_path)  # Clean up uploaded file

        else:
            error = 'Invalid file type. Allowed types are xlsx, xls'

    return render_template('index.html', explanation_html=explanation_html, error=error, model_name=model_name)  # Pass model_name to template


@app.route('/export_docx')
def export_docx_route():
    """Exports the explanation to DOCX format and allows download."""
    explanation_markdown = session.get('explanation_markdown')
    if not explanation_markdown:
        return "No explanation available to export.", 400  # Or redirect with error message

    docx_stream = export_to_docx(explanation_markdown)
    if docx_stream:
        return send_file(
            docx_stream,
            as_attachment=True,
            download_name=DEFAULT_DOCX_FILENAME,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        return "Error exporting to DOCX.", 500  # Or redirect with error message


if __name__ == '__main__':
    if configure_api():  # Only start the app if API is configured
        app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))  # Use PORT env var for Render
