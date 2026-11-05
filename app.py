
import streamlit as st
import fitz
import docx
from pptx import Presentation
from openpyxl import load_workbook
import io
import re

st.set_page_config(
    page_title="AcademicQuery AI",
    page_icon="📚",
    layout="wide"
)

st.title("📚 AcademicQuery AI")
st.write("Upload your academic documents and ask questions about them.")
st.info("Free document-based question answering — no API key or payment required.")

def extract_text(uploaded_file):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    text = ""

    try:
        if name.endswith(".pdf"):
            pdf = fitz.open(stream=data, filetype="pdf")
            text = "\n".join(page.get_text() for page in pdf)

        elif name.endswith(".docx"):
            document = docx.Document(io.BytesIO(data))
            text = "\n".join(p.text for p in document.paragraphs)
            for table in document.tables:
                for row in table.rows:
                    text += "\n" + " | ".join(
                        cell.text for cell in row.cells
                    )

        elif name.endswith(".pptx"):
            presentation = Presentation(io.BytesIO(data))
            slides = []
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        slides.append(shape.text)
            text = "\n".join(slides)

        elif name.endswith(".xlsx"):
            workbook = load_workbook(
                io.BytesIO(data), read_only=True, data_only=True
            )
            rows = []
            for sheet in workbook.worksheets:
                rows.append("Sheet: " + sheet.title)
                for row in sheet.iter_rows(values_only=True):
                    values = [str(v) for v in row if v is not None]
                    if values:
                        rows.append(" | ".join(values))
            text = "\n".join(rows)

        elif name.endswith(".txt"):
            text = data.decode("utf-8", errors="ignore")

        elif name.endswith(".csv"):
            text = data.decode("utf-8", errors="ignore")

        else:
            return ""

    except Exception as error:
        st.warning(f"Could not read {uploaded_file.name}: {error}")
        return ""

    return text


def tokenize(text):
    words = re.findall(r"\b[a-zA-Z0-9]+\b", text.lower())
    stopwords = {
        "the", "is", "are", "was", "were", "a", "an", "and",
        "or", "of", "to", "in", "on", "for", "with", "what",
        "who", "when", "where", "why", "how", "explain",
        "tell", "me", "about", "please", "this", "that"
    }
    return [word for word in words if word not in stopwords]


def make_chunks(text, chunk_size=1000):
    words = text.split()
    chunks = []
    for start in range(0, len(words), 150):
        chunk = " ".join(words[start:start + chunk_size // 5])
        if chunk.strip():
            chunks.append(chunk)
    return chunks


def find_answers(question, documents):
    question_words = set(tokenize(question))

    if not question_words:
        return []

    results = []

    for filename, text in documents.items():
        for chunk in make_chunks(text):
            chunk_words = set(tokenize(chunk))
            matches = question_words.intersection(chunk_words)

            if matches:
                score = len(matches) / len(question_words)
                results.append((score, filename, chunk))

    results.sort(key=lambda item: item[0], reverse=True)
    return results[:5]


st.sidebar.header("📁 Upload Documents")

uploaded_files = st.sidebar.file_uploader(
    "Choose PDF, Word, PowerPoint, Excel, TXT or CSV files",
    type=["pdf", "docx", "pptx", "xlsx", "txt", "csv"],
    accept_multiple_files=True
)

documents = {}

if uploaded_files:
    with st.spinner("Reading your documents..."):
        for uploaded_file in uploaded_files:
            extracted = extract_text(uploaded_file)
            if extracted.strip():
                documents[uploaded_file.name] = extracted

    st.sidebar.success(f"{len(documents)} document(s) ready.")

    for filename in documents:
        st.sidebar.write("✅ " + filename)

st.divider()

st.subheader("💬 Ask a Question")

if not documents:
    st.write("👈 Upload one or more documents using the sidebar to begin.")
else:
    with st.form("question_form"):
        question = st.text_input(
            "Enter your question",
            placeholder="Example: What is machine learning?"
        )
        submitted = st.form_submit_button("Search Documents")

    if submitted:
        if not question.strip():
            st.warning("Please enter a question.")
        else:
            with st.spinner("Searching your documents..."):
                answers = find_answers(question, documents)

            if answers:
                st.success("Found relevant information in your documents.")

                for index, (score, filename, passage) in enumerate(answers, 1):
                    st.markdown(f"### Result {index}")
                    st.caption(
                        f"Source: {filename} | Keyword match: {score:.0%}"
                    )
                    st.write(passage)
                    st.divider()
            else:
                st.warning(
                    "No matching information was found. "
                    "Try different keywords or check whether your document "
                    "contains selectable text."
                )

st.caption(
    "AcademicQuery AI | Free document search | No paid AI API required"
)
