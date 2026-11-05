import streamlit as st
import fitz
import docx
from pptx import Presentation
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import inch
import io
import re
import csv
import html
from collections import Counter
from datetime import datetime


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="SCSVMV AcademicQuery AI",
    page_icon="📚",
    layout="wide"
)

st.markdown("""
<style>
.main-title {
    font-size: 34px;
    font-weight: 800;
    color: #2457A6;
}
.subtitle {
    color: #64748b;
    font-size: 15px;
}
div.stButton > button {
    border-radius: 10px;
    min-height: 42px;
    font-weight: 600;
}
.study-card {
    padding: 16px;
    border-radius: 12px;
    border: 1px solid #dbe3ef;
    margin-bottom: 10px;
}
.small-note {
    color: #64748b;
    font-size: 13px;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {
    "documents": {},
    "notes": "",
    "selected_topic": "",
    "last_question": "",
    "search_results": [],
    "seminar_content": "",
    "project_guide": "",
    "page": "Home"
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# TEXT PROCESSING
# ============================================================

STOPWORDS = {
    "the", "is", "are", "was", "were", "a", "an", "and",
    "or", "of", "to", "in", "on", "for", "with", "what",
    "who", "when", "where", "why", "how", "explain",
    "tell", "me", "about", "please", "this", "that",
    "from", "which", "their", "there", "these", "those",
    "can", "could", "would", "should", "into", "also",
    "give", "write", "describe", "define", "discuss"
}


def tokenize(text):
    words = re.findall(r"\b[a-zA-Z0-9]+\b", text.lower())
    return [word for word in words if word not in STOPWORDS and len(word) > 1]


def clean_text(text):
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def split_sentences(text):
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    return [
        sentence.strip()
        for sentence in sentences
        if len(sentence.strip()) > 20
    ]


def extract_text(uploaded_file):
    filename = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    try:
        if filename.endswith(".pdf"):
            pdf = fitz.open(stream=data, filetype="pdf")
            pages = []
            for page_number, page in enumerate(pdf, start=1):
                page_text = page.get_text()
                if page_text.strip():
                    pages.append(
                        f"\n[Page {page_number}]\n{page_text}"
                    )
            return clean_text("\n".join(pages))

        if filename.endswith(".docx"):
            document = docx.Document(io.BytesIO(data))
            parts = [p.text for p in document.paragraphs]

            for table in document.tables:
                for row in table.rows:
                    parts.append(
                        " | ".join(cell.text for cell in row.cells)
                    )

            return clean_text("\n".join(parts))

        if filename.endswith(".pptx"):
            presentation = Presentation(io.BytesIO(data))
            parts = []

            for slide_number, slide in enumerate(
                presentation.slides, start=1
            ):
                parts.append(f"\nSlide {slide_number}")
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        parts.append(shape.text)

            return clean_text("\n".join(parts))

        if filename.endswith(".xlsx"):
            workbook = load_workbook(
                io.BytesIO(data),
                read_only=True,
                data_only=True
            )
            parts = []

            for sheet in workbook.worksheets:
                parts.append(f"\nSheet: {sheet.title}")
                for row in sheet.iter_rows(values_only=True):
                    values = [
                        str(value)
                        for value in row
                        if value is not None
                    ]
                    if values:
                        parts.append(" | ".join(values))

            return clean_text("\n".join(parts))

        if filename.endswith(".csv"):
            decoded = data.decode("utf-8-sig", errors="ignore")
            reader = csv.reader(io.StringIO(decoded))
            return clean_text(
                "\n".join(" | ".join(row) for row in reader)
            )

        if filename.endswith(".txt"):
            return clean_text(
                data.decode("utf-8-sig", errors="ignore")
            )

    except Exception as error:
        st.error(f"Could not read {uploaded_file.name}: {error}")
        return ""

    return ""


def make_chunks(text, words_per_chunk=180, overlap=30):
    words = text.split()
    chunks = []

    if not words:
        return chunks

    step = max(1, words_per_chunk - overlap)

    for start in range(0, len(words), step):
        chunk = " ".join(words[start:start + words_per_chunk])
        if chunk.strip():
            chunks.append(chunk)

    return chunks


def search_documents(question, documents, limit=8):
    question_words = set(tokenize(question))

    if not question_words:
        return []

    results = []

    for filename, text in documents.items():
        for chunk in make_chunks(text):
            chunk_words = set(tokenize(chunk))
            matched = question_words.intersection(chunk_words)

            if not matched:
                continue

            score = len(matched) / max(1, len(question_words))

            results.append({
                "score": score,
                "filename": filename,
                "text": chunk,
                "matched": sorted(matched)
            })

    results.sort(
        key=lambda result: (result["score"], len(result["text"])),
        reverse=True
    )

    return results[:limit]


def find_relevant_sentences(topic, documents, limit=12):
    topic_words = set(tokenize(topic))
    results = []

    if not topic_words:
        return []

    for filename, text in documents.items():
        for sentence in split_sentences(text):
            sentence_words = set(tokenize(sentence))
            matched = topic_words.intersection(sentence_words)

            if matched:
                score = len(matched) / len(topic_words)
                results.append({
                    "score": score,
                    "filename": filename,
                    "sentence": sentence
                })

    results.sort(
        key=lambda item: item["score"],
        reverse=True
    )

    return results[:limit]


# ============================================================
# QUESTION PAPER ANALYSIS
# ============================================================

QUESTION_START = re.compile(
    r"^\s*(?:Q\.?\s*)?(?:\d{1,2}\s*[.)]|\([a-zA-Z0-9]+\))\s+"
)


def extract_questions(text):
    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    questions = []

    for line in lines:
        if len(line) < 12:
            continue

        looks_like_question = (
            "?" in line
            or QUESTION_START.match(line)
            or line.lower().startswith(
                ("explain ", "define ", "describe ", "discuss ",
                 "what ", "write ", "compare ", "differentiate ",
                 "illustrate ", "list ", "state ", "derive ")
            )
        )

        if looks_like_question:
            cleaned = re.sub(
                r"^\s*(?:Q\.?\s*)?(?:\d{1,2}\s*[.)]|\([a-zA-Z0-9]+\))\s*",
                "",
                line
            )
            if len(cleaned) >= 10:
                questions.append(cleaned)

    return questions


def question_similarity(question_a, question_b):
    words_a = set(tokenize(question_a))
    words_b = set(tokenize(question_b))

    if not words_a or not words_b:
        return 0

    intersection = len(words_a.intersection(words_b))
    union = len(words_a.union(words_b))

    return intersection / union if union else 0


def analyse_previous_papers(documents):
    paper_questions = {}

    for filename, text in documents.items():
        questions = extract_questions(text)

        if questions:
            paper_questions[filename] = questions

    groups = []

    for filename, questions in paper_questions.items():
        for question in questions:
            matched_group = None

            for group in groups:
                similarity = question_similarity(
                    question, group["representative"]
                )

                if similarity >= 0.55:
                    matched_group = group
                    break

            if matched_group:
                matched_group["count"] += 1
                matched_group["papers"].add(filename)
                matched_group["versions"].append(question)
            else:
                groups.append({
                    "representative": question,
                    "count": 1,
                    "papers": {filename},
                    "versions": [question]
                })

    groups.sort(
        key=lambda group: group["count"],
        reverse=True
    )

    return paper_questions, groups


# ============================================================
# STUDY NOTES AND ANSWER FORMATS
# ============================================================

def make_main_points(topic, documents, limit=10):
    relevant = find_relevant_sentences(
        topic, documents, limit=limit * 2
    )

    points = []
    seen = set()

    for item in relevant:
        sentence = item["sentence"]
        key = sentence.lower()

        if key not in seen:
            seen.add(key)
            points.append({
                "text": sentence,
                "source": item["filename"]
            })

        if len(points) >= limit:
            break

    return points


def make_study_notes(topic, documents):
    points = make_main_points(topic, documents, 12)

    if not points:
        return (
            f"Topic: {topic}\n\n"
            "No matching passages were found in the uploaded documents. "
            "Try another topic name or upload relevant study material."
        )

    lines = [
        f"TOPIC: {topic}",
        "",
        "1. IMPORTANT POINTS",
        ""
    ]

    for index, point in enumerate(points, start=1):
        lines.append(f"{index}. {point['text']}")
        lines.append(f"   Source: {point['source']}")

    lines.extend([
        "",
        "2. QUICK REVISION",
        "",
        "Read the definition and important terms.",
        "Remember the numbered points.",
        "Practise explaining the topic using your own words.",
        "",
        "3. EXAM PREPARATION",
        "",
        "Begin with a definition or introduction.",
        "Explain the main points under separate headings.",
        "Add an example or diagram when appropriate.",
        "Finish with a short conclusion.",
        "",
        "Note: These notes are extracted and organized from uploaded "
        "documents. Check the original material for full context."
    ])

    return "\n".join(lines)


def answer_template(topic, marks, documents):
    points = make_main_points(topic, documents, 15)

    if not points:
        return (
            f"QUESTION TOPIC: {topic}\n\n"
            "No supporting information was found in the uploaded documents. "
            "Upload relevant material before preparing this answer."
        )

    lines = [
        f"QUESTION: Explain {topic}",
        f"Suggested answer format: {marks} marks",
        "",
        "1. INTRODUCTION / DEFINITION",
        points[0]["text"],
        "",
        "2. MAIN EXPLANATION",
        ""
    ]

    for index, point in enumerate(points[1:], start=1):
        lines.append(f"   {index}. {point['text']}")

    lines.extend([
        "",
        "3. EXAMPLE OR DIAGRAM",
        "Add a suitable example or labelled diagram from your study material.",
        "",
        "4. CONCLUSION",
        "Summarise the main idea in two or three clear sentences.",
        "",
        "EXAM PRESENTATION CHECKLIST",
        "- Write the question heading.",
        "- Underline important technical terms.",
        "- Use separate headings and numbered points.",
        "- Draw a labelled diagram if relevant.",
        "- Check that every part of the question is answered.",
        "",
        "SOURCE MATERIAL"
    ])

    for point in points:
        lines.append(f"- {point['source']}")

    lines.extend([
        "",
        "Important: This is a structured draft based on uploaded material. "
        "Review it for completeness and follow your university's syllabus "
        "and marking scheme."
    ])

    return "\n".join(lines)


# ============================================================
# PDF CREATOR
# ============================================================

def create_pdf(title, content):
    buffer = io.BytesIO()

    pdf = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=45,
        leftMargin=45,
        topMargin=50,
        bottomMargin=45
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "CustomTitle",
        parent=styles["Title"],
        fontSize=19,
        leading=24,
        textColor=colors.HexColor("#2457A6"),
        alignment=TA_CENTER,
        spaceAfter=18
    )

    heading_style = ParagraphStyle(
        "CustomHeading",
        parent=styles["Heading2"],
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#2457A6"),
        spaceBefore=10,
        spaceAfter=6
    )

    body_style = ParagraphStyle(
        "CustomBody",
        parent=styles["BodyText"],
        fontSize=9.5,
        leading=14,
        spaceAfter=6
    )

    story = [
        Paragraph(html.escape(title), title_style),
        Paragraph(
            "SCSVMV AcademicQuery AI | Student Study Notes",
            body_style
        ),
        Spacer(1, 10)
    ]

    for line in content.splitlines():
        safe_line = html.escape(line)

        if not safe_line.strip():
            story.append(Spacer(1, 5))
        elif (
            len(line) < 85
            and (
                line.isupper()
                or re.match(r"^\d+[.)]\s", line)
            )
        ):
            story.append(Paragraph(safe_line, heading_style))
        else:
            story.append(Paragraph(safe_line, body_style))

    pdf.build(story)
    buffer.seek(0)

    return buffer.getvalue()


# ============================================================
# PROJECT RESOURCE LIBRARY
# ============================================================

PROJECT_RESOURCES = [
    {
        "id": "DEV-001",
        "category": "Development tools",
        "name": "Visual Studio Code",
        "purpose": "Code editor for Python, HTML, CSS and JavaScript.",
        "url": "https://code.visualstudio.com/"
    },
    {
        "id": "DEV-002",
        "category": "Development tools",
        "name": "PyCharm",
        "purpose": "Python development and debugging.",
        "url": "https://www.jetbrains.com/pycharm/"
    },
    {
        "id": "DEV-003",
        "category": "Development tools",
        "name": "Jupyter Notebook",
        "purpose": "Python experiments and data analysis.",
        "url": "https://jupyter.org/"
    },
    {
        "id": "DEV-004",
        "category": "Browser coding",
        "name": "Replit",
        "purpose": "Browser-based coding environment.",
        "url": "https://replit.com/"
    },
    {
        "id": "DEV-005",
        "category": "Browser coding",
        "name": "Google Colab",
        "purpose": "Browser-based Python notebooks.",
        "url": "https://colab.research.google.com/"
    },
    {
        "id": "DEV-006",
        "category": "Browser coding",
        "name": "GitHub Codespaces",
        "purpose": "Cloud development environment.",
        "url": "https://github.com/features/codespaces"
    },
    {
        "id": "DEP-001",
        "category": "Deployment",
        "name": "Render",
        "purpose": "Host Python and web applications.",
        "url": "https://render.com/"
    },
    {
        "id": "DEP-002",
        "category": "Deployment",
        "name": "Vercel",
        "purpose": "Deploy frontend and supported full-stack applications.",
        "url": "https://vercel.com/"
    },
    {
        "id": "DEP-003",
        "category": "Deployment",
        "name": "Netlify",
        "purpose": "Deploy static websites and supported web projects.",
        "url": "https://www.netlify.com/"
    },
    {
        "id": "DEP-004",
        "category": "Deployment",
        "name": "Streamlit Community Cloud",
        "purpose": "Deploy Streamlit applications.",
        "url": "https://share.streamlit.io/"
    },
    {
        "id": "DEP-005",
        "category": "Deployment",
        "name": "GitHub Pages",
        "purpose": "Host static HTML, CSS and JavaScript websites.",
        "url": "https://pages.github.com/"
    },
    {
        "id": "DEP-006",
        "category": "Deployment",
        "name": "PythonAnywhere",
        "purpose": "Host Python web applications.",
        "url": "https://www.pythonanywhere.com/"
    }
]


def project_steps(project_type):
    if project_type == "Streamlit":
        return """
PROJECT EXECUTION AND DEPLOYMENT

1. Install Python.
2. Install VS Code or PyCharm.
3. Create a project folder.
4. Create app.py.
5. Create requirements.txt.
6. Open the terminal in the project folder.
7. Install dependencies using:
   pip install -r requirements.txt
8. Run locally using:
   streamlit run app.py
9. Open the local URL displayed in the terminal.
10. Test the application.
11. Upload the project to GitHub.
12. Create a new Streamlit or Render deployment.
13. Connect the GitHub repository.
14. Set the start command:
    streamlit run app.py --server.port $PORT --server.address 0.0.0.0
15. Deploy and test the public website URL.

Check the hosting platform's current requirements and free-tier limits.
"""

    if project_type == "Flask":
        return """
PROJECT EXECUTION AND DEPLOYMENT

1. Install Python.
2. Install VS Code or PyCharm.
3. Create a project folder.
4. Create app.py and templates folder.
5. Install Flask:
   pip install flask
6. Run locally:
   python app.py
7. Open the local URL in a browser.
8. Test all routes and forms.
9. Upload the project to GitHub.
10. Select a Python web hosting platform.
11. Configure the start command according to that platform.
12. Deploy and test the public URL.
"""

    if project_type == "Static website":
        return """
PROJECT EXECUTION AND DEPLOYMENT

1. Install VS Code.
2. Create a project folder.
3. Create index.html.
4. Create style.css.
5. Create script.js if needed.
6. Open index.html in a browser.
7. Test links, forms and page layout.
8. Upload the files to GitHub.
9. Choose GitHub Pages, Netlify or Vercel.
10. Connect or upload the project.
11. Publish the website.
12. Test the public URL.
"""

    return """
PROJECT EXECUTION AND DEPLOYMENT

1. Identify the programming language.
2. Install the required development tools.
3. Create the project folder.
4. Add the source code and dependency files.
5. Install the required libraries.
6. Run the project locally.
7. Test the output.
8. Fix errors.
9. Upload the source code to GitHub.
10. Choose a hosting platform that supports the framework.
11. Configure build and start commands.
12. Deploy and test the live output.
"""


# ============================================================
# SIDEBAR NAVIGATION
# ============================================================

st.sidebar.markdown("## 🎓 SCSVMV")
st.sidebar.caption("AcademicQuery AI")

pages = [
    "Home",
    "Student Chatbot",
    "Student Upload",
    "Previous Question Papers",
    "Exam Answer Builder",
    "Quick Learning",
    "My Notes & PDF Creator",
    "Academic Calculator",
    "Project Development",
    "Seminar Assistant",
    "Admin Upload",
    "About Project"
]

selected_page = st.sidebar.radio(
    "Navigation",
    pages,
    index=pages.index(st.session_state.page)
    if st.session_state.page in pages else 0
)

st.session_state.page = selected_page

st.sidebar.divider()
st.sidebar.caption("Free Python-based document processing")
st.sidebar.caption("No paid AI API or external AI compiler")


# ============================================================
# HOME PAGE
# ============================================================

if selected_page == "Home":

    st.markdown(
        '<div class="main-title">SCSVMV AcademicQuery AI</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Your examination preparation, '
        'academic and project learning assistant.</div>',
        unsafe_allow_html=True
    )

    st.divider()

    st.info(
        "Upload academic materials, explore previous-year questions, "
        "prepare structured answers and create your own study PDFs."
    )

    col1, col2 = st.columns(2)

    home_cards = [
        ("📚 Student Chatbot", "Search and understand uploaded study materials."),
        ("📄 Previous Question Papers", "Find repeated and similar questions."),
        ("✍️ Exam Answer Builder", "Prepare answers according to marks."),
        ("⏱️ Quick Learning", "Revise topics in 5, 15 or 20 minutes."),
        ("📝 My Notes & PDF Creator", "Edit notes and create personal PDFs."),
        ("🎓 Academic Calculator", "Calculate marks, SGPA and CGPA."),
        ("💻 Project Development", "Learn coding, execution and deployment."),
        ("🎤 Seminar Assistant", "Prepare seminar topics and presentations.")
    ]

    for index, (heading, description) in enumerate(home_cards):
        target_col = col1 if index % 2 == 0 else col2

        with target_col:
            with st.container(border=True):
                st.subheader(heading)
                st.write(description)


# ============================================================
# STUDENT AND ADMIN DOCUMENT UPLOAD
# ============================================================

elif selected_page in ["Student Upload", "Admin Upload"]:

    st.title("📁 " + selected_page)

    st.write(
        "Upload academic documents for topic search, question analysis "
        "and exam preparation."
    )

    if selected_page == "Admin Upload":
        st.warning(
            "This is a demonstration admin page. "
            "It does not include secure login or access control. "
            "Do not use it for private staff documents until authentication "
            "and permissions are implemented."
        )

    upload_category = st.selectbox(
        "Document category",
        [
            "Study Material",
            "Previous Year Question Paper",
            "Arrear Question Paper",
            "Staff Question Paper",
            "Syllabus",
            "Other"
        ]
    )

    uploaded_files = st.file_uploader(
        "Select your documents",
        type=["pdf", "docx", "pptx", "xlsx", "txt", "csv"],
        accept_multiple_files=True,
        key="main_upload"
    )

    if st.button("Process Uploaded Documents"):

        if not uploaded_files:
            st.warning("Please select at least one document.")
        else:
            count = 0

            with st.spinner("Reading and indexing documents..."):
                for uploaded_file in uploaded_files:
                    text = extract_text(uploaded_file)

                    if text.strip():
                        st.session_state.documents[
                            uploaded_file.name
                        ] = {
                            "text": text,
                            "category": upload_category,
                            "uploaded_at": datetime.now().strftime(
                                "%Y-%m-%d %H:%M"
                            )
                        }
                        count += 1

            if count:
                st.success(f"{count} document(s) processed successfully.")
            else:
                st.error("No readable text was found in the selected files.")

    st.subheader("Available documents")

    if st.session_state.documents:
        for filename, details in st.session_state.documents.items():
            with st.expander(filename):
                st.write("Category:", details["category"])
                st.write("Uploaded:", details["uploaded_at"])
                st.write(
                    "Extracted characters:",
                    len(details["text"])
                )
    else:
        st.info("No documents have been uploaded in this session.")


# ============================================================
# DOCUMENT SEARCH
# ============================================================

elif selected_page == "Student Chatbot":

    st.title("🤖 Academic AI Chat")
    st.write(
        "Ask questions about your uploaded academic documents. "
        "Answers are based on matching passages from those documents."
    )

    if not st.session_state.documents:
        st.warning("Upload documents first using Student Upload.")
    else:
        document_texts = {
            name: details["text"]
            for name, details in st.session_state.documents.items()
        }

        question = st.text_input(
            "Enter your question",
            placeholder="Example: Explain semantic role labeling"
        )

        if st.button("Search Academic Materials"):

            if not question.strip():
                st.warning("Enter a question.")
            else:
                results = search_documents(
                    question, document_texts
                )

                st.session_state.search_results = results
                st.session_state.last_question = question

        if st.session_state.search_results:
            st.subheader("Relevant information")

            for index, result in enumerate(
                st.session_state.search_results, start=1
            ):
                with st.container(border=True):
                    st.markdown(f"**Result {index}**")
                    st.caption(
                        f"Source: {result['filename']} | "
                        f"Keyword match: {result['score']:.0%}"
                    )
                    st.write(result["text"])

                    if st.button(
                        "Add this passage to My Notes",
                        key=f"add_passage_{index}"
                    ):
                        st.session_state.notes += (
                            "\n\n" + result["text"]
                        )
                        st.success("Passage added to My Notes.")


# ============================================================
# PREVIOUS QUESTION PAPERS
# ============================================================

elif selected_page == "Previous Question Papers":

    st.title("📄 Previous-Year Question Paper Analysis")

    st.write(
        "Upload previous-year papers, arrear papers or staff question "
        "papers to find frequently repeated and similar questions."
    )

    paper_files = st.file_uploader(
        "Upload question papers",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
        key="paper_upload"
    )

    if st.button("Analyse Question Papers"):

        if not paper_files:
            st.warning("Upload at least one question paper.")
        else:
            with st.spinner("Extracting and comparing questions..."):
                paper_texts = {}

                for file in paper_files:
                    extracted = extract_text(file)
                    if extracted.strip():
                        paper_texts[file.name] = extracted

                paper_questions, groups = analyse_previous_papers(
                    paper_texts
                )

                st.session_state["paper_questions"] = paper_questions
                st.session_state["question_groups"] = groups

    if "question_groups" in st.session_state:

        groups = st.session_state.question_groups

        st.subheader("Frequently repeated questions")

        if not groups:
            st.warning(
                "No question-like lines were identified. "
                "Check the extracted text or upload clearer papers."
            )

        for index, group in enumerate(groups, start=1):
            with st.container(border=True):
                st.markdown(
                    f"### {index}. {group['representative']}"
                )
                st.metric(
                    "Occurrences in uploaded papers",
                    group["count"]
                )
                st.write(
                    "Found in:",
                    ", ".join(sorted(group["papers"]))
                )

                if len(group["versions"]) > 1:
                    with st.expander("Similar question wordings"):
                        for version in group["versions"]:
                            st.write("- " + version)

        st.warning(
            "Frequency is calculated only from the papers uploaded here. "
            "It does not predict or guarantee future examination questions."
        )


# ============================================================
# EXAM ANSWER BUILDER
# ============================================================

elif selected_page == "Exam Answer Builder":

    st.title("✍️ Exam Answer Builder")

    st.write(
        "Prepare an answer structure from your uploaded study materials."
    )

    if not st.session_state.documents:
        st.warning("Upload study materials first.")
    else:
        topic = st.text_input(
            "Enter the question or topic",
            placeholder="Example: Explain the applications of NLP"
        )

        marks = st.selectbox(
            "Select marks",
            [2, 3, 5, 8, 10, 13, 15, 16, 20]
        )

        if st.button("Prepare Exam Answer"):

            if not topic.strip():
                st.warning("Enter a topic.")
            else:
                document_texts = {
                    name: details["text"]
                    for name, details in st.session_state.documents.items()
                }

                answer = answer_template(
                    topic, marks, document_texts
                )

                st.session_state.notes = answer
                st.success("Answer draft prepared.")

        if st.session_state.notes:
            st.text_area(
                "Editable answer",
                key="notes",
                height=450
            )

            st.download_button(
                "Download answer as TXT",
                data=st.session_state.notes,
                file_name="exam_answer.txt",
                mime="text/plain"
            )


# ============================================================
# QUICK LEARNING
# ============================================================

elif selected_page == "Quick Learning":

    st.title("⏱️ Quick Learning")

    st.write(
        "Learn and revise any topic through simple, "
        "step-by-step learning sessions."
    )

    duration = st.radio(
        "Choose your learning time",
        [5, 15, 20],
        horizontal=True,
        format_func=lambda x: f"{x} Minutes"
    )

    topic = st.text_input(
        "Enter topic to learn",
        placeholder="Example: Machine Learning"
    )

    if st.button("Start Quick Learning"):

        if not topic.strip():
            st.warning("Please enter a topic.")

        elif not st.session_state.documents:
            st.warning(
                "Please upload your study materials first."
            )

        else:
            document_texts = {
                name: details["text"]
                for name, details in st.session_state.documents.items()
            }

            points = make_main_points(
                topic,
                document_texts,
                15
            )

            if not points:
                st.warning(
                    "No matching information was found. "
                    "Try another topic or upload relevant material."
                )

            else:
                st.divider()

                st.subheader(
                    f"📖 {topic} — {duration}-Minute Learning Session"
                )

                st.success(
                    f"Your {duration}-minute learning session is ready!"
                )

                if duration == 5:
                    learning_steps = [
                        (
                            "Step 1 — Understand the Topic (1 minute)",
                            "Read the first important point and understand "
                            "what the topic means."
                        ),
                        (
                            "Step 2 — Learn Key Points (2 minutes)",
                            "Read the main concepts and remember the "
                            "important information."
                        ),
                        (
                            "Step 3 — Review Supporting Details (1 minute)",
                            "Review additional details from your uploaded "
                            "study material."
                        ),
                        (
                            "Step 4 — Quick Memory Test (1 minute)",
                            "Close your notes and try to explain the topic "
                            "using your own words."
                        )
                    ]

                elif duration == 15:
                    learning_steps = [
                        (
                            "Step 1 — Topic Introduction (3 minutes)",
                            "Understand the topic and identify its main "
                            "definition or meaning."
                        ),
                        (
                            "Step 2 — Understand Main Concepts (5 minutes)",
                            "Study the important points one by one. "
                            "Understand how each point relates to the topic."
                        ),
                        (
                            "Step 3 — Review Details and Examples (4 minutes)",
                            "Read supporting explanations, examples and "
                            "other relevant information in your material."
                        ),
                        (
                            "Step 4 — Revision and Self-Test (3 minutes)",
                            "Recall the important points and practise "
                            "explaining the topic without looking at "
                            "your notes."
                        )
                    ]

                else:
                    learning_steps = [
                        (
                            "Step 1 — Understand the Topic (4 minutes)",
                            "Read the introduction and understand the "
                            "basic meaning of the topic."
                        ),
                        (
                            "Step 2 — Learn Important Concepts (6 minutes)",
                            "Study the major concepts carefully. "
                            "Understand each point instead of memorising "
                            "sentences."
                        ),
                        (
                            "Step 3 — Study Supporting Information (4 minutes)",
                            "Review additional explanations, examples, "
                            "applications and diagrams available in "
                            "your uploaded documents."
                        ),
                        (
                            "Step 4 — Exam Preparation (3 minutes)",
                            "Practise writing a short answer using "
                            "headings and numbered points."
                        ),
                        (
                            "Step 5 — Final Revision (3 minutes)",
                            "Recall the definition, main concepts and "
                            "important keywords."
                        )
                    ]

                st.divider()
                st.subheader("🪜 Your Learning Steps")

                for index, (heading, instruction) in enumerate(
                    learning_steps,
                    start=1
                ):
                    with st.container(border=True):
                        st.markdown(f"### {heading}")
                        st.write(instruction)

                        if index == 1:
                            st.caption(
                                "Start by understanding the topic."
                            )

                        elif index == len(learning_steps):
                            st.caption(
                                "Finish by checking what you remember."
                            )

                st.divider()
                st.subheader("📌 Important Learning Points")

                for index, point in enumerate(points, start=1):
                    with st.container(border=True):
                        st.markdown(
                            f"**Point {index}:** {point['text']}"
                        )
                        st.caption(
                            "Study material: " + point["source"]
                        )

                st.divider()
                st.subheader("✅ Quick Revision Checklist")

                st.write(
                    "Use this checklist after completing your "
                    "learning session."
                )

                st.checkbox(
                    "I understood the meaning of the topic.",
                    key="quick_check_1"
                )

                st.checkbox(
                    "I learned the important points.",
                    key="quick_check_2"
                )

                st.checkbox(
                    "I reviewed the supporting information.",
                    key="quick_check_3"
                )

                st.checkbox(
                    "I can explain the topic in my own words.",
                    key="quick_check_4"
                )

                st.info(
                    "The learning steps organize information from "
                    "your uploaded documents. This feature does not "
                    "invent definitions or examples when they are "
                    "missing from the source material."
                )


# ============================================================
# PERSONAL NOTES AND PDF CREATOR
# ============================================================

elif selected_page == "My Notes & PDF Creator":

    st.title("📝 My Notes and PDF Creator")

    st.write(
        "Create, edit and download your own study notes."
    )

    st.text_input(
        "Document title",
        key="personal_document_title",
        value=st.session_state.get(
            "personal_document_title", "My Study Notes"
        )
    )

    st.text_area(
        "Write or edit your notes",
        key="notes",
        height=450,
        placeholder="Write your study notes here..."
    )

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "Download as TXT",
            data=st.session_state.notes,
            file_name="my_study_notes.txt",
            mime="text/plain"
        )

    with col2:
        if st.button("Prepare PDF"):
            if not st.session_state.notes.strip():
                st.warning("Write some notes first.")
            else:
                try:
                    pdf_data = create_pdf(
                        st.session_state.get(
                            "personal_document_title",
                            "My Study Notes"
                        ),
                        st.session_state.notes
                    )

                    st.download_button(
                        "Download My PDF",
                        data=pdf_data,
                        file_name="my_study_notes.pdf",
                        mime="application/pdf"
                    )
                except Exception as error:
                    st.error(f"PDF creation failed: {error}")


# ============================================================
# ACADEMIC CALCULATOR
# ============================================================

elif selected_page == "Academic Calculator":

    st.title("🎓 Academic Performance Calculator")

    st.write(
        "Calculate marks and estimate GPA using the official rules "
        "provided by your university."
    )

    calculator_mode = st.radio(
        "Choose calculation",
        ["Total marks", "SGPA / CGPA"],
        horizontal=True
    )

    if calculator_mode == "Total marks":

        internal_max = st.number_input(
            "Maximum internal marks", min_value=1.0, value=25.0
        )
        internal_obtained = st.number_input(
            "Internal marks obtained",
            min_value=0.0,
            max_value=float(internal_max),
            value=20.0
        )

        external_max = st.number_input(
            "Maximum external marks", min_value=1.0, value=75.0
        )
        external_obtained = st.number_input(
            "External marks obtained",
            min_value=0.0,
            max_value=float(external_max),
            value=50.0
        )

        if st.button("Calculate Total"):

            total_max = internal_max + external_max
            total_obtained = (
                internal_obtained + external_obtained
            )

            percentage = (
                total_obtained / total_max * 100
                if total_max else 0
            )

            st.metric(
                "Total marks",
                f"{total_obtained:g} / {total_max:g}"
            )
            st.metric(
                "Percentage",
                f"{percentage:.2f}%"
            )

    else:

        st.info(
            "Enter the credit and grade point for each subject. "
            "Use the grade points and rules specified in your official "
            "university regulations."
        )

        number_subjects = st.number_input(
            "Number of subjects",
            min_value=1,
            max_value=20,
            value=5
        )

        subjects = []

        for index in range(int(number_subjects)):
            st.markdown(f"**Subject {index + 1}**")

            col1, col2 = st.columns(2)

            with col1:
                credit = st.number_input(
                    f"Credits {index + 1}",
                    min_value=0.0,
                    value=3.0,
                    key=f"credit_{index}"
                )

            with col2:
                grade_point = st.number_input(
                    f"Grade point {index + 1}",
                    min_value=0.0,
                    max_value=10.0,
                    value=8.0,
                    key=f"grade_{index}"
                )

            subjects.append((credit, grade_point))

        if st.button("Calculate SGPA"):

            total_credits = sum(
                credit for credit, grade in subjects
            )

            weighted_points = sum(
                credit * grade
                for credit, grade in subjects
            )

            if total_credits > 0:
                sgpa = weighted_points / total_credits
                st.metric("Estimated SGPA", f"{sgpa:.2f}")
            else:
                st.error("Total credits must be greater than zero.")

        st.caption(
            "For CGPA, enter all applicable semester subjects and credits "
            "using the university's official cumulative calculation rules. "
            "This basic calculator does not automatically determine "
            "SCSVMV grades or percentage conversion."
        )


# ============================================================
# PROJECT DEVELOPMENT ASSISTANT
# ============================================================

elif selected_page == "Project Development":

    st.title("💻 Beginner Project Development Assistant")

    st.write(
        "Learn how to plan, code, run, test and deploy a project."
    )

    project_type = st.selectbox(
        "Choose project type",
        ["Streamlit", "Flask", "Static website", "Other Python project"]
    )

    project_name = st.text_input(
        "Enter your project title",
        placeholder="Example: Student Management System"
    )

    skill_level = st.selectbox(
        "Your programming experience",
        ["Beginner", "Basic knowledge", "Intermediate"]
    )

    editor = st.radio(
        "Choose development tool",
        ["VS Code", "PyCharm", "Jupyter Notebook"],
        horizontal=True
    )

    if st.button("Create My Project Guide"):

        guide = f"""
PROJECT TITLE: {project_name or "My Project"}

PROJECT TYPE: {project_type}
EXPERIENCE LEVEL: {skill_level}
DEVELOPMENT TOOL: {editor}

STEP 1: PROJECT PLANNING

Write the problem statement.
Identify the users.
List the required features.
Define the expected output.

STEP 2: SOFTWARE SETUP

Install Python if required.
Install {editor}.
Create a new project folder.
Open the folder in your editor.

STEP 3: PROJECT STRUCTURE

Create separate files for application code,
styles, templates, data and documentation
as required by your project type.

STEP 4: INSTALLATION

Create a requirements.txt file for Python libraries.
Install the required libraries using pip.

STEP 5: EXECUTION

{project_steps(project_type)}

STEP 6: TESTING

Test every feature.
Check input validation.
Check incorrect input.
Check page layout.
Check errors and fix them.

STEP 7: DOCUMENTATION

Prepare project abstract, objectives, methodology,
system requirements, diagrams, screenshots,
testing results and conclusion.

STEP 8: FINAL REVIEW

Confirm that the application works locally.
Confirm that the deployed application works.
Keep a copy of the project source code.
"""

        st.session_state.project_guide = guide

    if st.session_state.project_guide:
        st.text_area(
            "Your project development guide",
            value=st.session_state.project_guide,
            height=450
        )

        st.download_button(
            "Download Project Guide",
            data=st.session_state.project_guide,
            file_name="project_development_guide.txt",
            mime="text/plain"
        )

    st.divider()

    st.subheader("Development and deployment resources")

    search_resource = st.text_input(
        "Search tools or deployment platforms",
        placeholder="Example: Python, deployment, browser"
    )

    for resource in PROJECT_RESOURCES:
        searchable = (
            resource["id"] + " " +
            resource["category"] + " " +
            resource["name"] + " " +
            resource["purpose"]
        ).lower()

        if search_resource.lower() in searchable:
            with st.container(border=True):
                st.markdown(
                    f"**{resource['id']} — {resource['name']}**"
                )
                st.caption(resource["category"])
                st.write(resource["purpose"])
                st.link_button(
                    "Open official resource",
                    resource["url"]
                )

    st.warning(
        "Hosting platforms have different supported frameworks, "
        "free-tier limits and deployment requirements. Check their "
        "official documentation before deploying."
    )


# ============================================================
# SEMINAR ASSISTANT
# ============================================================

elif selected_page == "Seminar Assistant":

    st.title("🎤 Seminar Topic Assistant")

    st.write(
        "Prepare a seminar outline using your academic materials."
    )

    seminar_topic = st.text_input(
        "Enter seminar topic",
        placeholder="Example: Applications of Artificial Intelligence"
    )

    seminar_subject = st.text_input(
        "Subject name",
        placeholder="Example: Computer Science"
    )

    presentation_duration = st.selectbox(
        "Presentation duration",
        ["5 minutes", "10 minutes", "15 minutes", "20 minutes"]
    )

    if st.button("Prepare Seminar Outline"):

        if not seminar_topic.strip():
            st.warning("Enter a seminar topic.")
        else:
            seminar_text = f"""
SEMINAR TITLE: {seminar_topic}

SUBJECT: {seminar_subject or "Not specified"}

PRESENTATION DURATION: {presentation_duration}

1. INTRODUCTION
Introduce the topic and explain why it is important.

2. DEFINITION
Write a clear definition using your study materials.

3. BACKGROUND
Explain the basic concepts and relevant history.

4. MAIN CONTENT
Divide the topic into three to five important headings.
Explain each heading using simple language.

5. DIAGRAM OR FLOWCHART
Add a suitable labelled diagram or process flowchart.

6. EXAMPLE
Give a real-world or textbook example.

7. APPLICATIONS
Explain where the topic is used.

8. ADVANTAGES AND LIMITATIONS
Mention important benefits and limitations.

9. CONCLUSION
Summarise the main ideas.

10. REFERENCES
List the textbooks, papers and other sources actually consulted.

POSSIBLE QUESTIONS FROM THE AUDIENCE

1. What is the main purpose of this topic?
2. What are its important features?
3. Where is it used?
4. What are its advantages?
5. What are its limitations?
"""

            st.session_state.seminar_content = seminar_text

    if st.session_state.seminar_content:
        st.text_area(
            "Editable seminar outline",
            value=st.session_state.seminar_content,
            height=450
        )

        st.download_button(
            "Download Seminar Outline",
            data=st.session_state.seminar_content,
            file_name="seminar_outline.txt",
            mime="text/plain"
        )


# ============================================================
# ABOUT PROJECT
# ============================================================

elif selected_page == "About Project":

    st.title("ℹ️ About SCSVMV AcademicQuery AI")

    st.write("""
    AcademicQuery AI is a student-focused academic assistance
    application designed to support university examination preparation.

    Its main objectives are:

    - Organize academic study materials.
    - Search uploaded documents.
    - Analyse previous-year question papers.
    - Identify frequently repeated questions.
    - Prepare structured examination answer drafts.
    - Support quick revision.
    - Create editable student notes and PDF documents.
    - Calculate marks and estimate GPA using entered rules.
    - Guide beginners through project development.
    - Provide project execution and deployment resources.
    - Support seminar preparation.

    This application uses Python-based document processing,
    keyword search, text comparison and predefined templates.

    It does not connect to paid AI APIs or external AI compilers.
    It does not independently understand every subject or guarantee
    which questions will appear in future examinations.
    """)

    st.subheader("Technology used")

    st.markdown("""
    - Python
    - Streamlit
    - PyMuPDF
    - python-docx
    - python-pptx
    - openpyxl
    - ReportLab
    - Regular expressions
    - Python text-processing algorithms
    """)

    st.subheader("Important limitations")

    st.write("""
    Uploaded documents are processed during the current browser session.
    This initial version does not provide permanent shared storage,
    student accounts, secure administrator authentication, or
    cross-session document indexing.

    Scanned PDFs may require OCR, which is not included in this version.
    """)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "SCSVMV AcademicQuery AI | Free Python-based academic assistance | "
    "No paid AI API required"
)
