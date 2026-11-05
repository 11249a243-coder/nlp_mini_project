import streamlit as st
import fitz
import re
import io
import csv
from collections import Counter
from datetime import datetime

from docx import Document
from pptx import Presentation
from openpyxl import load_workbook

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer
)
from reportlab.lib.styles import getSampleStyleSheet


# =====================================================
# PAGE CONFIGURATION
# =====================================================

st.set_page_config(
    page_title="SCSVMV AcademicQuery AI",
    page_icon="🎓",
    layout="wide"
)


# =====================================================
# CUSTOM DESIGN
# =====================================================

st.markdown("""
<style>

.stApp {
    background-color: #f5f7fb;
}

.main-title {
    font-size: 35px;
    font-weight: bold;
    color: #183153;
    text-align: center;
}

.subtitle {
    font-size: 16px;
    color: #52657a;
    text-align: center;
}

.section-title {
    color: #183153;
    font-size: 24px;
    font-weight: bold;
}

div.stButton > button {
    border-radius: 10px;
    font-weight: bold;
}

</style>
""", unsafe_allow_html=True)


# =====================================================
# SESSION STORAGE
# =====================================================

if "documents" not in st.session_state:
    st.session_state.documents = {}

if "notes" not in st.session_state:
    st.session_state.notes = ""

if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""

if "seminar_content" not in st.session_state:
    st.session_state.seminar_content = ""


# =====================================================
# DOCUMENT EXTRACTION
# =====================================================

def extract_text(uploaded_file):

    filename = uploaded_file.name.lower()

    try:

        if filename.endswith(".pdf"):

            pdf = fitz.open(
                stream=uploaded_file.read(),
                filetype="pdf"
            )

            pages = []

            for page in pdf:
                pages.append(page.get_text())

            return "\n".join(pages)

        elif filename.endswith(".docx"):

            document = Document(uploaded_file)

            return "\n".join(
                paragraph.text
                for paragraph in document.paragraphs
            )

        elif filename.endswith(".pptx"):

            presentation = Presentation(uploaded_file)

            content = []

            for slide in presentation.slides:

                for shape in slide.shapes:

                    if hasattr(shape, "text"):
                        content.append(shape.text)

            return "\n".join(content)

        elif filename.endswith(".xlsx"):

            workbook = load_workbook(
                uploaded_file,
                data_only=True
            )

            content = []

            for sheet in workbook.worksheets:

                for row in sheet.iter_rows(values_only=True):

                    values = [
                        str(value)
                        for value in row
                        if value is not None
                    ]

                    if values:
                        content.append(" ".join(values))

            return "\n".join(content)

        elif filename.endswith(".csv"):

            text = uploaded_file.read().decode(
                "utf-8",
                errors="ignore"
            )

            return text

        elif filename.endswith(".txt"):

            return uploaded_file.read().decode(
                "utf-8",
                errors="ignore"
            )

        else:

            return "Unsupported file format."

    except Exception as error:

        return "Could not read file: " + str(error)


# =====================================================
# TEXT PROCESSING
# =====================================================

def tokenize(text):

    return re.findall(
        r"[a-zA-Z0-9]+",
        text.lower()
    )


def split_sentences(text):

    sentences = re.split(
        r"(?<=[.!?])\s+|\n+",
        text
    )

    return [
        sentence.strip()
        for sentence in sentences
        if len(sentence.strip()) > 15
    ]


def find_relevant_sentences(text, topic, limit=12):

    topic_words = set(tokenize(topic))

    if not topic_words:
        return []

    sentences = split_sentences(text)

    results = []

    for sentence in sentences:

        sentence_words = set(tokenize(sentence))

        score = len(
            topic_words.intersection(sentence_words)
        )

        if score > 0:

            results.append(
                (score, sentence)
            )

    results.sort(
        key=lambda item: item[0],
        reverse=True
    )

    return [
        sentence
        for score, sentence in results[:limit]
    ]


# =====================================================
# DOCUMENT SEARCH
# =====================================================

def search_documents(topic):

    all_results = []

    for filename, content in st.session_state.documents.items():

        results = find_relevant_sentences(
            content,
            topic,
            limit=10
        )

        all_results.extend(results)

    return all_results


# =====================================================
# ANSWER GENERATOR
# =====================================================

def generate_answer(topic, marks=5):

    results = search_documents(topic)

    if not results:

        return (
            "No matching information was found in the "
            "uploaded documents.\n\n"
            "Try using different keywords or upload "
            "a relevant study material or question paper."
        )

    answer = []

    answer.append("TOPIC: " + topic)

    answer.append("")
    answer.append("INTRODUCTION")

    answer.append(
        "The following points are collected from "
        "your uploaded study materials."
    )

    answer.append("")
    answer.append("MAIN ANSWER")

    for number, point in enumerate(results, start=1):

        answer.append(
            str(number) + ". " + point
        )

    answer.append("")
    answer.append("CONCLUSION")

    answer.append(
        "Review the above points and write the answer "
        "in a clear and organized manner."
    )

    answer.append("")
    answer.append("EXAM PRESENTATION TIPS")

    if marks <= 2:

        answer.append("- Write a short definition.")
        answer.append("- Include one or two important points.")

    elif marks <= 5:

        answer.append("- Start with a suitable definition.")
        answer.append("- Explain the main points.")
        answer.append("- Add an example if relevant.")
        answer.append("- End with a short conclusion.")

    elif marks <= 10:

        answer.append("- Write an introduction.")
        answer.append("- Explain each heading separately.")
        answer.append("- Include examples and diagrams if suitable.")
        answer.append("- Write a clear conclusion.")

    else:

        answer.append("- Start with an introduction.")
        answer.append("- Organize the answer using headings.")
        answer.append("- Explain each concept in detail.")
        answer.append("- Add examples, diagrams and flowcharts.")
        answer.append("- Finish with a conclusion.")

    answer.append("")
    answer.append(
        "Note: This is a keyword-based study aid. "
        "Check your textbook and class notes before "
        "using the answer in an examination."
    )

    return "\n".join(answer)


# =====================================================
# PDF CREATOR
# =====================================================

def create_pdf(title, content):

    buffer = io.BytesIO()

    pdf = SimpleDocTemplate(
        buffer,
        pagesize=A4
    )

    styles = getSampleStyleSheet()

    story = []

    story.append(
        Paragraph(title, styles["Title"])
    )

    story.append(Spacer(1, 20))

    for line in content.split("\n"):

        safe_line = (
            line.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
        )

        if safe_line.strip():

            story.append(
                Paragraph(
                    safe_line,
                    styles["BodyText"]
                )
            )

        story.append(Spacer(1, 7))

    pdf.build(story)

    buffer.seek(0)

    return buffer


# =====================================================
# HEADER
# =====================================================

st.markdown(
    '<div class="main-title">🎓 SCSVMV AcademicQuery AI</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Academic Learning | Exam Preparation | Smart Study Assistant'
    '</div>',
    unsafe_allow_html=True
)

st.divider()


# =====================================================
# SIDEBAR NAVIGATION
# =====================================================

menu = st.sidebar.radio(

    "📚 MAIN MENU",

    [
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
)


# =====================================================
# HOME PAGE
# =====================================================

if menu == "Home":

    st.header("Welcome to AcademicQuery AI")

    st.write(
        "Your academic assistant for learning, "
        "exam preparation and project guidance."
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.info("📖 Upload study materials")

    with col2:

        st.success("✍️ Prepare exam answers")

    with col3:

        st.warning("📄 Create your own notes")

    st.subheader("Available Features")

    st.markdown("""
    - Student academic chatbot
    - Study material upload
    - Previous question paper analysis
    - Exam answer preparation
    - Quick topic learning
    - Editable notes
    - PDF creation
    - Internal marks calculator
    - Project development guidance
    - Seminar preparation
    """)

    st.caption(
        "This application uses local Python-based searching "
        "and does not require a paid AI API."
    )


# =====================================================
# STUDENT UPLOAD
# =====================================================

elif menu == "Student Upload":

    st.header("📤 Student Study Material Upload")

    uploaded_files = st.file_uploader(

        "Upload your study materials",

        type=[
            "pdf",
            "docx",
            "pptx",
            "xlsx",
            "csv",
            "txt"
        ],

        accept_multiple_files=True
    )

    if st.button("Process Student Files"):

        if uploaded_files:

            for file in uploaded_files:

                text = extract_text(file)

                if not text.startswith("Could not read"):

                    st.session_state.documents[
                        file.name
                    ] = text

            st.success("Files processed.")

        else:

            st.warning("Please select a file first.")

    st.subheader("Uploaded Documents")

    if st.session_state.documents:

        for filename in st.session_state.documents:

            st.write("📄 " + filename)

    else:

        st.info("No documents uploaded yet.")


# =====================================================
# ADMIN UPLOAD
# =====================================================

elif menu == "Admin Upload":

    st.header("🔐 Admin Study Material Upload")

    st.warning(
        "This is a demonstration upload page. "
        "It does not currently have administrator login "
        "or access protection."
    )

    admin_files = st.file_uploader(

        "Upload academic materials",

        type=[
            "pdf",
            "docx",
            "pptx",
            "xlsx",
            "csv",
            "txt"
        ],

        accept_multiple_files=True,

        key="admin_upload"
    )

    if st.button("Add Admin Materials"):

        if admin_files:

            for file in admin_files:

                text = extract_text(file)

                if not text.startswith("Could not read"):

                    st.session_state.documents[
                        file.name
                    ] = text

            st.success("Admin materials added.")

        else:

            st.warning("Select files to upload.")


# =====================================================
# STUDENT CHATBOT
# =====================================================

elif menu == "Student Chatbot":

    st.header("💬 Student Academic Chatbot")

    st.write(
        "Ask a question based on your uploaded materials."
    )

    question = st.text_input(
        "Enter your academic question"
    )

    marks = st.selectbox(
        "Choose answer marks",
        [2, 5, 10, 13, 15, 16, 20]
    )

    if st.button("Generate Answer"):

        if question.strip():

            answer = generate_answer(
                question,
                marks
            )

            st.session_state.last_answer = answer

            st.text_area(
                "Generated Answer",
                value=answer,
                height=400
            )

        else:

            st.warning("Please enter a question.")


# =====================================================
# QUICK LEARNING
# =====================================================

elif menu == "Quick Learning":

    st.header("⚡ Quick Learning")

    topic = st.text_input(
        "Enter the topic you want to learn"
    )

    duration = st.selectbox(
        "Select learning time",
        ["5 minutes", "15 minutes", "20 minutes"]
    )

    if st.button("Start Quick Learning"):

        if topic.strip():

            results = search_documents(topic)

            if results:

                st.subheader("Important Learning Points")

                for number, point in enumerate(
                    results,
                    start=1
                ):

                    st.markdown(
                        f"**{number}.** {point}"
                    )

                st.subheader("Quick Revision Plan")

                if duration == "5 minutes":

                    st.write(
                        "Read the important points "
                        "and revise the definitions."
                    )

                elif duration == "15 minutes":

                    st.write(
                        "Read the points, understand "
                        "the examples and revise."
                    )

                else:

                    st.write(
                        "Read the points, write a short "
                        "answer and revise the key terms."
                    )

            else:

                st.warning(
                    "No matching topic found. "
                    "Upload relevant study materials."
                )

        else:

            st.warning("Enter a topic.")


# =====================================================
# PREVIOUS QUESTION PAPERS
# =====================================================

elif menu == "Previous Question Papers":

    st.header("📚 Previous Question Paper Analysis")

    st.write(
        "Upload previous question papers to identify "
        "frequently occurring questions and topics."
    )

    question_files = st.file_uploader(

        "Upload question papers",

        type=["pdf", "docx", "txt"],

        accept_multiple_files=True,

        key="question_papers"
    )

    if st.button("Analyze Question Papers"):

        if question_files:

            combined_text = []

            for file in question_files:

                text = extract_text(file)

                if text and not text.startswith("Could not read"):

                    combined_text.append(text)

            if combined_text:

                all_text = "\n".join(combined_text)

                lines = split_sentences(all_text)

                counts = Counter(
                    line.lower()
                    for line in lines
                )

                repeated = [
                    (question, count)
                    for question, count in counts.items()
                    if count > 1
                ]

                repeated.sort(
                    key=lambda item: item[1],
                    reverse=True
                )

                st.subheader("Repeated Text")

                if repeated:

                    for question, count in repeated[:20]:

                        st.write(
                            f"**Found {count} times:** {question}"
                        )

                else:

                    st.info(
                        "No exact repeated sentences found. "
                        "Similar questions may be worded differently."
                    )

        else:

            st.warning("Upload question papers first.")


# =====================================================
# EXAM ANSWER BUILDER
# =====================================================

elif menu == "Exam Answer Builder":

    st.header("✍️ Exam Answer Builder")

    topic = st.text_input(
        "Enter the question or topic",
        key="builder_topic"
    )

    marks = st.selectbox(
        "Select marks",
        [2, 5, 10, 13, 15, 16, 20],
        key="builder_marks"
    )

    if st.button("Build Exam Answer"):

        if topic.strip():

            st.session_state.last_answer = generate_answer(
                topic,
                marks
            )

        else:

            st.warning("Enter a topic.")

    st.text_area(
        "Edit your answer",
        key="builder_edit",
        height=400
    )

    if st.button("Load Generated Answer"):

        st.session_state.builder_edit = (
            st.session_state.last_answer
        )

        st.rerun()


# =====================================================
# MY NOTES AND PDF CREATOR
# =====================================================

elif menu == "My Notes & PDF Creator":

    st.header("📝 My Notes & PDF Creator")

    title = st.text_input(
        "Document title",
        value="My Academic Notes"
    )

    notes = st.text_area(
        "Write or edit your notes",
        value=st.session_state.notes,
        height=400
    )

    st.session_state.notes = notes

    if st.button("Save Notes"):

        st.session_state.notes = notes

        st.success("Notes saved for this session.")

    st.download_button(

        "Download Notes as TXT",

        data=notes,

        file_name="academic_notes.txt",

        mime="text/plain"
    )

    if st.button("Create PDF"):

        if notes.strip():

            pdf_file = create_pdf(
                title,
                notes
            )

            st.download_button(

                "Download PDF",

                data=pdf_file,

                file_name="academic_notes.pdf",

                mime="application/pdf"
            )

        else:

            st.warning("Write some notes first.")


# =====================================================
# ACADEMIC CALCULATOR
# =====================================================

elif menu == "Academic Calculator":

    st.header("🧮 Academic Calculator")

    st.subheader("Internal Marks Percentage")

    obtained = st.number_input(
        "Marks obtained",
        min_value=0.0,
        value=0.0
    )

    maximum = st.number_input(
        "Maximum marks",
        min_value=1.0,
        value=100.0
    )

    if st.button("Calculate Percentage"):

        if obtained <= maximum:

            percentage = (
                obtained / maximum
            ) * 100

            st.success(
                f"Percentage: {percentage:.2f}%"
            )

        else:

            st.error(
                "Obtained marks cannot exceed maximum marks."
            )

    st.divider()

    st.subheader("Weighted GPA Estimate")

    st.caption(
        "Enter the grade points and credits for each subject. "
        "This is only a generic estimate, not an official "
        "university calculation."
    )

    number_subjects = st.number_input(
        "Number of subjects",
        min_value=1,
        max_value=15,
        value=5
    )

    total_points = 0.0
    total_credits = 0.0

    for i in range(int(number_subjects)):

        col1, col2 = st.columns(2)

        with col1:

            grade = st.number_input(
                f"Subject {i + 1} grade point",
                min_value=0.0,
                max_value=10.0,
                value=0.0,
                key=f"grade_{i}"
            )

        with col2:

            credit = st.number_input(
                f"Subject {i + 1} credits",
                min_value=0.0,
                value=3.0,
                key=f"credit_{i}"
            )

        total_points += grade * credit

        total_credits += credit

    if st.button("Calculate GPA"):

        if total_credits > 0:

            gpa = total_points / total_credits

            st.success(
                f"Estimated GPA: {gpa:.2f}"
            )

        else:

            st.warning("Enter valid credits.")


# =====================================================
# PROJECT DEVELOPMENT
# =====================================================

elif menu == "Project Development":

    st.header("💻 Project Development Assistant")

    st.write(
        "Beginner-friendly guidance for developing "
        "academic projects."
    )

    project_type = st.selectbox(

        "Choose your project area",

        [
            "Python Project",
            "Machine Learning Project",
            "Web Development Project",
            "Data Science Project",
            "Artificial Intelligence Project",
            "Database Project"
        ]
    )

    st.subheader("Development Steps")

    steps = [
        "Choose a simple project idea.",
        "Understand the project requirements.",
        "Install Python and a code editor.",
        "Create a project folder.",
        "Write the program step by step.",
        "Install the required Python packages.",
        "Run the program on your computer.",
        "Test the program and correct errors.",
        "Upload your project to GitHub.",
        "Deploy the project if required."
    ]

    for number, step in enumerate(steps, start=1):

        st.write(
            f"**Step {number}:** {step}"
        )

    st.subheader("Useful Development Platforms")

    resources = {
        "Visual Studio Code": "https://code.visualstudio.com/",
        "PyCharm": "https://www.jetbrains.com/pycharm/",
        "Jupyter": "https://jupyter.org/",
        "Google Colab": "https://colab.research.google.com/",
        "Replit": "https://replit.com/",
        "GitHub": "https://github.com/",
        "GitHub Codespaces": "https://github.com/features/codespaces",
        "Streamlit Community Cloud": "https://streamlit.io/cloud",
        "Render": "https://render.com/",
        "Vercel": "https://vercel.com/",
        "Netlify": "https://www.netlify.com/",
        "PythonAnywhere": "https://www.pythonanywhere.com/",
        "GitHub Pages": "https://pages.github.com/"
    }

    for name, url in resources.items():

        st.markdown(
            f"- [{name}]({url})"
        )

    st.info(
        "Some platforms have free usage limits. "
        "Check their current terms before using them."
    )


# =====================================================
# SEMINAR ASSISTANT
# =====================================================

elif menu == "Seminar Assistant":

    st.header("🎤 Seminar Topic Assistant")

    topic = st.text_input(
        "Enter your seminar topic"
    )

    if st.button("Prepare Seminar"):

        if topic.strip():

            content = f"""
SEMINAR TOPIC: {topic}

1. INTRODUCTION

Introduce the topic and explain its importance.

2. DEFINITION

Write a clear definition of the topic.

3. BACKGROUND

Explain the basic background and development.

4. MAIN CONCEPTS

Explain the important concepts using headings.

5. WORKING PRINCIPLE

Describe how the topic works.

6. APPLICATIONS

Give real-world applications and examples.

7. ADVANTAGES

Explain the benefits.

8. LIMITATIONS

Explain the challenges and disadvantages.

9. FUTURE SCOPE

Describe possible future developments.

10. CONCLUSION

Summarize the seminar topic.

Note: Add verified subject-specific information,
examples and references before presenting.
"""

            st.session_state.seminar_content = content

        else:

            st.warning("Enter a seminar topic.")

    seminar_text = st.text_area(
        "Edit your seminar content",
        value=st.session_state.seminar_content,
        height=400
    )

    st.session_state.seminar_content = seminar_text

    st.download_button(

        "Download Seminar Notes",

        data=seminar_text,

        file_name="seminar_notes.txt",

        mime="text/plain"
    )


# =====================================================
# ABOUT PROJECT
# =====================================================

elif menu == "About Project":

    st.header("ℹ️ About SCSVMV AcademicQuery AI")

    st.write("""
    SCSVMV AcademicQuery AI is an academic assistance
    application designed to support students during
    examination preparation.

    Main objectives:

    - Help students search uploaded study materials.
    - Identify important topics.
    - Support exam answer preparation.
    - Analyze previous question papers.
    - Create editable academic notes.
    - Generate PDF documents.
    - Provide basic academic calculations.
    - Guide students in project development.
    - Support seminar preparation.

    TECHNICAL DETAILS

    - Python
    - Streamlit
    - PyMuPDF
    - python-docx
    - python-pptx
    - openpyxl
    - ReportLab

    This version uses keyword-based document searching.
    It does not use a paid AI API or guarantee exam results.
    """)

    st.caption(
        "SCSVMV AcademicQuery AI | Student Academic Project"
    )

