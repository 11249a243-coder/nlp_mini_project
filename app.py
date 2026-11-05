
import os
import io
import csv
import hashlib
from pathlib import Path

import numpy as np
import streamlit as st
import pymupdf
from docx import Document
from pptx import Presentation
from openpyxl import load_workbook
from sentence_transformers import SentenceTransformer
from openai import OpenAI

st.set_page_config(
    page_title="SCSVMV AcademicQuery AI",
    page_icon="🎓",
    layout="wide",
)

# --------------------------- STYLING ---------------------------

st.markdown("""
<style>
.stApp { background: #f4f7fc; }

[data-testid="stSidebar"] {
    background: #102b50;
}

[data-testid="stSidebar"] * {
    color: white;
}

.main-title {
    color: #102b50;
    font-size: 34px;
    font-weight: 800;
}

.hero {
    background: linear-gradient(120deg,#102b50,#24558a);
    padding: 28px;
    border-radius: 18px;
    color: white;
    margin: 16px 0 24px;
}

.hero h1 { color: white; }
.hero p { color: #e5edf8; font-size: 16px; }

[data-testid="stChatMessage"] {
    background: white;
    border: 1px solid #e5eaf2;
    border-radius: 14px;
}

.stButton button {
    background: #173f70;
    color: white;
    border-radius: 9px;
    border: 0;
}

.stButton button:hover {
    background: #d5a943;
    color: #102b50;
}
</style>
""", unsafe_allow_html=True)


# --------------------------- AI MODEL ---------------------------

@st.cache_resource
def load_embedding_model():
    return SentenceTransformer(
        "sentence-transformers/all-MiniLM-L6-v2",
        device="cpu"
    )


def get_openai_client():

    api_key = os.getenv("OPENAI_API_KEY", "")

    if not api_key and hasattr(st, "secrets"):
        try:
            api_key = st.secrets.get("OPENAI_API_KEY", "")
        except Exception:
            pass

    if not api_key:
        return None

    return OpenAI(api_key=api_key)


# --------------------------- DOCUMENT TYPES ---------------------------

SUPPORTED_TYPES = [
    "pdf",
    "docx",
    "txt",
    "md",
    "csv",
    "xlsx",
    "pptx"
]


# --------------------------- PDF EXTRACTION ---------------------------

def extract_pdf(data):

    parts = []

    with pymupdf.open(
        stream=data,
        filetype="pdf"
    ) as pdf:

        for page_no, page in enumerate(pdf, start=1):

            text = page.get_text("text").strip()

            if text:
                parts.append((page_no, text))

    return parts


# --------------------------- WORD EXTRACTION ---------------------------

def extract_docx(data):

    doc = Document(io.BytesIO(data))

    pieces = [
        p.text.strip()
        for p in doc.paragraphs
        if p.text.strip()
    ]

    for table in doc.tables:

        for row in table.rows:

            values = [
                cell.text.strip()
                for cell in row.cells
            ]

            if any(values):
                pieces.append(" | ".join(values))

    return [(None, "\n".join(pieces))] if pieces else []


# --------------------------- POWERPOINT EXTRACTION ---------------------------

def extract_pptx(data):

    prs = Presentation(io.BytesIO(data))

    pieces = []

    for slide_no, slide in enumerate(
        prs.slides,
        start=1
    ):

        slide_text = []

        for shape in slide.shapes:

            if (
                getattr(shape, "has_text_frame", False)
                and shape.text.strip()
            ):
                slide_text.append(shape.text.strip())

            if getattr(shape, "has_table", False):

                for row in shape.table.rows:

                    slide_text.append(
                        " | ".join(
                            cell.text.strip()
                            for cell in row.cells
                        )
                    )

        if slide_text:
            pieces.append(
                (slide_no, "\n".join(slide_text))
            )

    return pieces


# --------------------------- EXCEL EXTRACTION ---------------------------

def extract_xlsx(data):

    workbook = load_workbook(
        io.BytesIO(data),
        read_only=True,
        data_only=True
    )

    pieces = []

    for sheet in workbook.worksheets:

        rows = []

        for row in sheet.iter_rows(values_only=True):

            values = [
                str(value)
                for value in row
                if value is not None
            ]

            if values:
                rows.append(" | ".join(values))

        if rows:

            pieces.append((
                None,
                f"Sheet: {sheet.title}\n"
                + "\n".join(rows)
            ))

    workbook.close()

    return pieces


# --------------------------- CSV EXTRACTION ---------------------------

def extract_csv(data):

    decoded = data.decode(
        "utf-8-sig",
        errors="replace"
    )

    rows = list(
        csv.reader(io.StringIO(decoded))
    )

    text = "\n".join(
        " | ".join(row)
        for row in rows
        if any(cell.strip() for cell in row)
    )

    return [(None, text)] if text.strip() else []


# --------------------------- FILE EXTRACTION ---------------------------

def extract_file(uploaded_file):

    data = uploaded_file.getvalue()

    ext = Path(
        uploaded_file.name
    ).suffix.lower().lstrip(".")

    if ext == "pdf":
        return extract_pdf(data)

    if ext == "docx":
        return extract_docx(data)

    if ext == "pptx":
        return extract_pptx(data)

    if ext == "xlsx":
        return extract_xlsx(data)

    if ext == "csv":
        return extract_csv(data)

    if ext in ("txt", "md"):

        text = data.decode(
            "utf-8-sig",
            errors="replace"
        )

        return [(None, text)] if text.strip() else []

    raise ValueError(
        f"Unsupported file type: .{ext}"
    )


# --------------------------- TEXT CHUNKING ---------------------------

def split_text(
    text,
    chunk_words=180,
    overlap=35
):

    words = text.split()

    if not words:
        return []

    chunks = []

    step = max(
        1,
        chunk_words - overlap
    )

    for start in range(
        0,
        len(words),
        step
    ):

        chunk = " ".join(
            words[start:start + chunk_words]
        ).strip()

        if chunk:
            chunks.append(chunk)

    return chunks


# --------------------------- SESSION INITIALIZATION ---------------------------

def initialize_session():

    defaults = {
        "records": [],
        "processed_files": set(),
        "chat_history": [],
        "uploaded_names": set(),
    }

    for key, value in defaults.items():

        if key not in st.session_state:
            st.session_state[key] = value


# --------------------------- PROCESS DOCUMENTS ---------------------------

def process_documents(files):

    model = load_embedding_model()

    new_records = []

    processed_count = 0
    skipped_count = 0

    progress = st.progress(
        0,
        text="Starting document processing..."
    )

    for index, uploaded_file in enumerate(files):

        raw = uploaded_file.getvalue()

        file_hash = hashlib.sha256(
            raw
        ).hexdigest()

        if file_hash in st.session_state.processed_files:

            skipped_count += 1

            progress.progress(
                (index + 1) / len(files),
                text=f"Already processed: {uploaded_file.name}"
            )

            continue

        try:

            pages = extract_file(uploaded_file)

            file_chunks = []

            for page_no, page_text in pages:

                for chunk in split_text(page_text):

                    file_chunks.append({
                        "file": uploaded_file.name,
                        "page": page_no,
                        "text": chunk,
                    })

            if not file_chunks:

                raise ValueError(
                    "No readable text found. "
                    "If this is a scanned PDF, OCR is required."
                )

            new_records.extend(file_chunks)

            st.session_state.processed_files.add(
                file_hash
            )

            st.session_state.uploaded_names.add(
                uploaded_file.name
            )

            processed_count += 1

        except Exception as exc:

            st.error(
                f"Could not process {uploaded_file.name}: {exc}"
            )

        progress.progress(
            (index + 1) / len(files),
            text=f"Processed {index + 1} of {len(files)} files"
        )

    if new_records:

        texts = [
            record["text"]
            for record in new_records
        ]

        with st.spinner(
            "Creating searchable document representations..."
        ):

            vectors = model.encode(
                texts,
                convert_to_numpy=True,
                normalize_embeddings=True,
                show_progress_bar=False,
                batch_size=32,
            )

        for record, vector in zip(
            new_records,
            vectors
        ):

            record["embedding"] = vector

        st.session_state.records.extend(
            new_records
        )

    progress.empty()

    return processed_count, skipped_count


# --------------------------- SEMANTIC SEARCH ---------------------------

def retrieve_context(
    question,
    top_k=5
):

    records = st.session_state.records

    if not records:
        return []

    model = load_embedding_model()

    query_vector = model.encode(
        [question],
        convert_to_numpy=True,
        normalize_embeddings=True
    )[0]

    scored = []

    for record in records:

        score = float(
            np.dot(
                query_vector,
                record["embedding"]
            )
        )

        scored.append(
            (score, record)
        )

    scored.sort(
        key=lambda item: item[0],
        reverse=True
    )

    return scored[:top_k]


# --------------------------- AI CHATBOT ---------------------------

def generate_answer(question):

    client = get_openai_client()

    if client is None:

        raise RuntimeError(
            "AI API key is missing. Add OPENAI_API_KEY "
            "in Render Environment settings and redeploy."
        )

    matches = retrieve_context(
        question,
        top_k=5
    )

    context_blocks = []

    for score, record in matches:

        if score < 0.18:
            continue

        location = (
            f"page/slide {record['page']}"
            if record["page"]
            else "document"
        )

        context_blocks.append(
            f"[Document: {record['file']} | {location}]\n"
            f"{record['text']}"
        )

    context = "\n\n---\n\n".join(
        context_blocks
    ) if context_blocks else (
        "No strongly relevant passage was found "
        "in the uploaded documents."
    )

    system_prompt = """
You are AcademicQuery AI, a friendly and careful
university academic assistant.

Chat naturally and clearly, like a helpful tutor.

Answer the student's question directly.
Explain concepts step by step when useful.
Use previous conversation messages to understand
follow-up questions.

Use supplied document excerpts when relevant.

Do not invent university rules, dates, fees,
attendance requirements, exam schedules or
official facts.

If uploaded documents do not contain an answer,
say so plainly.

You may still give general educational explanations,
but clearly label them as general information rather
than official university policy.

Do not include website links, URLs, markdown links,
or suggestions to visit external websites.

You may mention uploaded document filenames and
page or slide numbers as plain-text references.

Never claim a document says something unless
the supplied excerpt supports it.
"""

    messages = [
        {
            "role": "system",
            "content": system_prompt
        }
    ]

    for item in st.session_state.chat_history[-10:]:

        if item["role"] in (
            "user",
            "assistant"
        ):

            messages.append({
                "role": item["role"],
                "content": item["content"]
            })

    messages.append({
        "role": "user",
        "content": (
            f"Student question:\n{question}\n\n"
            f"Relevant uploaded-document excerpts:\n{context}"
        )
    })

    response = client.chat.completions.create(

        model=os.getenv(
            "OPENAI_MODEL",
            "gpt-4o-mini"
        ),

        messages=messages,

        temperature=0.3
    )

    answer = response.choices[0].message.content

    source_records = [
        record
        for score, record in matches
        if score >= 0.18
    ]

    return (
        answer or "I could not generate a response.",
        source_records
    )


# --------------------------- MAIN APPLICATION ---------------------------

initialize_session()

st.sidebar.markdown(
    "<h2 style='text-align:center;color:white;'>SCSVMV</h2>",
    unsafe_allow_html=True
)

st.sidebar.markdown(
    "<p style='text-align:center;color:#d5a943;'>AcademicQuery AI</p>",
    unsafe_allow_html=True
)

st.sidebar.divider()

page = st.sidebar.radio(
    "Navigation",
    [
        "Home",
        "Student Chatbot",
        "Admin Upload",
        "About Project"
    ]
)


# --------------------------- HOME ---------------------------

if page == "Home":

    st.markdown(
        '<div class="main-title">'
        'SCSVMV Academic Portal</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<p style="color:#64748b;">'
        'Your smart university information assistant'
        '</p>',
        unsafe_allow_html=True
    )

    st.markdown("""
    <div class="hero">
        <h1>Welcome to AcademicQuery AI 🎓</h1>
        <p>
        Upload academic documents and ask questions
        in a natural conversation.
        Get clear explanations grounded in your
        uploaded material.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Documents loaded",
        len(st.session_state.uploaded_names)
    )

    col2.metric(
        "Searchable sections",
        len(st.session_state.records)
    )

    col3.metric(
        "Chat messages",
        len(st.session_state.chat_history)
    )

    st.info(
        "Open Student Chatbot to upload documents "
        "and start chatting, or use Admin Upload."
    )


# --------------------------- STUDENT CHATBOT ---------------------------

elif page == "Student Chatbot":

    st.title("🤖 Academic AI Chat")

    st.caption(
        "Upload your documents and ask questions naturally. "
        "No external links are added to answers."
    )

    with st.expander(
        "📎 Attach documents to this chat",
        expanded=not bool(st.session_state.records)
    ):

        chat_files = st.file_uploader(
            "Choose documents",
            type=SUPPORTED_TYPES,
            accept_multiple_files=True,
            key="chat_document_uploader",
            help=(
                "Supported: PDF, DOCX, TXT, MD, "
                "CSV, XLSX and PPTX."
            )
        )

        if st.button(
            "Process attached documents",
            key="process_chat_files"
        ):

            if chat_files:

                done, skipped = process_documents(
                    chat_files
                )

                st.success(
                    f"Finished: {done} new file(s) "
                    f"processed; {skipped} already processed."
                )

            else:

                st.warning(
                    "Choose at least one document first."
                )

    if st.session_state.uploaded_names:

        st.caption(
            "Available in this browser session: "
            + ", ".join(
                sorted(st.session_state.uploaded_names)
            )
        )

    for message in st.session_state.chat_history:

        with st.chat_message(message["role"]):

            st.markdown(
                message["content"]
            )

            if message.get("sources"):

                with st.expander(
                    "Document references"
                ):

                    for source in message["sources"]:

                        loc = (
                            f" — page/slide {source['page']}"
                            if source["page"]
                            else ""
                        )

                        st.write(
                            f"**{source['file']}**{loc}"
                        )

    question = st.chat_input(
        "Ask anything about your documents or an academic topic..."
    )

    if question:

        st.session_state.chat_history.append({
            "role": "user",
            "content": question
        })

        with st.chat_message("user"):

            st.markdown(question)

        with st.chat_message("assistant"):

            try:

                with st.spinner(
                    "Reading your documents and preparing an answer..."
                ):

                    answer, sources = generate_answer(
                        question
                    )

                st.markdown(answer)

                if sources:

                    with st.expander(
                        "Document references"
                    ):

                        unique_sources = {}

                        for source in sources:

                            key = (
                                source["file"],
                                source["page"]
                            )

                            unique_sources[key] = source

                        for filename, page_no in unique_sources:

                            loc = (
                                f" — page/slide {page_no}"
                                if page_no
                                else ""
                            )

                            st.write(
                                f"**{filename}**{loc}"
                            )

                st.session_state.chat_history.append({
                    "role": "assistant",
                    "content": answer,
                    "sources": sources
                })

            except Exception as exc:

                error_message = str(exc)

                st.error(error_message)

                st.session_state.chat_history.append({
                    "role": "assistant",
                    "content": (
                        "Sorry, I could not answer that yet. "
                        + error_message
                    ),
                    "sources": []
                })

    col_a, col_b = st.columns(2)

    with col_a:

        if st.button("Clear chat history"):

            st.session_state.chat_history = []

            st.rerun()

    with col_b:

        if st.button("Clear uploaded documents"):

            st.session_state.records = []

            st.session_state.processed_files = set()

            st.session_state.uploaded_names = set()

            st.session_state.chat_history = []

            st.rerun()


# --------------------------- ADMIN UPLOAD ---------------------------

elif page == "Admin Upload":

    st.title("📚 Admin Document Upload")

    st.write(
        "Upload university documents to make them "
        "searchable in this browser session."
    )

    files = st.file_uploader(
        "Choose documents",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
        key="admin_document_uploader"
    )

    if st.button(
        "Process documents",
        key="admin_process"
    ):

        if files:

            done, skipped = process_documents(
                files
            )

            st.success(
                f"Finished: {done} new file(s) "
                f"processed; {skipped} already processed."
            )

        else:

            st.warning(
                "Please select at least one document."
            )

    st.divider()

    st.subheader("Processed documents")

    if st.session_state.uploaded_names:

        for name in sorted(
            st.session_state.uploaded_names
        ):

            st.write("📄", name)

    else:

        st.info(
            "No documents processed yet."
        )

    if st.button(
        "Clear all documents and chat"
    ):

        st.session_state.records = []

        st.session_state.processed_files = set()

        st.session_state.uploaded_names = set()

        st.session_state.chat_history = []

        st.rerun()


# --------------------------- ABOUT PROJECT ---------------------------

else:

    st.title("About AcademicQuery AI")

    st.write(
        "AcademicQuery AI is a document-assisted "
        "university question-answering prototype "
        "built with Streamlit, semantic retrieval "
        "and an AI language model."
    )

    st.subheader("Supported document types")

    st.write(
        "PDF, DOCX, TXT, MD, CSV, XLSX and PPTX."
    )

    st.warning(
        "Scanned/image-only PDFs need OCR before "
        "their text can be searched. Uploaded "
        "documents and chat history are currently "
        "held in the active Streamlit session. "
        "They are not automatically shared with "
        "other devices or permanently stored."
    )
