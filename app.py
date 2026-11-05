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
