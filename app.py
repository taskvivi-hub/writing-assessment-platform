import os
import json
import uuid
import base64
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO

import streamlit as st
from openai import OpenAI
from openpyxl import Workbook

APP_NAME = "Writing Assessment"
TASKS_FILE = Path("tasks.json")
SUBMISSIONS_FILE = Path("submissions.json")

RUBRIC = {
    "Content & Task Fulfillment": {
        4: ("Exceeds Expectations", "Fully addresses all task requirements. Content is relevant, sufficiently developed, and supported with specific information, examples, or evidence appropriate to the academic or professional purpose."),
        3: ("Meets Expectations", "Addresses the main task requirements. Content is relevant and adequately developed, with enough supporting information to complete the task successfully."),
        2: ("Approaching Expectations", "Addresses only part of the task or develops ideas unevenly. Support may be limited, repetitive, or insufficient for the intended purpose."),
        1: ("Needs Development", "Does not adequately address the task. Content is minimal, off-topic, or too incomplete to fulfill the required academic or professional purpose.")
    },
    "Organization & Coherence": {
        4: ("Exceeds Expectations", "Information is well organized at both paragraph and whole-text levels. Ideas progress logically, transitions are effective, and the reader can follow the argument or message without difficulty."),
        3: ("Meets Expectations", "Organization is generally clear. Paragraphing and sequencing are appropriate, and most ideas are connected logically, with only minor lapses in coherence."),
        2: ("Approaching Expectations", "Some organization is evident, but paragraphing, sequencing, or transitions are inconsistent and occasionally make the text difficult to follow."),
        1: ("Needs Development", "Organization is weak or unclear. Ideas are fragmented, poorly sequenced, or insufficiently connected, making the text difficult to follow.")
    },
    "Language Use": {
        4: ("Exceeds Expectations", "Uses an effective range of vocabulary and sentence structures, including appropriate academic or professional language. Word choice is generally precise, and errors are minor and do not affect meaning."),
        3: ("Meets Expectations", "Uses sufficient vocabulary and sentence structures to complete the task. Academic or professional language is generally appropriate, and errors rarely interfere with meaning."),
        2: ("Approaching Expectations", "Uses a limited range of vocabulary or structures. Repetition, imprecise wording, or frequent errors sometimes reduce clarity or accuracy."),
        1: ("Needs Development", "Language resources are too limited for the task. Frequent or serious errors in wording or sentence construction make important parts of the text difficult to understand.")
    },
    "Genre & Professional Appropriacy": {
        4: ("Exceeds Expectations", "Consistently follows the expected purpose, organization, format, tone, and conventions of the assigned genre. The writing is well suited to its intended audience and professional or academic context."),
        3: ("Meets Expectations", "Generally follows the expected purpose, organization, format, and tone of the genre. Minor inconsistencies do not interfere with the intended communication."),
        2: ("Approaching Expectations", "Shows partial control of the genre. Format, tone, organization, or audience awareness is inconsistent and sometimes weakens the effectiveness of the text."),
        1: ("Needs Development", "Shows limited awareness of the assigned genre. Format, tone, organization, or audience expectations are frequently inappropriate for the task.")
    }
}


def read_json(path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return default
    return default


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_tasks():
    return read_json(TASKS_FILE, {})


def save_tasks(tasks):
    write_json(TASKS_FILE, tasks)


def load_submissions():
    return read_json(SUBMISSIONS_FILE, [])


def save_submissions(items):
    write_json(SUBMISSIONS_FILE, items)


def secret_or_env(name, default=""):
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default
    return value or os.getenv(name, default)


def get_base_url():
    return secret_or_env("APP_BASE_URL", "").rstrip("/")


def student_link(task_id):
    base = get_base_url()
    return f"{base}/?task={task_id}" if base else f"?task={task_id}"


def image_to_data_url(uploaded_file):
    mime = uploaded_file.type or "image/jpeg"
    b64 = base64.b64encode(uploaded_file.getvalue()).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def build_prompt(task):
    requirements = task.get("requirements", "").strip()
    return f'''You are an English writing assessor. Read the student's uploaded composition carefully.

TEACHER SETTINGS
Task Title: {task["title"]}
Genre / Writing Type: {task["genre"]}
Task Requirements:
{requirements}

ASSESSMENT RULES
Evaluate exactly these four dimensions. Give an INTEGER score from 1 to 4 for each:
1. Content & Task Fulfillment
2. Organization & Coherence
3. Language Use
4. Genre & Professional Appropriacy

Important:
- No half points.
- Do not add criteria.
- Do not double-penalize the same issue.
- Judge Content & Task Fulfillment against the stated Task Requirements.
- Evaluate only the writing the student is required to produce.
- Some assignments may already provide fixed genre elements outside the student's response, such as a subject line, greeting, opening, closing, or signature.
- Do NOT penalize a student for omitting any element that is not explicitly required in Task Requirements.
- If the Genre / Writing Type is "Email Body", evaluate only the body paragraphs for appropriate purpose, organization, tone, audience awareness, and professional/academic appropriacy. Do not require a greeting, closing, or signature.
- If the image is not readable enough, do not guess.

Then identify genuine errors ONLY in Grammar, Spelling, and Punctuation.
For each genuine error, provide type, original, correction, and a brief A2-B1 English explanation.
Do not rewrite the full composition. Preserve the student's intended meaning. Do not list stylistic preferences as grammar errors.

Finally, provide 1 to 3 brief suggestions about CONTENT and ORGANIZATION only. Do not provide a model essay.

Return VALID JSON ONLY:
{{
  "image_readable": true,
  "transcription": "faithful transcription",
  "scores": {{
    "Content & Task Fulfillment": 1,
    "Organization & Coherence": 1,
    "Language Use": 1,
    "Genre & Professional Appropriacy": 1
  }},
  "corrections": [
    {{
      "type": "Grammar",
      "original": "student text",
      "correction": "corrected text",
      "explanation": "brief explanation"
    }}
  ],
  "content_organization_suggestions": ["brief suggestion"]
}}'''


def assess(uploaded_file, task):
    api_key = secret_or_env("OPENAI_API_KEY")
    model = secret_or_env("OPENAI_MODEL", "gpt-5.6-luna")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not configured.")

    client = OpenAI(api_key=api_key)
    response = client.responses.create(
        model=model,
        input=[{
            "role": "user",
            "content": [
                {"type": "input_text", "text": build_prompt(task)},
                {"type": "input_image", "image_url": image_to_data_url(uploaded_file)}
            ]
        }]
    )
    raw = response.output_text.strip()
    if raw.startswith("```"):
        raw = raw.strip("`").strip()
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
    result = json.loads(raw)
    for dim in RUBRIC:
        score = int(result["scores"][dim])
        if score not in (1, 2, 3, 4):
            raise ValueError(f"Invalid score for {dim}")
        result["scores"][dim] = score
    return result


def taipei_now():
    return datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y-%m-%d %H:%M:%S")


def save_submission(task_id, task, student_id, student_name, result):
    submissions = load_submissions()
    row = {
        "submission_id": uuid.uuid4().hex,
        "task_id": task_id,
        "class_name": task.get("class_name", ""),
        "task_date": task.get("task_date", ""),
        "task_title": task.get("title", ""),
        "genre": task.get("genre", ""),
        "student_id": student_id.strip(),
        "student_name": student_name.strip(),
        "content_score": result["scores"]["Content & Task Fulfillment"],
        "organization_score": result["scores"]["Organization & Coherence"],
        "language_score": result["scores"]["Language Use"],
        "genre_score": result["scores"]["Genre & Professional Appropriacy"],
        "total": sum(result["scores"].values()),
        "submitted_at": taipei_now()
    }
    submissions.append(row)
    save_submissions(submissions)
    return row


def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Writing Results"
    headers = [
        "Class", "Task Date", "Task Title", "Genre / Writing Type",
        "Student ID", "Student Name",
        "Content & Task Fulfillment", "Organization & Coherence",
        "Language Use", "Genre & Professional Appropriacy",
        "Total /16", "Submission Time"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.get("class_name", ""), r.get("task_date", ""), r.get("task_title", ""), r.get("genre", ""),
            r.get("student_id", ""), r.get("student_name", ""),
            r.get("content_score", ""), r.get("organization_score", ""),
            r.get("language_score", ""), r.get("genre_score", ""),
            r.get("total", ""), r.get("submitted_at", "")
        ])
    widths = [18, 12, 28, 20, 16, 18, 24, 24, 16, 30, 12, 22]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = width
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def teacher_authenticated():
    required = secret_or_env("TEACHER_PASSWORD", "")
    if not required:
        st.warning("Teacher password is not configured yet. Add TEACHER_PASSWORD in Streamlit Secrets before using real student data.")
        return True
    if st.session_state.get("teacher_ok"):
        return True
    st.subheader("Teacher Login")
    pwd = st.text_input("Password", type="password")
    if st.button("Log in"):
        if pwd == required:
            st.session_state["teacher_ok"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


st.set_page_config(page_title=APP_NAME, page_icon="✍️", layout="centered")
st.markdown('''
<style>
.block-container{max-width:900px;padding-top:2rem;padding-bottom:4rem}
.taskbox,.card,.totalbox{border:1px solid rgba(120,120,120,.28);border-radius:14px;padding:1rem 1.1rem;margin:.7rem 0}
.totalbox{text-align:center;border-width:2px}.totalnum{font-size:2.2rem;font-weight:800}
.cardtitle{font-size:1.05rem;font-weight:700}.score{float:right;font-weight:800}.level{font-weight:700;margin-top:.35rem}.meta{opacity:.75}
</style>
''', unsafe_allow_html=True)

tasks = load_tasks()
task_id = st.query_params.get("task")

if task_id and task_id in tasks:
    task = tasks[task_id]
    st.title("Writing Assessment")
    st.write("Follow the steps below to check your writing.")
    st.markdown(
        f'''<div class="taskbox"><b>Class:</b> {task.get("class_name","")}<br>
        <b>Task Date:</b> {task.get("task_date","")}<br>
        <b>Task Title:</b> {task["title"]}<br>
        <b>Genre / Writing Type:</b> {task["genre"]}</div>''',
        unsafe_allow_html=True
    )
    if task.get("requirements"):
        st.subheader("Task Requirements")
        st.write(task["requirements"])

    st.subheader("Step 1. Enter your information")
    student_id = st.text_input("Student ID")
    student_name = st.text_input("Student Name")

    st.subheader("Step 2. Upload your writing")
    st.write("Take a clear photo of your writing and upload it here.")
    uploaded = st.file_uploader("Choose an image", type=["jpg", "jpeg", "png", "webp"])

    st.subheader("Step 3. Submit your writing")
    if st.button("Submit for Assessment", type="primary", use_container_width=True):
        if not student_id.strip() or not student_name.strip():
            st.warning("Please enter your Student ID and Student Name.")
        elif not uploaded:
            st.warning("Please upload an image first.")
        else:
            with st.spinner("Checking your writing..."):
                try:
                    result = assess(uploaded, task)
                    if not result.get("image_readable", True):
                        st.error("The image is not clear enough to read. Please upload a clearer photo.")
                    else:
                        save_submission(task_id, task, student_id, student_name, result)
                        st.session_state["result"] = result
                        st.session_state["show_revision"] = False
                except Exception as e:
                    st.error(f"Assessment could not be completed: {e}")

    result = st.session_state.get("result")
    if result:
        st.divider()
        st.header("Your Writing Score")
        total = sum(result["scores"][d] for d in RUBRIC)
        st.markdown(f'<div class="totalbox"><div class="meta">Total Score</div><div class="totalnum">{total} / 16</div></div>', unsafe_allow_html=True)
        for dim, levels in RUBRIC.items():
            score = result["scores"][dim]
            level, desc = levels[score]
            st.markdown(f'<div class="card"><span class="cardtitle">{dim}</span><span class="score">{score}/4</span><div class="level">{level}</div><div>{desc}</div></div>', unsafe_allow_html=True)
        if st.button("See Revision Suggestions", use_container_width=True):
            st.session_state["show_revision"] = True
        if st.session_state.get("show_revision"):
            st.divider()
            st.header("Revision Suggestions")
            st.subheader("Language Corrections")
            corrections = result.get("corrections", [])
            if not corrections:
                st.success("No clear grammar, spelling, or punctuation errors were found.")
            else:
                for i, c in enumerate(corrections, 1):
                    st.markdown(f"**{i}. {c.get('type','Correction')}**")
                    st.markdown(f"- **Original:** {c.get('original','')}")
                    st.markdown(f"- **Correction:** {c.get('correction','')}")
                    st.markdown(f"- **Why:** {c.get('explanation','')}")
                    st.write("")
            st.subheader("Content & Organization Suggestions")
            suggestions = result.get("content_organization_suggestions", [])
            if suggestions:
                for s in suggestions[:3]:
                    st.markdown(f"- {s}")
            else:
                st.write("No additional suggestions.")
            st.info("Revise the errors in your own writing. Do not copy a new essay.")
else:
    st.title("Writing Assessment")
    if teacher_authenticated():
        tab1, tab2 = st.tabs(["Create Task", "Results"])
        with tab1:
            st.subheader("Teacher Setup")
            st.write("Create a task and share the student link.")
            class_name = st.text_input("Class", placeholder="e.g., English Communication A")
            task_date = st.date_input("Task Date")
            title = st.text_input("Task Title", placeholder="e.g., Use of Learning Resources")
            genre = st.text_input("Genre / Writing Type", placeholder="e.g., Email Body")
            requirements = st.text_area(
                "Task Requirements",
                placeholder="1. Explain how you have benefited from these resources.\n2. Give suggestions about what the department can do to encourage students to use them more.",
                height=120
            )
            if st.button("Create Student Link", type="primary", use_container_width=True):
                if not class_name.strip() or not title.strip() or not genre.strip() or not requirements.strip():
                    st.warning("Please complete Class, Task Title, Genre / Writing Type, and Task Requirements.")
                else:
                    code = uuid.uuid4().hex[:10]
                    tasks[code] = {
                        "class_name": class_name.strip(),
                        "task_date": task_date.isoformat(),
                        "title": title.strip(),
                        "genre": genre.strip(),
                        "requirements": requirements.strip()
                    }
                    save_tasks(tasks)
                    st.session_state["created_link"] = student_link(code)
            if st.session_state.get("created_link"):
                st.success("Task created.")
                st.markdown("**Student Link**")
                st.code(st.session_state["created_link"])
            if tasks:
                st.divider()
                st.subheader("Created Tasks")
                for code, task in reversed(list(tasks.items())):
                    with st.container(border=True):
                        st.markdown(f"**{task.get('title','')}**")
                        st.write(f"Class: {task.get('class_name','')}")
                        st.write(f"Task Date: {task.get('task_date','')}")
                        st.write(f"Genre / Writing Type: {task.get('genre','')}")
                        st.code(student_link(code))

                        c1, c2 = st.columns(2)
                        with c1:
                            if st.button("Edit Task", key=f"edit_{code}", use_container_width=True):
                                st.session_state["editing_task_code"] = code
                        with c2:
                            if st.button("Delete Task", key=f"delete_{code}", use_container_width=True):
                                st.session_state["deleting_task_code"] = code

                        if st.session_state.get("editing_task_code") == code:
                            st.markdown("### Edit Task")
                            edit_class = st.text_input(
                                "Class",
                                value=task.get("class_name", ""),
                                key=f"edit_class_{code}"
                            )
                            try:
                                from datetime import date
                                current_date = date.fromisoformat(task.get("task_date", ""))
                            except Exception:
                                current_date = task_date
                            edit_date = st.date_input(
                                "Task Date",
                                value=current_date,
                                key=f"edit_date_{code}"
                            )
                            edit_title = st.text_input(
                                "Task Title",
                                value=task.get("title", ""),
                                key=f"edit_title_{code}"
                            )
                            edit_genre = st.text_input(
                                "Genre / Writing Type",
                                value=task.get("genre", ""),
                                key=f"edit_genre_{code}"
                            )
                            edit_requirements = st.text_area(
                                "Task Requirements",
                                value=task.get("requirements", ""),
                                key=f"edit_req_{code}",
                                height=120
                            )

                            s1, s2 = st.columns(2)
                            with s1:
                                if st.button("Save Changes", key=f"save_{code}", type="primary", use_container_width=True):
                                    if not edit_class.strip() or not edit_title.strip() or not edit_genre.strip() or not edit_requirements.strip():
                                        st.warning("Please complete all task fields.")
                                    else:
                                        tasks[code] = {
                                            "class_name": edit_class.strip(),
                                            "task_date": edit_date.isoformat(),
                                            "title": edit_title.strip(),
                                            "genre": edit_genre.strip(),
                                            "requirements": edit_requirements.strip()
                                        }
                                        save_tasks(tasks)
                                        st.session_state.pop("editing_task_code", None)
                                        st.success("Task updated.")
                                        st.rerun()
                            with s2:
                                if st.button("Cancel", key=f"cancel_edit_{code}", use_container_width=True):
                                    st.session_state.pop("editing_task_code", None)
                                    st.rerun()

                        if st.session_state.get("deleting_task_code") == code:
                            st.warning("Delete this task? The student link will stop working. Existing submission records will not be deleted.")
                            d1, d2 = st.columns(2)
                            with d1:
                                if st.button("Yes, Delete", key=f"confirm_delete_{code}", use_container_width=True):
                                    tasks.pop(code, None)
                                    save_tasks(tasks)
                                    st.session_state.pop("deleting_task_code", None)
                                    st.success("Task deleted.")
                                    st.rerun()
                            with d2:
                                if st.button("Cancel", key=f"cancel_delete_{code}", use_container_width=True):
                                    st.session_state.pop("deleting_task_code", None)
                                    st.rerun()
        with tab2:
            st.subheader("Results")
            submissions = load_submissions()
            if not submissions:
                st.info("No student submissions yet.")
            else:
                classes = sorted({r.get("class_name", "") for r in submissions if r.get("class_name", "")})
                titles = sorted({r.get("task_title", "") for r in submissions if r.get("task_title", "")})
                dates = sorted({r.get("task_date", "") for r in submissions if r.get("task_date", "")})
                c1, c2, c3 = st.columns(3)
                with c1:
                    class_filter = st.selectbox("Class", ["All"] + classes)
                with c2:
                    date_filter = st.selectbox("Task Date", ["All"] + dates)
                with c3:
                    title_filter = st.selectbox("Task Title", ["All"] + titles)
                filtered = []
                for r in submissions:
                    if class_filter != "All" and r.get("class_name") != class_filter: continue
                    if date_filter != "All" and r.get("task_date") != date_filter: continue
                    if title_filter != "All" and r.get("task_title") != title_filter: continue
                    filtered.append(r)
                st.write(f"Submissions: **{len(filtered)}**")
                table_rows = [{
                    "Class": r.get("class_name", ""), "Task Date": r.get("task_date", ""), "Task Title": r.get("task_title", ""),
                    "Student ID": r.get("student_id", ""), "Student Name": r.get("student_name", ""),
                    "Content": r.get("content_score", ""), "Organization": r.get("organization_score", ""),
                    "Language": r.get("language_score", ""), "Genre": r.get("genre_score", ""),
                    "Total": r.get("total", ""), "Submitted": r.get("submitted_at", "")
                } for r in filtered]
                st.dataframe(table_rows, use_container_width=True, hide_index=True)
                if filtered:
                    st.download_button(
                        "Download Results (Excel)",
                        data=build_excel(filtered),
                        file_name="writing_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                st.warning("Current prototype storage is local to the Streamlit app. For semester-long use, results should be moved to a persistent database before relying on this as the only grade record.")
