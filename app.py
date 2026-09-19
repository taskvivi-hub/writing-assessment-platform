import os
import json
import base64
import html
from datetime import date
from io import BytesIO

import streamlit as st
from openai import OpenAI
from openpyxl import Workbook
import requests

APP_NAME = "Writing Assessment"

RUBRIC = {
    "Content & Task Fulfillment": {
        4: (
            "Fully addresses all task requirements. Content is relevant, sufficiently developed, and supported with specific information, examples, or evidence appropriate to the academic or professional purpose.",
            "完整回應所有任務要求。內容切題且發展充分，並提供符合學術或專業目的的具體資訊、例子或證據作為支持。"
        ),
        3: (
            "Addresses the main task requirements. Content is relevant and adequately developed, with enough supporting information to complete the task successfully.",
            "回應主要的任務要求。內容切題且有適當發展，並提供足夠的相關資訊，能夠完成此項任務。"
        ),
        2: (
            "Addresses only part of the task or develops ideas unevenly. Support may be limited, repetitive, or insufficient for the intended purpose.",
            "僅回應部分任務要求，或想法發展不均。支持內容可能有限、重複，或不足以達成任務目的。"
        ),
        1: (
            "Does not adequately address the task. Content is minimal, off-topic, or too incomplete to fulfill the required academic or professional purpose.",
            "未能充分回應任務。內容過少、偏離主題，或過於不完整，無法達成所要求的學術或專業目的。"
        )
    },
    "Organization & Coherence": {
        4: (
            "Information is well organized at both paragraph and whole-text levels. Ideas progress logically, transitions are effective, and the reader can follow the argument or message without difficulty.",
            "段落與全文的資訊組織良好。想法發展有邏輯，轉承有效，讀者能輕鬆理解文章的論述或訊息。"
        ),
        3: (
            "Organization is generally clear. Paragraphing and sequencing are appropriate, and most ideas are connected logically, with only minor lapses in coherence.",
            "整體組織大致清楚。段落安排與內容順序適當，大部分想法之間具有合理連結，僅有少數地方銜接不夠順暢。"
        ),
        2: (
            "Some organization is evident, but paragraphing, sequencing, or transitions are inconsistent and occasionally make the text difficult to follow.",
            "可看出部分組織安排，但段落、內容順序或轉承不一致，有時會使文章較難理解。"
        ),
        1: (
            "Organization is weak or unclear. Ideas are fragmented, poorly sequenced, or insufficiently connected, making the text difficult to follow.",
            "組織薄弱或不清楚。想法零散、順序不佳或缺乏足夠連結，使文章難以理解。"
        )
    },
    "Language Use": {
        4: (
            "Uses an effective range of vocabulary and sentence structures, including appropriate academic or professional language. Word choice is generally precise, and errors are minor and do not affect meaning.",
            "能有效運用多樣的字彙與句型，包括適當的學術或專業用語。用字大致精確，錯誤輕微且不影響意思。"
        ),
        3: (
            "Uses sufficient vocabulary and sentence structures to complete the task. Academic or professional language is generally appropriate, and errors rarely interfere with meaning.",
            "能使用足夠的字彙與句型完成任務。學術或專業用語大致適當，錯誤很少影響意思理解。"
        ),
        2: (
            "Uses a limited range of vocabulary or structures. Repetition, imprecise wording, or frequent errors sometimes reduce clarity or accuracy.",
            "字彙或句型的運用範圍較有限。重複、不精確的用字或較頻繁的錯誤，有時會降低表達的清楚度或正確性。"
        ),
        1: (
            "Language resources are too limited for the task. Frequent or serious errors in wording or sentence construction make important parts of the text difficult to understand.",
            "語言能力不足以完成任務。用字或句子結構出現頻繁或嚴重錯誤，使文章的重要部分難以理解。"
        )
    },
    "Genre & Professional Appropriacy": {
        4: (
            "Consistently follows the expected purpose, organization, format, tone, and conventions of the assigned genre. The writing is well suited to its intended audience and professional or academic context.",
            "能一致地符合指定文類的目的、組織、格式、語氣與慣例。文章非常適合預定讀者以及專業或學術情境。"
        ),
        3: (
            "Generally follows the expected purpose, organization, format, and tone of the genre. Minor inconsistencies do not interfere with the intended communication.",
            "大致符合該文類預期的目的、組織、格式與語氣。少數不一致之處不影響原本的溝通目的。"
        ),
        2: (
            "Shows partial control of the genre. Format, tone, organization, or audience awareness is inconsistent and sometimes weakens the effectiveness of the text.",
            "對該文類僅有部分掌握。格式、語氣、組織或讀者意識不一致，有時會降低文章的溝通效果。"
        ),
        1: (
            "Shows limited awareness of the assigned genre. Format, tone, organization, or audience expectations are frequently inappropriate for the task.",
            "對指定文類的掌握有限。格式、語氣、組織或讀者期待經常不符合任務需求。"
        )
    }
}



def secret_or_env(name, default=""):
    try:
        value = st.secrets.get(name, default)
    except Exception:
        value = default
    return value or os.getenv(name, default)


def get_supabase_config():
    url = secret_or_env("SUPABASE_URL").rstrip("/")
    key = secret_or_env("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise RuntimeError("Supabase is not configured.")
    return url, key


def supabase_request(method, table, *, params=None, json_body=None, prefer=None):
    url, key = get_supabase_config()
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer

    response = requests.request(
        method,
        f"{url}/rest/v1/{table}",
        headers=headers,
        params=params,
        json=json_body,
        timeout=30,
    )

    if not response.ok:
        detail = response.text.strip()
        raise RuntimeError(
            f"Supabase request failed ({response.status_code}). "
            f"{detail[:500]}"
        )

    if not response.content:
        return []
    try:
        return response.json()
    except ValueError:
        return []


def list_tasks():
    rows = supabase_request(
        "GET",
        "tasks",
        params={"select": "*", "order": "created_at.desc"},
    )
    return rows or []


def get_task(task_id):
    rows = supabase_request(
        "GET",
        "tasks",
        params={"select": "*", "id": f"eq.{task_id}", "limit": "1"},
    )
    return rows[0] if rows else None


def create_task_record(class_name, task_date, title, genre, requirements):
    payload = {
        "class_name": class_name.strip(),
        "task_date": task_date.isoformat(),
        "title": title.strip(),
        "genre": genre.strip(),
        "requirements": requirements.strip(),
    }
    rows = supabase_request(
        "POST",
        "tasks",
        json_body=payload,
        prefer="return=representation",
    )
    if not rows:
        raise RuntimeError("Task was not saved.")
    return rows[0]


def update_task_record(task_id, class_name, task_date, title, genre, requirements):
    payload = {
        "class_name": class_name.strip(),
        "task_date": task_date.isoformat(),
        "title": title.strip(),
        "genre": genre.strip(),
        "requirements": requirements.strip(),
    }
    supabase_request(
        "PATCH",
        "tasks",
        params={"id": f"eq.{task_id}"},
        json_body=payload,
        prefer="return=minimal",
    )


def delete_task_record(task_id):
    supabase_request(
        "DELETE",
        "tasks",
        params={"id": f"eq.{task_id}"},
        prefer="return=minimal",
    )


def list_submissions():
    rows = supabase_request(
        "GET",
        "submissions",
        params={"select": "*", "order": "submitted_at.desc"},
    )
    return rows or []

def get_base_url():
    return secret_or_env("APP_BASE_URL", "").rstrip("/")


def student_link(task_id):
    base = get_base_url()
    return f"{base}/?task={task_id}" if base else f"?task={task_id}"


def image_to_data_url(uploaded_file):
    mime = uploaded_file.type or "image/jpeg"
    b64 = base64.b64encode(uploaded_file.getvalue()).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def rubric_for_prompt():
    lines = []
    for dim, levels in RUBRIC.items():
        lines.append(dim)
        for score in (4, 3, 2, 1):
            english_desc, _ = levels[score]
            lines.append(f"{score}: {english_desc}")
    return "\n".join(lines)


def score_overview_chart(scores):
    colors = {
        "Content & Task Fulfillment": "#4E79A7",
        "Organization & Coherence": "#59A14F",
        "Language Use": "#F28E2B",
        "Genre & Professional Appropriacy": "#B07AA1",
    }
    short_labels = {
        "Content & Task Fulfillment": "Content",
        "Organization & Coherence": "Organization",
        "Language Use": "Language",
        "Genre & Professional Appropriacy": "Genre",
    }

    bars = []
    for dim in RUBRIC:
        score = int(scores[dim])
        height = 36 * score
        bars.append(
            '<div class="scorebar-item">'
            f'<div class="scorebar-value">{score}/4</div>'
            '<div class="scorebar-track">'
            f'<div class="scorebar-fill" style="height:{height}px;background:{colors[dim]};"></div>'
            '</div>'
            f'<div class="scorebar-label">{html.escape(short_labels[dim])}</div>'
            '</div>'
        )

    return '<div class="scorechart">' + ''.join(bars) + '</div>'


def build_prompt(task):
    requirements = task.get("requirements", "").strip()
    rubric_text = rubric_for_prompt()

    return f"""You are an English writing assessor. Read the student's uploaded composition carefully.

TEACHER SETTINGS
Task Title: {task["title"]}
Genre / Writing Type: {task["genre"]}
Task Requirements:
{requirements}

OFFICIAL RUBRIC
Use these descriptors exactly as the basis for scoring:

{rubric_text}

ASSESSMENT RULES
Evaluate exactly these four dimensions. Give an INTEGER score from 1 to 4 for each:
1. Content & Task Fulfillment
2. Organization & Coherence
3. Language Use
4. Genre & Professional Appropriacy

SCORING CALIBRATION
- Score each dimension by matching the student's actual performance to the closest official descriptor.
- Do NOT deliberately score generously or harshly.
- Do NOT lower a score merely because the writing is simple or written by an A2-B1 learner.
- Do NOT raise a score merely because the response is understandable or attempts the task.
- A score of 1 is appropriate only when the score-1 descriptor is genuinely the closest match.
- A score of 2 is appropriate when there is partial control, limited development, inconsistent organization, limited language range, or frequent errors as described in the rubric.
- A score of 3 is appropriate when the score-3 descriptor is genuinely met overall.
- A score of 4 is appropriate only when the score-4 descriptor is genuinely met overall.
- Judge every dimension independently. One weak dimension must not automatically lower the others.
- Do not double-penalize language errors under Content, Organization, or Genre unless those errors actually affect that dimension.
- For Content & Task Fulfillment, focus on whether the required points are addressed and sufficiently developed.
- For Organization & Coherence, focus on sequencing, paragraphing, progression, and logical connections. Simple writing can still earn 2 or 3 if its organization matches those descriptors.
- For Language Use, consider range, accuracy, and how much errors affect clarity. Frequent errors can still be score 2 when the main meaning remains understandable.
- For Genre & Professional Appropriacy, focus on purpose, tone, audience awareness, format, and genre conventions actually required by the task.

IMPORTANT TASK RULES
- No half points.
- Do not add criteria.
- Judge Content & Task Fulfillment against the stated Task Requirements.
- Evaluate only the writing the student is required to produce.
- Some assignments may already provide fixed genre elements outside the student's response, such as a subject line, greeting, opening, closing, or signature.
- Do NOT penalize a student for omitting any element that is not explicitly required in Task Requirements.
- If the Genre / Writing Type is "Email Body", evaluate only the body paragraphs for appropriate purpose, organization, tone, audience awareness, and professional/academic appropriacy. Do not require a greeting, closing, or signature.
- If the image is not readable enough, do not guess.

LANGUAGE CORRECTIONS
Identify ALL clear, genuine errors in these three categories:
- Grammar
- Spelling
- Punctuation

Do not impose an artificial maximum number of corrections. If there are many genuine errors, list all of them.
For every error:
- identify the category,
- quote the student's original wording,
- give the corrected wording,
- give one short Traditional Chinese explanation.
Do not rewrite the full composition.
Preserve the student's intended meaning.
Do not list style preferences as errors.
Do not invent errors when the original wording is acceptable.

CONTENT & ORGANIZATION REVISION
Provide 1 to 3 useful revision suggestions about CONTENT and ORGANIZATION only.
For each suggestion, provide:
- a short A2-B1 English suggestion,
- a clear Traditional Chinese translation.
Do not provide a model essay.

Return VALID JSON ONLY.
The values "<score 1-4>" below are placeholders. Replace each one with the student's actual integer score.

{{
  "image_readable": true,
  "transcription": "faithful transcription",
  "scores": {{
    "Content & Task Fulfillment": "<score 1-4>",
    "Organization & Coherence": "<score 1-4>",
    "Language Use": "<score 1-4>",
    "Genre & Professional Appropriacy": "<score 1-4>"
  }},
  "corrections": [
    {{
      "type": "Grammar",
      "original": "student text",
      "correction": "corrected text",
      "explanation_zh": "簡短的繁體中文說明"
    }}
  ],
  "content_organization_suggestions": [
    {{
      "en": "short A2-B1 English suggestion",
      "zh": "清楚的繁體中文翻譯"
    }}
  ]
}}"""


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



def save_submission(task, student_id, student_name, result):
    payload = {
        "task_id": task["id"],
        "class_name": task.get("class_name", ""),
        "task_date": task.get("task_date"),
        "task_title": task.get("title", ""),
        "genre": task.get("genre", ""),
        "student_id": student_id.strip(),
        "student_name": student_name.strip(),
        "content_score": result["scores"]["Content & Task Fulfillment"],
        "organization_score": result["scores"]["Organization & Coherence"],
        "language_score": result["scores"]["Language Use"],
        "genre_score": result["scores"]["Genre & Professional Appropriacy"],
        "total": sum(result["scores"].values()),
    }
    supabase_request(
        "POST",
        "submissions",
        json_body=payload,
        prefer="return=minimal",
    )


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
.cardtitle{font-size:1.05rem;font-weight:700}.score{float:right;font-weight:800}.meta{opacity:.75}
.desc-en{margin-top:.65rem;line-height:1.55}
.desc-zh{margin-top:.5rem;line-height:1.65;opacity:.88}
.scorechart{display:flex;justify-content:space-around;align-items:flex-end;gap:12px;border:1px solid rgba(120,120,120,.22);border-radius:14px;padding:18px 12px 12px;margin:.8rem 0 1.2rem}
.scorebar-item{width:22%;min-width:92px;text-align:center}
.scorebar-value{font-weight:800;margin-bottom:6px}
.scorebar-track{height:144px;display:flex;align-items:flex-end;justify-content:center;border-bottom:1px solid rgba(120,120,120,.35)}
.scorebar-fill{width:58px;max-width:80%;border-radius:8px 8px 0 0}
.scorebar-label{font-size:.88rem;font-weight:650;margin-top:8px;line-height:1.2}
@media (max-width:640px){.scorechart{gap:5px;padding-left:5px;padding-right:5px}.scorebar-item{min-width:0;width:25%}.scorebar-fill{width:42px}.scorebar-label{font-size:.75rem}}
</style>
''', unsafe_allow_html=True)


task_id = st.query_params.get("task")

# ---------------------------
# STUDENT VIEW
# ---------------------------
if task_id:
    try:
        task = get_task(task_id)
    except Exception as e:
        task = None
        st.error(f"Database connection error: {e}")

    if not task:
        st.title("Writing Assessment")
        st.error("This task link is not available.")
    else:
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
                            save_submission(task, student_id, student_name, result)
                            st.session_state["result"] = result
                            st.session_state["show_revision"] = False
                    except Exception as e:
                        st.error(f"Assessment could not be completed: {e}")

        result = st.session_state.get("result")
        if result:
            st.divider()
            st.header("Your Writing Score")
            total = sum(result["scores"][d] for d in RUBRIC)
            st.markdown(
                f'<div class="totalbox"><div class="meta">Total Score</div><div class="totalnum">{total} / 16</div></div>',
                unsafe_allow_html=True
            )

            st.subheader("Score Overview")
            st.markdown(
                score_overview_chart(result["scores"]),
                unsafe_allow_html=True
            )

            for dim, levels in RUBRIC.items():
                score = result["scores"][dim]
                desc_en, desc_zh = levels[score]
                st.markdown(
                    f'<div class="card"><span class="cardtitle">{html.escape(dim)}</span>'
                    f'<span class="score">{score}/4</span>'
                    f'<div class="desc-en">{html.escape(desc_en)}</div>'
                    f'<div class="desc-zh">{html.escape(desc_zh)}</div></div>',
                    unsafe_allow_html=True
                )

            if st.button("See Revision Suggestions", use_container_width=True):
                st.session_state["show_revision"] = True

            if st.session_state.get("show_revision"):
                st.divider()
                st.header("Revision Suggestions")

                st.subheader("Content & Organization Suggestions")
                suggestions = result.get("content_organization_suggestions", [])
                if suggestions:
                    for i, s in enumerate(suggestions, 1):
                        if isinstance(s, dict):
                            en = s.get("en", "")
                            zh = s.get("zh", "")
                            st.markdown(f"**{i}. {en}**")
                            if zh:
                                st.markdown(f"{zh}")
                        else:
                            st.markdown(f"**{i}. {s}**")
                        st.write("")
                else:
                    st.write("No additional suggestions.")

                st.subheader("Language Corrections")
                st.caption("Grammar • Spelling • Punctuation")
                corrections = result.get("corrections", [])
                if not corrections:
                    st.success("No clear grammar, spelling, or punctuation errors were found.")
                else:
                    for i, c in enumerate(corrections, 1):
                        error_type = c.get("type", "Correction")
                        original = c.get("original", "")
                        correction = c.get("correction", "")
                        explanation_zh = c.get("explanation_zh", c.get("explanation", ""))
                        st.markdown(f"**{i}. {error_type}**")
                        st.markdown(f"- **Original:** {original}")
                        st.markdown(f"- **Correction:** {correction}")
                        if explanation_zh:
                            st.markdown(f"- **說明：** {explanation_zh}")
                        st.write("")

                st.info("Revise the errors in your own writing. Do not copy a new essay.")

# ---------------------------
# TEACHER VIEW
# ---------------------------
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
                placeholder="1. Explain how you have benefited from these resources.\n"
                            "2. Give suggestions about what the department can do to encourage students to use them more.",
                height=120
            )

            if st.button("Create Student Link", type="primary", use_container_width=True):
                if not class_name.strip() or not title.strip() or not genre.strip() or not requirements.strip():
                    st.warning("Please complete Class, Task Title, Genre / Writing Type, and Task Requirements.")
                else:
                    try:
                        task = create_task_record(class_name, task_date, title, genre, requirements)
                        st.session_state["created_link"] = student_link(task["id"])
                        st.success("Task created and saved.")
                    except Exception as e:
                        st.error(f"Task could not be created: {e}")

            if st.session_state.get("created_link"):
                st.markdown("**Student Link**")
                st.code(st.session_state["created_link"])

            st.divider()
            st.subheader("Created Tasks")
            try:
                tasks = list_tasks()
            except Exception as e:
                tasks = []
                st.error(f"Tasks could not be loaded: {e}")

            if not tasks:
                st.info("No tasks yet.")
            else:
                for task in tasks:
                    code = task["id"]
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
                            edit_class = st.text_input("Class", value=task.get("class_name", ""), key=f"edit_class_{code}")
                            try:
                                current_date = date.fromisoformat(str(task.get("task_date", "")))
                            except Exception:
                                current_date = date.today()
                            edit_date = st.date_input("Task Date", value=current_date, key=f"edit_date_{code}")
                            edit_title = st.text_input("Task Title", value=task.get("title", ""), key=f"edit_title_{code}")
                            edit_genre = st.text_input("Genre / Writing Type", value=task.get("genre", ""), key=f"edit_genre_{code}")
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
                                        update_task_record(code, edit_class, edit_date, edit_title, edit_genre, edit_requirements)
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
                                    delete_task_record(code)
                                    st.session_state.pop("deleting_task_code", None)
                                    st.success("Task deleted.")
                                    st.rerun()
                            with d2:
                                if st.button("Cancel", key=f"cancel_delete_{code}", use_container_width=True):
                                    st.session_state.pop("deleting_task_code", None)
                                    st.rerun()

        with tab2:
            st.subheader("Results")
            try:
                submissions = list_submissions()
            except Exception as e:
                submissions = []
                st.error(f"Results could not be loaded: {e}")

            if not submissions:
                st.info("No student submissions yet.")
            else:
                classes = sorted({r.get("class_name", "") for r in submissions if r.get("class_name", "")})
                titles = sorted({r.get("task_title", "") for r in submissions if r.get("task_title", "")})
                dates = sorted({str(r.get("task_date", "")) for r in submissions if r.get("task_date", "")})

                c1, c2, c3 = st.columns(3)
                with c1:
                    class_filter = st.selectbox("Class", ["All"] + classes)
                with c2:
                    date_filter = st.selectbox("Task Date", ["All"] + dates)
                with c3:
                    title_filter = st.selectbox("Task Title", ["All"] + titles)

                filtered = []
                for r in submissions:
                    if class_filter != "All" and r.get("class_name") != class_filter:
                        continue
                    if date_filter != "All" and str(r.get("task_date", "")) != date_filter:
                        continue
                    if title_filter != "All" and r.get("task_title") != title_filter:
                        continue
                    filtered.append(r)

                st.write(f"Submissions: **{len(filtered)}**")

                table_rows = [{
                    "Class": r.get("class_name", ""),
                    "Task Date": r.get("task_date", ""),
                    "Task Title": r.get("task_title", ""),
                    "Student ID": r.get("student_id", ""),
                    "Student Name": r.get("student_name", ""),
                    "Content": r.get("content_score", ""),
                    "Organization": r.get("organization_score", ""),
                    "Language": r.get("language_score", ""),
                    "Genre": r.get("genre_score", ""),
                    "Total": r.get("total", ""),
                    "Submitted": r.get("submitted_at", "")
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
