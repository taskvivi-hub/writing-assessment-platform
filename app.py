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
            "Fully addresses all task requirements. Content is directly relevant, sufficiently developed, and supported with specific information, examples, or evidence appropriate to the assigned task and purpose.",
            "完整回應所有任務要求。內容與指定任務直接相關且發展充分，並提供符合任務與目的的具體資訊、例子或證據作為支持。"
        ),
        3: (
            "Addresses the main task requirements. Content is relevant and adequately developed, with enough supporting information to complete the assigned task successfully.",
            "回應主要的任務要求。內容與指定任務相關且有適當發展，並提供足夠的相關資訊，能夠完成此項任務。"
        ),
        2: (
            "Addresses only part of the assigned task or develops required ideas unevenly. Important task content may be missing, weakly developed, repetitive, or only partly relevant.",
            "僅回應部分指定任務，或對必要內容的發展不均。重要任務內容可能缺漏、發展不足、重複，或只有部分與任務相關。"
        ),
        1: (
            "Does not adequately address the assigned task. The response is largely off-topic, minimal, or too incomplete to fulfill the required purpose.",
            "未能充分回應指定任務。內容大多偏題、過少，或過於不完整，無法達成任務要求的目的。"
        )
    },
    "Organization & Coherence": {
        4: (
            "The required task content is well organized at both paragraph and whole-text levels. Ideas progress logically, transitions are effective, and the assigned message or purpose is easy to follow.",
            "任務所要求的內容在段落與全文層次皆組織良好。想法發展有邏輯，轉承有效，讀者能輕鬆理解指定的訊息或目的。"
        ),
        3: (
            "The required task content is generally organized clearly. Sequencing and connections are appropriate, with only minor lapses in coherence while carrying out the assigned task.",
            "任務所要求的內容整體組織大致清楚。內容順序與連結適當，在完成指定任務時僅有少數銜接不夠順暢之處。"
        ),
        2: (
            "Some organization of task-related content is evident, but sequencing, paragraphing, or transitions are inconsistent. Off-topic or loosely connected ideas sometimes weaken the assigned message.",
            "可看出部分與任務相關內容的組織安排，但內容順序、段落或轉承不一致。偏題或連結鬆散的想法有時會削弱指定訊息的表達。"
        ),
        1: (
            "The response does not organize the required task content effectively. Ideas are largely off-topic, fragmented, poorly sequenced, or insufficiently connected for the assigned purpose.",
            "未能有效組織任務所要求的內容。想法大多偏題、零散、順序不佳或缺乏足夠連結，無法達成指定目的。"
        )
    },
    "Language Use": {
        4: (
            "Uses an effective range of vocabulary and sentence structures to express the assigned task content accurately and appropriately. Word choice is generally precise, and errors are minor and do not affect the intended meaning.",
            "能有效運用多樣的字彙與句型，正確且適切地表達指定任務內容。用字大致精確，錯誤輕微且不影響預定意思。"
        ),
        3: (
            "Uses sufficient vocabulary and sentence structures to communicate the assigned task content. Language is generally appropriate to the purpose, and errors rarely interfere with the intended meaning.",
            "能使用足夠的字彙與句型表達指定任務內容。語言大致符合任務目的，錯誤很少影響預定意思。"
        ),
        2: (
            "Language only partly supports the assigned task. Vocabulary or structures are limited, repetitive, imprecise, or not consistently relevant to the required content, and frequent errors sometimes reduce clarity.",
            "語言僅能部分支援指定任務。字彙或句型較有限、重複、不精確，或未能持續用於表達任務所要求的內容；較頻繁的錯誤有時會降低清楚度。"
        ),
        1: (
            "Language does not adequately support the assigned task. The response may be largely off-topic, too limited, or contain frequent or serious errors that prevent effective communication of the required content.",
            "語言未能充分支援指定任務。內容可能大多偏題、語言資源過少，或有頻繁／嚴重錯誤，使任務所要求的內容無法有效傳達。"
        )
    },
    "Genre & Professional Appropriacy": {
        4: (
            "Consistently fulfills the assigned communicative purpose and follows the expected organization, format, tone, and conventions of the specified genre. The writing is well suited to its intended audience and context.",
            "能一致地達成指定的溝通目的，並符合指定文類預期的組織、格式、語氣與慣例。文章非常適合預定讀者與情境。"
        ),
        3: (
            "Generally fulfills the assigned communicative purpose and follows the expected organization, format, and tone of the specified genre. Minor inconsistencies do not interfere with the intended communication.",
            "大致能達成指定的溝通目的，並符合指定文類預期的組織、格式與語氣。少數不一致之處不影響原本的溝通目的。"
        ),
        2: (
            "Shows partial control of the assigned genre and purpose. Format, tone, organization, audience awareness, or task relevance is inconsistent and sometimes weakens communication.",
            "對指定文類與目的僅有部分掌握。格式、語氣、組織、讀者意識或與任務的相關性不一致，有時會降低溝通效果。"
        ),
        1: (
            "Shows limited awareness of the assigned genre and communicative purpose. The response is largely inappropriate, off-task, or inconsistent with the expected audience, format, tone, or conventions.",
            "對指定文類與溝通目的的掌握有限。內容大多不適切、偏離任務，或與預期讀者、格式、語氣或文類慣例不一致。"
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
    if task_has_submissions(task_id):
        raise RuntimeError("This task already has student submissions and can no longer be edited.")

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
    if task_has_submissions(task_id):
        raise RuntimeError("This task already has student submissions and cannot be deleted.")

    supabase_request(
        "DELETE",
        "tasks",
        params={"id": f"eq.{task_id}"},
        prefer="return=minimal",
    )


def set_task_closed(task_id, is_closed):
    supabase_request(
        "PATCH",
        "tasks",
        params={"id": f"eq.{task_id}"},
        json_body={"is_closed": bool(is_closed)},
        prefer="return=minimal",
    )


def list_submissions():
    rows = supabase_request(
        "GET",
        "submissions",
        params={"select": "*", "order": "submitted_at.desc"},
    )
    return rows or []


def submission_count_for_task(task_id):
    rows = supabase_request(
        "GET",
        "submissions",
        params={
            "select": "id",
            "task_id": f"eq.{task_id}",
        },
    )
    return len(rows or [])


def task_has_submissions(task_id):
    return submission_count_for_task(task_id) > 0


def submission_exists(task_id, student_name):
    rows = supabase_request(
        "GET",
        "submissions",
        params={
            "select": "id",
            "task_id": f"eq.{task_id}",
            "student_name": f"eq.{student_name.strip()}",
            "limit": "1",
        },
    )
    return bool(rows)


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

OFFICIAL TASK-ANCHORED RUBRIC
Use these descriptors exactly as the basis for scoring. Every dimension is evaluated within the context of THIS assigned task, not as a general measure of the student's writing ability.

{rubric_text}

CORE SCORING PRINCIPLE: DYNAMIC TASK ANCHORING
The Teacher Settings above are the ONLY task-specific source for this assessment. Do not assume any fixed topic, genre, audience, or required idea from previous tasks or examples.

Before assigning any score, derive the current assessment target from THIS task only:
1. Task Title = the current topic/context.
2. Task Requirements = the required content, ideas, examples, explanations, length, and other explicit instructions.
3. Genre / Writing Type = the expected communicative purpose, organization, tone, audience awareness, and conventions.

Then evaluate ALL FOUR dimensions against this dynamically derived task target.

Do NOT treat Content as the only task-dependent dimension. Organization, Language Use, and Genre must also be judged by how effectively the student organizes, expresses, and realizes the CURRENT task content and communicative purpose.

GENERALIZATION RULE
- Never hard-code or carry over a topic, keyword, example, genre expectation, or audience from another assignment.
- If the teacher creates a different task later, all four scores must automatically be recalibrated to that new Task Title, Task Requirements, and Genre / Writing Type.
- A response that would be strong for a different assignment can still receive a low score here if it does not perform THIS assignment.

TASK-RELEVANCE GATE
Classify the response against the CURRENT Teacher Settings first:
- STRONGLY ON TASK: most or all current requirements are meaningfully addressed.
- PARTLY ON TASK: some current requirements are meaningfully addressed, but important required content is missing or underdeveloped.
- MINIMALLY ON TASK / LARGELY OFF TOPIC: the response mostly discusses material unrelated to the current requirements, merely repeats isolated words from the current prompt, or fails to carry out the current communicative purpose.

If the response is MINIMALLY ON TASK / LARGELY OFF TOPIC:
- Content should normally be 1.
- Organization should be 1 when the required task content is not meaningfully organized, even if unrelated sentences have some local order.
- Language Use should be 1 when the language does not effectively communicate the required task content. Correct grammar in unrelated content does NOT by itself justify a high Language score.
- Genre should be 1 when the response does not carry out the assigned communicative purpose or genre in a meaningful way.
- Do not reward an off-topic response with high scores simply because it is readable, grammatical, long, or locally coherent.

ASSESSMENT RULES
Evaluate exactly these four dimensions. Give an INTEGER score from 1 to 4 for each:
1. Content & Task Fulfillment
2. Organization & Coherence
3. Language Use
4. Genre & Professional Appropriacy

SCORING CALIBRATION
- Score each dimension by matching the student's task-specific performance to the closest official descriptor.
- Do NOT deliberately score generously or harshly.
- A score of 1 is appropriate when the student fails to perform that dimension adequately for the assigned task.
- A score of 2 is appropriate for partial, limited, inconsistent, or weak task performance.
- A score of 3 represents solid, generally successful performance: the main task is completed, development and control are adequate, but the writing may still be simple, uneven, limited, or contain noticeable weaknesses.
- A score of 4 is NOT the default for a complete response. It represents clearly stronger performance than score 3 and requires the score-4 descriptor to be fully and consistently demonstrated.
- Completing every listed requirement does NOT automatically justify score 4.
- If a response is mainly simple listing, thin explanation, basic development, predictable structure, limited language range, or has repeated noticeable errors, the relevant dimension should normally remain at 3 or below.
- When deciding between 3 and 4, choose 4 only when there is clear positive evidence of fuller development, stronger control, effective organization, and/or broader, more precise task-appropriate language as required by that dimension.
- Do not automatically give the same score to all four dimensions. Judge each dimension separately, but always within the same task context.

HIGH-SCORE GATE: CHECK BEFORE ANY SCORE OF 4
Before assigning 4 in any dimension, ask whether the response is clearly stronger than an adequate score-3 response for THIS task. If the answer is uncertain, assign 3 rather than 4.

For Content = 4:
- All important requirements must be addressed, not merely mentioned.
- Ideas must be sufficiently developed with specific, relevant support, explanation, examples, or details.
- A basic response that simply covers each required point without meaningful development is normally 3, not 4.

For Organization = 4:
- Organization must be consistently effective, not merely understandable.
- Ideas must progress logically and connections/transitions must actively support the task message.
- A straightforward list or simple sequence with weak or mechanical connections is normally 3 or below.

For Language = 4:
- The response must show an effective range of task-appropriate vocabulary and sentence structures, not just grammatical comprehensibility.
- Word choice should be generally precise and errors should be genuinely minor and infrequent.
- Repeated grammar, wording, spelling, punctuation, capitalization, agreement, article, tense, or sentence-structure errors are incompatible with score 4.
- If several noticeable language errors recur across the response, score 3 or below even if meaning remains clear.

For Genre = 4:
- The response must consistently fulfill the assigned communicative purpose and genre conventions for the current task.
- Tone, organization, format, audience awareness, and purpose should be consistently appropriate.
- Merely using the expected format or broadly matching the genre is normally 3, not 4, unless control is consistently strong.

DIMENSION CHECKS
Content & Task Fulfillment:
- Check every explicit Task Requirement.
- Merely mentioning the title, one keyword, or one isolated related detail does not count as meaningful task fulfillment.
- If most required content is absent or replaced by unrelated content, score 1.

Organization & Coherence:
- Evaluate the organization of the REQUIRED TASK CONTENT, not just whether unrelated sentences happen to follow one another.
- A coherent paragraph about the wrong topic is not strong task organization.
- Off-topic, fragmented, random, or loosely connected material should substantially lower this score.

Language Use:
- Evaluate vocabulary, sentence structures, accuracy, spelling, punctuation, capitalization, and how effectively the language expresses the REQUIRED TASK CONTENT.
- Correct grammar alone does not earn a high score if the language is not being used to perform the assigned task.
- Frequent grammar, spelling, punctuation, capitalization, wording, or sentence-structure problems should lower the score according to severity and frequency.
- Language 4 requires both strong task-relevant expression and an effective range of accurate language.

Genre & Professional Appropriacy:
- Evaluate whether the response actually performs the assigned communicative purpose and genre.
- Surface features alone are not enough. First person, paragraph form, or a title does not prove genre control.
- If the response does not meaningfully carry out the assigned purpose, Genre should be low.

FINAL SCORE CHECK
Before returning the JSON, verify:
- Did you use only the CURRENT Task Title, Task Requirements, and Genre / Writing Type to define what counts as successful performance?
- Is each score based on THIS task rather than on general writing ability or a previous assignment?
- Would each score still make sense if the current Teacher Settings were shown beside it?
- Have unrelated but grammatical sentences been prevented from inflating Organization or Language?
- Have frequent grammar, spelling, punctuation, and capitalization errors been reflected in Language Use?
- For every score of 4, is there clear evidence that the performance is stronger than score 3, rather than merely complete or understandable?
- Have simple listing, thin development, predictable organization, limited language range, or repeated noticeable errors prevented unjustified 4s?
- Does Genre reflect the current assigned communicative purpose rather than generic paragraph-writing conventions?

IMPORTANT TASK RULES
- No half points.
- Scores must be 1, 2, 3, or 4. There is no score 0 in the current rubric.
- Evaluate only the writing the student is required to produce.
- Some assignments may already provide fixed genre elements outside the student's response, such as a subject line, greeting, opening, closing, or signature.
- Do NOT penalize a student for omitting any element that is not explicitly required in Task Requirements.
- If the Genre / Writing Type is "Email Body", evaluate only the body paragraphs for appropriate purpose, organization, tone, audience awareness, and professional/academic appropriacy. Do not require a greeting, closing, or signature.
- If the image is not readable enough, do not guess.

LANGUAGE CORRECTIONS
Identify ALL clear, genuine errors in these categories:
- Grammar
- Spelling
- Punctuation
- Capitalization

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



def save_submission(task, seat_number, student_id, student_name, result):
    payload = {
        "task_id": task["id"],
        "class_name": task.get("class_name", ""),
        "task_date": task.get("task_date"),
        "task_title": task.get("title", ""),
        "genre": task.get("genre", ""),
        "seat_number": seat_number.strip(),
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
        "Seat No.", "Student ID", "Student Name",
        "Content & Task Fulfillment", "Organization & Coherence",
        "Language Use", "Genre & Professional Appropriacy",
        "Total /16", "Submission Time"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.get("class_name", ""), r.get("task_date", ""), r.get("task_title", ""), r.get("genre", ""),
            r.get("seat_number", ""), r.get("student_id", ""), r.get("student_name", ""),
            r.get("content_score", ""), r.get("organization_score", ""),
            r.get("language_score", ""), r.get("genre_score", ""),
            r.get("total", ""), r.get("submitted_at", "")
        ])
    widths = [18, 12, 28, 20, 10, 16, 18, 24, 24, 16, 30, 12, 22]
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
:root{
  --wa-bg:#f7fbff;
  --wa-card:#ffffff;
  --wa-text:#17324d;
  --wa-muted:#5f7183;
  --wa-border:#d8e6f2;
  --wa-primary:#246bfd;
  --wa-primary-2:#4f8cff;
  --wa-soft:#eef6ff;
  --wa-accent:#ffb84d;
  --wa-success:#2fa36b;
  --wa-shadow:0 8px 24px rgba(33,79,125,.10);
}

html, body, [class*="css"] {font-size:18px;}
.stApp{background:linear-gradient(180deg,#f8fbff 0%,#f5f9ff 45%,#ffffff 100%);color:var(--wa-text)}
.block-container{max-width:980px;padding-top:2rem;padding-bottom:4.5rem}

h1{font-size:2.3rem !important;line-height:1.15 !important;color:var(--wa-text) !important;margin-bottom:.35rem !important}
h2{font-size:1.65rem !important;color:var(--wa-text) !important;margin-top:1.25rem !important}
h3{font-size:1.25rem !important;color:var(--wa-text) !important;margin-top:1.15rem !important}
p, label, .stCaption{line-height:1.6}

/* task and result cards */
.taskbox,.card,.totalbox,.scorechart{
  background:var(--wa-card);
  color:var(--wa-text);
  border:1px solid var(--wa-border);
  border-radius:20px;
  box-shadow:var(--wa-shadow);
}
.taskbox{padding:1.25rem 1.35rem;margin:1rem 0 1.25rem;font-size:1.02rem;line-height:1.8;border-left:7px solid var(--wa-primary)}
.card{padding:1.2rem 1.35rem;margin:.9rem 0}
.totalbox{text-align:center;border-width:1px;padding:1.25rem 1rem;margin:1rem 0 1.3rem;background:linear-gradient(135deg,#eff6ff,#ffffff)}
.totalnum{font-size:2.7rem;font-weight:850;color:var(--wa-primary);letter-spacing:.02em}
.cardtitle{font-size:1.16rem;font-weight:800;color:var(--wa-text)}
.score{float:right;font-size:1.15rem;font-weight:850;color:var(--wa-primary)}
.meta{opacity:.78;font-size:1rem;color:var(--wa-muted)}
.desc-en{margin-top:.72rem;line-height:1.65;font-size:1rem;color:var(--wa-text)}
.desc-zh{margin-top:.55rem;line-height:1.7;font-size:.98rem;color:var(--wa-muted)}

/* Score chart */
.scorechart{display:flex;justify-content:space-around;align-items:flex-end;gap:16px;padding:24px 16px 16px;margin:1rem 0 1.5rem}
.scorebar-item{width:22%;min-width:100px;text-align:center}
.scorebar-value{font-size:1.05rem;font-weight:850;margin-bottom:8px;color:var(--wa-text)}
.scorebar-track{height:150px;display:flex;align-items:flex-end;justify-content:center;border-bottom:2px solid #d8e3ec}
.scorebar-fill{width:62px;max-width:82%;border-radius:10px 10px 0 0;box-shadow:0 4px 10px rgba(0,0,0,.08)}
.scorebar-label{font-size:.95rem;font-weight:800;margin-top:10px;line-height:1.2;color:var(--wa-text)}

/* Inputs */
div[data-baseweb="input"] > div,
div[data-baseweb="textarea"] > div{
  border-radius:14px !important;
  border:1.5px solid #bfd3e6 !important;
  background:#ffffff !important;
  box-shadow:0 2px 8px rgba(26,73,115,.05) !important;
}
div[data-baseweb="input"] input,
div[data-baseweb="textarea"] textarea{
  color:#16324a !important;
  font-size:1rem !important;
}
.stTextInput label,.stTextArea label,.stDateInput label,.stSelectbox label,.stFileUploader label{font-weight:800 !important;color:var(--wa-text) !important;font-size:1rem !important}

/* Buttons */
.stButton > button, .stDownloadButton > button{
  border-radius:14px !important;
  min-height:3.05rem;
  font-size:1rem !important;
  font-weight:800 !important;
  border:1px solid #c9d9e8 !important;
  box-shadow:0 4px 12px rgba(33,79,125,.08) !important;
}
.stButton > button[kind="primary"]{
  background:linear-gradient(135deg,var(--wa-primary),var(--wa-primary-2)) !important;
  color:#ffffff !important;
  border:none !important;
}
.stButton > button[kind="primary"]:hover{filter:brightness(.98)}

/* Upload area */
[data-testid="stFileUploaderDropzone"]{
  background:#f2f8ff !important;
  border:2px dashed #8fb8e6 !important;
  border-radius:18px !important;
  padding:1.1rem !important;
}
[data-testid="stFileUploaderDropzone"] *{color:#1d3d5c !important}

/* Tabs and status messages */
button[data-baseweb="tab"]{font-size:1rem !important;font-weight:800 !important}
[data-testid="stAlert"]{border-radius:16px !important;font-size:1rem !important}
hr{border-color:#dce7f0 !important;margin:1.6rem 0 !important}

/* Tables */
[data-testid="stDataFrame"]{border:1px solid var(--wa-border);border-radius:16px;overflow:hidden;box-shadow:var(--wa-shadow)}

/* Mobile */
@media (max-width:640px){
  html, body, [class*="css"]{font-size:17px}
  .block-container{padding-top:1.25rem;padding-left:1rem;padding-right:1rem}
  h1{font-size:2rem !important}
  h2{font-size:1.45rem !important}
  .taskbox{padding:1.05rem 1rem;font-size:.98rem}
  .totalnum{font-size:2.35rem}
  .scorechart{gap:6px;padding:18px 6px 14px}
  .scorebar-item{min-width:0;width:25%}
  .scorebar-fill{width:44px}
  .scorebar-label{font-size:.78rem}
  .scorebar-value{font-size:.9rem}
}

/* Dark-mode safety: keep custom cards readable even when device/browser uses dark mode */
@media (prefers-color-scheme: dark){
  .stApp{background:linear-gradient(180deg,#101827 0%,#141d2a 45%,#10151d 100%) !important;color:#f4f7fb !important}
  h1,h2,h3,.taskbox,.card,.totalbox,.scorechart,.cardtitle,.scorebar-value,.scorebar-label,.desc-en{color:#f4f7fb !important}
  .taskbox,.card,.totalbox,.scorechart{background:#182434 !important;border-color:#31445a !important;box-shadow:none !important}
  .meta,.desc-zh{color:#c4d0dc !important}
  div[data-baseweb="input"] > div,div[data-baseweb="textarea"] > div{background:#111a25 !important;border-color:#40566d !important}
  div[data-baseweb="input"] input,div[data-baseweb="textarea"] textarea{color:#f4f7fb !important}
  .stTextInput label,.stTextArea label,.stDateInput label,.stSelectbox label,.stFileUploader label{color:#eef4fa !important}
  [data-testid="stFileUploaderDropzone"]{background:#152233 !important;border-color:#587ca4 !important}
  [data-testid="stFileUploaderDropzone"] *{color:#eef4fa !important}
  .scorebar-track{border-bottom-color:#51657a !important}
}
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

        if task.get("is_closed", False):
            st.warning("This task is closed. Submissions are no longer accepted.")
            st.info("If you think you still need to submit, please contact your teacher.")
            st.stop()

        st.subheader("Step 1. Enter your information")
        st.caption("Use your real name. Each student can submit this task only once.")
        info_c1, info_c2, info_c3 = st.columns([1, 2.2, 2.2])
        with info_c1:
            seat_number = st.text_input("Seat No.", max_chars=2, placeholder="e.g., 1 or 12")
        with info_c2:
            student_id = st.text_input("Student ID", max_chars=9, placeholder="9 characters")
        with info_c3:
            student_name = st.text_input("Student Name")

        st.subheader("Step 2. Upload your writing")
        st.write("Take a clear photo of your writing and upload it here.")
        uploaded = st.file_uploader("Choose an image", type=["jpg", "jpeg", "png", "webp"])

        st.subheader("Step 3. Submit your writing")
        if st.button("Submit for Assessment", type="primary", use_container_width=True):
            if not seat_number.strip() or not student_id.strip() or not student_name.strip():
                st.warning("Please enter your Seat No., Student ID, and Student Name.")
            elif not (seat_number.strip().isascii() and seat_number.strip().isdigit() and 1 <= len(seat_number.strip()) <= 2):
                st.warning("Seat No. must be 1 or 2 digits, for example 1, 2, or 12.")
            elif len(student_id.strip()) != 9:
                st.warning("Student ID must contain exactly 9 characters. Letters and numbers are both accepted.")
            elif not uploaded:
                st.warning("Please upload an image first.")
            else:
                try:
                    if submission_exists(task["id"], student_name):
                        st.error("This student has already submitted this task. Each student can submit only once.")
                    else:
                        with st.spinner("Processing your writing... Please do not click Submit again."):
                            result = assess(uploaded, task)
                            if not result.get("image_readable", True):
                                st.error("The image is not clear enough to read. Please upload a clearer photo.")
                            else:
                                save_submission(task, seat_number, student_id, student_name, result)
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

                        is_closed = bool(task.get("is_closed", False))
                        status_label = "Closed" if is_closed else "Open"
                        try:
                            submission_count = submission_count_for_task(code)
                        except Exception as e:
                            submission_count = None
                            st.error(f"Submission count could not be loaded: {e}")

                        if submission_count is None:
                            st.caption(f"Status: {status_label}")
                            has_submissions = True
                        else:
                            st.caption(f"Status: {status_label} · Submissions: {submission_count}")
                            has_submissions = submission_count > 0

                        if has_submissions:
                            st.session_state.pop("editing_task_code", None)
                            st.session_state.pop("deleting_task_code", None)

                        c1, c2, c3 = st.columns(3)
                        with c1:
                            if st.button(
                                "Edit Task",
                                key=f"edit_{code}",
                                use_container_width=True,
                                disabled=has_submissions,
                                help="Editing is locked after the first student submission." if has_submissions else None,
                            ):
                                st.session_state["editing_task_code"] = code
                        with c2:
                            close_label = "Reopen Task" if is_closed else "Close Task"
                            if st.button(close_label, key=f"close_{code}", use_container_width=True):
                                set_task_closed(code, not is_closed)
                                st.success("Task reopened." if is_closed else "Task closed. Students can no longer submit using this link.")
                                st.rerun()
                        with c3:
                            if st.button(
                                "Delete Task",
                                key=f"delete_{code}",
                                use_container_width=True,
                                disabled=has_submissions,
                                help="Deletion is locked after the first student submission." if has_submissions else None,
                            ):
                                st.session_state["deleting_task_code"] = code

                        if has_submissions:
                            st.caption("Edit and Delete are locked because this task already has student submissions. You can still Close or Reopen it.")

                        if st.session_state.get("editing_task_code") == code and not has_submissions:
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
                                        try:
                                            update_task_record(code, edit_class, edit_date, edit_title, edit_genre, edit_requirements)
                                            st.session_state.pop("editing_task_code", None)
                                            st.success("Task updated.")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"Task could not be updated: {e}")
                            with s2:
                                if st.button("Cancel", key=f"cancel_edit_{code}", use_container_width=True):
                                    st.session_state.pop("editing_task_code", None)
                                    st.rerun()

                        if st.session_state.get("deleting_task_code") == code and not has_submissions:
                            st.warning("Delete this task? The student link will stop working. Existing submission records will not be deleted.")
                            d1, d2 = st.columns(2)
                            with d1:
                                if st.button("Yes, Delete", key=f"confirm_delete_{code}", use_container_width=True):
                                    try:
                                        delete_task_record(code)
                                        st.session_state.pop("deleting_task_code", None)
                                        st.success("Task deleted.")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Task could not be deleted: {e}")
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
                    "Seat No.": r.get("seat_number", ""),
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
