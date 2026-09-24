import os
import json
import base64
import html
import hashlib
import time
from datetime import date
from io import BytesIO

from PIL import Image, ImageOps
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

    last_error = None
    for attempt in range(3):
        try:
            response = requests.request(
                method,
                f"{url}/rest/v1/{table}",
                headers=headers,
                params=params,
                json=json_body,
                timeout=45,
            )
            if response.ok:
                if not response.content:
                    return []
                try:
                    return response.json()
                except ValueError:
                    return []

            detail = response.text.strip()
            last_error = RuntimeError(
                f"Supabase request failed ({response.status_code}). {detail[:500]}"
            )
            if response.status_code not in (408, 425, 429, 500, 502, 503, 504):
                raise last_error
        except (requests.Timeout, requests.ConnectionError) as e:
            last_error = e

        if attempt < 2:
            time.sleep(1.2 * (attempt + 1))

    raise RuntimeError(f"Supabase request failed after retries. {last_error}")




STORAGE_BUCKET = "writing-submissions"


def prepare_storage_image(uploaded_file):
    raw = uploaded_file.getvalue()
    try:
        image = Image.open(BytesIO(raw))
        image = ImageOps.exif_transpose(image)
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        elif image.mode == "L":
            image = image.convert("RGB")

        max_side = 1800
        if max(image.size) > max_side:
            scale = max_side / max(image.size)
            new_size = (max(1, int(image.width * scale)), max(1, int(image.height * scale)))
            image = image.resize(new_size)

        out = BytesIO()
        image.save(out, format="JPEG", quality=82, optimize=True)
        return out.getvalue(), "image/jpeg", ".jpg"
    except Exception:
        mime = uploaded_file.type or "image/jpeg"
        suffix = ".png" if "png" in mime else ".webp" if "webp" in mime else ".jpg"
        return raw, mime, suffix


def upload_submission_image(task_id, student_id, student_name, uploaded_file):
    url, key = get_supabase_config()
    image_bytes, mime, suffix = prepare_storage_image(uploaded_file)
    safe_student_id = "".join(ch for ch in student_id.strip() if ch.isalnum() or ch in ("-", "_")) or "student"
    normalized_name = " ".join(student_name.strip().lower().split())
    name_hash = hashlib.sha256(normalized_name.encode("utf-8")).hexdigest()[:12]
    object_path = f"{task_id}/{safe_student_id}_{name_hash}{suffix}"
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": mime,
        # Idempotent retry: if a previous attempt stored the image but failed later,
        # a retry should replace the same object instead of returning 409 Duplicate.
        "x-upsert": "true",
    }

    last_error = None
    for attempt in range(3):
        try:
            response = requests.post(
                f"{url}/storage/v1/object/{STORAGE_BUCKET}/{object_path}",
                headers=headers,
                data=image_bytes,
                timeout=90,
            )
            if response.ok:
                return object_path

            detail = response.text.strip()
            last_error = RuntimeError(
                f"Submission image could not be stored ({response.status_code}). {detail[:500]}"
            )
            if response.status_code not in (408, 425, 429, 500, 502, 503, 504):
                raise last_error
        except (requests.Timeout, requests.ConnectionError) as e:
            last_error = e

        if attempt < 2:
            time.sleep(1.5 * (attempt + 1))

    raise RuntimeError(f"Submission image could not be stored after retries. {last_error}")


def download_submission_image(object_path):
    if not object_path:
        return None
    url, key = get_supabase_config()
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
    }
    response = requests.get(
        f"{url}/storage/v1/object/{STORAGE_BUCKET}/{object_path}",
        headers=headers,
        timeout=60,
    )
    if not response.ok:
        return None
    return response.content


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


def response_requires_zero(transcription):
    text = (transcription or "").strip()
    if not text:
        return False
    latin_letters = sum(ch.isascii() and ch.isalpha() for ch in text)
    cjk_chars = sum(("\u4e00" <= ch <= "\u9fff") or ("\u3400" <= ch <= "\u4dbf") for ch in text)
    # If the response is predominantly Chinese/non-English rather than English,
    # treat it as a language-of-response violation for this English writing assessment.
    return cjk_chars >= 8 and cjk_chars > latin_letters


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

LANGUAGE-OF-RESPONSE OVERRIDE
This is an ENGLISH writing assessment.
- If the student writes the response predominantly in Chinese or another non-English language instead of English, set ALL FOUR dimension scores to 0.
- This 0-score rule overrides the normal 1-4 rubric because the student did not provide an English response to the assigned English writing task.
- Do not apply this rule for an English response that only contains a few isolated Chinese words, names, labels, or necessary proper nouns.
- Preserve the student's original language faithfully in transcription; do not translate it before deciding this rule.

ASSESSMENT RULES
Evaluate exactly these four dimensions. Normally give an INTEGER score from 1 to 4 for each, except the language-of-response override above may require 0 for all four:
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
- Normal rubric scores are 1, 2, 3, or 4. Score 0 is reserved only for the language-of-response override when the response is predominantly not written in English.
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
The values "<score 0-4>" below are placeholders. Use 0 only when the language-of-response override applies; otherwise use 1-4.

{{
  "image_readable": true,
  "transcription": "faithful transcription",
  "english_response": true,
  "scores": {{
    "Content & Task Fulfillment": "<score 0-4>",
    "Organization & Coherence": "<score 0-4>",
    "Language Use": "<score 0-4>",
    "Genre & Professional Appropriacy": "<score 0-4>"
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
    request_input = [{
        "role": "user",
        "content": [
            {"type": "input_text", "text": build_prompt(task)},
            {"type": "input_image", "image_url": image_to_data_url(uploaded_file)}
        ]
    }]

    last_error = None
    result = None

    # Whole-class submissions can arrive in a short burst. Retry transient OpenAI
    # failures (rate limits, timeouts, temporary server errors) with backoff.
    for attempt in range(3):
        try:
            response = client.responses.create(
                model=model,
                input=request_input,
                timeout=120,
            )
            raw = response.output_text.strip()
            if raw.startswith("```"):
                raw = raw.strip("`").strip()
                if raw.lower().startswith("json"):
                    raw = raw[4:].strip()
            result = json.loads(raw)
            break
        except Exception as e:
            last_error = e
            status_code = getattr(e, "status_code", None)
            retryable = status_code in (408, 409, 425, 429, 500, 502, 503, 504)
            name = type(e).__name__.lower()
            if any(token in name for token in ("timeout", "connection", "ratelimit", "internalserver")):
                retryable = True

            # A malformed/truncated JSON response can also happen during a transient
            # failure. One or two retries are safer than immediately failing a student.
            if isinstance(e, json.JSONDecodeError):
                retryable = True

            if not retryable or attempt == 2:
                break

            time.sleep(1.5 * (2 ** attempt))

    if result is None:
        raise RuntimeError(
            "The assessment service is temporarily busy or unavailable. "
            "Please wait about 20–30 seconds and try again. "
            "This attempt has not been recorded as a submission."
        ) from last_error

    transcription = result.get("transcription", "")
    model_says_english = result.get("english_response", True)
    force_zero = (model_says_english is False) or response_requires_zero(transcription)

    if force_zero:
        result["english_response"] = False
        result["scores"] = {dim: 0 for dim in RUBRIC}
        result["corrections"] = []
        result["content_organization_suggestions"] = []
    else:
        result["english_response"] = True
        for dim in RUBRIC:
            score = int(result["scores"][dim])
            if score not in (1, 2, 3, 4):
                raise ValueError(f"Invalid score for {dim}")
            result["scores"][dim] = score
    return result



def save_submission(task, seat_number, student_id, student_name, result, image_path):
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
        "image_path": image_path,
        "transcription": result.get("transcription", ""),
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

/* Student hero + icon-led section headings */
.wa-hero{
  text-align:center;
  padding:.35rem 0 .9rem;
}
.wa-hero-icon{
  width:92px;height:92px;margin:0 auto .65rem;
  display:flex;align-items:center;justify-content:center;
  border-radius:28px;
  background:linear-gradient(135deg,#eaf3ff,#dfeeff);
  border:1px solid #c8ddf2;
  box-shadow:0 10px 28px rgba(36,107,253,.12);
  font-size:3.15rem;line-height:1;
}
.wa-hero-title{
  font-size:2.35rem;font-weight:850;line-height:1.15;color:var(--wa-text);
  letter-spacing:-.02em;
}
.wa-hero-subtitle{
  margin-top:.45rem;color:var(--wa-muted);font-size:1.14rem;line-height:1.6;font-weight:600;
}
.wa-section-title{
  display:flex;align-items:center;gap:.65rem;
  font-size:1.35rem;font-weight:850;color:var(--wa-text);
  margin:1.35rem 0 .55rem;
}
.wa-section-icon{
  width:36px;height:36px;display:inline-flex;align-items:center;justify-content:center;
  border-radius:11px;background:var(--wa-soft);border:1px solid #cfe0f1;
  font-size:1.15rem;flex:0 0 auto;
}
.wa-field-label{font-weight:800;color:var(--wa-text);font-size:1.1rem;margin-bottom:.25rem}
.wa-instruction{font-size:1.1rem;line-height:1.7;color:var(--wa-muted);font-weight:600;margin:.2rem 0 .7rem}
.wa-requirements{font-size:1.12rem;line-height:1.85;color:var(--wa-text);margin:.15rem 0 .85rem}

/* task and result cards */
.taskbox,.card,.totalbox,.scorechart{
  background:var(--wa-card);
  color:var(--wa-text);
  border:1px solid var(--wa-border);
  border-radius:20px;
  box-shadow:var(--wa-shadow);
}
.taskbox{padding:1.3rem 1.4rem;margin:1rem 0 1.25rem;font-size:1.13rem;line-height:1.9;border-left:7px solid var(--wa-primary)}
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
.stTextInput label,.stTextArea label,.stDateInput label,.stSelectbox label,.stFileUploader label{font-weight:800 !important;color:var(--wa-text) !important;font-size:1.1rem !important}

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
[data-testid="stFileUploaderDropzone"] *{color:#1d3d5c !important;font-size:1.02rem !important}

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
  .taskbox{padding:1.05rem 1rem;font-size:1.05rem}
  .wa-hero-subtitle{font-size:1.05rem}
  .wa-instruction{font-size:1.03rem}
  .wa-requirements{font-size:1.05rem}
  .stTextInput label,.stFileUploader label{font-size:1.03rem !important}
  .totalnum{font-size:2.35rem}
  .scorechart{gap:6px;padding:18px 6px 14px}
  .scorebar-item{min-width:0;width:25%}
  .scorebar-fill{width:44px}
  .scorebar-label{font-size:.78rem}
  .scorebar-value{font-size:.9rem}
  .wa-hero-icon{width:78px;height:78px;font-size:2.65rem;border-radius:23px}
  .wa-hero-title{font-size:2rem}
  .wa-section-title{font-size:1.2rem}
}

/* Force one consistent light UI on every device/browser mode.
   This prevents iOS/Android dark mode from turning labels or input text invisible. */
:root, html, body, .stApp{
  color-scheme: light !important;
}
html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"]{
  background:linear-gradient(180deg,#f8fbff 0%,#f5f9ff 45%,#ffffff 100%) !important;
  color:#17324d !important;
}
/* General Streamlit text */
.stApp p,.stApp span,.stApp div,.stApp label,.stApp small,
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"]{
  color:#17324d;
}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5,.stApp h6{
  color:#17324d !important;
}
/* Field labels and helper text */
.stTextInput label,.stTextArea label,.stDateInput label,.stSelectbox label,.stFileUploader label,
[data-testid="stWidgetLabel"] *, [data-testid="stCaptionContainer"] *{
  color:#17324d !important;
  -webkit-text-fill-color:#17324d !important;
  opacity:1 !important;
}
/* Text, date, select and textarea controls */
div[data-baseweb="input"] > div,
div[data-baseweb="textarea"] > div,
div[data-baseweb="select"] > div,
[data-baseweb="base-input"]{
  background:#ffffff !important;
  color:#16324a !important;
  border-color:#bfd3e6 !important;
}
div[data-baseweb="input"] input,
div[data-baseweb="textarea"] textarea,
div[data-baseweb="select"] input,
input, textarea{
  color:#16324a !important;
  -webkit-text-fill-color:#16324a !important;
  caret-color:#16324a !important;
  background:#ffffff !important;
}
input::placeholder, textarea::placeholder{
  color:#7b8fa4 !important;
  -webkit-text-fill-color:#7b8fa4 !important;
  opacity:1 !important;
}
/* Selectbox/date input displayed values and icons */
[data-baseweb="select"] *, [data-baseweb="input"] svg, [data-baseweb="select"] svg{
  color:#16324a !important;
  fill:currentColor !important;
}
/* Upload area */
[data-testid="stFileUploaderDropzone"]{
  background:#f2f8ff !important;
  border-color:#8fb8e6 !important;
}
[data-testid="stFileUploaderDropzone"] *{
  color:#1d3d5c !important;
  -webkit-text-fill-color:#1d3d5c !important;
}
/* Custom cards */
.taskbox,.card,.totalbox,.scorechart{
  background:#ffffff !important;
  color:#17324d !important;
  border-color:#d8e6f2 !important;
}
.taskbox *,.card *,.totalbox *,.scorechart *,
.wa-hero-title,.wa-section-title,.wa-field-label,.wa-requirements,.desc-en,.scorebar-value,.scorebar-label{
  color:#17324d !important;
}
.wa-hero-subtitle,.wa-instruction,.meta,.desc-zh{
  color:#5f7183 !important;
}
.wa-hero-icon,.wa-section-icon{
  background:#eef6ff !important;
  border-color:#cfe0f1 !important;
}
/* Buttons: keep primary blue and secondary white in all modes */
.stButton > button,.stDownloadButton > button{
  background:#ffffff !important;
  color:#17324d !important;
  -webkit-text-fill-color:#17324d !important;
}
.stButton > button[kind="primary"]{
  background:linear-gradient(135deg,#246bfd,#4f8cff) !important;
  color:#ffffff !important;
  -webkit-text-fill-color:#ffffff !important;
}
.stButton > button *,.stDownloadButton > button *{color:inherit !important;-webkit-text-fill-color:inherit !important}
/* Alerts and tabs */
[data-testid="stAlert"] *{color:inherit !important}
button[data-baseweb="tab"],button[data-baseweb="tab"] *{color:#17324d !important;-webkit-text-fill-color:#17324d !important}
/* Explicitly neutralize system dark-mode repainting */
@media (prefers-color-scheme: dark){
  :root,html,body,.stApp{color-scheme:light !important}
  html,body,.stApp,[data-testid="stAppViewContainer"],[data-testid="stMain"]{
    background:linear-gradient(180deg,#f8fbff 0%,#f5f9ff 45%,#ffffff 100%) !important;
    color:#17324d !important;
  }
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
        st.markdown(
            """<div class="wa-hero">
                <div class="wa-hero-icon">✍️</div>
                <div class="wa-hero-title">Writing Assessment</div>
                <div class="wa-hero-subtitle">Follow the steps below to check your writing. Each student can submit only once.</div>
            </div>""",
            unsafe_allow_html=True
        )

        st.markdown(
            f'''<div class="taskbox"><b>Class:</b> {task.get("class_name","")}<br>
            <b>Task Date:</b> {task.get("task_date","")}<br>
            <b>Task Title:</b> {task["title"]}<br>
            <b>Genre / Writing Type:</b> {task["genre"]}</div>''',
            unsafe_allow_html=True
        )

        if task.get("requirements"):
            st.markdown('<div class="wa-section-title"><span class="wa-section-icon">📝</span><span>Task Requirements</span></div>', unsafe_allow_html=True)
            req_html = html.escape(task["requirements"]).replace("\n", "<br>")
            st.markdown(f'<div class="wa-requirements">{req_html}</div>', unsafe_allow_html=True)

        if task.get("is_closed", False):
            st.warning("This task is closed. Submissions are no longer accepted.")
            st.info("If you think you still need to submit, please contact your teacher.")
            st.stop()

        st.markdown('<div class="wa-section-title"><span class="wa-section-icon">👤</span><span>Step 1. Enter your information</span></div>', unsafe_allow_html=True)
        st.markdown('<div class="wa-instruction">Use your real name. Each student can submit this task only once.</div>', unsafe_allow_html=True)
        info_c1, info_c2, info_c3 = st.columns([1, 2.2, 2.2])
        with info_c1:
            seat_number = st.text_input("🔢 Seat No.", max_chars=2, placeholder="e.g., 1 or 12")
        with info_c2:
            student_id = st.text_input("🪪 Student ID", max_chars=9, placeholder="9 characters")
        with info_c3:
            student_name = st.text_input("👤 Student Name")

        st.markdown('<div class="wa-section-title"><span class="wa-section-icon">📷</span><span>Step 2. Upload your writing</span></div>', unsafe_allow_html=True)
        st.markdown('<div class="wa-instruction">Take a clear photo of your writing and upload it here. Each student can upload only once.</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Choose an image", type=["jpg", "jpeg", "png", "webp"])

        st.markdown('<div class="wa-section-title"><span class="wa-section-icon">✅</span><span>Step 3. Submit your writing</span></div>', unsafe_allow_html=True)
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
                                image_path = upload_submission_image(task["id"], student_id, student_name, uploaded)
                                save_submission(task, seat_number, student_id, student_name, result, image_path)
                                st.session_state["result"] = result
                                st.session_state["show_revision"] = False
                except Exception as e:
                    st.error(f"Assessment could not be completed: {e}")
                    st.info("Your submission was not completed. You may try again. If this message appears again, please show it to your teacher.")

        result = st.session_state.get("result")
        if result:
            st.divider()
            st.markdown('<div class="wa-section-title"><span class="wa-section-icon">🏆</span><span>Your Writing Score</span></div>', unsafe_allow_html=True)
            total = sum(result["scores"][d] for d in RUBRIC)
            st.markdown(
                f'<div class="totalbox"><div class="meta">Total Score</div><div class="totalnum">{total} / 16</div></div>',
                unsafe_allow_html=True
            )

            st.markdown('<div class="wa-section-title"><span class="wa-section-icon">📊</span><span>Score Overview</span></div>', unsafe_allow_html=True)
            st.markdown(
                score_overview_chart(result["scores"]),
                unsafe_allow_html=True
            )

            for dim, levels in RUBRIC.items():
                score = result["scores"][dim]
                if score == 0:
                    desc_en = "The response was not written predominantly in English, so this English writing assessment receives 0 for this dimension."
                    desc_zh = "作答內容主要不是以英文書寫，因此本英文寫作評量此向度為 0 分。"
                else:
                    desc_en, desc_zh = levels[score]
                st.markdown(
                    f'<div class="card"><span class="cardtitle">{html.escape(dim)}</span>'
                    f'<span class="score">{score}/4</span>'
                    f'<div class="desc-en">{html.escape(desc_en)}</div>'
                    f'<div class="desc-zh">{html.escape(desc_zh)}</div></div>',
                    unsafe_allow_html=True
                )

            if not result.get("english_response", True):
                st.info("This response was not written predominantly in English. Please complete the assigned task in English.")
            elif st.button("See Revision Suggestions", use_container_width=True):
                st.session_state["show_revision"] = True

            if result.get("english_response", True) and st.session_state.get("show_revision"):
                st.divider()
                st.markdown('<div class="wa-section-title"><span class="wa-section-icon">💡</span><span>Revision Suggestions</span></div>', unsafe_allow_html=True)

                st.markdown('<div class="wa-section-title"><span class="wa-section-icon">🧩</span><span>Content & Organization Suggestions</span></div>', unsafe_allow_html=True)
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

                st.markdown('<div class="wa-section-title"><span class="wa-section-icon">🔤</span><span>Language Corrections</span></div>', unsafe_allow_html=True)
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

            # If a task was duplicated, load its settings into the main Create Task form.
            pending_duplicate = st.session_state.pop("pending_duplicate_task", None)
            if pending_duplicate:
                st.session_state["create_class"] = pending_duplicate.get("class_name", "")
                try:
                    st.session_state["create_date"] = date.fromisoformat(str(pending_duplicate.get("task_date", "")))
                except Exception:
                    st.session_state["create_date"] = date.today()
                st.session_state["create_title"] = pending_duplicate.get("title", "")
                st.session_state["create_genre"] = pending_duplicate.get("genre", "")
                st.session_state["create_requirements"] = pending_duplicate.get("requirements", "")

            if st.session_state.pop("show_duplicate_notice", False):
                st.success("Task copied to the fields above. You can edit any details before creating the new task.")

            class_name = st.text_input("Class", placeholder="e.g., English Communication A", key="create_class")
            task_date = st.date_input("Task Date", key="create_date")
            title = st.text_input("Task Title", placeholder="e.g., Use of Learning Resources", key="create_title")
            genre = st.text_input("Genre / Writing Type", placeholder="e.g., Email Body", key="create_genre")
            requirements = st.text_area(
                "Task Requirements",
                placeholder="1. Explain how you have benefited from these resources.\n"
                            "2. Give suggestions about what the department can do to encourage students to use them more.",
                height=120,
                key="create_requirements",
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

                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            if st.button(
                                "Duplicate",
                                key=f"duplicate_{code}",
                                use_container_width=True,
                                help="Copy this task into the Create Task fields above.",
                            ):
                                st.session_state["pending_duplicate_task"] = dict(task)
                                st.session_state["show_duplicate_notice"] = True
                                st.session_state.pop("editing_task_code", None)
                                st.session_state.pop("deleting_task_code", None)
                                st.rerun()
                        with c2:
                            if st.button(
                                "Edit Task",
                                key=f"edit_{code}",
                                use_container_width=True,
                                disabled=has_submissions,
                                help="Editing is locked after the first student submission." if has_submissions else None,
                            ):
                                st.session_state["editing_task_code"] = code
                                st.session_state.pop("duplicating_task_code", None)
                        with c3:
                            close_label = "Reopen Task" if is_closed else "Close Task"
                            if st.button(close_label, key=f"close_{code}", use_container_width=True):
                                set_task_closed(code, not is_closed)
                                st.success("Task reopened." if is_closed else "Task closed. Students can no longer submit using this link.")
                                st.rerun()
                        with c4:
                            if st.button(
                                "Delete Task",
                                key=f"delete_{code}",
                                use_container_width=True,
                                disabled=has_submissions,
                                help="Deletion is locked after the first student submission." if has_submissions else None,
                            ):
                                st.session_state["deleting_task_code"] = code
                                st.session_state.pop("duplicating_task_code", None)

                        if has_submissions:
                            st.caption("Edit and Delete are locked because this task already has student submissions. Duplicate and Close/Reopen are still available.")

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
                            st.warning("Are you sure you want to delete this task? The student link will stop working. This action cannot be undone. Existing submission records will not be deleted.")
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
            st.markdown("### Find Student Submission")
            try:
                submissions = list_submissions()
            except Exception as e:
                submissions = []
                st.error(f"Results could not be loaded: {e}")

            if not submissions:
                st.info("No student submissions yet.")
            else:
                # 1) Class
                classes = sorted({str(r.get("class_name", "")).strip() for r in submissions if str(r.get("class_name", "")).strip()})
                selected_class = st.selectbox(
                    "Class",
                    classes,
                    key="results_class",
                )

                class_rows = [r for r in submissions if str(r.get("class_name", "")).strip() == selected_class]

                # 2) Task: use task_id as the real value so duplicated/similarly named tasks remain distinct.
                task_map = {}
                for r in class_rows:
                    task_id_value = str(r.get("task_id", "")).strip()
                    if not task_id_value:
                        # Fallback for very old rows, if any, that do not contain task_id.
                        task_id_value = f"legacy::{r.get('task_date','')}::{r.get('task_title','')}"
                    task_map.setdefault(task_id_value, r)

                task_ids = sorted(
                    task_map.keys(),
                    key=lambda tid: (
                        str(task_map[tid].get("task_date", "")),
                        str(task_map[tid].get("submitted_at", "")),
                    ),
                    reverse=True,
                )

                selected_task_id = st.selectbox(
                    "Task",
                    task_ids,
                    format_func=lambda tid: f"{task_map[tid].get('task_date','')} · {task_map[tid].get('task_title','')}",
                    key="results_task",
                )

                if selected_task_id.startswith("legacy::"):
                    task_date_value = str(task_map[selected_task_id].get("task_date", ""))
                    task_title_value = str(task_map[selected_task_id].get("task_title", ""))
                    task_rows = [
                        r for r in class_rows
                        if str(r.get("task_date", "")) == task_date_value
                        and str(r.get("task_title", "")) == task_title_value
                    ]
                else:
                    task_rows = [r for r in class_rows if str(r.get("task_id", "")) == selected_task_id]

                def seat_sort_key(row):
                    raw = str(row.get("seat_number", "")).strip()
                    try:
                        return (0, int(raw), str(row.get("student_name", "")))
                    except Exception:
                        return (1, 999999, raw, str(row.get("student_name", "")))

                task_rows = sorted(task_rows, key=seat_sort_key)

                # 3) Student: submission ID is the actual selectbox value.
                submission_map = {str(r.get("id")): r for r in task_rows if r.get("id") is not None}
                submission_ids = list(submission_map.keys())

                if not submission_ids:
                    st.info("No student submissions are available for this task.")
                else:
                    selected_submission_id = st.selectbox(
                        "Student",
                        submission_ids,
                        format_func=lambda sid: (
                            f"{submission_map[sid].get('seat_number','')} · "
                            f"{submission_map[sid].get('student_id','')} · "
                            f"{submission_map[sid].get('student_name','')}"
                        ),
                        key="results_student",
                    )

                    selected = submission_map[selected_submission_id]

                    st.divider()
                    st.markdown("### Student Information")
                    st.markdown(
                        f"**Seat No.:** {selected.get('seat_number','')} · "
                        f"**Student ID:** {selected.get('student_id','')} · "
                        f"**Name:** {selected.get('student_name','')}"
                    )
                    st.markdown(
                        f"**Class:** {selected.get('class_name','')} · "
                        f"**Task Date:** {selected.get('task_date','')} · "
                        f"**Task Title:** {selected.get('task_title','')}"
                    )
                    st.caption(f"Submitted: {selected.get('submitted_at','')}")

                    st.markdown("### Assessment Results")
                    score_cols = st.columns(5)
                    score_items = [
                        ("Content", selected.get("content_score", ""), 4),
                        ("Organization", selected.get("organization_score", ""), 4),
                        ("Language", selected.get("language_score", ""), 4),
                        ("Genre", selected.get("genre_score", ""), 4),
                        ("Total", selected.get("total", ""), 16),
                    ]
                    for col, (label, value, maximum) in zip(score_cols, score_items):
                        with col:
                            st.caption(label)
                            st.markdown(f"## {value}/{maximum}")

                    st.divider()
                    st.markdown("### Original Submission & Double Check")
                    image_path = selected.get("image_path", "")
                    image_bytes = download_submission_image(image_path)
                    if image_bytes:
                        st.markdown("#### Original Submission")
                        st.image(image_bytes, use_container_width=True)
                    else:
                        st.info("No stored image is available for this submission. Older submissions created before image storage was enabled will not have an image.")

                    transcription = selected.get("transcription", "")
                    if transcription:
                        st.markdown("#### AI Transcription")
                        st.text_area(
                            "Transcribed text",
                            value=transcription,
                            height=240,
                            disabled=True,
                            key=f"transcription_{selected_submission_id}",
                        )
                    else:
                        st.info("No AI transcription is stored for this submission.")

                    st.caption(
                        "Revision Suggestions and detailed AI feedback are not currently stored in the submissions table, so they cannot be reopened here for older or completed submissions."
                    )

                    st.divider()
                    st.markdown("### Task Results")
                    table_rows = [{
                        "Seat No.": r.get("seat_number", ""),
                        "Student ID": r.get("student_id", ""),
                        "Student Name": r.get("student_name", ""),
                        "Content": r.get("content_score", ""),
                        "Organization": r.get("organization_score", ""),
                        "Language": r.get("language_score", ""),
                        "Genre": r.get("genre_score", ""),
                        "Total": r.get("total", ""),
                        "Submitted": r.get("submitted_at", ""),
                    } for r in task_rows]
                    st.dataframe(table_rows, use_container_width=True, hide_index=True)

                    st.download_button(
                        "Download This Task Results (Excel)",
                        data=build_excel(task_rows),
                        file_name="writing_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
