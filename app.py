import os
import json
import base64
import html
import hashlib
import time
import re
import difflib
from datetime import date
from io import BytesIO

from PIL import Image, ImageOps
import streamlit as st
from openai import OpenAI
from openpyxl import Workbook
import requests

APP_NAME = "Writing Assessment"

# Authentic teacher stamp provided by the teacher; no extra file deployment needed.
VIVIAN_STAMP_REFERENCE_B64 = "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"

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
            "Meaningfully addresses at least part of the assigned task, but one or more important requirements remain missing or insufficiently developed. There is some substantive task-related information beyond isolated statements or repeated claims.",
            "對指定任務至少有部分具體回應，但仍缺少一項以上重要要求，或發展不充分；內容須有一定實質資訊，不能只有零碎陳述或重複主張。"
        ),
        1: (
            "Provides only isolated, skeletal, or extremely limited task-related statements, or is largely off-topic. Most essential task requirements are unmet, and there is too little meaningful development to fulfill the assigned purpose.",
            "僅提供零碎、極簡或非常有限的相關陳述，或大多偏題；多數必要任務要求未完成，缺乏實質發展，無法達成指定目的。"
        )
    },
    "Organization & Coherence": {
        4: (
            "The required task content is well organized at both paragraph and whole-text levels. Ideas progress logically, transitions are effective, and the assigned message or purpose is easy to follow.",
            "任務所要求的內容在段落與全文層次皆組織良好。想法發展有邏輯，轉承有效，讀者能輕鬆理解指定的訊息或目的。"
        ),
        3: (
            "The ideas actually presented are meaningfully developed and organized into a clear progression appropriate to the assigned task. Connections show how distinct ideas relate, with only minor lapses in coherence. A short repetitive statement with a warning is not sufficient evidence of this level.",
            "實際呈現的想法有適當發展，並依指定任務形成清楚的論述進展；不同想法之間的關係明確，僅有少數連貫性問題。簡短重複的陳述加上一句提醒，不足以證明達到此程度。"
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
            "Generally achieves the assigned communicative purpose with sufficiently informative content for the intended reader, while following the expected format, tone, and conventions of the specified genre. Minor inconsistencies do not undermine communication.",
            "能以對預期讀者而言足夠且有用的資訊，大致達成指定溝通目的，並符合文類的格式、語氣與慣例；少數不一致之處不妨礙溝通。"
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


def response_is_task_requirements_copy(transcription, requirements):
    """Return True only when the student's response is essentially copied from Task Requirements.

    This guardrail is intentionally conservative. Reusing a few necessary task words or phrases
    should NOT trigger a zero. The zero-score rule applies only when most of the student's actual
    response is copied from the teacher-provided Task Requirements with little original writing.
    """
    def _tokens(value):
        return re.findall(r"[a-z0-9]+", (value or "").lower())

    response_tokens = _tokens(transcription)
    requirement_tokens = _tokens(requirements)

    if len(response_tokens) < 5 or len(requirement_tokens) < 5:
        return False

    response_norm = " ".join(response_tokens)
    requirement_norm = " ".join(requirement_tokens)

    # Exact or near-exact copying of a requirement sentence / block.
    if response_norm in requirement_norm:
        return True

    matcher = difflib.SequenceMatcher(
        None, requirement_tokens, response_tokens, autojunk=False
    )
    matched_words = sum(block.size for block in matcher.get_matching_blocks())
    copied_fraction = matched_words / max(1, len(response_tokens))

    # Treat as "copy only" when at least 80% of the student's response is traceable
    # to the Task Requirements and there is very little additional original writing.
    little_original_writing = len(response_tokens) <= len(requirement_tokens) + 12
    return copied_fraction >= 0.80 and little_original_writing


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

DYNAMIC REQUIREMENT AUDIT (MANDATORY BEFORE CONTENT SCORING)
- Read the CURRENT Task Requirements and identify each distinct essential content obligation, including any requested number of distinct ideas, explanations, reasons, examples, recommendations, or comparisons ONLY when this task actually asks for them.
- For each obligation, check whether the student's own response fulfills it meaningfully, partially, or not at all. Repeated descriptions of the same idea do not satisfy a request for distinct items.
- Distinguish essential obligations from optional illustrations and minor supporting details. Do not assume every task needs examples, reasons, advice, or any fixed number of points.
- Check the quality and sufficiency of development as well as coverage. Merely naming an item is not necessarily explaining it.
- If an essential content obligation is missing or only superficially addressed, Content normally cannot exceed 2. However, score 2 is NOT an automatic floor for mentioning the topic: a skeletal response that barely addresses the task may merit 1. A minor omission or uneven but meaningful development may still justify 3 if the main task is fulfilled. Do not mechanically subtract points per missing item.
- Before finalizing Content 3 or 4, verify that the requirement audit supports successful completion of the task. If it does not, reconsider the score.

WORD-COUNT POLICY (ALL TASKS)
- When the teacher specifies a word-count range, its LOWER bound is the required minimum. The upper bound is guidance, not a penalty threshold.
- NEVER deduct points merely because the student exceeds the stated maximum; sufficiently developed longer writing may receive 4.
- When the response falls below the minimum, reflect the shortfall in Content & Task Fulfillment, taking into account its magnitude and whether the required ideas have been developed. If it is substantially below the minimum AND consists only of skeletal statements with most essential requirements missing, Content should ordinarily be 1 rather than 2. A slight shortfall is different from a substantial one; do not apply an automatic fixed-point deduction.
- Word-count shortfall alone must not automatically reduce Organization, Language Use, or Genre. Judge those dimensions independently.
- If no minimum word count is specified, do not invent one.

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

TASK-REQUIREMENTS-COPY OVERRIDE
- If the student's submitted response merely copies or substantially reproduces the teacher-provided Task Requirements instead of producing an original response to the task, set ALL FOUR dimension scores to 0.
- Apply this 0-score rule only when the submission is essentially the Task Requirements themselves with little or no original student writing.
- Do NOT apply this rule merely because the student appropriately reuses necessary topic words, names, phrases, or wording from the task.
- In the JSON output, set "task_requirements_copy_only" to true only when this override applies.

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
- A score of 2 is appropriate for partial, limited, inconsistent, weak, or noticeably flawed performance.
- A score of 3 is NOT simply "the task was completed" or "the meaning can be understood." A 3 requires clearly adequate, generally successful performance in THAT specific dimension.
- A score of 4 is NOT the default for a complete response. It represents clearly stronger performance than score 3 and requires the score-4 descriptor to be fully and consistently demonstrated.
- Completing every listed requirement does NOT automatically justify score 3 or 4 in Organization, Language Use, or Genre.
- If a response is mainly simple listing, thin explanation, basic development, predictable structure, limited language range, or has repeated noticeable errors, the relevant dimension should normally remain at 2 or 3 depending on the descriptor match.
- Do not raise Organization or Genre to 3 merely because the response is on topic, has a beginning and ending, or includes a few expected surface features.
- Do not raise Language Use to 3 merely because the overall meaning is understandable.
- When deciding between 2 and 3, choose 3 only when the evidence shows adequate control of that specific dimension across most of the response. If performance is frequent, repeated, or noticeably weak in that dimension, choose 2.
- When deciding between 3 and 4, choose 4 only when there is clear positive evidence of fuller development, stronger control, effective organization, and/or broader, more precise task-appropriate language as required by that dimension.
- Do not automatically give the same score to all four dimensions. Judge each dimension separately, but always within the same task context.

OFFICIAL RUBRIC BOUNDARY DECISIONS
The OFFICIAL RUBRIC descriptors above remain the scoring authority. The rules below do NOT replace or rewrite them. Use these rules only to decide between two adjacent scores when the response sits near a boundary.

CONTENT & TASK FULFILLMENT
- 1 vs 2: Choose 1 when the response offers only a few isolated claims, repeated versions of one point, or generic advice, and most essential requirements remain unmet. Merely mentioning the topic, identifying one item without developing it, or adding a brief warning does NOT establish score 2. Choose 2 when the student meaningfully develops at least one essential component with substantive, relevant explanation or detail, even if other components are missing. Consider a severe shortfall below the specified minimum as additional evidence of skeletal development, not as an automatic numeric cutoff. A short but genuinely developed response must be judged on its actual substance.
- 2 vs 3: Choose 3 only when the MAIN task requirements are addressed and there is enough relevant development to complete the task successfully. If important required content is missing, weakly developed, repetitive, or uneven, choose 2.
- 3 vs 4: Choose 4 only when ALL important requirements are fully addressed with clearly stronger, specific, sufficiently developed support. Adequate completion alone is 3.

ORGANIZATION & COHERENCE
- Assess the organization of ideas that the student actually developed. Do not pretend missing task ideas exist, and do not mechanically copy the Content score.
- A logical order of two or three short sentences, a heading, or a concluding warning is only basic sequencing; it does not establish adequately developed coherence.
- Repeated restatements of one idea are not meaningful progression. Look for distinct developed points and intelligible relationships among them, relative to the assigned task and response length.
- A short response may still earn 3 if it demonstrates genuine, clear development and relationships; do not impose a fixed length or paragraph-count requirement. For an extremely skeletal response, do not mistake a sequence of isolated sentences for actual organization.
- 1 vs 2: Choose 2 only when some recognizable organization of task-related ideas is present. If ideas are fragmented, poorly sequenced, or insufficiently connected for the assigned purpose, choose 1.
- 2 vs 3: Choose 3 only when the overall sequence and connections are generally clear and appropriate, with only MINOR lapses. A few connectors such as "First" or "Also" do not by themselves justify 3. If sequencing, paragraphing, or transitions are noticeably inconsistent or ideas are loosely connected, choose 2.
- 3 vs 4: Choose 4 only when organization is consistently strong at both paragraph and whole-text levels, idea progression is logical, and transitions are effective throughout. Generally clear organization is 3.

LANGUAGE USE
- 1 vs 2: Choose 2 when limited or inaccurate language still communicates enough of the task that the intended meaning is generally recoverable. Choose 1 when language is too limited or errors are so frequent/serious that effective communication of the required content breaks down.
- 2 vs 3: Choose 3 only when vocabulary and sentence structures are sufficient across MOST of the response and errors are not frequent enough to noticeably reduce clarity. If basic grammar, wording, spelling, punctuation, capitalization, or sentence-structure errors recur across many sentences, choose 2 even if the overall message can still be understood.
- 3 vs 4: Choose 4 only when language shows clearly stronger range, precision, accuracy, and appropriacy, with only minor errors. Sufficient but ordinary control is 3.

GENRE & PROFESSIONAL APPROPRIACY
- Evaluate whether the text gives the intended audience enough relevant, usable communication to substantially achieve the assigned genre purpose. A title, warning tone, or appropriate salutation alone is not sufficient.
- Distinguish communicative success from the Content checklist: do not mechanically copy the Content score, but a merely skeletal message that does not substantially perform the genre's purpose normally merits 2 rather than 3.
- A brief response can still earn 3 when its communicative purpose is substantially achieved with appropriate audience awareness and conventions.
- 1 vs 2: Choose 2 when the student shows partial awareness of the assigned communicative purpose and genre, even if format, tone, organization, or audience awareness is inconsistent. Choose 1 when awareness of the assigned genre/purpose is very limited or the response is largely inappropriate for the expected audience, format, tone, or conventions.
- 2 vs 3: Choose 3 when the response generally carries out the assigned communicative purpose and generally follows the expected organization, format, and tone, with only minor inconsistencies. Do NOT lower Genre merely because Language Use contains grammar errors; avoid double-penalizing language errors unless they directly prevent the genre purpose or audience appropriacy.
- 3 vs 4: Choose 4 only when genre control is consistent and the writing is clearly well suited to the intended audience and context throughout. Generally appropriate genre performance is 3.

ANTI-INFLATION CHECK: CHECK BEFORE ANY SCORE OF 3
Before assigning 3 in any dimension, ask whether the response demonstrates ADEQUATE CONTROL of that dimension, not merely partial success.

For Content = 3:
- The main task requirements are addressed clearly enough to fulfill the task.
- Ideas may still be simple or only moderately developed.
- A response can earn Content = 3 even when its Organization, Language Use, or Genre scores are only 2.

For Organization = 3:
- Ideas must show genuine progression and meaningful connections across the response, beyond repetition or a simple problem-then-warning sequence.
- Basic connectors alone (for example, "First" or "Also") do not justify 3 if ideas are mostly listed, loosely connected, or weakly developed.
- If coherence is basic, inconsistent, or frequently awkward, Organization should normally be 2.

For Language Use = 3:
- Grammar, vocabulary, spelling, punctuation, capitalization, and sentence structure must be adequately controlled across most of the response.
- Some errors are acceptable, but errors should not be frequent across many sentences.
- Repeated basic errors such as subject-verb agreement, missing articles/prepositions, malformed clauses, fragments, run-ons, incorrect word forms, or unnatural sentence patterns normally keep Language Use at 2 or below, even when the meaning remains understandable.

For Genre & Professional Appropriacy = 3:
- The response must substantially achieve the expected communicative purpose for its intended readers, with sufficient useful information, appropriate tone, audience awareness, and genre conventions.
- A greeting, closing, title, or other surface feature alone does not justify 3.
- If tone, phrasing, format, audience awareness, or genre conventions are noticeably inconsistent or weak, Genre should normally be 2.

CALIBRATION PRINCIPLE
A poor but on-task response may reasonably receive a pattern such as Content = 3 while Organization = 2, Language Use = 2, and Genre = 2. Equally, a partially completed task with repetitive, skeletal development should not earn Organization or Genre = 3 merely because its few sentences are understandable or its warning sounds appropriate. Do not let successful task completion inflate the other dimensions.

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
- Check every explicit Task Requirement, using the mandatory dynamic requirement audit and word-count policy above.
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
- Did Content reflect all essential requirements and any minimum-word-count shortfall without penalizing writing above the maximum?

IMPORTANT TASK RULES
- No half points.
- Normal rubric scores are 1, 2, 3, or 4. Score 0 is reserved only for: (a) the language-of-response override, or (b) the task-requirements-copy override.
- Evaluate only the writing the student is required to produce.
- Some assignments may already provide fixed genre elements outside the student's response, such as a subject line, greeting, opening, closing, or signature.
- Do NOT penalize a student for omitting any element that is not explicitly required in Task Requirements.
- If the Genre / Writing Type is "Email Body", evaluate only the body paragraphs for appropriate purpose, organization, tone, audience awareness, and professional/academic appropriacy. Do not require a greeting, closing, or signature.
- If the image is not readable enough, do not guess.

LANGUAGE CORRECTIONS — SENTENCE BY SENTENCE
Check the student's transcription sentence by sentence for ALL clear, genuine errors in these categories:
- Grammar
- Spelling
- Punctuation
- Capitalization
- Sentence Structure

Use Sentence Structure for problems such as sentence fragments, run-on sentences, comma splices, incorrect word order, incomplete or malformed clauses, or sentence patterns that make the intended meaning unclear. Do not use Sentence Structure for mere style preferences.

IMPORTANT OUTPUT RULES:
- Create ONE correction item for each sentence that contains one or more genuine errors.
- Do NOT create separate correction items for multiple errors in the same sentence.
- "original_sentence" must quote the student's complete original sentence exactly as it appears in the transcription.
- "corrected_sentence" must give one complete corrected version of that same sentence while preserving the student's intended meaning.
- "explanation_zh" must explain ALL errors in that sentence in concise Traditional Chinese.
- Separate each error explanation with a Chinese semicolon "；".
- Begin each explanation item with its category, for example: "Grammar：..."；"Sentence Structure：...".
- If the sentence has only one error, give only one explanation item.
- Do not rewrite sentences that are already acceptable.
- Do not invent wording that does not appear in the student's transcription.
- Do not rewrite the full composition.

CONTENT & ORGANIZATION REVISION
Provide 1 to 3 useful revision suggestions about CONTENT and ORGANIZATION only.
For each suggestion, provide:
- a short A2-B1 English suggestion,
- a clear Traditional Chinese translation.
Do not provide a model essay.

Return VALID JSON ONLY.
The values "<score 0-4>" below are placeholders. Use 0 only when the language-of-response override or task-requirements-copy override applies; otherwise use 1-4.

{{
  "image_readable": true,
  "transcription": "faithful transcription",
  "english_response": true,
  "task_requirements_copy_only": false,
  "scores": {{
    "Content & Task Fulfillment": "<score 0-4>",
    "Organization & Coherence": "<score 0-4>",
    "Language Use": "<score 0-4>",
    "Genre & Professional Appropriacy": "<score 0-4>"
  }},
  "corrections": [
    {{
      "original_sentence": "complete original sentence exactly from the transcription",
      "corrected_sentence": "complete corrected sentence",
      "explanation_zh": "Grammar：...；Spelling：...；Punctuation：...；Capitalization：...；Sentence Structure：..."
    }}
  ],
  "content_organization_suggestions": [
    {{
      "en": "short A2-B1 English suggestion",
      "zh": "清楚的繁體中文翻譯"
    }}
  ]
}}"""


def verify_vivian_stamp(uploaded_file):
    """Fail closed: only confirmed matching stamps can proceed to assessment."""
    api_key = secret_or_env("OPENAI_API_KEY")
    model = secret_or_env("OPENAI_MODEL", "gpt-5.6-luna")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not configured.")
    client = OpenAI(api_key=api_key)
    # The first image is the teacher-provided reference; the second is the student's entire photo.
    reference_url = "data:image/png;base64," + VIVIAN_STAMP_REFERENCE_B64
    response = client.responses.create(
        model=model,
        input=[{"role": "user", "content": [
            {"type": "input_text", "text": (
                "You are checking an image-submission prerequisite, NOT grading writing. "
                "Image 1 is the authentic teacher stamp reference: blue ink, Chinese name and 'Vivian' with a small cartoon. "
                "Image 2 is a student's photographed worksheet. Determine whether a clearly visible stamp matching "
                "the reference is physically present in the bottom-right region of the worksheet in image 2. "
                "Allow ordinary rotation, scale, faint ink, and minor blur, but do not accept arbitrary stamps, "
                "printed words saying Vivian, or a stamp elsewhere on the page. "
                "If the stamp is missing, cropped out, indistinct, or you are uncertain, return false. "
                "Reply ONLY with JSON: {\"stamp_verified\": true} or {\"stamp_verified\": false}."
            )},
            {"type": "input_image", "image_url": reference_url},
            {"type": "input_image", "image_url": image_to_data_url(uploaded_file)},
        ]}],
        timeout=120,
    )
    raw = response.output_text.strip()
    if raw.startswith("```"):
        raw = raw.strip("`").strip()
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
    data = json.loads(raw)
    return data.get("stamp_verified") is True


def minimum_required_words(requirements):
    """Extract an explicit minimum from common English word-count instructions."""
    text = str(requirements or "")
    ranges = re.findall(r"(?<!\d)(\d{1,4})\s*[-–—~]\s*(\d{1,4})\s*(?:words?\b|word\b)", text, re.I)
    if ranges:
        return min(int(a) for a, b in ranges if int(a) <= int(b))
    minimum = re.search(r"(?:at least|minimum(?: of)?|no fewer than)\s*(\d{1,4})\s*words?\b", text, re.I)
    return int(minimum.group(1)) if minimum else None


def audit_skeletal_content(client, model, task, transcription):
    """Independent task-specific 1-vs-2 audit; only used for severe underlength cases."""
    prompt = (
        "Independently audit CONTENT & TASK FULFILLMENT only, for the student's ACTUAL writing. "
        "Interpret the teacher's task-specific essential obligations, not a generic checklist. "
        "A score of 1 means only skeletal, isolated or repetitive statements, with most "
        "essential obligations unmet and no meaningful development. A score of 2 requires "
        "at least one essential component to be meaningfully developed with substantive, "
        "task-relevant explanation or detail. Mentioning a topic or generic warning alone "
        "does not suffice. Do not require examples unless explicitly asked. "
        "This response is substantially shorter than the teacher's stated minimum; "
        "consider this as evidence but NOT as an automatic score rule. "
        "Also independently evaluate ORGANIZATION & COHERENCE: return organization_level 1 if "
        "the response is just disconnected, repeated, or skeletal statements with no "
        "meaningful development or relationships among ideas; return 2 if there is "
        "recognizable basic organization of at least some distinct ideas. Do not "
        "automatically equate Organization with Content: a developed but off-topic "
        "response can still be organized. "
        "Return ONLY JSON with integer content_level (1 or 2) and "
        "organization_level (1 or 2).\n\n"
        f"Task Title: {task.get('title', '')}\nGenre: {task.get('genre', '')}\n"
        f"Requirements: {task.get('requirements', '')}\n"
        f"Student response: {transcription}"
    )
    response = client.responses.create(model=model, input=[{"role": "user", "content": prompt}], timeout=120)
    raw = response.output_text.strip()
    if raw.startswith("```"):
        raw = raw.strip("`").strip()
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()
    audit = json.loads(raw)
    content_level = int(audit["content_level"])
    organization_level = int(audit["organization_level"])
    return (content_level, organization_level) if content_level in (1, 2) and organization_level in (1, 2) else None


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
    non_english_zero = (model_says_english is False) or response_requires_zero(transcription)
    task_copy_zero = bool(result.get("task_requirements_copy_only", False)) or response_is_task_requirements_copy(
        transcription, task.get("requirements", "")
    )
    force_zero = non_english_zero or task_copy_zero

    if force_zero:
        result["zero_reason"] = "non_english" if non_english_zero else "task_requirements_copy"
        result["english_response"] = not non_english_zero
        result["task_requirements_copy_only"] = bool(task_copy_zero)
        result["scores"] = {dim: 0 for dim in RUBRIC}
        result["corrections"] = []
        result["content_organization_suggestions"] = []
    else:
        result["zero_reason"] = ""
        result["english_response"] = True
        result["task_requirements_copy_only"] = False
        for dim in RUBRIC:
            score = int(result["scores"][dim])
            if score not in (1, 2, 3, 4):
                raise ValueError(f"Invalid score for {dim}")
            result["scores"][dim] = score

        # A separate model pass resolves the known 1-vs-2 inflation in extremely
        # short submissions. It is task-specific, not a hard-coded word-count penalty.
        # Run only for a severe shortfall and an initial Content score of 2.
        min_words = minimum_required_words(task.get("requirements", ""))
        words_written = len(re.findall(r"\b[A-Za-z]+(?:['’][A-Za-z]+)?\b", transcription))
        if (min_words and min_words > 0 and words_written < 0.6 * min_words
                and result["scores"]["Content & Task Fulfillment"] == 2):
            try:
                audited_levels = audit_skeletal_content(client, model, task, transcription)
                if audited_levels:
                    content_level, organization_level = audited_levels
                    if content_level == 1:
                        result["scores"]["Content & Task Fulfillment"] = 1
                    # Only adjust Organization if the audit independently finds
                    # skeletal sequencing, not merely because Content is low.
                    if organization_level == 1:
                        result["scores"]["Organization & Coherence"] = 1
            except Exception:
                # Never block a valid submission solely because the secondary audit failed.
                pass

        # Verify every original sentence against the transcription.
        # Any correction for wording the student did not write is discarded.
        def _norm_match(s):
            return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()

        transcription_norm = _norm_match(transcription)
        verified = []
        seen = set()
        for item in result.get("corrections", []) or []:
            if not isinstance(item, dict):
                continue
            original = str(item.get("original_sentence", item.get("original", ""))).strip()
            corrected = str(item.get("corrected_sentence", item.get("correction", ""))).strip()
            explanation = str(item.get("explanation_zh", item.get("explanation", ""))).strip()
            original_norm = _norm_match(original)
            if not original_norm or original_norm not in transcription_norm:
                continue
            if original_norm in seen:
                continue
            seen.add(original_norm)
            verified.append({
                "original_sentence": original,
                "corrected_sentence": corrected,
                "explanation_zh": explanation,
            })
        result["corrections"] = verified
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
/* Uploaded file chip/icon: make the file icon light instead of the default dark tile */
[data-testid="stFileUploaderFile"]{
  background:#f7fbff !important;
  border:1px solid #cfe0f1 !important;
  border-radius:12px !important;
}
[data-testid="stFileUploaderFile"] svg{
  color:#ffffff !important;
  fill:#ffffff !important;
  stroke:#ffffff !important;
}
[data-testid="stFileUploaderFile"] > div:first-child,
[data-testid="stFileUploaderFile"] [data-testid="stFileUploaderFileIcon"]{
  background:#6f9fd0 !important;
  color:#ffffff !important;
  border-radius:9px !important;
}
[data-testid="stFileUploaderFile"] > div:first-child *{
  color:#ffffff !important;
  -webkit-text-fill-color:#ffffff !important;
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

/* Strong fallback for Streamlit's current uploaded-file icon markup.
   The internal DOM can vary by Streamlit version, so target icon-bearing
   elements anywhere inside the file uploader instead of relying on one testid. */
[data-testid="stFileUploader"] [data-testid^="stIconMaterial"],
[data-testid="stFileUploader"] svg{
  color:#4f8cff !important;
  fill:#4f8cff !important;
  stroke:#4f8cff !important;
}
[data-testid="stFileUploader"] div:has(> svg),
[data-testid="stFileUploader"] div:has(> [data-testid^="stIconMaterial"]),
[data-testid="stFileUploader"] span:has(> svg){
  background:#eef6ff !important;
  color:#4f8cff !important;
  border-radius:8px !important;
}
[data-testid="stFileUploader"] [data-testid^="stIconMaterial"]{
  background:#eef6ff !important;
  border-radius:8px !important;
  padding:2px !important;
}


/* Hide Streamlit's built-in uploaded-file icon. Its dark tile is rendered
   inconsistently across Streamlit/browser versions, so we use our own
   upload-complete confirmation below the uploader instead. */
[data-testid="stFileUploader"] svg,
[data-testid="stFileUploader"] [data-testid^="stIconMaterial"]{
  display:none !important;
}


/* Replace Streamlit's blank/dark uploaded-file tile with a real document icon. */
[data-testid="stFileUploaderFile"] > div:first-child{
  display:none !important;
}
[data-testid="stFileUploaderFile"]::before{
  content:"📄";
  display:inline-flex;
  align-items:center;
  justify-content:center;
  width:32px;
  height:32px;
  margin-right:8px;
  border-radius:8px;
  background:#eef6ff;
  font-size:20px;
  line-height:1;
  flex:0 0 auto;
}


/* Custom uploaded-file card used after upload is complete. */
.wa-upload-card{
  display:flex;
  align-items:center;
  gap:12px;
  background:#f7fbff;
  border:1px solid #cfe0f1;
  border-radius:14px;
  padding:12px 14px;
  margin:.45rem 0 .55rem;
}
.wa-upload-doc{
  width:38px;
  height:38px;
  flex:0 0 38px;
  display:flex;
  align-items:center;
  justify-content:center;
  border-radius:10px;
  background:#eaf3ff;
}
.wa-upload-doc svg{
  width:22px;
  height:22px;
  display:block;
}
.wa-upload-name{
  font-weight:800;
  color:#17324d;
  line-height:1.35;
  word-break:break-all;
}
.wa-upload-size{
  color:#5f7183;
  font-size:.9rem;
  margin-top:2px;
}

</style>
''', unsafe_allow_html=True)


task_id = st.query_params.get("task")

# ---------------------------
# STUDENT VIEW
# ---------------------------
if task_id:

    st.markdown(
        """
        <style>
        :root{
          --wa-student-ink:#10233f;
          --wa-student-muted:#5e6f8b;
          --wa-student-line:#d8e0f1;
          --wa-student-glass:rgba(255,255,255,.88);
          --wa-student-shadow:0 12px 30px rgba(58,75,120,.10);
          --wa-student-shadow-raised:
              0 7px 14px rgba(64,75,110,.12),
              inset 0 1px 0 rgba(255,255,255,.95);
        }

        html, body, .stApp,
        [data-testid="stAppViewContainer"],
        [data-testid="stMain"]{
          color-scheme:light !important;
          background:
            radial-gradient(circle at 92% 5%, rgba(193,168,255,.24) 0, rgba(193,168,255,0) 26%),
            radial-gradient(circle at 8% 82%, rgba(143,229,215,.20) 0, rgba(143,229,215,0) 28%),
            linear-gradient(180deg,#edf2ff 0%,#f5f8ff 47%,#eefbf8 100%) !important;
          color:var(--wa-student-ink) !important;
        }

        .block-container{
          max-width:940px !important;
          padding-top:1.35rem !important;
          padding-bottom:5rem !important;
        }

        .wa-hero{
          padding:.15rem 0 1.15rem !important;
          text-align:center !important;
        }
        .wa-hero-icon{
          width:72px !important;
          height:72px !important;
          margin:0 auto .7rem !important;
          border-radius:20px !important;
          background:linear-gradient(145deg,#8a68ff 0%,#6840ee 100%) !important;
          border:1px solid rgba(255,255,255,.72) !important;
          box-shadow:
            0 12px 24px rgba(104,64,238,.23),
            inset 0 1px 0 rgba(255,255,255,.35) !important;
          color:#fff !important;
          font-size:2rem !important;
          display:flex !important;
          align-items:center !important;
          justify-content:center !important;
        }
        .wa-hero-icon svg{width:36px !important;height:36px !important;display:block !important;}
        .wa-hero-title{
          font-size:2.15rem !important;
          font-weight:850 !important;
          letter-spacing:-.025em !important;
          color:var(--wa-student-ink) !important;
          line-height:1.15 !important;
        }
        .wa-hero-subtitle{
          max-width:720px;
          margin:.48rem auto 0 !important;
          color:var(--wa-student-muted) !important;
          font-size:1.05rem !important;
          font-weight:500 !important;
          line-height:1.6 !important;
        }

        .taskbox,
        .wa-requirements-card,
        .card,
        .totalbox,
        .scorechart,
        .wa-upload-card{
          background:var(--wa-student-glass) !important;
          border:1px solid rgba(255,255,255,.92) !important;
          border-radius:24px !important;
          box-shadow:var(--wa-student-shadow) !important;
          backdrop-filter:blur(10px);
          -webkit-backdrop-filter:blur(10px);
        }

        .taskbox{
          padding:1.35rem 1.55rem !important;
          margin:.85rem 0 1.45rem !important;
          border-left:none !important;
          line-height:1.9 !important;
          color:var(--wa-student-ink) !important;
        }
        .taskbox b{color:#24354f !important;font-weight:800 !important;}

        .wa-requirements-card{
          padding:1.25rem 1.4rem !important;
          margin:.35rem 0 1.5rem !important;
        }
        .wa-requirements{
          margin:0 !important;
          color:#4b5f7c !important;
          font-size:1.05rem !important;
          line-height:1.85 !important;
          font-weight:500 !important;
        }

        .wa-section-title{
          margin:1.55rem 0 .65rem !important;
          color:var(--wa-student-ink) !important;
          font-size:1.28rem !important;
          font-weight:850 !important;
          gap:.65rem !important;
        }
        .wa-section-icon{
          width:36px !important;
          height:36px !important;
          border-radius:11px !important;
          background:linear-gradient(145deg,#f6f3ff,#eeeaff) !important;
          border:1px solid #e2dcff !important;
          box-shadow:0 4px 10px rgba(86,64,180,.07) !important;
          color:#6045d9 !important;
        }
        .wa-instruction{
          color:#718099 !important;
          font-size:1rem !important;
          font-weight:500 !important;
          line-height:1.65 !important;
          margin:.15rem 0 .75rem !important;
        }

        .stTextInput label,
        .stTextArea label,
        .stSelectbox label,
        .stFileUploader label,
        [data-testid="stWidgetLabel"] *{
          color:#263a56 !important;
          -webkit-text-fill-color:#263a56 !important;
          font-weight:750 !important;
          opacity:1 !important;
        }

        div[data-baseweb="input"] > div,
        div[data-baseweb="textarea"] > div,
        div[data-baseweb="select"] > div,
        [data-baseweb="base-input"]{
          background:rgba(255,255,255,.84) !important;
          border:1px solid #d7deec !important;
          border-radius:12px !important;
          box-shadow:
            inset 0 1px 2px rgba(48,63,94,.05),
            0 2px 6px rgba(54,72,110,.03) !important;
        }
        div[data-baseweb="input"] input,
        div[data-baseweb="textarea"] textarea,
        div[data-baseweb="select"] input,
        input, textarea{
          color:#223550 !important;
          -webkit-text-fill-color:#223550 !important;
          background:transparent !important;
          font-weight:600 !important;
        }
        input::placeholder, textarea::placeholder{
          color:#99a7bd !important;
          -webkit-text-fill-color:#99a7bd !important;
          opacity:1 !important;
        }

        [data-testid="stFileUploaderDropzone"]{
          background:rgba(255,255,255,.74) !important;
          border:1.5px dashed #b9c7df !important;
          border-radius:18px !important;
          box-shadow:inset 0 1px 0 rgba(255,255,255,.8) !important;
        }
        [data-testid="stFileUploaderDropzone"] *{
          color:#425571 !important;
          -webkit-text-fill-color:#425571 !important;
        }

        .wa-upload-card{
          padding:12px 14px !important;
          border-radius:16px !important;
          margin:.5rem 0 .65rem !important;
        }
        .wa-upload-doc{
          background:linear-gradient(145deg,#f0ecff,#e8f4ff) !important;
          border:1px solid #ddd9fa !important;
          box-shadow:0 4px 10px rgba(83,67,165,.08) !important;
        }
        .wa-upload-name{color:#243651 !important;font-weight:800 !important;}
        .wa-upload-size{color:#73839a !important;}

        .stButton > button,
        .stDownloadButton > button{
          min-height:3.05rem !important;
          border-radius:13px !important;
          background:linear-gradient(180deg,#ffffff 0%,#f3f5fb 100%) !important;
          color:#20334e !important;
          -webkit-text-fill-color:#20334e !important;
          border:1px solid #cfd8e8 !important;
          font-weight:800 !important;
          box-shadow:var(--wa-student-shadow-raised) !important;
          transition:transform .08s ease, box-shadow .08s ease, background .08s ease !important;
        }
        .stButton > button[kind="primary"]{
          background:linear-gradient(180deg,#fbfaff 0%,#eeeaff 100%) !important;
          color:#2c2850 !important;
          -webkit-text-fill-color:#2c2850 !important;
          border:1px solid #cbc2f5 !important;
          box-shadow:
            0 7px 14px rgba(91,72,176,.13),
            inset 0 1px 0 rgba(255,255,255,.95) !important;
        }
        .stButton > button:hover,
        .stDownloadButton > button:hover{
          transform:translateY(-1px) !important;
          background:linear-gradient(180deg,#ffffff 0%,#eef1f8 100%) !important;
          border-color:#becbe0 !important;
        }
        .stButton > button:active,
        .stDownloadButton > button:active{
          transform:translateY(1px) !important;
          box-shadow:
            0 2px 5px rgba(55,70,100,.10),
            inset 0 2px 5px rgba(65,75,110,.08) !important;
        }
        .stButton > button:disabled{
          background:#f1f3f7 !important;
          color:#9aa6b8 !important;
          -webkit-text-fill-color:#9aa6b8 !important;
          border-color:#dde3ec !important;
          box-shadow:none !important;
          opacity:1 !important;
        }
        .stButton > button *,
        .stDownloadButton > button *{
          color:inherit !important;
          -webkit-text-fill-color:inherit !important;
        }

        .card{padding:1.25rem 1.4rem !important;margin:.9rem 0 !important;}
        .totalbox{
          padding:1.35rem 1rem !important;
          background:
            radial-gradient(circle at 85% 15%,rgba(196,178,255,.23),transparent 34%),
            rgba(255,255,255,.90) !important;
        }
        .totalnum{color:#6045d9 !important;font-size:2.55rem !important;font-weight:900 !important;}
        .cardtitle,.scorebar-value,.scorebar-label{color:#243651 !important;}
        .score{color:#6045d9 !important;}
        .desc-en{color:#425571 !important;}
        .desc-zh,.meta{color:#73839a !important;}
        .scorechart{padding:1.5rem 1rem 1rem !important;}
        .scorebar-track{border-bottom-color:#dce4ef !important;}
        .scorebar-fill{
          opacity:.86 !important;
          box-shadow:0 5px 12px rgba(44,59,98,.09) !important;
        }

        [data-testid="stAlert"]{
          border-radius:16px !important;
          border:1px solid rgba(207,217,233,.95) !important;
          box-shadow:0 6px 18px rgba(58,75,120,.06) !important;
        }
        [data-testid="stAlert"] *{
          color:#263a56 !important;
          -webkit-text-fill-color:#263a56 !important;
        }

        hr{
          border-color:rgba(194,205,221,.72) !important;
          margin:1.7rem 0 !important;
        }

        @media (prefers-color-scheme:dark){
          :root, html, body, .stApp{color-scheme:light !important;}
          html, body, .stApp,
          [data-testid="stAppViewContainer"],
          [data-testid="stMain"]{
            background:
              radial-gradient(circle at 92% 5%, rgba(193,168,255,.24) 0, rgba(193,168,255,0) 26%),
              radial-gradient(circle at 8% 82%, rgba(143,229,215,.20) 0, rgba(143,229,215,0) 28%),
              linear-gradient(180deg,#edf2ff 0%,#f5f8ff 47%,#eefbf8 100%) !important;
            color:#10233f !important;
          }
        }

        @media (max-width:640px){
          .block-container{
            padding-top:.9rem !important;
            padding-left:1rem !important;
            padding-right:1rem !important;
          }
          .wa-hero-icon{
            width:64px !important;
            height:64px !important;
            border-radius:18px !important;
          }
          .wa-hero-icon svg{width:32px !important;height:32px !important;}
          .wa-hero-title{font-size:1.8rem !important;}
          .wa-hero-subtitle{font-size:.95rem !important;padding:0 .25rem;}
          .taskbox,
          .wa-requirements-card,
          .card,
          .totalbox,
          .scorechart{border-radius:19px !important;}
          .taskbox{padding:1.05rem 1.05rem !important;font-size:1rem !important;}
          .wa-requirements-card{padding:1.05rem 1rem !important;}
          .wa-section-title{font-size:1.12rem !important;margin-top:1.35rem !important;}
          .wa-instruction,.wa-requirements{font-size:.97rem !important;}
          .stButton > button,.stDownloadButton > button{min-height:3rem !important;}
          .scorechart{gap:6px !important;padding:1.1rem .4rem .85rem !important;}
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
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
                <div class="wa-hero-icon" aria-hidden="true">
                  <svg viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg">
                    <path d="M8.25 5.25H20.4L24.75 9.6V26.75H8.25V5.25Z"
                          stroke="white" stroke-width="2.1" stroke-linejoin="round"/>
                    <path d="M20 5.5V10H24.5"
                          stroke="white" stroke-width="2.1" stroke-linejoin="round"/>
                    <path d="M11.5 15H20.5M11.5 19H18.5"
                          stroke="white" stroke-width="2" stroke-linecap="round"/>
                    <path d="M10.9 24.1L12.1 20.9L21.55 11.45C22.25 10.75 23.4 10.75 24.1 11.45L24.55 11.9C25.25 12.6 25.25 13.75 24.55 14.45L15.1 23.9L11.9 25.1L10.9 24.1Z"
                          fill="rgba(255,255,255,.28)" stroke="white" stroke-width="1.8" stroke-linejoin="round"/>
                  </svg>
                </div>
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
            st.markdown(
                f'<div class="wa-requirements-card"><div class="wa-requirements">{req_html}</div></div>',
                unsafe_allow_html=True
            )

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
        st.markdown('<div class="wa-instruction">Take a clear photo of your writing and upload it here. Wait until you see <b>Upload complete</b> before pressing Submit.</div>', unsafe_allow_html=True)

        if "upload_widget_version" not in st.session_state:
            st.session_state["upload_widget_version"] = 0

        uploaded = st.file_uploader(
            "Choose an image",
            type=["jpg", "jpeg", "png", "webp"],
            key=f"writing_upload_{st.session_state['upload_widget_version']}",
        )

        if uploaded is None:
            st.info("Choose your photo and wait until the upload is complete.")
        else:
            # Hide Streamlit's native uploaded-file row completely after upload.
            # This avoids the browser/Streamlit-specific blank or dark square icon.
            st.markdown(
                """<style>
                [data-testid="stFileUploader"]{display:none !important;}
                </style>""",
                unsafe_allow_html=True,
            )
            size_mb = len(uploaded.getvalue()) / (1024 * 1024)
            safe_name = html.escape(uploaded.name)
            st.markdown(
                f"""
                <div class="wa-upload-card">
                  <div class="wa-upload-doc" aria-hidden="true">
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M7 2.75h6.7L19.25 8.3V21.25H7V2.75Z"
                            fill="#4F8CFF" opacity=".18"/>
                      <path d="M13.5 2.75V8.5H19.25"
                            stroke="#246BFD" stroke-width="1.8"
                            stroke-linejoin="round"/>
                      <path d="M7 2.75h6.5L19.25 8.5V21.25H7V2.75Z"
                            stroke="#246BFD" stroke-width="1.8"
                            stroke-linejoin="round"/>
                      <path d="M9.5 12.25H16.5M9.5 15.5H16.5M9.5 18.75H14.5"
                            stroke="#246BFD" stroke-width="1.6"
                            stroke-linecap="round"/>
                    </svg>
                  </div>
                  <div>
                    <div class="wa-upload-name">{safe_name}</div>
                    <div class="wa-upload-size">{size_mb:.2f} MB · Upload complete</div>
                  </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Choose a different photo", use_container_width=False):
                st.session_state["upload_widget_version"] += 1
                st.rerun()

        st.markdown('<div class="wa-section-title"><span class="wa-section-icon">✅</span><span>Step 3. Submit your writing</span></div>', unsafe_allow_html=True)
        if uploaded is None:
            st.caption("Submit will be available after the photo upload is complete.")
        if st.button("Submit for Assessment", type="primary", use_container_width=True, disabled=(uploaded is None)):
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
                            if not verify_vivian_stamp(uploaded):
                                st.error("Vivian's stamp could not be verified in the bottom-right corner of your worksheet. Please take a clearer photo that includes the stamp and upload it again. No assessment or submission has been recorded.")
                                st.stop()
                            result = assess(uploaded, task)
                            if not result.get("image_readable", True):
                                st.error("The image is not clear enough to read. Please upload a clearer photo.")
                            else:
                                image_path = upload_submission_image(task["id"], student_id, student_name, uploaded)
                                save_submission(task, seat_number, student_id, student_name, result, image_path)
                                st.session_state["result"] = result
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
                    if result.get("zero_reason") == "task_requirements_copy":
                        desc_en = "The submitted response mainly reproduces the Task Requirements instead of providing an original response, so this dimension receives 0."
                        desc_zh = "本次作答主要是在抄寫 Task Requirements，未提供原創的任務作答，因此此向度為 0 分。"
                    else:
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

            if result.get("zero_reason") == "task_requirements_copy":
                st.info("This submission mainly copies the Task Requirements. Please write your own response to the task.")
            elif not result.get("english_response", True):
                st.info("This response was not written predominantly in English. Please complete the assigned task in English.")
            else:
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
                st.caption("Grammar • Spelling • Punctuation • Capitalization • Sentence Structure")
                corrections = result.get("corrections", [])
                if not corrections:
                    st.success("No clear grammar, spelling, punctuation, capitalization, or sentence structure errors were found.")
                else:
                    for i, c in enumerate(corrections, 1):
                        original = c.get("original_sentence", c.get("original", ""))
                        correction = c.get("corrected_sentence", c.get("correction", ""))
                        explanation_zh = c.get("explanation_zh", c.get("explanation", ""))
                        st.markdown(f"**Sentence {i}**")
                        st.markdown(f"- **Original:** {original}")
                        st.markdown(f"- **Corrected:** {correction}")
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
