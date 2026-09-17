
import os, json, uuid, base64
from pathlib import Path
import streamlit as st
from openai import OpenAI

APP_NAME = "Writing Assessment"
TASKS_FILE = Path("tasks.json")

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

def load_tasks():
    if TASKS_FILE.exists():
        try:
            return json.loads(TASKS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_tasks(tasks):
    TASKS_FILE.write_text(json.dumps(tasks, ensure_ascii=False, indent=2), encoding="utf-8")

def get_base_url():
    base = ""
    try:
        base = st.secrets.get("APP_BASE_URL", "")
    except Exception:
        pass
    return (base or os.getenv("APP_BASE_URL", "")).rstrip("/")

def student_link(task_id):
    base = get_base_url()
    return f"{base}/?task={task_id}" if base else f"?task={task_id}"

def image_to_data_url(uploaded_file):
    mime = uploaded_file.type or "image/jpeg"
    b64 = base64.b64encode(uploaded_file.getvalue()).decode("utf-8")
    return f"data:{mime};base64,{b64}"

def build_prompt(task_title, genre):
    return f"""
You are an English writing assessor.

Teacher settings:
Task Title: {task_title}
Genre / Writing Type: {genre}

Score exactly these four dimensions, each with an integer from 1 to 4:
- Content & Task Fulfillment
- Organization & Coherence
- Language Use
- Genre & Professional Appropriacy

Rules:
- No half points.
- Do not add criteria.
- Do not double-penalize the same issue.
- The teacher provides only Task Title and Genre / Writing Type.
- Do not invent unstated task requirements.
- For Content & Task Fulfillment, judge relevance, development, support, and fulfillment of the communicative purpose reasonably implied by the title and genre.
- If the image is not readable enough, do not guess.

Then identify genuine errors ONLY in:
- Grammar
- Spelling
- Punctuation

For each error, provide:
- type
- original
- correction
- brief A2-B1 English explanation

Do not rewrite the full composition.
Preserve the student's intended meaning.

Finally, provide 1 to 3 brief suggestions about CONTENT and ORGANIZATION only.
Do not provide a model essay.

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
  "content_organization_suggestions": [
    "brief suggestion"
  ]
}}
"""

def assess(uploaded_file, task_title, genre):
    api_key = None
    model = "gpt-5.6-luna"
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
        model = st.secrets.get("OPENAI_MODEL", model)
    except Exception:
        pass

    api_key = api_key or os.getenv("OPENAI_API_KEY")
    model = os.getenv("OPENAI_MODEL", model)

    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not configured.")

    client = OpenAI(api_key=api_key)

    response = client.responses.create(
        model=model,
        input=[{
            "role": "user",
            "content": [
                {"type": "input_text", "text": build_prompt(task_title, genre)},
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

st.set_page_config(page_title=APP_NAME, page_icon="✍️", layout="centered")

st.markdown("""
<style>
.block-container{max-width:860px;padding-top:2rem;padding-bottom:4rem}
.taskbox,.card,.totalbox{border:1px solid rgba(120,120,120,.28);border-radius:14px;padding:1rem 1.1rem;margin:.7rem 0}
.totalbox{text-align:center;border-width:2px}
.totalnum{font-size:2.2rem;font-weight:800}
.cardtitle{font-size:1.05rem;font-weight:700}
.score{float:right;font-weight:800}
.level{font-weight:700;margin-top:.35rem}
.muted{opacity:.7}
</style>
""", unsafe_allow_html=True)

tasks = load_tasks()
task_id = st.query_params.get("task")

if task_id and task_id in tasks:
    task = tasks[task_id]
    st.title("Writing Assessment")
    st.write("Follow the steps below to check your writing.")

    st.markdown(
        f'<div class="taskbox"><b>Task Title:</b> {task["title"]}<br><b>Genre / Writing Type:</b> {task["genre"]}</div>',
        unsafe_allow_html=True
    )

    st.subheader("Step 1. Upload your writing")
    st.write("Take a clear photo of your writing and upload it here.")
    uploaded = st.file_uploader("Choose an image", type=["jpg","jpeg","png","webp"])

    st.subheader("Step 2. Submit your writing")
    if st.button("Submit for Assessment", type="primary", use_container_width=True):
        if not uploaded:
            st.warning("Please upload an image first.")
        else:
            with st.spinner("Checking your writing..."):
                try:
                    result = assess(uploaded, task["title"], task["genre"])
                    if not result.get("image_readable", True):
                        st.error("The image is not clear enough to read. Please upload a clearer photo.")
                    else:
                        st.session_state["result"] = result
                        st.session_state["show_revision"] = False
                except Exception as e:
                    st.error(f"Assessment could not be completed: {e}")

    result = st.session_state.get("result")
    if result:
        st.divider()
        st.header("Your Writing Score")
        total = sum(result["scores"][d] for d in RUBRIC)
        st.markdown(f'<div class="totalbox"><div class="muted">Total Score</div><div class="totalnum">{total} / 16</div></div>', unsafe_allow_html=True)

        for dim, levels in RUBRIC.items():
            score = result["scores"][dim]
            level, desc = levels[score]
            st.markdown(
                f'<div class="card"><span class="cardtitle">{dim}</span><span class="score">{score}/4</span><div class="level">{level}</div><div>{desc}</div></div>',
                unsafe_allow_html=True
            )

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
    st.subheader("Teacher Setup")
    st.write("Create a task and share the student link.")

    title = st.text_input("Task Title", placeholder="e.g., Use of Learning Resources")
    genre = st.text_input("Genre / Writing Type", placeholder="e.g., Email")

    if st.button("Create Student Link", type="primary", use_container_width=True):
        if not title.strip() or not genre.strip():
            st.warning("Please enter both Task Title and Genre / Writing Type.")
        else:
            code = uuid.uuid4().hex[:10]
            tasks[code] = {"title": title.strip(), "genre": genre.strip()}
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
                st.markdown(f"**{task['title']}**")
                st.write(f"Genre / Writing Type: {task['genre']}")
                st.code(student_link(code))
