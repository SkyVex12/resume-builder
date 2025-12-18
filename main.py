import os
import re
import json
import time
import contextlib
import io
from datetime import datetime
from pathlib import Path

import keyboard
import pyperclip
from winotify import Notification

from docx import Document
from docx2pdf import convert

from openai import OpenAI

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ================= CONFIG ================= #

MODEL = "gpt-4.1-mini"
TEMPERATURE_MAIN = 0.35
TEMPERATURE_REGEN = 0.70

COMPANY_COUNT = 4

# IMPORTANT: You generate 5 people, so reflect that everywhere
PERSON_ORDER = ["Timothy", "Wilfredo", "Lou", "Ryan", "James"]

TEMPLATES = {
    "Timothy": "templates/Timothy.docx",
    "Wilfredo": "templates/Wilfredo.docx",
    "Lou": "templates/Lou.docx",
    "Ryan": "templates/Ryan.docx",
    "James": "templates/James.docx"
}

SKILL_CATEGORY_ORDER = [
    "Programming Languages",
    "Frameworks & Libraries",
    "Databases",
    "Cloud & DevOps",
    "Testing",
    "Tools & Practices",
]

SOFT_SKILL_SIGNAL_WORDS = {
    "team", "teams", "cross-functional", "cross functional",
    "collaboration", "collaborative", "coordination", "partnering", "partnership",
    "communication", "communicating", "stakeholder", "stakeholders", "presentation", "presenting",
    "alignment", "aligned",
    "leadership", "leading", "lead", "ownership", "accountability", "responsibility", "responsible",
    "decision", "decision-making",
    "process", "processes", "process changes", "process improvement", "continuous improvement",
    "execution", "delivery", "planning", "prioritization", "roadmap", "strategy",
    "mentoring", "mentor", "mentorship", "coaching", "guidance", "training",
    "problem", "problem-solving", "analysis", "analytical", "investigation", "troubleshooting",
    "quality", "reliability", "best practices", "standards", "compliance",
    "customer", "customers", "client", "clients", "business", "business needs", "requirements",
    "adaptability", "flexibility", "change", "changes", "improvement", "optimization",
    "empathy", "empathetic", "empathic",
    "attention to detail",
    "independently", "collaborate", "pressure", "creative", "organized", "consistent",
    "detail-oriented", "interpersonal", "multitasking", "proactive", "self-motivated", "strategic",
    "perfectionist", "diligent", "dependable", "dedicated", "innovative"
}

HARD_SKILL_KEYWORDS = {
    "java", "python", "golang", "php", "javascript", "typescript",
    "react", "next.js", "node.js", "django", "flask", "laravel",
    "mysql", "postgresql", "mongodb", "firebase", "redis",
    "aws", "azure", "docker", "kubernetes", "jenkins", "ci/cd",
    "jest", "cypress", "mocha",
    "jira", "git", "rest", "api", "microservices", "security", "ai"  # ensure lowercase
}

OUTPUT_DIR = "output"
BULLET = "  \u2022    "  # bullet point


# Uniqueness thresholds / loops
SIMILARITY_THRESHOLD = 0.65
MAX_REGEN_ROUNDS = 4
AVOID_TEXT_MAX_CHARS = 3500  # keep avoid block small to control tokens


# ================= PERSON PROFILES ================= #

PERSON_PROFILES = {
    "Timothy": {
        "style": "strategic, architecture-focused, leadership-driven",
        "summary": "senior software engineer focused on scalable systems",
        "metrics": "latency, uptime, scale, cost, reliability"
    },
    "Wilfredo": {
        "style": "hands-on, delivery-focused, optimization-oriented",
        "summary": "experienced senior full-stack engineer with strong execution",
        "metrics": "throughput, build time, cost, delivery speed, defect reduction"
    },
    "Lou": {
        "style": "collaborative, balanced, system improvement oriented",
        "summary": "senior software engineer with growing ownership and impact",
        "metrics": "MTTR, incident rate, test coverage, availability, customer satisfaction"
    },
    "Ryan": {
        "style": "practical, implementation-heavy, learning-focused",
        "summary": "senior software engineer with solid fundamentals and growth mindset",
        "metrics": "feature adoption, support tickets, time-to-ship, performance, usage"
    },
    "James": {
        "style": "results-driven, execution-focused, performance optimization oriented",
        "summary": "senior full-stack engineer with a strong track record of delivering high-impact features",
        "metrics": "conversion, revenue, latency, cost, retention"
    }
}


# ================= DEFAULT SKILLS ================= #

DEFAULT_SKILLS = {
    "Programming Languages": ["JavaScript", "TypeScript", "Python", "PHP", "Ruby", "Golang", "Java", "C#"],
    "Frameworks & Libraries": ["Microservices", "REST APIs", "React.js", "Next.js", "Node.js", "Django", "Flask", "Laravel"],
    "Databases": ["MySQL", "PostgreSQL", "MongoDB", "Firebase", "Redis"],
    "Testing": ["Jest", "React Testing Library", "Cypress", "Mocha", "Go-Testify"],
    "Cloud & DevOps": ["AWS", "Azure", "Docker", "Jenkins"],
    "Tools & Practices": ["Git", "Jira", "CI/CD", "Agile", "Scrum"]
}

MAX_SKILLS_PER_CATEGORY = 8


# ================= INIT ================= #

def notify(message: str):
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        Notification(
            app_id="Resume Generator",
            title="Resume Generator",
            msg=message,
            duration="short"
        ).show()


def get_client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY environment variable.")
    return OpenAI(api_key=api_key)


client = get_client()

output_base_dir = None
folder_opened = False
is_running = False


# ================= SOFT SKILL EXTRACTION ================= #

def extract_soft_skills(jd_text: str) -> list[str]:
    jd = jd_text.lower()
    candidates = re.findall(r"\b[a-z]+(?:[-\s][a-z]+){1,3}\b", jd)
    soft = set()

    for phrase in candidates:
        phrase = phrase.strip()
        if not any(sig in phrase for sig in SOFT_SKILL_SIGNAL_WORDS):
            continue
        if any(h in phrase for h in HARD_SKILL_KEYWORDS):
            continue
        soft.add(phrase)

    return sorted(soft)


# ================= PROMPTS ================= #

def people_block() -> str:
    return "\n".join(f"- {name}: {PERSON_PROFILES[name]['style']}" for name in PERSON_ORDER)


def build_prompt_all(jd_text: str, soft_skills_text: str) -> str:
    return f"""
You are an ATS optimization engine.
Output valid JSON only.

GOAL:
Generate FIVE DISTINCT senior-level resumes for the SAME Job Description.
Target ATS score: 90–95%+.

ABSOLUTE RULES:
- Use EXACT keywords and phrases from the Job Description where possible.
- Repeat critical JD keywords across Summary, Skills, and Experience (ATS).
- Each person MUST have unique wording, metrics, and sentence structure.
- Do NOT reuse bullets across people.
- Output JSON ONLY. No markdown. No explanations.
- All soft skills and hard skills mentioned in the JD MUST appear across ALL resumes.
- Treat items like 'front-end development', 'cloud infrastructure', 'Information Systems', 'web services', 'Software Development Lifecycle' as hard skills too when present.

PEOPLE & STYLE DIFFERENTIATION:
{people_block()}

SOFT SKILLS (USE VERBATIM WORDING FROM JD):
{soft_skills_text}

OUTPUT FORMAT (STRICT JSON — NO TRAILING COMMAS):
{{
  "Timothy": {{
    "summary": "string",
    "skills": {{
      "Programming Languages": [],
      "Frameworks & Libraries": [],
      "Databases": [],
      "Cloud & DevOps": [],
      "Testing": [],
      "Tools & Practices": []
    }},
    "experience": {{
      "company_1": [],
      "company_2": [],
      "company_3": [],
      "company_4": []
    }}
  }},
  "Wilfredo": {{
    "summary": "string",
    "skills": {{
      "Programming Languages": [],
      "Frameworks & Libraries": [],
      "Databases": [],
      "Cloud & DevOps": [],
      "Testing": [],
      "Tools & Practices": []
    }},
    "experience": {{
      "company_1": [],
      "company_2": [],
      "company_3": [],
      "company_4": []
    }}
  }},
  "Lou": {{
    "summary": "string",
    "skills": {{
      "Programming Languages": [],
      "Frameworks & Libraries": [],
      "Databases": [],
      "Cloud & DevOps": [],
      "Testing": [],
      "Tools & Practices": []
    }},
    "experience": {{
      "company_1": [],
      "company_2": [],
      "company_3": [],
      "company_4": []
    }}
  }},
  "Ryan": {{
    "summary": "string",
    "skills": {{
      "Programming Languages": [],
      "Frameworks & Libraries": [],
      "Databases": [],
      "Cloud & DevOps": [],
      "Testing": [],
      "Tools & Practices": []
    }},
    "experience": {{
      "company_1": [],
      "company_2": [],
      "company_3": [],
      "company_4": []
    }}
  }},
  "James": {{
    "summary": "string",
    "skills": {{
      "Programming Languages": [],
      "Frameworks & Libraries": [],
      "Databases": [],
      "Cloud & DevOps": [],
      "Testing": [],
      "Tools & Practices": []
    }},
    "experience": {{
      "company_1": [],
      "company_2": [],
      "company_3": [],
      "company_4": []
    }}
  }}
}}

SUMMARY RULES:
- Senior-level tone ONLY for ALL people.
- Start summary with "Senior Software Engineer" and include the EXACT job title from the Job Description.
- 4–5 concise lines.
- Include 6–8 exact JD keywords.

SKILLS RULES:
- Categorize skills strictly under the provided categories.
- Use JD keywords verbatim where possible.
- Include 6–10 skills per category.
- Prioritize skills mentioned multiple times in the JD.
- 'Tools & Practices' must only include tools and practices in the JD.

EXPERIENCE RULES:
- company_1 and company_2 → Senior-level responsibilities, 6–8 bullets each
- company_3 → Mid-level responsibilities, 6 bullets
- company_4 → Junior-level responsibilities, 4–6 bullets
- EACH bullet MUST:
  - Start with a strong action verb
  - Include at least ONE JD hard skill (verbatim)
  - Include at least ONE soft skill (verbatim)
  - Include a measurable metric (%,$,users,latency,scale)
- Enforce uniqueness across people: different verbs, different metric types, different framing.

JOB DESCRIPTION:
{jd_text}

Return ONLY valid JSON.
""".strip()


def build_prompt_one_person(jd_text: str, soft_skills_text: str, person: str, avoid_text: str) -> str:
    profile = PERSON_PROFILES[person]
    metric_hint = profile.get("metrics", "")

    return f"""
You are an ATS optimization engine.
Output valid JSON only.

Generate ONE resume for: {person}
Style: {profile["style"]}

HARD CONSTRAINT:
Do NOT reuse wording, sentence patterns, or bullet structures from the AVOID TEXT below.
Use different opening verbs, different clause structures, and different metric families.

AVOID TEXT:
{avoid_text}

SOFT SKILLS (USE VERBATIM WORDING FROM JD):
{soft_skills_text}

Prefer metrics like: {metric_hint}

OUTPUT JSON (NO TRAILING COMMAS):
{{
  "{person}": {{
    "summary": "string",
    "skills": {{
      "Programming Languages": [],
      "Frameworks & Libraries": [],
      "Databases": [],
      "Cloud & DevOps": [],
      "Testing": [],
      "Tools & Practices": []
    }},
    "experience": {{
      "company_1": [],
      "company_2": [],
      "company_3": [],
      "company_4": []
    }}
  }}
}}

RULES:
- Summary starts with "Senior Software Engineer" and includes the EXACT job title from the JD.
- Bullets must include: (1) one JD hard skill verbatim, (2) one soft skill verbatim, (3) one metric.
- Make all sentences structurally distinct from AVOID TEXT.

JOB DESCRIPTION:
{jd_text}

Return ONLY valid JSON.
""".strip()


# ================= OPENAI RESPONSE PARSING ================= #

def extract_text(response) -> str:
    if hasattr(response, "output_text") and response.output_text:
        return response.output_text.strip()

    texts = []
    for message in getattr(response, "output", []):
        for block in getattr(message, "content", []):
            t = getattr(block, "text", None)
            if t:
                texts.append(t)
    return "\n".join(texts).strip()


def safe_parse_json(raw_text: str):
    if not raw_text:
        return None
    raw_text = raw_text.strip()
    try:
        return json.loads(raw_text)
    except Exception:
        return None


def generate_all(jd: str, soft_skills_text: str) -> dict:
    response = client.responses.create(
        model=MODEL,
        input=build_prompt_all(jd, soft_skills_text),
        temperature=TEMPERATURE_MAIN,
        timeout=90
    )
    text = extract_text(response)
    data = safe_parse_json(text)
    if not isinstance(data, dict):
        raise ValueError("Invalid JSON from model.")
    return data


def regenerate_one(jd: str, soft_skills_text: str, person: str, avoid_text: str) -> dict:
    response = client.responses.create(
        model=MODEL,
        input=build_prompt_one_person(jd, soft_skills_text, person, avoid_text),
        temperature=TEMPERATURE_REGEN,
        timeout=90
    )
    text = extract_text(response)
    data = safe_parse_json(text)
    if not isinstance(data, dict) or person not in data:
        raise ValueError(f"Invalid regen JSON for {person}.")
    return data[person]


# ================= UNIQUENESS CHECKING ================= #

def _collect_all_bullets(all_data: dict) -> tuple[list[str], list[str]]:
    bullets = []
    owners = []
    for person in PERSON_ORDER:
        pdata = all_data.get(person, {})
        exp = pdata.get("experience", {})
        for i in range(1, COMPANY_COUNT + 1):
            for b in exp.get(f"company_{i}", []) or []:
                if isinstance(b, str) and b.strip():
                    bullets.append(b.strip())
                    owners.append(person)
    return bullets, owners


def find_flagged_people(all_data: dict, threshold: float) -> list[str]:
    bullets, owners = _collect_all_bullets(all_data)
    if len(bullets) < 2:
        return []

    vec = TfidfVectorizer(ngram_range=(1, 2), stop_words="english")
    X = vec.fit_transform(bullets)
    sim = cosine_similarity(X)

    flagged = set()
    n = len(bullets)
    for i in range(n):
        for j in range(i + 1, n):
            if owners[i] != owners[j] and sim[i, j] >= threshold:
                flagged.add(owners[i])
                flagged.add(owners[j])

    return sorted(flagged)


def build_avoid_text(all_data: dict, exclude_person: str, max_chars: int) -> str:
    chunks = []
    for person in PERSON_ORDER:
        if person == exclude_person:
            continue
        pdata = all_data.get(person, {})
        if pdata.get("summary"):
            chunks.append(pdata["summary"])
        exp = pdata.get("experience", {})
        for i in range(1, COMPANY_COUNT + 1):
            chunks.extend(exp.get(f"company_{i}", []) or [])

    joined = "\n".join([c for c in chunks if isinstance(c, str)])
    return joined[:max_chars]


def enforce_uniqueness(jd: str, soft_skills_text: str, all_data: dict,
                       threshold: float = SIMILARITY_THRESHOLD,
                       max_rounds: int = MAX_REGEN_ROUNDS) -> dict:
    for round_idx in range(1, max_rounds + 1):
        flagged = find_flagged_people(all_data, threshold)
        if not flagged:
            return all_data

        # Regenerate only the people contributing to overlap
        for person in flagged:
            avoid_text = build_avoid_text(all_data, person, AVOID_TEXT_MAX_CHARS)
            try:
                all_data[person] = regenerate_one(jd, soft_skills_text, person, avoid_text)
            except Exception as e:
                # If regen fails, keep existing and continue
                print(f"[WARN] Regen failed for {person}: {e}")

        # small pause to avoid rapid retry bursts
        time.sleep(0.2)

    return all_data


# ================= SKILLS MERGE ================= #

def merge_skills(gpt_skills: dict) -> dict:
    merged = {}
    for category in set(gpt_skills) | set(DEFAULT_SKILLS):
        s = set(gpt_skills.get(category, []) or [])
        s.update(DEFAULT_SKILLS.get(category, []) or [])
        merged[category] = sorted(s)[:MAX_SKILLS_PER_CATEGORY]
    return merged


# ================= DOCX HELPERS ================= #

def replace_placeholder_plain(doc: Document, placeholder: str, value: str):
    # NOTE: This destroys runs/formatting in that paragraph. Only use for plain sections.
    for p in doc.paragraphs:
        if placeholder in p.text:
            p.text = p.text.replace(placeholder, value)


def replace_skills_placeholder(doc: Document, placeholder: str, skills: dict):
    # Bold only categories using runs
    for p in doc.paragraphs:
        if placeholder in p.text:
            p.clear()
            for category in SKILL_CATEGORY_ORDER:
                items = skills.get(category, [])
                if not items:
                    continue
                run_cat = p.add_run(f"{category}: ")
                run_cat.bold = True
                p.add_run(", ".join(items))
                p.add_run("\n")


def fill_template(template_path: str, data: dict, output_path: str):
    doc = Document(template_path)

    replace_placeholder_plain(doc, "{{SUMMARY}}", data.get("summary", ""))

    merged_skills = merge_skills(data.get("skills", {}))
    replace_skills_placeholder(doc, "{{SKILLS}}", merged_skills)

    for i in range(1, COMPANY_COUNT + 1):
        bullets_list = data.get("experience", {}).get(f"company_{i}", []) or []
        bullets = f"\n{BULLET}".join(bullets_list)
        bullets = bullets.removesuffix(BULLET)
        bullets = BULLET + bullets if bullets else ""
        replace_placeholder_plain(doc, f"{{{{EXP_COMPANY_{i}}}}}", bullets)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)


def batch_convert_to_pdf(folder: str):
    folder_path = Path(folder)
    for docx_path in folder_path.glob("*.docx"):
        pdf_path = folder_path / f"{docx_path.stem}.pdf"
        convert(str(docx_path), str(pdf_path))


# ================= MAIN ================= #

def on_hotkey_generate():
    global output_base_dir, folder_opened, is_running

    if is_running:
        return
    is_running = True

    try:
        jd = pyperclip.paste()
        if not jd.strip():
            notify("Clipboard is empty")
            return

        if output_base_dir is None:
            output_base_dir = os.path.join(OUTPUT_DIR, "JD_resumes")
            os.makedirs(output_base_dir, exist_ok=True)

        soft_skills = extract_soft_skills(jd)
        soft_skills_text = ", ".join(soft_skills) if soft_skills else "None"

        notify("Generating resumes...")
        all_data = generate_all(jd, soft_skills_text)

        # Enforce uniqueness across people
        notify("Improving uniqueness...")
        all_data = enforce_uniqueness(jd, soft_skills_text, all_data)

        # Write DOCX
        for person in PERSON_ORDER:
            pdata = all_data.get(person)
            if not pdata:
                notify(f"{person} missing from model output")
                continue

            template = TEMPLATES[person]
            output_docx = os.path.join(output_base_dir, f"{person}_resume.docx")

            fill_template(template, pdata, output_docx)
            notify(f"{person} DOCX generated")

            if not folder_opened:
                try:
                    os.startfile(output_base_dir)
                except Exception:
                    pass
                folder_opened = True

        notify("DOCX generation done. Converting to PDF...")
        batch_convert_to_pdf(output_base_dir)
        notify("All PDFs generated")

    except Exception as e:
        print("Error:", e)
        notify("Error generating resumes")
    finally:
        is_running = False


def on_hotkey_pdf_only():
    global output_base_dir, is_running

    if is_running:
        return
    is_running = True

    try:
        output_base_dir = os.path.join(OUTPUT_DIR, "JD_resumes")
        if not os.path.isdir(output_base_dir):
            notify("No output folder found. Generate resumes first.")
            return

        notify("Converting DOCX to PDF...")
        batch_convert_to_pdf(output_base_dir)
        notify("All PDFs generated")

    except Exception as e:
        print("Error:", e)
        notify("Error generating PDFs")
    finally:
        is_running = False


def main():
    notify("Resume Generator is running")
    keyboard.add_hotkey("ctrl+q", on_hotkey_generate)
    keyboard.add_hotkey("ctrl+alt+p", on_hotkey_pdf_only)
    keyboard.wait()


if __name__ == "__main__":
    main()
