import os
import json
import keyboard
import pyperclip
from datetime import datetime
from docx import Document
from openai import OpenAI
from winotify import Notification

# ================= CONFIG ================= #

MODEL = "gpt-4.1-mini"
TEMPERATURE = 0.35
COMPANY_COUNT = 4

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
    "Tools & Methodologies",
    # "Soft Skills"  # if you use it
]

SOFT_SKILL_SIGNAL_WORDS = {
    # Team / Collaboration
    "team", "teams",
    "cross-functional", "cross functional",
    "collaboration", "collaborative",
    "coordination", "partnering", "partnership",

    # Communication
    "communication", "communicating",
    "stakeholder", "stakeholders",
    "presentation", "presenting",
    "alignment", "aligned",

    # Leadership / Ownership
    "leadership", "leading", "lead",
    "ownership", "accountability",
    "responsibility", "responsible",
    "decision", "decision-making",

    # Process / Execution
    "process", "processes",
    "process changes", "process improvement",
    "continuous improvement",
    "execution", "delivery",
    "planning", "prioritization",
    "roadmap", "strategy",

    # Mentorship / Growth
    "mentoring", "mentor",
    "mentorship", "coaching",
    "guidance", "training",

    # Problem / Analysis
    "problem", "problem-solving",
    "analysis", "analytical",
    "investigation", "troubleshooting",

    # Quality / Reliability
    "quality", "reliability",
    "best practices",
    "standards", "compliance",

    # Customer / Business
    "customer", "customers",
    "client", "clients",
    "business", "business needs",
    "requirements",

    # Adaptability
    "adaptability", "flexibility",
    "change", "changes",
    "improvement", "optimization",
    
    "empathy",
    "empathetic",
    "empathic", 
    "attention to detail",
    "highly competitive","ability to adapt",
    "motivating", "ability to learn quickly",
    "consistent"
}


HARD_SKILL_KEYWORDS = {
    "java", "python", "golang", "php",
    "javascript", "typescript",
    "react", "next.js", "node.js",
    "django", "flask", "laravel",
    "mysql", "postgresql", "mongodb",
    "aws", "azure", "docker", "kubernetes",
    "jenkins", "ci/cd",
    "jest", "cypress", "mocha",
    "jira", "git", "rest", "api", "microservices", "security","AI"
}

OUTPUT_DIR = "output"

# ================= PERSON PROFILES ================= #

PERSON_PROFILES = {
    "Timothy": {
        "style": "strategic, architecture-focused, leadership-driven",
        "summary": "senior software engineer focused on scalable systems"
    },
    "Wilfredo": {
        "style": "hands-on, delivery-focused, optimization-oriented",
        "summary": "experienced senior full-stack engineer with strong execution"
    },
    "Lou": {
        "style": "collaborative, balanced, system improvement oriented",
        "summary": "senior software engineer with growing ownership and impact"
    },
    "Ryan": {
        "style": "practical, implementation-heavy, learning-focused",
        "summary": "senior software engineer with solid fundamentals and growth mindset"
    },
    "James": {
        "style": "results-driven, execution-focused, performance optimization oriented",
        "summary": "senior full-stack engineer with a strong track record of delivering high-impact features and performance improvements"
    }
    
}


# ================= DEFAULT SKILLS ================= #

DEFAULT_SKILLS = {
    "Technical Skills": [
        "JavaScript", "TypeScript", "Python", "PHP",
        "Ruby", "Golang", "Microservices", "REST APIs"
    ],
    "Frameworks & Libraries": [
        "React.js", "Next.js", "Node.js",
        "Django", "Flask", "Laravel"
    ],
    "Databases": [
        "MySQL", "PostgreSQL", "MongoDB"
    ],
    "Testing": [
        "Jest", "React Testing Library",
        "Cypress", "Mocha", "Go-Testify"
    ],
    "Cloud & DevOps": [
        "AWS", "Azure", "Docker", "Jenkins"
    ],
    "Tools & Practices": [
        "Git", "Jira", "CI/CD", "Agile", "Scrum"
    ]
}

MAX_SKILLS_PER_CATEGORY = 8

# ================= INIT ================= #

client = OpenAI(api_key="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")

PERSON_KEYS = list(TEMPLATES.keys())
current_person_index = 0
output_base_dir = None
folder_opened = False
is_running = False
BULLET = "  \u2022    "  # Unicode for bullet point

# ================= NOTIFY ================= #
import contextlib
import io


def notify(message):
    with contextlib.redirect_stdout(io.StringIO()), \
         contextlib.redirect_stderr(io.StringIO()):
        Notification(
            app_id="Resume Generator",
            title="Resume Generator",
            msg=message,
            duration="short"
        ).show()

# ================= PROMPT ================= #
def build_prompt(jd_text, soft_skills_text):
    people_block = "\n".join(
        f"- {name}: {profile['style']}"
        for name, profile in PERSON_PROFILES.items()
    )

    return f"""
You are an ATS optimization engine.
Output valid JSON only.

GOAL:
Generate FOUR DISTINCT senior-level resumes for the SAME Job Description.
Target ATS score: 90–95%+.

ABSOLUTE RULES:
- Use EXACT keywords and phrases from the Job Description.
- Repeat critical JD keywords across Summary, Skills, and Experience.
- Each person MUST have unique wording, metrics, and sentence structure.
- Do NOT reuse bullets across people.
- Output JSON ONLY. No markdown. No explanations.

PEOPLE & STYLE DIFFERENTIATION:
{people_block}

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
      "Tools & Methodologies": []
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
      "Tools & Methodologies": []
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
      "Tools & Methodologies": []
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
      "Tools & Methodologies": []
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
      "Tools & Methodologies": []
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
- Start summary with the EXACT job title from the Job Description.
- 4–5 concise lines.
- Include 6–8 exact JD keywords.

SKILLS RULES:
- Categorize skills strictly under the provided categories.
- Use JD keywords verbatim where possible.
- Include Soft Skills explicitly.

EXPERIENCE RULES:
- company_1 and company_2 → Senior-level responsibilities
- company_3 → Mid-level responsibilities
- company_4 → Junior-level responsibilities
- 4–6 bullets per company.
- EACH bullet MUST:
  - Start with a strong action verb
  - Include at least ONE JD hard skill (verbatim)
  - Include at least ONE soft skill (verbatim)
  - Include a measurable metric (%,$,users,latency,scale)

JOB DESCRIPTION:
\"\"\"
{jd_text}
\"\"\"

Before returning, verify that the JSON is syntactically valid.
Return ONLY valid JSON.
"""


# ================= GPT ================= #
import re

def extract_soft_skills(jd_text: str) -> list[str]:
    """
    Extract ATS-safe soft skills from a Job Description.
    Soft skills are defined as JD phrases (2–4 words)
    that contain behavioral signal words and exclude hard skills.
    """

    jd_text = jd_text.lower()

    # Extract 2–4 word phrases
    candidates = re.findall(
        r"\b[a-z]+(?:[-\s][a-z]+){1,3}\b",
        jd_text
    )

    soft_skills = set()

    for phrase in candidates:
        phrase = phrase.strip()

        # Must contain a behavioral signal word
        if not any(signal in phrase for signal in SOFT_SKILL_SIGNAL_WORDS):
            continue

        # Must NOT contain hard skills
        if any(hard in phrase for hard in HARD_SKILL_KEYWORDS):
            continue

        soft_skills.add(phrase)

    return sorted(soft_skills)

def extract_text(response):
    # 1️⃣ Preferred: official shortcut
    if hasattr(response, "output_text") and response.output_text:
        return response.output_text.strip()

    # 2️⃣ Fallback: manual extraction
    texts = []
    for message in getattr(response, "output", []):
        for block in getattr(message, "content", []):
            text = getattr(block, "text", None)
            if text:
                texts.append(text)

    return "\n".join(texts).strip()

def safe_parse_json(raw_text: str):
    if not raw_text:
        return None

    raw_text = raw_text.strip()
    if not raw_text:
        return None

    try:
        return json.loads(raw_text)
    except Exception:
        return None
    
def generate_content_all(jd, soft_skills_text):
    response = client.responses.create(
        model=MODEL,
        input=build_prompt(jd, soft_skills_text),
        temperature=TEMPERATURE,
        timeout=90
    )

    text = extract_text(response)
    if not text:
        raise ValueError("Empty GPT response")

    data = safe_parse_json(text)
    if not isinstance(data, dict):
        raise ValueError("Invalid GPT JSON")

    return data


# ================= SKILLS ================= #

def merge_skills(gpt_skills):
    merged = {}
    for category in set(gpt_skills) | set(DEFAULT_SKILLS):
        skills = set(gpt_skills.get(category, []))
        skills.update(DEFAULT_SKILLS.get(category, []))
        merged[category] = sorted(skills)[:MAX_SKILLS_PER_CATEGORY]
    return merged

# def format_skills(skills):
#     return "\n".join(
#         f"{cat}: {', '.join(items)}"
#         for cat, items in skills.items()
#         if items
#     )

def format_skills(skills: dict) -> str:
    lines = []
    for category in SKILL_CATEGORY_ORDER:
        items = skills.get(category, [])
        if items:
            lines.append(f"{category}: {', '.join(items)}")
    return "\n".join(lines)

# ================= DOCX To PDF ================= #
from docx2pdf import convert
from pathlib import Path

def batch_convert(folder):
    global  is_running

    print('Starting batch conversion to PDF in folder:', folder)
    for doc in Path(f"{folder}").glob("*.docx"):
        print('Converting to PDF:', doc)
        convert(doc, f"{folder}/{doc.stem}.pdf")

# ================= DOCX ================= #

def replace_placeholder(doc, placeholder, value):
    for p in doc.paragraphs:
        if placeholder in p.text:
            p.text = p.text.replace(placeholder, value)

def fill_template(template_path, data, output_path):
    print('-------------------------------------------')
    doc = Document(template_path)

    print('data to fill:', data["summary"])
    replace_placeholder(doc, "{{SUMMARY}}", data["summary"])

    merged_skills = merge_skills(data["skills"])
    replace_placeholder(doc, "{{SKILLS}}", format_skills(merged_skills))
    for i in range(1, COMPANY_COUNT + 1):
        bullets = f"\n{BULLET}".join(data["experience"][f"company_{i}"])
        bullets = bullets.removesuffix(BULLET)
        bullets = BULLET + bullets if bullets else ""

        replace_placeholder(doc, f"{{{{EXP_COMPANY_{i}}}}}", bullets)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)


# ================= MAIN ================= #

def on_hotkey():
    global output_base_dir, folder_opened, is_running

    if is_running:
        return

    is_running = True

    try:
        jd = pyperclip.paste()
        if not jd.strip():
            notify("Clipboard is empty")
            return
        # Create output folder once
        if output_base_dir is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_base_dir = os.path.join(OUTPUT_DIR, f"JD_resumes")
            os.makedirs(output_base_dir, exist_ok=True)
        print('output_base_dir:', output_base_dir)
        soft_skills = extract_soft_skills(jd)
        soft_skills_text = ", ".join(soft_skills) if soft_skills else "None"
        notify(f"Started generating resume...")
        all_data = generate_content_all(jd, soft_skills_text)
        for index, person in enumerate(PERSON_KEYS, start=1):
            data = all_data.get(person)
            if not data:
                notify(f"{person} missing from GPT output")
                continue
            print('person_key:', PERSON_KEYS)
            notify(f"Generating {person} resume...")
            
            template = TEMPLATES[person]
            print('template:', template)

            output_docx = os.path.join(
                output_base_dir,
                f"{person}_resume.docx"
            )
            
            fill_template(template, data, output_docx)

            notify(f"{person} resume generated")

            # Open folder ONLY for first person
            if not folder_opened:
                try:
                    os.startfile(output_base_dir)
                except:
                    pass
                folder_opened = True

        notify("All resume docs generated")

        notify("Started PDF generation")

        batch_convert(output_base_dir)
        
        notify("All resume PDFs generated")

    except Exception as e:
        print('Error:', e)
        notify("Error generating resumes")

    finally:
        is_running = False

def convertToPDF():
    global output_base_dir, folder_opened, is_running
    output_base_dir = os.path.join(OUTPUT_DIR, f"JD_resumes")

    if is_running:
        return

    is_running = True

    try:
        if output_base_dir is None:
            notify("No resumes to convert. Generate resumes first.")
            return

        notify("Started PDF generation")

        batch_convert(output_base_dir)

        notify("All resume PDFs generated")

    except Exception as e:
        print('Error:', e)
        notify("Error generating PDFs")

    finally:
        is_running = False

# ================= HOTKEY ================= #

notify("Resume Generator is running")
keyboard.add_hotkey("ctrl+q", on_hotkey)
keyboard.add_hotkey("ctrl+alt+p", convertToPDF)
keyboard.wait()
