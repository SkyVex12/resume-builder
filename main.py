import os
import re
import json
import contextlib
import io
from pathlib import Path
from datetime import datetime
from collections import Counter

import keyboard
import pyperclip
from winotify import Notification

from docx import Document
from docx2pdf import convert

from openai import OpenAI

# Optional: local similarity check (NO extra OpenAI calls)
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ================= CONFIG ================= #

MODEL = "gpt-4.1-mini"
TEMPERATURE_MAIN = 0.35
COMPANY_COUNT = 4

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

OUTPUT_DIR = "output"
BULLET = "  \u2022    "  # bullet point

MAX_SKILLS_PER_CATEGORY = 8

# similarity warning only (no regen)
SIMILARITY_WARN_THRESHOLD = 0.70

# ================= DIVERSITY LANES (NO EXTRA REQUESTS) ================= #
# Disjoint opening verbs per person (biggest impact)
OPENING_VERBS = {
    "Timothy": ["Architected", "Defined", "Established", "Standardized", "Governed"],
    "Wilfredo": ["Implemented", "Optimized", "Refactored", "Automated", "Instrumented"],
    "Lou": ["Stabilized", "Hardened", "Coordinated", "Improved", "Streamlined"],
    "Ryan": ["Built", "Extended", "Enhanced", "Supported", "Integrated"],
    "James": ["Drove", "Accelerated", "Increased", "Delivered", "Reduced"],
}

# Different “metric families” per person
METRIC_LANES = {
    "Timothy": "latency, uptime, scalability, availability, reliability",
    "Wilfredo": "throughput, build time, runtime efficiency, cost per request, deployment frequency",
    "Lou": "MTTR, incident rate, test coverage, defect rate, operational stability",
    "Ryan": "feature adoption, support tickets, time-to-ship, engagement, usage growth",
    "James": "revenue impact, conversion, retention, cost reduction, performance gains"
}

# Different bullet framing styles per person (structure lanes)
BULLET_STYLE = {
    "Timothy": "architecture trade-offs + scalable design decision + measurable outcome",
    "Wilfredo": "implementation detail + optimization change + measurable performance/cost outcome",
    "Lou": "process/reliability improvement + cross-team collaboration + measurable stability outcome",
    "Ryan": "hands-on build/iteration + learning/ownership + measurable delivery/adoption outcome",
    "James": "business outcome + execution leadership + measurable ROI/performance outcome"
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


import re
from typing import List, Set

# You can expand these safely (they only match if found verbatim in the JD)
SOFT_SKILL_PHRASE_LIBRARY = [
     # Interpersonal / People skills
    "interpersonal skills",
    "strong interpersonal skills",
    "excellent interpersonal skills",
    "people skills",
    "relationship building",
    "build relationships",
    "build strong relationships",

    # Organization / Work style
    "highly organized",
    "strong organizational skills",
    "organizational skills",
    "well organized",
    "organized and detail-oriented",
    "detail-oriented",
    "detail oriented",

    # Communication effectiveness
    "communicate effectively",
    "effective communicator",
    "effective communication",
    "communicate clearly",
    "clear communicator",

    # Personality / Work ethic (Jobscan-safe)
    "perfectionist",
    "high attention to detail",
    "attention to detail",
    "results-oriented",
    "results driven",
    "self-starter",
    "self motivated",
    "highly motivated",
    "proactive mindset",
    "strong work ethic",

    # Collaboration & Cross-functional
    "cross-functional team",
    "cross-functional teams",
    "cross-functional collaboration",
    "cross functional collaboration",
    "cross-functional partnership",
    "collaborate closely",
    "collaborate effectively",
    "collaborate with stakeholders",
    "collaborate with product",
    "collaborate with design",
    "collaborate with engineering",
    "collaborate with leadership",
    "work closely with",
    "partner closely with",
    "partner with stakeholders",
    "work with stakeholders",
    "work across teams",
    "work across functions",
    "partner across teams",
    "coordinate across teams",
    "coordinate with stakeholders",
    "collaboration with",
    "build consensus",
    "align with stakeholders",
    "alignment with stakeholders",
    "cross-team alignment",
    "cross functional alignment",

    # Communication
    "strong communication skills",
    "excellent communication skills",
    "clear communication",
    "effective communication",
    "verbal communication",
    "written communication",
    "written and verbal communication",
    "verbal and written communication",
    "clear technical communication",
    "technical communication",
    "communicate complex concepts",
    "communicate technical concepts",
    "communicate clearly",
    "present to stakeholders",
    "present technical concepts",
    "stakeholder communication",
    "stakeholder management",
    "customer communication",
    "client communication",
    "document technical decisions",
    "produce technical documentation",
    "technical documentation",
    "create documentation",
    "write documentation",

    # Ownership / Accountability
    "end-to-end ownership",
    "ownership of projects",
    "take ownership",
    "strong ownership",
    "high ownership",
    "accountability",
    "sense of ownership",
    "responsible for delivery",
    "deliver outcomes",
    "deliver results",
    "drive delivery",
    "drive execution",
    "drive projects",
    "deliver high-quality software",
    "deliver production-ready software",
    "operate production systems",
    "production support",
    "on-call rotation",
    "incident response",
    "escalated issues",
    "level 3 support",
    "provide support",
    "support escalations",
    "support and maintenance",

    # Leadership / Influence
    "technical leadership",
    "lead technical initiatives",
    "lead cross-functional efforts",
    "lead projects",
    "lead engineering efforts",
    "influence technical direction",
    "influence architecture",
    "influence product roadmap",
    "drive technical strategy",
    "provide technical guidance",
    "set engineering standards",
    "set best practices",
    "establish best practices",
    "mentor junior engineers",
    "mentoring junior engineers",
    "mentor engineers",
    "coaching engineers",
    "coach engineers",
    "pair programming",
    "code review",
    "provide feedback",

    # Problem Solving / Analytical
    "problem-solving skills",
    "strong problem-solving skills",
    "analytical thinking",
    "critical thinking",
    "troubleshoot issues",
    "troubleshoot complex issues",
    "debug production issues",
    "deep technical analysis",
    "root cause analysis",
    "investigate issues",
    "identify issues",
    "resolve issues",
    "resolve defects",
    "debugging and troubleshooting",
    "diagnose issues",
    "data-driven decision making",
    "make data-driven decisions",

    # Process / Execution / Delivery
    "agile development",
    "agile development practices",
    "agile methodologies",
    "scrum",
    "kanban",
    "agile/scrum",
    "iterative development",
    "continuous improvement",
    "process improvement",
    "improve processes",
    "engineering best practices",
    "best engineering practices",
    "software engineering best practices",
    "continuous delivery",
    "continuous integration",
    "prioritize work",
    "prioritization",
    "manage priorities",
    "work under deadlines",
    "deliver on time",
    "fast-paced environment",
    "adapt to change",
    "adapt to changing priorities",
    "respond to feedback",
    "stakeholder feedback",
    "retrospective meeting",
    "continuous learning",
    "learning mindset",
    "growth mindset",

    # Quality / Reliability / Security Mindset
    "attention to detail",
    "high attention to detail",
    "commitment to quality",
    "focus on quality",
    "high-quality code",
    "code quality",
    "testing mindset",
    "quality assurance",
    "reliability",
    "operational excellence",
    "security best practices",
    "privacy and security",
    "compliance requirements",
    "risk management",
    "performance optimization",
    "optimize performance",

    # Customer / Product / UX
    "customer-focused mindset",
    "customer focused mindset",
    "customer-centric",
    "customer centric",
    "user-focused",
    "user focused",
    "user experience",
    "focus on user experience",
    "empathy for the user",
    "empathy for users",
    "high level of empathy",
    "product mindset",
    "business requirements",
    "translate business requirements",
    "understand business requirements",
    "translate requirements into",
    "requirements gathering",
    "work with product",
    "work with design",
    "stakeholder requirements",

    # Autonomy / Organization
    "work independently",
    "independent work",
    "self-motivated",
    "self motivated",
    "self-directed",
    "self directed",
    "manage multiple priorities",
    "multitask",
    "time management",
    "organizational skills",
    "organized",
    "detail-oriented",
    "detail oriented",
    "perfectionist",
    "client satisfaction"
]

HARD_SKILL_KEYWORDS = {
    # Core domains / “hard skill phrases” often in JDs
    "front-end development", "front end development",
    "back-end development", "back end development",
    "full stack", "full-stack", "full stack development", "full-stack development",
    "cloud infrastructure", "infrastructure", "information systems",
    "web services", "software development lifecycle", "sdlc","software development life cycle",
    "system design", "architecture", "cloud architecture",
    "deployment strategies", "containerization", "orchestration",

    # Programming languages
    "python", "java", "javascript", "typescript", 
    "c++", "c#", "go", "golang", "ruby", "php",
    "scala", "kotlin", "swift", "rust",

    # Frontend
    "react", "react.js", "next.js", "nextjs",
    "vue", "vue.js", "angular", "angularjs",
    "html", "css", "responsive design", "web browser fundamentals",

    # Backend frameworks
    "fastapi", "flask", "django",
    "node.js", "nodejs", "express", "nestjs",
    "spring", "spring boot",

    # APIs & integration
    "rest", "restful", "rest api", "restful api", "rest apis", "restful apis",
    "soap",
    "graphql",
    "websocket", "websockets",
    "api design", "api-based solutions",

    # Databases
    "postgres", "postgresql", "mysql", "mariadb",
    "mongodb", "redis", "sqlite",
    "dynamodb", "cassandra", "elasticsearch",
    "sql", "nosql",

    # Cloud providers
    "aws", "amazon web services",
    "azure", "microsoft azure",
    "gcp", "google cloud", "google cloud platform",

    # DevOps / CI/CD
    "devops", "ci/cd", "continuous integration", "continuous delivery",
    "jenkins", "github actions", "gitlab ci", "circleci",
    "docker", "kubernetes",
    "terraform", "cloudformation", "pulumi",
    "helm", "nginx",
    "linux",

    # Auth / Security
    "authentication", "authorization",
    "oauth", "oauth2",
    "jwt", "sso", "single sign-on",
    "secrets management",
    "security best practices", "compliance",

    # Testing
    "automated testing", "unit testing", "integration testing", "end-to-end testing", "e2e testing",
    "tdd", "test-driven development",
    "pytest", "unittest", "nose",
    "jest", "cypress", "playwright", "mocha", "junit",
    "postman", "swagger", "openapi",

    # Monitoring / Observability
    "monitoring", "logging", "alerting", "observability",
    "datadog", "prometheus", "grafana",
    "sentry",

    # Performance / scalability
    "performance optimization", "scalability", "high-throughput", "high throughput",
    "caching", "batching",

    # AI / ML / LLM / RAG
    "ai", "artificial intelligence",
    "machine learning", "ml",
    "deep learning",
    "llm", "llms", "large language models",
    "gpt-4", "gpt-4o", "claude", "gemini",
    "rag", "retrieval-augmented generation", "retrieval augmented generation",
    "prompt engineering",
    "embeddings", "embedding",
    "semantic search", "similarity search", "hybrid search",
    "vector database", "vector databases",
    "pinecone", "weaviate", "chroma", "faiss", "qdrant",
    "langchain", "llamaindex",
    "evaluation frameworks", "model evaluation",

    # Architecture patterns
    "event-driven architecture", "event driven architecture",
    "microservices", "distributed systems",

    # Tools
    "git", "jira", "confluence",
}

HARD_SKILL_KEYWORDS.update({
    # Programming paradigms / concepts
    "object-oriented programming",
    "object oriented programming",
    "object-oriented programming languages",
    "oop",
    "functional programming",
    "procedural programming",

    # Documentation / SDLC artifacts
    "project documentation",
    "technical documentation",
    "software documentation",
    "system documentation",
    "design documentation",
    "architecture documentation",
    "api documentation",
    "requirements documentation",

    # Design & architecture concepts
    "software architecture",
    "application architecture",
    "system architecture",
    "design patterns",
    "architectural patterns",

    # Engineering practices
    "code reviews",
    "version control",
    "source control",
    "build automation",
    "release management",

    # Analysis & design
    "requirements analysis",
    "technical design",
    "data model design",
    "database design",

    # Enterprise / systems language Jobscan likes
    "information technology",
    "information systems",
    "enterprise systems",
})

# ================= EXTRACT SKILLS ================= #
import re
from typing import Dict, List, Set

def extract_skills_with_counts(
    jd_text: str,
    soft_phrase_library: List[str],
    hard_skill_keywords: Set[str],
    *,
    min_hard_len: int = 3,
) -> Dict[str, Dict[str, int]]:
    """
    Clean extraction + counts:
    - HARD: counts only meaningful keywords/phrases (prevents 'c' matching everywhere)
    - SOFT: counts ONLY phrases from soft_phrase_library (verbatim match, hyphen/space tolerant)
    Returns:
      {
        "hard_skills": {"java": 2, "ci/cd": 2, ...},
        "soft_skills": {"highly organized": 1, "detail-oriented": 1, ...}
      }
    """

    jd = jd_text
    jd_norm = jd.lower().replace("–", "-").replace("—", "-")

    def norm(s: str) -> str:
        s = s.strip().lower().replace("–", "-").replace("—", "-")
        s = re.sub(r"\s+", " ", s)
        return s

    def flex_pattern(phrase: str) -> str:
        # spaces/hyphens interchangeable
        p = norm(phrase)
        esc = re.escape(p)
        esc = esc.replace(r"\ ", r"[-\s]+").replace(r"\-", r"[-\s]+")
        return esc

    def count_phrase(phrase: str) -> int:
        patt = flex_pattern(phrase)
        return len(list(re.finditer(patt, jd_norm, flags=re.IGNORECASE)))

    # ---------------- HARD SKILLS ----------------
    hard_counts: Dict[str, int] = {}

    for kw in sorted({norm(x) for x in hard_skill_keywords if x and x.strip()}, key=len, reverse=True):
        # Filter out noisy short tokens like "c", "r", "go" (optional)
        if len(kw) < min_hard_len and kw not in {"go"}:
            continue

        # Use word boundaries for short-ish clean words to avoid substring hits
        if re.fullmatch(r"[a-z0-9]+", kw):
            patt = r"\b" + re.escape(kw) + r"\b"
            c = len(list(re.finditer(patt, jd_norm, flags=re.IGNORECASE)))
        else:
            c = count_phrase(kw)

        if c:
            hard_counts[kw] = c

    # ---------------- SOFT SKILLS ----------------
    soft_counts: Dict[str, int] = {}

    for phrase in sorted({norm(p) for p in soft_phrase_library if p and p.strip()}, key=len, reverse=True):
        c = count_phrase(phrase)
        if c:
            soft_counts[phrase] = c

    # Sort by frequency desc
    hard_counts = dict(sorted(hard_counts.items(), key=lambda x: -x[1]))
    soft_counts = dict(sorted(soft_counts.items(), key=lambda x: -x[1]))

    return {"hard_skills": hard_counts, "soft_skills": soft_counts}

# ================= EXTRACT SKILLS ================= #

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


# ================= ATS KEYWORD EXTRACTION (CHEAP + EFFECTIVE) ================= #

STOPWORDS = {
    "the","and","or","to","of","in","a","an","for","with","on","as","is","are","be","by","at","from","that",
    "this","it","you","we","our","your","will","can","must","have","has","had","into","across","using","use",
    "ability","experience","years","year","strong","work","working","build","develop","design", ".",
}

def normalize_text(t: str) -> str:
    t = t.lower()
    t = re.sub(r"[^a-z0-9+\-./\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def extract_top_phrases(jd_text: str, top_n: int = 18) -> list[str]:
    """
    Extract high-signal 1–3 word phrases from the JD.
    This is NOT perfect NLP — it’s a cheap ATS anchor list.
    """
    jd = normalize_text(jd_text)
    tokens = [w for w in jd.split() if w not in STOPWORDS and len(w) > 2]

    # include some common tech patterns
    # keep original-cased variants later by using the JD itself in generation
    ngrams = []
    for n in (1, 2, 3):
        for i in range(len(tokens) - n + 1):
            phrase = " ".join(tokens[i:i+n])
            # avoid useless phrases
            if any(p in STOPWORDS for p in phrase.split()):
                continue
            ngrams.append(phrase)

    counts = Counter(ngrams)

    # Boost common ATS-important terms if present
    boosts = ["rag", "retrieval augmented generation", "llm", "fastapi", "react", "next.js",
              "kubernetes", "docker", "terraform", "postgresql", "mongodb", "redis", "graphql",
              "ci/cd", "oauth2", "jwt", "websocket", "datadog", "prometheus", "grafana"]
    for b in boosts:
        if b in jd:
            counts[b] += 6

    # Filter too-generic
    filtered = [(p, c) for p, c in counts.items() if c >= 2 or len(p.split()) >= 2]
    filtered.sort(key=lambda x: x[1], reverse=True)

    top = []
    for p, _ in filtered:
        if p not in top:
            top.append(p)
        if len(top) >= top_n:
            break
    return top


# ================= PROMPT ================= #

def build_prompt_all(jd_text: str, ats_phrases: list[str]) -> str:
    phrases_block = ", ".join(ats_phrases) if ats_phrases else "None"
    skills = extract_skills_with_counts(jd_text, SOFT_SKILL_PHRASE_LIBRARY, HARD_SKILL_KEYWORDS)
    print(skills)
    return f"""
You are an ATS-optimized resume generation engine.
Output valid JSON ONLY. No markdown. No explanations.

TASK:
Generate FIVE DISTINCT senior-level resumes for the SAME Job Description.
The resumes must be ATS-friendly, professional, and written for human reviewers.

CRITICAL ATS REQUIREMENT (NON-NEGOTIABLE):
You MUST guarantee that ATS scores are over 95% for ALL five resumes against the provided JD.

ATS KEYWORD ANCHORS (MANDATORY):
{phrases_block}

You MUST:
- Find ALL keywords and phrases from the JD text below
- Use ALL of them verbatim in EACH resume
- Especially include ALL hard skills and soft skills from the JD
- Use EVERY skill more times per resume than its appearance count in the JD

These are the extracted skills with required minimum appearance counts:
{json.dumps(skills, indent=2)}

These skills MUST appear verbatim in:
- Summary
- Skills
- Experience
And MUST exceed their JD frequency PER RESUME.

━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ BULLET UNIQUENESS ENFORCEMENT (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━

You are generating bullets across FIVE people.
You must internally track the structure, verbs, and order of phrases used in each bullet.

✅ You must enforce these hard rules:
- NO two bullets in ANY resume may begin with the same 4 words.
- NO two bullets may use the same verb-object pair (e.g., "Built APIs", "Developed systems").
- NO two bullets may use the same sentence pattern or clause order (e.g., "[Verb] using [Tech] to achieve [Result]").
- Every bullet must be a complete rewrite — do not swap nouns, verbs, or metrics and call it "different".
- Structure, rhythm, and phrasing must be fully unique across resumes.
- Use different sentence starters, clause positions, and focus (outcome, method, metric, collaboration).

━━━━━━━━━━━━━━━━━━━━━━━━━━
STRUCTURAL VARIATION (REQUIRED IN EACH RESUME)
━━━━━━━━━━━━━━━━━━━━━━━━━━

Each resume must contain:
- At least 4 bullets that begin with the metric (e.g., “Reduced X by…”)
- At least 4 that begin with the technology (e.g., “Using Java and…”)
- At least 4 that begin with collaboration (e.g., “In partnership with…”)
- At least 4 that begin with the outcome (e.g., “Enabled seamless…”)
- At least 2 that begin with a non-standard form (e.g., “For a system with 1M+ users…”)

Do not group them — spread these types randomly through the experience section.
Each resume must use a different number and order of these types.

━━━━━━━━━━━━━━━━━━━━━━━━━━
GLOBAL RULES
━━━━━━━━━━━━━━━━━━━━━━━━━━
- Use EXACT keywords and phrases from the Job Description verbatim.
- Repeat critical JD keywords naturally across Summary, Skills, and Experience.
- All resumes align to the SAME role, but wording and structure MUST differ.
- Do NOT reuse bullet sentence patterns across people.

━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT FORMAT (STRICT JSON, NO TRAILING COMMAS)
━━━━━━━━━━━━━━━━━━━━━━━━━━
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
  "Wilfredo": {{ "...same structure..." }},
  "Lou": {{ "...same structure..." }},
  "Ryan": {{ "...same structure..." }},
  "James": {{ "...same structure..." }}
}}

━━━━━━━━━━━━━━━━━━━━━━━━━━
SUMMARY RULES (EACH PERSON)
━━━━━━━━━━━━━━━━━━━━━━━━━━
- Start with "Senior Software Engineer"
- Include the EXACT job title from the JD
- 4–5 concise lines
- Senior-level tone only
- Include 6–8 JD keywords or phrases verbatim
- Include at least 2 soft skills verbatim

━━━━━━━━━━━━━━━━━━━━━━━━━━
SKILLS RULES (EACH PERSON)
━━━━━━━━━━━━━━━━━━━━━━━━━━
- Categorize skills strictly under the provided categories
- Use JD keywords verbatim
- Include 6–10 skills per category
- Ensure ALL hard and soft skills from the JD are represented

━━━━━━━━━━━━━━━━━━━━━━━━━━
EXPERIENCE RULES (EACH PERSON)
━━━━━━━━━━━━━━━━━━━━━━━━━━
- company_1 and company_2: 7–8 bullets each
- company_3: 6 bullets
- company_4: 5 bullets

━━━━━━━━━━━━━━━━━━━━━━━━━━
BULLET QUALITY RULES (NON-NEGOTIABLE)
━━━━━━━━━━━━━━━━━━━━━━━━━━
EACH bullet MUST include ALL of the following:
- At least ONE JD hard skill (verbatim)
- At least ONE JD soft skill (verbatim)
- At least ONE collaboration / communication phrase from the JD
- At least ONE measurable metric (%, $, latency, users, throughput, uptime, cost)

━━━━━━━━━━━━━━━━━━━━━━━━━━
PERSON-SPECIFIC EMPHASIS (DO NOT MENTION NAMES)
━━━━━━━━━━━━━━━━━━━━━━━━━━
- Timothy: architecture decisions, scalability, system design, long-term maintainability
- Wilfredo: implementation depth, optimization, delivery speed, performance tuning
- Lou: reliability, process improvement, cross-team coordination, stability
- Ryan: hands-on development, iteration, learning, incremental improvements
- James: business impact, execution ownership, performance gains, cost or revenue outcomes

━━━━━━━━━━━━━━━━━━━━━━━━━━
FINAL HARD STOP
━━━━━━━━━━━━━━━━━━━━━━━━━━
DO NOT generate output unless ALL of the following are true:
- ALL skills are used verbatim more than their JD frequency PER RESUME
- ALL experience bullets are completely unique in syntax and structure
- ATS relevance is guaranteed above 95% for EACH resume
- Output is 100% valid JSON with no extra text

JOB DESCRIPTION:
{jd_text}

Return ONLY valid JSON.
""".strip()


# ================= OPENAI CALL ================= #

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

def generate_all_one_call(jd: str) -> dict:
    ats_phrases = extract_top_phrases(jd, top_n=18)
    prompt = build_prompt_all(jd, ats_phrases)
    response = client.responses.create(
        model=MODEL,
        input=prompt,
        temperature=TEMPERATURE_MAIN,
        timeout=240
    )
    print('response:', response)
    text = extract_text(response)
    data = safe_parse_json(text)
    if not isinstance(data, dict):
        raise ValueError("Invalid JSON from model.")
    return data


# ================= LOCAL SIMILARITY WARNING (NO EXTRA API) ================= #

def collect_all_bullets(all_data: dict) -> tuple[list[str], list[str]]:
    bullets, owners = [], []
    for person in PERSON_ORDER:
        exp = all_data.get(person, {}).get("experience", {})
        for i in range(1, COMPANY_COUNT + 1):
            for b in exp.get(f"company_{i}", []) or []:
                if isinstance(b, str) and b.strip():
                    bullets.append(b.strip())
                    owners.append(person)
    return bullets, owners

def similarity_warning(all_data: dict, threshold: float = SIMILARITY_WARN_THRESHOLD) -> list[tuple[str,str,float]]:
    bullets, owners = collect_all_bullets(all_data)
    if len(bullets) < 2:
        return []

    vec = TfidfVectorizer(ngram_range=(1,2), stop_words="english")
    X = vec.fit_transform(bullets)
    sim = cosine_similarity(X)

    warnings = []
    n = len(bullets)
    for i in range(n):
        for j in range(i+1, n):
            if owners[i] != owners[j] and sim[i, j] >= threshold:
                warnings.append((owners[i], owners[j], float(sim[i, j])))

    warnings.sort(key=lambda x: x[2], reverse=True)
    return warnings[:8]


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
    # Destroys runs/formatting; use only for plain placeholders
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

        notify("Generating resumes (one API call)...")
        all_data = generate_all_one_call(jd)

        # local similarity warning only (no regen)
        warns = similarity_warning(all_data, SIMILARITY_WARN_THRESHOLD)
        if warns:
            notify("Warning: some bullets still similar (see console).")
            print("\n[SIMILARITY WARNINGS] (personA, personB, score)")
            for a, b, s in warns:
                print(a, b, f"{s:.2f}")
        else:
            print("[OK] Similarity check passed.")

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

        notify("Converting DOCX to PDF...")
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
