# 🚀 ATS Resume Generator (Multi-Person, Jobscan-Optimized)

An **ATS-optimized resume generation system** that produces **multiple senior-level resumes** from a single Job Description using OpenAI GPT models.

Designed to consistently reach **90–95%+ ATS scores (Jobscan-style)** while maintaining **person-specific differentiation**, strict JSON safety, and automated document generation.

---

## ✨ Key Features

- 🔥 **Single Job Description → Multiple Distinct Resumes**
- 🎯 **ATS-First Design (Jobscan-Optimized)**
- 🧠 **Strict Keyword Fidelity (No Paraphrasing)**
- 🧩 **Soft Skills Extraction & Enforcement**
- 📄 **DOCX Resume Generation (Template-Based)**
- 🔔 **Windows Toast Notifications**
- ⌨️ **Hotkey-Triggered (Ctrl + Q)**
- 📂 **Auto-Open Output Directory (Once)**
- 🛡️ **Strict JSON Validation (No Hallucinated Formats)**

---

## 🏗️ Project Structure

```
.
├── main.py
├── templates/
│   ├── Timothy.docx
│   ├── Wilfredo.docx
│   ├── Lou.docx
│   └── Ryan.docx
├── output/
│   └── JD_YYYYMMDD_HHMMSS/
│       ├── 1_Timothy_resume.docx
│       ├── 2_Wilfredo_resume.docx
│       ├── 3_Lou_resume.docx
│       └── 4_Ryan_resume.docx
└── README.md
```

---

## 🧠 How It Works

1. Copy a Job Description
2. Press **Ctrl + Q**
3. The system:
   - Extracts hard & soft skills
   - Builds a strict ATS prompt
   - Generates 4 distinct senior resumes
   - Saves each as `.docx`
   - Opens output folder (once)
   - Shows toast notifications

---

## 👤 Multi-Person Differentiation

Each resume uses a unique style profile to prevent duplication and ATS penalties.

All resumes:
- Remain **senior-level**
- Share the same JD
- Use different metrics and wording

---

## 🧩 ATS Optimization Strategy

### Hard Skills
- Copied verbatim from the Job Description
- Repeated across Summary, Skills, Experience

### Soft Skills
- Extracted via behavioral signal detection
- Enforced verbatim in experience bullets

### Experience Rules
Each bullet:
- Starts with an action verb
- Includes 1 hard skill, 1 soft skill, and 1 metric

---

## ⚙️ Requirements

- Python 3.10+
- Windows OS
- OpenAI API key

### Install Dependencies

```bash
pip install openai python-docx keyboard pyperclip winotify docx2pdf scikit-learn
```

---

## ▶️ Usage

```bash
python main.py
```

Copy a Job Description → Press **Ctrl + Q**

---

## 🏆 Result

✔ 90–95%+ ATS resumes  
✔ Fully automated  
✔ Deterministic & scalable  
