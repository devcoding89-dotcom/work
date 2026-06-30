
# Resume Tailor - Full Project (v7)

This version uses GPT to **identify** and **rewrite** only:
- Experience responsibility bullets (including fake bullets like "o ")
- Skills lines (including "Category: value" lines)

It does NOT modify:
- Objective/Summary/Profile
- Job titles, dates, headings, education, projects, certifications
- Company Summary/Overview/Description lines

DOCX safety:
- The code only edits text inside `<w:t>` nodes in `word/document.xml`.
- It preserves the DOCX XML structure, so Word can open the output.

## Run server
```bash
cd server
npm install
export OPENAI_API_KEY="sk-..."
npm start
```

Server runs on http://localhost:4000
