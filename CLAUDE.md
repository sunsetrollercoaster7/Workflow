# CV & cover letter workflow

This repo produces a tailored CV and cover letter (Word .docx) for each job the user pastes in.
Claude drafts; the user reviews and sends. Never submit or email an application.

**Privacy:** this GitHub repo is public. `master/` and `applications/` are gitignored. Never commit,
push or paste personal details (CVs, letters, contact info, job files) into tracked files.
Personal files are lost when the cloud session ends, so always send the finished .docx files to the user.

## Context
Australian job market (Victoria): community services, disability (NDIS), youth, family violence and
early childhood roles, mostly SCHADS Award Levels 1–3. Use Australian English spelling
(organise, behaviour, centre, programme only if the ad uses it) and Australian terms (WWCC, NDIS,
MARAM, CALD, Child Safe Standards, key selection criteria).

## Setup (once)
- Master CV lives at `master/master_cv.md` (plus optional `master/profile.md`, `master/cover_letter_notes.md`).
  If it's missing, ask the user to paste/upload it and save it there as Markdown.
- `pip install -r requirements.txt`

## For each job the user pastes (link or description)
1. If given a link, fetch it; if the page can't be read, ask for the description text.
2. Create `applications/YYYY-MM-DD_company_role/` (lowercase, underscores) and save `job.md`.
3. Write `cv.md`, tailored from the master CV:
   - Only use facts from `master/`. **Never invent** experience, skills, qualifications, dates or numbers.
   - Reorder and trim to what's relevant; mirror the job's keywords where they're truthfully supported.
   - Max 2 pages. Strong action verbs, quantified results where the master CV has them.
4. Write `cover_letter.md`: under one page, addressed to the hiring manager if named,
   specific to the company and role, links 2–3 achievements to the top requirements.
   - If the ad has **key selection criteria (KSC)**, also write `ksc.md` (built to
     `<Name>_Selection_Criteria_<Company>.docx`): one `## Criterion` heading each, answered with
     the STAR method (Situation, Task, Action, Result) in flowing prose, not labelled S/T/A/R,
     drawing on a different example per criterion where possible. Respect any word/page limit.
   - If there are no formal KSC, the cover letter still addresses the 3–4 key requirements of the
     job description, each backed by a brief STAR-style example.
5. Write `notes.md`: requirements matched, gaps or weak spots, anything the user should verify.
6. Build the Word files:
   `python3 scripts/build_docx.py applications/<folder>/cv.md "applications/<folder>/<Name>_CV_<Company>.docx"`
   `python3 scripts/build_docx.py applications/<folder>/cover_letter.md "applications/<folder>/<Name>_Cover_Letter_<Company>.docx" --letter`
7. Commit and push, then send the user the two .docx files and a short summary from `notes.md`.

## Writing style: must read as human-written
- Write like a thoughtful early-career practitioner, not a template. Vary sentence length; use plain words.
- Avoid AI tells: "I am excited/thrilled to", "passionate about", "delve", "leverage", "tapestry",
  "testament to", "navigate", "foster", "in today's fast-paced world", "I am confident that",
  "Furthermore/Moreover/Additionally" chains, rule-of-three lists everywhere, em-dash overuse,
  and generic closing lines. No em dashes in letters; use commas or full stops.
- Specific beats impressive: name the service, the client group, what she actually did, what changed.
- Keep the user's own voice and phrasing from the master CV where it works.
- Re-read the final text once purely to strip anything that sounds generic or machine-made.

## Markdown supported by build_docx.py
`# Name` (title), `## Section`, `### Role line`, `- bullet`, `**bold**`, `*italic*`,
plain paragraphs, and `---` for a horizontal rule.
