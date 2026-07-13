# PDKS Paid Leave Calculator

> Built during my internship after watching HR staff spend hours every month doing the same Excel calculations by hand.

🔗 **Live:** [pdks-izin-hesaplayici.netlify.app](https://pdks-izin-hesaplayici.netlify.app)

---

## Background

At my internship, the HR team used Tempo PDKS to track attendance. Every month, someone had to manually go through the exported Excel report, calculate overtime hours for each employee, figure out who qualified for paid leave, and highlight them one by one.

I automated it. Upload the file, get the result.

---

## Why not connect directly to Tempo?

As an intern I didn't have access to the company's database or Tempo's internal API. The Excel export was the only data I could work with — so I built around that constraint.

---

## What it does

Takes a standard Tempo PDKS Excel export and returns a processed file with two sheets:

- **Sheet 1** — original data with eligible employees highlighted in green and leave entitlement calculated in column Q
- **Sheet 2** — summary of only the eligible employees, with totals

The leave calculation follows Turkish labor law:

```
if (OT_50% + OT_100%) > 30 hours:
    leave_days = round((total_OT - 30) × 1.5 / 9.5, 1)
```

---

## Stack

- **Backend** — Python, FastAPI, openpyxl
- **Frontend** — HTML, CSS, vanilla JS (single file, no framework needed)
- **Hosting** — Render (API) + Netlify (frontend)
- **Uptime** — UptimeRobot keeps the Render free tier awake

---

## Run locally

```bash
git clone https://github.com/berencelik05-cell/pdks-izin-hesaplayici.git
cd pdks-izin-hesaplayici/backend
pip install -r requirements.txt
uvicorn main:app --reload
```

Open `frontend/index.html` in your browser.

---

## Project structure

```
pdks-izin-hesaplayici/
├── backend/
│   ├── main.py           # FastAPI endpoints + Excel processing
│   └── requirements.txt
├── frontend/
│   └── index.html        # drag-and-drop UI
└── README.md
```

---

## What I'd do differently with more time

- Add support for other PDKS formats (not just Tempo)
- Auth layer so only company staff can access
- Store monthly reports instead of discarding after download
- Dockerize the backend

---

*3rd year CS student. Built this in my internship to solve a real problem.*
