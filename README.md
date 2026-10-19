<div align="center">

<img src="frontend/public/logo.jpg" alt="AIFi logo" width="96" />

# AIFi: The AIML Club Website

### The official website and event registration platform for AIFi, the AI/ML club at REVA University, Bengaluru.

<p>
  <img src="https://img.shields.io/badge/React_19-20232A?style=for-the-badge&logo=react&logoColor=61DAFB" alt="React" />
  <img src="https://img.shields.io/badge/Vite-1B1B1F?style=for-the-badge&logo=vite&logoColor=646CFF" alt="Vite" />
  <img src="https://img.shields.io/badge/Framer_Motion-111111?style=for-the-badge&logo=framer&logoColor=white" alt="Framer Motion" />
  <img src="https://img.shields.io/badge/Node.js-1A1A1A?style=for-the-badge&logo=nodedotjs&logoColor=5FA04E" alt="Node.js" />
  <img src="https://img.shields.io/badge/Express-1E1E1E?style=for-the-badge&logo=express&logoColor=white" alt="Express" />
  <img src="https://img.shields.io/badge/ExcelJS-1D3B2A?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="ExcelJS" />
  <img src="https://img.shields.io/badge/Render-111827?style=for-the-badge&logo=render&logoColor=white" alt="Render" />
</p>

</div>

> **Fork notice:** this is a fork of [Bunny-95/ai-fi-website](https://github.com/Bunny-95/ai-fi-website), a team project for the AIFi club. The frontend was built by the club team. My part was the **registration backend** and connecting the frontend to it (see [My Contributions](#-my-contributions)).

---

## ✨ Overview

The site introduces **AIFi, the AIML Club** at REVA University: its mission, events, faculty and team. It also handles registrations for the club's national-level AI event series: a **24-hour AI Hackathon**, a **GFG Coding Contest** and an **AI Masterclass** (AiFi × GDG × OS Code). Registrations are sent to a small Express API that appends them to per-event Excel workbooks for the organisers.

## 🚀 Features

- **Animated landing site**: Hero, About, Events timeline, Faculty and Team sections built with Framer Motion.
- **Three registration flows**:
  - `/register/hackathon`: team registration for up to 4 members, with college, year and branch
  - `/register/gfg`: coding-contest registration
  - `/register/masterclass`: masterclass registration
- **Excel-backed registration API**: each event writes to its own `.xlsx` file (`contest.xlsx`, `hackathon.xlsx`, `masterclass.xlsx`) with fixed column schemas.
- **Auto-created workbooks**: the data directory and each workbook (with headers) are created on the first registration. If a workbook is open in Excel (`EBUSY`), the API returns a clear error.
- **Centralised error handling**: `AppError`, a `catchAsync` wrapper, a global error handler, and crash logging to `error.log`.

### API

| Method | Endpoint | Description |
| --- | --- | --- |
| `POST` | `/api/register-hackathon` | Register a hackathon team |
| `POST` | `/api/register-gfg` | Register for the GFG coding contest |
| `POST` | `/api/register-masterclass` | Register for the AI Masterclass |

## 🧱 Tech Stack

- **Frontend:** React 19, Vite, React Router, Framer Motion, Lucide icons
- **Backend:** Node.js, Express, ExcelJS, CORS
- **Hosting:** the backend is deployed on Render

## 🗂️ Project Structure

```
ai-fi-website/
├── frontend/
│   ├── public/                 # logos
│   └── src/
│       ├── components/         # Hero, About, Events, Faculty, Team, *Registration
│       └── context/            # AnimationContext
└── backend/
    ├── index.js                # Express app + error handling
    ├── routes/                 # registrationRoutes.js
    ├── controllers/            # registration + error controllers
    ├── services/excelService.js
    ├── config/constants.js     # file paths + Excel column schemas
    └── utils/                  # AppError, catchAsync
```

## ⚡ Getting Started

**Backend:**

```bash
cd backend
npm install
npm run dev        # node --watch index.js, listens on PORT (default 3000)
```

Registrations are written to `backend/data/*.xlsx` (created automatically).

**Frontend:**

```bash
cd frontend
npm install
# create frontend/.env and set VITE_API_URL to your backend URL, e.g. http://localhost:3000
npm run dev
```

**Environment variables:**

- Frontend: `VITE_API_URL`
- Backend: `PORT` (optional)

## 🙋 My Contributions

- Implemented the **Express registration backend**: routes, controllers, the ExcelJS storage service, error handling and constants.
- Fixed an ExcelJS row-insertion bug.
- Connected the React registration forms to the deployed backend through `VITE_API_URL`.

These changes were merged into the upstream repository through pull requests.

## 👤 Author

Original project by [Bunny-95](https://github.com/Bunny-95) and the AIFi team.
Backend by **Mohammed Shoaeb Malik** · AI & full-stack developer, Bengaluru · [@ShoaebMalik19](https://github.com/ShoaebMalik19)
