# MercerAssess

A Mercer-branded assessment, certification, and survey analytics dashboard, built with React + Vite + Tailwind CSS, with shared data storage via Upstash Redis (connected through Vercel's Marketplace integration).

## What this is

The UI is a fully client-side React app. Data (participants, surveys, grading criteria) is stored in **Upstash Redis** — connected via Vercel's Marketplace integration, since Vercel KV itself was deprecated — through small serverless API routes (`/api/data.js`, `/api/survey-response.js`, `/api/attendance.js`, `/api/send-reminder.js`), so every device that opens the deployed site reads and writes the *same* shared data.

### Participant accounts

Participants now create real accounts and sign in with their own email + password, instead of a shared mock login. How it works:

- An admin adds a participant's profile first (Admin → Participants), which includes their real email — this is the "employee record."
- The participant then clicks **Create one** on the login page, enters that same email plus a password they choose, and their account is created — tied to their existing profile, not a brand-new one.
- Passwords are hashed (SHA-256 with a per-account random salt) using the browser's built-in Web Crypto API before being stored — never in plain text. **This is still not production-grade security**: there's no email verification, no rate limiting on login attempts, and no server-side session/cookie hardening. Treat it as meaningfully better than plaintext, not as a bank-grade auth system.
- **On concurrent edits**: participant data is stored as one shared list under a single key, read-then-written as a whole. Signup and login specifically re-fetch the freshest copy right before checking/saving (rather than trusting a possibly-stale in-memory copy), which closes most of the realistic race window — e.g., two different people signing up moments apart on different devices. It isn't a fully atomic guarantee, though: two admins editing *different* participants at literally the same instant could still have one save overwrite the other's. A fully race-proof version would store each participant as its own record (e.g. a Redis hash with per-field atomic writes) rather than one shared list — a reasonable next step if this sees real concurrent multi-admin usage.
- Login sessions persist per-device via the browser's own `localStorage` (separate from the shared KV data), so refreshing the page keeps you signed in on that device.
- Admin → Participants → edit any participant to see whether they've created their account yet, and reset their password if they're locked out (forces them to create it again with a new one).
- Admin sign-in is unchanged — a separate, simpler mock flow, not tied to the new participant account system.

### Meetings, calendar, reminders, and attendance QR codes

In Admin → **Meetings**, schedule a meeting (title, date, start/end time, location, description, and audience — All Participants or a specific cohort). Meetings are grouped into **Upcoming** and **Past** agenda sections. Each meeting has four actions:
- **Remind** — sends a real email reminder (see setup below) to every participant in that meeting's audience, using their personal email on file. Subject and message are pre-filled from the meeting details and fully editable before sending.
- **Present** — opens a full-screen, high-contrast QR display meant to be projected on a screen for a room to scan (separate from the compact QR in the Share modal).
- **Share** — the compact QR code + link, same as before.
- **View check-ins** — the live roster of who's checked in, matched against real registered participants (see below).

Anyone who scans a meeting's QR code lands on a **"Sign in to check in"** page: they select their name from the actual registered participant list (matching their real employee ID, cohort, and department automatically), or choose "Not on this list…" for a manual entry, which is flagged **"Unmatched"** in the admin view. If someone who's already checked in tries to check in again (e.g. re-scanning the same QR code), they're shown "You're already checked in" instead of creating a duplicate record.

### Attendance Report

Admin → **Attendance Report** aggregates check-ins across every meeting into one view — each participant's attendance rate (meetings attended ÷ meetings that applied to them, based on cohort/audience matching), color-coded (green ≥80%, amber 50–79%, red <50%), plus their last-attended date. Sort by lowest rate first to surface who's falling behind, or by name. This is separate from the per-meeting check-in list under **Meetings**, which still shows who attended one specific session.

#### Setting up real email reminders (Resend)

The Remind button sends actual emails via [Resend](https://resend.com):

1. Sign up at resend.com (free, no credit card — 3,000 emails/month, 100/day)
2. Dashboard → API Keys → Create API Key, copy it
3. In your Vercel project → Settings → Environment Variables, add `RESEND_API_KEY` with that value, then redeploy
4. Optional: once you've verified your own sending domain in Resend (Settings → Domains), also add a `RESEND_FROM` environment variable like `"Your Name <you@yourdomain.com>"`. Until then, it sends from Resend's shared test address (`onboarding@resend.dev`), which works but is more likely to land in spam for real recipients.

Without `RESEND_API_KEY` set, clicking Remind shows a clear error explaining it's missing, rather than silently failing. Participants with a missing or malformed email address are automatically skipped (with a visible count) rather than failing the whole batch.

### Bulk importing participants and scores from Excel

Admin → Participants → **Import from Excel** accepts a real-world tracker file — even a messy one with multiple sheets, multi-row headers, and inconsistent column names. It scans every sheet, scores each by how many recognizable columns it finds (name, email, employee ID, batch/cohort, a total score), and uses the best match automatically. You get a preview table before anything is saved: each row shows whether it'll create a new participant or update an existing one (matched by email), with checkboxes to exclude any you don't want.

Every other column in the sheet — attendance, milestones, coaching sessions, per-component scores, whatever your tracker has — is also detected and shown as a **Grading Criteria Columns** list with checkboxes, so you can bring in the full breakdown per participant, not just one overall number. Duplicate headers (like two columns both called "ILM Score") are automatically disambiguated. This breakdown is stored per participant and shown read-only when editing them. Re-importing merges with any criteria a participant already has rather than wiping them, so importing a second file with different columns doesn't erase the first import's data.

Cohort/audience dropdowns throughout the app (Participants filter, meeting audience, survey audience) are now derived from whatever cohort values actually exist in your data, rather than a fixed "Cohort A/B/C" list — so imported batches (e.g. "Batch 1", "Batch 2") are immediately usable for targeting meetings and surveys, not just cosmetic labels.

### Separate participant lists and grading criteria per project

Admin → Participants has a project selector at the top, in a highlighted box that always shows which project you're currently viewing. Each project keeps its own participant roster, its own Leaderboard, and — new — its **own grading criteria**. When a project has zero participants (e.g. right after creating a new one), a clear message explains that specifically, rather than it looking like data disappeared — this replaces an earlier, more confusing version of this feature that got reverted for exactly that reason. Import from Excel also respects the current project, matching existing participants by email only within that project (so the same email in two different projects doesn't get mixed up).

A **Delete Project** button appears next to New Project whenever you're not viewing Default Project (which can't be deleted, since it's the fallback everything relies on). Deleting a project asks for confirmation and never silently discards participant data — everyone in that project is moved to Default Project first, so their records (including any imported grades) stay intact, just no longer grouped under the deleted project.

### Participant Meetings tab

Logged-in participants now have their own **Meetings** tab in the nav — a simplified, view-only version of the admin's Meetings tab. It shows only their upcoming meetings (matching their cohort/audience), each with a **Show QR Code** button that pulls up the same full-screen QR display admins use to project onto a room, so a participant can show their own device's screen to check in without needing an admin present.

### Colorful rating scale and per-project client logos

Survey rating questions (both the anonymous QR flow and the in-app Surveys tab) now use a proper color-coded Likert scale — green for Strongly Agree down to red for Strongly Disagree, each option's badge colored even before it's selected — instead of plain gray circles.

Each project can also have its own **client logo**: in Admin → Participants, next to the project selector, click **Add Client Logo** to upload an image for the currently-selected project. New surveys and meetings created while that project is selected use its logo (instead of the Mercer logo) on the survey-taking page, the check-in page, and the full-screen Present view. For surveys created *before* a project logo was set, open **Share Survey** and use the **Set Logo** control there to apply one directly to that specific survey. Meetings already pick this up automatically — editing and re-saving an existing meeting (Meetings → click it → Save Changes) applies the project's current logo if one wasn't already set.

### Grading criteria are per-project, editable, and importable

Admin → Grading Criteria now edits *the current project's* criteria (not one global list), saves live as you type, and has an **Import from Excel** button that reuses the same column-detection as participant import — every column in your tracker becomes a candidate criterion with an even starting weight you can adjust before confirming. Different projects can have completely different grading structures (e.g. one project might weight Assessment 50% / Attendance 20% / ALPs 25% / Modules 5%, while another uses a totally different rubric).

One criterion can be marked **"AI-graded from reflections"** (a radio button next to each row) — when a participant's reflection is graded by Gemini, that score is saved into whichever criterion is marked this way, contributing its configured weight to the *overall* grade. The reflection is one input among several, not the whole score: a participant's real total is the weighted sum across every criterion in their project (attendance, ALPs, modules, the AI-graded reflection, whatever else you've set up), computed live — not a stored number that goes stale. Criteria scores that don't parse as numbers (like "Y"/"not due" tracking values) are preserved and shown for reference but don't factor into the weighted math.

### Participant detail view

Clicking a participant's name (not just the edit pencil) in Admin → Participants opens a detail view: their full grading-criteria breakdown for their project, each criterion's weight and weighted contribution, the real computed overall score and pass/fail status, and any other tracked (non-numeric) data from an import that isn't part of the formal criteria. An "Edit Participant" button jumps from there into the existing edit form.

If you do need to run two genuinely separate programs with zero shared infrastructure (not just separate data within one app), the more robust path is two separate deployments (two Vercel projects, each with its own Redis database) rather than this in-app project switcher — that gives real isolation. Happy to help set that up if you get there.

### Nav is split by role

Participants (logged in via their own account) see **Dashboard, My Profile, Surveys** — nothing admin-facing. **Survey Analytics** and **Admin** are now admin-session-only, not shown to regular participants at all (previously they were visible to everyone, which meant any participant could technically click into the Admin panel — this closes that gap).

### Inviting multiple cohorts to one meeting or survey

Audience selection (for both meetings and surveys) is now a checklist instead of a single dropdown — check "Batch 1" and "Batch 2" together to invite both to the same meeting, instead of needing to create a separate one for each. Checking "All Participants" supersedes specific cohort picks (they're mutually exclusive). This also fixed a real bug: the "Build Survey" flow had an audience field that was never actually wired up to save — every built survey silently defaulted to "All Participants" no matter what you picked. It's now properly connected, and existing single-cohort audiences from before this change still work (read as a one-item list).

### Certificates actually download now

The "Download PDF" button on a participant's certificate previously had no action wired to it at all — it was a static mockup. It now generates a real, branded **image** (PNG) certificate — drawn on an in-memory canvas, needing no external library — with the participant's name and program, and downloads it immediately. This uses Canvas instead of a PDF library specifically so it works identically in every environment (including this chat's preview), rather than depending on a library that isn't universally available. It now also draws the real Mercer logo (a header mark plus subtle corner and center watermarks) instead of just text, and no longer shows the score on the image itself — just the name, program, issue date, and a certificate ID.

### Attendance check-in now requires signing in as yourself

Previously, scanning a meeting's QR code let anyone pick *any* registered participant's name from a dropdown and check them in — meaning someone physically present could check in on behalf of someone who wasn't there. Check-in now requires actually signing in with your own account (email + password) first; once signed in, you can only confirm attendance for yourself, with no way to select anyone else. This does mean a participant needs an account/password set before they can check in this way — if someone hasn't signed up yet, they'll need to do that first.

### Downloadable survey analysis

Survey Analytics now has a **Download Analysis (CSV)** button next to Share/Send, exporting each question's average score and response count as a CSV file, generated client-side.

### Groups

Admin → **Groups** (per-project) lets you organize participants into working groups — e.g. for group coaching sessions — independent of cohort. Create a group, expand it, and check off which participants belong to it.

### Resources (documents & videos, with completion tracking, viewable in-app)

Admin → **Resources** (per-project) lets you either **upload a file directly** (Word, PDF, PowerPoint, Excel, or a video file) or paste a link, for participants to review. Uploaded files are stored via Vercel Blob storage and get a real hosted URL automatically — no need to host it elsewhere yourself first. Participants see these under their own **Resources** tab; clicking one opens it **inside the app** in a viewer modal rather than just linking out — Word/PowerPoint/Excel files render through Microsoft's public Office viewer, video files play natively, and YouTube/Vimeo links get embedded players. Anything that can't be previewed inline (e.g. some Google Drive/SharePoint links) falls back to an explicit "open in new tab" button rather than a silent dead link.

**PDF viewing note**: Vercel Blob deliberately sets `X-Frame-Options: DENY` on every file it stores (a security measure to prevent it hosting embeddable content), which would otherwise make a directly-embedded PDF iframe silently render blank. To work around this, PDFs are fetched client-side first and shown via a local, same-origin blob URL instead of embedding the Vercel Blob URL directly — this sidesteps the restriction since no cross-origin frame request happens at that point. If a PDF still fails to load inline for some reason, the viewer shows a clear message and an "open in new tab" button rather than a blank frame.

Participants can mark each resource as completed — admin sees a running completion count per resource across the project's participants.

**Requires Blob storage connected**: like Redis, this needs a one-time setup — Vercel dashboard → your project → **Storage** tab → **Marketplace Database Providers** → **Blob** → connect to this project. That automatically injects the `BLOB_READ_WRITE_TOKEN` environment variable the upload route needs. Redeploy after connecting it for the first time.

### LinkedIn-shareable certification badge

Next to each certificate's PDF download button, there's now a **LinkedIn** button that opens LinkedIn's official "Add to Profile" flow for certifications, pre-filled with the program name, Mercer as the issuing organization, and the certificate ID — adding it to the participant's real Licenses & Certifications section on LinkedIn, rather than just sharing a static image.

### Email verification on signup

Previously, signup never confirmed the person actually controlled the email they typed — anyone could create an account using any email address, including someone else's. New signups are now marked unverified until they click a real emailed verification link. Verified status gates the highest-stakes action specifically: meeting check-in is blocked (with a clear message) until the account is verified. A banner on the Dashboard shows unverified accounts a "Resend verification email" option. Existing accounts created before this feature aren't affected — they're treated as verified automatically so no one gets suddenly locked out.

### "Remember me" now actually works

The checkbox on the sign-in page previously did nothing at all regardless of whether it was checked. It now controls where the session is stored: checked persists across browser restarts (localStorage), unchecked clears as soon as the tab or browser closes (sessionStorage) — the standard meaning of "remember me."

### Sort participants by score

Admin → Participants has a new "Sort" dropdown — Name, Highest Score, or Lowest Score — using the same weighted-total calculation as the leaderboard and certificates.

### Search in Groups and Resources

Both panels now have a search box once there's more than a couple of items, with a distinct "no matches for X" state versus the original "nothing here yet" empty state.

### Admin login is now actually secured

Previously, clicking the "Admin" tab on the sign-in page and submitting *any* email and password — even blank — granted full admin access. This was leftover mock behavior from early in the build. Admin accounts are now real, password-verified accounts (same hashing system as participants), with their own "Forgot password?" flow. The demo admin account still works: `admin.grader@mercer.com` / `demo1234`.

New **Admin → Admin Accounts** tab lets a signed-in admin add or remove other admin accounts — there's no public "sign up as admin" path, only an already-authenticated admin can grant access to someone else. The last remaining admin account can't be removed, so it's not possible to lock everyone out.

### Forgot password

The "Forgot password?" link on the sign-in page previously did nothing — it was decorative text. It now opens a real form: enter your email, get a reset link by email (same Resend infrastructure as meeting reminders), click it, set a new password, and you're signed in. Links expire after 1 hour. For privacy, the request always shows the same generic confirmation message regardless of whether the email actually matched an account.

### Meetings and Surveys are now project-scoped

Previously only Participants, Leaderboard, Groups, and Resources respected the selected project — Meetings and Surveys showed everything globally, mixed across projects. Both are now tagged with their project at creation and filtered everywhere: the admin panels, the participant-facing Meetings/Surveys tabs, notifications, and the Attendance Report.

### Admin Overview tab

A new **Overview** tab (now the default landing tab in Admin) shows a project's key numbers at a glance: participant count, average score, pass rate, upcoming meetings, survey count, resource completion rate, and group coverage — without needing to click through every other tab first.

### Resource completion reminders

Admin → Resources now has a **Remind** button per resource, sending a real email only to participants who haven't marked that specific resource as completed yet — same pattern as meeting and survey reminders.

### Downloadable survey slides (PowerPoint)

Survey Analytics and the Compare Surveys view both have a **Download Slides (PPTX)** button next to the CSV export — generates a real, branded PowerPoint deck client-side (via `pptxgenjs`), with the same navy-to-teal gradient title/closing slides used elsewhere in the app, native editable charts (not baked-in images) for question scores and sentiment breakdown, and a key-themes slide from open-text responses when available. The comparison version adds an overall-score delta slide and a side-by-side matched-question chart. Both open and edit normally in PowerPoint, Keynote, or Google Slides.

### Full roster export

Admin → Participants now has an **Export Roster (CSV)** button, exporting every participant in the current project — name, email, cohort, department, status, group, every configured grading criterion's score, and the computed weighted total — as a single CSV file, generated client-side.

### Which sheet gets used when multiple sheets could match

When a tracker file has more than one sheet that plausibly contains grading data (e.g. a detailed per-milestone tracking sheet *and* a clean results summary sheet), the importer now prefers whichever sheet has the more precise "total score" column — a sheet with an exact "Total out of 100%" column wins over one that only loosely contains the word "score" elsewhere (like a per-module score column), even if both sheets otherwise look equally matched. Previously, a tie was broken by whichever sheet happened to be read first, which could silently pick the wrong one.

### Sending a survey to participants

Built surveys now have a **Send to Participants** button (next to Share Survey) that emails everyone in the survey's audience — a real email via Resend, same infrastructure as meeting reminders — letting them know a new survey is ready and pointing them to the in-app Surveys tab. Subject/message are pre-filled and editable before sending, and participants with a missing/invalid email are automatically skipped with a visible count.

### If newly-created data doesn't show up for other people/sessions

Every save (creating a survey, meeting, or account) now surfaces a visible toast if it fails to persist to the shared database, instead of only logging to the browser console. If you ever create something and it doesn't show up elsewhere, check for that toast — it almost always means the database isn't connected yet (see the Deploying on Vercel section above).

### Submitting a reflection (paste text or upload a file)

On the Dashboard, participants now actually submit their own reflection instead of grading a fixed example. Two ways to do it: paste/type text directly, or upload a `.docx` or `.pdf` file — uploaded files are parsed server-side (Word via `mammoth`, PDF via `pdf-parse`, no client-side library weight or worker configuration needed) into editable text before submitting. Once submitted, the AI Reflection Grading panel grades *that* real text instead of the old hardcoded sample, and participants can update their reflection later (which re-grades against the new text).

### AI reflection grading (Gemini)

The **AI Reflection Grading** panel on the Dashboard now calls a real AI model instead of showing mock scores. Click **Grade with AI** to have Google's Gemini API evaluate the submitted reflection against your criteria and weights, returning a real score and specific feedback for each one.

Setup:

1. Go to **aistudio.google.com**, sign in, and click **Get API key → Create API key** (free, no credit card)
2. In Vercel → Settings → Environment Variables, add `GEMINI_API_KEY` with that value, then redeploy

Without `GEMINI_API_KEY` set, clicking "Grade with AI" shows a clear error explaining it's missing, rather than silently failing. Until you click that button, the panel shows an explicit "Not yet graded — showing placeholder scores" badge rather than pretending the mock numbers are real.

### Taking surveys from within the app

Participants no longer need a QR code to answer a built survey — a new **Surveys** tab in the nav (visible once logged in) shows every survey available to them (matching their cohort or "All Participants"), with a **Start Survey** button. Rating questions show the full labeled scale (Strongly Disagree → Strongly Agree), not just bare numbers — same as the anonymous QR flow. Unlike the anonymous QR link, responses submitted this way are tagged with the participant's identity, so a survey they've already completed shows "Completed" instead of letting them answer twice.

### Real notifications

The bell icon in the top nav now shows genuine, derived notifications rather than a decorative dot: unanswered surveys matching the participant's audience, and any meetings in the next 7 days that apply to them. The badge count reflects how many items are pending; clicking a survey notification jumps straight to the Surveys tab.

### QR codes are generated entirely client-side

Both the survey and attendance QR codes are generated locally in the browser using the `qrcode` npm package — no call to any third-party QR image service. This is more reliable (works offline, isn't dependent on a third-party service being reachable) than the alternative of loading a QR image from a public API.

### Building and sharing your own surveys

In Survey Analytics → **Add Survey → Build Survey**, you can author your own questions (rating 1–5, or open text) instead of only uploading an existing file. Once created, click **Share Survey** on it to get:
- A QR code anyone can scan to open the survey — no login required
- A copyable link to the same page

Responses collect live in the shared database as people submit them, and the survey's charts update automatically with real data — same analysis engine as the Excel upload feature, just fed by live responses instead of a spreadsheet.

### Comparing surveys

Click **Compare Surveys** (next to Add Survey) to view any two surveys side by side — response counts, average ratings, sentiment donut charts, and per-question breakdowns for each. If the two surveys share matching questions (e.g. a pre-program and post-program run of the same survey), it automatically detects the overlap and renders a **Pre & Post Impact** table: each matching statement's before/after % who Agree or Strongly Agree, color-coded increase/decrease/no-change, plus a summary count — the built-in "Pre Impact Survey" and "Post Impact Survey" sample data demonstrate this.

## Local development

```bash
npm install
npm run dev
```

Note: the storage API route (`/api/data.js`) only runs on Vercel, not in plain `vite dev`. For local development, run `vercel dev` instead (see below) if you want storage to work locally too. Otherwise the UI will still load, but saving/loading data will fail with a network error until it's deployed.

## Deploying on Vercel

1. **Push this project to a GitHub repo** (see the git commands below).
2. **Import the repo into Vercel** (vercel.com → Add New → Project). Leave the framework preset as **Vite** — build command `npm run build`, output directory `dist`.
3. **Before your first successful deploy with working storage**, connect a Redis database. Vercel KV itself was deprecated and removed — the current path is through the Marketplace:
   - In your Vercel project, go to the **Storage** tab
   - Look for **Marketplace Database Providers** (or a similar prompt to browse the Marketplace) and choose **Upstash** → **Redis**
   - Follow the prompts — you can either let Vercel manage a new Upstash account for you, or connect an existing one. Either way, once connected, Vercel automatically injects the required environment variables (`KV_REST_API_URL`, `KV_REST_API_TOKEN`) for you. You don't need to copy/paste any keys yourself.
4. **Redeploy** (Deployments tab → click the three dots on the latest deployment → Redeploy) so the new environment variables take effect.

That's it — from then on, every visitor to your deployed URL shares the same participant list, surveys, grading criteria, etc.

If storage calls fail after this (e.g. "Failed to save"), the two things to check are: (1) the Redis integration in Storage actually shows as connected to *this* project, and (2) you redeployed after connecting it — connecting it doesn't retroactively apply to an already-running deployment.

## Pushing to GitHub for the first time

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

## Project structure

```
├── index.html          # HTML entry point
├── package.json        # dependencies
├── vite.config.js      # build tool config
├── tailwind.config.js  # Tailwind setup
├── postcss.config.js   # required by Tailwind
├── api/
│   ├── data.js              # serverless API route — shared storage backed by Upstash Redis
│   ├── survey-response.js  # serverless API route — collects real responses to built surveys
│   ├── attendance.js        # serverless API route — collects attendance check-ins
│   ├── send-reminder.js    # serverless API route — sends meeting reminder emails via Resend
│   ├── grade-reflection.js # serverless API route — grades reflections via Gemini
│   └── extract-text.js     # serverless API route — extracts text from uploaded .docx/.pdf reflections
├── src/
│   ├── main.jsx        # React entry point
│   ├── index.css       # Tailwind directives
│   └── App.jsx         # the entire app (all components, all data)
```

## A note on the `xlsx` dependency

`package.json` points at SheetJS's own CDN (`cdn.sheetjs.com`) instead of the plain npm registry for the `xlsx` package — that's intentional and matches SheetJS's own official installation instructions, not a mistake. If `npm install` ever complains about it, check SheetJS's install docs for the current recommended version URL.

## Known limitations of this version

- No real authentication — any email/password combination logs you in. Data is shared across everyone who visits the link, but there's no per-user login separation.
- "Sending" a survey or "grading" a reflection with AI still doesn't call any real email or AI service — those remain mocked.
- QR codes are generated by fetching an image from a free public API (`api.qrserver.com`) rather than a bundled library — this keeps the app dependency-free, but means QR generation depends on that third-party service being reachable. If you'd rather not depend on it, swap in an npm QR library (e.g. `qrcode`) instead.
- Survey responses are stored per-survey using Redis's list operations (`rpush`/`lrange`), so concurrent submissions from many people at once don't overwrite each other — but there's no de-duplication, so someone could submit the same survey multiple times from the same device.
- If you publish this publicly, remember the participant data currently baked in as defaults is sample/mock data — update it before sharing the link widely, since anyone with the link can see and edit the shared data.
