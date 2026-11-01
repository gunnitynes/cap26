# Catalyst Operations Test: setup

Two parts:

- **The page** (`docs/`) is hosted on GitHub Pages. It holds no data.
- **The data** lives in one private Google Sheet on your account. A small Apps Script inside that Sheet is the only thing the page talks to, and it checks every teacher's password.

Your co-teachers need no Google or claude.ai account, only the link, their name and a password.

## 1. The Sheet and its script (about 5 minutes, once)

1. Create a new Google Sheet called **Catalyst Operations Test**.
2. **Extensions ▸ Apps Script.** Delete what is there and paste in all of `apps-script/Code.gs`, then save.
3. In the editor, choose the function **setup** and press **Run**. Allow the permissions it asks for: "this spreadsheet only", because of `@OnlyCurrentDoc`. The tabs Teachers, Students, Marks, Days and Results appear.
4. **Deploy ▸ New deployment ▸ Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**

   Press **Deploy** and copy the address that ends in `/exec`.
5. Back in the Sheet, run **Ops Test ▸ Set up / tidy the sheet** once. It builds and styles every tab, and adds a **Start here** tab that explains everything.

## Adding teachers

This is done in the **Teachers** tab. Anyone you share the Sheet with can do it.

- **Name and password:** type a name and a password of 6 or more characters. The password turns into "●●●●●● set" the moment you press Enter, and the teacher can sign in.
- **Name only:** the first time that teacher signs in on the page, they choose their own password, typed twice.
- **Forgotten password:** clear the dots. The teacher is signed out everywhere and chooses a new password next time.
- **Status column:** says whether each person is Ready, Waiting, or has a password that is too short.

"Anyone" access means anyone can reach the script. Without a valid name and password it answers nothing but "Please sign in again".

## 2. The page

1. Put the `/exec` address into `docs/config.js`:
   ```js
   window.OPS_API = 'https://script.google.com/macros/s/…/exec';
   ```
2. Publish the `docs/` folder with GitHub Pages, for example as `gunnitynes/cap26`. The link is then `https://gunnitynes.github.io/cap26/`.

## 3. Send it

Send each teacher the link plus their own name and password, ideally in person. On iPhone they should use **Share ▸ Add to Home Screen**, which opens it full screen like an app.

## Changing the script later

Saving the script is not enough. Use **Deploy ▸ Manage deployments ▸ ✎ Edit ▸ Version: New version ▸ Deploy**. The `/exec` address stays the same.

## Changing the questions

Teachers can do this on the page: **teacher menu ▸ My questions** edits their own tab (an admin gets **Edit questions** for any teacher's tab or the shared one). Reword, move, add or delete questions, set the Ask-out-loud line, Column name, Studio and stage, then **Save**. A save is refused if the same tab was changed in the Sheet in the meantime.

Or edit the **Questions** tab in the Sheet. Each row is one question, in test order:

| Column | What it does |
| --- | --- |
| Stage | The question's stage, picked from a dropdown. The stages are one list for every teacher, kept under **Test stages** on Start here: renaming one there renames it in every tab, and Grades has one column per stage. A new row left blank takes the stage above. |
| Stage heading | The heading on the page. Needed on one row of the stage. |
| Question | What the teacher marks. Delete the row to remove the question. |
| Ask out loud | Optional question for the teacher to read out. |
| Studio | Both, K3 or K4. |
| Column name | Optional short name: the Marks column heading. The same name in two teachers' tabs shares one column. Cleared when a different question is typed over the row. |
| ID (hidden) | Given by the sheet; keeps marks attached when a question is reworded or moved. |

Phones pick up the change within about 10 seconds. A graded student keeps the questions they were graded on, so grades already given never move (rewording still shows there); every student not graded yet gets the current questions straight away. **Ops Test ▸ Use the current questions for a test date** moves a marked date onto the current questions.

`OPS_SEED=1 node dev/server.mjs` adds example students (with times and marks) three and five days ahead, for trying the start day and the Grades screen.

Try it locally: `OPS_TRY_QUESTIONS=1 node dev/server.mjs` edits the tab 45 seconds after start, so an open page shows the change.

## If the Sheet can't be reached

Every mark is also kept on the teacher's phone, one copy per student. A save, a sign-out or the Sheet never clears that copy, and it is kept for 60 days. **Export my marks** in the teacher menu gives a spreadsheet file (CSV) of everything that teacher marked: the score, result, critical errors and notes, every question's mark, and an **In the Sheet** column saying *Yes*, *Not yet: waiting on this phone*, or *Missing from the Sheet: this is the phone copy*.

## What is where in the Sheet

| Tab | What it holds |
| --- | --- |
| Start here | The page link, how to add teachers, what each tab holds, how grades work. |
| Teachers | Name, password (shown as dots), role, status, last sign-in. The scrambled password sits in two hidden columns. |
| Students | One row per student and test date. Removing a student marks the row "Deleted" (with who and when), so nothing is lost. |
| Marks | One row per student per teacher: Result, Score % and one % column per test stage (the same columns for every teacher), then critical errors and notes. Every single skill is folded away on the right (press + above the Updated column). A hidden last column holds the page's own copy. |
| Days | Which test dates have their grades revealed. |
| Results | The tally. The page writes it whenever grades are revealed or downloaded. |

## Trying it on this Mac without deploying

```bash
node dev/server.mjs
```

Then open `http://localhost:8787/?api=/api` and sign in as Gunnar with the password `catalyst`. Anna has no password yet, so signing in as her tries out the first sign-in. This runs the same `Code.gs` with stand-ins for Google's services.

`node dev/test.mjs` runs the automatic checks for sign-in and the Sheet.
