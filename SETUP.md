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
5. Go back to the Sheet and reload it. A menu **Ops Test** appears. Use **Ops Test ▸ Add or reset a teacher…**:
   - add yourself first and answer **Yes** to Admin;
   - then add each co-teacher with a password of 6 or more characters.

   You can also add teachers later from the page itself: tap your name ▸ Teachers.

"Anyone" access means anyone can reach the script. Without a valid name and password it answers nothing but "Please sign in again".

## 2. The page

1. Put the `/exec` address into `docs/config.js`:
   ```js
   window.OPS_API = 'https://script.google.com/macros/s/…/exec';
   ```
2. Publish the `docs/` folder with GitHub Pages, for example as `gunnitynes/ops-test`. The link is then `https://gunnitynes.github.io/ops-test/`.

## 3. Send it

Send each teacher the link plus their own name and password, ideally in person. On iPhone they should use **Share ▸ Add to Home Screen**, which opens it full screen like an app.

## Changing the script later

Saving the script is not enough. Use **Deploy ▸ Manage deployments ▸ ✎ Edit ▸ Version: New version ▸ Deploy**. The `/exec` address stays the same.

## What is where in the Sheet

| Tab | What it holds |
| --- | --- |
| Teachers | Name, role, and the password as a salted hash. The password itself is never stored. |
| Students | One row per student and test date. Removing a student marks the row "Deleted" (with who and when), so nothing is lost. |
| Marks | One row per student per teacher, as JSON. |
| Days | Which test dates have their grades revealed. |
| Results | The tally. The page writes it whenever grades are revealed or downloaded. |

## Trying it on this Mac without deploying

```bash
node dev/server.mjs
```

Then open `http://localhost:8787/?api=/api` and sign in as Gunnar with the password `catalyst`. This runs the same `Code.gs` with stand-ins for Google's services.
