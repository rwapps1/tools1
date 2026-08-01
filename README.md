# Report Builder

A standalone, hosted version of the Report Builder tool. Everything builds and
deploys automatically via GitHub Actions — you don't need Node, git, or a
terminal.

## Getting it live (drag-and-drop, ~5 minutes)

1. **Create the repository**
   - Go to https://github.com/new
   - Give it a name (e.g. `report-builder`)
   - Choose Public or Private (Private repos can still use free GitHub Pages)
   - Do **not** tick "Add a README" — leave it empty
   - Click **Create repository**

2. **Upload these files**
   - On the new (empty) repo page, click **"uploading an existing file"**
   - Drag the *entire contents* of this folder in (not the folder itself —
     the files and subfolders: `src`, `.github`, `index.html`, `package.json`,
     etc.) into the browser drop zone
   - Modern Chrome/Edge will preserve the folder structure (`src/...`,
     `.github/workflows/...`) as long as you drag folders, not just top-level
     files. If it flattens anything, upload the `src` folder and the
     `.github` folder as separate drag-and-drop steps to be safe.
   - Scroll down and click **Commit changes**

3. **Turn on GitHub Pages**
   - In your repo, go to **Settings → Pages**
   - Under "Build and deployment", set **Source** to **GitHub Actions**
   - That's it — no further settings needed

4. **Wait for the build**
   - Go to the **Actions** tab in your repo — you'll see a "Deploy to GitHub
     Pages" run kick off automatically (triggered by your upload)
   - It takes about 1–2 minutes. A green check means it worked
   - Your live URL will appear in the workflow run, and also under
     **Settings → Pages** once done. It'll look like:
     `https://<your-username>.github.io/<repo-name>/`

Any time you upload changed files again in future, it will rebuild and
redeploy automatically — no extra steps.

## Notes

- Saved presets are stored in the visitor's own browser (`localStorage`), not
  on a server — each person who uses the tool keeps their own presets
  locally, nothing is shared or synced between devices.
- All spreadsheet processing happens entirely in the browser — no files are
  uploaded anywhere.
