
# The Lot Ledger

A CSV-based inventory search tool. Runs entirely in the browser — no backend,
no API keys, no database.

## Run it locally first (recommended before deploying)

You'll need [Node.js](https://nodejs.org) installed (any recent version).

```bash
npm install
npm run dev
```

Open the URL it prints (usually `http://localhost:5173`). Try importing a
CSV and confirm everything works before deploying.

## Deploy to GitHub Pages

1. Push this whole folder to your GitHub repo (root of the repo, so
   `package.json` sits at the top level).
2. In the repo on GitHub: **Settings → Pages → Source → GitHub Actions**.
   (Not "Deploy from a branch" — this project uses the included Actions
   workflow instead.)
3. Push to `main`. The workflow in `.github/workflows/deploy.yml` will
   automatically build the app and publish it.
4. After the Action finishes (check the "Actions" tab), your app is live at:
   `https://<your-github-username>.github.io/<your-repo-name>/`

## If your repo name isn't "DealerInventoryScanner"

Open `vite.config.js` and change the `base` value to match your actual repo
name exactly (case-sensitive), e.g.:

```js
base: "/your-repo-name/",
```

If you ever move this to a *custom domain* instead of the default
`github.io` address, set `base: "/"` instead.

## Notes

- All data is session-only by design — nothing is saved between visits.
  Each time you open the site, import that day's CSV fresh.
- Use the in-app **Export to CSV** button to keep your own permanent copies.
