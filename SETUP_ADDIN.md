# Color Palette — Office.js add-in (Mac + Windows + web)

A cross-platform task-pane version of the palette. Same idea as the VBA one —
**Font / Fill / Outline** toggle, click a swatch to apply to the selection,
**Transparent / No Fill / No Line**, and **custom hex** swatches that persist —
but it runs on PowerPoint for **Mac, Windows, and the web** from one codebase.

> Note: Office.js can't read the open deck's theme colours, so the top "Palette"
> row is a **defined, editable** set (open `docs/taskpane.js` and edit the
> `PALETTE` / `STANDARD` arrays — e.g. drop in your exact EDF hex codes).

## How it's hosted

The panel's web files (`docs/`) are served free over HTTPS by **GitHub Pages**.
The `manifest.xml` (which you load into PowerPoint) points at that Pages URL.

### Step 1 — Put your GitHub username in the manifest

`manifest.xml` has a placeholder `GITHUB_USER` in several URLs. Replace every
occurrence with your GitHub username. On Mac/Linux:

```bash
sed -i '' 's/GITHUB_USER/yourname/g' manifest.xml   # macOS
# sed -i 's/GITHUB_USER/yourname/g' manifest.xml     # Linux
```

(Or just open the file and Find-and-Replace `GITHUB_USER` → `yourname`.)

### Step 2 — Push and enable Pages

```bash
git add -A && git commit -m "Add Office.js add-in"
git push
```

On GitHub: **Settings → Pages →** Source: *Deploy from a branch*, Branch:
`main`, Folder: `/docs` → Save. After a minute, confirm the panel loads at
`https://yourname.github.io/powerpoint-color-palette/index.html`.

### Step 3 — Sideload the manifest

**Mac PowerPoint**
1. Copy `manifest.xml` into:
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
   (create the `wef` folder if it doesn't exist).
2. Restart PowerPoint → **Home tab → Color Palette** (or **Insert → Add-ins →
   My Add-ins**).

**Windows PowerPoint**
1. Put `manifest.xml` in a folder, then **share that folder** (right-click →
   Properties → Sharing) and copy its network path (`\\PC\share`).
2. PowerPoint → **File → Options → Trust Center → Trust Center Settings →
   Trusted Add-in Catalogs** → paste the path → **Add catalog** → tick
   *Show in Menu* → OK, then restart PowerPoint.
3. **Insert → My Add-ins → Shared Folder →** select **Color Palette**.

**PowerPoint on the web** (quickest test): **Insert → Add-ins → Upload My
Add-in →** choose `manifest.xml`.

## Using it

- Click **Font**, **Fill**, or **Outline** to choose what gets coloured.
- Select shape(s) or text on a slide, then click a swatch.
- **Transparent / No Fill / No Line** clears fill or outline (per the active mode).
- Type `#RRGGBB` then **Set** to apply, or **+** to save it as a custom swatch.

## Editing the default palette

Open `docs/taskpane.js` and edit the two arrays near the top:

```js
const PALETTE  = ["#002060", "#FFFFFF", ...];  // your brand row
const STANDARD = ["#C00000", "#FF0000", ...];  // standard colours
```

Commit and push; the change is live after Pages rebuilds.
