# PowerPoint Color Palette

A small, think-cell-style floating colour palette for **PowerPoint on Windows**, written in VBA.
Pick a mode — **Font / Fill / Outline** — then click a swatch to apply that colour to the
selected shape(s) or text. The swatches are read **live from the open presentation's theme**,
so they always match the deck you're working in.

![mode: font / fill / outline · live theme swatches · transparent · custom hex](https://img.shields.io/badge/PowerPoint-VBA%20add--in-orange)

## Features

- **Mode toggle** — Font, Fill, or Outline; the active mode is highlighted.
- **Live theme swatches** — the 10 theme colours of the active deck, each with 5 tint/shade
  variants (the same grid PowerPoint shows), plus the standard-colours row underneath.
  Reopen the palette on another deck and the swatches update automatically.
- **Transparent / No Fill / No Line** — one button, applied per the active mode.
- **Custom hex** — type `#RRGGBB` and **Set** to apply, or **+** to save it as a swatch.
  Custom swatches persist between sessions.
- **Modeless** — stays open while you keep editing slides.

## Why VBA (and not an Office.js add-in)

Office.js can't read a presentation's theme colour scheme, and this tool's whole point is to
mirror the open deck's palette. VBA reads it directly via
`ActivePresentation.SlideMaster.Theme.ThemeColorScheme`, so a classic VBA add-in is the right fit.
Windows only.

## Install (~2 minutes)

1. Open PowerPoint → `Alt`+`F11` (VBA editor).
   If macros aren't visible: **File → Options → Customize Ribbon → tick *Developer***.
2. **Insert → UserForm.** Press `F4`, set **(Name)** to `frmPalette`. Double-click the form,
   then paste the entire contents of [`src/frmPalette.frm.vba`](src/frmPalette.frm.vba).
3. **File → Import File…** and import [`src/clsSwatch.cls`](src/clsSwatch.cls) and
   [`src/ColorPalette.bas`](src/ColorPalette.bas).
4. Click inside the `ColorPalette` module and press `F5` (runs `ShowPalette`). The palette opens.

### Make it a permanent add-in

5. **File → Save As → PowerPoint Add-in (`*.ppam`)**, saved to the default AddIns folder.
6. **File → Options → Add-ins → Manage: PowerPoint Add-ins → Go… → Add New →** pick your `.ppam`.
7. Add a launch button: right-click the ribbon (or QAT) → **Customize** → *Choose commands from:
   Macros* → add **`ShowPalette`**.

## Usage notes

- Font mode with **text** selected recolours only that text; with a **shape** selected it
  recolours all text in the shape.
- *Transparent* applies in Fill and Outline modes only (font has no transparency).
- The palette is built in code rather than as a designed `.frm/.frx` form, so the repo stays
  diff-friendly and there's no binary form resource to corrupt.

## Files

| File | Role |
|------|------|
| `src/ColorPalette.bas` | Standard module: theme reading, tint/shade maths, apply logic, persistence |
| `src/clsSwatch.cls` | Class that wraps a dynamic label and routes its click |
| `src/frmPalette.frm.vba` | Code to paste into the blank `frmPalette` UserForm |

## License

MIT — see [`LICENSE`](LICENSE).
