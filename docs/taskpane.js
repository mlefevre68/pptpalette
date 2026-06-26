/* Color Palette - Office.js task pane (PowerPoint, Mac + Windows + web) */

let MODE = "FILL"; // FONT | FILL | LINE

// Edit these freely - they are your default swatches.
const PALETTE = [
  "#002060", "#FFFFFF", "#44546A", "#E7E6E6", "#4472C4",
  "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"
];
const STANDARD = [
  "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050",
  "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"
];

const LS_KEY = "ppColorPalette.custom";

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("Open this in PowerPoint.", true);
    return;
  }
  document.getElementById("mFONT").onclick = () => setMode("FONT");
  document.getElementById("mFILL").onclick = () => setMode("FILL");
  document.getElementById("mLINE").onclick = () => setMode("LINE");
  document.getElementById("noneBtn").onclick = () => applyColor(null, true);
  document.getElementById("hexSet").onclick = onHexSet;
  document.getElementById("hexAdd").onclick = onHexAdd;

  renderRow("themeRow", PALETTE);
  renderRow("standardRow", STANDARD);
  renderRow("customRow", loadCustom());

  setMode("FILL");
  setStatus("Select a shape or text, then click a colour.");
});

function setMode(m) {
  MODE = m;
  ["FONT", "FILL", "LINE"].forEach((k) => {
    document.getElementById("m" + k).classList.toggle("active", k === m);
  });
}

function setStatus(msg, isErr) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.classList.toggle("err", !!isErr);
}

function renderRow(id, colors) {
  const row = document.getElementById(id);
  row.innerHTML = "";
  colors.forEach((hex) => {
    const b = document.createElement("button");
    b.className = "swatch";
    b.style.background = hex;
    b.title = hex;
    b.onclick = () => applyColor(hex, false);
    row.appendChild(b);
  });
}

/* ---------- custom swatch persistence ---------- */
function loadCustom() {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || []; }
  catch (e) { return []; }
}
function saveCustom(arr) {
  try { localStorage.setItem(LS_KEY, JSON.stringify(arr)); } catch (e) {}
}

function normalizeHex(v) {
  if (!v) return null;
  v = v.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]{6}$/.test(v)) return null;
  return "#" + v.toUpperCase();
}

function onHexSet() {
  const hex = normalizeHex(document.getElementById("hexInput").value);
  if (!hex) { setStatus("Enter a colour like #FE5715.", true); return; }
  applyColor(hex, false);
}
function onHexAdd() {
  const hex = normalizeHex(document.getElementById("hexInput").value);
  if (!hex) { setStatus("Enter a colour like #FE5715.", true); return; }
  const arr = loadCustom();
  if (!arr.includes(hex)) { arr.push(hex); saveCustom(arr); renderRow("customRow", arr); }
  applyColor(hex, false);
}

/* ---------- apply to current selection ---------- */
async function applyColor(hex, clear) {
  try {
    await PowerPoint.run(async (ctx) => {
      const shapes = ctx.presentation.getSelectedShapes();
      shapes.load("items");
      await ctx.sync();

      // FONT mode: prefer a live text selection, else recolour shape text.
      if (MODE === "FONT") {
        if (clear) { setStatus("Transparent doesn't apply to font.", true); return; }
        let handled = false;
        try {
          const tr = ctx.presentation.getSelectedTextRange();
          await ctx.sync();
          tr.font.color = hex;
          await ctx.sync();
          handled = true;
        } catch (e) { handled = false; }
        if (!handled) {
          if (!shapes.items.length) { setStatus("Select a shape or text first.", true); return; }
          shapes.items.forEach((sh) => { try { sh.textFrame.textRange.font.color = hex; } catch (e) {} });
          await ctx.sync();
        }
        setStatus("Applied font colour " + (hex || ""));
        return;
      }

      if (!shapes.items.length) { setStatus("Select a shape first.", true); return; }

      shapes.items.forEach((sh) => {
        if (MODE === "FILL") {
          if (clear) sh.fill.clear();
          else sh.fill.setSolidColor(hex);
        } else if (MODE === "LINE") {
          if (clear) sh.lineFormat.visible = false;
          else { sh.lineFormat.visible = true; sh.lineFormat.color = hex; }
        }
      });
      await ctx.sync();
      setStatus((clear ? "Cleared " : "Applied ") + MODE.toLowerCase() + (hex ? " " + hex : ""));
    });
  } catch (err) {
    setStatus("Couldn't apply: " + (err && err.message ? err.message : err), true);
  }
}
