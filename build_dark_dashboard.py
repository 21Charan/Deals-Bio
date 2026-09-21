#!/usr/bin/env python3
"""
Dark dashboard — the full Employee dashboard (Directory, Skill Atlas, Capability Risk,
Pulse, Team Analytics, Rate Analysis) re-skinned to match the Scroll Story landing page.

    reads   ../../03_Output files/Employee_Dashboard.html   (built by the root generator)
    returns the re-skinned HTML as a string — it writes nothing

This is a module, not a script: build_scroll_story.py calls dark_html() and embeds the result,
so the whole practice site is one file. The root dashboard is only ever read.

    cd "02_Scripts & ETL" && python generate_report.py          # the usual monthly build
    cd "../08_Scroll Story/02_Scripts & ETL" && python build_scroll_story.py

What it changes, and nothing else:
  1. A dark PwC skin appended after the dashboard's own <style>. Colours stay inside the
     PwC palette — orange, tangerine, yellow, rose, pink, burgundy and greys; only the
     greys are re-pointed for a dark background.
  2. The chart palette constants that are dark by design (black, dark grey) are swapped
     for their light counterparts so series stay visible on a dark canvas.
  3. The dashboard's own landing/overview is removed — the Scroll Story is the landing
     page — so it opens on the Directory tab, and the header gets an Overview button
     that closes the panel.
"""
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
EXPERIMENT = HERE.parent
ROOT = EXPERIMENT.parent
SOURCE = ROOT / "03_Output files" / "Employee_Dashboard.html"
# The dashboard is embedded in the landing page, so "Overview" closes the panel rather than
# navigating anywhere. If it is ever opened on its own the button simply does nothing.
BACK_BTN = ('<button type="button" class="ds-back" '
            "onclick=\"if(window.parent!==window)window.parent.postMessage('ds-close','*')\">"
            "&larr; Overview</button>")

# Chart series that are dark by design in the light dashboard -> light equivalents.
SERIES_SWAP = {"BLACK='#2D2D2D'": "BLACK='#D8D8DE'", "GREY_D='#464646'": "GREY_D='#9A9AA2'"}

# Pulse's gender KPI is two independent utilization figures, but "50% : 79%" sits a few cards away
# from headcount ratios written the same way ("9 : 9 male : female") and reads like one. The maths
# is the dashboard's and is untouched; only how it is written changes.
TEXT_SWAP = {
    "pct(mU) + ' : ' + pct(fU)": "'M ' + pct(mU) + '  \u00b7  F ' + pct(fU)",
    'Utilization &middot; M : F': 'Utilization &middot; men vs women',
}

SKIN = r"""
/* ============ Dark PwC skin — matches the Scroll Story landing page ============ */
:root{
  /* PwC accents unchanged; the greys are re-pointed for a dark canvas */
  --pwc-grey-900:#F5F5F4; --pwc-grey-700:#D4D4D8; --pwc-grey-600:#A5A5AC; --pwc-grey-400:#7E7E86;
  --pwc-grey-300:rgba(255,255,255,.16); --pwc-grey-100:rgba(255,255,255,.07); --pwc-grey-50:rgba(255,255,255,.035);
  --white:#16161A; --pwc-black:#0B0B0D;
  --ds-line:rgba(255,255,255,.09); --ds-line2:rgba(255,255,255,.18);
  --ds-bg:#0B0B0D; --ds-card:#16161A; --ds-tx:#F5F5F4; --ds-mu:#A5A5AC; --ds-mu2:#72727A;
  --serif:Georgia,'Times New Roman',serif;
  color-scheme:dark;
}
body{background:var(--ds-bg);color:var(--ds-tx)}
h1,h2,h3,.ov-h,.section-title{font-family:var(--serif);font-weight:400;letter-spacing:-.015em}

/* header + tab bar */
header.brand{background:#111114;border-bottom:1px solid var(--ds-line);padding:16px 40px}
header.brand,header.brand *{color:var(--ds-tx)}
header.brand .logo-img{background:#fff}
.brand-home{pointer-events:none}
.ds-back{margin-left:auto;display:inline-flex;align-items:center;gap:8px;font:600 13px/1 'Helvetica Neue',Arial,sans-serif;
  color:#fff;background:var(--pwc-orange);border:0;cursor:pointer;border-radius:999px;padding:11px 18px;text-decoration:none;pointer-events:auto;transition:transform .2s,background .2s}
.ds-back:hover{transform:translateY(-1px);background:#ff6a2b}
nav.tabs{background:rgba(17,17,20,.86);-webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);
  border-bottom:1px solid var(--ds-line);box-shadow:none}
nav.tabs button{color:var(--ds-mu)}
nav.tabs button:hover{color:var(--ds-tx)}
nav.tabs button.active{color:var(--ds-tx)}

/* surfaces */
.stat-card,.dir-card,.bio-card,.comp-card,.ana-card,.modal,.insight-callout,.mdi-block,.util-chart,
.ms-panel,.sc-sel,.atlas-view-toggle,.exp-filter,.cr-card,.card,.panel{
  border:1px solid var(--ds-line);box-shadow:none}
.stat-card,.dir-card,.comp-card,.ana-card,.insight-callout,.mdi-block,.util-chart{border-radius:16px}
.modal,.bio-card{border-radius:20px}
table{border-color:var(--ds-line)}
th,td{border-color:var(--ds-line)!important}
thead th{background:#121215;color:var(--ds-mu2)}
tbody tr:hover{background:rgba(255,255,255,.03)}
hr{border-color:var(--ds-line)}

/* form controls */
input,select,textarea{background:rgba(255,255,255,.05)!important;border:1px solid var(--ds-line2)!important;color:var(--ds-tx)!important}
input::placeholder{color:var(--ds-mu2)}
select option{background:#16161A;color:var(--ds-tx)}
.controls .reset-btn{background:rgba(255,255,255,.05)!important;color:var(--ds-tx)!important;border-color:var(--ds-line2)!important}
.ms-trigger{background:rgba(255,255,255,.05)!important;color:var(--ds-tx)!important}
.ms-panel{background:#16161A!important}
input[type=range]{accent-color:var(--pwc-orange)}
/* the text-input rule above boxes every <input>; a slider is drawn by its own track, so no box */
input[type=range]{background:transparent!important;border:0!important;box-shadow:none!important}
/* Team Analytics: the people count sits at the toolbar's right end, so it stays put when a view
   hides the fiscal-year toggle (Workforce Mix, Competency Teams) instead of jumping left */
#ana-meta{margin-left:auto}
/* Keep the scrollbar's space even when a view is too short to scroll (Manager Scorecard), so the
   page does not slide sideways as the scrollbar comes and goes between views and tabs */
html{scrollbar-gutter:stable}

/* spots where the light build painted text or a panel for a light background */
.logo,.header-title,.header-meta .val{color:var(--ds-tx)!important}
.header-meta .lbl{color:var(--ds-mu2)!important}
.dir-card .photo-wrap{background:linear-gradient(135deg,#1c1c22,#121215)!important}
.heatmap .hm-corner,.heatmap .hm-month-header{background:#1a1a1f!important;color:var(--ds-tx)!important}
.heatmap .hm-month-header.hm-fy-col{background:var(--pwc-orange)!important;color:#fff!important}
.heatmap .hm-spacer{background:transparent!important}
.bio-banner{background:linear-gradient(120deg,#1b1b21 0%,#26262e 100%)!important}
.bio-banner .who,.bio-banner .who h2,.bio-banner .who .pill{color:var(--ds-tx)!important}
.rt-notice{background:rgba(255,182,0,.10)!important;border:1px solid rgba(255,182,0,.35)!important;color:var(--ds-tx)!important}
.rt-notice b{color:var(--pwc-yellow)!important}
.util-chart{background:rgba(255,255,255,.03)!important}
/* pale heatmap cells keep dark text — the light build's .text-dark rule follows grey-900, which is now light */
.heatmap .hm-cell.text-dark,.hm-cell.text-dark{color:#2D2D2D!important}
.cr-cell.text-dark,.text-on-light{color:#2D2D2D!important}

/* the light dashboard writes near-black onto light chips; on a dark card that reads as a hole */
[fill="#2D2D2D"]{fill:#D8D8DE}[stroke="#2D2D2D"]{stroke:#D8D8DE}
[fill="#464646"]{fill:#9A9AA2}[stroke="#464646"]{stroke:#9A9AA2}
[stroke="#F2F2F2"]{stroke:rgba(255,255,255,.12)}
[fill="#F2F2F2"]{fill:rgba(255,255,255,.06)}
svg text{fill:var(--ds-mu)}

/* the dashboard's own landing is replaced by the Scroll Story */
#tab-overview{display:none!important}
body.landing nav.tabs{display:flex}


/* ---- Pulse: one two-handle band instead of the From/To pair ----
   The two range inputs stay in the page and keep driving everything; they are just hidden and
   written to by the handles below, so Pulse's own logic (crossing guard, labels, re-render) is
   untouched. */
.cr-thr-row.pl-two{display:none}
.pl-band{display:flex;align-items:center;gap:14px;margin-top:2px}
.pl-rail{position:relative;flex:1 1 auto;min-width:150px;max-width:320px;height:22px;cursor:pointer;touch-action:none}
.pl-rail::before{content:"";position:absolute;left:9px;right:9px;top:50%;height:4px;margin-top:-2px;
  border-radius:2px;background:rgba(255,255,255,.15)}
.pl-fill{position:absolute;top:50%;height:4px;margin-top:-2px;border-radius:2px;background:var(--pwc-orange);
  left:calc(9px + var(--a,0) * (100% - 18px));width:calc(var(--w,1) * (100% - 18px));pointer-events:none}
.pl-dot{position:absolute;top:50%;left:calc(9px + var(--p,0) * (100% - 18px));width:18px;height:18px;
  margin:-9px 0 0 -9px;padding:0;border:0;border-radius:50%;background:var(--pwc-orange);
  box-shadow:0 1px 5px rgba(0,0,0,.55);cursor:grab;touch-action:none}
.pl-dot:active{cursor:grabbing}
.pl-dot:focus-visible{outline:2px solid var(--pwc-yellow);outline-offset:3px}
.pl-band-val{font:600 13px/1 'Helvetica Neue',Arial,sans-serif;color:var(--pwc-orange);white-space:nowrap}

@media (max-width:900px){
  header.brand{padding:14px 16px;flex-wrap:wrap}
  .ds-back{padding:9px 14px}
}
"""

BOOT = """
<script>
/* The Scroll Story is the landing page for this build: drop the dashboard's own
   overview and open on the Directory, exactly as the dashboard's own fallback does. */
(function(){
  try{
    document.body.classList.remove('landing');
    var ov = document.getElementById('tab-overview'); if (ov) ov.remove();
    var btn = document.querySelector('nav.tabs button.active') ||
              document.querySelector('nav.tabs button[data-tab="directory"]');
    if (btn) btn.click();
    window.scrollTo(0, 0);
  }catch(e){ if (window.console) console.warn('dark build boot:', e); }
})();
</script>
"""


VIEW_ANCHOR = '<button data-view="mix">Workforce Mix</button>'
BODY_ANCHOR = '<div class="ana-grid" id="ana-body"></div>'



BAND_JS = """
<script>
/* One band control for Pulse, writing to the two sliders the page already has. */
(function(){
  var lo = document.getElementById('pl-min-slider'), hi = document.getElementById('pl-max-slider'),
      rail = document.getElementById('pl-rail'), out = document.getElementById('pl-band-val'),
      fill = document.getElementById('pl-fill');
  if (!lo || !hi || !rail) return;
  var MIN = +lo.min || 0, MAX = +lo.max || 150, STEP = +lo.step || 5,
      dots = [].slice.call(rail.querySelectorAll('.pl-dot')), drag = null;

  function pos(v){ return (v - MIN) / (MAX - MIN || 1); }
  function paint(){
    var a = +lo.value, b = +hi.value;
    dots[0].style.setProperty('--p', pos(a));
    dots[1].style.setProperty('--p', pos(b));
    dots[0].setAttribute('aria-valuenow', a);
    dots[1].setAttribute('aria-valuenow', b);
    if (fill){ fill.style.setProperty('--a', pos(a)); fill.style.setProperty('--w', Math.max(0, pos(b) - pos(a))); }
    if (out) out.textContent = a + '–' + b + '%' + (b >= MAX ? '+' : '');
  }
  function set(end, v){
    v = Math.max(MIN, Math.min(MAX, Math.round(v / STEP) * STEP));
    var el = end === 'min' ? lo : hi;
    if (end === 'min' && v > +hi.value) v = +hi.value;      /* the ends cannot cross */
    if (end === 'max' && v < +lo.value) v = +lo.value;
    if (+el.value === v) return;
    el.value = v;
    el.dispatchEvent(new Event('input', {bubbles: true}));   /* Pulse recalculates from here */
    paint();
  }
  function valueAt(clientX){
    var r = rail.getBoundingClientRect(), inner = r.width - 18;
    return MIN + Math.max(0, Math.min(1, (clientX - r.left - 9) / (inner || 1))) * (MAX - MIN);
  }

  rail.addEventListener('pointerdown', function(ev){
    var v = valueAt(ev.clientX),
        end = Math.abs(v - +lo.value) <= Math.abs(v - +hi.value) ? 'min' : 'max';
    if (ev.target.classList.contains('pl-dot')) end = ev.target.getAttribute('data-end');
    drag = end; set(end, v);
    if (rail.setPointerCapture) { try { rail.setPointerCapture(ev.pointerId); } catch (e) {} }
    ev.preventDefault();
  });
  rail.addEventListener('pointermove', function(ev){ if (drag) set(drag, valueAt(ev.clientX)); });
  ['pointerup', 'pointercancel', 'lostpointercapture'].forEach(function(t){
    rail.addEventListener(t, function(){ drag = null; });
  });
  dots.forEach(function(d){
    d.addEventListener('keydown', function(ev){
      var end = d.getAttribute('data-end'), cur = +(end === 'min' ? lo : hi).value, k = ev.key;
      if (k === 'ArrowLeft' || k === 'ArrowDown') set(end, cur - STEP);
      else if (k === 'ArrowRight' || k === 'ArrowUp') set(end, cur + STEP);
      else if (k === 'Home') set(end, MIN);
      else if (k === 'End') set(end, MAX);
      else return;
      ev.preventDefault();
    });
  });
  /* Keep in step when something else moves the sliders. Pulse's own Reset writes .value straight
     and re-renders itself without firing an event, so a repaint after any click in the tab is the
     reliable catch-all — paint() only writes a few styles. */
  lo.addEventListener('input', paint);
  hi.addEventListener('input', paint);
  lo.addEventListener('change', paint);
  hi.addEventListener('change', paint);
  var tab = document.getElementById('tab-pulse');
  if (tab) tab.addEventListener('click', function(){ requestAnimationFrame(paint); }, true);
  paint();
})();
</script>
"""

def dark_html(extra=None):
    """The root dashboard, re-skinned dark, as a string. build_scroll_story.py embeds it.

    `extra` adds one view to the Team Analytics tab: {"html", "css", "js", "label", "view"}.
    It is appended to the tab's own toggle and panel — the dashboard's markup is not rewritten,
    and if either anchor goes missing in a future root build the view is skipped with a warning.
    """
    if not SOURCE.exists():
        raise SystemExit(f"Could not find {SOURCE}. Run the root generate_report.py first.")
    html = SOURCE.read_text(encoding="utf-8")

    for old, new in SERIES_SWAP.items():
        if old not in html:
            print(f"[warn] chart palette constant {old} not found — series colours may stay dark")
        html = html.replace(old, new)

    for old, new in TEXT_SWAP.items():
        if old not in html:
            print(f"[warn] {old[:40]!r} not found — the gender utilization KPI keeps its own wording")
        html = html.replace(old, new)

    if "id=\"ds-skin\"" in html:
        raise SystemExit("The source dashboard already carries the dark skin — point SOURCE at the light build.")
    skin = SKIN + (extra.get("css") or "" if extra else "")
    html = html.replace("</head>", f'<style id="ds-skin">{skin}</style>\n</head>', 1)

    if extra and extra.get("html"):
        if VIEW_ANCHOR in html and BODY_ANCHOR in html:
            btn = f'<button data-view="{extra["view"]}">{extra["label"]}</button>'
            html = html.replace(VIEW_ANCHOR, VIEW_ANCHOR + btn, 1)
            html = html.replace(BODY_ANCHOR, BODY_ANCHOR + extra["html"], 1)
            html = html.replace("</body>", f'<script>window.__CT={extra["data"]};</script>\n'
                                           f'<script>{extra["js"]}</script>\n</body>', 1)
        else:
            print("[warn] Team Analytics anchors not found — the Competency Teams view was left out")

    # the button back to the landing page, in the header
    m = re.search(r"(<header class=\"brand\"[^>]*>)(.*?)(</header>)", html, re.S)
    if m:
        html = html[:m.end(2)] + BACK_BTN + html[m.end(2):]
    else:
        print("[warn] header not found — no Overview button was added")


    # Pulse: one two-handle band in place of the From/To pair (the inputs stay, hidden)
    i = html.find('id="pl-min-slider"')
    g = html.rfind('<div class="cr-thr-group">', 0, i) if i > 0 else -1
    r = html.find('<div class="cr-thr-row">', g) if g > 0 else -1
    if r > 0:
        widget = ('<div class="pl-band" id="pl-band">'
                  '<span class="pl-rail" id="pl-rail"><i class="pl-fill" id="pl-fill"></i>'
                  '<button type="button" class="pl-dot" data-end="min" role="slider" aria-label="Lowest utilization to show"'
                  ' aria-valuemin="0" aria-valuemax="150" aria-valuenow="0"></button>'
                  '<button type="button" class="pl-dot" data-end="max" role="slider" aria-label="Highest utilization to show"'
                  ' aria-valuemin="0" aria-valuemax="150" aria-valuenow="150"></button></span>'
                  '<span class="pl-band-val" id="pl-band-val">0\u2013150%</span></div>')
        html = html[:r] + widget + html[r:].replace('<div class="cr-thr-row">',
                                                    '<div class="cr-thr-row pl-two">', 1)
        html = html.replace("</body>", BAND_JS + "</body>", 1)
    else:
        print("[warn] the Pulse utilization sliders were not found — the band control was left out")

    return html.replace("</body>", BOOT + "</body>", 1)


if __name__ == "__main__":
    raise SystemExit("This is now a module. Run build_scroll_story.py — it embeds the dashboard "
                     "in the one output file.")
