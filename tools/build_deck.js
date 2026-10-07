// Builds docs/Copilot-Studio-Agents-SDK-SharePoint.pptx from the validated results and screenshots.
// Run with:  $env:NODE_PATH = (npm root -g); node tools/build_deck.js   (needs pptxgenjs, react, react-dom, react-icons, sharp)
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const ROOT = path.resolve(__dirname, "..");
const IMG = path.join(ROOT, "docs", "images");
const RES = path.join(ROOT, "docs", "test-results");
const OUT = path.join(ROOT, "docs", "Copilot-Studio-Agents-SDK-SharePoint.pptx");
const REPO = process.env.REPO_URL || "https://github.com/patmeh1/copilot-studio-agents-sdk-sharepoint";
const SITE = process.env.SITE_URL || "https://patmeh1.github.io/copilot-studio-agents-sdk-sharepoint/";

const C = {
  dark: "2B2A29", stat: "45423F", bg: "F7F4EF", card: "FFFFFF", border: "DEDEDE", text: "242424", muted: "5C5C5C",
  accent: "B11F4B", accentSoft: "F6E4EA", ok: "15803D", okSoft: "E3F4E8", bad: "C62828", badSoft: "FBE5E5",
  warn: "B45309", warnSoft: "FDF0D8", white: "FFFFFF", code: "FFFFFF", pink: "FD8EA1",
};
const H = "Segoe UI Semibold";
const B = "Segoe UI";
const W = 13.333, HGT = 7.5, M = 0.5;
const NBH = "\u2011"; // non-breaking hyphen
const NBSP = "\u00A0";

// Crops (left, top, width, height) in source pixels - keep the part that matters readable on a slide.
const CROPS = {
  "10-agentA-authenticate-with-microsoft.png": [545, 90, 550, 286],
  "11-agentA-web-app-connection-string.png": [1272, 10, 640, 366],
  "12-agentA-share.png": [530, 192, 860, 503],
  "16-enable-app.png": [1336, 430, 576, 330],
  "19-webpart-property-pane.png": [1532, 303, 296, 251],
  "24-userB-agentB-embed-answers.png": [0, 0, 600, 603],
  "25-site-access-denied.png": [640, 270, 800, 457],
  "26-embed-anonymous-browser.png": [0, 228, 500, 505],
  "36-agent-link-userC-no-site-a.png": [1140, 405, 425, 285],
  "37-sdk-userC-no-site-a-works.png": [0, 0, 587, 310],
  "38-agent-link-userC-file-read-only-works.png": [100, 48, 1500, 952],
};

async function imgData(file) {
  const p = path.join(IMG, file);
  const crop = CROPS[file];
  let pipe = sharp(p);
  if (crop) pipe = pipe.extract({ left: crop[0], top: crop[1], width: crop[2], height: crop[3] });
  const buf = await pipe.png().toBuffer();
  const meta = await sharp(buf).metadata();
  return { data: "image/png;base64," + buf.toString("base64"), w: meta.width, h: meta.height };
}

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

const shadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 1, angle: 90, opacity: 0.12 });

function footer(slide, n) {
  slide.addText("Copilot Studio x SharePoint  |  Microsoft 365 Agents SDK", { x: M, y: HGT - 0.42, w: 7, h: 0.3, fontFace: B, fontSize: 10, color: C.muted, margin: 0 });
  slide.addText(String(n), { x: W - 1.0, y: HGT - 0.42, w: 0.5, h: 0.3, fontFace: B, fontSize: 10, color: C.muted, align: "right", margin: 0 });
}

function title(slide, kicker, text) {
  slide.addText(kicker.toUpperCase(), { x: M, y: 0.32, w: 12, h: 0.28, fontFace: H, fontSize: 12, color: C.accent, charSpacing: 2, margin: 0 });
  slide.addText(text, { x: M, y: 0.64, w: 12.33, h: 0.6, fontFace: H, fontSize: 28, color: C.text, margin: 0, valign: "top" });
}

function card(pres, slide, x, y, w, h, accent = C.accent) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.card }, line: { color: C.border, width: 0.75 }, shadow: shadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.07, h, fill: { color: accent }, line: { color: accent, width: 0 } });
}

function numberBadge(pres, slide, x, y, num) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: 0.45, h: 0.45, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  slide.addText(String(num), { x, y, w: 0.45, h: 0.45, fontFace: H, fontSize: 14, color: C.white, align: "center", valign: "middle", margin: 0 });
}

function pill(pres, slide, text, x, y, kind, w) {
  const map = { ok: [C.okSoft, C.ok], bad: [C.badSoft, C.bad], warn: [C.warnSoft, C.warn], muted: ["F0F0F0", C.muted] };
  const [fill, line] = map[kind] || map.muted;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.3, rectRadius: 0.15, fill: { color: fill }, line: { color: line, width: 0.75 } });
  slide.addText(text, { x, y, w, h: 0.3, fontFace: H, fontSize: 10, color: C.text, align: "center", valign: "middle", margin: 0 });
}

function kind(v) {
  const u = v.toUpperCase();
  if (u.startsWith("WORKS") || u === "OK") return "ok";
  if (u.includes("NOT SUPPORTED") || u.startsWith("BLOCKED") || u.includes("DENIED")) return "bad";
  if (u.includes("ANSWERS")) return "warn";
  return "muted";
}

// Places an image contained in the outer box (x, y, maxW, maxH) inside a thin white frame; align: "center" | "left".
// Returns the outer frame box, so frames respect the same margins as cards and tables.
function place(pres, slide, im, x, y, maxW, maxH, align = "center") {
  const pad = 0.03, bw = maxW - 2 * pad, bh = maxH - 2 * pad;
  let w = bw, h = (bw * im.h) / im.w;
  if (h > bh) { h = bh; w = (bh * im.w) / im.h; }
  const fx = align === "center" ? x + (maxW - (w + 2 * pad)) / 2 : x;
  slide.addShape(pres.shapes.RECTANGLE, { x: fx, y, w: w + 2 * pad, h: h + 2 * pad, fill: { color: C.white }, line: { color: C.border, width: 0.75 }, shadow: shadow() });
  slide.addImage({ data: im.data, x: fx + pad, y: y + pad, w, h });
  return { x: fx, y, w: w + 2 * pad, h: h + 2 * pad };
}

// Fixed-size white frame (fw x fh) with the image contained at its top-left - keeps a row of screenshots uniform.
function placeFramed(pres, slide, im, x, y, fw, fh) {
  const pad = 0.03, bw = fw - 2 * pad, bh = fh - 2 * pad;
  let w = bw, h = (bw * im.h) / im.w;
  if (h > bh) { h = bh; w = (bh * im.w) / im.h; }
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w: fw, h: fh, fill: { color: C.white }, line: { color: C.border, width: 0.75 }, shadow: shadow() });
  slide.addImage({ data: im.data, x: x + pad, y: y + pad, w, h });
  return { x, y, w: fw, h: fh };
}

// Outer frame height of an image placed at outer width fw.
const frameH = (im, fw) => ((fw - 0.06) * im.h) / im.w + 0.06;

function caption(slide, text, x, y, w) {
  slide.addText(text, { x, y, w, h: 0.3, fontFace: B, fontSize: 11, color: C.muted, align: "center", margin: 0 });
}

function box(pres, slide, x, y, w, h, t1, t2, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color: opts.fill || C.white },
    line: { color: opts.line || "919191", width: opts.lineW || 1, dashType: opts.dash || "solid" },
  });
  slide.addText([
    { text: t1, options: { fontFace: H, fontSize: 13, color: C.text, breakLine: true } },
    { text: t2, options: { fontFace: B, fontSize: opts.subSize || 10, color: C.muted } },
  ], { x, y, w, h, align: "center", valign: "middle", margin: 0.06 });
}

function arrow(pres, slide, x1, y1, x2, y2, color = "7A7A7A", dash = "solid", width = 1.5) {
  slide.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001,
    flipH: x2 < x1, flipV: y2 < y1, line: { color, width, dashType: dash, endArrowType: "triangle" },
  });
}

function codeBox(pres, slide, x, y, w, h, text, size = 11) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.code }, line: { color: "C9C4BC", width: 0.75 } });
  slide.addText(text, { x: x + 0.15, y, w: w - 0.3, h, fontFace: "Consolas", fontSize: size, color: C.text, margin: 0, valign: "middle" });
}

(async () => {
  const summary = JSON.parse(fs.readFileSync(path.join(RES, "summary.json"), "utf8"));
  const im = {};
  for (const f of fs.readdirSync(IMG).filter(f => f.endsWith(".png"))) im[f] = await imgData(f);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "GitHub Copilot";
  pres.title = "Authenticated Copilot Studio agents on many SharePoint sites - Microsoft 365 Agents SDK";
  const icCheck = await icon(fa.FaCheckCircle, C.ok);
  const icCheckDark = await icon(fa.FaCheckCircle, "4ADE80");
  const icWarn = await icon(fa.FaExclamationTriangle, C.warn);
  const icX = await icon(fa.FaTimesCircle, C.bad);
  let n = 0;

  // 1 Title (no footer)
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.dark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: HGT, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
    s.addText("CUSTOMER USE CASE  |  VALIDATED 7 OCT 2026", { x: 0.8, y: 1.15, w: 11.7, h: 0.4, fontFace: H, fontSize: 14, color: C.pink, charSpacing: 2, margin: 0 });
    s.addText("One authenticated Copilot Studio agent on every SharePoint site", { x: 0.8, y: 1.7, w: 11.7, h: 1.7, fontFace: H, fontSize: 44, color: C.white, margin: 0, valign: "top" });
    s.addText(`Microsoft 365 Agents SDK vs. SharePoint channel + Agent link - built, deployed and tested with 4${NBSP}users, 3${NBSP}sites, 2${NBSP}agents and 2${NBSP}harnesses`,
      { x: 0.8, y: 3.55, w: 11.7, h: 0.9, fontFace: B, fontSize: 18, color: "DEDEDE", margin: 0, valign: "top" });
    // Third value = measured glyph width (in) of the number at 40pt Segoe UI Semibold, so every label sits 0.25in after its number.
    const stats = [["15", "persona x site test runs", 0.532], ["45", "web-part test cells", 0.628], ["0", "extra site permissions with the SDK web part", 0.308]];
    const bw = (11.7 - 2 * 0.3) / 3;
    stats.forEach(([v, l, vw], i) => {
      const x = 0.8 + i * (bw + 0.3);
      const lx = x + 0.3 + vw + 0.25;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 5.0, w: bw, h: 1.25, fill: { color: C.stat }, line: { color: "5A5653", width: 0.75 } });
      s.addText(v, { x: x + 0.3, y: 5.05, w: vw + 0.12, h: 1.15, fontFace: H, fontSize: 40, color: C.pink, valign: "middle", margin: 0 });
      s.addText(l, { x: lx, y: 5.05, w: x + bw - 0.25 - lx, h: 1.15, fontFace: B, fontSize: 14, color: "EDEDED", valign: "middle", margin: 0 });
    });
  }

  // 2 Customer setup
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "The use case", "What the customer built - and what they observed");
    card(pres, s, M, 1.55, 6.0, 4.2);
    s.addText("Customer setup", { x: 0.8, y: 1.7, w: 5.5, h: 0.4, fontFace: H, fontSize: 18, color: C.text, margin: 0 });
    const steps = [
      "Agent deployed to Site A with the Copilot Studio SharePoint channel",
      "On Site B, Site C and others: an Agent link web part points to that agent",
      "Users are added (shared) on the agent",
      "Users get access to Site A, Site B and every site that shows the agent",
    ];
    // Text is top-aligned so badges and dots sit level with the first line of each item.
    steps.forEach((t, i) => {
      const y = 2.3 + i * 0.82;
      numberBadge(pres, s, 0.85, y, i + 1);
      s.addText(t, { x: 1.45, y: y + 0.08, w: 4.8, h: 0.62, fontFace: B, fontSize: 14, color: C.text, valign: "top", margin: 0 });
    });
    card(pres, s, 6.83, 1.55, 6.0, 4.2, C.muted);
    s.addText("Customer observations", { x: 7.13, y: 1.7, w: 5.5, h: 0.4, fontFace: H, fontSize: 18, color: C.text, margin: 0 });
    const obs = [
      "Chat works from any site with the Agent link web part",
      "Removed from the agent -> \"You don't have access to this agent\"; re-added -> access restored",
      "Users must have access to Site A, even when chatting from another site",
      "Same behaviour with and without an M365 Copilot license",
    ];
    obs.forEach((t, i) => {
      const y = 2.3 + i * 0.82;
      s.addShape(pres.shapes.OVAL, { x: 7.22, y: y + 0.155, w: 0.14, h: 0.14, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
      s.addText(t, { x: 7.5, y: y + 0.08, w: 5.05, h: 0.62, fontFace: B, fontSize: 14, color: C.text, valign: "top", margin: 0 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 6.0, w: 12.33, h: 0.8, fill: { color: C.accentSoft }, line: { color: C.accentSoft, width: 0 } });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: 6.0, w: 0.07, h: 0.8, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
    s.addText([{ text: "Their question: ", options: { bold: true } }, { text: "do you see any concerns or potential issues with this approach?" }],
      { x: 0.8, y: 6.0, w: 11.8, h: 0.8, fontFace: B, fontSize: 16, color: C.text, valign: "middle", margin: 0 });
    footer(s, n);
  }

  // 3 Answer at a glance
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Answer at a glance", "Valid approach - with one structural dependency to manage");
    const cw = (12.33 - 2 * 0.25) / 3;
    const cards = [
      [icCheck, C.ok, "Works", `Standard-harness agent + Agents SDK web part on any number of sites. Silent Entra SSO, the agent knows the user, Copilot Studio sharing enforced (revoke and re${NBH}grant verified).`],
      [icWarn, C.warn, "Depends on Site A", `Customer pattern (SharePoint channel + Agent link): every user needs Read on the .agent file in Site${NBSP}A. Verified: that file alone is enough - the site can stay closed.`],
      [icX, C.bad, "Not today", "Agents SDK calls to GitHub Copilot harness agents fail with \"This action doesn't support agents built with the GitHub Copilot harness\". Their embed code is not authenticated."],
    ];
    cards.forEach(([ic, col, h, t], i) => {
      const x = M + i * (cw + 0.25);
      card(pres, s, x, 1.55, cw, 3.15, col);
      s.addImage({ data: ic, x: x + 0.3, y: 1.8, w: 0.5, h: 0.5 });
      s.addText(h, { x: x + 0.95, y: 1.78, w: cw - 1.1, h: 0.55, fontFace: H, fontSize: 20, color: C.text, valign: "middle", margin: 0 });
      s.addText(t, { x: x + 0.3, y: 2.5, w: cw - 0.55, h: 2.1, fontFace: B, fontSize: 16.5, color: C.text, valign: "top", margin: 0 });
    });
    card(pres, s, M, 4.95, 12.33, 1.45, C.accent);
    s.addText([
      { text: "Bottom line for the customer: ", options: { bold: true, color: C.accent } },
      { text: "keep the SharePoint-channel pattern if no-code matters, but host the agent in a dedicated, locked-down site or library and grant one security group Read on the .agent file. Use the Agents SDK web part when you want no host-site dependency, inline/branded chat or several agents per page." },
    ], { x: 0.8, y: 5.0, w: 11.8, h: 1.35, fontFace: B, fontSize: 15, color: C.text, valign: "middle", margin: 0 });
    footer(s, n);
  }

  // 4 Two patterns, one runtime
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Two patterns, one runtime", "Same runtime - the Agent link adds a hop through Site A");
    card(pres, s, M, 1.55, 6.0, 4.45);
    s.addText("SharePoint channel + Agent link (customer)", { x: 0.8, y: 1.65, w: 5.6, h: 0.4, fontFace: H, fontSize: 15, color: C.text, margin: 0 });
    box(pres, s, 0.8, 2.25, 1.6, 0.95, "Site B / C page", "Agent link card");
    box(pres, s, 2.7, 2.25, 1.6, 0.95, "Copilot pane", "in SharePoint");
    box(pres, s, 4.6, 2.25, 1.7, 0.95, "Copilot Studio", "api.powerplatform.com");
    box(pres, s, 0.8, 4.3, 1.6, 0.95, "Site A (host)", "Copilot Studio Agents <agent>.agent", { fill: C.accentSoft, line: C.bad, dash: "dash", subSize: 9 });
    box(pres, s, 4.6, 4.3, 1.7, 0.95, "Agent", "sharing check");
    arrow(pres, s, 2.4, 2.72, 2.7, 2.72);
    arrow(pres, s, 4.3, 2.72, 4.6, 2.72);
    arrow(pres, s, 3.2, 3.2, 2.1, 4.3, C.bad, "dash", 1.75);
    s.addText([{ text: "reads .agent file", options: { breakLine: true } }, { text: "(user needs Read)" }],
      { x: 3.05, y: 3.5, w: 1.5, h: 0.5, fontFace: B, fontSize: 10.5, color: C.bad, margin: 0 });
    arrow(pres, s, 5.45, 3.2, 5.45, 4.3);
    s.addText("Host-site dependency - every user needs Read on the .agent file", { x: 0.8, y: 5.45, w: 5.6, h: 0.3, fontFace: B, fontSize: 11, color: C.bad, margin: 0 });
    card(pres, s, 6.83, 1.55, 6.0, 4.45, C.ok);
    s.addText("Agents SDK web part (this implementation)", { x: 7.13, y: 1.65, w: 5.6, h: 0.4, fontFace: H, fontSize: 15, color: C.text, margin: 0 });
    box(pres, s, 7.13, 2.25, 1.9, 0.95, "Any site page", "SPFx web part, inline chat");
    box(pres, s, 10.4, 2.25, 2.15, 0.95, "Copilot Studio", "Agents SDK API");
    box(pres, s, 7.13, 4.3, 1.9, 0.95, "Microsoft Entra ID", "MSAL, SPA app");
    box(pres, s, 10.4, 4.3, 2.15, 0.95, "Agent", "sharing check");
    arrow(pres, s, 9.03, 2.72, 10.4, 2.72, C.ok, "solid", 2);
    s.addText("connection string", { x: 9.03, y: 2.36, w: 1.37, h: 0.3, fontFace: B, fontSize: 10, color: C.ok, align: "center", margin: 0 });
    arrow(pres, s, 7.7, 3.2, 7.7, 4.3);
    s.addText("sign-in", { x: 7.0, y: 3.6, w: 0.6, h: 0.3, fontFace: B, fontSize: 10, color: C.muted, margin: 0, align: "right" });
    arrow(pres, s, 8.45, 4.3, 8.45, 3.2);
    s.addText("token", { x: 8.55, y: 3.6, w: 0.8, h: 0.3, fontFace: B, fontSize: 10, color: C.muted, margin: 0 });
    arrow(pres, s, 11.47, 3.2, 11.47, 4.3);
    s.addText("No host site - nothing to read in any other site", { x: 7.13, y: 5.45, w: 5.4, h: 0.3, fontFace: B, fontSize: 11, color: C.ok, margin: 0 });
    s.addText([
      { text: "Under the hood: ", options: { bold: true } },
      { text: "the .agent file written by the SharePoint channel holds copilotStudioMetadata { botSchema, environmentID, hostnameSuffix: api.powerplatform.com, transport: rest } - the runtime the Agents SDK client calls directly. That is why sharing, revocation and licensing behave the same." },
    ], { x: M, y: 6.3, w: 12.33, h: 0.65, fontFace: B, fontSize: 12, color: C.text, margin: 0, valign: "top" });
    footer(s, n);
  }

  // 5 Head-to-head
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Head-to-head on Site B", "Does the user need Site A? Only the pattern changes");
    const rows = [
      ["1", "User A", "Shared, Read on Site A, no Copilot license", "WORKS", "WORKS"],
      ["2", "User C", "Not shared, no Site A access", "BLOCKED", "BLOCKED"],
      ["3", "User C", "Shared, no Site A access", "BLOCKED", "WORKS"],
      ["4", "User C", "Shared, Read on only the .agent file", "WORKS", "WORKS"],
    ];
    const colW = [0.4, 0.85, 3.45, 1.25, 1.25];
    const hdr = ["#", "User", "State", "Agent link", "Agents SDK"].map(t => ({ text: t, options: { bold: true, fill: { color: "EFEFEF" }, color: C.text, fontFace: H } }));
    const data = [hdr, ...rows.map(r => [r[0], r[1], r[2], "", ""].map(t => ({ text: t, options: { fontFace: B } })))];
    s.addTable(data, { x: M, y: 1.55, w: 7.2, colW, rowH: 0.55, fontSize: 12, color: C.text, border: { type: "solid", pt: 0.75, color: C.border }, fill: { color: C.white }, valign: "middle" });
    const xAL = M + colW[0] + colW[1] + colW[2], xSDK = xAL + colW[3];
    rows.forEach((r, i) => {
      const y = 1.55 + 0.55 * (i + 1) + 0.125;
      pill(pres, s, r[3], xAL + (colW[3] - 1.0) / 2, y, kind(r[3]), 1.0);
      pill(pres, s, r[4], xSDK + (colW[4] - 1.0) / 2, y, kind(r[4]), 1.0);
    });
    s.addText([
      { text: "Test 3 is the customer's Site A dependency. ", options: { bold: true } },
      { text: `The Agent link says "You need a Microsoft${NBSP}365 Copilot license to access this agent" - misleading, because unlicensed User${NBSP}A works. The Agents SDK web part works for the same user without any Site${NBSP}A access.` },
    ], { x: M, y: 4.55, w: 7.2, h: 0.75, fontFace: B, fontSize: 13, color: C.text, margin: 0, valign: "top" });
    card(pres, s, M, 5.55, 7.2, 0.8, C.ok);
    s.addText([{ text: "Fix for the customer pattern: ", options: { bold: true } }, { text: "grant Read on only the .agent file (test 4) - Site A itself can stay Access Denied." }],
      { x: 0.8, y: 5.55, w: 6.6, h: 0.8, fontFace: B, fontSize: 13, color: C.text, margin: 0, valign: "middle" });
    // The license message is tiny in the screenshot, so it is repeated as a readable callout next to it.
    const a = place(pres, s, im["36-agent-link-userC-no-site-a.png"], 8.0, 1.55, 2.2, 1.6, "left");
    const qx = a.x + a.w + 0.3;
    s.addText([
      { text: "Test 3 - Agent link", options: { bold: true, breakLine: true } },
      { text: "\"You need a Microsoft 365", options: { italic: true, breakLine: true } },
      { text: "Copilot license to access", options: { italic: true, breakLine: true } },
      { text: "this agent.\"", options: { italic: true } },
    ], { x: qx, y: a.y, w: 12.83 - qx, h: a.h, fontFace: B, fontSize: 12, color: C.text, margin: 0, valign: "middle" });
    const b = place(pres, s, im["37-sdk-userC-no-site-a-works.png"], 8.0, 3.4, 4.83, 2.55);
    caption(s, "Test 3 - Agents SDK web part: works, greets User C by name", 8.0, b.y + b.h + 0.1, 4.83);
    footer(s, n);
  }

  // 6 Mitigation
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Least-privilege fix for the customer pattern", "Grant Read on the .agent file - not on the whole host site");
    const a = place(pres, s, im["38-agent-link-userC-file-read-only-works.png"], M, 1.6, 6.6, 4.4, "left");
    caption(s, "Test 4 - Read on only the .agent file: Agent link works, Site A still Access Denied", M, a.y + a.h + 0.1, a.w);
    const tips = [
      ["Dedicated host", "Deploy the agent to a dedicated, locked-down \"agent host\" site or library with no other content."],
      ["One group, two places", "Share the agent with an Entra security group and grant that group Read on just the .agent file."],
      ["Few editors", "Members with Edit can delete or rename the .agent file and break every Agent link - keep Edit to a few owners."],
      ["Help desk note", "\"You need a Microsoft 365 Copilot license\" usually means the user cannot read the .agent file."],
    ];
    // Cards end level with the screenshot caption (image 4.19 + caption -> 6.19; cards 1.6 + 4 x 1.0 + 3 x 0.2 = 6.2).
    const ch = 1.0, gap = 0.2;
    tips.forEach(([h, t], i) => {
      const y = 1.6 + i * (ch + gap);
      card(pres, s, 7.45, y, 5.38, ch);
      s.addText(h, { x: 7.72, y: y + 0.09, w: 4.95, h: 0.32, fontFace: H, fontSize: 14, color: C.text, margin: 0 });
      s.addText(t, { x: 7.72, y: y + 0.41, w: 4.95, h: 0.52, fontFace: B, fontSize: 12, color: C.text, margin: 0, valign: "top" });
    });
    footer(s, n);
  }

  // 7 Comparison
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Comparison", "SharePoint channel + Agent link vs. Agents SDK web part");
    const rows = [
      ["Effort", "No code", "One-time SPFx package, Entra app, app catalog"],
      ["Host-site dependency", "Yes - Read on the .agent file; one host site at a time", "None - each page stores the connection string"],
      ["Sharing and revocation", "Enforced (customer test)", "Enforced (verified incl. revoke and re-grant)"],
      ["M365 Copilot license", "Not required; consumes Copilot Studio capacity", "Not required; consumes Copilot Studio capacity"],
      ["Experience", "Card on the page opens the Copilot side pane", "Inline chat, themable, several agents per page"],
      ["Error when access is missing", "Misleading license message", "Agent's own access message or explicit errors"],
      ["Governance", "Host-site editors can delete or rename the .agent file", "Central versioned package; per-page configuration"],
      ["GitHub Copilot harness", "SharePoint channel not available", "Not supported by the client library"],
    ];
    const hdr = ["Aspect", "SharePoint channel + Agent link", "Agents SDK web part"].map(t => ({ text: t, options: { bold: true, fill: { color: "EFEFEF" }, fontFace: H } }));
    const data = [hdr, ...rows.map(r => [{ text: r[0], options: { bold: true, fontFace: H } }, { text: r[1], options: { fontFace: B } }, { text: r[2], options: { fontFace: B } }])];
    s.addTable(data, { x: M, y: 1.55, w: 12.33, colW: [2.75, 4.79, 4.79], rowH: 0.56, fontSize: 13, color: C.text, border: { type: "solid", pt: 0.75, color: C.border }, fill: { color: C.white }, valign: "middle", margin: 0.08 });
    footer(s, n);
  }

  // 8 Concerns
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Answering the customer", "Concerns and recommendations for the customer pattern");
    const items = [
      ["HIGH", C.bad, "Single point of failure", "Archiving Site A, changing its permissions or deleting/renaming the .agent file breaks the agent on every site."],
      ["HIGH", C.bad, "Over-permissioning", "Granting Site A to every agent user exposes Site A content. Grant Read on only the .agent file instead (verified)."],
      ["MEDIUM", C.warn, "Two permission systems", "Copilot Studio sharing plus SharePoint file permission - drive both from one Entra security group."],
      ["MEDIUM", C.warn, "Misleading error", "An unreadable .agent file shows \"You need a Microsoft 365 Copilot license\" although unlicensed users work."],
      ["MEDIUM", C.warn, "Capacity, not licenses", "Agents in SharePoint consume Copilot Studio capacity - plan and monitor capacity or pay-as-you-go."],
      ["LOW", C.muted, "One deployment site at a time", "Moving the host means undeploying, redeploying and re-pointing every Agent link."],
    ];
    const cw = (12.33 - 0.25) / 2, ch = 1.38, gap = 0.25;
    items.forEach(([sev, col, h, t], i) => {
      const x = M + (i % 2) * (cw + 0.25), y = 1.65 + Math.floor(i / 2) * (ch + gap);
      card(pres, s, x, y, cw, ch, col);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 0.17, w: 0.95, h: 0.3, rectRadius: 0.05, fill: { color: C.white }, line: { color: col, width: 1 } });
      s.addText(sev, { x: x + 0.3, y: y + 0.17, w: 0.95, h: 0.3, fontFace: H, fontSize: 10, color: col, align: "center", valign: "middle", margin: 0 });
      s.addText(h, { x: x + 1.4, y: y + 0.13, w: cw - 1.6, h: 0.38, fontFace: H, fontSize: 15, color: C.text, margin: 0, valign: "middle" });
      s.addText(t, { x: x + 0.3, y: y + 0.6, w: cw - 0.6, h: 0.72, fontFace: B, fontSize: 13, color: C.text, margin: 0, valign: "top" });
    });
    footer(s, n);
  }

  // 9 Implementation architecture
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "The Agents SDK implementation", "One SPFx package and one Entra app serve every site");
    const comps = [
      ["SPFx web part", "SPFx 1.21.1, deployed tenant-wide from the app catalog"],
      ["MSAL.js with its own SPA app", "acquireTokenSilent, then ssoSilent with the SharePoint user as login hint; popup fallback"],
      ["Agents SDK client", "@microsoft/agents-copilotstudio-client 1.9.1, called with the connection string"],
      ["Bot Framework Web Chat", "4.19.1 with the Fluent theme, bundled - no CDN, no SharePoint API-access approvals"],
    ];
    const ch = 0.98, gap = 0.2;
    comps.forEach(([h, t], i) => {
      const y = 1.7 + i * (ch + gap);
      card(pres, s, M, y, 7.3, ch);
      s.addText(h, { x: 0.8, y: y + 0.1, w: 6.8, h: 0.35, fontFace: H, fontSize: 15, color: C.text, margin: 0 });
      s.addText(t, { x: 0.8, y: y + 0.47, w: 6.8, h: 0.45, fontFace: B, fontSize: 12, color: C.muted, margin: 0, valign: "top" });
    });
    card(pres, s, 8.05, 1.7, 4.78, 4.52, C.warn);
    // Heading and body form one group centred in the card; the body (14pt, 9 lines) is about 2.45in tall.
    const gy = 1.7 + (4.52 - (0.45 + 0.2 + 2.45)) / 2;
    s.addImage({ data: icWarn, x: 8.35, y: gy + 0.015, w: 0.42, h: 0.42 });
    s.addText("Why MSAL, not the SPFx token provider?", { x: 8.9, y: gy, w: 3.8, h: 0.45, fontFace: H, fontSize: 14, color: C.text, margin: 0, valign: "middle" });
    s.addText([
      { text: `SPFx AadTokenProvider cannot obtain CopilotStudio.Copilots.Invoke: the SharePoint first${NBH}party app is not preauthorized by the Power Platform API (AADSTS65002). Microsoft treats this as by design (issue sp${NBH}dev${NBH}docs #10775).`, options: { breakLine: true } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: `So the web part uses its own single-tenant SPA app with delegated CopilotStudio.Copilots.Invoke and tenant-wide consent, created by New${NBH}AgentsSdkSpaApp.ps1.` },
    ], { x: 8.35, y: gy + 0.65, w: 4.3, h: 2.75, fontFace: B, fontSize: 14, color: C.text, margin: 0, valign: "top" });
    s.addText("Tested: silent SSO for MFA users (no prompts); anonymous call to the connection string returns HTTP 401.",
      { x: M, y: 6.45, w: 12.33, h: 0.35, fontFace: B, fontSize: 12.5, color: C.text, margin: 0 });
    footer(s, n);
  }

  // 10 Steps 1-4
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Build it - part 1", "Copilot Studio and Microsoft Entra ID");
    const steps = [
      ["Standard agent, Microsoft auth, publish", "10-agentA-authenticate-with-microsoft.png"],
      ["Copy the Agents SDK connection string", "11-agentA-web-app-connection-string.png"],
      ["Share with users or security groups", "12-agentA-share.png"],
    ];
    const sw = (12.33 - 2 * 0.3) / 3;
    const fh = Math.max(...steps.map(([, f]) => frameH(im[f], sw)));
    steps.forEach(([t, f], i) => {
      const x = M + i * (sw + 0.3);
      numberBadge(pres, s, x, 1.55, i + 1);
      s.addText(t, { x: x + 0.55, y: 1.5, w: sw - 0.55, h: 0.55, fontFace: H, fontSize: 13, color: C.text, margin: 0, valign: "middle" });
      placeFramed(pres, s, im[f], x, 2.25, sw, fh);
    });
    numberBadge(pres, s, M, 5.05, 4);
    s.addText("Register one Entra SPA app for the tenant (scripted, idempotent)", { x: 1.05, y: 5.0, w: 11.7, h: 0.55, fontFace: H, fontSize: 13, color: C.text, margin: 0, valign: "middle" });
    codeBox(pres, s, 1.05, 5.65, 11.78, 1.15,
      "az login --tenant <tenant>.onmicrosoft.com --allow-no-subscriptions\n.\\scripts\\New-AgentsSdkSpaApp.ps1 -TenantId <tenant>.onmicrosoft.com -SharePointRootUrl https://<tenant>.sharepoint.com\n# Power Platform API / CopilotStudio.Copilots.Invoke + SPA redirect = SharePoint origin + tenant-wide consent");
    footer(s, n);
  }

  // 11 Steps 5-9
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Build it - part 2", "Package once, deploy tenant-wide, add to any page");
    numberBadge(pres, s, M, 1.55, 5);
    s.addText("Build the package with tenant defaults - Node 22, everything bundled", { x: 1.05, y: 1.5, w: 11.7, h: 0.55, fontFace: H, fontSize: 13, color: C.text, margin: 0, valign: "middle" });
    codeBox(pres, s, 1.05, 2.15, 11.78, 0.8,
      ".\\scripts\\Build-SpfxPackage.ps1 -AppClientId <appClientId> -Version 1.0.1.0 -NodePath <node22-folder>\n# output: spfx-copilot-agent\\sharepoint\\solution\\copilot-studio-agent-webpart.sppkg");
    const steps = [
      [6, "App catalog: enable for all sites", "16-enable-app.png"],
      [7, "Add web part, paste connection string", "19-webpart-property-pane.png"],
      [8, "Grant site access (Visitor is enough)", "25-site-access-denied.png"],
    ];
    const sw = (12.33 - 2 * 0.3) / 3;
    steps.forEach(([num, t, f], i) => {
      const x = M + i * (sw + 0.3);
      numberBadge(pres, s, x, 3.15, num);
      s.addText(t, { x: x + 0.55, y: 3.1, w: sw - 0.55, h: 0.55, fontFace: H, fontSize: 13, color: C.text, margin: 0, valign: "middle" });
      placeFramed(pres, s, im[f], x, 3.8, sw, 2.26);
    });
    numberBadge(pres, s, M, 6.3, 9);
    s.addText("Validate: greeting shows the user's name, footer shows the signed-in UPN, unshared users are blocked, anonymous calls get 401.",
      { x: 1.05, y: 6.25, w: 11.7, h: 0.55, fontFace: H, fontSize: 13, color: C.text, margin: 0, valign: "middle" });
    footer(s, n);
  }

  // 12 Test design
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Permission test design", "4 users x 3 sites x 2 agents, plus the maker as control");
    const hdr = ["Persona", "Agent A (standard)", "Agent B (GitHub)", "Site A", "Site B", "Site C"].map(t => ({ text: t, options: { bold: true, fill: { color: "EFEFEF" }, fontFace: H } }));
    const rows = [
      ["User A", "Shared", "Shared", "Visitor", "Member", "Owner"],
      ["User B", "Shared", "-", "Member", "Visitor", "No access"],
      ["User C", "-", "Shared", "No access", "Member", "Visitor"],
      ["User D", "-", "-", "Visitor", "No access", "Member"],
      ["MarioR (maker)", "Owner", "Owner", "Owner", "Owner", "Owner"],
    ];
    // The table and the card share top and bottom edges (1.6 + 6 x 0.64 = 5.44).
    const tTop = 1.6, rowH = 0.64, tBottom = tTop + 6 * rowH;
    s.addTable([hdr, ...rows.map(r => r.map((t, j) => ({ text: t, options: { fontFace: j === 0 ? H : B, color: t === "No access" ? C.bad : C.text } })))],
      { x: M, y: tTop, w: 8.0, colW: [1.6, 1.45, 1.4, 1.18, 1.18, 1.19], rowH, fontSize: 14, color: C.text, border: { type: "solid", pt: 0.75, color: C.border }, fill: { color: C.white }, valign: "middle" });
    card(pres, s, 8.85, tTop, 3.98, tBottom - tTop);
    const blt = (t, last) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: !last } });
    // Content block is ~3.05in tall; centre it in the card.
    const cy = tTop + (tBottom - tTop - 3.05) / 2;
    s.addText("Each site page hosts 3 web parts", { x: 9.15, y: cy, w: 3.55, h: 0.3, fontFace: H, fontSize: 14.5, color: C.text, margin: 0, valign: "top" });
    s.addText([blt("Agent A via Agents SDK"), blt("Agent B via Agents SDK"), blt("Agent B via hosted embed (comparison)", true)],
      { x: 9.15, y: cy + 0.42, w: 3.55, h: 0.9, fontFace: B, fontSize: 13, color: C.text, margin: 0, valign: "top", paraSpaceAfter: 4 });
    s.addText("How it was tested", { x: 9.15, y: cy + 1.5, w: 3.55, h: 0.3, fontFace: H, fontSize: 14.5, color: C.text, margin: 0, valign: "top" });
    s.addText([blt("Edge + Playwright, one profile per user"), blt("MFA with a software TOTP"), blt("A greeting and a question per chat"), blt("Screenshot + JSON result per web part", true)],
      { x: 9.15, y: cy + 1.92, w: 3.55, h: 1.15, fontFace: B, fontSize: 13, color: C.text, margin: 0, valign: "top", paraSpaceAfter: 4 });
    const ny = tBottom + 0.3;
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: ny, w: 12.33, h: 0.6, fill: { color: C.accentSoft }, line: { color: C.accentSoft, width: 0 } });
    s.addShape(pres.shapes.RECTANGLE, { x: M, y: ny, w: 0.07, h: 0.6, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
    s.addText([{ text: "Licensing: ", options: { bold: true } }, { text: "no tester had a Microsoft 365 Copilot or Copilot Studio license - Microsoft 365 E5 without Teams only." }],
      { x: 0.8, y: ny, w: 11.8, h: 0.6, fontFace: B, fontSize: 14, color: C.text, margin: 0, valign: "middle" });
    footer(s, n);
  }

  // 13 Results
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Results", "Sharing decides who chats; site access decides who sees it");
    const order = [["a", "User A"], ["b", "User B"], ["c", "User C"], ["d", "User D"], ["mario", "MarioR"]];
    const hdr = ["Persona", "Site", "Page", "Agent A via SDK", "Agent B via SDK", "Agent B embed"].map(t => ({ text: t, options: { bold: true, fill: { color: "EFEFEF" }, fontFace: H } }));
    const rows = [];
    order.forEach(([k, name]) => ["A", "B", "C"].forEach(site => {
      const c = summary[k] && summary[k][site];
      if (!c) return;
      const ok = c.site === "OK";
      rows.push([name, site, ok ? "OK" : "DENIED", ok ? c.A_sdk : "-", ok ? c.B_sdk : "-", ok ? (c.B_embed.startsWith("ANSWERS") ? "ANSWERS (no auth)" : c.B_embed) : "-"]);
    }));
    const colorFor = v => ({ ok: C.ok, bad: C.bad, warn: C.warn, muted: C.muted })[kind(v)];
    const tbl = [hdr, ...rows.map(r => r.map((t, j) => ({ text: t, options: { fontFace: j < 2 ? H : B, bold: j >= 2 && t !== "-", color: j >= 2 ? colorFor(t) : C.text } })))];
    const tblTop = 1.6, tblRowH = 0.315;
    s.addTable(tbl, { x: M, y: tblTop, w: 8.6, colW: [1.25, 0.6, 1.0, 2.05, 1.9, 1.8], rowH: tblRowH, fontSize: 10.5, border: { type: "solid", pt: 0.5, color: C.border }, fill: { color: C.white }, valign: "middle" });
    const notes = [
      [C.ok, "Agent A via SDK works for every shared user on every site they can open (any permission level)"],
      [C.bad, "Unshared users get \"You don't have access to talk to this bot\" on every site"],
      [C.bad, "Agent B via SDK is rejected for everyone, including the owner"],
      [C.warn, "Agent B embed answered unshared users and a signed-out browser"],
      [C.ok, "Revocation and re-grant took effect on the next page load"],
    ];
    // Note cards span exactly the table height with 0.2in gaps.
    const gap = 0.2, ch = (tbl.length * tblRowH - gap * (notes.length - 1)) / notes.length;
    notes.forEach(([col, t], i) => {
      const y = tblTop + i * (ch + gap);
      card(pres, s, 9.4, y, 3.43, ch, col);
      s.addText(t, { x: 9.62, y, w: 3.0, h: ch, fontFace: B, fontSize: 11.5, color: C.text, margin: 0, valign: "middle" });
    });
    footer(s, n);
  }

  // 14 Evidence (three equal screenshots; the Agent link evidence is on slides 5 and 6)
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "Evidence", "Agents SDK web part - screenshots from the automated runs");
    const ev = [
      ["21-userA-agentA-sdk-works.png", "User A: greeted by name and answered", "Silent SSO - the footer shows the signed-in UPN"],
      ["22-userD-agentA-sdk-not-shared.png", "User D (not shared): blocked", "\"You don't have access to talk to this bot\""],
      ["28-userB-after-revoke.png", "User B right after revocation", "Same block; re-granting restored access"],
    ];
    const sw = (12.33 - 2 * 0.3) / 3, top = 1.6;
    ev.forEach(([f, cap, sub], i) => {
      const x = M + i * (sw + 0.3);
      const p = place(pres, s, im[f], x, top, sw, 4.1);
      s.addText(cap, { x, y: p.y + p.h + 0.15, w: sw, h: 0.3, fontFace: H, fontSize: 13, color: C.text, align: "center", margin: 0 });
      s.addText(sub, { x, y: p.y + p.h + 0.47, w: sw, h: 0.3, fontFace: B, fontSize: 11, color: C.muted, align: "center", margin: 0 });
    });
    footer(s, n);
  }

  // 15 GitHub Copilot harness
  {
    const s = pres.addSlide(); n++; s.background = { color: C.bg };
    title(s, "GitHub Copilot harness", "Weather agent: not reachable via the SDK, embed not secured");
    card(pres, s, M, 1.55, 6.04, 1.15, C.bad);
    s.addText([{ text: "Agents SDK: not supported (yet). ", options: { bold: true } }, { text: `Every request, even the owner's, failed with "This action doesn't support agents built with the GitHub Copilot harness." The client supports standard${NBH}harness agents only.` }],
      { x: 0.8, y: 1.55, w: 5.55, h: 1.15, fontFace: B, fontSize: 12.5, color: C.text, margin: 0, valign: "middle" });
    card(pres, s, 6.79, 1.55, 6.04, 1.15, C.warn);
    s.addText([{ text: "App-channel embed: not secured. ", options: { bold: true } }, { text: "It answered users the agent was not shared with and a signed-out browser (anonymous Direct Line token). Treat the existing button and embed on the root site as public." }],
      { x: 7.09, y: 1.55, w: 5.55, h: 1.15, fontFace: B, fontSize: 12.5, color: C.text, margin: 0, valign: "middle" });
    // Equal-height screenshots spread across the full content width (outer frame widths include the 0.03in frame padding).
    const shots = [["23-agentB-sdk-not-supported.png", "Via Agents SDK"], ["24-userB-agentB-embed-answers.png", "Embed answering an unshared user"], ["26-embed-anonymous-browser.png", "Embed answering a signed-out browser"]];
    const ih = 3.43, top = 2.95;
    const ws = shots.map(([f]) => ((ih - 0.06) * im[f].w) / im[f].h + 0.06);
    const gap = (12.33 - ws.reduce((t, v) => t + v, 0)) / (shots.length - 1);
    let x = M;
    shots.forEach(([f, cap], i) => {
      place(pres, s, im[f], x, top, ws[i], ih);
      caption(s, cap, x, top + ih + 0.1, ws[i]);
      x += ws[i] + gap;
    });
    footer(s, n);
  }

  // 16 Recommendations (no footer)
  {
    const s = pres.addSlide(); n++; s.background = { color: C.dark };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.18, h: HGT, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
    s.addText("RECOMMENDATIONS", { x: 0.8, y: 0.6, w: 11, h: 0.4, fontFace: H, fontSize: 14, color: C.pink, charSpacing: 2, margin: 0 });
    s.addText("What to tell the customer", { x: 0.8, y: 1.05, w: 11.7, h: 0.8, fontFace: H, fontSize: 34, color: C.white, margin: 0 });
    const recs = [
      "Their pattern is valid and supported - keep it if no-code matters most.",
      "Host the agent in a dedicated, locked-down site or library; keep Edit rights minimal.",
      "Use one security group for Copilot Studio sharing and Read on the .agent file - not on Site A.",
      "Plan Copilot Studio capacity; unlicensed users consume it.",
      "Choose the Agents SDK web part to remove the host-site dependency or to embed chat inline.",
      "Keep SharePoint-facing agents on the standard harness; retire the unauthenticated GitHub-harness embed.",
    ];
    recs.forEach((t, i) => {
      const y = 2.1 + i * 0.66;
      s.addImage({ data: icCheckDark, x: 0.85, y: y + 0.09, w: 0.32, h: 0.32 });
      s.addText(t, { x: 1.35, y, w: 11.3, h: 0.5, fontFace: B, fontSize: 16, color: "EDEDED", margin: 0, valign: "middle" });
    });
    s.addText([
      { text: `Web pages (tutorial, use cases, test scenarios, customer use case): ${SITE}`, options: { breakLine: true } },
      { text: `Source code and guide: ${REPO}` },
    ], { x: 0.8, y: 6.2, w: 11.9, h: 0.7, fontFace: B, fontSize: 12, color: "C8C8C8", margin: 0, valign: "top" });
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT, "slides:", n);
})();
