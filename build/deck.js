// Šta čini nešto lepim? — build script (pptxgenjs + sharp)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const RAW = process.env.RAW || path.join(__dirname, "img_raw");
const OUT_IMG = path.join(__dirname, "img");
fs.mkdirSync(OUT_IMG, { recursive: true });
const OUT = path.join(__dirname, "..", "Sta_cini_nesto_lepim.pptx");

// ---------- palette ----------
const HEX = {
  text: "1B1916", white: "F7F4EE", ink: "13110F", bone: "F2EEE7",
  gold: "B8924A", cinn: "A3392B", stone: "6E665C", ash: "A69D90", card: "E6E0D6", panel: "221F1B",
};
const THEME = {
  name: "Galerija",
  headFontFace: "Georgia",
  bodyFontFace: "Arial",
  colors: {
    dk1: HEX.text, lt1: HEX.white, dk2: HEX.ink, lt2: HEX.bone,
    accent1: HEX.gold, accent2: HEX.cinn, accent3: HEX.stone, accent4: HEX.ash,
    accent5: HEX.card, accent6: HEX.panel, hlink: HEX.cinn, folHlink: HEX.stone,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "Šta čini nešto lepim? Estetsko shvatanje lepog";
pres.subject = "Estetika";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const S = pres.SchemeColor;
const C = {
  text: S.text1, white: S.background1, ink: S.text2, bone: S.background2,
  gold: S.accent1, cinn: S.accent2, stone: S.accent3, ash: S.accent4, card: S.accent5, panel: S.accent6,
};
const HF = "Georgia"; // heading face (also the theme major font)
const W = 13.333, H = 7.5, M = 0.7;

// ---------- layouts ----------
const footer = (col) => [
  { text: { text: "ŠTA ČINI NEŠTO LEPIM?", options: { x: M, y: 6.98, w: 5, h: 0.3, fontSize: 10, charSpacing: 3, color: col, margin: 0, valign: "middle" } } },
];
pres.defineSlideMaster({ title: "Svetli", background: { color: HEX.bone }, objects: footer(HEX.stone),
  slideNumber: { x: W - M - 0.6, y: 6.98, w: 0.6, h: 0.3, fontSize: 10, color: HEX.stone, align: "right", margin: 0 } });
pres.defineSlideMaster({ title: "Tamni", background: { color: HEX.ink }, objects: footer(HEX.ash),
  slideNumber: { x: W - M - 0.6, y: 6.98, w: 0.6, h: 0.3, fontSize: 10, color: HEX.ash, align: "right", margin: 0 } });
pres.defineSlideMaster({ title: "Pun kadar", background: { color: HEX.ink }, objects: [] });

// ---------- helpers ----------
const meta = JSON.parse(fs.readFileSync(path.join(RAW, "meta.json"), "utf8"));
function rawPath(key) {
  const f = fs.readdirSync(RAW).find((n) => n.startsWith(key + "."));
  if (!f) throw new Error("missing image " + key);
  return path.join(RAW, f);
}
// crop to exact aspect around focus point (fx, fy in 0..1), zoom >= 1 tightens the crop
async function pic(key, wIn, hIn, fx = 0.5, fy = 0.5, zoom = 1, tag = "") {
  const src = rawPath(key);
  const m = await sharp(src).metadata();
  const ar = wIn / hIn;
  let cw = m.width, ch = Math.round(cw / ar);
  if (ch > m.height) { ch = m.height; cw = Math.round(ch * ar); }
  cw = Math.round(cw / zoom); ch = Math.round(ch / zoom);
  let left = Math.round(fx * m.width - cw / 2), top = Math.round(fy * m.height - ch / 2);
  left = Math.max(0, Math.min(m.width - cw, left)); top = Math.max(0, Math.min(m.height - ch, top));
  const ppi = Math.min(cw / wIn, 220);
  const out = path.join(OUT_IMG, `${key}_${tag || `${wIn}x${hIn}`}.jpg`);
  await sharp(src).extract({ left, top, width: cw, height: ch })
    .resize(Math.round(wIn * ppi), Math.round(hIn * ppi)).jpeg({ quality: 88 }).toFile(out);
  return out;
}
async function gradient(name, w, h, dir, stops) {
  const [x2, y2] = dir === "right" ? ["100%", "0%"] : ["0%", "100%"];
  const st = stops.map(([o, a]) => `<stop offset="${o}" stop-color="#${HEX.ink}" stop-opacity="${a}"/>`).join("");
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><defs><linearGradient id="g" x1="0%" y1="0%" x2="${x2}" y2="${y2}">${st}</linearGradient></defs><rect width="100%" height="100%" fill="url(#g)"/></svg>`;
  const out = path.join(OUT_IMG, name + ".png");
  await sharp(Buffer.from(svg)).png().toFile(out);
  return out;
}
function T(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", fontSize: 16, color: C.text, paraSpaceAfter: 0 }, o));
}
const label = (slide, text, x, y, w, color, name) =>
  T(slide, text.toUpperCase(), { x, y, w, h: 0.3, fontSize: 11, bold: true, charSpacing: 4, color, objectName: name });
function frame(slide, x, y, w, h, name, color = C.gold, pad = 0.12) {
  slide.addShape(pres.shapes.RECTANGLE, { x: x - pad, y: y - pad, w: w + 2 * pad, h: h + 2 * pad, line: { color, width: 0.75 }, objectName: name });
}
// museum wall label: artist (bold) · title (italic) · date / place
function wall(slide, x, y, w, artist, title, rest, color, name, align = "left") {
  T(slide, [
    { text: artist, options: { bold: true } },
    { text: "  " + title, options: { italic: true } },
    { text: rest ? "  ·  " + rest : "" },
  ], { x, y, w, h: 0.32, fontSize: 10, color, objectName: name, align });
}
const add = (layout, section) => pres.addSlide({ masterName: layout, sectionTitle: section });

// transitions to inject after writing: slide index (1-based) -> kind
const TRANS = {};

(async () => {
  const gLeft = await gradient("grad_left", 1600, 900, "right", [[0, 0.94], [0.45, 0.78], [0.75, 0.15], [1, 0]]);
  const gBottom = await gradient("grad_bottom", 1600, 900, "down", [[0, 0.55], [0.3, 0.05], [0.55, 0.1], [1, 0.92]]);
  const gHalfBottom = await gradient("grad_halfbottom", 800, 900, "down", [[0, 0.75], [0.18, 0.15], [0.42, 0.45], [0.58, 0.9], [1, 0.97]]);

  const venus169 = await pic("venus", 16, 9, 0.5, 0.5, 1, "169");
  const venusHero = await pic("venus", 16, 9, 0.47, 0.42, 1.3, "hero");

  // ============ SEKCIJA: Uvod ============
  pres.addSection({ title: "Uvod" });

  // 1 — Naslov
  let s = add("Pun kadar", "Uvod"); TRANS[1] = "fade";
  s.addImage({ path: venusHero, x: 0, y: 0, w: W, h: H, objectName: "!!venus", altText: "Sandro Botičeli, Rođenje Venere" });
  s.addImage({ path: gLeft, x: 0, y: 0, w: W, h: H, objectName: "!!veil" });
  label(s, "Estetika  ·  filozofija lepog", M, 0.7, 6, C.gold, "!!kicker");
  T(s, [
    { text: "Šta čini", options: { breakLine: true } },
    { text: "nešto ", options: {} },
    { text: "lepim?", options: { italic: true, color: C.gold } },
  ], { x: M, y: 2.0, w: 7.6, h: 2.8, fontFace: HF, fontSize: 72, color: C.white, lineSpacingMultiple: 0.95, objectName: "!!title" });
  T(s, "Estetsko shvatanje lepog — od Platona do neuronauke", { x: M, y: 4.95, w: 6.4, h: 0.8, fontSize: 20, color: C.white, objectName: "!!sub" });
  wall(s, M, 6.75, 8, "Sandro Botičeli", "Rođenje Venere", "oko 1485, Galerija Ufici, Firenca", C.ash, "!!cap1");
  s.addNotes("Otvaranje. Pitanje koje deluje jednostavno: šta čini nešto lepim? Botičelijeva Venera je vekovima bila sinonim za ideal lepote — ali da li je ona lepa sama po sebi, ili zato što smo naučili da je tako vidimo? To je pitanje kojim se bavi estetika.");

  // 2 — Uvod (Morph: slika se smanjuje u ram)
  s = add("Tamni", "Uvod"); TRANS[2] = "morph";
  const vx = 6.55, vy = 1.2, vw = 6.08, vh = vw * 9 / 16;
  frame(s, vx, vy, vw, vh, "!!vframe");
  s.addImage({ path: venus169, x: vx, y: vy, w: vw, h: vh, objectName: "!!venus", altText: "Sandro Botičeli, Rođenje Venere" });
  wall(s, vx, vy + vh + 0.3, vw, "Sandro Botičeli", "Rođenje Venere", "oko 1485, Galerija Ufici, Firenca", C.ash, "!!cap1");
  label(s, "Uvod", M, 0.7, 4, C.gold, "!!kicker");
  T(s, [
    { text: "Svi znamo šta je lepo — ", options: {} },
    { text: "dok nas neko ne pita zašto.", options: { italic: true, color: C.gold } },
  ], { x: M, y: 1.15, w: 5.3, h: 2.0, fontFace: HF, fontSize: 34, color: C.white, lineSpacingMultiple: 1.0, objectName: "!!title" });
  T(s, "Lepim nazivamo lice, melodiju, zalazak sunca, pa i matematički dokaz. Šta je zajedničko svim tim stvarima?",
    { x: M, y: 3.45, w: 5.2, h: 1.1, fontSize: 16, color: C.ash, lineSpacingMultiple: 1.15, objectName: "!!sub" });
  T(s, "αἴσθησις", { x: M, y: 4.85, w: 5.2, h: 0.6, fontFace: HF, fontSize: 30, italic: true, color: C.gold });
  T(s, "aisthēsis — grč. „opažanje, čulni doživljaj“. Od nje potiče reč estetika; disciplinu je tako nazvao A. G. Baumgarten (Aesthetica, 1750).",
    { x: M, y: 5.5, w: 5.2, h: 1.0, fontSize: 14, color: C.ash, lineSpacingMultiple: 1.15 });
  s.addNotes("Estetika je filozofska disciplina o lepom i o čulnom doživljaju. Ime dolazi od grčke reči aisthesis — opažanje. Aleksandar Baumgarten ju je 1750. izdvojio kao posebnu nauku. Danas o lepoti govore i filozofi, i istoričari umetnosti, i psiholozi, i neuronaučnici.");

  // 3 — Četiri pitanja
  s = add("Svetli", "Uvod"); TRANS[3] = "fade";
  label(s, "Četiri pitanja", M, 0.7, 6, C.cinn);
  T(s, "Put kroz jedno staro pitanje", { x: M, y: 1.05, w: 11, h: 0.8, fontFace: HF, fontSize: 38, color: C.text });
  const Q = [
    ["Da li je lepota objektivna ili subjektivna?", "Platon · Aristotel · Hjum"],
    ["Da li je lepota u predmetu ili u našem doživljaju?", "Kant · neuroestetika"],
    ["Šta je ukus i zašto se razlikuje?", "Hjum · Zajonc · Burdije"],
    ["Zašto se merila lepote menjaju kroz istoriju i kulture?", "Od Vilendorfa do Dišana"],
  ];
  const colW = 2.7, colGap = (W - 2 * M - 4 * colW) / 3;
  Q.forEach(([q, who], i) => {
    const x = M + i * (colW + colGap);
    T(s, "0" + (i + 1), { x, y: 2.35, w: colW, h: 1.1, fontFace: HF, fontSize: 64, italic: true, color: C.gold, objectName: "!!num" + (i + 1) });
    T(s, q, { x, y: 3.65, w: colW, h: 1.6, fontFace: HF, fontSize: 21, color: C.text, lineSpacingMultiple: 1.05, objectName: "!!q" + (i + 1) });
    T(s, who, { x, y: 5.55, w: colW, h: 0.4, fontSize: 13, color: C.stone });
  });
  s.addNotes("Prezentacija prati četiri pitanja. Svako pitanje je jedna sala ove male izložbe: objektivno ili subjektivno; predmet ili doživljaj; ukus; i promene kroz istoriju i kulture.");

  // ============ SEKCIJA 01 ============
  pres.addSection({ title: "01 Objektivna ili subjektivna" });
  const SEC1 = "01 Objektivna ili subjektivna";

  async function divider(n, title, sub, key, fx, fy, cap, section, transKey, name) {
    const sl = add("Pun kadar", section);
    const ix = 8.4, iw = W - ix;
    sl.addImage({ path: await pic(key, iw, H, fx, fy, 1, "div"), x: ix, y: 0, w: iw, h: H, objectName: name || "!!divimg", altText: cap[1] });
    T(sl, n, { x: M - 0.1, y: 0.35, w: 7, h: 3.4, fontFace: HF, fontSize: 190, italic: true, color: C.gold, objectName: "!!num" + Number(n) });
    T(sl, title, { x: M, y: 3.55, w: 7.3, h: 1.6, fontFace: HF, fontSize: 46, color: C.white, lineSpacingMultiple: 1.0, objectName: "!!q" + Number(n) });
    T(sl, sub, { x: M, y: 5.45, w: 6.6, h: 0.9, fontSize: 17, color: C.ash, lineSpacingMultiple: 1.15 });
    wall(sl, ix + 0.3, 6.95, iw - 0.5, cap[0], cap[1], cap[2], C.white, undefined);
    TRANS[pres.slides.length] = transKey;
    return sl;
  }

  // 4 — Divider 01 (Morph iz 03: broj raste)
  s = await divider("01", "Objektivna ili subjektivna?", "Postoji li lepota nezavisno od nas — ili je, kako kaže izreka, „u oku posmatrača“?",
    "plato", 0.5, 0.42, ["Silanion (rimska kopija)", "Platon", "Kapitolski muzeji, Rim"], SEC1, "morph", "!!plato");
  s.addNotes("Prvo pitanje: da li je lepota objektivna — svojstvo stvari, merljivo i isto za sve — ili subjektivna, zavisna od onoga ko gleda.");

  // 5 — Objektivizam
  s = add("Svetli", SEC1); TRANS[5] = "fade";
  const lw = 5.0;
  s.addImage({ path: await pic("vitruvian", lw, H, 0.5, 0.42), x: 0, y: 0, w: lw, h: H, altText: "Leonardo da Vinči, Vitruvijev čovek" });
  wall(s, 0.4, 6.98, lw - 0.6, "Leonardo da Vinči", "Vitruvijev čovek", "oko 1490, Venecija", C.text);
  const rx = lw + 0.8, rw = W - rx - M;
  label(s, "Stanovište 1  ·  objektivizam", rx, 0.7, rw, C.cinn);
  T(s, "Lepota je u meri i srazmeri", { x: rx, y: 1.05, w: rw, h: 0.9, fontFace: HF, fontSize: 34, color: C.text });
  const OBJ = [
    ["6. v. p. n. e.", "Pitagorejci", "Harmonija je broj: skladni muzički intervali počivaju na prostim razmerama — 2:1, 3:2, 4:3."],
    ["4. v. p. n. e.", "Platon", "Lepo je večna ideja. Pojedinačne lepe stvari samo „učestvuju“ u Lepoti po sebi (Gozba)."],
    ["4. v. p. n. e.", "Aristotel", "„Glavni oblici lepote su red, simetrija i određenost.“ (Metafizika, XIII)"],
  ];
  OBJ.forEach(([when, who, what], i) => {
    const y = 2.4 + i * 1.45;
    T(s, when, { x: rx, y, w: 1.75, h: 0.5, fontFace: HF, fontSize: 18, italic: true, color: C.gold });
    T(s, who, { x: rx + 1.9, y, w: rw - 1.9, h: 0.4, fontSize: 16, bold: true, color: C.text });
    T(s, what, { x: rx + 1.9, y: y + 0.4, w: rw - 1.9, h: 0.85, fontSize: 15, color: C.stone, lineSpacingMultiple: 1.1 });
  });
  s.addNotes("Najstarija tradicija vidi lepotu kao nešto objektivno. Pitagorejci su otkrili da skladni tonovi odgovaraju prostim brojčanim odnosima. Platon smatra da lepota postoji kao večna ideja. Aristotel lepotu vezuje za red, simetriju i određenost. Leonardov Vitruvijev čovek je renesansna slika te ideje: telo upisano u krug i kvadrat.");

  // 6 — Mit o zlatnom preseku
  s = add("Tamni", SEC1); TRANS[6] = "fade";
  const nx = 7.55, ny = 0.95, nw = 5.08, nh = 5.4;
  frame(s, nx, ny, nw, nh);
  s.addImage({ path: await pic("nautilus", nw, nh, 0.5, 0.5), x: nx, y: ny, w: nw, h: nh, altText: "Presek ljušture nautilusa" });
  wall(s, nx, ny + nh + 0.25, nw, "Presek ljušture nautilusa", "", "logaritamska spirala", C.ash);
  label(s, "Provera jednog mita", M, 0.7, 6, C.gold);
  T(s, "Postoji li formula lepote?", { x: M, y: 1.05, w: 6.6, h: 0.9, fontFace: HF, fontSize: 32, color: C.white });
  T(s, "φ ≈ 1,618", { x: M, y: 2.05, w: 6.4, h: 1.5, fontFace: HF, fontSize: 80, italic: true, color: C.gold });
  T(s, "Zlatni presek se često navodi kao univerzalni ključ lepote. Ali psihološka istraživanja sklonosti prema njemu daju mešovite rezultate, a spirala nautilusa — najčešći „dokaz“ — u stvari nije zlatna spirala.",
    { x: M, y: 3.75, w: 6.2, h: 1.7, fontSize: 16, color: C.ash, lineSpacingMultiple: 1.2 });
  T(s, [{ text: "Mera pomaže da razumemo lepo, ", options: {} }, { text: "ali ga ne iscrpljuje.", options: { italic: true, color: C.gold } }],
    { x: M, y: 5.45, w: 6.4, h: 0.9, fontFace: HF, fontSize: 21, color: C.white });
  T(s, "Izvori: Green, Perception, 1995; Falbo, College Mathematics Journal, 2005.", { x: M, y: 6.5, w: 6.4, h: 0.3, fontSize: 10, color: C.ash });
  s.addNotes("Ako je lepota objektivna, da li postoji formula? Zlatni presek je najpoznatiji kandidat. Međutim, pregled istraživanja (Green, 1995) pokazuje da su rezultati mešoviti, a matematičar Falbo (2005) pokazuje da nautilus raste po logaritamskoj spirali koja nije zlatna. Proporcije su važne, ali nisu cela priča.");

  // 7 — Subjektivizam (Hjum)
  s = add("Svetli", SEC1); TRANS[7] = "fade";
  const hx = M + 0.12, hy = 1.0, hw = 3.6, hh = 4.6;
  frame(s, hx, hy, hw, hh, "!!hume_f");
  s.addImage({ path: await pic("hume", hw, hh, 0.5, 0.4), x: hx, y: hy, w: hw, h: hh, objectName: "!!hume", altText: "Alan Remzi, portret Dejvida Hjuma" });
  wall(s, hx - 0.12, hy + hh + 0.3, hw + 0.4, "Alan Remzi", "Dejvid Hjum", "1766", C.text);
  const qx = 5.3, qw = W - qx - M;
  label(s, "Stanovište 2  ·  subjektivizam", qx, 0.7, qw, C.cinn);
  T(s, "“", { x: qx - 0.1, y: 0.75, w: 1.5, h: 1.6, fontFace: HF, fontSize: 140, color: C.cinn });
  T(s, "Lepota nije svojstvo samih stvari: ona postoji samo u umu koji ih posmatra, i svaki um opaža drugačiju lepotu.",
    { x: qx, y: 2.05, w: qw, h: 2.6, fontFace: HF, fontSize: 30, italic: true, color: C.text, lineSpacingMultiple: 1.12 });
  T(s, [{ text: "Dejvid Hjum", options: { bold: true } }, { text: "  ·  O merilu ukusa, 1757." }], { x: qx, y: 4.85, w: qw, h: 0.4, fontSize: 15, color: C.stone });
  s.addShape(pres.shapes.RECTANGLE, { x: qx, y: 5.55, w: qw, h: 0.9, fill: { color: C.card }, line: { color: C.card, width: 0 } });
  T(s, [{ text: "De gustibus non est disputandum", options: { italic: true, color: C.text } }, { text: "  —  o ukusima se ne raspravlja.", options: { color: C.stone } }],
    { x: qx + 0.3, y: 5.55, w: qw - 0.6, h: 0.9, fontSize: 16, valign: "middle" });
  s.addNotes("Suprotno stanovište: lepota je u posmatraču. Škotski filozof Dejvid Hjum piše 1757. da lepota nije svojstvo stvari nego osećanje u umu. Zato svako vidi drugačiju lepotu. Ali — kao što ćemo videti — Hjum se tu ne zaustavlja.");

  // 8 — Šta kažu istraživanja
  s = add("Svetli", SEC1); TRANS[8] = "fade";
  label(s, "Dokazi", M, 0.7, 6, C.cinn);
  T(s, "Šta kažu istraživanja?", { x: M, y: 1.05, w: 11, h: 0.9, fontFace: HF, fontSize: 38, color: C.text });
  T(s, "Koliko se ljudi slažu oko toga šta je lepo — prema domenu", { x: M, y: 2.0, w: 8, h: 0.4, fontSize: 15, color: C.stone });
  // ordinal spectrum
  const ly = 3.35, lx1 = M + 0.3, lx2 = W - M - 0.3;
  s.addShape(pres.shapes.LINE, { x: lx1, y: ly, w: lx2 - lx1, h: 0, line: { color: C.stone, width: 1 } });
  const DOM = [["Lica", 0.72], ["Prirodni pejzaži", 0.6], ["Arhitektura", 0.42], ["Umetnička dela", 0.3]];
  DOM.forEach(([name, d], i) => {
    const cx = lx1 + 0.6 + i * ((lx2 - lx1 - 1.2) / 3);
    s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: ly - d / 2, w: d, h: d, fill: { color: i < 2 ? C.gold : C.card }, line: { color: C.gold, width: 1 } });
    T(s, name, { x: cx - 1.3, y: ly + 0.55, w: 2.6, h: 0.4, fontSize: 16, bold: true, color: C.text, align: "center" });
  });
  T(s, "← više zajedničkog ukusa", { x: lx1, y: 2.55, w: 4, h: 0.3, fontSize: 12, color: C.stone });
  T(s, "više ličnog ukusa →", { x: lx2 - 4, y: 2.55, w: 4, h: 0.3, fontSize: 12, color: C.stone, align: "right" });
  T(s, "Prikazan je redosled, ne tačne vrednosti.", { x: lx2 - 5, y: 4.45, w: 5, h: 0.3, fontSize: 10, italic: true, color: C.stone, align: "right" });
  const cards = [
    ["Priroda nas ujedinjuje", "Oko lica i prirodnih pejzaža ljudi se u velikoj meri slažu; oko umetnosti i arhitekture ukusi se najviše razilaze.", "Vessel i sar., Cognition, 2018"],
    ["Bebe već „biraju“", "Bebe od nekoliko meseci duže gledaju lica koja odrasli ocenjuju kao privlačna — pre nego što nauče kulturne norme.", "Langlois i sar., Developmental Psychology, 1987"],
    ["Zaključak", "Lepota nije ni čisto objektivna ni čisto subjektivna: postoji zajedničko jezgro i širok prostor ličnog ukusa.", ""],
  ];
  const cw3 = (W - 2 * M - 2 * 0.35) / 3;
  cards.forEach(([h, b, src], i) => {
    const x = M + i * (cw3 + 0.35), y = 4.95, dark = i === 2;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: cw3, h: 1.75, fill: { color: dark ? C.ink : C.card }, line: { color: dark ? C.ink : C.card, width: 0 } });
    T(s, h, { x: x + 0.25, y: y + 0.2, w: cw3 - 0.5, h: 0.35, fontFace: HF, fontSize: 17, color: dark ? C.gold : C.text });
    T(s, b, { x: x + 0.25, y: y + 0.6, w: cw3 - 0.5, h: 0.85, fontSize: 13, color: dark ? C.white : C.stone, lineSpacingMultiple: 1.08 });
    if (src) T(s, src, { x: x + 0.25, y: y + 1.45, w: cw3 - 0.5, h: 0.22, fontSize: 10, italic: true, color: C.stone });
  });
  s.addNotes("Empirijska estetika daje nijansiran odgovor. Vessel i saradnici (2018) su pokazali da se ljudi mnogo više slažu oko lepote lica i prirodnih pejzaža nego oko arhitekture i umetničkih dela — kod umetnosti je nečije omiljeno delo često tuđe najmanje omiljeno. Langlois i saradnici (1987) su pokazali da i bebe duže gledaju lica koja odrasli smatraju privlačnim. Dakle: postoji zajedničko jezgro, ali i veliki prostor za lični ukus.");

  // ============ SEKCIJA 02 ============
  const SEC2 = "02 Predmet ili doživljaj";
  pres.addSection({ title: SEC2 });

  // 9 — Divider 02
  s = await divider("02", "U predmetu ili u doživljaju?", "Da li je lepota svojstvo stvari — ili se rađa u načinu na koji ih doživljavamo?",
    "kant", 0.5, 0.4, ["Nepoznati autor", "Imanuel Kant", "oko 1790"], SEC2, "fadeBlack", "!!kant");
  s.addNotes("Drugo pitanje pomera fokus: ako lepota nije samo u stvari ni samo u nama, gde se tačno nalazi? Ključna figura je Imanuel Kant.");

  // 10 — Kant (Morph: portret se uvlači u ram)
  s = add("Tamni", SEC2); TRANS[10] = "morph";
  const kx = M + 0.12, ky = 0.95, kw = 3.3, kh = kw * H / 4.933;
  frame(s, kx, ky, kw, kh, "!!kant_f");
  s.addImage({ path: await pic("kant", 4.933, H, 0.5, 0.4, 1, "div"), x: kx, y: ky, w: kw, h: kh, objectName: "!!kant", altText: "Imanuel Kant" });
  wall(s, kx - 0.12, ky + kh + 0.3, kw + 0.4, "Imanuel Kant", "Kritika moći suđenja", "1790", C.ash);
  const ax = 4.95, aw = W - ax - M;
  label(s, "Imanuel Kant  ·  1790", ax, 0.7, aw, C.gold);
  T(s, [{ text: "„Lepo je ono što se ", options: {} }, { text: "bez pojma", options: { italic: true, color: C.gold } }, { text: " dopada opšte.“", options: {} }],
    { x: ax, y: 1.05, w: aw, h: 1.3, fontFace: HF, fontSize: 32, color: C.white, lineSpacingMultiple: 1.05 });
  const KM = [
    ["Bez interesa", "Uživamo u samom izgledu stvari, a ne u tome da je posedujemo ili iskoristimo."],
    ["Bez pojma, a ipak opšte", "Ne dokazujemo lepotu pravilom, ali očekujemo da se i drugi slože sa nama."],
    ["Svrhovitost bez svrhe", "Delo deluje skladno „smišljeno“, iako ne služi nekoj spoljnoj svrsi."],
    ["Nužno dopadanje", "Sud o lepom pretpostavlja zajedničko čulo (sensus communis) svih ljudi."],
  ];
  const kgw = (aw - 0.35) / 2;
  KM.forEach(([h, b], i) => {
    const x = ax + (i % 2) * (kgw + 0.35), y = 2.75 + Math.floor(i / 2) * 1.85;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: kgw, h: 1.6, fill: { color: C.panel }, line: { color: C.panel, width: 0 } });
    T(s, String(i + 1), { x: x + 0.25, y: y + 0.2, w: 0.5, h: 0.6, fontFace: HF, fontSize: 30, italic: true, color: C.gold });
    T(s, h, { x: x + 0.85, y: y + 0.25, w: kgw - 1.1, h: 0.4, fontSize: 16, bold: true, color: C.white });
    T(s, b, { x: x + 0.85, y: y + 0.7, w: kgw - 1.1, h: 0.8, fontSize: 13, color: C.ash, lineSpacingMultiple: 1.08 });
  });
  s.addNotes("Kant u Kritici moći suđenja (1790) nudi treći put. Sud o lepom je subjektivan — zasniva se na osećanju zadovoljstva, a ne na pojmu — ali ipak polaže pravo na opštost: kada kažem da je nešto lepo, očekujem da se i drugi slože. Četiri momenta: dopadanje bez interesa, opštost bez pojma, svrhovitost bez svrhe i nužnost zasnovana na zajedničkom čulu.");

  // 11 — Tri sastojka (circles apart)
  s = add("Svetli", SEC2); TRANS[11] = "fade";
  label(s, "Odnos, a ne stvar", M, 0.7, 6, C.cinn);
  T(s, "Tri sastojka estetskog doživljaja", { x: M, y: 1.05, w: 11.5, h: 0.9, fontFace: HF, fontSize: 38, color: C.text });
  const CIR = [
    ["Predmet", "simetrija i proporcija\nkontrast i ritam\nmaterijal i boja", C.gold],
    ["Posmatrač", "čula i telo\npamćenje i emocije\nznanje i iskustvo", C.cinn],
    ["Kontekst", "kultura i epoha\nmesto i prilika\nnorme zajednice", C.stone],
  ];
  const cd = 3.3, cgap = (W - 2 * M - 3 * cd) / 2;
  CIR.forEach(([n, items, col], i) => {
    const x = M + i * (cd + cgap), y = 2.35;
    s.addShape(pres.shapes.OVAL, { x, y, w: cd, h: cd, fill: { color: col, transparency: 82 }, line: { color: col, width: 1.25 }, objectName: "!!c" + i });
    T(s, n, { x, y: y + 0.65, w: cd, h: 0.55, fontFace: HF, fontSize: 26, color: C.text, align: "center", objectName: "!!cn" + i });
    T(s, items, { x: x + 0.3, y: y + 1.35, w: cd - 0.6, h: 1.3, fontSize: 14, color: C.stone, align: "center", lineSpacingMultiple: 1.25 });
  });
  T(s, "Svaki ugao ima svoje zastupnike: objektivisti ističu predmet, subjektivisti posmatrača, istoričari i sociolozi kontekst.",
    { x: M, y: 5.95, w: W - 2 * M, h: 0.6, fontSize: 15, color: C.stone, align: "center" });
  s.addNotes("Umesto izbora jedne strane, savremena estetika lepotu vidi kao odnos. U njemu učestvuju tri činioca: svojstva predmeta, posmatrač sa svojim čulima, sećanjima i znanjem, i kontekst — kultura, epoha, prilika.");

  // 12 — Lepota u preseku (Morph: krugovi se spajaju)
  s = add("Svetli", SEC2); TRANS[12] = "morph";
  label(s, "Odnos, a ne stvar", M, 0.7, 6, C.cinn);
  T(s, "Lepota nastaje u susretu", { x: M, y: 1.05, w: 11.5, h: 0.9, fontFace: HF, fontSize: 38, color: C.text });
  const vcx = W / 2, vr = 1.45, vd = 2 * vr;
  const VC = [[vcx - 0.85, 3.65], [vcx + 0.85, 3.65], [vcx, 5.1]];
  VC.forEach(([cx, cy], i) => {
    const col = CIR[i][2];
    s.addShape(pres.shapes.OVAL, { x: cx - vr, y: cy - vr, w: vd, h: vd, fill: { color: col, transparency: 78 }, line: { color: col, width: 1.25 }, objectName: "!!c" + i });
  });
  T(s, "Predmet", { x: vcx - 0.85 - 1.55, y: 2.85, w: 1.9, h: 0.45, fontFace: HF, fontSize: 20, color: C.text, align: "center", objectName: "!!cn0" });
  T(s, "Posmatrač", { x: vcx + 0.85 - 0.5, y: 2.85, w: 1.9, h: 0.45, fontFace: HF, fontSize: 20, color: C.text, align: "center", objectName: "!!cn1" });
  T(s, "Kontekst", { x: vcx - 0.95, y: 5.85, w: 1.9, h: 0.45, fontFace: HF, fontSize: 20, color: C.text, align: "center", objectName: "!!cn2" });
  T(s, "lepota", { x: vcx - 0.8, y: 4.0, w: 1.6, h: 0.5, fontFace: HF, fontSize: 24, italic: true, bold: true, color: C.cinn, align: "center" });
  T(s, "Lepota nije ni samo u stvari ni samo u oku. Ona se rađa kada se svojstva predmeta, doživljaj posmatrača i kontekst susretnu.",
    { x: M, y: 2.6, w: 3.4, h: 2.6, fontSize: 16, color: C.text, lineSpacingMultiple: 1.2 });
  T(s, [{ text: "Primer", options: { bold: true, color: C.text, breakLine: true } },
        { text: "Ista Botičelijeva slika za Firentinca 1485. je učena mitološka alegorija, a za današnjeg turistu — ikona sa razglednice." }],
    { x: W - M - 3.3, y: 2.6, w: 3.3, h: 2.6, fontSize: 14, color: C.stone, lineSpacingMultiple: 1.2 });
  s.addNotes("Kada se tri kruga preklope, dobijamo odgovor na drugo pitanje: lepota nije u jednom od njih, nego u njihovom preseku. Ista slika menja značenje kada se promeni posmatrač ili kontekst.");

  // 13 — Neuroestetika
  s = add("Tamni", SEC2); TRANS[13] = "fade";
  label(s, "Neuroestetika", M, 0.7, 6, C.gold);
  T(s, "Lepota u mozgu posmatrača", { x: M, y: 1.05, w: 11.5, h: 0.9, fontFace: HF, fontSize: 38, color: C.white });
  const STATS = [
    ["21", "ispitanik u fMRI skeneru ocenjivao je lepotu"],
    ["60", "umetničkih dela: 30 slika i 30 muzičkih odlomaka"],
    ["1", "ista oblast reaguje na vizuelnu i muzičku lepotu: medijalni orbitofrontalni korteks"],
  ];
  const sw = (W - 2 * M - 2 * 0.5) / 3;
  STATS.forEach(([n, t], i) => {
    const x = M + i * (sw + 0.5);
    T(s, n, { x, y: 2.25, w: sw, h: 1.5, fontFace: HF, fontSize: 88, italic: true, color: C.gold });
    T(s, t, { x, y: 3.85, w: sw - 0.3, h: 1.0, fontSize: 16, color: C.white, lineSpacingMultiple: 1.15 });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: M, y: 5.2, w: W - 2 * M, h: 1.25, fill: { color: C.panel }, line: { color: C.panel, width: 0 } });
  T(s, [{ text: "Što je doživljaj lepote jači, to je ta aktivnost veća. ", options: { color: C.white } },
        { text: "Ali mozak uvek reaguje na nešto — na oblik, boju, zvuk. Lepota je i u predmetu i u doživljaju.", options: { color: C.ash } }],
    { x: M + 0.3, y: 5.2, w: W - 2 * M - 0.6, h: 1.25, fontSize: 16, valign: "middle", lineSpacingMultiple: 1.15 });
  T(s, "Izvor: Ishizu i Zeki, „Toward a Brain-Based Theory of Beauty“, PLoS ONE, 2011.", { x: M, y: 6.55, w: 9, h: 0.3, fontSize: 10, color: C.ash });
  s.addNotes("Neuronauka daje zanimljiv argument. Ishizu i Zeki (2011) su snimali mozak 21 ispitanika dok su gledali slike i slušali muziku. Kod oba čula, doživljaj lepote bio je povezan sa aktivnošću u istoj oblasti — medijalnom orbitofrontalnom korteksu — i to proporcionalno jačini doživljaja. Sami autori naglašavaju da je teorija preliminarna.");

  // ============ SEKCIJA 03 ============
  const SEC3 = "03 Ukus";
  pres.addSection({ title: SEC3 });

  // 14 — Divider 03
  s = await divider("03", "Ukus i različita shvatanja lepote", "Zašto se oko lepote toliko razlikujemo — i da li je svaki ukus podjednako dobar?",
    "fountain", 0.5, 0.5, ["Marsel Dišan", "Fontana", "1917, fotografija A. Stiglica"], SEC3, "fadeBlack", "!!divimg3");
  s.addNotes("Treće pitanje je pitanje ukusa. Dišanova Fontana — fabrički pisoar potpisan pseudonimom R. Mutt — i danas izaziva podele: da li je to umetnost? Da li je lepo? Upravo takva neslaganja su tema ukusa.");

  // 15 — Hjumov pravi kritičar
  s = add("Svetli", SEC3); TRANS[15] = "fade";
  label(s, "Merilo ukusa", M, 0.7, 6, C.cinn);
  T(s, "Postoji li „dobar“ ukus?", { x: M, y: 1.05, w: 11.5, h: 0.9, fontFace: HF, fontSize: 38, color: C.text });
  T(s, "Hjum priznaje da se ukusi razlikuju, ali tvrdi da nisu svi jednako vredni. Merilo je saglasan sud „pravih kritičara“ — ljudi koji imaju pet osobina:",
    { x: M, y: 2.0, w: 10.5, h: 0.8, fontSize: 16, color: C.stone, lineSpacingMultiple: 1.15 });
  const HQ = [
    ["Zdrav razum", "razume svrhu i celinu dela"],
    ["Istančan osećaj", "zapaža najfinije razlike"],
    ["Vežbu", "ukus se izoštrava čestim susretom sa delima"],
    ["Poređenje", "poznaje mnoga dela i epohe"],
    ["Slobodu od predrasuda", "ne sudi iz interesa ni mode"],
  ];
  const hw5 = (W - 2 * M - 4 * 0.3) / 5;
  HQ.forEach(([h, b], i) => {
    const x = M + i * (hw5 + 0.3), y = 3.1, d = 1.0;
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, line: { color: C.gold, width: 1.25 } });
    T(s, String(i + 1), { x, y, w: d, h: d, fontFace: HF, fontSize: 30, italic: true, color: C.gold, align: "center", valign: "middle" });
    T(s, h, { x, y: y + 1.25, w: hw5, h: 0.65, fontSize: 16, bold: true, color: C.text });
    T(s, b, { x, y: y + 1.9, w: hw5, h: 0.8, fontSize: 14, color: C.stone, lineSpacingMultiple: 1.1 });
  });
  T(s, [{ text: "„Snažan razum, sjedinjen sa istančanim osećanjem, unapređen vežbom, usavršen poređenjem i očišćen od svih predrasuda.“", options: { italic: true } },
        { text: "   — Hjum, 1757.", options: { color: C.stone } }],
    { x: M, y: 6.05, w: W - 2 * M, h: 0.6, fontFace: HF, fontSize: 16, color: C.text });
  s.addNotes("Hjum, koji je rekao da je lepota u umu, ipak ne prihvata da je svaki ukus jednako dobar. Neki sudovi su bolji od drugih. Merilo ukusa je saglasnost iskusnih kritičara sa pet osobina: zdrav razum, istančan osećaj, vežba, poređenje i oslobođenost od predrasuda. Ukus se, dakle, može razvijati.");

  // 16 — Kako nastaje ukus
  s = add("Tamni", SEC3); TRANS[16] = "fade";
  label(s, "Kako nastaje ukus", M, 0.7, 6, C.gold);
  T(s, "Ukus se uči", { x: M, y: 1.05, w: 11.5, h: 0.9, fontFace: HF, fontSize: 38, color: C.white });
  const TF = [
    ["Navika", "Efekat pukog izlaganja: ono što češće viđamo i slušamo, obično nam se više dopada.", "R. Zajonc, 1968"],
    ["Obrazovanje", "Iskustvo, vežba i poređenje izoštravaju opažanje — vidimo ono što ranije nismo primećivali.", "D. Hjum, 1757"],
    ["Društvo", "„Ukus klasifikuje, i klasifikuje onoga koji klasifikuje.“ Ukus je i znak pripadnosti društvenoj grupi.", "P. Burdije, Distinkcija, 1979"],
  ];
  const tw = (W - 2 * M - 2 * 0.4) / 3;
  TF.forEach(([h, b, src], i) => {
    const x = M + i * (tw + 0.4), y = 2.3;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: tw, h: 4.2, fill: { color: C.panel }, line: { color: C.panel, width: 0 } });
    T(s, "0" + (i + 1), { x: x + 0.4, y: y + 0.35, w: 2, h: 1.0, fontFace: HF, fontSize: 54, italic: true, color: C.gold });
    T(s, h, { x: x + 0.4, y: y + 1.55, w: tw - 0.8, h: 0.6, fontFace: HF, fontSize: 26, color: C.white });
    T(s, b, { x: x + 0.4, y: y + 2.25, w: tw - 0.8, h: 1.3, fontSize: 15, color: C.ash, lineSpacingMultiple: 1.15 });
    T(s, src, { x: x + 0.4, y: y + 3.65, w: tw - 0.8, h: 0.3, fontSize: 11, italic: true, color: C.gold });
  });
  s.addNotes("Tri sile oblikuju ukus. Navika: psiholog Robert Zajonc je pokazao da ponovljeno izlaganje povećava dopadanje. Obrazovanje: kao kod Hjuma, vežba i poređenje izoštravaju sud. Društvo: sociolog Pjer Burdije pokazuje da ukus označava društveni položaj — ukusom se razlikujemo od drugih i pripadamo svojima.");

  // ============ SEKCIJA 04 ============
  const SEC4 = "04 Istorija i kulture";
  pres.addSection({ title: SEC4 });

  // 17 — Divider 04
  s = await divider("04", "Kroz istoriju i kulture", "Zašto ono što je lepo u jednom vremenu i jednoj kulturi, u drugoj ne mora biti?",
    "wave", 0.3, 0.5, ["Kacušika Hokusai", "Veliki talas kod Kanagave", "oko 1831"], SEC4, "fadeBlack", "!!divimg4");
  s.addNotes("Četvrto pitanje: merila lepote nisu stalna. Prvo ćemo proći kroz šest tačaka na vremenskoj liniji, a zatim uporediti dve kulture.");

  // 18–19 — Vremenska linija (Morph: fokus klizi)
  const TL = [
    ["willendorf", "pre ~30.000 god.", "Vilendorfska Venera", "Obilje i plodnost: telo kao znak opstanka.", 0.5, 0.5],
    ["doryphoros", "oko 440. p. n. e.", "Polikleit, Dorifor", "Kanon: idealne proporcije tela izražene brojem.", 0.5, 0.35],
    ["venus", "oko 1485.", "Botičeli, Rođenje Venere", "Renesansa: povratak antičkom idealu sklada.", 0.5, 0.35],
    ["graces", "1630–1635.", "Rubens, Tri gracije", "Barok: puna, raskošna tela, pokret i čulnost.", 0.5, 0.4],
    ["impression", "1872.", "Mone, Impresija, izlazak sunca", "Lepota trenutka: svetlost i utisak umesto savršenog oblika.", 0.5, 0.5],
    ["fountain", "1917.", "Dišan, Fontana", "Avangarda: umetnost ne mora biti lepa — mora da pokrene misao.", 0.5, 0.5],
  ];
  const trY = 6.15, trX1 = M, trX2 = W - M;
  const dotX = (i) => trX1 + 0.35 + i * ((trX2 - trX1 - 0.7) / 5);
  const cardW = 3.7, cardGap = (W - 2 * M - 3 * cardW) / 2, imgH = 2.3;
  for (const part of [0, 1]) {
    s = add("Svetli", SEC4); TRANS[pres.slides.length] = part === 0 ? "fade" : "morph";
    label(s, "Vremenska linija", M, 0.7, 6, C.cinn, "!!tl_k");
    T(s, part === 0 ? "Lepota kroz vreme: od plodnosti do sklada" : "Lepota kroz vreme: od raskoši do ideje",
      { x: M, y: 1.05, w: 11.8, h: 0.8, fontFace: HF, fontSize: 34, color: C.text, objectName: "!!tl_t" });
    // focus band behind the three active dots
    const fx1 = dotX(part * 3) - 0.35, fx2 = dotX(part * 3 + 2) + 0.35;
    s.addShape(pres.shapes.RECTANGLE, { x: fx1, y: trY - 0.22, w: fx2 - fx1, h: 0.44, fill: { color: C.card }, line: { color: C.card, width: 0 }, objectName: "!!tl_focus" });
    s.addShape(pres.shapes.LINE, { x: trX1, y: trY, w: trX2 - trX1, h: 0, line: { color: C.stone, width: 1 }, objectName: "!!tl_line" });
    TL.forEach((it, i) => {
      const on = Math.floor(i / 3) === part, d = on ? 0.2 : 0.14;
      s.addShape(pres.shapes.OVAL, { x: dotX(i) - d / 2, y: trY - d / 2, w: d, h: d, fill: { color: on ? C.cinn : C.stone }, line: { color: on ? C.cinn : C.stone, width: 0 }, objectName: "!!tl_dot" + i });
      T(s, it[1], { x: dotX(i) - 0.9, y: trY + 0.25, w: 1.8, h: 0.3, fontSize: 11, bold: on, color: on ? C.text : C.stone, align: "center", objectName: "!!tl_y" + i });
    });
    for (let j = 0; j < 3; j++) {
      const it = TL[part * 3 + j];
      const x = M + j * (cardW + cardGap), y = 1.95;
      s.addImage({ path: await pic(it[0], cardW, imgH, it[4], it[5], 1, "tl"), x, y, w: cardW, h: imgH, altText: it[2] });
      T(s, it[1], { x, y: y + imgH + 0.2, w: cardW, h: 0.45, fontFace: HF, fontSize: 20, italic: true, color: C.cinn });
      T(s, it[2], { x, y: y + imgH + 0.62, w: cardW, h: 0.35, fontSize: 14, bold: true, color: C.text });
      T(s, it[3], { x, y: y + imgH + 0.98, w: cardW, h: 0.55, fontSize: 13, color: C.stone, lineSpacingMultiple: 1.05 });
    }
    s.addNotes(part === 0
      ? "Najstarije predstave, poput Vilendorfske Venere stare oko 30.000 godina, slave obilje i plodnost. Grčki vajar Polikleit u 5. veku p. n. e. stvara Kanon — sistem idealnih proporcija. Renesansa, sa Botičelijem, vraća antički ideal sklada."
      : "Barok, kod Rubensa, slavi puna, raskošna tela i pokret. Mone 1872. slika utisak svetlosti — po ovoj slici je impresionizam dobio ime. Dišan 1917. izlaže fabrički pisoar i postavlja pitanje da li umetnost uopšte mora biti lepa.");
  }

  // 20 — Dve kulture (split screen)
  s = add("Pun kadar", SEC4); TRANS[20] = "fade";
  const hw2 = W / 2;
  s.addImage({ path: await pic("parthenon", hw2, H, 0.5, 0.5), x: 0, y: 0, w: hw2, h: H, altText: "Partenon, Atina" });
  s.addImage({ path: await pic("ryoanji", hw2, H, 0.5, 0.5), x: hw2, y: 0, w: hw2, h: H, altText: "Bašta Rjoan-đi, Kjoto" });
  s.addImage({ path: gHalfBottom, x: 0, y: 0, w: hw2, h: H });
  s.addImage({ path: gHalfBottom, x: hw2, y: 0, w: hw2, h: H });
  T(s, "Dva ideala lepote", { x: M, y: 0.55, w: W - 2 * M, h: 0.8, fontFace: HF, fontSize: 36, color: C.white, align: "center" });
  const CU = [
    ["Grčka  ·  5. v. p. n. e.", "Savršen oblik", "Simetrija, proporcija i jasna mera. Partenon (447–432. p. n. e.) je oličenje harmonije celine i delova.", "Partenon, Atina"],
    ["Japan  ·  vabi-sabi", "Lepota nesavršenog", "Jednostavnost, prolaznost, tragovi vremena. U bašti Rjoan-đi (Kjoto, kraj 15. v.) od 15 kamenova uvek je bar jedan skriven pogledu.", "Rjoan-đi, Kjoto"],
  ];
  CU.forEach(([k, h, b, cap], i) => {
    const x = i * hw2 + M, w = hw2 - 2 * M;
    label(s, k, x, 4.35, w, C.gold);
    T(s, h, { x, y: 4.7, w, h: 0.7, fontFace: HF, fontSize: 30, italic: true, color: C.white });
    T(s, b, { x, y: 5.45, w, h: 1.1, fontSize: 15, color: C.white, lineSpacingMultiple: 1.15 });
    T(s, cap, { x, y: 6.85, w, h: 0.3, fontSize: 10, color: C.ash });
  });
  s.addNotes("Kulture ne razlikuju samo ukuse, nego i same ideje o tome šta lepota jeste. Grčka klasika traži savršen, simetričan i proporcionalan oblik. Japanska estetika vabi-sabi nalazi lepotu u nesavršenom, prolaznom i jednostavnom — kao u bašti Rjoan-đi, gde se sa terase nikada ne vidi svih petnaest kamenova odjednom.");

  // 21 — Zašto se merila menjaju
  s = add("Svetli", SEC4); TRANS[21] = "fade";
  const ix0 = M + 0.12, iw0 = 3.1, ih0 = 2.45;
  frame(s, ix0, 0.82, iw0, ih0);
  s.addImage({ path: await pic("elizabeth", iw0, ih0, 0.5, 0.3), x: ix0, y: 0.82, w: iw0, h: ih0, altText: "Elizabeta I, portret Armada" });
  T(s, [{ text: "Elizabeta I", options: { bold: true } }, { text: "  „Armada“ portret, oko 1588.", options: { italic: true } }], { x: M, y: 3.42, w: iw0 + 0.3, h: 0.3, fontSize: 10, color: C.text });
  frame(s, ix0, 4.07, iw0, ih0);
  s.addImage({ path: await pic("plum", iw0, ih0, 0.5, 0.5), x: ix0, y: 4.07, w: iw0, h: ih0, altText: "Van Gog, Rascvetali šljivik (po Hirošigeu)" });
  T(s, [{ text: "Vinsent van Gog", options: { bold: true } }, { text: "  Rascvetali šljivik (po Hirošigeu), 1887.", options: { italic: true } }], { x: M, y: 6.67, w: iw0 + 0.9, h: 0.3, fontSize: 10, color: C.text });
  const gx = 4.85, gw = W - gx - M;
  label(s, "Pokretači promene", gx, 0.7, gw, C.cinn);
  T(s, "Zašto se merila menjaju?", { x: gx, y: 1.05, w: gw, h: 0.9, fontFace: HF, fontSize: 36, color: C.text });
  const WHY = [
    ["Status i ekonomija", "Dok je rad značio rad na polju, bleda koža bila je znak privilegije; u 20. veku preplanulost postaje znak dokolice."],
    ["Tehnologija i mediji", "Fotografija (1839) oslobađa slikarstvo verne kopije stvarnosti; danas društvene mreže i filteri brzo šire nove ideale."],
    ["Ideje i verovanja", "Srednji vek traži duhovnu lepotu, renesansa ljudsku meru, romantizam uzvišeno i divlju prirodu."],
    ["Susret kultura", "Japanski drvorezi menjaju evropsku umetnost (japonizam): Van Gog kopira Hirošigea, a Mone sakuplja japanske grafike."],
  ];
  const gcw = (gw - 0.35) / 2;
  WHY.forEach(([h, b], i) => {
    const x = gx + (i % 2) * (gcw + 0.35), y = 2.2 + Math.floor(i / 2) * 2.25;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: gcw, h: 2.0, fill: { color: C.card }, line: { color: C.card, width: 0 } });
    T(s, "0" + (i + 1), { x: x + 0.3, y: y + 0.22, w: 0.8, h: 0.5, fontFace: HF, fontSize: 24, italic: true, color: C.cinn });
    T(s, h, { x: x + 1.0, y: y + 0.27, w: gcw - 1.3, h: 0.4, fontSize: 16, bold: true, color: C.text });
    T(s, b, { x: x + 0.3, y: y + 0.8, w: gcw - 0.6, h: 1.1, fontSize: 13, color: C.stone, lineSpacingMultiple: 1.1 });
  });
  s.addNotes("Zašto se merila menjaju? Prvo, status: kada je većina radila napolju, bleda koža — kao kod Elizabete I — značila je privilegiju; kasnije preplanulost postaje znak slobodnog vremena. Drugo, tehnologija: fotografija oslobađa slikarstvo, a danas mreže šire ideale brže nego ikad. Treće, ideje i religija svake epohe. Četvrto, susret kultura: japanske grafike su duboko promenile evropsku umetnost — Van Gog je ručno kopirao Hirošigea.");

  // ============ SEKCIJA: Zaključak ============
  const SEC5 = "Zaključak";
  pres.addSection({ title: SEC5 });

  // 22 — Ključne poruke
  s = add("Tamni", SEC5); TRANS[22] = "fadeBlack";
  label(s, "Odgovori", M, 0.7, 6, C.gold);
  T(s, "Četiri pitanja, četiri odgovora", { x: M, y: 1.05, w: 11, h: 0.8, fontFace: HF, fontSize: 38, color: C.white });
  const ANS = [
    "Ni čisto objektivna ni čisto subjektivna: postoji zajedničko jezgro i širok prostor ličnog ukusa.",
    "Lepota nastaje u susretu svojstava predmeta, doživljaja posmatrača i konteksta.",
    "Ukus je i dar i navika: oblikuju ga iskustvo, obrazovanje i društvo.",
    "Merila se menjaju jer se menjaju potrebe, tehnologije, ideje i susreti kultura.",
  ];
  ANS.forEach((a, i) => {
    const x = M + i * (colW + colGap);
    T(s, "0" + (i + 1), { x, y: 2.35, w: colW, h: 1.1, fontFace: HF, fontSize: 64, italic: true, color: C.gold });
    T(s, a, { x, y: 3.65, w: colW, h: 2.4, fontFace: HF, fontSize: 19, color: C.white, lineSpacingMultiple: 1.12 });
  });
  s.addNotes("Sažetak u četiri rečenice — po jedna za svako pitanje sa početka.");

  // 23 — Završna poruka
  s = add("Pun kadar", SEC5); TRANS[23] = "fade";
  const wx = 7.4, ww = W - wx;
  s.addImage({ path: await pic("wanderer", ww, H, 0.5, 0.45), x: wx, y: 0, w: ww, h: H, altText: "Kaspar David Fridrih, Putnik iznad mora magle" });
  T(s, "“", { x: M - 0.1, y: 0.6, w: 1.5, h: 1.6, fontFace: HF, fontSize: 140, color: C.gold });
  T(s, "Lepota će spasti svet.", { x: M, y: 2.0, w: 6.2, h: 2.0, fontFace: HF, fontSize: 54, italic: true, color: C.white, lineSpacingMultiple: 1.0 });
  T(s, [{ text: "F. M. Dostojevski", options: { bold: true } }, { text: "  ·  Idiot, 1869." }], { x: M, y: 4.15, w: 6, h: 0.4, fontSize: 16, color: C.ash });
  T(s, "A šta je za vas lepo — i zašto?", { x: M, y: 5.45, w: 6.2, h: 0.6, fontFace: HF, fontSize: 24, color: C.gold });
  T(s, "Hvala na pažnji.", { x: M, y: 6.1, w: 6, h: 0.4, fontSize: 16, color: C.ash });
  wall(s, wx + 0.3, 6.95, ww - 0.5, "Kaspar David Fridrih", "Putnik iznad mora magle", "oko 1818, Hamburg", C.white);
  s.addNotes("Završna misao iz Dostojevskog Idiota: lepota će spasti svet. Možda zato što je lepota upravo ono mesto na kome se svet i čovek susreću. Otvaram diskusiju: šta je za vas lepo — i zašto?");

  // 24 — Izvori
  s = add("Svetli", SEC5); TRANS[24] = "fade";
  label(s, "Izvori", M, 0.7, 6, C.cinn);
  T(s, "Literatura i slike", { x: M, y: 1.05, w: 11, h: 0.8, fontFace: HF, fontSize: 34, color: C.text });
  const LIT = [
    "Aristotel, Metafizika, knj. XIII (M), 1078a.",
    "Platon, Gozba; Hipija Veći.",
    "Hume, D. (1757). Of the Standard of Taste. Four Dissertations.",
    "Kant, I. (1790). Kritik der Urteilskraft (Kritika moći suđenja), §§ 1–22.",
    "Baumgarten, A. G. (1750). Aesthetica.",
    "Bourdieu, P. (1979). La Distinction. Paris: Minuit.",
    "Zajonc, R. B. (1968). Attitudinal effects of mere exposure. JPSP, 9(2).",
    "Langlois, J. H. i sar. (1987). Infant preferences for attractive faces. Developmental Psychology, 23(3).",
    "Vessel, E. A., Maurer, N., Denker, A. H., Starr, G. G. (2018). Stronger shared taste for natural aesthetic domains than for artifacts of human culture. Cognition, 179, 121–131.",
    "Ishizu, T., Zeki, S. (2011). Toward a brain-based theory of beauty. PLoS ONE, 6(7), e21852.",
    "Green, C. D. (1995). All that glitters: a review of psychological research on the aesthetics of the golden section. Perception, 24.",
    "Falbo, C. (2005). The golden ratio: a contrary viewpoint. College Mathematics Journal, 36(2).",
    "Stanford Encyclopedia of Philosophy: „Hume’s Aesthetics“; „Kant’s Aesthetics and Teleology“.",
  ];
  T(s, LIT.map((t, i) => ({ text: t, options: { breakLine: i < LIT.length - 1 } })),
    { x: M, y: 2.0, w: 6.6, h: 4.8, fontSize: 10, color: C.text, paraSpaceAfter: 4, lineSpacingMultiple: 1.05 });
  const order = ["venus", "plato", "vitruvian", "nautilus", "hume", "kant", "fountain", "wave", "willendorf", "doryphoros", "graces", "impression", "parthenon", "ryoanji", "elizabeth", "plum", "wanderer"];
  const strip = (h) => String(h || "").replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim();
  const IMG = order.filter((k) => meta[k]).map((k) => {
    const m = meta[k];
    let lic = strip(m.license) || "javno vlasništvo";
    if (/public domain/i.test(lic)) lic = "javno vlasništvo";
    const art = strip(m.artist).slice(0, 60);
    return `${m.title.replace(/\.(jpe?g|png|tif+)$/i, "")} — ${art ? art + ", " : ""}${lic}`;
  });
  label(s, "Slike: Wikimedia Commons", 7.65, 2.0, 5, C.stone);
  T(s, IMG.map((t, i) => ({ text: t, options: { breakLine: i < IMG.length - 1 } })),
    { x: 7.65, y: 2.35, w: W - 7.65 - M, h: 4.45, fontSize: 9, color: C.stone, paraSpaceAfter: 2 });
  s.addNotes("Svi navodi i slike. Slike su preuzete sa Wikimedia Commons; licence su navedene uz svaku.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  fs.writeFileSync(path.join(__dirname, "transitions.json"), JSON.stringify(TRANS));
  console.log("wrote", OUT, pres.slides.length, "slides");
})().catch((e) => { console.error(e); process.exit(1); });
