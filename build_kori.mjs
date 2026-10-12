import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const workspaceDir = '/Users/atom/Documents/Codex/2026-10-12/fai';
const SKILL_DIR = '/Users/atom/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/Presentations';
const TMP_DIR = path.join(workspaceDir, 'work/build');
const FINAL_PPTX = path.join(workspaceDir, 'outputs/KORI_Objet_01_Diaporama_Canvas_final.pptx');
const CHARTER = '/Users/atom/Documents/KORI_Charte_graphique.pdf';
const SOURCE = '/Users/atom/Downloads/Diaporama Café Premium Marketing-2.pdf';
const HERO = path.join(workspaceDir, 'work/assets/kori-hero.png');
const DETAIL = path.join(workspaceDir, 'work/assets/kori-detail.png');
const OUT_W = 1440, OUT_H = 810;
const C = {
  pine: '#17201C', paper: '#EBE7E0', ember: '#E8641C', concrete: '#CFCAC2',
  graphite: '#757472', ink: '#171A18', moss: '#B9C2BC', stone: '#5E5A54', emberDark: '#A3440F',
  white: '#F8F6F2', border: '#C7C1B7', muted: '#D7D2CA', darkText: '#292B28',
};
const DISPLAY = 'Newsreader';
const BODY = 'Jost';
const noLine = { style: 'solid', fill: 'none', width: 0 };
const presentation = Presentation.create({ slideSize: { width: OUT_W, height: OUT_H } });

function rect(slide, x, y, w, h, fill, lineFill = 'none', lineWidth = 0, name) {
  return slide.shapes.add({
    geometry: 'rect', name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: 'solid', fill: lineFill, width: lineWidth },
  });
}
function rule(slide, x, y, w, h = 2, fill = C.ember) {
  return rect(slide, x, y, w, h, fill, 'none', 0);
}
function txt(slide, text, x, y, w, h, opts = {}) {
  const shape = slide.shapes.add({
    geometry: 'textbox', name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill: 'none', line: noLine,
  });
  shape.text = text;
  shape.text.style = {
    typeface: opts.font ?? BODY,
    fontSize: opts.size ?? 20,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    color: opts.color ?? C.ink,
    alignment: opts.align ?? 'left',
    verticalAlignment: opts.valign ?? 'top',
    wrap: 'square', autoFit: 'none',
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return shape;
}
function logo(slide, x = 72, y = 42, color = C.ink, size = 23) {
  return txt(slide, 'K Ō R I', x, y, 210, 34, { font: BODY, size, color, name: 'kori-wordmark' });
}
function header(slide, section, title, dark = false) {
  const primary = dark ? C.paper : C.ink;
  const secondary = dark ? C.moss : C.graphite;
  logo(slide, 72, 34, primary, 20);
  txt(slide, section.toUpperCase(), 72, 102, 490, 26, { font: BODY, size: 16, bold: true, color: C.ember });
  txt(slide, title, 72, 140, 1180, 72, { font: DISPLAY, size: 48, color: primary, name: 'slide-title' });
  rule(slide, 72, 228, 64, 2, C.ember);
  return { primary, secondary };
}
function footer(slide, n, dark = false) {
  const col = dark ? C.moss : C.graphite;
  txt(slide, 'KŌRI · OBJET 01', 72, 774, 320, 18, { font: BODY, size: 13, color: col });
  txt(slide, `${String(n).padStart(2, '0')} / 06`, 1290, 774, 80, 18, { font: BODY, size: 13, color: col, align: 'right' });
}
function notes(slide, text) { slide.speakerNotes.textFrame.setText(text); }
function addImage(slide, bytes, alt, x, y, w, h) {
  return slide.images.add({ blob: bytes, contentType: 'image/png', alt, fit: 'cover', position: { left: x, top: y, width: w, height: h } });
}
function canvasCell(slide, x, y, w, h, num, label, body, opts = {}) {
  rect(slide, x, y, w, h, opts.fill ?? '#F2EFE9', C.border, 1, `canvas-cell-${num}`);
  txt(slide, String(num).padStart(2, '0'), x + 14, y + 12, 28, 21, { font: BODY, size: 15, bold: true, color: C.ember, name: `canvas-number-${num}` });
  txt(slide, label.toUpperCase(), x + 47, y + 11, w - 60, 24, { font: BODY, size: opts.headerSize ?? 16, bold: true, color: C.pine, name: `canvas-heading-${num}` });
  rule(slide, x + 14, y + 40, 34, 1.5, C.concrete);
  txt(slide, body, x + 14, y + 51, w - 28, h - 59, { font: BODY, size: opts.bodySize ?? 15.5, color: C.darkText, name: `canvas-body-${num}` });
}

const heroBytes = new Uint8Array(await fs.readFile(HERO));
const detailBytes = new Uint8Array(await fs.readFile(DETAIL));

// 1 — Cover
{
  const slide = presentation.slides.add();
  addImage(slide, heroBytes, 'Machine à espresso KŌRI en béton, photographiée en lumière latérale', 0, 0, OUT_W, OUT_H);
  logo(slide, 84, 48, C.paper, 25);
  txt(slide, 'OBJET 01', 86, 214, 340, 25, { font: BODY, size: 17, bold: true, color: C.ember });
  txt(slide, 'La machine,\nredéfinie.', 82, 264, 610, 196, { font: DISPLAY, size: 67, color: C.paper, name: 'cover-title' });
  rule(slide, 86, 493, 68, 2, C.paper);
  txt(slide, 'Une machine à espresso en béton\npensée pour le rituel.', 86, 519, 510, 76, { font: BODY, size: 22, color: C.paper });
  txt(slide, 'BÉTON COULÉ  ·  ACIER BROSSÉ  ·  NOYER', 86, 741, 580, 24, { font: BODY, size: 14, color: C.moss });
  txt(slide, '01 / 06', 1280, 759, 80, 20, { font: BODY, size: 13, color: C.paper, align: 'right' });
  notes(slide, `Sources : KORI_Charte_graphique.pdf, pages 3, 7, 12 et 14. Visuel produit généré à partir de l’image de référence de la machine fournie par l’utilisateur.`);
}

// 2 — Product
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  addImage(slide, detailBytes, 'Détail du groupe d’extraction, du béton et de la poignée en noyer', 850, 0, 590, 810);
  logo(slide, 72, 36, C.ink, 20);
  txt(slide, '01 · OBJET 01', 72, 102, 400, 25, { font: BODY, size: 16, bold: true, color: C.ember });
  txt(slide, 'Une machine\ncoulée en béton', 72, 145, 690, 143, { font: DISPLAY, size: 52, color: C.ink, name: 'product-title' });
  txt(slide, 'Béton, acier brossé et poignée en noyer.\nLa matière reste visible, le geste reste simple.', 74, 312, 680, 65, { font: BODY, size: 20, color: C.stone });
  rule(slide, 74, 417, 680, 1.5, C.concrete);
  const stats = [
    { x: 74, value: '93 °C', label: 'TEMPÉRATURE' },
    { x: 302, value: '9 bar', label: 'PRESSION' },
    { x: 530, value: '58 mm', label: 'PORTE-FILTRE' },
  ];
  for (const s of stats) {
    txt(slide, s.value, s.x, 447, 210, 60, { font: DISPLAY, size: 42, color: C.ember });
    txt(slide, s.label, s.x, 509, 210, 24, { font: BODY, size: 15, bold: true, color: C.graphite });
  }
  rule(slide, 74, 577, 680, 1.5, C.concrete);
  txt(slide, 'Corps coulé à la main, séché 28 jours puis scellé.', 74, 600, 680, 52, { font: BODY, size: 19, color: C.ink });
  txt(slide, 'Source des spécifications : charte KŌRI, édition 01.', 74, 722, 680, 24, { font: BODY, size: 14, color: C.graphite });
  footer(slide, 2);
  notes(slide, 'Sources : KORI_Charte_graphique.pdf, pages 3, 7, 9, 10 et 11. Les valeurs 93 °C, 9 bar, 58 mm et le séchage de 28 jours proviennent de la charte KŌRI.');
}

// 3 — Value proposition
{
  const slide = presentation.slides.add();
  slide.background.fill = C.pine;
  const { primary, secondary } = header(slide, '02 · Proposition de valeur', 'Proposition de valeur', true);
  txt(slide, 'Le client choisit un modèle et des options, puis visualise son projet en 3D.', 74, 258, 1220, 74, { font: DISPLAY, size: 31, color: C.paper });
  rule(slide, 74, 367, 1290, 1, '#5E665F');
  txt(slide, 'B2C', 74, 405, 170, 52, { font: DISPLAY, size: 42, color: C.ember });
  txt(slide, 'PARTICULIERS', 74, 466, 480, 28, { font: BODY, size: 17, bold: true, color: C.moss });
  rule(slide, 74, 511, 44, 2, C.ember);
  txt(slide, 'Une machine premium et durable,\npersonnalisée selon les goûts et les envies.\n\nPour les amateurs de café et les clients\nsensibles au design.', 74, 540, 565, 174, { font: BODY, size: 21, color: C.paper });
  rect(slide, 710, 397, 1.5, 318, '#5E665F');
  txt(slide, 'B2B', 768, 405, 170, 52, { font: DISPLAY, size: 42, color: C.ember });
  txt(slide, 'HÔTELLERIE', 768, 466, 480, 28, { font: BODY, size: 17, bold: true, color: C.moss });
  rule(slide, 768, 511, 44, 2, C.ember);
  txt(slide, 'Une solution haut de gamme pour les\nhôtels prestigieux qui souhaitent proposer\nune expérience café premium à leurs clients.', 768, 540, 550, 140, { font: BODY, size: 21, color: C.paper });
  footer(slide, 3, true);
  notes(slide, 'Sources : Diaporama Café Premium Marketing-2.pdf, page 5; KORI_Charte_graphique.pdf, pages 3, 7, 11 et 12. Le choix du modèle, les options et la visualisation 3D sont repris du canvas fourni.');
}

// 4 — Distribution
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  const { primary, secondary } = header(slide, '03 · Distribution', 'Un circuit direct, deux canaux', false);
  txt(slide, 'Une expérience pensée pour les\nparticuliers et l’hôtellerie haut de gamme.', 74, 262, 580, 84, { font: DISPLAY, size: 30, color: C.stone });
  txt(slide, 'B2C  ·  PARTICULIERS', 74, 375, 520, 25, { font: BODY, size: 16, bold: true, color: C.ember });
  txt(slide, 'B2B  ·  HÔTELS HAUT DE GAMME', 74, 413, 520, 25, { font: BODY, size: 16, bold: true, color: C.ember });
  rule(slide, 74, 472, 580, 1.5, C.concrete);
  txt(slide, 'OMNICANAL', 74, 508, 510, 28, { font: BODY, size: 16, bold: true, color: C.ember });
  txt(slide, 'Vente directe sur le site et en boutique.', 74, 552, 560, 72, { font: DISPLAY, size: 29, color: C.pine });
  // Editorial channel divider and second channel
  rect(slide, 706, 265, 1.5, 438, C.concrete);
  txt(slide, '01', 758, 270, 90, 45, { font: DISPLAY, size: 35, color: C.ember });
  txt(slide, 'SITE WEB', 758, 330, 510, 40, { font: BODY, size: 25, bold: true, color: C.pine });
  rule(slide, 758, 383, 46, 2, C.ember);
  txt(slide, 'Visualisation 3D\nPersonnalisation des modèles et options', 758, 409, 530, 80, { font: BODY, size: 19, color: C.stone });
  rule(slide, 758, 524, 560, 1, C.concrete);
  txt(slide, '02', 758, 552, 90, 45, { font: DISPLAY, size: 35, color: C.ember });
  txt(slide, 'BOUTIQUE', 758, 610, 510, 40, { font: BODY, size: 25, bold: true, color: C.pine });
  rule(slide, 758, 662, 46, 2, C.ember);
  txt(slide, 'Visualisation 3D et expérience d’achat\npersonnalisée en magasin', 758, 684, 530, 56, { font: BODY, size: 18, color: C.stone });
  footer(slide, 4);
  notes(slide, 'Sources : Diaporama Café Premium Marketing-2.pdf, pages 4 et 5. Le document indique une distribution exclusive par circuit direct, en B2B et B2C, avec une stratégie omnicanale. Les canaux web et boutique sont décrits dans le canvas source.');
}

// 5 — Customer journey (synthesized)
{
  const slide = presentation.slides.add();
  slide.background.fill = C.pine;
  header(slide, '04 · Parcours client', 'Parcours client envisagé', true);
  txt(slide, 'Du premier choix au suivi après-vente.', 74, 256, 900, 45, { font: DISPLAY, size: 28, color: C.moss });
  const stages = [
    { x: 76, n: '01', title: 'Découvrir', body: 'Site KŌRI\nou boutique' },
    { x: 406, n: '02', title: 'Personnaliser', body: 'Choix du modèle\net des options\nVisualisation 3D' },
    { x: 736, n: '03', title: 'Acheter', body: 'Circuit direct\nEn ligne ou en boutique' },
    { x: 1066, n: '04', title: 'Après l’achat', body: 'Service après-vente\nFidélité, événements\net nouveautés' },
  ];
  rule(slide, 112, 403, 1156, 1.5, '#5E665F');
  for (const s of stages) {
    rect(slide, s.x + 4, 394, 18, 18, C.ember, 'none', 0);
    txt(slide, s.n, s.x, 444, 68, 37, { font: DISPLAY, size: 28, color: C.ember });
    txt(slide, s.title, s.x, 490, 282, 38, { font: BODY, size: 22, bold: true, color: C.paper });
    txt(slide, s.body, s.x, 540, 270, 112, { font: BODY, size: 19, color: C.moss });
  }
  txt(slide, 'Parcours reconstitué à partir des éléments disponibles ; le PDF source ne détaille pas ces étapes.', 74, 718, 1100, 25, { font: BODY, size: 14, color: C.moss });
  footer(slide, 5, true);
  notes(slide, 'Synthèse de travail, à valider : le PDF source mentionne « Parcours client » dans son sommaire, mais ne détaille pas ce parcours. Les étapes proposées s’appuient sur les canaux web et boutique, la visualisation 3D et la personnalisation, le circuit direct, le service après-vente, les statuts de fidélité, les événements et la découverte des nouveautés décrits dans Diaporama Café Premium Marketing-2.pdf, pages 2, 4 et 5.');
}

// 6 — Business Model Canvas, all elements editable
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  logo(slide, 56, 24, C.ink, 18);
  txt(slide, '05 · MODÈLE ÉCONOMIQUE', 56, 59, 510, 21, { font: BODY, size: 14, bold: true, color: C.ember });
  txt(slide, 'Business Model Canvas', 56, 82, 880, 50, { font: DISPLAY, size: 39, color: C.ink, name: 'canvas-title' });

  const x0 = 56, gap = 8, cw = 259.2, topY = 146, fullH = 400, halfH = 196, splitGap = 8;
  const x1 = x0 + cw + gap, x2 = x1 + cw + gap, x3 = x2 + cw + gap, x4 = x3 + cw + gap;
  canvasCell(slide, x0, topY, cw, fullH, 1, 'Partenaires clés', 'Fournisseurs de café\nFournisseurs de matières premières\nFournisseurs et partenaires');
  canvasCell(slide, x1, topY, cw, halfH, 2, 'Activités clés', 'Conception et fabrication\nPersonnalisation des modèles\nGestion des boutiques et des ventes\nService après-vente', { bodySize: 15 });
  canvasCell(slide, x1, topY + halfH + splitGap, cw, halfH, 3, 'Ressources clés', 'Savoir-faire et équipements\nTechnologie et bornes 3D\nBoutiques et équipe\nMarque, partenaires, finances', { bodySize: 15 });
  canvasCell(slide, x2, topY, cw, fullH, 4, 'Proposition de valeur', 'Machine premium au design en béton\n\nPersonnalisation selon les goûts\n\nVisualisation 3D en ligne et en boutique\n\nExpérience café premium pour particuliers et hôtels', { bodySize: 16 });
  canvasCell(slide, x3, topY, cw, halfH, 5, 'Relations clients', 'Achat personnalisé\nStatuts selon les montants d’achat\nInvitations et nouveautés en exclusivité', { bodySize: 15 });
  canvasCell(slide, x3, topY + halfH + splitGap, cw, halfH, 6, 'Canaux', 'Site web avec visualisation 3D\nBoutiques avec visualisation 3D\nPersonnalisation en boutique', { bodySize: 15 });
  canvasCell(slide, x4, topY, cw, fullH, 7, 'Segments clients', 'B2C · Particuliers\n\nCadres et jeunes cadres, 35–50 ans\nAmateurs de café premium\nSensibles au design et à la durabilité\n\nB2B · Hôtels haut de gamme', { bodySize: 15.5 });
  const bottomY = topY + fullH + 10;
  const bw = (1328 - 8) / 2;
  canvasCell(slide, x0, bottomY, bw, 184, 8, 'Structure des coûts', 'Fabrication, matières et composants\nSalaires, boutiques et technologie\nMarketing et communication\nTransport, livraison, service après-vente\nFidélisation client', { bodySize: 17 });
  canvasCell(slide, x0 + bw + 8, bottomY, bw, 184, 9, 'Sources de revenus', 'Vente de cafetières\nVente de café en grains premium', { bodySize: 18 });
  txt(slide, 'Synthèse du canvas fourni · version reformulée pour la lisibilité', 56, 783, 900, 17, { font: BODY, size: 12.5, color: C.graphite });
  txt(slide, '06 / 06', 1288, 783, 80, 17, { font: BODY, size: 12.5, color: C.graphite, align: 'right' });
  notes(slide, 'Source principale : Diaporama Café Premium Marketing-2.pdf, page 5. Les libellés ont été reformulés pour corriger les erreurs de transcription et condenser le contenu sans ajouter de chiffres. La mention ambiguë « technologie pélygale » du PDF source n’est pas reprise. Les « statuts selon les montants » et les invitations aux événements/nouveautés sont conservés comme éléments de relation client. Les éléments visuels suivent KORI_Charte_graphique.pdf, pages 4–14.');
}

// Write previews and candidate, then run the required finalizer.
const previewDir = path.join(TMP_DIR, 'previews');
const stagingDir = path.join(workspaceDir, '.codex-finalizer');
await fs.mkdir(previewDir, { recursive: true });
await fs.mkdir(stagingDir, { recursive: true });
await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
for (let i = 0; i < presentation.slides.items.length; i++) {
  const slide = presentation.slides.items[i];
  const image = await slide.export({ format: 'png', scale: 1 });
  await fs.writeFile(path.join(previewDir, `slide-${String(i + 1).padStart(2, '0')}.png`), new Uint8Array(await image.arrayBuffer()));
}
const candidatePath = path.join(stagingDir, 'candidate.pptx');
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, 'container_tools/artifact_tool_utils.mjs')).href);
const charterHash = createHash('sha256').update(await fs.readFile(CHARTER)).digest('hex');
const result = await finalizePresentation({
  explicitTotalSlideCount: 6,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
  workspaceDir,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: '/Users/atom/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3',
  integrityValidatorPath: path.join(SKILL_DIR, 'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(SKILL_DIR, 'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '13716000,7715250', '--validate-bullet-geometry', '--validate-heading-fit'],
  fontPolicy: { basis: 'reference', families: [DISPLAY, BODY], referencePath: CHARTER, referenceSha256: charterHash },
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, 'KORI_Objet_01_Diaporama_Canvas_final.validation.json'),
});
console.log(JSON.stringify({ finalPath: FINAL_PPTX, slideCount: presentation.slides.items.length, result }, null, 2));
