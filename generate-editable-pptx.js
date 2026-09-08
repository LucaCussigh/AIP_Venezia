const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const projectRoot = __dirname;
const markdownPath = path.join(projectRoot, 'presentation.md');
const markdown = fs.readFileSync(markdownPath, 'utf8')
  .replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, '')
  .replace(/<!--[\s\S]*?-->/g, '')
  .trim();
const slides = markdown.split(/\r?\n---\r?\n/).map((slide) => slide.trim()).filter(Boolean);

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Luca Cussigh';
pptx.subject = 'AIP Pisa presentation';
pptx.title = 'Dai volti ai gruppi';
pptx.company = 'Universita degli Studi di Padova';
pptx.lang = 'it-IT';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'it-IT'
};

const colors = {
  ink: '13231F',
  paper: 'F6F4EE',
  accent: 'B7492E',
  accentDark: '82351F',
  line: 'D7D1C5'
};

function cleanMarkdown(value) {
  return value
    .replace(/<[^>]+>/g, ' ')
    .replace(/[*_`]/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

function imageSources(value) {
  return [...value.matchAll(/<img[^>]+src=["']([^"']+)["'][^>]*>/gi)]
    .map((match) => path.join(projectRoot, match[1].replaceAll('/', path.sep)))
    .filter((filePath) => fs.existsSync(filePath));
}

function addImages(slide, source) {
  const images = imageSources(source);
  if (!images.length) return;
  const usableWidth = 11.7;
  const columns = images.length === 1 ? 1 : Math.min(images.length, 4);
  const rows = Math.ceil(images.length / columns);
  const cellWidth = usableWidth / columns;
  const cellHeight = Math.min(2.25, 5.6 / rows);
  const startY = 1.55;

  images.forEach((filePath, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    slide.addImage({
      path: filePath,
      x: 0.55 + column * cellWidth,
      y: startY + row * cellHeight,
      w: cellWidth - 0.18,
      h: cellHeight - 0.18,
      sizingContain: true
    });
  });
}

function addBody(slide, source) {
  const lines = source.split(/\r?\n/);
  const textLines = [];
  let y = 1.48;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('|') || trimmed.startsWith('<img')) continue;
    if (/^#{1,6}\s/.test(trimmed)) continue;
    if (/^https?:\/\//.test(trimmed)) {
      slide.addText(trimmed, { x: 0.7, y, w: 11.9, h: 0.35, fontFace: 'Aptos', fontSize: 13, color: '0563C1', breakLine: false });
      y += 0.45;
      continue;
    }
    const isBullet = /^[-*+]\s+/.test(trimmed);
    const text = cleanMarkdown(trimmed.replace(/^[-*+]\s+/, '').replace(/^\d+\.\s+/, ''));
    if (!text || text === '---') continue;
    textLines.push({ text: `${isBullet ? '• ' : ''}${text}`, options: { bullet: false } });
  }

  if (textLines.length) {
    slide.addText(textLines, {
      x: 0.72,
      y,
      w: 11.85,
      h: 5.35,
      fontFace: 'Aptos',
      fontSize: textLines.length > 8 ? 15 : 19,
      color: colors.ink,
      breakLine: true,
      paraSpaceAfterPt: 8,
      fit: 'shrink'
    });
  }
}

slides.forEach((source) => {
  const slide = pptx.addSlide();
  slide.background = { color: colors.paper };
  const heading = source.match(/^#\s+(.+)$/m);
  const title = heading ? cleanMarkdown(heading[1]) : 'AIP Pisa';
  slide.addText(title, {
    x: 0.62,
    y: 0.34,
    w: 12.05,
    h: 0.72,
    fontFace: 'Aptos Display',
    fontSize: title.length > 55 ? 23 : 29,
    bold: false,
    color: colors.ink,
    margin: 0,
    fit: 'shrink'
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 0.65,
    y: 1.16,
    w: 12.0,
    h: 0,
    line: { color: colors.accent, width: 1.2 }
  });
  addBody(slide, source);
  addImages(slide, source);
});

pptx.writeFile({ fileName: path.join(projectRoot, 'presentation-editable.pptx') });