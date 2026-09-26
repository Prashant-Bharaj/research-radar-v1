const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');

// ---------------------------------------------------------------------------
//  .env laden (Key bleibt ausschliesslich serverseitig)
// ---------------------------------------------------------------------------
for (const p of [path.join(__dirname, '..', '.env'), path.join(__dirname, '.env')]) {
  if (fs.existsSync(p)) {
    for (const line of fs.readFileSync(p, 'utf8').split(/\r?\n/)) {
      const m = line.match(/^\s*([\w.]+)\s*=\s*(.*)\s*$/);
      if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
    }
  }
}

const MGA_BASE = process.env.MGA_BASE_URL || 'https://chat.int.bayer.com/api/v2';
const MGA_MODEL = process.env.MGA_MODEL || 'claude-opus-5';

const app = express();
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// ---------------------------------------------------------------------------
//  Produktionsbetrieb (ein Dienst): Passwortschutz, /api-Prefix, Frontend
//  Lokal unveraendert: ohne APP_PASSWORD kein Schutz, ohne dist kein Static.
// ---------------------------------------------------------------------------
const APP_PASSWORD = process.env.APP_PASSWORD || '';
if (APP_PASSWORD) {
  app.use((req, res, next) => {
    const header = req.headers.authorization || '';
    const [scheme, value] = header.split(' ');
    if (scheme === 'Basic' && value) {
      const pass = Buffer.from(value, 'base64').toString('utf8').split(':').slice(1).join(':');
      if (pass === APP_PASSWORD) return next();
    }
    res.set('WWW-Authenticate', 'Basic realm="Research Radar"').status(401).send('Zugang geschuetzt');
  });
}

// Bildet den vite-/nginx-Proxy nach: /api/search -> /search
app.use((req, res, next) => {
  if (req.url === '/api' || req.url.startsWith('/api/')) req.url = req.url.slice(4) || '/';
  next();
});

const DIST = path.join(__dirname, '..', 'frontend', 'dist');
const HAS_DIST = fs.existsSync(DIST);
if (HAS_DIST) app.use(express.static(DIST));

// ---------------------------------------------------------------------------
//  KI-Aufruf (myGenAssist, OpenAI-kompatibel)
// ---------------------------------------------------------------------------
async function llm(messages, { maxTokens = 1500, temperature = 0.2, json = false } = {}) {
  if (!process.env.MGA_API_KEY) throw Object.assign(new Error('KI-Schlüssel fehlt'), { statusCode: 503 });
  const res = await fetch(`${MGA_BASE}/chat/completions`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${process.env.MGA_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MGA_MODEL, messages, max_tokens: maxTokens, temperature }),
  });
  if (!res.ok) throw Object.assign(new Error('KI-Dienst nicht erreichbar'), { statusCode: 502, code: 'LLM_' + res.status });
  const data = await res.json();
  if (data.model !== MGA_MODEL) console.warn(`MGA: Fallback-Modell ${data.model} statt ${MGA_MODEL}`);
  let text = data.choices?.[0]?.message?.content ?? '';
  if (json) {
    text = text.replace(/```json|```/g, '').trim();
    const start = text.indexOf('['), startObj = text.indexOf('{');
    const s = start === -1 ? startObj : startObj === -1 ? start : Math.min(start, startObj);
    return JSON.parse(text.slice(s));
  }
  return text;
}

// ---------------------------------------------------------------------------
//  Quellen
// ---------------------------------------------------------------------------
const SOURCES = {
  pubmed: { label: 'PubMed', live: true },
};

const fmt = (d) => d.toISOString().slice(0, 10).replace(/-/g, '/');

async function searchPubMed(query, months, max = 12) {
  const to = new Date();
  const from = new Date();
  from.setMonth(from.getMonth() - months);
  const base = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils';
  const term = encodeURIComponent(`(${query}) AND (ophthalmology OR eye OR retina OR cornea OR glaucoma)`);
  const es = await fetch(
    `${base}/esearch.fcgi?db=pubmed&term=${term}&datetype=pdat&mindate=${fmt(from)}&maxdate=${fmt(to)}&retmode=json&retmax=${max}&sort=relevance`
  ).then((r) => r.json());
  const ids = es.esearchresult?.idlist || [];
  if (!ids.length) return [];
  const sum = await fetch(`${base}/esummary.fcgi?db=pubmed&id=${ids.join(',')}&retmode=json`).then((r) => r.json());
  return ids.map((id) => {
    const r = sum.result?.[id] || {};
    const pmc = (r.articleids || []).find((a) => a.idtype === 'pmc')?.value;
    const doi = (r.articleids || []).find((a) => a.idtype === 'doi')?.value;
    return {
      id: `pm-${id}`,
      source: 'pubmed',
      title: r.title || '(ohne Titel)',
      authors: (r.authors || []).slice(0, 3).map((a) => a.name).join(', ') + ((r.authors || []).length > 3 ? ' et al.' : ''),
      journal: r.fulljournalname || r.source || '',
      date: r.pubdate || '',
      url: `https://pubmed.ncbi.nlm.nih.gov/${id}/`,
      doi,
      accessible: Boolean(pmc),
      fulltextUrl: pmc ? `https://pmc.ncbi.nlm.nih.gov/articles/${pmc}/` : null,
    };
  });
}

// Demo-Einträge für Quellen ohne Live-Anbindung (klar als Demo gekennzeichnet)
function demoEntries(source, query) {
  const label = SOURCES[source].label;
  const mk = (n, accessible) => ({
    id: `${source}-demo-${n}`,
    source,
    demo: true,
    title: `[Demo] ${label}: Beitrag ${n} zu „${query}“`,
    authors: 'Muster A, Beispiel B',
    journal: label,
    date: '2026',
    url: 'https://example.org/demo',
    accessible,
    fulltextUrl: accessible ? 'https://example.org/demo' : null,
  });
  return [mk(1, true), mk(2, false)];
}

// ---------------------------------------------------------------------------
//  In-Memory Datenpool
// ---------------------------------------------------------------------------
const pool = { papers: [], uploads: [] };

const clean = (v, max = 200) => (typeof v === 'string' ? v.trim().slice(0, max) : '');

app.get('/health', (req, res) => res.json({ status: 'ok', llm: Boolean(process.env.MGA_API_KEY) }));
app.get('/sources', (req, res) => res.json({ sources: SOURCES }));

// Suche
app.post('/search', async (req, res, next) => {
  try {
    const query = clean(req.body?.query, 120);
    const months = Math.min(60, Math.max(1, Number(req.body?.months) || 12));
    const sources = (Array.isArray(req.body?.sources) ? req.body.sources : []).filter((s) => SOURCES[s]);
    if (!query) return res.status(400).json({ error: 'BAD_REQUEST', message: 'Stichwort fehlt', statusCode: 400 });
    if (!sources.length) return res.status(400).json({ error: 'BAD_REQUEST', message: 'Keine Quelle gewählt', statusCode: 400 });

    // Stichwort (auch deutsch) in einen englischen PubMed-Suchbegriff übersetzen
    let term = query;
    if (process.env.MGA_API_KEY) {
      try {
        const t = await llm(
          [
            { role: 'system', content: 'Du formulierst PubMed-Suchbegriffe. Antworte nur mit JSON.' },
            { role: 'user', content: `Übersetze dieses ophthalmologische Thema in einen präzisen englischen PubMed-Suchbegriff (MeSH-Begriffe und gängige Synonyme mit OR verknüpfen, maximal 4 Synonyme, keine Erklärung). Thema: "${query}"
Antwort: {"term":"..."}` },
          ],
          { json: true, maxTokens: 200 }
        );
        if (t?.term) term = clean(t.term, 300);
      } catch (e) {
        console.warn('Übersetzung übersprungen:', e.code || e.message);
      }
    }

    let papers = [];
    for (const s of sources) {
      if (s === 'pubmed') papers.push(...(await searchPubMed(term, months)));
      else papers.push(...demoEntries(s, query));
    }

    // KI: Relevanz + Ein-Satz-Begründung
    if (papers.length && process.env.MGA_API_KEY) {
      try {
        const list = papers.map((p, i) => `${i}. ${p.title} (${p.journal}, ${p.date})`).join('\n');
        const ranked = await llm(
          [
            { role: 'system', content: 'Du bist ein ophthalmologischer Research-Assistent. Antworte ausschließlich mit JSON.' },
            {
              role: 'user',
              content: `Thema: "${query}". Bewerte jede Publikation nach Relevanz für ein Update-Referat (Score 1-10) und gib eine deutsche Ein-Satz-Begründung.\n\n${list}\n\nAntworte als JSON-Array: [{"i":0,"score":8,"why":"..."}]`,
            },
          ],
          { json: true, maxTokens: 2500 }
        );
        for (const r of ranked) if (papers[r.i]) Object.assign(papers[r.i], { score: r.score, why: r.why });
        papers.sort((a, b) => (b.score || 0) - (a.score || 0));
      } catch (e) {
        console.warn('Ranking übersprungen:', e.code || e.message);
      }
    }

    pool.papers = papers.map((p) => ({ ...p, license: p.accessible ? 'free' : null, upload: null }));
    res.json({ papers: pool.papers, query, term, months });
  } catch (err) {
    next(err);
  }
});

app.get('/pool', (req, res) => res.json(pool));

// Manueller Upload zu einem Paper (mit Lizenz-Erklärung)
app.post('/papers/:id/upload', (req, res) => {
  const p = pool.papers.find((x) => x.id === req.params.id);
  if (!p) return res.status(404).json({ error: 'NOT_FOUND', message: 'Paper nicht gefunden', statusCode: 404 });
  const license = req.body?.license;
  if (!['free', 'paid'].includes(license))
    return res.status(400).json({ error: 'BAD_REQUEST', message: 'Lizenz muss "free" oder "paid" sein', statusCode: 400 });
  const filename = clean(req.body?.filename, 120) || 'upload';
  const size = Number(req.body?.size) || 0;
  p.license = license;
  p.upload = { filename, size, at: new Date().toISOString() };
  // Kostenpflichtig erworbene Inhalte: nur Nachweis speichern, Inhalt nicht weiterverwenden
  p.content = license === 'free' ? clean(req.body?.text, 200000) : null;
  res.json(p);
});

// Freier Zusatz-Upload (Datenpool anreichern)
app.post('/uploads', (req, res) => {
  const filename = clean(req.body?.filename, 120);
  if (!filename) return res.status(400).json({ error: 'BAD_REQUEST', message: 'Dateiname fehlt', statusCode: 400 });
  const u = { id: `up-${Date.now()}`, filename, size: Number(req.body?.size) || 0, text: clean(req.body?.text, 200000), at: new Date().toISOString() };
  pool.uploads.push(u);
  res.status(201).json({ id: u.id, filename: u.filename, size: u.size, at: u.at });
});

app.delete('/uploads/:id', (req, res) => {
  pool.uploads = pool.uploads.filter((u) => u.id !== req.params.id);
  res.json({ ok: true });
});

// ---------------------------------------------------------------------------
//  Outline: Vortragsgliederung aus dem Datenpool, iterativ per Chat
// ---------------------------------------------------------------------------
function poolContext() {
  const usable = pool.papers.filter((p) => p.license === 'free');
  const lines = usable.map(
    (p, i) => `[${i + 1}] ${p.title} – ${p.authors} (${p.journal}, ${p.date})${p.why ? ' – ' + p.why : ''}${p.content ? ' Auszug: ' + p.content.slice(0, 1500) : ''}`
  );
  const ups = pool.uploads.map((u) => `[Upload] ${u.filename}${u.text ? ': ' + u.text.slice(0, 1500) : ''}`);
  return { usable, text: [...lines, ...ups].join(String.fromCharCode(10)) };
}

app.post('/outline', async (req, res, next) => {
  try {
    const topic = clean(req.body?.topic, 120);
    const history = (Array.isArray(req.body?.messages) ? req.body.messages : [])
      .filter((m) => ['user', 'assistant'].includes(m.role) && typeof m.content === 'string')
      .slice(-12)
      .map((m) => ({ role: m.role, content: m.content.slice(0, 6000) }));
    const { usable, text } = poolContext();
    if (!usable.length && !pool.uploads.length)
      return res.status(400).json({ error: 'BAD_REQUEST', message: 'Datenpool ist leer', statusCode: 400 });

    const system = `Du bist ein erfahrener ophthalmologischer Referent und erstellst die Gliederung für einen Update-Vortrag zum Thema "${topic}".
Nutze AUSSCHLIESSLICH die folgenden Quellen aus dem Datenpool (kostenpflichtig erworbene Volltexte sind bereits ausgeschlossen):
${text}

Regeln:
- Antworte auf Deutsch, in Markdown.
- Gliederung als nummerierte Folien: "## Folie N: Titel", darunter 2-4 Stichpunkte, Quellen als [Nr] referenzieren.
- 8-12 Folien, beginnend mit Titel/Agenda, endend mit Take-Home-Messages.
- Bei Feedback des Nutzers: die KOMPLETTE überarbeitete Gliederung erneut ausgeben, nicht nur die Änderung.`;

    const messages = [{ role: 'system', content: system }];
    if (!history.length) messages.push({ role: 'user', content: 'Erstelle den ersten Entwurf der Gliederung.' });
    else messages.push(...history);

    const content = await llm(messages, { maxTokens: 3000, temperature: 0.3 });
    pool.outline = content;
    res.json({ content });
  } catch (err) {
    next(err);
  }
});

// ---------------------------------------------------------------------------
//  PowerPoint: Outline -> strukturierte Folien (KI) -> PPTX (modernes Template)
// ---------------------------------------------------------------------------
const THEME = { bg: 'FFFFFF', ink: '1C2430', muted: '5B6675', accent: '0A6CFF', accent2: '0B3D91', light: 'EEF4FF', line: 'E6E9EE' };

function addFooter(slide, pptx, n, total) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.32, w: 10, h: 0.02, fill: { color: THEME.line }, line: { color: THEME.line } });
  slide.addText('Research Radar · Update Ophthalmologie', { x: 0.4, y: 5.35, w: 6, h: 0.25, fontSize: 9, color: THEME.muted, fontFace: 'Calibri' });
  slide.addText(`${n} / ${total}`, { x: 8.6, y: 5.35, w: 1, h: 0.25, fontSize: 9, color: THEME.muted, align: 'right', fontFace: 'Calibri' });
}

async function buildPptx(topic, deck) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'Research Radar';
  pptx.title = deck.title || topic;
  const slides = deck.slides || [];
  const total = slides.length + 2;

  // Titelfolie
  let s = pptx.addSlide();
  s.background = { color: THEME.accent2 };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.25, h: 5.63, fill: { color: THEME.accent }, line: { color: THEME.accent } });
  s.addText(deck.title || topic, { x: 0.7, y: 1.4, w: 8.6, h: 1.6, fontSize: 34, bold: true, color: 'FFFFFF', fontFace: 'Calibri', valign: 'bottom' });
  s.addText(deck.subtitle || 'Update Ophthalmologie', { x: 0.7, y: 3.05, w: 8.6, h: 0.6, fontSize: 18, color: 'CFE0FF', fontFace: 'Calibri' });
  s.addText(new Date().toLocaleDateString('de-DE', { year: 'numeric', month: 'long' }), { x: 0.7, y: 4.5, w: 8, h: 0.4, fontSize: 12, color: 'CFE0FF', fontFace: 'Calibri' });

  // Inhaltsfolien
  slides.forEach((sl, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: THEME.bg };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: THEME.light }, line: { color: THEME.light } });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: 0.9, fill: { color: THEME.accent }, line: { color: THEME.accent } });
    slide.addText(sl.title || `Folie ${i + 1}`, { x: 0.4, y: 0.12, w: 9.2, h: 0.66, fontSize: 22, bold: true, color: THEME.ink, fontFace: 'Calibri', valign: 'middle' });
    const bullets = (sl.bullets || []).slice(0, 6).map((b) => ({ text: String(b), options: { bullet: { indent: 18 }, breakLine: true } }));
    if (bullets.length) slide.addText(bullets, { x: 0.5, y: 1.15, w: 9, h: 3.4, fontSize: 16, color: THEME.ink, fontFace: 'Calibri', valign: 'top', paraSpaceAfter: 8 });
    if (sl.sources) slide.addText(`Quellen: ${sl.sources}`, { x: 0.5, y: 4.7, w: 9, h: 0.5, fontSize: 10, italic: true, color: THEME.muted, fontFace: 'Calibri' });
    addFooter(slide, pptx, i + 2, total);
  });

  // Quellenfolie
  const refs = pool.papers.filter((p) => p.license === 'free');
  s = pptx.addSlide();
  s.background = { color: THEME.bg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: THEME.light }, line: { color: THEME.light } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: 0.9, fill: { color: THEME.accent }, line: { color: THEME.accent } });
  s.addText('Quellen', { x: 0.4, y: 0.12, w: 9.2, h: 0.66, fontSize: 22, bold: true, color: THEME.ink, fontFace: 'Calibri', valign: 'middle' });
  const refText = refs.map((p, i) => ({ text: `[${i + 1}] ${p.authors}. ${p.title} ${p.journal} ${p.date}.`, options: { breakLine: true } }));
  if (refText.length) s.addText(refText, { x: 0.5, y: 1.1, w: 9, h: 4.1, fontSize: 9, color: THEME.ink, fontFace: 'Calibri', valign: 'top', paraSpaceAfter: 4 });
  addFooter(s, pptx, total, total);

  return pptx.write({ outputType: 'nodebuffer' });
}

app.post('/pptx', async (req, res, next) => {
  try {
    const topic = clean(req.body?.topic, 120) || 'Update';
    const outline = typeof req.body?.outline === 'string' && req.body.outline.trim() ? req.body.outline.slice(0, 20000) : pool.outline;
    if (!outline) return res.status(400).json({ error: 'BAD_REQUEST', message: 'Keine Outline vorhanden', statusCode: 400 });

    const deck = await llm(
      [
        { role: 'system', content: 'Du wandelst eine Vortragsgliederung in strukturierte Folien um. Antworte ausschließlich mit JSON.' },
        {
          role: 'user',
          content: `Thema: "${topic}". Wandle diese Gliederung 1:1 in Folien um. Pro Folie: "title" (kurz), "bullets" (3-5 prägnante Stichpunkte, je maximal 110 Zeichen, deutsch), "sources" (Quellenverweise wie "[1], [3]" oder ""). Titelfolie und Quellenfolie NICHT erzeugen, die kommen automatisch.\n\nGliederung:\n${outline}\n\nJSON: {"title":"...","subtitle":"...","slides":[{"title":"...","bullets":["..."],"sources":"[1]"}]}`,
        },
      ],
      { json: true, maxTokens: 4000 }
    );

    const buf = await buildPptx(topic, deck);
    const fname = `Update_${topic.replace(/[^\w\u00C0-\u017F-]+/g, '_').slice(0, 60)}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${fname}"`);
    res.setHeader('X-Slide-Count', String((deck.slides || []).length + 2));
    res.send(buf);
  } catch (err) {
    next(err);
  }
});

// ---------------------------------------------------------------------------
app.use((err, req, res, next) => {
  console.error('Server error:', err.code || 'INTERNAL_ERROR', err.message);
  const code = err.statusCode || 500;
  res.status(code).json({
    error: err.code || 'INTERNAL_ERROR',
    message: code < 500 ? err.message : 'Ein unerwarteter Fehler ist aufgetreten',
    statusCode: code,
  });
});

if (HAS_DIST) app.get('*', (req, res) => res.sendFile(path.join(DIST, 'index.html')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Backend läuft auf Port ${PORT} · KI: ${process.env.MGA_API_KEY ? 'aktiv' : 'fehlt'}`));
