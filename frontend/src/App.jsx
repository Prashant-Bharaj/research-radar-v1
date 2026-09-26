import { useEffect, useRef, useState } from 'react';

const STEPS = ['Recherche', 'Datenpool', 'Outline', 'PowerPoint'];

const fmtBytes = (n) => (n > 1e6 ? (n / 1e6).toFixed(1) + ' MB' : Math.round(n / 1e3) + ' kB');

function readFile(file) {
  return new Promise((resolve) => {
    const r = new FileReader();
    r.onload = () => resolve(typeof r.result === 'string' ? r.result : '');
    r.onerror = () => resolve('');
    if (file.type.startsWith('text') || /\.(txt|md|csv)$/i.test(file.name)) r.readAsText(file);
    else resolve('');
  });
}

export default function App() {
  const [sources, setSources] = useState({});
  const [selected, setSelected] = useState(['pubmed']);
  const [query, setQuery] = useState('');
  const [months, setMonths] = useState(12);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState({ pct: 0, label: '' });
  const [papers, setPapers] = useState([]);
  const [uploads, setUploads] = useState([]);
  const [error, setError] = useState('');
  const [term, setTerm] = useState('');
  const [searched, setSearched] = useState(false);
  const [uploadFor, setUploadFor] = useState(null);
  const [step, setStep] = useState(0);
  const [chat, setChat] = useState([]);
  const [outline, setOutline] = useState('');
  const [outlineBusy, setOutlineBusy] = useState(false);
  const [feedback, setFeedback] = useState('');
  const [pptx, setPptx] = useState({ busy: false, url: '', name: '', slides: 0 });
  const timer = useRef();

  useEffect(() => {
    fetch('/api/sources').then((r) => r.json()).then((d) => setSources(d.sources || {})).catch(() => {});
  }, []);

  function toggle(key) {
    setSelected((s) => (s.includes(key) ? s.filter((k) => k !== key) : [...s, key]));
  }

  function startProgress() {
    const phases = ['Portale werden abgefragt…', 'Treffer werden gesammelt…', 'Relevanz wird bewertet…', 'Zugänglichkeit wird geprüft…'];
    let pct = 0;
    setProgress({ pct: 3, label: phases[0] });
    timer.current = setInterval(() => {
      pct = Math.min(92, pct + (pct < 40 ? 4 : pct < 75 ? 2 : 0.6));
      setProgress({ pct, label: phases[Math.min(3, Math.floor(pct / 25))] });
    }, 250);
  }

  async function search(e) {
    e.preventDefault();
    if (!query.trim() || loading) return;
    setError('');
    setLoading(true);
    setPapers([]);
    setSearched(false);
    startProgress();
    try {
      const res = await fetch('/api/search', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ query: query.trim(), months, sources: selected }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.message || 'Suche fehlgeschlagen');
      setPapers(data.papers || []);
      setTerm(data.term || '');
      setSearched(true);
      setStep(1);
    } catch (err) {
      setError(err.message);
    } finally {
      clearInterval(timer.current);
      setProgress({ pct: 100, label: 'Fertig' });
      setTimeout(() => setLoading(false), 350);
    }
  }

  async function submitUpload(paper, file, license) {
    const text = await readFile(file);
    const res = await fetch(`/api/papers/${paper.id}/upload`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: file.name, size: file.size, license, text }),
    });
    const updated = await res.json();
    setPapers((ps) => ps.map((p) => (p.id === updated.id ? updated : p)));
    setUploadFor(null);
  }

  async function addUpload(file) {
    const text = await readFile(file);
    const res = await fetch('/api/uploads', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: file.name, size: file.size, text }),
    });
    const created = await res.json();
    setUploads((u) => [...u, created]);
  }

  async function removeUpload(id) {
    await fetch(`/api/uploads/${id}`, { method: 'DELETE' });
    setUploads((u) => u.filter((x) => x.id !== id));
  }

  async function requestOutline(messages) {
    setOutlineBusy(true);
    setError('');
    try {
      const res = await fetch('/api/outline', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ topic: query.trim(), messages }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.message || 'Outline fehlgeschlagen');
      setOutline(data.content);
      setChat([...messages, { role: 'assistant', content: data.content }]);
    } catch (err) {
      setError(err.message);
    } finally {
      setOutlineBusy(false);
    }
  }

  function goOutline() {
    setStep(2);
    setChat([]);
    setOutline('');
    requestOutline([]);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function sendFeedback(e) {
    e.preventDefault();
    if (!feedback.trim() || outlineBusy) return;
    const msgs = [...chat, { role: 'user', content: feedback.trim() }];
    setFeedback('');
    requestOutline(msgs);
  }

  async function createPptx() {
    setStep(3);
    setPptx({ busy: true, url: '', name: '', slides: 0 });
    setError('');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    try {
      const res = await fetch('/api/pptx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ topic: query.trim(), outline }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.message || 'PowerPoint konnte nicht erstellt werden');
      }
      const blob = await res.blob();
      const cd = res.headers.get('Content-Disposition') || '';
      const name = (cd.match(/filename="?([^"]+)"?/) || [])[1] || 'Update.pptx';
      const url = URL.createObjectURL(blob);
      setPptx({ busy: false, url, name, slides: Number(res.headers.get('X-Slide-Count')) || 0 });
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      a.click();
    } catch (err) {
      setError(err.message);
      setPptx({ busy: false, url: '', name: '', slides: 0 });
    }
  }

  const usable = papers.filter((p) => p.license === 'free').length;
  const excluded = papers.filter((p) => p.license === 'paid').length;
  const open = papers.filter((p) => !p.license).length;

  return (
    <div className="page">
      <header className="topbar">
        <div className="brand">
          <span className="dot" />
          Research Radar · Ophthalmologie
        </div>
        <nav className="stepper">
          {STEPS.map((s, i) => (
            <span key={s} className={`step ${i === step ? 'active' : i < step ? 'done' : ''}`}>
              <b>{i + 1}</b> {s}
            </span>
          ))}
        </nav>
      </header>

      <main className="container wide">
        {/* ---------------- Suche ---------------- */}
        {step < 2 && (
        <section className="card">
          <h1>Was ist neu zu deinem Thema?</h1>
          <p className="lead">Stichwort eingeben, Zeitraum wählen – PubMed wird automatisch durchsucht und per KI nach Relevanz sortiert.</p>
          <form onSubmit={search} className="searchform">
            <div className="row">
              <input
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="z. B. Faricimab bei diabetischem Makulaödem"
                maxLength={120}
                autoFocus
              />
              <button type="submit" disabled={loading || !query.trim()}>
                {loading ? 'Recherchiert…' : 'Recherche starten'}
              </button>
            </div>

            <div className="slider">
              <label>
                Zeitraum: <b>letzte {months} Monate</b>
              </label>
              <input type="range" min="1" max="36" value={months} onChange={(e) => setMonths(Number(e.target.value))} />
              <div className="ticks"><span>1 Monat</span><span>12</span><span>24</span><span>36 Monate</span></div>
            </div>

            <div className="chips">
              {Object.entries(sources).map(([key, s]) => (
                <label key={key} className={`chip ${selected.includes(key) ? 'on' : ''}`}>
                  <input type="checkbox" checked={selected.includes(key)} onChange={() => toggle(key)} />
                  {s.label}
                  {!s.live && <small>Demo</small>}
                </label>
              ))}
            </div>
          </form>

          {loading && (
            <div className="progress">
              <div className="bar"><div style={{ width: progress.pct + '%' }} /></div>
              <span>{progress.label}</span>
            </div>
          )}
          {error && <p className="error">{error}</p>}
          {!loading && term && (
            <p className="muted termline">PubMed-Suchbegriff: <code>{term}</code></p>
          )}
          {!loading && searched && papers.length === 0 && !error && (
            <p className="error">Keine Treffer im gewählten Zeitraum. Zeitraum erweitern oder Stichwort anpassen.</p>
          )}
        </section>
        )}

        {/* ---------------- Ergebnisse ---------------- */}
        {step < 2 && papers.length > 0 && (
          <section className="card">
            <div className="cardhead">
              <h2>{papers.length} relevante Publikationen</h2>
              <div className="stats">
                <span className="pill ok">✓ {usable} nutzbar</span>
                <span className="pill warn">? {open} offen</span>
                <span className="pill bad">✕ {excluded} ausgeschlossen</span>
              </div>
            </div>
            <ul className="papers">
              {papers.map((p) => (
                <li key={p.id} className={`paper ${p.license === 'paid' ? 'excluded' : ''}`}>
                  <div className={`status ${p.license === 'free' ? 'green' : p.license === 'paid' ? 'gray' : 'red'}`}>
                    {p.license === 'free' ? '✓' : p.license === 'paid' ? '✕' : '!'}
                  </div>
                  <div className="body">
                    <a href={p.url} target="_blank" rel="noreferrer" className="title">{p.title}</a>
                    <div className="meta">
                      <span className="src">{sources[p.source]?.label}</span>
                      {p.authors && <span>{p.authors}</span>}
                      {p.journal && <span>{p.journal}</span>}
                      {p.date && <span>{p.date}</span>}
                      {p.score != null && <span className="score">Relevanz {p.score}/10</span>}
                    </div>
                    {p.why && <p className="why">{p.why}</p>}
                    <div className="actions">
                      {p.license === 'free' && p.fulltextUrl && (
                        <a href={p.fulltextUrl} target="_blank" rel="noreferrer" className="linkbtn">Volltext frei verfügbar</a>
                      )}
                      {p.license === 'free' && p.upload && <span className="tag">Hochgeladen: {p.upload.filename}</span>}
                      {p.license === 'paid' && <span className="tag">Kostenpflichtig erworben – wird nicht weiterverwendet</span>}
                      {!p.license && (
                        <>
                          <span className="tag red">Volltext nicht frei zugänglich</span>
                          <button type="button" className="ghost" onClick={() => setUploadFor(p)}>Manuell hochladen</button>
                        </>
                      )}
                    </div>
                  </div>
                </li>
              ))}
            </ul>
          </section>
        )}

        {/* ---------------- Zusatz-Uploads ---------------- */}
        {step < 2 && papers.length > 0 && (
          <section className="card">
            <div className="cardhead">
              <h2>Weitere Unterlagen zum Datenpool hinzufügen</h2>
              <label className="ghost filebtn">
                + Datei hochladen
                <input type="file" hidden onChange={(e) => e.target.files[0] && addUpload(e.target.files[0])} />
              </label>
            </div>
            {uploads.length === 0 ? (
              <p className="empty">Noch keine zusätzlichen Unterlagen (eigene Notizen, Vortragsfolien, Abstracts …).</p>
            ) : (
              <ul className="list">
                {uploads.map((u) => (
                  <li key={u.id} className="uploadrow">
                    <span>📄 {u.filename} <small>{fmtBytes(u.size)}</small></span>
                    <button type="button" className="ghost small" onClick={() => removeUpload(u.id)}>Entfernen</button>
                  </li>
                ))}
              </ul>
            )}
            <div className="next">
              <span className="muted">
                Datenpool: {usable} Publikationen + {uploads.length} Unterlagen
              </span>
              <button type="button" onClick={goOutline}>Outline erstellen →</button>
            </div>
          </section>
        )}
        {/* ---------------- Outline ---------------- */}
        {step === 2 && (
          <section className="card outline">
            <div className="cardhead">
              <div>
                <h1>Vortrags-Outline: {query}</h1>
                <p className="lead">Basis: {usable} Publikationen + {uploads.length} Unterlagen. Gib Feedback, bis die Gliederung passt.</p>
              </div>
              <button type="button" className="ghost" onClick={() => setStep(1)}>← Zurück zum Datenpool</button>
            </div>

            <div className="split">
              <div className="draft">
                {outlineBusy && (
                  <div className="progress">
                    <div className="bar indeterminate"><div /></div>
                    <span>{outline ? 'Gliederung wird überarbeitet…' : 'Gliederung wird aus dem Datenpool erstellt…'}</span>
                  </div>
                )}
                {outline && <Markdown text={outline} />}
                {error && <p className="error">{error}</p>}
              </div>

              <aside className="chatbox">
                <h2>Feedback</h2>
                <div className="msgs">
                  {chat.filter((m) => m.role === 'user').length === 0 && (
                    <p className="empty">z. B. „Folie 3 nach vorn, Take-Home-Messages kürzen, mehr zu Anti-VEGF-Intervallen“</p>
                  )}
                  {chat.map((m, i) =>
                    m.role === 'user' ? (
                      <div key={i} className="msg me">{m.content}</div>
                    ) : (
                      i > 0 && <div key={i} className="msg ai">Gliederung aktualisiert ✓</div>
                    )
                  )}
                </div>
                <form onSubmit={sendFeedback} className="row">
                  <input value={feedback} onChange={(e) => setFeedback(e.target.value)} placeholder="Änderungswunsch…" maxLength={500} disabled={outlineBusy} />
                  <button type="submit" disabled={outlineBusy || !feedback.trim()}>Senden</button>
                </form>
                <button type="button" className="go" disabled={!outline || outlineBusy} onClick={createPptx}>
                  ✓ PowerPoint erstellen
                </button>
              </aside>
            </div>
          </section>
        )}
        {/* ---------------- PowerPoint ---------------- */}
        {step === 3 && (
          <section className="card pptx">
            <h1>PowerPoint: {query}</h1>
            {pptx.busy && (
              <div className="progress">
                <div className="bar indeterminate"><div /></div>
                <span>Folien werden aus der Outline gebaut und gestaltet…</span>
              </div>
            )}
            {error && <p className="error">{error}</p>}
            {pptx.url && (
              <div className="done">
                <div className="bigcheck">✓</div>
                <h2>Präsentation fertig</h2>
                <p className="lead">{pptx.name} · {pptx.slides} Folien · Download wurde gestartet</p>
                <div className="row center">
                  <a className="btn" href={pptx.url} download={pptx.name}>⬇ Erneut herunterladen</a>
                  <button type="button" className="ghost" onClick={() => setStep(2)}>← Outline anpassen</button>
                </div>
              </div>
            )}
            {!pptx.busy && !pptx.url && !error && null}
            {!pptx.busy && error && (
              <div className="row center"><button type="button" className="ghost" onClick={() => setStep(2)}>← Zurück zur Outline</button></div>
            )}
          </section>
        )}
      </main>

      {uploadFor && <UploadDialog paper={uploadFor} onClose={() => setUploadFor(null)} onSubmit={submitUpload} />}
    </div>
  );
}

function UploadDialog({ paper, onClose, onSubmit }) {
  const [file, setFile] = useState(null);
  const [license, setLicense] = useState('');
  const [busy, setBusy] = useState(false);

  async function go() {
    if (!file || !license) return;
    setBusy(true);
    await onSubmit(paper, file, license);
  }

  return (
    <div className="overlay" onClick={onClose}>
      <div className="dialog" onClick={(e) => e.stopPropagation()}>
        <h2>Volltext manuell hinzufügen</h2>
        <p className="title">{paper.title}</p>

        <label className="dropzone">
          <input type="file" hidden onChange={(e) => setFile(e.target.files[0] || null)} />
          {file ? <span>📄 {file.name} · {fmtBytes(file.size)}</span> : <span>Datei auswählen (PDF, Text …)</span>}
        </label>

        <fieldset className="license">
          <legend>Woher stammt dieser Volltext?</legend>
          <label className={license === 'free' ? 'on' : ''}>
            <input type="radio" name="lic" value="free" checked={license === 'free'} onChange={() => setLicense('free')} />
            <b>Frei verfügbar</b>
            <small>Open Access, Preprint, eigene Arbeit – darf ausgewertet werden</small>
          </label>
          <label className={license === 'paid' ? 'on' : ''}>
            <input type="radio" name="lic" value="paid" checked={license === 'paid'} onChange={() => setLicense('paid')} />
            <b>Kostenpflichtig erworben</b>
            <small>Wird nur als Nachweis abgelegt, nicht für Outline oder PowerPoint verwendet</small>
          </label>
        </fieldset>

        <div className="row end">
          <button type="button" className="ghost" onClick={onClose}>Abbrechen</button>
          <button type="button" disabled={!file || !license || busy} onClick={go}>
            {busy ? 'Speichert…' : 'Hinzufügen'}
          </button>
        </div>
      </div>
    </div>
  );
}

function Markdown({ text }) {
  const lines = text.split(/\r?\n/);
  const out = [];
  let list = [];
  const flush = () => {
    if (list.length) out.push(<ul key={out.length}>{list.map((l, i) => <li key={i}>{l}</li>)}</ul>);
    list = [];
  };
  lines.forEach((raw) => {
    const line = raw.replace(/\*\*(.+?)\*\*/g, '$1').trim();
    if (!line) return flush();
    if (/^#{1,3}\s/.test(line)) { flush(); out.push(<h3 key={out.length}>{line.replace(/^#+\s*/, '')}</h3>); }
    else if (/^[-*•]\s/.test(line)) list.push(line.replace(/^[-*•]\s*/, ''));
    else if (/^\d+[.)]\s/.test(line) && !/^\d+[.)]\s*Folie/i.test(line)) list.push(line.replace(/^\d+[.)]\s*/, ''));
    else { flush(); out.push(<p key={out.length}>{line}</p>); }
  });
  flush();
  return <div className="md">{out}</div>;
}
