# Research Radar — Recherche mit Datenpool

Prototyp aus dem Mini-Hackathon beim Ophthalmologen-Kongress (25.–26.09.2026).

Recherche-Werkzeug für ophthalmologische Update-Referate: Live-Suche in PubMed,
Sammeln der relevanten Arbeiten in einem Datenpool, eigene Dokumente je Paper
hochladen, daraus eine Gliederung erzeugen und als PPTX ausgeben.

> Aus diesem Stand ist eine zweite, anders ausgerichtete Fassung entstanden:
> [research-radar-v2](https://github.com/NiclasBayer/research-radar-v2) — dort
> stehen Volltext-Belege, Crosscheck und Abbildungen im Vordergrund statt des
> Datenpools.

## Aufbau

- `frontend/` — React + Vite
- `backend/` — Express: PubMed (NCBI E-utilities), LLM über myGenAssist, PPTX via pptxgenjs

Im Betrieb ist es **ein Dienst**: Express liefert die API und das gebaute Frontend aus.

## Lokal starten

```bash
npm install --prefix backend
npm install --prefix frontend
echo "MGA_API_KEY=..." > backend/.env
npm start --prefix backend        # Backend auf :3000
npm run dev --prefix frontend     # Frontend auf :8080
```

## Konfiguration

Alle Werte kommen aus der Umgebung — **im Repo steht kein Schlüssel**.

| Variable | Zweck |
|---|---|
| `MGA_API_KEY` | Zugang zum LLM-Gateway. Ohne ihn läuft die Suche, die KI-Schritte antworten mit 503. |
| `APP_PASSWORD` | Schützt die öffentlich erreichbare Instanz per Passwortabfrage. Nicht gesetzt = kein Schutz (nur lokal sinnvoll). |
| `MGA_MODEL` | Modell-ID, Vorgabe `claude-opus-5`. |
| `PORT` | Vom Hoster gesetzt, lokal 3000. |

Alle Beispieldaten sind synthetisch — keine echten Patientendaten.
