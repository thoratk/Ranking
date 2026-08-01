# Ranking

F&O Friday Ranking Tracker — upload an Excel file, pick a base date, and download ranked results with weekly Friday top-20 summary.

## Run locally

```bat
python app.py
```

Open http://127.0.0.1:8000

## Deploy on Render

1. Connect repo `thoratk/Ranking` on [render.com](https://render.com)
2. **Root Directory:** leave **blank** (repo root)
3. **Build:** `pip install -r requirements.txt`
4. **Start:** `uvicorn app:app --host 0.0.0.0 --port $PORT`
5. After deploy, open `/health` — must show `v2.4-close-values`

If output looks old: **Manual Deploy → Clear build cache & deploy**.

## Features

- Close prices only (2 decimals), NSE holiday aware
- % Diff = (Points / Base Price) × 100
- Base/current prices and Friday ranks with color bands
- **Friday Top 20** sheet with Entry/Exit tracking
