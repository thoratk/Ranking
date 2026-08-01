from datetime import date
from io import BytesIO

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from excel_processor import process_workbook

app = FastAPI(title="F&O Friday Ranking Tracker")


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>F&O Friday Ranking Tracker</title>
  <style>
    :root {
      color-scheme: light dark;
      --bg: #0f1419;
      --card: #1a2332;
      --text: #e7ecf3;
      --muted: #9aa7b8;
      --accent: #3b82f6;
      --accent-hover: #2563eb;
      --border: #2d3a4d;
      --success: #22c55e;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", system-ui, sans-serif;
      background: linear-gradient(160deg, #0b1020, #111827 45%, #0f172a);
      color: var(--text);
      min-height: 100vh;
    }
    .wrap {
      max-width: 760px;
      margin: 0 auto;
      padding: 48px 20px 64px;
    }
    h1 {
      margin: 0 0 8px;
      font-size: 1.9rem;
      font-weight: 700;
    }
    .sub {
      color: var(--muted);
      margin-bottom: 28px;
      line-height: 1.5;
    }
    .card {
      background: rgba(26, 35, 50, 0.92);
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 24px;
      box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35);
    }
    label {
      display: block;
      font-weight: 600;
      margin-bottom: 8px;
    }
    .field { margin-bottom: 20px; }
    input[type="file"],
    input[type="date"] {
      width: 100%;
      padding: 12px 14px;
      border-radius: 10px;
      border: 1px solid var(--border);
      background: #0f172a;
      color: var(--text);
      font-size: 1rem;
    }
    button {
      width: 100%;
      padding: 14px 18px;
      border: none;
      border-radius: 10px;
      background: var(--accent);
      color: white;
      font-size: 1rem;
      font-weight: 600;
      cursor: pointer;
    }
    button:hover { background: var(--accent-hover); }
    button:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }
    .note {
      margin-top: 20px;
      padding: 14px 16px;
      border-radius: 10px;
      background: rgba(59, 130, 246, 0.08);
      border: 1px solid rgba(59, 130, 246, 0.25);
      color: var(--muted);
      font-size: 0.92rem;
      line-height: 1.55;
    }
    .status {
      margin-top: 16px;
      min-height: 1.2rem;
      color: var(--success);
      font-weight: 600;
    }
    ul { margin: 8px 0 0 18px; padding: 0; }
    li { margin: 4px 0; }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>F&O Friday Ranking Tracker</h1>
    <p class="sub">
      Upload your F&O Excel file and pick a base date.
      Columns <strong>A, B, C</strong> are copied exactly from your file.
      Calculations start from column <strong>D</strong>.
    </p>

    <div class="card">
      <form id="upload-form">
        <div class="field">
          <label for="file">Excel file (.xlsx)</label>
          <input id="file" name="file" type="file" accept=".xlsx" required />
        </div>
        <div class="field">
          <label for="base_date">Base date (column D price)</label>
          <input id="base_date" name="base_date" type="date" required />
        </div>
        <button id="submit-btn" type="submit">Generate ranked Excel</button>
        <div id="status" class="status"></div>
      </form>
    </div>

    <div class="note">
      <strong>Output layout</strong>
      <ul>
        <li><strong>A–C</strong> Same as your uploaded file (Script, Script-EQ, Sector, etc.)</li>
        <li><strong>D</strong> Base price (your selected date)</li>
        <li><strong>E</strong> Current price</li>
        <li><strong>F</strong> Points (E − D)</li>
        <li><strong>G</strong> % diff from base to current</li>
        <li><strong>H</strong> Today's rank</li>
        <li><strong>I onward</strong> Friday ranks (color-coded)</li>
        <li><strong>Friday Top 20</strong> sheet — top 20 stock names per Friday</li>
      </ul>
      <p style="margin:12px 0 0;color:var(--muted);font-size:0.9rem;">
        Rank columns (H and I onward) are color-coded:
        1–20 blue, 21–40 green, 41–60 yellow, 61–90 brown, 91+ no color.
      </p>
    </div>
  </div>

  <script>
    const form = document.getElementById("upload-form");
    const statusEl = document.getElementById("status");
    const submitBtn = document.getElementById("submit-btn");

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      statusEl.textContent = "Fetching prices and building ranks...";
      submitBtn.disabled = true;

      const formData = new FormData(form);

      try {
        const response = await fetch("/process", {
          method: "POST",
          body: formData,
        });

        if (!response.ok) {
          const payload = await response.json().catch(() => ({}));
          throw new Error(payload.detail || "Processing failed.");
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        const stamp = new Date().toISOString().slice(0, 10);
        anchor.href = url;
        anchor.download = `fo-ranked-${stamp}.xlsx`;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        URL.revokeObjectURL(url);

        statusEl.textContent = "Done. Download started.";
      } catch (error) {
        statusEl.textContent = error.message;
        statusEl.style.color = "#f87171";
      } finally {
        submitBtn.disabled = false;
      }
    });
  </script>
</body>
</html>"""


@app.post("/process")
async def process(
    file: UploadFile = File(...),
    base_date: str = Form(...),
):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")

    try:
        parsed_base_date = date.fromisoformat(base_date)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="Invalid base date.") from exc

    content = await file.read()
    try:
        output = process_workbook(content, parsed_base_date)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Failed to process workbook: {exc}",
        ) from exc

    return StreamingResponse(
        BytesIO(output),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="fo-ranked.xlsx"'},
    )


if __name__ == "__main__":
    import os

    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
