## 📋 Usage

Run with default settings (AAPL, Jan–May 2026):
\`\`\`bash
python gappers.py
\`\`\`

Or specify your own parameters:
\`\`\`bash
python gappers.py --ticker TSLA --start 2026-01-01 --end 2026-06-01 --threshold 0.02
\`\`\`

**Arguments:**
- `--ticker` — stock symbol (default: AAPL)
- `--start` / `--end` — date range (YYYY-MM-DD)
- `--threshold` — minimum gap % to include (default: 0.015 = 1.5%)
- `--direction` — which gaps to detect: `up` (default), `down`, or `both`
- `--output` — custom output filename