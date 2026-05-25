# ChemSeg Analyst — V12

Chemical Import/Export Segmentation & Price Analysis Tool.

## Features
- Multi-year workbook processing (each year = separate sheet)
- Company Entity Resolution Engine (VN/EN legal suffixes, fused suffix detection, dedup)
- Numeric Backfill Engine (recovers missing Unit Price / Total Value / Quantity)
- Granular Data Quality (Missing, Zero, Negative, Invalid, Text Placeholder, Backfilled, Unrecoverable)
- Industry Master matching (4-column upload)
- Segment-specific keyword dictionaries (columnar Excel import)
- Unit conversion to Metric Tons (TNE, KGM, GRM, LBS, TONNE)
- IQR outlier detection per year per segment
- Full Excel export with Before/After outlier sheets + audit logs

## Deploy on Vercel
1. Push this repo to GitHub
2. Go to [vercel.com/new](https://vercel.com/new)
3. Import the GitHub repo
4. Click **Deploy** — no config needed

## Local Development
```bash
npm install
npm run dev
```
Open [http://localhost:3000](http://localhost:3000)
