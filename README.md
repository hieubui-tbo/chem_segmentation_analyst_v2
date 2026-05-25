# ⚗️ ChemSeg V17 — Chemical Import/Export Segmentation Tool

Internal analyst tool for multi-year chemical import/export data processing.

## Features
- 14-step processing pipeline (Upload → Export)
- Company entity resolution with Vietnamese-English alias detection
- Top 80% purchaser ranking with placeholder exclusion (V16)
- Enhanced entity standardization with risk flags (V17)
- IQR outlier detection per Year + Segment
- Full audit workbook export

## Local Development

```bash
npm install
npm run dev
```

Open http://localhost:5173

## Build for Production

```bash
npm run build
```

Output in `dist/` folder.
