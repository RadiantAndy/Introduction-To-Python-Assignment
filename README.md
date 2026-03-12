---
title: Schneider Demo 2
emoji: ⚡
colorFrom: green
colorTo: teal
sdk: docker
pinned: false
---

# Schneider Electric – Lever Planning Tool

A React application for planning and visualising Schneider Electric emission-reduction levers.

## Features

- Upload trajectory baseline data (`df_2025.xlsx`) and levers (`levers.xlsx`)
- AI-assisted lever entry via chat (powered by HuggingFace Inference API)
- Scenario trajectory visualisation with line and bar charts
- Export levers input / levers tab as Excel

## Data files

| File | Description |
|------|-------------|
| `public/df_2025.xlsx` | Trajectory baseline (auto-loaded on startup) |
| `public/levers.xlsx` | Levers definitions (optional – can also be uploaded via UI) |

## Local development

```bash
npm install
npm run dev
```

## Deployment

The app is packaged as a Docker image (nginx on port 7860) for HuggingFace Spaces.
