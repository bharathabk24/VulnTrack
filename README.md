# VulnTrack — Vulnerability Dashboard

A local, browser-based dashboard for ManageEngine Endpoint Central vulnerability exports.
Supports **multiple days**, **day-over-day diff**, **resolution tracking**, and **multi-day trend**.

## Quick Start

### Option A — VS Code Live Server (recommended)
1. Unzip and open the folder in VS Code
2. Install **Live Server** extension (ritwickdey.LiveServer)
3. Right-click `index.html` → Open with Live Server

### Option B — Open directly in browser
Double-click `index.html` in Chrome or Edge.

> Internet required on first load for Chart.js + SheetJS from CDN.

---

## Daily Workflow

### Upload vulnerability file(s)
Click "Upload vuln file(s)" — select one or multiple .xlsx files at once.

Date detection:
- Filename with date like 2024-01-10 or 10-01-2024 → used automatically
- No date in filename → a prompt will ask for the date
- File with "Discovered Date" column → rows are grouped by that date automatically

### Upload resolution file (optional)
Click "Upload resolution file". Expected columns (flexible, case-insensitive):
  Computer Name | Vulnerability (or Patch Name) | Status | Date | Office | Notes

Status values: Resolved, Pending, In Progress, Not Applicable

---

## Tabs

| Tab        | Description |
|------------|-------------|
| Overview   | Day snapshot with metrics, charts, and deltas vs previous day |
| Records    | Paginated table with search, filters, and resolution status column |
| By Office  | Per-office cards with severity breakdown |
| Day Diff   | Compare any two days: new vulns, resolved, severity and office changes |
| Resolution | Patch/fix tracking with donut chart and resolution rate trend |
| Trend      | All days side-by-side: total count, severity stacked, resolution rate |

---

## File Structure

  vuln-dashboard/
  index.html   — markup
  style.css    — dark theme
  app.js       — all logic
  README.md

---

## Offline Mode

Download into libs/ subfolder and update script src tags in index.html:
  https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
  https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js
