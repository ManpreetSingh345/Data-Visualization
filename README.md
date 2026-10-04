# 📊 GlassVibe — Advanced Data Visualizer

A sleek, client-side data visualization dashboard built with vanilla HTML, CSS, and JavaScript. Upload any Excel spreadsheet and instantly get interactive charts and key statistics — no backend, no frameworks, no installs required.

---

## ✨ Features

- **Excel File Import** — Drag & drop or browse to upload `.xlsx` / `.xls` files
- **Multi-Sheet Support** — Switch between sheets from a dropdown selector
- **Auto KPI Cards** — Automatically displays Total Rows, Total Columns, and a numeric column sum
- **3 Interactive Charts:**
  - 📊 **Bar Chart** — Distribution Analysis
  - 📈 **Line Chart** — Trend Analysis  
  - 🍩 **Doughnut Chart** — Proportional Share
- **Axis Controls** — Pick any column for X and Y axes on each chart independently
- **Smart Column Detection** — Auto-classifies columns as categorical or numerical
- **Dark / Light Theme** — Toggle with persistence via `localStorage`
- **Glassmorphism UI** — Frosted glass aesthetic with animated background blobs

---

## 🛠️ Tech Stack

| Technology | Purpose |
|---|---|
| HTML5 / CSS3 | Structure & glassmorphism styling |
| Vanilla JavaScript (ES Modules) | App logic & state management |
| [Chart.js](https://www.chartjs.org/) | Chart rendering |
| [SheetJS (xlsx)](https://sheetjs.com/) | Excel file parsing |
| [Font Awesome 6](https://fontawesome.com/) | Icons |

---

## 🚀 Getting Started

No installation or build step is needed.

1. **Clone or download** this repository
2. Open `index.html` in any modern browser (Chrome, Edge, Firefox)
3. Upload an Excel file and explore your data!

> **Note:** Because it uses ES Modules (`type="module"`), open via a local server or the browser's file system directly. If charts don't render, try `Live Server` in VS Code.

---

## 📁 Project Structure

```
Data visualization/
├── index.html    # App shell & dashboard layout
├── style.css     # Glassmorphism theme, dark/light modes
├── app.js        # Core logic: file parsing, data analysis, chart rendering
└── README.md
```

---

## 📋 How It Works

1. User uploads an `.xlsx` / `.xls` file via drag-and-drop or file picker
2. **SheetJS** parses the workbook and extracts sheet data
3. App auto-detects **categorical** vs **numerical** columns (using a 70% numeric threshold heuristic)
4. KPI cards are computed and displayed
5. **Chart.js** renders Bar, Line, and Doughnut charts with the auto-selected axes
6. User can change axes via dropdowns to re-render charts on the fly
7. Theme toggle re-renders all charts with updated color theming

---

## 🎓 Course Info

**Minor Project** — Data Visualization  
CEC · ST 2026
