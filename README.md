# Gantt Chart Generator

A zero-install, local HTML/JS application for creating professional Gantt charts with a think-cell-inspired aesthetic and PowerPoint export.

## Quick Start

1. **Open directly** – Double-click `index.html` in any modern browser.
2. **Or use a local server** (recommended for best compatibility):
   ```bash
   # Python 3
   cd GanttChart
   python3 -m http.server 8080

   # Node.js (npx, no install needed)
   npx serve .
   ```
   Then open `http://localhost:8080`.

> **Windows users**: The `file://` protocol works fine in Chrome and Edge. On Firefox, localStorage may be restricted under `file://` — use a local server instead.

## Features

| Feature | Description |
|---------|-------------|
| **Task Management** | Add, edit, delete tasks and milestones inline |
| **Interactive Gantt** | Drag bars to reschedule, resize to change duration |
| **View Modes** | Day / Week / Month / Quarter / Year |
| **Dependencies** | Link tasks with comma-separated IDs |
| **Color Coding** | Per-task color picker with preset palette |
| **Progress Tracking** | 0–100% progress per task |
| **PPTX Export** | One-click PowerPoint export (widescreen 16:9) |
| **JSON Import/Export** | Save and reload projects as JSON files |
| **Auto-Save** | All changes persist in localStorage automatically |
| **Sample Project** | Click "Sample" to see a complete demo project |

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl/⌘ + N` | Add new task |
| `Ctrl/⌘ + E` | Export to PPTX |
| `Delete` / `Backspace` | Delete selected task |

## File Structure

```
GanttChart/
├── index.html          # Entry point – open this
├── css/
│   └── style.css       # Think-cell-inspired styling
├── js/
│   ├── app.js          # Task CRUD, state, table UI, persistence
│   ├── gantt-render.js # Frappe Gantt integration
│   └── pptx-export.js  # PptxGenJS PowerPoint export
└── README.md
```

## Libraries (loaded via CDN)

- **[Frappe Gantt v1.0.3](https://github.com/nicedoc/frappe-gantt)** — Interactive SVG Gantt chart renderer
- **[PptxGenJS v3.12.0](https://gitbrent.github.io/PptxGenJS/)** — Client-side PowerPoint file generation

No npm, no build step, no Node.js required.

## Browser Compatibility

| Browser | Status | Notes |
|---------|--------|-------|
| **Chrome 90+** | Fully supported | Recommended |
| **Edge 90+** | Fully supported | Chromium-based, same as Chrome |
| **Firefox 90+** | Supported | Use local server for localStorage under `file://` |
| **Safari 15+** | Supported | Minor date-input styling differences |
| **Opera 80+** | Supported | Chromium-based |
| **IE 11** | Not supported | ES6+ features required |

### Notes

- The app uses `<input type="date">`, which renders a native date picker. Appearance varies by browser and OS.
- PPTX export uses `Blob` and `URL.createObjectURL`, supported in all modern browsers.
- Frappe Gantt renders to SVG. Performance is best with < 100 tasks.
- localStorage stores data per-origin. Under `file://`, some browsers treat each file as a separate origin.

## PPTX Export Details

The exported PowerPoint file contains:
- Widescreen (13.33" × 7.5") slide layout
- Time-scale header with month labels
- Alternating row shading
- Color-coded task bars with progress overlay
- Diamond-shaped milestone markers
- Dependency arrows (horizontal → vertical → horizontal)
- Dashed "Today" line
- Task labels in the left column

All elements are native PowerPoint shapes (not images), so they remain editable after export.

## Data Format

Projects are stored as JSON arrays. Example:

```json
[
  {
    "id": "task_1",
    "name": "Project Kickoff",
    "start": "2026-02-19",
    "end": "2026-02-19",
    "progress": 100,
    "color": "#e8923f",
    "dependencies": "",
    "isMilestone": true
  },
  {
    "id": "task_2",
    "name": "Design Phase",
    "start": "2026-02-28",
    "end": "2026-03-11",
    "progress": 40,
    "color": "#6aab73",
    "dependencies": "task_1",
    "isMilestone": false
  }
]
```

## License

MIT
