# Streamlit Demo Playground

A self-contained reference application that demonstrates every major Streamlit capability in a clean, modular, and scalable architecture. Built for internal training — all data is generated locally, no cloud services or databases required.

---

## Quick Start

### 1. Create & activate a virtual environment

```bash
# Windows
cd streamlit_demo
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
cd streamlit_demo
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the app

```bash
streamlit run main_menu.py
```

Or on **Windows**, simply double-click **`run_app.bat`** (it activates the venv and launches the app automatically).

---

## Project Structure

```
streamlit_demo/
├── main_menu.py              # Main entry point – page config & sidebar routing
├── menu.py                   # Page registry + home page renderer
├── pages/                    # One module per demo (each exposes a render() function)
│   ├── __init__.py
│   ├── data_input.py         # Input widgets demo
│   ├── visualization.py      # Tables, charts, KPI metrics
│   ├── crud_demo.py          # Create / Read / Update / Delete operations
│   ├── file_upload.py        # File upload & download
│   ├── process_runner.py     # Simulated long-running task with live logs
│   └── state_management.py   # Session state patterns
├── data/
│   ├── __init__.py
│   ├── loader.py             # Centralised data-loading helper
│   └── mock_data/            # Auto-generated CSV / JSON files
├── scripts/
│   ├── __init__.py
│   └── generate_mock_data.py # Standalone script to (re)generate mock data
├── requirements.txt
├── run_app.bat               # Windows double-click launcher
└── README.md                 # This file
```

---

## Demo Pages

### Data Input
**File:** `pages/data_input.py`

Showcases every common input widget:
- Text input, text area, number input
- Slider, select box, multi-select
- Date and time pickers
- Checkbox, radio buttons
- A grouped form with a submit button (`st.form`)

**Why it exists:** Understanding how widgets work and interact with session state is the foundation of every Streamlit app.

---

### Visualization
**File:** `pages/visualization.py`

Demonstrates data display and charting:
- `st.dataframe` – scrollable, sortable read-only table
- `st.data_editor` – editable inline table
- `st.line_chart`, `st.bar_chart`, `st.scatter_chart` – built-in chart types
- `st.metric` – KPI cards with delta indicators
- Dynamic filtering controls that update charts in real time

**Why it exists:** Dashboards and data exploration tools are the most common use case for internal Streamlit apps.

---

### CRUD Operations
**File:** `pages/crud_demo.py`

Manages an in-memory employee table with full CRUD capability:
- **Create** – add new records via a form
- **Read** – browse the full dataset
- **Update** – edit cells inline with `st.data_editor`
- **Delete** – multi-select IDs and remove them
- **Reset** – restore the original mock data

Changes are persisted across reruns using `st.session_state`.

**Why it exists:** Nearly every internal tool revolves around a table of records that users need to manage.

---

### File Handling
**File:** `pages/file_upload.py`

Upload and download files:
- **Upload** CSV, TXT, or JSON files and instantly preview their contents
- **Download** the employees dataset as CSV, JSON, or TXT

**Why it exists:** Many internal tools need to accept user-provided data files and produce downloadable exports.

---

### Process Runner
**File:** `pages/process_runner.py`

Simulates a long-running background task:
- A button triggers the process
- A `st.progress` bar tracks completion
- Live log lines stream into the UI via `st.empty`
- `st.status` shows a running / complete indicator
- Configurable step count and warning probability

**Why it exists:** Internal automation tools often wrap shell scripts or ETL jobs. This pattern shows how to keep users informed while a task runs.

---

### State Management
**File:** `pages/state_management.py`

Deep dive into `st.session_state`:
- **Counter** – a persistent value that survives reruns, with code example
- **Shared Notepad** – a value written on one page that any other module can read
- **State Inspector** – dumps the entire session state dictionary for debugging

**Why it exists:** Streamlit reruns the script on every interaction. Understanding session state is essential for building stateful apps.

---

## Mock Data

The `scripts/generate_mock_data.py` script produces:

| File | Records | Used By |
|---|---|---|
| `employees.csv` / `.json` | 50 employees | CRUD, Visualization, File Handling |
| `sales.csv` | 200 sales transactions | Visualization charts |

If mock data files are missing when the app starts, they are generated automatically via `data/loader.py`.

To regenerate manually:

```bash
python scripts/generate_mock_data.py
```

---

## Adding a New Demo

1. Create `pages/my_new_demo.py` with a `render()` function.
2. Import it in `menu.py` and append an entry to the `PAGES` list.
3. Done — the sidebar picks it up automatically.

---

## Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | UI framework |
| `pandas` | Data manipulation & display |

All other functionality uses Python's standard library.

---

## License

Internal use only — built for training and reference purposes.
