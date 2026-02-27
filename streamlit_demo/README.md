# Streamlit Capabilities Demo

A complete, modular Streamlit application that demonstrates all major framework features in self-contained mini demos. Built as an internal training and reference project.

---

## Quick Start

### 1. Create the virtual environment

```bash
cd streamlit_demo
python -m venv .venv
```

### 2. Install dependencies

```bash
# Windows
.venv\Scripts\pip install -r requirements.txt

# Linux / macOS
.venv/bin/pip install -r requirements.txt
```

### 3. Generate mock data

```bash
# Windows
.venv\Scripts\python scripts\generate_mock_data.py

# Linux / macOS
.venv/bin/python scripts/generate_mock_data.py
```

### 4. Run the app

```bash
# Windows (double-click or from terminal)
run_app.bat

# Or manually
.venv\Scripts\python -m streamlit run app.py

# Linux / macOS
.venv/bin/python -m streamlit run app.py
```

The app opens at **http://localhost:8501**.

---

## Project Structure

```
streamlit_demo/
├── app.py                          # Main entry point & sidebar navigation
├── menu.py                         # Page registry and dynamic loader
├── pages/
│   ├── __init__.py
│   ├── data_input.py               # Input widgets demo
│   ├── visualization.py            # Charts, tables, and metrics demo
│   ├── crud_demo.py                # CRUD operations demo
│   ├── file_upload.py              # File upload & download demo
│   ├── process_runner.py           # Long-running process simulation
│   └── state_management.py         # Session state demo
├── data/
│   └── mock_data/                  # Generated CSV/JSON data files
│       ├── employees.csv
│       ├── sales.json
│       └── timeseries.csv
├── scripts/
│   └── generate_mock_data.py       # Deterministic mock data generator
├── requirements.txt
├── run_app.bat                     # Windows double-click launcher
└── README.md
```

---

## Demo Pages

### 1. Data Input (`pages/data_input.py`)

**What it shows:** Every major input widget Streamlit provides.

| Widget | Streamlit API |
|---|---|
| Text / password / textarea | `st.text_input`, `st.text_area` |
| Numbers / sliders | `st.number_input`, `st.slider`, `st.select_slider` |
| Select / multiselect | `st.selectbox`, `st.multiselect` |
| Date / time | `st.date_input`, `st.time_input` |
| Color / checkbox / radio | `st.color_picker`, `st.checkbox`, `st.radio` |
| Forms | `st.form`, `st.form_submit_button` |

**Why it exists:** Forms prevent unnecessary reruns when you have many inputs. Understanding widget keys and return values is foundational.

---

### 2. Data Visualization (`pages/visualization.py`)

**What it shows:** Tables, editable dataframes, KPI metrics, and charts.

- **KPI cards** with `st.metric` (value + delta)
- **Line chart** of visitors and signups over time
- **Bar chart** of revenue by product category
- **Scatter chart** of amount vs. quantity
- **Interactive DataFrame** with column configuration
- **Editable DataFrame** with `st.data_editor`

**Why it exists:** Data visualization is the core use case for Streamlit. This page shows the full range of built-in chart types and the powerful `st.data_editor` for inline editing.

---

### 3. CRUD Operations (`pages/crud_demo.py`)

**What it shows:** Full Create-Read-Update-Delete workflow on an in-memory dataset.

- **Browse** with multi-column filtering
- **Create** new records via a form
- **Update** existing records by selecting an ID
- **Delete** with confirmation

All changes persist in `st.session_state` for the current session.

**Why it exists:** Most real apps need CRUD. This demo shows how to build it without a database, using session state as the persistence layer.

---

### 4. File Handling (`pages/file_upload.py`)

**What it shows:** Uploading and downloading files.

- Upload CSV → auto-parsed into a DataFrame
- Upload JSON → displayed as a table or raw JSON
- Upload TXT → displayed in a text area
- Download a generated CSV, JSON, or custom text file

**Why it exists:** File I/O is essential for data tools. `st.file_uploader` and `st.download_button` are the two APIs you need.

---

### 5. Process Runner (`pages/process_runner.py`)

**What it shows:** Simulating a long-running pipeline with live feedback.

- **Progress bar** updated step-by-step
- **Streaming log output** rendered in a code block
- **Configurable speed** and **failure simulation**
- **Status indicators** (success / error)

**Why it exists:** Many internal tools run batch jobs, ETL pipelines, or CI tasks. This shows the pattern for real-time progress and log streaming.

---

### 6. State Management (`pages/state_management.py`)

**What it shows:** How `st.session_state` works under the hood.

- **Counter demo** — persisting a value across reruns
- **Callback pattern** — `on_change` / `on_click` hooks
- **Cross-page state** — reading data set by other demo pages
- **State inspector** — raw view of all session state keys

**Why it exists:** Session state is the #1 source of confusion for new Streamlit developers. This page demystifies it.

---

## Architecture Decisions

| Decision | Rationale |
|---|---|
| **One file per demo** | Each page is isolated — easy to understand, test, or remove independently. |
| **Dynamic import via `menu.py`** | Adding a new demo is a one-line change in the registry. No changes to `app.py`. |
| **No database** | All state is in-memory (`st.session_state`) or on disk (CSV/JSON). Zero infrastructure. |
| **Deterministic mock data** | Seeded random generator ensures reproducible outputs across runs. |
| **`.bat` launcher** | Windows users can double-click to start — no terminal knowledge needed. |

---

## Adding a New Demo Page

1. Create `pages/my_new_demo.py` with a `render()` function.
2. Register it in `menu.py`:

```python
DemoPage(
    key="my_new_demo",
    title="My New Demo",
    icon="🆕",
    description="Description of the new demo.",
    module_name="my_new_demo",
),
```

3. That's it. The sidebar and home page update automatically.

---

## Constraints

- **Python only** — no JavaScript or custom components.
- **Streamlit only** — no Flask, FastAPI, or other frameworks.
- **No cloud services** — everything runs locally.
- **No databases** — in-memory or local file storage only.
