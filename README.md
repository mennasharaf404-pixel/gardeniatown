# Gardenia Town Streamlit Dashboard

## Run

1. Open PowerShell in this folder.
2. Create/activate a virtual environment if desired.
3. Install requirements:

```powershell
python -m pip install -r requirements.txt
```

4. Start Streamlit:

```powershell
python -m streamlit run app.py
```

The dashboard reads `data/All Inventory Project(1).xlsx`, sheet `Gardenia Town`.

## Features
- ALBA vs ORCHID comparison
- Dynamic KPIs
- Sidebar filters
- Status chart
- Sales conversion chart
- Available inventory by building
- Unit type mix
- Search by unit/client
- Filtered unit detail table
