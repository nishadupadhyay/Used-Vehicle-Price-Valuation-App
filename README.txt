# Used Vehicle Price Valuation — Streamlit starter

## Files to keep together
- `app.py`
- `best_used_car_price_model.pkl`
- `Used_Car_Price_Prediction_Clean.xls` (the supplied file is actually CSV text; the app reads it with `pd.read_csv`)
- `requirements.txt`

## Windows setup (PowerShell)
Open PowerShell in this folder and run:

```powershell
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

If PowerShell blocks activation, run this once in the current terminal only:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

If `py -3.13` is not available, use `python -m venv .venv` instead.

## Important compatibility note
The provided notebook's saved model warns as being serialized with scikit-learn 1.6.1. The requirements pin that version to improve compatibility. Use the same Python environment for installation and for `streamlit run app.py`.

## What the app does
- Loads the saved model pipeline.
- Uses dropdown options from the training dataset.
- Sends all 20 input features in the order/labels expected by the notebook.
- Displays a predicted price in rupees.

The model was trained on listing records from 2019–2021, so the listing-year input is kept within the years represented in the dataset. The displayed R² score (90.58%) comes from the supplied notebook's held-out test set.
