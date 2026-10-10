import joblib
import pandas as pd
import streamlit as st
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
MODEL_PATH = BASE_DIR / "best_used_car_price_model.pkl"
DATA_PATH = BASE_DIR / "Used_Car_Price_Prediction_Clean.xls"

st.set_page_config(page_title="Used Vehicle Price Valuation", page_icon="V", layout="wide")

st.markdown(r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
:root { --ink:#172b46; --muted:#5f7187; --line:#dce5ef; --blue:#173553; --teal:#258a91; --surface:#ffffff; --page:#f3f7fc; }
html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
h1,h2,h3,h4 { font-family:'Manrope','DM Sans',sans-serif !important; letter-spacing:-.025em; }
.stApp { background:linear-gradient(180deg,#f3f7fc 0%,#f8fafc 52%,#eef4fa 100%); color:var(--ink); }
.block-container { max-width:1280px; padding-top:1.7rem; padding-bottom:2.4rem; }
.hero { position:relative; overflow:hidden; display:flex; align-items:center; justify-content:space-between; gap:2rem;
  padding:2.1rem 2.35rem; margin-bottom:1.25rem; border-radius:24px;
  background:linear-gradient(118deg,#10243b 0%,#173553 54%,#258a91 100%); color:white;
  box-shadow:0 18px 45px rgba(19,47,78,.16); }
.hero:after { content:''; position:absolute; width:300px;height:300px;border:1px solid rgba(255,255,255,.12);border-radius:50%;right:105px;top:-185px;box-shadow:0 0 0 35px rgba(255,255,255,.035),0 0 0 70px rgba(255,255,255,.025); }
.hero-copy { position:relative; z-index:2; max-width:680px; }
.eyebrow { font-size:.75rem; text-transform:uppercase; letter-spacing:.18em; font-weight:700; color:#9edbd8; margin-bottom:.8rem; }
.hero h1 { color:#fff !important; font-size:clamp(2rem,4vw,3.05rem); line-height:1.12; margin:0 0 .75rem 0; }
.hero p { color:#d7e5f1; font-size:1.02rem; line-height:1.7; margin:0; max-width:620px; }
.hero-art { z-index:1; flex:0 0 230px; width:230px; opacity:.98; }
.section-card { background:#fff; border:1px solid var(--line); border-radius:20px; padding:1.35rem 1.45rem .7rem; margin:1rem 0; box-shadow:0 8px 24px rgba(21,38,59,.035); }
.section-heading { display:flex; align-items:center; gap:.85rem; margin-bottom:1rem; }
.section-number { display:flex; align-items:center; justify-content:center; width:40px;height:40px; flex-shrink:0; border-radius:12px; color:#fff; background:linear-gradient(135deg,#173553,#258a91); font-family:'Manrope',sans-serif; font-weight:800; }
.section-heading h3 { font-size:1.17rem; margin:0; color:var(--ink); }
.section-heading p { margin:.2rem 0 0; color:var(--muted); font-size:.86rem; }
div[data-testid="stSelectbox"] label, div[data-testid="stNumberInput"] label { color:#334b66; font-weight:600; font-size:.88rem; }
.stSelectbox div[data-baseweb="select"] > div, .stNumberInput div[data-baseweb="input"] > div { border-radius:10px; border-color:#dce5ef; }
.stButton > button, div[data-testid="stFormSubmitButton"] > button { border-radius:12px; min-height:3.1rem; font-weight:700; letter-spacing:.01em; background:linear-gradient(100deg,#173553,#258a91); border:0; box-shadow:0 8px 18px rgba(23,53,83,.18); }
.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover { border:0; filter:brightness(1.06); }
div[data-testid="stMetric"] { background:#fff !important; border:1px solid #dce6ef !important; padding:1.25rem 1.35rem !important; border-radius:18px !important; box-shadow:0 8px 22px rgba(21,38,59,.045); }
div[data-testid="stMetric"] label, div[data-testid="stMetricLabel"] p { color:#53657a !important; font-weight:600 !important; }
div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] div, div[data-testid="stMetricValue"] > div { color:#172b46 !important; -webkit-text-fill-color:#172b46 !important; font-family:'Manrope',sans-serif !important; font-weight:800 !important; }
.result-heading { margin-top:1.6rem; padding-top:.2rem; }
.result-heading h2 { color:#172b46; font-size:1.6rem; }
.result-panel { border-radius:18px; padding:1.25rem 1.5rem; background:linear-gradient(110deg,#eaf2fb 0%,#e7f6f3 100%); border:1px solid #d5e5eb; margin-bottom:.8rem; }
.result-label { color:#52677d; font-size:.9rem; font-weight:600; }
.result-price { color:#173553; font-family:'Manrope',sans-serif; font-size:clamp(2.2rem,4vw,3.25rem); line-height:1.25; font-weight:800; letter-spacing:-.04em; }
.result-sub { color:#60758a; font-size:.85rem; margin-top:.25rem; }
.footer-line { border-top:1px solid #dce5ef; margin-top:2rem; padding-top:1rem; color:#8290a1; font-size:.8rem; }
@media (max-width:760px) { .hero { padding:1.5rem; } .hero-art { display:none; } .section-card { padding:1rem .9rem .5rem; } }
</style>
<div class="hero">
  <div class="hero-copy">
    <div class="eyebrow">Automotive intelligence · Price estimation</div>
    <h1>Used Vehicle<br>Price Valuation</h1>
    <p>Discover an estimated resale value based on vehicle specifications, usage history and listing details.</p>
  </div>
  <div class="hero-art" aria-hidden="true">
    <svg viewBox="0 0 280 170" width="100%" xmlns="http://www.w3.org/2000/svg">
      <defs><linearGradient id="body" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#e7f5fa"/><stop offset="1" stop-color="#80c7ce"/></linearGradient></defs>
      <circle cx="150" cy="85" r="78" fill="none" stroke="#ffffff" stroke-opacity=".13" stroke-width="1.5"/>
      <path d="M29 111 L42 81 Q46 72 59 68 L90 61 L116 34 Q122 28 134 28 L180 28 Q191 29 199 40 L220 70 L241 78 Q251 82 254 94 L258 111 L248 122 L34 122 Z" fill="url(#body)" stroke="#f5fbff" stroke-width="2.5" stroke-linejoin="round"/>
      <path d="M101 60 L121 38 Q124 35 134 35 L153 35 L153 60 Z" fill="#193d5b"/>
      <path d="M161 35 L179 36 Q187 37 193 46 L207 66 L161 66 Z" fill="#193d5b"/>
      <path d="M43 84 L92 70 L209 75" fill="none" stroke="#ffffff" stroke-opacity=".6" stroke-width="2"/>
      <path d="M35 103 L53 103" stroke="#f3d7a0" stroke-width="5" stroke-linecap="round"/>
      <path d="M235 91 L250 95" stroke="#f5b0a8" stroke-width="5" stroke-linecap="round"/>
      <circle cx="81" cy="119" r="22" fill="#10243b" stroke="#c7e3e9" stroke-width="4"/><circle cx="81" cy="119" r="9" fill="#b8dce2"/>
      <circle cx="211" cy="119" r="22" fill="#10243b" stroke="#c7e3e9" stroke-width="4"/><circle cx="211" cy="119" r="9" fill="#b8dce2"/>
      <path d="M20 145 H260" stroke="#ffffff" stroke-opacity=".28" stroke-width="2" stroke-linecap="round"/>
      <path d="M44 153 H91 M104 153 H141 M154 153 H191 M204 153 H236" stroke="#8ad0d0" stroke-opacity=".8" stroke-width="3" stroke-linecap="round"/>
    </svg>
  </div>
</div>
""", unsafe_allow_html=True)

@st.cache_resource
def load_model():
    return joblib.load(MODEL_PATH)

@st.cache_data
def load_dataset():
    # The supplied .xls file is comma-separated text (CSV), not an Excel workbook.
    return pd.read_csv(DATA_PATH)

if not MODEL_PATH.exists():
    st.error(f"Model file not found: {MODEL_PATH.name}. Keep it in the same folder as app.py.")
    st.stop()
if not DATA_PATH.exists():
    st.error(f"Dataset file not found: {DATA_PATH.name}. Keep it in the same folder as app.py.")
    st.stop()

try:
    df = load_dataset()
    model = load_model()
except Exception as exc:
    st.error("The model or dataset could not be loaded.")
    st.exception(exc)
    st.info("If the error mentions scikit-learn version compatibility, install scikit-learn==1.6.1 in the same Python environment used to run Streamlit.")
    st.stop()

feature_columns = [
    "fuel_type", "kms_run", "city", "body_type", "transmission", "variant",
    "assured_buy", "registered_city", "registered_state", "rto", "source",
    "make", "model", "total_owners", "fitness_certificate", "warranty_avail",
    "ad_year", "ad_month", "ad_dayofweek", "car_age"
]
missing = [col for col in feature_columns + ["sale_price"] if col not in df.columns]
if missing:
    st.error("The dataset is missing columns required by the trained model: " + ", ".join(missing))
    st.stop()

def choices(column):
    values = df[column].dropna().unique().tolist()
    return sorted(values, key=lambda x: str(x).lower())

def filtered_choices(column, filters):
    """Return dataset values matching the selected parent fields."""
    subset = df
    for field, value in filters.items():
        subset = subset[subset[field].astype(str) == str(value)]
    values = subset[column].dropna().unique().tolist()
    return sorted(values, key=lambda x: str(x).lower())

def display_value(value, column=None):
    """Format display labels without changing the raw values passed to the model."""
    text = str(value)
    if column == "rto":
        return text.upper()
    text = text.replace("_", " ")
    return text[:1].upper() + text[1:]

def select_index(column, preferred):
    vals = choices(column)
    for item in preferred:
        for i, val in enumerate(vals):
            if str(val).lower() == item.lower():
                return i
    return 0

with st.container():
    st.markdown('<div class="section-card"><div class="section-heading"><div class="section-number">01</div><div><h3>Vehicle profile</h3><p>Core details that influence resale value</p></div></div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        make = st.selectbox("Make / brand", choices("make"), index=select_index("make", ["maruti"]), format_func=display_value)
        model_name = st.selectbox("Model", choices("model"), index=select_index("model", ["swift"]), format_func=display_value)
        variant = st.selectbox("Variant", choices("variant"), format_func=display_value)
        fuel_type = st.selectbox("Fuel type", choices("fuel_type"), format_func=display_value)
    with c2:
        body_type = st.selectbox("Body type", choices("body_type"), format_func=display_value)
        transmission = st.selectbox("Transmission", choices("transmission"), format_func=display_value)
        kms_run = st.number_input("Kilometres driven", min_value=0, max_value=2000000, value=int(df["kms_run"].median()), step=1000)
        car_age = st.number_input("Vehicle age (years)", min_value=0, max_value=40, value=int(df["car_age"].median()), step=1)
    with c3:
        total_owners = st.selectbox("Number of previous/current owners", sorted(int(x) for x in df["total_owners"].dropna().unique()))
        assured_buy = st.selectbox("Assured-buy listing?", options=[("Yes", 1), ("No", 0)], format_func=lambda x: x[0])[1]
        fitness_certificate = st.selectbox("Fitness certificate available?", options=[("Yes", 1.0), ("No", 0.0)], format_func=lambda x: x[0])[1]
        warranty_avail = st.selectbox("Warranty available?", options=[("Yes", 1), ("No", 0)], format_func=lambda x: x[0])[1]
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-heading"><div class="section-number">02</div><div><h3>Location & listing</h3><p>Registration and listing context</p></div></div>', unsafe_allow_html=True)
    d1, d2, d3 = st.columns(3)
    with d1:
        registered_state = st.selectbox(
            "Registration state", choices("registered_state"),
            index=select_index("registered_state", ["gujarat"]),
            format_func=display_value, key="registered_state"
        )
        registration_cities = filtered_choices("registered_city", {"registered_state": registered_state})
        if not registration_cities:
            st.warning("No registration cities are available for this state in the training dataset.")
            registration_cities = ["Unknown"]
        registered_city = st.selectbox(
            "Registration city", registration_cities,
            format_func=display_value, key=f"registered_city_{registered_state}"
        )
        state_rtos = filtered_choices("rto", {"registered_state": registered_state})
        if not state_rtos:
            st.warning("No RTO codes are available for this state in the training dataset.")
            state_rtos = ["Unknown"]
        rto = st.selectbox(
            "RTO code", state_rtos,
            format_func=lambda value: display_value(value, "rto"),
            key=f"rto_{registered_state}"
        )
    with d2:
        city = st.selectbox("Listing city", choices("city"), format_func=display_value)
        source = st.selectbox("Listing source", choices("source"), format_func=display_value)
        ad_year = st.selectbox("Listing year", sorted(int(x) for x in df["ad_year"].dropna().unique()))
    with d3:
        ad_month = st.selectbox("Listing month", list(range(1, 13)), index=3, format_func=lambda x: pd.Timestamp(2024, x, 1).strftime("%B"))
        ad_dayofweek = st.selectbox("Listing weekday", list(range(7)), index=2, format_func=lambda x: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][x])
    st.markdown('</div>', unsafe_allow_html=True)

    submitted = st.button("Calculate estimated value  →", type="primary", use_container_width=True)

if submitted:
    input_data = pd.DataFrame([{
        "fuel_type": fuel_type, "kms_run": int(kms_run), "city": city,
        "body_type": body_type, "transmission": transmission, "variant": variant,
        "assured_buy": int(assured_buy), "registered_city": registered_city,
        "registered_state": registered_state, "rto": rto, "source": source,
        "make": make, "model": model_name, "total_owners": int(total_owners),
        "fitness_certificate": float(fitness_certificate), "warranty_avail": int(warranty_avail),
        "ad_year": float(ad_year), "ad_month": float(ad_month),
        "ad_dayofweek": float(ad_dayofweek), "car_age": float(car_age),
    }], columns=feature_columns)
    try:
        prediction = float(model.predict(input_data)[0])
        st.markdown('<div class="result-heading"><h2>Estimated selling price</h2></div>', unsafe_allow_html=True)
        if prediction < 0:
            st.warning("The model returned a negative estimate. Review the inputs and training data before interpreting this result.")
        else:
            left, right = st.columns([2, 1])
            with left:
                st.markdown(f'<div class="result-panel"><div class="result-label">Predicted vehicle value</div><div class="result-price">₹{prediction:,.0f}</div><div class="result-sub">Estimated resale price</div></div>', unsafe_allow_html=True)
            with right:
                st.metric("Model test R²", "90.58%")
                st.caption("Recorded test-set score")
            with st.expander("Review vehicle details used for prediction"):
                review = input_data.copy()
                for col in ["fuel_type", "city", "body_type", "transmission", "variant", "registered_city", "registered_state", "rto", "source", "make", "model"]:
                    review[col] = review[col].map(lambda value: display_value(value, col))
                st.dataframe(review, use_container_width=True, hide_index=True)
    except Exception as exc:
        st.error("Prediction failed. Check that the saved model is compatible with the installed scikit-learn version.")
        st.exception(exc)

st.markdown('<div class="footer-line">Used Vehicle Price Valuation · Machine-learning powered estimate</div>', unsafe_allow_html=True)
