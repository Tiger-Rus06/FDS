"""Streamlit app for the Titanic logistic regression model.

Run locally with:  streamlit run app.py
"""
import json
from pathlib import Path

import joblib
import numpy as np
import pandas as pd
import streamlit as st

BASE = Path(__file__).parent
MODEL_PATH = BASE / "model" / "titanic_logreg.pkl"
METRICS_PATH = BASE / "model" / "titanic_metrics.json"
FEATURES = ["Age", "SibSp", "Parch", "Fare", "Pclass", "Sex", "Embarked"]

st.set_page_config(page_title="Titanic Survival Predictor", page_icon="🚢", layout="centered")


@st.cache_resource
def load_model():
    return joblib.load(MODEL_PATH)


@st.cache_data
def load_metrics():
    return json.loads(METRICS_PATH.read_text())


model = load_model()
info = load_metrics()

st.title("🚢 Titanic Survival Predictor")
st.write(
    "Enter a passenger's details to estimate their probability of surviving the Titanic disaster. "
    "Predictions come from a logistic regression model trained on 891 passengers."
)

# ---- User inputs
with st.sidebar:
    st.header("Passenger details")
    pclass = st.selectbox("Ticket class", [1, 2, 3],
                          format_func=lambda c: {1: "1st (Upper)", 2: "2nd (Middle)", 3: "3rd (Lower)"}[c])
    sex = st.radio("Sex", ["female", "male"], horizontal=True)
    age = st.slider("Age (years)", 0, 80, 28)
    sibsp = st.number_input("Siblings / spouses aboard", 0, 8, 0)
    parch = st.number_input("Parents / children aboard", 0, 6, 0)
    default_fare = {1: 60.0, 2: 15.0, 3: 8.0}[pclass]
    fare = st.number_input("Fare (£)", 0.0, 520.0, default_fare, step=1.0)
    embarked = st.selectbox("Port of embarkation", ["S", "C", "Q"],
                            format_func=lambda p: {"S": "Southampton", "C": "Cherbourg", "Q": "Queenstown"}[p])

passenger = pd.DataFrame([{
    "Age": age, "SibSp": sibsp, "Parch": parch, "Fare": fare,
    "Pclass": str(pclass), "Sex": sex, "Embarked": embarked,
}])[FEATURES]

# ---- Prediction
proba = float(model.predict_proba(passenger)[0, 1])
survived = proba >= 0.5

st.subheader("Prediction")
col1, col2 = st.columns(2)
col1.metric("Survival probability", f"{proba:.1%}")
col2.metric("Predicted outcome", "Survived ✅" if survived else "Did not survive ❌")
st.progress(proba)

if survived:
    st.success("The model predicts this passenger would **survive**.")
else:
    st.error("The model predicts this passenger would **not survive**.")

# ---- Explanation: contribution of each feature to the log-odds
with st.expander("Why this prediction?"):
    pre = model.named_steps["preprocessor"]
    clf = model.named_steps["classifier"]
    x = pre.transform(passenger)
    x = x.toarray() if hasattr(x, "toarray") else x
    names = [n.split("__")[1] for n in pre.get_feature_names_out()]
    contrib = pd.DataFrame({"feature": names, "contribution": x[0] * clf.coef_[0]})
    contrib = contrib[contrib["contribution"].abs() > 1e-6].sort_values("contribution")
    st.write(
        "Each bar shows how much a feature pushes the log-odds of survival up (positive) or down (negative) "
        "for this passenger, relative to the baseline of a 1st-class woman from Cherbourg with average age, fare and family size. "
        f"Intercept (baseline log-odds): **{clf.intercept_[0]:.2f}**."
    )
    st.bar_chart(contrib.set_index("feature"), horizontal=True)

# ---- Model info
with st.expander("Model performance (test set)"):
    m = info["metrics"]
    cols = st.columns(len(m))
    for c, (k, v) in zip(cols, m.items()):
        c.metric(k, f"{v:.3f}")
    st.caption(f"Trained on {info['n_train']} passengers, evaluated on {info['n_test']} held-out passengers.")
    coef = pd.DataFrame(info["coefficients"])[["feature", "coefficient", "odds_ratio"]]
    st.dataframe(coef, hide_index=True, use_container_width=True)
