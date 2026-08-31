# NM Associates CRM — Streamlit Dashboard

Interactive dashboard hai — filters, live charts, editable tables. Chalane ka tareeqa:

## 1. Install (ek dafa)
```
pip install -r requirements.txt
```

## 2. Run
Teeno files (`app.py`, `requirements.txt`, `NM_Associates_CRM_Dashboard.xlsx`) ek hi folder mein rakhein, phir:
```
streamlit run app.py
```

Browser mein khud khul jayega (`http://localhost:8501`). Agar khud na khule to yeh link manually open kar lein.

## Kya milta hai
- **Leads Dashboard tab**: Country / Status / Source filters (sidebar), live KPI cards, Pipeline donut chart, Country bar chart, Source bar chart, Country × Status heatmap, aur editable leads table.
- **Team Referrals tab**: Referrals-by-member bar chart, status breakdown, aur editable referral log (jahan aap rows add kar sakte hain).
- Har jagah **"Export to Excel"** button hai — jo changes aap dashboard mein karo ge, wapis usi `.xlsx` file mein save ho jayenge.

## Note
Agar `NM_Associates_CRM_Dashboard.xlsx` file present na ho to app khali template ke sath khul jayega — aap upar se apna data daal sakte hain.
