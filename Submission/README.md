# Data Science Case Studies

Fundamentals of Data Science (22CSC59) – course-end project by Sarapu Rushil (160124733197), CSE-3.

One Jupyter Notebook containing two independent case studies:

| Part | Case study | Result |
|---|---|---|
| A | Titanic survival prediction using logistic regression | 84.4% accuracy, ROC-AUC 0.898; sex, class and age decide survival; deployed with Streamlit |
| B | Comparative analysis of classifiers for bank-marketing prediction | Random Forest best (ROC-AUC 0.801); calling its top 20% of customers reaches over 60% of subscribers |

## Project structure

| Path | Purpose |
|---|---|
| `Data_Science_Case_Studies.ipynb` | Combined notebook: shared setup, Part A (Titanic), Part B (Bank Marketing) |
| `data/` | `titanic.csv`, `bank-additional-full.csv` (+ attribute descriptions) |
| `figures/titanic/`, `figures/bank/` | All plots produced by the notebook; `figures/report/` holds the report diagrams |
| `model/` | Saved Titanic pipeline (`titanic_logreg.pkl`) and metrics of both parts |
| `app.py` | Streamlit app for Titanic survival predictions |
| `make_report.py` | Builds the combined Word report from `report_template.docx` |
| `export_pdf.ps1` | Refreshes page numbers in Word and saves the PDF (Windows + MS Word) |
| `Data_Science_Case_Studies_Report.docx` / `.pdf` | Combined course-end project report |

## How to run

```bash
pip install -r requirements.txt
jupyter notebook Data_Science_Case_Studies.ipynb    # Kernel -> Restart & Run All (about 30 seconds)
streamlit run app.py                                # Titanic prediction app
python make_report.py                               # rebuild the Word report
powershell -ExecutionPolicy Bypass -File export_pdf.ps1   # update page numbers + save PDF
```

All paths in the notebook are relative to this folder, so the whole folder can be moved or uploaded to GitHub as is.
