
# VPK/TVPP Cleaning Schedule → Act Generator

Automated system that converts monthly cleaning schedules (ODS/XLSX) into completed VPK/TVPP cleaning act Excel files.

The app reads the cleaning schedule, applies weekday **X** marks, calculates **Periodiškumas**, and generates formulas for **Kaina** and **Suma be PVM** using `TRUNC(..., 2)` (no rounding).

Supports multiple objects (Ignalina, Anykščiai, etc.) with fixed month grid ranges or automatic detection.

---

## 🚀 Features

- Upload **Aktas (.xlsx)** and **Grafikas (.ods/.xlsx)**
- Automatically:
  - Writes weekday **X** markings (Pn–Pn)
  - Calculates **Periodiškumas**
  - Inserts **TRUNC** pricing formulas
  - Updates **Suma be PVM**
- Supports:
  - Fixed monthly grid ranges (Sigitos nustatymai)
  - Autodetection (fallback)
- Works on **Streamlit Cloud** and locally
- Dark neon UI theme

---

## 📦 Project Structure
