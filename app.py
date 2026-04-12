# OBS! Kör den här filen i terminalen med kommandot:
#     streamlit run app.py

import streamlit as st
import requests
from openpyxl import load_workbook
from io import BytesIO

FILNAMN = "Allsvenskan_Tipstävling_2026.xlsx"

def hämta_och_uppdatera_excel():
    url = "https://allsvenskan.se/data-endpoint/statistics/standings/2026/total"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://allsvenskan.se/",
        "Origin": "https://allsvenskan.se"
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        api_data = response.json()
    except requests.RequestException as e:
        st.error(f"❌ Kunde inte hämta data från API: {e}")
        return None

    lagdata = {}
    for team_info in api_data.values():
        name = team_info.get("name")
        stats = {stat["name"]: stat["value"] for stat in team_info.get("stats", [])}
        if not name:
            continue
        S = stats.get("gp")
        V = stats.get("w")
        O = stats.get("t")
        F = stats.get("l")
        GM = stats.get("gf")
        IM = stats.get("ga")
        if None in (S, V, O, F, GM, IM):
            continue
        MS = GM - IM
        P = V * 3 + O
        lagdata[name] = {
            "S": S, "V": V, "O": O, "F": F,
            "GM": GM, "IM": IM, "MS": MS, "P": P
        }

    sorterade_lag = sorted(
        lagdata.items(),
        key=lambda x: (-x[1]["P"], -x[1]["MS"], -x[1]["GM"])
    )

    try:
        wb = load_workbook(FILNAMN)
        ws = wb["Allsvenskan"]
    except Exception as e:
        st.error(f"❌ Kunde inte öppna Excel-filen: {e}")
        return None

    for i, (lag_namn, stats) in enumerate(sorterade_lag[:16], start=2):
        ws.cell(row=i, column=1).value = i - 1
        ws.cell(row=i, column=2).value = lag_namn
        ws.cell(row=i, column=3).value = stats["S"]
        ws.cell(row=i, column=4).value = stats["V"]
        ws.cell(row=i, column=5).value = stats["O"]
        ws.cell(row=i, column=6).value = stats["F"]
        ws.cell(row=i, column=7).value = stats["GM"]
        ws.cell(row=i, column=8).value = stats["IM"]
        ws.cell(row=i, column=9).value = stats["MS"]
        ws.cell(row=i, column=10).value = stats["P"]

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

st.title("Allsvenskan Tipstävling 2026")

# === Uppdatera-knappen först
if st.button("🔄 Uppdatera Excel-fil med senaste data"):
    excel_stream = hämta_och_uppdatera_excel()
    if excel_stream:
        st.success("✅ Excel-filen är uppdaterad! Du kan ladda ner den nedan.")
        st.session_state["excel_stream"] = excel_stream
    else:
        st.error("❌ Kunde inte uppdatera Excel-filen.")

# === Info om knappens funktion om fil ej laddats ner ännu (flyttad hit)
if "excel_stream" not in st.session_state:
    st.info("Tryck på knappen ovan för att hämta senaste data och ladda ner Excel-filen.")

# === Ladda ner-knappen direkt efter (om fil finns)
if "excel_stream" in st.session_state:
    st.download_button(
        label="📥 Ladda ner uppdaterad Excel-fil",
        data=st.session_state["excel_stream"],
        file_name=FILNAMN,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# === Viktig info om redigering
st.markdown("### ℹ️ När du öppnar Excel-filen")
st.markdown("""
👉 **Klicka på `Aktivera redigering`** upptill i fönstret om du använder Excel.  
Det behövs för att formlerna och resultaten ska visas korrekt.
""")

# === Hjälp för de utan Excel
with st.expander("📂 Har du inte Excel installerat?"):
    st.markdown("""
Du kan fortfarande öppna och använda filen:

**🔹 Excel Online:**  
Gå till [office.com/launch/excel](https://office.com/launch/excel) – kräver Microsoft-konto.

**🔹 Google Kalkylark:**  
1. Gå till [Google Drive](https://drive.google.com)  
2. Ladda upp Excel-filen  
3. Högerklicka → **Öppna med → Google Kalkylark**  
4. Observera att vissa formler kan visas annorlunda än i Excel.
""")
