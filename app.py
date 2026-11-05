import io
import math
import os
import re
import pandas as pd
import streamlit as st

# Importi za generiranje PDF-ova i registraciju fonta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(
    page_title="Kontrola GLS Logističkih Računa", page_icon="📦", layout="wide"
)

# Sigurna registracija fontova s provjerom postojanja datoteka
FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

try:
    if os.path.exists("DejaVuSans.ttf") and os.path.exists("DejaVuSans-Bold.ttf"):
        pdfmetrics.registerFont(TTFont("DejaVuSans", "DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", "DejaVuSans-Bold.ttf"))
        FONT_REGULAR = "DejaVuSans"
        FONT_BOLD = "DejaVuSans-Bold"
    else:
        sys_fonts = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ]
        if os.path.exists(sys_fonts[0]) and os.path.exists(sys_fonts[1]):
            pdfmetrics.registerFont(TTFont("DejaVuSans", sys_fonts[0]))
            pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", sys_fonts[1]))
            FONT_REGULAR = "DejaVuSans"
            FONT_BOLD = "DejaVuSans-Bold"
except Exception as e:
    st.warning(f"Greška pri registraciji fonta: {e}")

st.title("📦 Sustav za Kontrolu i Analizu GLS Računa")
st.write(
    "Automatska kontrola troškova prijevoza, dodatka za gorivo i dodatnih usluga prema ugovoru Makromikro grupa d.o.o."
)

# Ugovorne osnovne cijene iz GLS cjenika 2026.
OSNOVNA_CIJENA_STANDARD = 3.66  # Business Parcel do 1 paketa/adresa (EUR)
OSNOVNA_CIJENA_SMALL = 4.38     # BusinessSmallParcel (EUR)


def izracunaj_osnovnu_cijenu_gls(masa, connected_services=""):
    """
    Određuje ugovorenu osnovnu cijenu prijevoza.
    Ako je primijenjen Paketomat/PaketShop popust (PLD/PSD/PSD/PSD...), 
    osnovna tarifa ovisi o ugovornim stavkama.
    """
    if masa < 2.0:
        return OSNOVNA_CIJENA_STANDARD
    else:
        return OSNOVNA_CIJENA_SMALL


def izracunaj_dodatak_goriva_gls(cijena_dizela):
    """
    Dodatak za gorivo prema GLS cjeniku:
    Iznad 1,00 €/l za svaku započetu promjenu cijene dizela od 0,03 €/l (kategoriju) 
    obračunava se +0,02 € po paketu.
    """
    if cijena_dizela <= 1.00:
        return 0.0
    kategorije = math.ceil((cijena_dizela - 1.00) / 0.03)
    return round(kategorije * 0.02, 2)


def ekstrahiraj_pbr_i_grad(adresa):
    """Izvlači poštanski broj i naziv grada iz stupca Delivery address (npr. 'HR-47240 Slunj ...')."""
    if not isinstance(adresa, str):
        return 10000, "Nepoznato"
    match = re.search(r"(?:HR-)?(\d{5})\s+([^,]+)", adresa)
    if match:
        pbr = int(match.group(1))
        grad = match.group(2).split()[0]
        return pbr, grad
    return 10000, "Nepoznato"


def izracunaj_radne_dane(datum_slanja, datum_dostave):
    try:
        d1 = pd.to_datetime(datum_slanja, errors="coerce")
        d2 = pd.to_datetime(datum_dostave, errors="coerce")
        if pd.isna(d1) or pd.isna(d2):
            return None
        radni_dani = pd.bdate_range(start=d1, end=d2).shape[0] - 1
        return max(0, radni_dani)
    except:
        return None


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Izvjestaj")
    return output.getvalue()


# Sidebar parametri
st.sidebar.header("Parametri obračuna GLS")
trenutna_cijena_goriva = st.sidebar.number_input(
    "Prosječna cijena dizela (€/l bez PDV-a):",
    min_value=0.80,
    max_value=2.50,
    value=1.35,
    step=0.01,
)
iznos_goriva_po_paketu = izracunaj_dodatak_goriva_gls(trenutna_cijena_goriva)
st.sidebar.info(
    f"Izračunati ugovoreni dodatak za gorivo: **{iznos_goriva_po_paketu:.2f} € po paketu**"
)

uploaded_file = st.file_uploader(
    "Učitaj GLS Excel ili CSV tablicu (npr. gls2.xlsx)", type=["xlsx", "csv"]
)

if uploaded_file is not None:
    if uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file)
    else:
        df = pd.read_csv(uploaded_file)

    st.success("GLS tablica uspješno učitana!")

    if st.button("Generiraj GLS izvještaje"):
        rezultati = []
        ukupno_pošiljaka = len(df)

        for idx, row in df.iterrows():
            parcel_num = str(row.get("Parcel number", f"PKT-{idx+1}"))
            masa = float(row.get("Weight / Size", 0.0))
            
            naplaceni_transport = float(row.get("Transport fee", 0.0))
            naplaceno_gorivo = float(row.get("Diesel fee", 0.0))
            naplaceni_servisi = float(row.get("Total service fee / Total service fees", 0.0))
            naplaceni_surcharges = float(row.get("Total surcharges fee", 0.0))
            sveukupno_naplaceno_faktura = float(row.get("Total amount", 0.0))
            
            connected_services = str(row.get("Connected Services", ""))
            orientation = str(row.get("Parcel orientation", "[DomesticDelivery]"))
            dest_country = str(row.get("Destination country", "HR"))
            
            adresa_dostave = str(row.get("Delivery address", ""))
            pbr, grad = ekstrahiraj_pbr_i_grad(adresa_dostave)

            d_slanja = row.get("Pick up date", None)
            d_dostave = row.get("Delivery date", None)
            tranzit_dani = izracunaj_radne_dane(d_slanja, d_dostave)
            
            # Za HR tranzit je u pravilu 1 radni dan (24h), za inozemstvo ovisi o državi
            dopušteni_rok = 1 if dest_country == "HR" else 3

            if tranzit_dani is not None:
                kasni = tranzit_dani > dopušteni_rok
                tranzit_status = (
                    f"{tranzit_dani} rad. dana (U roku)"
                    if not kasni
                    else f"🔴 {tranzit_dani} rad. dana (Kasni, rok je {dopušteni_rok})"
                )
            else:
                kasni = False
                tranzit_status = "Nema informacije"

            # Izračun ugovorenih očekivanih troškova
            ugovorena_osnova = izracunaj_osnovnu_cijenu_gls(masa, connected_services)
            ugovoreno_gorivo = iznos_goriva_po_paketu

            # Očekivano sveukupno (osnova + gorivo + servisi naplaćeni od GLS-a)
            očekivano_sveukupno = ugovorena_osnova + ugovoreno_gorivo + naplaceni_servisi + naplaceni_surcharges

            red_podataka = {
                "RedniBroj": idx + 1,
                "Parcel Number": parcel_num,
                "Kupac / Primatelj": row.get("Addressee name", row.get("Receiver's name", "")),
                "Grad": grad,
                "ZIP": pbr,
                "Adresa Dostave": adresa_dostave,
                "Destinacija": dest_country,
                "Tip Pošiljke": orientation,
                "Masa (kg)": masa,
                "Pick Up Date": d_slanja,
                "Delivery Date": d_dostave,
                "Stvarni Tranzit (dani)": tranzit_dani if tranzit_dani is not None else -1,
                "Dopušteni Rok (dani)": dopušteni_rok,
                "Status Dostave": tranzit_status,
                "Kasni": kasni,
                "Connected Services": connected_services,
                "Naplaćeni Transport (€)": round(naplaceni_transport, 2),
                "Ugovorena Osnova (€)": round(ugovorena_osnova, 2),
                "Naplaćeno Gorivo (€)": round(naplaceno_gorivo, 2),
                "Ugovoreno Gorivo (€)": round(ugovoreno_gorivo, 2),
                "Naplaćene Dodatne Usluge / Popusti (€)": round(naplaceni_servisi, 2),
                "Sveukupno Naplaćeno Faktura (€)": round(sveukupno_naplaceno_faktura, 2),
                "Sveukupno Očekivano (€)": round(očekivano_sveukupno, 2),
                "Razlika / Preplata (€)": round(sveukupno_naplaceno_faktura - očekivano_sveukupno, 2),
            }

            rezultati.append(red_podataka)

        res_df = pd.DataFrame(rezultati)

        valid_tranzit = res_df[res_df["Stvarni Tranzit (dani)"] >= 0]
        uk_validnih = len(valid_tranzit)
        if uk_validnih > 0:
            uk_kasni = int(valid_tranzit["Kasni"].sum())
            postotak_kasnjenja = (uk_kasni / uk_validnih) * 100.0
            postotak_urednih = 100.0 - postotak_kasnjenja
        else:
            uk_kasni = 0
            postotak_kasnjenja = 0.0
            postotak_urednih = 100.0

        uk_naplaceni_transport = res_df["Naplaćeni Transport (€)"].sum()
        uk_ugovorena_osnova = res_df["Ugovorena Osnova (€)"].sum()
        uk_naplaceno_gorivo = res_df["Naplaćeno Gorivo (€)"].sum()
        uk_ugovoreno_gorivo = res_df["Ugovoreno Gorivo (€)"].sum()
        uk_servisi = res_df["Naplaćene Dodatne Usluge / Popusti (€)"].sum()

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 1. Tranzit i rokovi",
            "⚖️ 2. Usporedba cijena",
            "🚨 3. Preplate",
            "📈 4. Zbirne sume",
            "🏙️ 5. Izvještaj po gradovima",
        ])

        with tab1:
            st.subheader("Analiza tranzita GLS pošiljaka")
            col_a, col_b, col_c = st.columns(3)
            col_a.metric(
                label="Uredno isporučeno u roku",
                value=f"{postotak_urednih:.1f}%",
                delta=f"{uk_validnih - uk_kasni} pošiljaka",
            )
            col_b.metric(
                label="Kašnjenje",
                value=f"{postotak_kasnjenja:.1f}%",
                delta=f"-{uk_kasni} pošiljaka",
                delta_color="inverse",
            )
            col_c.metric(label="Ukupno analizirano pošiljaka", value=f"{uk_validnih}")
            st.markdown("---")
            st.dataframe(
                res_df[[
                    "RedniBroj",
                    "Parcel Number",
                    "Kupac / Primatelj",
                    "Grad",
                    "ZIP",
                    "Pick Up Date",
                    "Delivery Date",
                    "Stvarni Tranzit (dani)",
                    "Dopušteni Rok (dani)",
                    "Status Dostave",
                ]],
                use_container_width=True,
            )

        with tab2:
            st.subheader("Detaljna usporedba svih stavki GLS fakture")
            st.dataframe(res_df, use_container_width=True)
            st.download_button(
                "📥 Preuzmi GLS Usporedbu (Excel)",
                to_excel(res_df),
                "gls_sve_usporedbe.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab3:
            st.subheader("Izdvojene preplate na GLS fakturi")
            sumnjive = res_df[res_df["Razlika / Preplata (€)"] > 0.05]
            if sumnjive.empty:
                st.success("Nema detektiranih preplata na ovoj fakturi!")
            else:
                st.warning(f"Pronađeno sumnjivih stavki: {len(sumnjive)}")
                st.dataframe(sumnjive, use_container_width=True)
                st.download_button(
                    "📥 Preuzmi Preplate (Excel)",
                    to_excel(sumnjive),
                    "gls_preplate.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        with tab4:
            st.subheader("📈 Zbirni financijski pregled GLS fakture (Bez PDV-a)")
            sveukupno_naplaceno_racun = res_df["Sveukupno Naplaćeno Faktura (€)"].sum()
            sveukupno_ocekivano_ugovor = res_df["Sveukupno Očekivano (€)"].sum()
            konačna_preplata = sveukupno_naplaceno_racun - sveukupno_ocekivano_ugovor

            col1, col2, col3 = st.columns(3)
            col1.metric(
                label="Ukupno naplaćeno na računu",
                value=f"{sveukupno_naplaceno_racun:,.2f} €",
            )
            col2.metric(
                label="Ukupno trebalo po ugovoru",
                value=f"{sveukupno_ocekivano_ugovor:,.2f} €",
            )
            col3.metric(
                label="Razlika / Povrat sredstava",
                value=f"{konačna_preplata:,.2f} €",
            )

            st.markdown("---")
            zbirni_detalji = pd.DataFrame([
                {
                    "Kategorija troška": "Transport (Osnovna cijena)",
                    "Što su naplatili (€)": round(uk_naplaceni_transport, 2),
                    "Što je trebalo biti (€)": round(uk_ugovorena_osnova, 2),
                },
                {
                    "Kategorija troška": f"Dodatak za gorivo ({iznos_goriva_po_paketu:.2f} €/paket)",
                    "Što su naplatili (€)": round(uk_naplaceno_gorivo, 2),
                    "Što je trebalo biti (€)": round(uk_ugovoreno_gorivo, 2),
                },
                {
                    "Kategorija troška": "Usluge i popusti (FDS, PLD Paketomat, COD...)",
                    "Što su naplatili (€)": round(uk_servisi, 2),
                    "Što je trebalo biti (€)": round(uk_servisi, 2),
                },
                {
                    "Kategorija troška": "SVEUKUPNO ZA GLS FAKTURU",
                    "Što su naplatili (€)": round(sveukupno_naplaceno_racun, 2),
                    "Što je trebalo biti (€)": round(sveukupno_ocekivano_ugovor, 2),
                },
            ])
            st.dataframe(zbirni_detalji, use_container_width=True)

        with tab5:
            st.subheader("🏙️ Zbirni izvještaj po gradovima i kupcima")
            sumarno_gradovi = (
                res_df.groupby(["Grad", "ZIP"])
                .agg(
                    Ukupno_Pošiljaka=("Parcel Number", "count"),
                    Ukupna_Masa_kg=("Masa (kg)", "sum"),
                    Ukupno_Naplaćeno_EUR=("Sveukupno Naplaćeno Faktura (€)", "sum"),
                )
                .reset_index()
                .sort_values(by="Ukupno_Pošiljaka", ascending=False)
            )
            st.dataframe(sumarno_gradovi, use_container_width=True)
            st.download_button(
                "📥 Preuzmi Izvještaj po Gradovima (Excel)",
                to_excel(sumarno_gradovi),
                "gls_izvjestaj_po_gradovima.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
