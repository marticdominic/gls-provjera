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

# Osnovne ugovorne cijene po cjeniku
OSNOVNA_CIJENA_PAKETA = 3.66  # Business Parcel do 1 paketa/adresa
OSNOVNA_CIJENA_SMALL = 4.38   # BusinessSmallParcel


def izracunaj_osnovnu_cijenu_gls(masa, je_small=False):
    """Izračun osnovne cijene prema ugovoru za GLS."""
    baza = OSNOVNA_CIJENA_SMALL if je_small or masa < 2.0 else OSNOVNA_CIJENA_PAKETA
    return baza


def izracunaj_dodatak_goriva_gls(cijena_dizela):
    """Izračun dodatka za gorivo: iznad 1.00 €/l, svakih 0.03 €/l je +0.02 € po paketu."""
    if cijena_dizela <= 1.00:
        return 0.0
    kategorije = math.ceil((cijena_dizela - 1.00) / 0.03)
    return round(kategorije * 0.02, 2)


def izracunaj_radne_dane(datum_slanja, datum_dostave):
    try:
        d1 = pd.to_datetime(datum_slanja, format="%d.%m.%Y", errors="coerce")
        d2 = pd.to_datetime(datum_dostave, format="%d.%m.%Y", errors="coerce")
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
    f"Izračunati dodatak za gorivo: **{iznos_goriva_po_paketu:.2f} € po paketu**"
)

uploaded_file = st.file_uploader(
    "Učitaj Excel ili CSV tablicu s GLS pošiljkama", type=["xlsx", "csv"]
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
        ukupno_paketa = (
            int(df["Number of Parcels"].sum())
            if "Number of Parcels" in df.columns
            else ukupno_pošiljaka
        )

        st.info(
            f"📊 Obrađeno pošiljaka: **{ukupno_pošiljaka}** | Ukupno paketa: **{ukupno_paketa}**"
        )

        usluge_lista = [
            "FlexDelivery",
            "Guaranteed24",
            "COD",
            "BankCard",
            "ExchangeService",
            "DocumentReturn",
            "AddresseeOnly",
            "Overweight",
            "Oversized",
        ]

        for idx, row in df.iterrows():
            pbr = row.get("Consignee ZIP Code", 10410)
            try:
                pbr_int = int(pbr)
            except:
                pbr_int = 10410

            masa = float(row.get("Weight", 1.0))
            naplaceni_transport = float(row.get("Transport Price", 0.0))
            naplaceno_gorivo = float(row.get("Fuel Surcharge", 0.0))

            d_slanja = row.get("Shipping Date", None)
            d_dostave = row.get("Delivery Time", None)
            tranzit_dani = izracunaj_radne_dane(d_slanja, d_dostave)
            dopušteni_rok = 1  # Domaći promet GLS-a je većinom 1 radni dan (24h)

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

            ugovorena_osnova = izracunaj_osnovnu_cijenu_gls(masa)
            ugovoreno_gorivo = iznos_goriva_po_paketu

            zbroj_naplacenih_dodatnih = 0.0
            postoji_dodatna_naplata = False

            red_podataka = {
                "RedniBroj": idx + 1,
                "Parcel Number": row.get("Parcel Number", row.get("Shipment ID", "")),
                "Consignee Name": row.get("Consignee Name", ""),
                "Consignee Town": row.get("Consignee Town", ""),
                "ZIP": pbr_int,
                "Masa (kg)": masa,
                "Slanje": d_slanja,
                "Dostava": d_dostave,
                "Stvarni Tranzit (dani)": (
                    tranzit_dani if tranzit_dani is not None else -1
                ),
                "Dopušteni Rok (dani)": dopušteni_rok,
                "Status Dostave": tranzit_status,
                "Kasni": kasni,
                "Naplaćeni Transport (€)": round(naplaceni_transport, 2),
                "Ugovorena Osnova (€)": round(ugovorena_osnova, 2),
                "Naplaćeno Gorivo (€)": round(naplaceno_gorivo, 2),
                "Ugovoreno Gorivo (€)": round(ugovoreno_gorivo, 2),
            }

            for usluga in usluge_lista:
                p_col = next(
                    (
                        c
                        for c in df.columns
                        if c.lower().replace(" ", "")
                        == f"price{usluga.lower().replace(' ', '')}"
                        or c.lower().replace(" ", "") == f"{usluga.lower()}"
                    ),
                    None,
                )
                p_iznos = (
                    float(row[p_col]) if p_col and pd.notna(row[p_col]) else 0.0
                )
                if p_iznos > 0:
                    postoji_dodatna_naplata = True
                    zbroj_naplacenih_dodatnih += p_iznos
                red_podataka[f"{usluga} - Naplaćeno (€)"] = round(p_iznos, 2)

            red_podataka["Naplaćene Dodatne Usluge Ukupno (€)"] = round(
                zbroj_naplacenih_dodatnih, 2
            )
            sveukupno_naplaceno = (
                naplaceni_transport
                + naplaceno_gorivo
                + zbroj_naplacenih_dodatnih
            )
            sveukupno_ocekivano = (
                ugovorena_osnova + ugovoreno_gorivo + zbroj_naplacenih_dodatnih
            )

            red_podataka["Sveukupno Naplaćeno (€)"] = round(sveukupno_naplaceno, 2)
            red_podataka["Sveukupno Očekivano (€)"] = round(sveukupno_ocekivano, 2)
            red_podataka["Ima Dodatnih Usluga"] = postoji_dodatna_naplata

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

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 1. Tranzit i rokovi",
            "⚖️ 2. Usporedba cijena",
            "🚨 3. Preplate",
            "📈 4. Zbirne sume",
            "🏙️ 5. Izvještaj po gradovima",
        ])

        with tab1:
            st.subheader("Analiza tranzita GLS pošiljaka (24h rok)")
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
            col_c.metric(label="Ukupno analizirano", value=f"{uk_validnih}")
            st.markdown("---")
            st.dataframe(
                res_df[[
                    "RedniBroj",
                    "Parcel Number",
                    "Consignee Name",
                    "Consignee Town",
                    "ZIP",
                    "Slanje",
                    "Dostava",
                    "Stvarni Tranzit (dani)",
                    "Status Dostave",
                ]],
                use_container_width=True,
            )

        with tab2:
            st.subheader("Detaljna usporedba cijena (Transport i Gorivo)")
            st.dataframe(res_df, use_container_width=True)
            st.download_button(
                "📥 Preuzmi GLS Usporedbu (Excel)",
                to_excel(res_df),
                "gls_sve_usporedbe.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab3:
            st.subheader("Izdvojene preplate na GLS fakturi")
            res_df["Razlika Fakture (€)"] = (
                res_df["Sveukupno Naplaćeno (€)"]
                - res_df["Sveukupno Očekivano (€)"]
            )
            sumnjive = res_df[res_df["Razlika Fakture (€)"] > 0.05]
            st.dataframe(sumnjive, use_container_width=True)
            st.download_button(
                "📥 Preuzmi Preplate (Excel)",
                to_excel(sumnjive),
                "gls_preplate.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab4:
            st.subheader("📈 Zbirni financijski pregled GLS fakture (Bez PDV-a)")
            sveukupno_naplaceno_racun = res_df["Sveukupno Naplaćeno (€)"].sum()
            sveukupno_ocekivano_ugovor = res_df["Sveukupno Očekivano (€)"].sum()
            konačna_preplata = (
                sveukupno_naplaceno_racun - sveukupno_ocekivano_ugovor
            )

            col1, col2, col3 = st.columns(3)
            col1.metric(
                label="Naplaćeno po računu",
                value=f"{sveukupno_naplaceno_racun:,.2f} €",
            )
            col2.metric(
                label="Trebalo po ugovoru",
                value=f"{sveukupno_ocekivano_ugovor:,.2f} €",
            )
            col3.metric(
                label="Ukupna preplata / Višak",
                value=f"{max(0, konačna_preplata):,.2f} €",
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
                    "Kategorija troška": "SVEUKUPNO ZA GLS FAKTURU",
                    "Što su naplatili (€)": round(
                        sveukupno_naplaceno_racun, 2
                    ),
                    "Što je trebalo biti (€)": round(
                        sveukupno_ocekivano_ugovor, 2
                    ),
                },
            ])
            st.dataframe(zbirni_detalji, use_container_width=True)

        with tab5:
            st.subheader("🏙️ Izvještaj po gradovima i poštanskim brojevima")
            if "Consignee Town" in res_df.columns:
                sumarno_gradovi = (
                    res_df.groupby(["Consignee Town", "ZIP"])
                    .agg(
                        Ukupno_Pošiljaka=("Parcel Number", "count"),
                        Ukupna_Masa_kg=("Masa (kg)", "sum"),
                        Ukupno_Naplaćeno_EUR=(
                            "Sveukupno Naplaćeno (€)",
                            "sum",
                        ),
                    )
                    .reset_index()
                )
                st.dataframe(sumarno_gradovi, use_container_width=True)
                st.download_button(
                    "📥 Preuzmi Izvještaj po Gradovima (Excel)",
                    to_excel(sumarno_gradovi),
                    "gls_izvjestaj_po_gradovima.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
