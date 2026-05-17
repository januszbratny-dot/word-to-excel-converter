import xml.etree.ElementTree as ET
import streamlit as st

def parse_ksef(xml_file):
    # Obsługa przestrzeni nazw (namespaces) w KSeF
    tree = ET.parse(xml_file)
    root = tree.getroot()
    ns = {'ns': 'http://crd.gov.pl/wzor/2025/06/25/13775/'}

    # Wyciąganie danych
    try:
        nr_faktury = root.find('.//ns:Fa/ns:P_2', ns).text
        data_wyst = root.find('.//ns:Fa/ns:P_1', ns).text
        nip_sprzedawcy = root.find('.//ns:Podmiot1/ns:DaneIdentyfikacyjne/ns:NIP', ns).text
        nazwa_sprzedawcy = root.find('.//ns:Podmiot1/ns:DaneIdentyfikacyjne/ns:Nazwa', ns).text
        netto = root.find('.//ns:Fa/ns:P_13_1', ns).text
        vat = root.find('.//ns:Fa/ns:P_14_1', ns).text
        brutto = root.find('.//ns:Fa/ns:P_15', ns).text

        return {
            "nr": nr_faktury,
            "data": data_wyst,
            "nip": nip_sprzedawcy,
            "kontrahent": nazwa_sprzedawcy,
            "netto": netto,
            "vat": vat,
            "brutto": brutto
        }
    except AttributeError:
        return None

# Streamlit UI
st.title("KSeF ➔ Rachmistrz GT (EPP)")
files = st.file_uploader("Wgraj pliki XML", accept_multiple_files=True)

if files:
    results = []
    for f in files:
        data = parse_ksef(f)
        if data:
            results.append(data)
    
    st.table(results)
    
    if st.button("Generuj EPP"):
        # Tutaj musisz sformatować tekst zgodnie ze specyfikacją EDI++ (EPP)
        # Np. [ZAWARTOSC],1,1,2, ... itd.
        st.info("Generowanie pliku EPP (wymaga dopisania stałego nagłówka EPP)")
