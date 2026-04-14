import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
import io
import zipfile

# --- FUNKCJE POMOCNICZE (Te same co wcześniej) ---

def format_kwoty(kwota):
    return f"{float(kwota):.2f}".replace('.', ',')

def kwota_slownie(kwota):
    try:
        zlote = int(kwota)
        grosze = int(round((kwota - zlote) * 100))
        slownie_zlote = num2words(zlote, lang='pl')
        return f"{slownie_zlote} PLN {grosze}/100"
    except:
        return f"{kwota:.2f} PLN"

# --- KLASA GENERATORA (Lekka modyfikacja pod Streamlit) ---

class FakturaPDF(FPDF):
    def stworz_fakture(self, row, sprzedawca):
        self.add_page()
        # Ważne: Czcionki muszą być w tym samym folderze na serwerze!
        self.add_font("DejaVu", style="", fname="DejaVuSans-Regular.ttf")
        self.add_font("DejaVu", style="B", fname="DejaVuSans-Bold.ttf")
        
        # Logika rysowania (identyczna jak w Twoim poprzednim kodzie)
        self.set_y(20)
        self.set_font("DejaVu", style="B", size=18)
        self.cell(0, 10, txt=f"Faktura nr {row['Nr_faktury']}", ln=True, align='C')
        
        # ... (tutaj reszta kodu rysującego z poprzedniego etapu) ...
        # (Dla oszczędności miejsca pomijam powtórkę rysowania tabeli, 
        # ale wklej tam całą swoją funkcję stworz_fakture)
        
        self.set_font("DejaVu", style="", size=10)
        self.set_y(100)
        self.cell(0, 10, txt=f"Klient: {row['Klient']}", ln=True)
        self.cell(0, 10, txt=f"Produkt: {row['Produkt']}", ln=True)
        self.cell(0, 10, txt=f"Do zapłaty: {format_kwoty(row['Netto'] * 1.23)} PLN", ln=True)

# --- INTERFEJS STREAMLIT ---

st.set_page_config(page_title="Generator Faktur", page_icon="📄")
st.title("📄 Automat do Faktur")
st.write("Wgraj plik Excel, a ja przygotuję dla Ciebie paczkę PDF-ów.")

# 1. Upload pliku
uploaded_file = st.file_uploader("Wybierz plik Excel (.xlsx)", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Czyszczenie danych
        df['Netto'] = df['Netto'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
        df['Netto'] = pd.to_numeric(df['Netto'])
        
        st.success(f"Wczytano pomyślnie {len(df)} wierszy!")
        st.dataframe(df.head()) # Podgląd danych

        if st.button("🚀 Generuj faktury i przygotuj ZIP"):
            # Tworzymy archiwum ZIP w pamięci RAM
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for index, row in df.iterrows():
                    pdf = FakturaPDF()
                    # Tutaj przekazujemy dane sprzedawcy (możesz je pobrać z bocznego paska Streamlit!)
                    sprzedawca_demo = {'nazwa': 'Moja Firma'} 
                    pdf.stworz_fakture(row, sprzedawca_demo)
                    
                    # Zamiast zapisywać na dysk, pobieramy bajty PDF
                    pdf_output = pdf.output(dest='S') 
                    
                    # Dodajemy do ZIP
                    filename = f"Faktura_{str(row['Nr_faktury']).replace('/', '_')}.pdf"
                    zf.writestr(filename, pdf_output)
            
            # Przygotowanie przycisku pobierania
            st.download_button(
                label="📥 Pobierz wszystkie faktury (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="faktury_wynik.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Wystąpił błąd: {e}")