import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
import io
import zipfile
import os

# --- KONFIGURACJA (Stałe używane w kodzie) ---
CZCIONKA_REG = "DejaVuSans-Regular.ttf"
CZCIONKA_BOLD = "DejaVuSans-Bold.ttf"

# Dane Twojej firmy (Wykonawcy)
SPRZEDAWCA = {
    'nazwa': 'Twoja Firma Sp. z o.o.',
    'adres1': 'ul. Przykładowa 123',
    'adres2': '00-001 Warszawa',
    'nip': '1234567890',
    'email': 'biuro@twojafirma.pl',
    'bank': 'Bank Pekao S.A.',
    'konto': '00 0000 0000 0000 0000 0000 0000'
}

# --- FUNKCJE POMOCNICZE ---

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

# --- KLASA GENERATORA FAKTURY ---

class FakturaPDF(FPDF):
    def stworz_fakture(self, row):
        self.add_page()
        
        # Ładowanie czcionek z obsługą polskich znaków
        # Upewnij się, że pliki .ttf są w tym samym folderze na serwerze!
        self.add_font("DejaVu", style="", fname=CZCIONKA_REG)
        self.add_font("DejaVu", style="B", fname=CZCIONKA_BOLD)
        
        # Formatowanie daty - usuwamy godzinę
        try:
            data_faktury = pd.to_datetime(row['Data']).strftime('%Y-%m-%d')
        except:
            data_faktury = str(row['Data'])
        
        # 1. Tytuł
        self.set_y(20)
        self.set_font("DejaVu", style="B", size=18)
        self.cell(0, 10, txt=f"Faktura nr {row['Nr_faktury']}", align='C', new_x="LMARGIN", new_y="NEXT")
        
        # 2. Dane o wystawieniu (Prawa góra)
        curr_y = 35
        self.set_font("DejaVu", style="", size=9)
        self.set_xy(135, curr_y)
        self.cell(40, 5, txt="Data wystawienia:")
        self.set_xy(175, curr_y)
        self.cell(30, 5, txt=data_faktury)
        
        self.set_xy(135, curr_y + 5)
        self.cell(40, 5, txt="Miejsce wystawienia:")
        self.set_xy(175, curr_y + 5)
        self.cell(30, 5, txt="Warszawa")

        self.line(10, 50, 200, 50) # Linia oddzielająca nagłówek

        # 3. Sekcja Sprzedawca / Nabywca
        y_sekcja = 55
        # Sprzedawca
        self.set_xy(10, y_sekcja)
        self.set_font("DejaVu", style="", size=9)
        self.cell(90, 5, "Sprzedawca:", new_x="LMARGIN", new_y="NEXT")
        self.set_font("DejaVu", style="B", size=10)
        self.cell(90, 5, SPRZEDAWCA['nazwa'], new_x="LMARGIN", new_y="NEXT")
        self.set_font("DejaVu", style="", size=9)
        self.cell(90, 5, SPRZEDAWCA['adres1'], new_x="LMARGIN", new_y="NEXT")
        self.cell(90, 5, SPRZEDAWCA['adres2'], new_x="LMARGIN", new_y="NEXT")
        self.cell(90, 5, f"NIP: {SPRZEDAWCA['nip']}", new_x="LMARGIN", new_y="NEXT")

        # Nabywca
        self.set_xy(115, y_sekcja)
        self.set_font("DejaVu", style="", size=9)
        self.cell(90, 5, "Nabywca:", new_x="LMARGIN", new_y="NEXT")
        self.set_xy(115, self.get_y())
        self.set_font("DejaVu", style="B", size=10)
        self.cell(90, 5, str(row['Klient']), new_x="LMARGIN", new_y="NEXT")
        self.set_font("DejaVu", style="", size=9)
        
        adres_linii = str(row['Adres']).split(',')
        for linia in adres_linii:
            self.set_x(115)
            self.cell(90, 5, linia.strip(), new_x="LMARGIN", new_y="NEXT")

        # 4. Tabela Pozycji (Ręczne rysowanie, aby uniknąć "schodów")
        y_tab = 95
        kolumny = [10, 75, 12, 12, 25, 15, 41] # Szerokości kolumn
        headers = ["Lp", "Nazwa towaru / usługi", "Jm", "Ilość", "Cena Netto", "VAT", "Wartość Brutto"]
        
        self.set_font("DejaVu", style="B", size=8)
        self.set_fill_color(245, 245, 245)
        
        x_start = 10
        for i, h in enumerate(headers):
            self.set_xy(x_start, y_tab)
            self.rect(x_start, y_tab, kolumny[i], 10)
            self.cell(kolumny[i], 10, h, align='C', fill=True)
            x_start += kolumny[i]

        y_dane = y_tab + 10
        netto = float(row['Netto'])
        brutto = netto * 1.23
        
        wartosci = ["1", str(row['Produkt']), "szt.", "1", format_kwoty(netto), "23%", format_kwoty(brutto)]
        
        self.set_font("DejaVu", style="", size=9)
        x_start = 10
        for i, v in enumerate(wartosci):
            self.set_xy(x_start, y_dane)
            self.rect(x_start, y_dane, kolumny[i], 10)
            align = 'L' if i == 1 else 'C'
            self.cell(kolumny[i], 10, v, align=align)
            x_start += kolumny[i]

        # 5. Podsumowanie
        y_suma = y_dane + 20
        self.set_xy(120, y_suma)
        self.set_font("DejaVu", style="B", size=11)
        self.cell(35, 10, "DO ZAPŁATY:")
        self.cell(45, 10, f"{format_kwoty(brutto)} PLN", border=1, align='C', new_x="LMARGIN", new_y="NEXT")
        
        self.set_xy(10, y_suma + 12)
        self.set_font("DejaVu", style="", size=9)
        self.cell(0, 5, f"Słownie: {kwota_slownie(brutto)}", new_x="LMARGIN", new_y="NEXT")
        
        self.ln(5)
        self.set_font("DejaVu", style="B", size=9)
        self.cell(0, 5, "Dane do przelewu:", new_x="LMARGIN", new_y="NEXT")
        self.set_font("DejaVu", style="", size=9)
        self.cell(0, 5, f"Bank: {SPRZEDAWCA['bank']}", new_x="LMARGIN", new_y="NEXT")
        self.cell(0, 5, f"Nr konta: {SPRZEDAWCA['konto']}", new_x="LMARGIN", new_y="NEXT")

        # 6. SEKCJA PODPISÓW
        y_podpis = 250
        self.set_font("DejaVu", style="", size=8)
        
        # Podpis wystawcy
        self.set_xy(20, y_podpis)
        self.cell(60, 0, "", border="T") 
        self.set_xy(20, y_podpis + 2)
        self.cell(60, 5, "Osoba upoważniona do wystawienia", align='C')
        
        # Podpis odbiorcy
        self.set_xy(130, y_podpis)
        self.cell(60, 0, "", border="T") 
        self.set_xy(130, y_podpis + 2)
        self.cell(60, 5, "Osoba upoważniona do odbioru", align='C')

# --- INTERFEJS STREAMLIT ---

st.set_page_config(page_title="Generator Faktur", page_icon="📄")
st.title("📄 Profesjonalny Automat do Faktur")
st.write("Wgraj plik Excel, a ja przygotuję paczkę PDF-ów zgodną z Twoim wzorem.")

uploaded_file = st.file_uploader("Wybierz plik Excel (.xlsx)", type="xlsx")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Czyszczenie danych Netto
        df['Netto'] = df['Netto'].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
        df['Netto'] = pd.to_numeric(df['Netto'], errors='coerce')
        df = df.dropna(subset=['Netto', 'Nr_faktury']) # Usuwamy puste wiersze
        
        st.success(f"Wczytano pomyślnie {len(df)} faktur!")
        st.dataframe(df[['Nr_faktury', 'Klient', 'Netto']].head())

        if st.button("🚀 Generuj faktury i przygotuj ZIP"):
            # Sprawdzenie czy czcionki istnieją na serwerze
            if not os.path.exists(CZCIONKA_REG):
                st.error(f"Błąd: Nie znaleziono pliku czcionki {CZCIONKA_REG} w folderze aplikacji!")
            else:
                # Archiwum ZIP w pamięci RAM
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for index, row in df.iterrows():
                        pdf = FakturaPDF()
                        pdf.stworz_fakture(row)
                        
                        # Pobieramy bajty PDF (fpdf2 domyślnie zwraca bajty przy .output())
                        pdf_output = pdf.output()
                        
                        # Nazwa pliku (bezpieczna dla systemów)
                        filename = f"Faktura_{str(row['Nr_faktury']).replace('/', '_').replace('\\', '_')}.pdf"
                        zf.writestr(filename, pdf_output)
                
                st.download_button(
                    label="📥 Pobierz wszystkie faktury (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="faktury_wynik.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Wystąpił błąd podczas przetwarzania: {e}")
