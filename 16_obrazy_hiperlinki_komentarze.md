Zanim napiszę moduł, zweryfikuję kluczowe elementy API z briefu — obrazy (jednostki, alt text, cykl życia pliku), komentarze, hiperlinki i zachowanie `copy_worksheet` względem nich.

Teraz zweryfikuję hiperlinki, `copy_worksheet`, walidację scalonych komórek z komentarzami oraz tryby `read_only`/`write_only`.

Weryfikuję ostatnie kwestie: ostrzeżenie o komentarzach w scalonych komórkach, tekst alternatywny, setter `cell.hyperlink` oraz znane błędy obrazów w changelogu.

Ostatnia partia weryfikacji: pełne źródło `Image` w 3.1.5 (cykl życia pliku i strumienia), lista wbudowanych stylów nazwanych oraz `add_image` w arkuszu.

Ostatnia weryfikacja — sposób zapisu hiperlinków (zewnętrzne vs wewnętrzne) oraz funkcje konwersji jednostek.

Ostatnia weryfikacja: dokumentacja stylów (czy `Hyperlink` jest wbudowanym stylem nazwanym) oraz zmiany w 3.1 dotyczące stylów nazwanych.

Jedna weryfikacja techniczna przed pisaniem — klasy pomocnicze anchora, których użyję w kodzie.

Zweryfikowałem wszystko, co potrzebne. Kluczowe ustalenia, które **prostują lub uzupełniają brief**:

- `img.width`/`img.height` to **piksele** (z `PIL.Image.size`) — potwierdzone w `_check_anchor()`: obrazy → `pixels_to_EMU`, wykresy → `cm_to_EMU`. Nie ma `units.cm_to_pixels()`.
- **Hiperlinki wewnętrzne nie działają przez `"#Arkusz!A1"`.** `Hyperlink.__attrs__` zawiera `location`, a `target` staje się relacją `TargetMode="External"`. Wewnętrzny link = `Hyperlink(location=...)`, nie prefiks `#`.
- **`cell.style = "Hyperlink"` jest poprawne** — `'Hyperlink'` to wbudowany styl nazwany (sekcja „Text styles" w dokumentacji stylów).
- **Obrazy w `write_only` DZIAŁAJĄ** (`WriteOnlyWorksheet.add_image = Worksheet.add_image.__get__`). Niedostępne są w `read_only` i na `Chartsheet` (`AttributeError`).
- **Tekst alternatywny: openpyxl wpisuje dosłownie `descr="Picture"` dla każdego obrazu** (`_picture_frame`) i nie ma do tego API. `img.anchor` to pozycja, nie opis — brief myli te rzeczy.
- `set_categories`… nie dotyczy tego modułu. Natomiast: **komentarz na scalonej pod-komórce rzuca `AttributeError`** już przy próbie przypisania.

---

# Moduł 16 — Obrazy, hiperlinki i komentarze: elementy, które nie są komórkami

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–15

## 0. W tym module nauczysz się

- **Zrozumiesz, dlaczego obrazy, hiperlinki i komentarze mają „inne zasady"** niż wszystko, co robiłeś dotąd. Poznasz wspólną cechę tej trójki: **żadne z nich nie jest wartością w komórce** — więc nie da się ich odczytać przez `ws["A1"].value`, nie mają typu Excela i nie zachowują się jak dane.
- **Opanujesz wstawianie obrazów** (`ws.add_image`), poznasz **trzy jednostki miary**, którymi openpyxl się posługuje — **piksel, punkt i centymetr** — i dowiesz się, dlaczego to jedno z najczęstszych źródeł „obraz jest za duży / za mały".
- **Zrozumiesz cykl życia pliku obrazu i strumienia**: dlaczego `_data()` **zamyka** strumień, dlaczego ten sam obiekt `Image` może zawieść przy drugim zapisie, i jaka praktyka jest odporna na wszystkie te pułapki.
- **Wygenerujesz obraz w locie** (`matplotlib` → PNG w pamięci → `add_image`) — wzorzec, który **omija wszystkie ograniczenia wykresów** z modułu 15 (mapa, treemap, histogram, waterfall: narysuj w `matplotlib` i wstaw jako obraz).
- **Nauczysz się trzech rodzajów hiperlinków** — zewnętrznego, wewnętrznego i `mailto`/pliku — oraz zrozumiesz model `Hyperlink` z jego rozdzieleniem na `target` (relacja zewnętrzna) i `location` (nawigacja w skoroszycie). To tu poprawimy pewien powszechny błąd.
- **Poznasz komentarze** (notatki), ich rozmiary w pikselach, automatyczne kopiowanie przy wielokrotnym przypisaniu — i **ostrzeżenie, które openpyxl wypisuje dla komentarzy w scalonych komórkach**.
- **Dowie się wprost, czego nie ma:** tekstu alternatywnego (openpyxl wpisuje dosłownie `descr="Picture"`), obrazów i komentarzy w `read_only`, obrazów na `Chartsheet`, obrazów i wykresów w `copy_worksheet`.

## 1. Intuicja i analogia

### 1.1. Trójka, której nie znajdziesz w komórce

Wszystko, co robiłeś od modułu 04, miało jedną wspólną własność: **było wartością w komórce**. Liczba, tekst, data, formuła, kolor, obramowanie, format liczby, reguła warunkowa — za każdym razem mogłeś zapytać „co jest w `B3`?" i dostać odpowiedź (albo obiekt, który tę odpowiedź niesie).

Obraz, hiperlink i komentarz wypadają z tego schematu. Spójrz:

```python
ws["B2"] = "Logo:"
img = Image("logo.png")
ws.add_image(img, "B2")        # obraz "przy" komórce B2
print(ws["B2"].value)          # → 'Logo:'   ← obraz NIE jest wartością B2
```

Pytasz `ws["B2"].value` i widzisz tylko tekst. Gdzie jest obraz? Nie „w" `B2`. Jest **nałożony** na `B2` — przypięty do siatki, ale nie będący jej częścią.

Wyobraź sobie **tablicę korkową w biurze**:

- **Siatka korkowa z liniami** = komórki arkusza. Każde pole ma przypiętą kartkę z jakąś treścią (albo jest puste).
- **Karteczka na pinezce** = obraz. Leży **na** tablicy, zasłania kilka pól, ale nie znalazła się w żadnym z nich. Możesz ją przesunąć, a pola pod spodem nie zmienią swojej treści.
- **Kartka z wypisanym adresem** = hiperlink. To też treść pola, ale zawiera **wskazówkę, gdzie pójść**, a nie sam cel.
- **Żółta karteczka przyklejona z boku** = komentarz. Nie jest częścią wypowiedzi, jest **przypisem** do niej.

Wszystkie trzy są **nakładkami** albo **przypisami**. Żadna nie jest „danymi" w rozumieniu Excela. To nie jest drobiazg techniczny — to jest właśnie powód, dla którego w tym module pojawiają się ograniczenia, których nie widziałeś nigdzie wcześniej.

### 1.2. Trzy linijki do jednego rysunku

W module 15 poznałeś obraz **dwóch linijek przy jednym rysunku** (dwie osie wykresu — sprzedaż i marża na różnych skalach). Teraz wraca to samo, ale w innej postaci.

Wyobraź sobie, że dostajesz od architekta **rysunek drogi** i chcesz opisać na nim wszystko:

- **grubość warstwy asfaltu** mierzysz w **centymetrach**,
- **szerokość jezdni** też w **centymetrach**, ale w innej skali,
- **rozdzielczość skanowania rysunku** w **pikselach na cal**.

Jeden dokument, **trzy różne miary**. Jeśli pomieszasz je w głowie, dostaniesz absurd: „asfalt o grubości 2000" (to piksele!) albo „jezdnia 7,5" (to centymetry, ale wygląda jak piksele).

W openpyxl jest dokładnie tak samo, i to jest **największa pułapka tego modułu**:

| Element | Jednostka openpyxl | Dlaczego |
|---|---|---|
| Wykres | **centymetry** (`chart.width`, `chart.height`) | moduł 15 |
| Obraz | **piksele** (`img.width`, `img.height`) | naturalny rozmiar pliku |
| Szerokość kolumny | **znaki** | spadek po starym Excelu |
| Wysokość wiersza | **punkty** | typografia |
| Komentarz | **piksele** | rozmiar okienka |
| Kotwica (wewnętrznie) | **EMU** | specyfikacja OOXML |

I jeszcze: **polecenie w OOXML mierzy się w EMU** (*English Metric Unit*), gdzie 1 cal = 914 400 EMU, 1 cm = 360 000 EMU, a **1 piksel = 9525 EMU** (przy 96 DPI — to rozdzielczość, której domyślnie używa Excel dla bitmap).

Wszystko to zobaczysz w źródle openpyxl w jednej funkcji — i ta funkcja jest kluczem do całego modułu:

```python
# openpyxl/drawing/spreadsheet_drawing.py
elif isinstance(obj, Image):
    anchor.ext.width = pixels_to_EMU(obj.width)     # obraz: PIKSELE
# openpyxl/chart/_chart.py -> ChartBase
    anchor.ext.width = cm_to_EMU(obj.width)         # wykres: CENTYMETRY
```

**Zapamiętaj to rozróżnienie raz na zawsze.** Obraz w openpyxl mierzy się w pikselach, wykres w centymetrach. Jeśli będziesz to mylić, obrazy będą albo mikroskopijne, albo gigantyczne.

**Konsekwencja praktyczna: nie ma `units.cm_to_pixels()`.** Sprawdziłem cały moduł `openpyxl.utils.units` — są tam `pixels_to_EMU`, `EMU_to_pixels`, `cm_to_EMU`, `EMU_to_cm`, `inch_to_EMU`, `EMU_to_inch`, `pixels_to_points`, `points_to_pixels`, `dxa_to_cm`, `cm_to_dxa`, `degrees_to_angle` — ale **nie ma konwersji cm ↔ px**. Napiszesz ją sam, w jednej linii ($1$ cal $= 2{,}54$ cm, $96$ px na cal):

$$
\text{px} = \frac{\text{cm}}{2{,}54} \times 96
$$

### 1.3. Pinezka, nie koperta

W module 15 powiedziałem: wykres to kartka przypięta pinezką nad siatką. Obraz działa **identycznie**, z jednym wzmocnieniem — bo obraz jest *dosłownie* bitmapą, więc intuicja „kartki na tablicy" jest tu jeszcze bardziej dosłowna.

**Pinezka** to w terminologii openpyxl **anchor** (kotwica). Domyślnie obraz ma kotwicę w `A1`:

```python
# fragment klasy Image w openpyxl/drawing/image.py
class Image:
    _id = 1
    _path = "/xl/media/image{0}.{1}"
    anchor = "A1"          # ← DOMYŚLNA POZYCJA. Zawsze A1, jeśli nie podasz innej!
```

To wyjaśnia, dlaczego początkujący piszą „obraz zasłonił mi dane" — bo `ws.add_image(img)` **bez drugiego argumentu** wstawia obraz w `A1`, na same dane.

Pinezka przytrzymuje **lewy górny róg** obrazu. Popatrz na cztery konsekwencje, które zbijają ludzi z tropu:

1. **Sortowanie i filtrowanie nie przenoszą obrazu.** Obraz jest przyczepiony do **kotwicy** (komórki), nie do **danych**. Jeśli w komórce-kotwicy są dane i posortujesz kolumnę, wartość pojedzie na inne miejsce, a obraz zostanie. Tablica się przekręciła, kartka została.
2. **Scalanie komórek psuje pozycję.** Scalenie zmienia geometrię siatki, więc obraz „przeskakuje" albo ląduje w dziwnym miejscu. (Scalanie a wartości — moduł 11.)
3. **Wstawienie wiersza/w kolumny nad obrazem przesuwa go razem z komórkami** — bo kotwica też się przesunęła. To akurat zachowanie pożądane.
4. **Obraz nie jest częścią danych, więc nie da się go „odczytać"** tak jak wartości. Jedynym sposobem jest zajrzenie do listy `ws._images` (atrybut prywatny, sekcja 2.13).

### 1.4. Hiperlink to drzwi, komentarz to przypis

Dwa pozostałe elementy są prostsze, ale mają po jednym zaskoczeniu.

**Hiperlink to drzwi w ścianie.** W komórce jest napis („Otwórz raport"), a pod napisem — **nie widoczny dla ciebie** — adres, dokąd prowadzą. Kluczowe: adres i napis to **dwie różne rzeczy**. Możesz zmienić napis, nie ruszając drzwi. I możesz mieć drzwi bez napisu (pusto, głupio) albo napis bez drzwi (tekst wyglądający jak link, ale nieklikalny).

To rozróżnienie wyjaśnia najczęstszą pomyłkę: **wpisanie adresu jako tekstu to nie hiperlink.**

```python
ws["A1"] = "https://example.com"                    # ← zwykły TEKST. Nie klika się.
ws["A2"] = "example.com"
ws["A2"].hyperlink = "https://example.com"          # ← DRZWI. Klikalne.
```

W pliku różnica jest dramatyczna: pierwszy przypadek to napis w `sheet1.xml`. Drugi to napis **plus osobny plik relacji** (`xl/worksheets/_rels/sheet1.xml.rels`) z wpisem `<Relationship Target="https://example.com" TargetMode="External"/>`. Pamiętasz moduł 02? To ten sam mechanizm, który łączył arkusz z rysunkiem.

**Komentarz to przypis na marginesie** (w Excelu: „notatka"). Nie wpływa na wartość komórki, nie jest widoczny, dopóki ktoś nie najedzie myszą. Ma **tekst** i **autora** — i oba trzeba podać.

I tu zaskoczenie, które pokazuje całą różnicę między „obiektem" a „treścią":

```python
c = Comment("Sprawdź z księgowością", "Anna")
ws["B2"].comment = c
ws["C3"].comment = c              # ta sama zmienna!
print(ws["B2"].comment is c)      # True
print(ws["C3"].comment is c)      # False!  ← openpyxl zrobił KOPIĘ
```

Dlaczego? Bo openpyxl dba o to, żeby **jeden komentarz nie był jednocześnie w dwóch miejscach** (w pliku komentarz należy do konkretnej komórki, `ref`). To ta sama filozofia co przy stylach w module 09: **obiekt jest współdzielony, ale gdy coś ma należeć do konkretnego miejsca, powstaje kopia.**

### 1.5. Trzy kategorie prawdy — w wersji „nakładki"

W module 03 obiecałem tabelę „potrafi / potrafi częściowo / gubi przy zapisie". Oto ona w wydaniu dla tego modułu:

| Element | openpyxl **potrafi** | **potrafi częściowo** | **gubi przy zapisie** |
|---|---|---|---|
| **Obrazy** | dodać PNG/JPEG/GIF, skalować, zakotwiczyć | odczytać z pliku, odtworzyć pozycję | **tekst alternatywny** (`descr="Picture"`), WMF, obrazy spoza pakietu, obrazy gdy brak `pillow` |
| **Hiperlinki** | zewnętrzne, wewnętrzne, `mailto`, pliki, tooltip, styl | — | — (ta kategoria jest solidna) |
| **Komentarze** | tekst, autor, rozmiar okienka | odczytać tekst i autora | **formatowanie tekstu**, rozmiar okna przy odczycie, **komentarze w scalonych komórkach**, komentarze w `read_only` |

Zwróć uwagę na asymetrię. **Hiperlinki są solidne** — jedyna naprawdę „bezpieczna" z tej trójki. **Obrazy i komentarze tracą**. To nie jest przypadek: hiperlink to dwie linijki XML, a obraz i komentarz to całe osobne części pakietu z własnymi rysunkami VML.

## 2. Teoria

### 2.1. Gdzie to mieszka w pliku (powrót do modułu 02)

Zanim wejdziemy w API, zajrzyjmy do środka `.xlsx`. Ten rzut oka uzasadni **każde** ograniczenie tego modułu.

Wygeneruj skoroszyt z obrazem, hiperlinkiem i komentarzem (Przykład 1 i 3), zmień rozszerzenie na `.zip`, rozpakuj:

```text
raport.xlsx (rozpakowany)
├── [Content_Types].xml            ← deklaruje typy: png, vml, comments
├── docProps/
│   └── core.xml                   ← autor dokumentu, data (moduł 04)
└── xl/
    ├── workbook.xml
    ├── sheets/sheet1.xml          ← DANE + <hyperlinks> (dla linków WEWNĘTRZNYCH)
    ├── media/
    │   └── image1.png             ← SAM PLIK OBRAZU (kopia bitów!)
    ├── drawings/
    │   ├── drawing1.xml           ← POZYCJA obrazu (kotwice, sekcja 2.4)
    │   └── _rels/drawing1.xml.rels← powiązanie rysunku → plik w media/
    ├── comments1.xml              ← TREŚĆ KOMENTARZY (tekst, autor)
    ├── vmlDrawing1.vml            ← WYGLĄD I POZYCJA OKIENKA komentarza (VML!)
    ├── worksheets/_rels/
    │   └── sheet1.xml.rels        ← relacje: linki ZEWNĘTRZNE, rysunek, komentarze
    └── styles.xml
```

**Cztery rzeczy, które trzeba z tego wyczytać:**

**1. Obraz jest kopiowany, nie wskazywany.** `xl/media/image1.png` to **fizyczna kopia bitów** Twojego pliku, wciśnięta do archiwum ZIP. To fundamentalna różnica wobec wykresu (moduł 15): wykres trzymał **adres** (`'Sprzedaz'!$B$2`), obraz trzyma **kopię bajtów**. Dlatego obraz w pliku `.xlsx` zwiększa jego rozmiar, a wykres nie. I dlatego **wykres reaguje na zmianę komórki, a obraz nie** — obraz jest zamrożoną migawką.

**2. Komentarz ma dwie części w dwóch różnych technologiach.** Treść siedzi w `comments1.xml` (nowoczesny XML), a **wygląd i pozycja okienka** — w `vmlDrawing1.vml`, czyli w **VML** (*Vector Markup Language*), technologii sprzed ćwierćwiecza. To dlatego openpyxl obsługuje z komentarza „tekst i autora", a rozmiar okienka odtwarza tylko częściowo: druga połowa komentarza żyje w starym, pokręconym formacie.

**3. Kotwica obrazu jest w osobnym pliku.** `drawing1.xml` opisuje **gdzie** obraz leży, a `media/image1.png` — **co** to jest. Dwa pliki, dwa zadania. Stąd trudność z „przesuwaniem" i stąd przewidywalność pozycji (można ją czytać i pisać).

**4. Hiperlink zewnętrzny jest relacją, wewnętrzny — atrybutem.** Zapamiętaj to zdanie, wrócimy do niego w 2.8 — to sedno poprawki do najczęstszego błędu w internecie.

### 2.2. Obraz: `Image()` i warunek istnienia `pillow`

Klasa `Image` **nie jest częścią openpyxl** — jest cienką nakładką na `PIL.Image` z biblioteki **Pillow**. Pełne źródło z 3.1.5:

```python
# openpyxl/drawing/image.py — CAŁA klasa, nic nie pomijam
from io import BytesIO

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = False

def _import_image(img):
    if not PILImage:
        raise ImportError('You must install Pillow to fetch image objects')
    if not isinstance(img, PILImage.Image):
        img = PILImage.open(img)
    return img

class Image:
    """Image in a spreadsheet"""

    _id = 1
    _path = "/xl/media/image{0}.{1}"
    anchor = "A1"

    def __init__(self, img):
        self.ref = img
        mark_to_close = isinstance(img, str)
        image = _import_image(img)
        self.width, self.height = image.size
        try:
            self.format = image.format.lower()
        except AttributeError:
            self.format = "png"
        if mark_to_close:
            # PIL instances created for metadata should be closed.
            image.close()

    def _data(self):
        """Return image data, convert to supported types if necessary"""
        img = _import_image(self.ref)
        # don't convert these file formats
        if self.format in ['gif', 'jpeg', 'png']:
            img.fp.seek(0)
            fp = img.fp
        else:
            fp = BytesIO()
            img.save(fp, format="png")
            fp.seek(0)
        data = fp.read()
        fp.close()
        return data

    @property
    def path(self):
        return self._path.format(self._id, self.format)
```

Przeanalizujmy to zdanie po zdaniu, bo **cała wiedza o obrazach w openpyxl mieści się w tych 40 linijkach.**

**`if not PILImage: raise ImportError('You must install Pillow to fetch image objects')`** — jeśli nie masz Pillow, każdy obraz wywoła ten komunikat. To ten sam komunikat, który w czytniku powoduje **ciche pominięcie obrazów z wczytanego pliku** (sekcja 2.11).

**`self.ref = img`** — *referencja* do źródła. Albo ścieżka (`str`), albo obiekt `PIL.Image`, albo `BytesIO`. To ona zostanie użyta **przy zapisie**, nie teraz.

**`mark_to_close = isinstance(img, str)`** — jeśli podałeś **ścieżkę do pliku**, openpyxl otworzy go tylko po to, żeby poznać rozmiar i format, a potem **zamknie uchwyt**. Jeśli podałeś **strumień lub obiekt PIL**, nie zamknie — bo nie jest właścicielem tego zasobu.

**`self.width, self.height = image.size`** — **piksele**, prosto z Pillow. Zapamiętaj: to są *oryginalne* wymiary pliku.

**`self.format = image.format.lower()`** — rozszerzenie wynikające z samego pliku (`png`, `jpeg`, `gif`...). Gdy Pillow nie potrafi podać formatu (np. dla obrazu wygenerowanego w pamięci), openpyxl przyjmuje `"png"`.

**`_data()` — i tu jest najważniejszy szczegół całego modułu:**

- openpyxl **otwiera źródło ponownie** (`_import_image(self.ref)`),
- dla `png`/`jpeg`/`gif` **czyta bajty bezpośrednio** (bez konwersji),
- dla innych formatów (np. BMP, TIFF) **konwertuje do PNG** w pamięci — to dlatego „obraz nie jest prostokątem" nie jest problemem, ale też dlatego tracisz format źródłowy,
- i na końcu **`fp.close()`** — zamyka odczytany strumień.

**Skąd `ImportError` w praktyce:** `pip install pillow`. Pillow **nie jest** zależnością obowiązkową openpyxl (moduł 00), więc na czystym środowisku `Image(...)` się wywali. To jedna z rzeczy, o których warto pamiętać przy budowaniu obrazu Dockera (moduł 19).

### 2.3. Jednostki: kompletna mapa i dwie funkcje, które napiszesz sam

Zbierzmy wszystko w jednym miejscu. To tabela, do której będziesz wracać.

| Jednostka | Symbol w kodzie | Relacja | Używana do |
|---|---|---|---|
| piksel | px | $1$ px $= 9525$ EMU | `img.width`, `img.height`, `comment.width/height` |
| punkt | pt | $1$ cal $= 72$ pt | `row_dimensions[h].height`, `Font(size=)`, marginesy |
| centymetr | cm | $1$ cm $= 360\,000$ EMU | `chart.width`, `chart.height` (moduł 15) |
| cal | inch | $1$ in $= 914\,400$ EMU | `page_margins` (moduł 11) |
| dxa (twip) | dxa | $\text{dxa} = \text{pt} \times 20$ | niskopoziomowe szerokości |
| EMU | — | jednostka wewnętrzna OOXML | wszystko, pod spodem |

Funkcje dostępne w `openpyxl.utils.units` (wszystkie sprawdzone w źródle 3.1.5):

```python
pixels_to_EMU(value)      # 1 px = 9525 EMU
EMU_to_pixels(value)      # round(value / 9525)
cm_to_EMU(value)          # 1 cm = 360000 EMU
EMU_to_cm(value)          # round(value / 360000, 4)
inch_to_EMU(value)        # 1 in = 914400 EMU
EMU_to_inch(value)        # round(value / 914400, 4)
inch_to_dxa(value)        # 1 in = 1440 dxa
dxa_to_inch(value)
dxa_to_cm(value)
cm_to_dxa(value)
pixels_to_points(value, dpi=96)     # px → pt (typografia!)
points_to_pixels(value, dpi=96)     # pt → px  (okienka komentarzy!)
degrees_to_angle(value)             # → 1/60000 stopnia
angle_to_degrees(value)
short_color(color)                  # usuwa kanal alfa z koloru
```

**Czego brakuje: `cm_to_pixels` i `pixels_to_cm`.** Sprawdziłem cały moduł — ich tam nie ma. Napisz je sam, raz, w jednym miejscu projektu:

```python
# units_extra.py — male uzupelnienie, ktorego openpyxl nie ma
DPI_EXCEL = 96          # rozdzielczosc, ktorej domyslnie uzywa Excel dla bitmap

def cm_to_pixels(cm: float, dpi: int = DPI_EXCEL) -> int:
    """Centymetry -> piksele (jednostka img.width/img.height)."""
    return round(cm / 2.54 * dpi)

def pixels_to_cm(px: float, dpi: int = DPI_EXCEL) -> float:
    """Piksele -> centymetry (przydatne przy planowaniu ukladu)."""
    return round(px * 2.54 / dpi, 2)
```

Sprawdźmy na liczbach: obraz szeroki na $800$ px to $800 \times 2{,}54 / 96 = 21{,}17$ cm. Jeśli chcesz, żeby obraz miał dokładnie $12$ cm szerokości: $12 / 2{,}54 \times 96 = 453$ px. Wartości zaokrąglaj do liczb całkowitych — `img.width` jest wpisywane do XML jako `int` (przez `pixels_to_EMU`, które robi `int(value * 9525)`).

> **Uwaga o proporcjach.** `Image` **nie ma** `resize_proportional` (to atrybut klasy `Drawing`, która obsługuje kształty — nie `Image`). Zmieniając `img.width` sam, **rozciągniesz obraz**. Musisz policzyć skalę samodzielnie:
>
> ```python
> skala = nowa_szerokosc / img.width
> img.height = round(img.height * skala)
> img.width  = nowa_szerokosc
> ```

### 2.4. Anchor: pinezka w szczegółach

**Domyślna pozycja to `A1`** (atrybut klasowy `anchor = "A1"`). Możesz ją podać na dwa sposoby:

```python
img = Image("logo.png")
ws.add_image(img, "D5")        # sposob 1: drugi argument = kotwica
# albo:
img.anchor = "D5"
ws.add_image(img)              # sposob 2: atrybut obiektu
```

W źródle `add_image` to zaledwie trzy linijki:

```python
def add_image(self, img, anchor=None):
    """Add an image to the sheet. Optionally provide a cell for the top-left anchor"""
    if anchor is not None:
        img.anchor = anchor
    self._images.append(img)
```

Zwróć uwagę: **`add_image` nic nie liczy, tylko dopisuje do listy.** Cała geometria powstaje dopiero przy **zapisie**, w funkcji `_check_anchor` (widzieliśmy ją w 1.2). To ta sama filozofia „obiekt żywy, serializacja na końcu" z modułu 03.

**Trzy rodzaje kotwic** (z `openpyxl.drawing.spreadsheet_drawing`):

| Klasa | `tagname` | Co opisuje | Zachowanie w Excelu |
|---|---|---|---|
| `OneCellAnchor` | `oneCellAnchor` | lewy górny róg + **stały rozmiar w EMU** | obraz nie rośnie przy zmianie rozmiaru komórek |
| `TwoCellAnchor` | `twoCellAnchor` | lewy górny róg + **prawy dolny róg** (dwie komórki) | obraz **rozciąga się** z komórkami |
| `AbsoluteAnchor` | `absoluteAnchor` | pozycja i rozmiar w EMU, bez komórek | obraz „wisi" niezależnie od siatki |

**Domyślnie openpyxl tworzy `OneCellAnchor`** — a gdy przekazujesz tekst („B2"), robi to w tej gałęzi:

```python
anchor = obj.anchor
if not isinstance(anchor, _AnchorBase):
    row, col = coordinate_to_tuple(anchor.upper())
    anchor = OneCellAnchor()
    anchor._from.row = row - 1        # ← MINUS JEDEN!
    anchor._from.col = col - 1        # ← MINUS JEDEN!
```

> **⚠️ Pułapka odwrotnego indeksowania.** Wewnątrz kotwicy indeksy są **liczone od zera** (`row - 1`, `col - 1`), podczas gdy **wszędzie indziej w openpyxl od jedynki** (moduł 03). To dziedzictwo OOXML. Gdy używasz tekstu („B2"), openpyxl przelicza to za Ciebie. Ale gdy budujesz kotwicę **ręcznie**, musisz pamiętać o zera. To ta sama „pułapka dwóch numeracji", którą widziałeś w module 15 przy `AnchorMarker`.

**Precyzyjna kotwica ręczna** (gdy potrzebujesz pozycji co do piksela albo rozciągania z komórkami):

```python
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D, XDRPoint2D
from openpyxl.utils.units import pixels_to_EMU

# --- OneCellAnchor: kotwica w komorce + STALY rozmiar ---
img = Image(BytesIO(dane_png))
img.anchor = OneCellAnchor(
    _from=AnchorMarker(col=4, row=1),                        # col=4 -> E, row=1 -> wiersz 2 (0-based!)
    ext=XDRPositiveSize2D(cx=pixels_to_EMU(300), cy=pixels_to_EMU(150)),
)
ws.add_image(img)

# --- TwoCellAnchor: obraz rozciaga sie miedzy dwiema komorkami ---
img2 = Image(BytesIO(dane_png))
img2.anchor = TwoCellAnchor(
    editAs="twoCell",                                       # razem z komorkami
    _from=AnchorMarker(col=4, row=10),
    to=AnchorMarker(col=8, row=16),
)
ws.add_image(img2)
```

`AnchorMarker` ma cztery pola: `col`, `colOff`, `row`, `rowOff` — dwa ostatnie to **przesunięcie w EMU** od krawędzi komórki. Dzięki nim możesz ustawić obraz „przy prawej krawędzi komórki". To rzadko potrzebne; wzorce `OneCellAnchor` + `XDRPositiveSize2D` pokrywają prawie wszystko.

**Praktyczna rekomendacja:** używaj prostego `ws.add_image(img, "E2")` plus `img.width`/`img.height`. Ręczne kotwice stosuj tylko wtedy, gdy naprawdę potrzebujesz relatywnego rozmiaru (`TwoCellAnchor`) — i wtedy **testuj okiem**, bo geometria komórek jest zależna od szerokości kolumn i wysokości wierszy.

### 2.5. Cykl życia pliku obrazu — `mark_to_close` i zamknięty strumień

To najważniejsza sekcja praktyczna modułu. Wróć do `_data()` z 2.2 i popatrz na ostatnie trzy linijki:

```python
data = fp.read()
fp.close()          # ← TU ZAMYKAMY STRUMIEN
return data
```

Ten `fp.close()` jest **poprawny** — openpyxl wie, że otworzył ten plik, więc go zamyka (naprawiało to błędy `#750 Adding images keeps file handles open` z 2.4.5 i `#1684 Image files not closed when workbooks are saved` z 3.0.10). Ale ma **dwa skutki uboczne**, o których musisz wiedzieć.

**Skutek 1 — nie usuwaj pliku źródłowego przed `save()`.**

`_data()` **otwiera źródło ponownie** w momencie zapisu. Jeśli podałeś ścieżkę i plik zniknął pomiędzy `add_image` a `wb.save()`:

```python
img = Image("logo.png")
ws.add_image(img, "A1")
os.remove("logo.png")      # :-)
wb.save("raport.xlsx")
# → FileNotFoundError: [Errno 2] No such file or directory: 'logo.png'
```

Analogia: **pinezka trzyma karteczkę, ale karteczka nie jest jeszcze przyklejona do tablicy.** Wstawiasz *wskazanie*, dokument powstaje dopiero przy zapisie. Dopóki plik istnieje przy `save()`, wszystko jest dobrze — ale nie zakładaj, że openpyxl „skopiował go już wcześniej".

**Skutek 2 — ten sam obiekt `Image` może zawieść przy DRUGIM zapisie.**

Prześledź to dla obrazu utworzonego ze **strumienia**:

```python
dane_png = BytesIO(otrzymane_bajty)
img = Image(dane_png)           # self.ref = dane_png  (ten sam obiekt BytesIO)
ws1.add_image(img, "A1")
wb1.save("plik1.xlsx")          # _data():  fp = dane_png.fp ; fp.read() ; fp.close()
                                #           ^^^ dane_png jest TERAZ ZAMKNIĘTY

ws2.add_image(img, "A1")        # <-- ten SAM obiekt img
wb2.save("plik2.xlsx")          # _data():  PILImage.open(<zamkniety BytesIO>)
                                # → ValueError: I/O operation on closed file
```

**Reguła, która to wszystko rozwiązuje — jedna, prosta i odporna na wersję:**

> **Zachowuj BADANIE (bytes), nie obiekt `Image`. Twórz świeży `Image` dla każdego `wb.save()`.**

```python
# DOBRZE: bajty sa zrodlem prawdy, Image jest jednorazowy
dane_png: bytes = zrob_png_bytes()        # raz

for sciezka in ("raport_a.xlsx", "raport_b.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.add_image(Image(BytesIO(dane_png)), "B2")   # swiezy obiekt za kazdym razem
    wb.save(sciezka)
    wb.close()
```

> **Uwaga o ścieżkach.** Obraz utworzony **ze ścieżki** (`Image("logo.png")`) jest odporny na ten problem — bo `_data()` otwiera plik od nowa przy każdym zapisie, a zamyka ten świeżo otwarty uchwyt. Ale powyższa reguła (świeży `Image` per zapis) jest **lepsza**, bo nie zależy od wersji biblioteki ani od tego, czy źródło jest plikiem, czy strumieniem. To jest dokładnie ten rodzaj decyzji defensywnej, którą w module 03 nazwaliśmy „świadomością cyklu życia".

**Skutek 3 — uwaga przy zapisie do `BytesIO` (moduł 20).**

W testach `wb.save(buf)` zapisuje do strumienia. Jeśli w tym samym procesie **ponownie użyjesz tego samego obiektu `Image`** — wpadniesz w dokładnie ten sam problem. Testy zapisujące do pamięci są więc **najlepszym miejscem na przestrzeganie zasady „świeży `Image`"**, bo tam łatwo wpaść w pułapkę przez pośpiech.

### 2.6. Obrazy ze strumienia i generowane w locie

Skoro umiesz już wstawić plik z dysku, przejdźmy do tego, co daje największą moc: **obraz, który powstaje w pamięci i nigdy nie dotyka dysku.**

```python
from io import BytesIO
from openpyxl.drawing.image import Image

# obraz z bajtow (np. z bazy, z API, z sieci)
bajty = requests.get(url_obrazu, timeout=10).content
img = Image(BytesIO(bajty))
ws.add_image(img, "B2")
```

I teraz najważniejszy wzorzec całego modułu — **generowanie obrazu w locie**:

```python
import matplotlib
matplotlib.use("Agg")               # backend bez wyswietlania (serwery!)
import matplotlib.pyplot as plt
from io import BytesIO

def wykres_do_bajtow(labels, values, tytul: str, dpi: int = 120) -> bytes:
    """Rysuje wykres matplotlib i zwraca PNG jako bajty. Nic nie zapisuje na dysk."""
    fig, ax = plt.subplots(figsize=(7, 3.5), dpi=dpi)
    ax.bar(labels, values, color="#5B5CF0")
    ax.set_title(tytul)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, format="png")   # zapis do PAMIECI, nie do pliku
    plt.close(fig)                   # KLUCZOWE: zwolnij pamiec matplotlib!
    return buf.getvalue()            # same bajty
```

Trzy rzeczy, które trzeba tu docenić:

**1. `matplotlib.use("Agg")`** — backend *Agg* rysuje do pliku/strumienia, a nie na ekran. Bez niego na serwerze bez wyświetlacza możesz dostać błędy. To standard w automatyzacji.

**2. `plt.close(fig)`** — jeśli tego nie zrobisz, matplotlib będzie trzymał każdy wykres w pamięci. Przy pętli 500 raportów zjesz RAM. Analogia: **brudne naczynia w zlewie** — zmywak przetwarza, ale nie zmywa; po dwustu talerzach nie ma gdzie pracować.

**3. `buf.getvalue()`** zwraca **bajty**, a nie strumień — to dokładnie ta forma, o której mówiłem w 2.5. Bajty są wieczne, strumień jest jednorazowy.

**Dlaczego to jest ważne — most z modułem 15.** openpyxl **nie potrafi** narysować mapy, treemapu, sunburst, histogramu, Pareto, waterfall, lejka ani pudełka-wąsów (moduł 15, sekcja o „chartex"). Ale `matplotlib` potrafi **wszystko** — i możesz jego wynik wstawić jako **obraz**:

```python
# wykres, ktorego openpyxl NIE potrafi zrobic
def wykres_waterfall(bajty_zrodlowe) -> bytes:
    import numpy as np
    import matplotlib.pyplot as plt
    ...

# a potem po prostu:
ws.add_image(Image(BytesIO(wykres_waterfall(...))), "F2")
```

**Cena tego obejścia (powiedzmy wprost):** obraz to **migawka**, a nie żywy wykres. Zmienisz dane w komórce — obraz się **nie zmieni**. Nie ma też tooltipów, legendy ani klikania. Więc: **żywotne wykresy rób natywnie (moduł 15), a te „niemożliwe" — jako obraz i z wyraźnym podpisem, że to statyczna ilustracja.** To dokładnie ten sam kompromis, co w 2.1: wykres = adres, obraz = kopia bitów.

### 2.7. Dostępność: tekst alternatywny, którego openpyxl nie potrafi ustawić

W instytucjach publicznych (i w firmach objętych regulacjami) dokument musi być **dostępny**: czytnik ekranu dla osoby niewidomej musi umieć opisać, co jest na obrazie. W OOXML robi się to atrybutem **`descr`** na elemencie `xdr:cNvPr` (obok `name` i `title`):

```xml
<xdr:cNvPr id="1" name="Image 1" descr="Wykres sprzedaży wzrósł o 47% r/r"/>
```

**Co robi openpyxl?** Sprawdźmy w `SpreadsheetDrawing._picture_frame`:

```python
def _picture_frame(self, idx):
    pic = PictureFrame()
    pic.nvPicPr.cNvPr.descr = "Picture"       # ← DOSLOWNIE "Picture", dla KAZDEGO obrazu
    pic.nvPicPr.cNvPr.id = idx
    pic.nvPicPr.cNvPr.name = "Image {0}".format(idx)
    ...
```

**openpyxl wpisuje jako opis dosłownie `"Picture"`.** Czyli w praktyce: **każdy obraz dodany przez openpyxl ma tekst alternatywny „Picture"**, niezależnie od treści. Czytnik ekranu przeczyta „Picture, Picture, Picture". To nie tekst alternatywny — to jego atrapa.

I teraz najważniejsze: **klasa `Image` nie ma żadnego API do jego zmiany.** Przeczytałem całą klasę (2.2) — pola to tylko `ref`, `width`, `height`, `format`, `anchor`, `path`, `_id`. **Nie ma `img.descr`, `img.alt_text` ani `img.title`.** `descr` jest generowane twardo przy zapisie i **nie da się go nadpisać z poziomu publicznego API**.

> **⚠️ Sprostowanie do briefu.** Brief sugeruje `img.anchor` jako miejsce tekstu alternatywnego. To **nieprawda** — `img.anchor` to **pozycja** obrazu (kotwica z 2.4), a nie opis. Nie ma w `Image` żadnego atrybutu opisu. Mieszanie tych dwóch pojęć jest częste w tutorialach, bo `anchor` jest jedynym atrybutem obrazu „o treści", na jaki się natykają.

**Co więc zrobić, jeśli dostępność jest wymogiem?**

**1. Najprościej: ustaw opis w Excelu.** Otwórz plik, kliknij obraz prawym, **Formatuj obraz → Tekst alternatywny**, wpisz. Jeśli generujesz raport raz na miesiąc i ktoś go „domyka" — to wystarczy, **pod warunkiem** że raport nie przejdzie potem przez openpyxl (moduł 18!).

**2. Obejście: dopisz `descr` do XML po zapisie.** Skoro openpyxl generuje `descr="Picture"`, ten sam plik można podmienić w archiwum. To **nieoficjalne** i wymaga ostrożności (moduł 18: edycja XML psuje spójność, jeśli zrobisz to źle), ale jest wykonalne:

```python
"""NIEOWERSYJNE obejscie: podmiana descr="Picture" w xl/drawings/drawing1.xml.
Testuj na swojej wersji. Rob na KOPII pliku."""
import re
import shutil
import zipfile
from pathlib import Path

def ustaw_alt_text(sciezka: Path, opisy: list[str]) -> Path:
    """Podmienia kolejne descr="Picture" na podane opisy. Zwraca nowy plik."""
    nowa = sciezka.with_name(sciezka.stem + "_alt" + sciezka.suffix)
    with zipfile.ZipFile(sciezka) as src, zipfile.ZipFile(nowa, "w",
                                                          zipfile.ZIP_DEFLATED) as dst:
        licznik = 0
        for item in src.infolist():
            dane = src.read(item.filename)
            if item.filename.startswith("xl/drawings/drawing") and item.filename.endswith(".xml"):
                tekst = dane.decode("utf-8")
                def podmien(_m, _it=iter(opisy)):
                    nonlocal licznik
                    opis = next(_it, "Picture")
                    licznik += 1
                    return f'descr="{opis}"'
                tekst = re.sub(r'descr="[^"]*"', podmien, tekst)
                dane = tekst.encode("utf-8")
            dst.writestr(item, dane)
    return nowa
```

**Uczciwe ostrzeżenie:** to hack. Zależny od tego, jak openpyxl zapisze XML w Twojej wersji, i od tego, że regex nie zepsuje encji. **Używaj tylko wtedy, gdy dostępność jest twardym wymogiem, testuj plik w Excelu, i zapisz to w komentarzu w kodzie.** Normalnie: **poproś o uzupełnienie opisów w procesie** albo **dodaj podpisy jako zwykły tekst w komórkach obok obrazu** (co jest zresztą lepsze dla wszystkich czytelników, nie tylko dla czytników ekranu).

**Wniosek praktyczny:** **traktuj dostępność obrazów jako ograniczenie openpyxl, nie jako funkcję.** Zaplanuj to z góry, a nie na końcu projektu.

### 2.8. Hiperlinki: model `Hyperlink` i poprawka do najczęstszego błędu

Wróćmy do 2.1: **hiperlink zewnętrzny to relacja, wewnętrzny to atrybut.** Zobaczmy, skąd to wynika — oto cała klasa z 3.1.5:

```python
# openpyxl/worksheet/hyperlink.py
class Hyperlink(Serialisable):
    tagname = "hyperlink"

    ref = String()                        # komorka, do ktorej nalezy link ("B2")
    location = String(allow_none=True)    # ← NAWIGACJA W SKOROSZYCIE ("Dane!A1")
    tooltip = String(allow_none=True)     # podpowiedz przy najechaniu
    display = String(allow_none=True)     # tekst wyswietlany
    id = Relation()                       # Identyfikator relacji (dla linkow ZEWNETRZNYCH)
    target = String(allow_none=True)      # ← ADRES ZEWNETRZNY ("https://...")

    __attrs__ = ("ref", "location", "tooltip", "display", "id")   # target NIE jest atrybutem XML!
```

**Zwróć uwagę na `__attrs__`.** Nie ma tam `target`. To nie przeoczenie — to architektura. `target` **nie trafia do XML jako atrybut**; trafia do **pliku relacji**. Zobaczmy w writerze:

```python
# openpyxl/worksheet/_writer.py
def write_hyperlinks(self):
    links = self.ws._hyperlinks
    for link in links:
        if link.target:
            rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target)
            self._rels.append(rel)
            link.id = rel.id
    if links:
        self.xf.send(HyperlinkList(links).to_tree())
```

Trzy linijki i cała prawda:

1. **Jeśli link ma `target`** → powstaje relacja z `TargetMode="External"` (adres zewnętrzny: `https://`, `mailto:`, plik).
2. **Jeśli link ma tylko `location`** → nie powstaje żadna relacja; `<hyperlink ref="B2" location="Dane!A1"/>` ląduje w XML arkusza. To jest **link wewnętrzny**.
3. `link.id = rel.id` dopisuje identyfikator relacji, gdy relacja powstała.

**Teraz poprawka.** W internecie (i w briefie) znajdziesz wzorzec:

```python
ws["B4"].hyperlink = "#Dane!A1"      # ← ŹLE!
```

To **nie działa poprawnie**. Ponieważ przypisujesz **tekst**, setter (sekcja niżej) opakuje go w `Hyperlink(target="#Dane!A1")`. A skoro `target` jest ustawiony, writer utworzy **relację zewnętrzną** z adresem `#Dane!A1` — czyli Excel spróbuje potraktować to jako zewnętrzny zasób o dziwnym adresie, zamiast nawigować w skoroszycie. Efekt: link, który nie prowadzi tam, gdzie chcesz (albo prowadzi w sposób zależny od wersji/przeglądarki pliku).

**Poprawnie — link wewnętrzny przez `location`:**

```python
from openpyxl.utils import quote_sheetname
from openpyxl.worksheet.hyperlink import Hyperlink

nazwa = "Dane 2026"
ws["B4"].hyperlink = Hyperlink(
    ref="B4",                                       # zostanie nadpisane przez setter, ale musi byc
    location=f"{quote_sheetname(nazwa)}!A1",        # → "'Dane 2026'!A1"
    display="Przejdź do danych",
)
ws["B4"].value = "Przejdź do danych"                # napis widoczny (drzwi + napis)
```

`quote_sheetname` to znajomy pomocnik z modułu 07 — wstawia apostrofy wokół nazwy ze spacją. **Dla nazw bez spacji możesz podać `location="Dane!A1"`; dla nazw ze spacjami apostrofy są konieczne.** Możesz też linkować do **nazwy zdefiniowanej** (moduł 26): `location="MojaNazwa"`.

**A teraz setter `cell.hyperlink` — zachowanie, które trzeba znać:**

```python
# openpyxl/cell/cell.py
@hyperlink.setter
def hyperlink(self, val):
    """Set value and display for hyperlinks in a cell.
    Automatically sets the `value` of the cell with link text,
    but you can modify it afterwards by setting the `value`
    property, and the hyperlink will remain.
    Hyperlink is removed if set to ``None``."""
    if val is None:
        self._hyperlink = None                  # ← usuwanie linku
    else:
        if not isinstance(val, Hyperlink):
            val = Hyperlink(ref="", target=val)  # ← tekst staje sie Targetem (ZEWNETRZNY!)
        val.ref = self.coordinate               # ← ref zawsze = koordynat komorki
        self._hyperlink = val
        if self._value is None:
            self.value = val.target or val.location   # ← AUTOMATYCZNE uzupelnienie napisu!
```

Cztery konsekwencje, każda praktyczna:

**1. `cell.hyperlink = None` usuwa link.** (Dawniej był to błąd `#697 Cannot unset hyperlinks`, naprawiony w 2.4.1.)

**2. `val.ref` jest zawsze nadpisywane.** Dlatego podanie `ref=` w konstruktorze jest wymagane technicznie (deskryptor `String()` nie lubi `None`), ale i tak zostanie zastąpione poprawnym koordynatem. **Nie musisz się tym przejmować.**

**3. Tekst staje się `target` — czyli linkiem ZEWNĘTRZNYM.** To wyjaśnia, dlaczego `"#Dane!A1"` nie robi tego, co chcesz: każdy `str` idzie tą samą drogą.

**4. ⚠️ Komórka NIE zostanie pusta — openpyxl wpisze w nią adres.** To najważniejszy szczegół:

```python
ws["D2"].hyperlink = "https://example.com/auto"   # komorka byla PUSTA
print(ws["D2"].value)      # → 'https://example.com/auto'   ← openpyxl wpisal adres!
```

> **Sprostowanie do briefu.** Brief mówi „hiperlink bez `cell.value` (puste pole)". Prawda jest inna i, szczerze mówiąc, **gorsza dla estetyki**: openpyxl **sam uzupełni wartość komórki adresem**, jeśli komórka jest pusta. Nie dostaniesz pustego pola — dostaniesz brzydki surowy URL jako napis. **Poprawna kolejność to: najpierw `cell.value = "ładny napis"`, potem `cell.hyperlink = adres`.** Wtedy `_value is not None` i openpyxl nie nadpisze Twojego napisu. (Albo odwrotnie — ustaw link, a potem nadpisz `value`; setter sam mówi: *„you can modify it afterwards"*.)

**Styl hiperlinku.** Excel nie nadaje linkom niebieskiego podkreślenia automatycznie (to robi dopiero po kliknięciu, gdy stosuje „Odwiedzony hiperlink"). W openpyxl możesz użyć wbudowanego stylu nazwanego (moduł 09) — dokumentacja stylów wymienia go w grupie **„Text styles"**:

```python
ws["B2"].style = "Hyperlink"            # wbudowany styl nazwany (niebieski + podkreslenie)
```

Wbudowane style nazwane dostępne w openpyxl (z dokumentacji; **tylko angielskie nazwy, dokładnie tak zapisane**):
`'Normal'`, `'Comma'`, `'Comma [0]'`, `'Currency'`, `'Currency [0]'`, `'Percent'`, `'Calculation'`, `'Total'`, `'Note'`, `'Warning Text'`, `'Explanatory Text'`, `'Title'`, `'Headline 1'`–`'Headline 4'`, **`'Hyperlink'`**, **`'Followed Hyperlink'`**, `'Linked Cell'`, `'Input'`, `'Output'`, `'Check Cell'`, `'Good'`, `'Bad'`, `'Neutral'`, `'Accent1'`–`'Accent6'` i ich warianty procentowe, `'Pandas'`.

**Zasada ostrożności:** przypisanie `cell.style = "Hyperlink"` **nadpisuje cały styl komórki** (moduł 09: styl to komplet sześciu pieczątek). Jeśli komórka ma już format liczby albo wypełnienie, użyj `copy()` i zmodyfikuj tylko `font` — albo ustaw `style` **przed** innymi modyfikacjami.

**Trzy rodzaje linków i ich `target`:**

```python
# 1. ZEWNETRZNY (relacja External)
ws["B2"].hyperlink = "https://example.com"              # strona
ws["B3"].hyperlink = "mailto:raporty@example.com"       # e-mail
ws["B4"].hyperlink = "file:///C:/raporty/2026.xlsx"     # plik lokalny (URI!)
ws["B5"].hyperlink = "raport_pomocniczy.xlsx"           # plik relatywny

# 2. WEWNETRZNY (location, bez relacji)
ws["B6"].hyperlink = Hyperlink(ref="B6", location="'Dane 2026'!A1")
ws["B7"].hyperlink = Hyperlink(ref="B7", location="MojaNazwaZdefiniowana")

# 3. Z TOOLTIPEM
ws["B8"].hyperlink = Hyperlink(ref="B8", target="https://example.com",
                               tooltip="Otworzy się w przeglądarce")
```

> **Uwaga o adresach plików.** W `target` używaj **URI** (`file:///...`), nie zwykłej ścieżki Windows (`C:\...`). Backslashe w URI są niepoprawne. Ścieżka relatywna (`raport.xlsx`) działa jako relacja względem lokalizacji pliku — ale jeśli skoroszyt zostanie przeniesiony, link się zerwie (moduł 18: to jeden z „ukrytych" powodów, dla których plik wygląda inaczej na innym komputerze).

### 2.9. Komentarze: tekst, autor, rozmiar — i co się traci

Pełne źródło klasy (3.1.5):

```python
# openpyxl/comments/comments.py
class Comment(object):
    _parent = None

    def __init__(self, text, author, height=79, width=144):
        self.content = text          # ← tekst trzymany w .content
        self.author = author
        self.height = height         # ← PIKSELE
        self.width = width           # ← PIKSELE

    def __eq__(self, other):
        return self.content == other.content and self.author == other.author

    def __repr__(self):
        return "Comment: {0} by {1}".format(self.content, self.author)

    def __copy__(self):
        """Create a detached copy of this comment."""
        clone = self.__class__(self.content, self.author, self.height, self.width)
        return clone

    def unbind(self):
        """Unbind a comment from a cell"""
        self._parent = None

    @property
    def text(self):
        """Any comment text stripped of all formatting."""
        return self.content

    @text.setter
    def text(self, value):
        self.content = value
```

Czytaj to uważnie, bo tu są trzy niespodzianki:

**1. `text` to alias `content`.** Oba działają (`c.text = "x"` i `c.content = "x"` robią to samo). W dokumentacji używa się `text`, w XML-u i w `from_cell` — `content`. Używaj `text` (czytelniejszy), ale wiedz, że istnieje drugi.

**2. Domyślne wymiary: $144 \times 79$ pikseli.** To okienko wąskie i niskie — tekst się w nim nie zmieści, jeśli jest dłuższy niż jedno zdanie. Ustawiaj rozmiar ręcznie. Pamiętaj, że to **piksele**, więc do przeliczenia z punktów (typografia!) użyj `points_to_pixels`:

```python
from openpyxl.comments import Comment
from openpyxl.utils import units

c = Comment("Tekst", "Autor")
c.width = 300
c.height = 90
# albo z punktow (typografia):
c.width = units.points_to_pixels(220)      # ~293 px
c.height = units.points_to_pixels(60)      # ~80 px
```

**3. `__copy__` i `_parent`.** Widzisz mechanizm automatycznego kopiowania z 1.4 — `__copy__` tworzy „odłączoną kopię". I widzisz `_parent` (komórka, do której komentarz należy) oraz `unbind()` — wewnętrzne narzędzia openpyxl.

**Co się traci przy odczycie (dokumentacja mówi wprost):**

> *„Openpyxl currently supports the reading and writing of comment text only. Formatting information is lost. Comment dimensions are lost upon reading, but can be written."*

Czyli:

- ✅ **tekst i autor** — przechodzą w obie strony;
- ❌ **formatowanie tekstu** (pogrubienie, kolor fragmentu) — tracone;
- ❌ **wymiary i pozycja okienka przy odczycie** — tracone (choć można je **zapisać**; to asymetria: napiszesz, ale nie odczytasz);
- ❌ **nowe komentarze wątkowe** (`xl/threadedComments/`, Excel 365) — openpyxl modeluje **klasyczne notatki** (`xl/comments{n}.xml` + `xl/vmlDrawing{n}.vml`). Nie zakładaj, że komentarze wątkowe przetrwają zapis — **sprawdź na swojej wersji i obejrzyj plik w Excelu.**

To wyjaśnia zdanie z 2.1: **komentarz ma dwie części w dwóch technologiach, a openpyxl obsługuje dobrze tylko nowszą.** Wygląd okienka (VML) jest drugorzędny, więc autor tego kodu postawił na treść.

### 2.10. Komentarze w scalonych komórkach — ostrzeżenie, które trzeba wypisać wprost

To mały, ale bardzo konkretny kawałek wiedzy — i openpyxl wypisuje go **własnym komunikatem**. Oto warunek (fragment `read_worksheets` z `openpyxl/reader/excel.py`, 3.1.5):

```python
comment_warning = """Cell '{0}':{1} is part of a merged range but has a comment which will be removed because merged cells cannot contain any data."""
...
# assign any comments to cells
for r in rels.find(COMMENTS_NS):
    src = self.archive.read(r.target)
    comment_sheet = CommentSheet.from_tree(fromstring(src))
    for ref, comment in comment_sheet.comments:
        try:
            ws[ref].comment = comment
        except AttributeError:
            c = ws[ref]
            if isinstance(c, MergedCell):
                warnings.warn(comment_warning.format(ws.title, c.coordinate))
                continue
```

Przetłumaczmy to na praktykę (dokładny komunikat to: *„Cell 'Arkusz':C2 is part of a merged range but has a comment which will be removed because merged cells cannot contain any data."*):

**1. Przy próbie ustawienia komentarza na „pod-komórce" scalonego zakresu dostajesz `AttributeError`.** Bo `ws["C2"]` w scalonym `B2:D2` zwraca `MergedCell`, a `MergedCell` **nie ma** atrybutu `comment`:

```python
ws.merge_cells("B2:D2")
ws["B2"].comment = Comment("To zadziala (lewy gorny)", "Anna")     # ✅
try:
    ws["C2"].comment = Comment("To NIE zadziala", "Anna")           # ❌
except AttributeError as e:
    print("AttributeError:", e)   # 'MergedCell' object has no attribute 'comment'
```

**2. Komentarze w scalonych komórkach istnieją — Excel je robi.** Excel pozwala „przykleić" notatkę do scalonej komórki. openpyxl takiego pliku **nie odtworzy**: przy wczytaniu wypisze ostrzeżenie i **pominie** komentarz. Analogia: **kartka przyklejona do środka złączonych kart** — przy przepisywaniu na czysto nie ma jej gdzie umieścić, więc zostaje w notatkach na marginesie (to właśnie ostrzeżenie), a nie w nowym dokumencie.

**3. Dlaczego to ostrzeżenie ma znaczenie w module 18.** Jeśli przepuszczasz przez openpyxl plik z Excela i pojawia się to ostrzeżenie — **wiesz dokładnie, że tracisz dane.** To jeden z tych sygnałów, które warto **zbierać**, a nie ignorować: w module 18 zbudujesz nasłuch ostrzeżeń i zamienisz go w raport ryzyka.

**Praktyczne rozwiązanie:** komentarz umieszczaj **w lewej górnej komórce** scalonego zakresu (to działa) albo **w osobnej komórce obok**, zamiast „w środku" scalenia.

### 2.11. Tryby pracy: `read_only`, `write_only`, `Chartsheet`

Sprawdźmy tę trójkę dla naszych elementów — i wynik jest ** odwrotnie proporcjonalny do intuicji**.

**`read_only=True` — obrazy i komentarze niedostępne, hiperlinki dostępne.**

W czytniku, dla arkusza czytanego leniwie, wykonywany jest `continue` **przed** przetwarzaniem rysunków (moduł 15). Do tego dochodzi ostrzeżenie z dokumentacji komentarzy: *„Comments are not currently supported if `read_only=True` is used."* Hiperlinki natomiast są zwykłymi atrybutami w `sheet1.xml` — więc przechodzą.

**`write_only=True` — wszystko działa, w tym obrazy.** To jest zaskoczenie. Zobacz źródło:

```python
# openpyxl/worksheet/write_only.py
# Methods from Worksheet
self._add_row = Worksheet._add_row.__get__(self)
self._add_column = Worksheet._add_column.__get__(self)
self.add_chart = Worksheet.add_chart.__get__(self)
self.add_image = Worksheet.add_image.__get__(self)     # ← OBRAZY DZIALAJA
self.add_table = Worksheet.add_table.__get__(self)
```

`WriteOnlyWorksheet` **podpina** metody `add_image` i `add_chart` ze zwykłego `Worksheet`. I to ma sens, gdy się nad tym zastanowić: dodanie obrazu **nie wymaga czytania komórek** — wystarczy kotwica (adres) i bajty. Historia biblioteki to potwierdza: `#729 Allow images in write-only mode` (2.4.4) i *„Write-only workbooks support charts and images"* (2.3.0-b1).

**Praktyczna kolejność w `write_only`:** obraz i wykres dodawaj **przed `wb.save()`**, a najbezpieczniej jako **pierwszą operację na arkuszu** (przed `append`). Ogólna reguła dla trybu strumieniowego — *„wszystko, co pojawia się w pliku przed danymi komórek, musi być utworzone przed dodaniem komórek"* — dotyczy wymiarów kolumn, ustawień wydruku i ochrony; obrazy i wykresy są obsługiwane przy zapisie, ale trzymanie ich na początku jest najprostszą dyscypliną.

**`Chartsheet` — brak `add_image`.** Próba wstawienia obrazu na arkusz wykresów kończy się `AttributeError: 'Chartsheet' object has no attribute 'add_image'`. To potwierdzone zachowanie biblioteki (odpowiedź autora openpyxl w dyskusji o tym przypadku brzmi krótko: *„Can't be done with openpyxl"*). Jeśli obraz ma być na „dużym arkuszu wykresów" — wstaw go na normalny arkusz i ukryj siatkę (moduł 11).

**Obrazy w skoroszycie z makrami.** Historycznie bywał z tym problem (`#464 Cannot use images when preserving macros`). Przy `keep_vba=True` (moduł 18) **sprawdź obraz w Excelu** — to obszar, który warto traktować jako „potrafi częściowo".

### 2.12. `copy_worksheet` a nasza trójka

W module 08 powiedziałeś: „`copy_worksheet` nie kopiuje obrazków i wykresów". Teraz dopiszmy do tego pełny obraz — dokumentacja tutoriala mówi:

> *„Only cells (including values, styles, hyperlinks and comments) and certain worksheet attributes (including dimensions, format and properties) are copied. All other workbook / worksheet attributes are not copied - e.g. Images, Charts."*

Czyli, dla jasności:

| Element | Kopiowany przez `copy_worksheet`? |
|---|---|
| **Wartości i style komórek** | ✅ tak |
| **Hiperlinki** | ✅ **tak** |
| **Komentarze** | ✅ **tak** |
| **Obrazy** | ❌ **nie** |
| **Wykresy** | ❌ **nie** |
| **Formatowanie warunkowe, walidacja, tabele** | ❌ nie (to nie jest „atrybut arkusza" w tym sensie) |

Potwierdza to źródło `openpyxl/worksheet/copier.py` — kopiowane są komórki (`_copy_cells`, z hiperlinkami i komentarzami), wymiary i zestaw atrybutów. **Ani `_images`, ani `_charts` nie są w tym zestawie.**

**Praktyczny wniosek dla „generatora miniatur" (2.13):** jeśli chcesz mnożyć arkusze z obrazami (np. karta produktu na arkusz), **`copy_worksheet` Ci nie wystarczy** — musisz **dodać obraz do każdego nowego arkusza jawnie**. I tu przydaje się zasada z 2.5: trzymaj **bajty**, nie obiekt `Image`, żeby móc tworzyć nowe obrazy dla każdego arkusza bez ryzyka zamkniętego strumienia.

### 2.13. Wzorzec: „generator miniatur" i jedna funkcja = jedna czynność

Brief nazywa to „generatorem miniatur". Rozszerzmy to na wzorzec ogólniejszy, bo to most do modułu 24 (Facade).

Popatrz na naiwny kod, który tworzy kartę produktu z miniaturą:

```python
# ZLE: wszystko w jednym miejscu, jednostki "na oko", obraz z nazwy pliku
for produkt in produkty:
    ws = wb.create_sheet(produkt.nazwa[:31])
    ws["A1"] = produkt.nazwa
    img = Image(f"photos/{produkt.sku}.jpg")
    img.width = 200                                   # skad 200? nie wiadomo
    img.height = 150                                  # znieksztalcone!
    ws.add_image(img, "B3")
    ws["A1"].hyperlink = f"https://sklep.example/{produkt.sku}"   # link na A1, nie B3!
    ws["D3"].comment = Comment("Cena netto", "system")            # rozmiar domyslny, za maly
    ws.row_dimensions[3].height = 130                 # "na oko", czy zmiesci obraz?
```

Cztery problemy, każdy inny:

1. **Jednostki i proporcje liczone ręcznie** — `200 × 150` to nie jest proporcja zdjęcia, więc obraz jest rozciągnięty. Brak `cm_to_pixels`/skalowania (2.3).
2. **Kotwica i wysokość wiersza nie są uzgodnione** — obraz w `B3` o wysokości $150$ px potrzebuje wiersza o wysokości $\approx 113$ pt. Wartość `130` nie wynika z niczego.
3. **Ten sam błąd z linkiem** — link trafia w `A1`, a napis jest w `B2`/`B4`, więc użytkownik klika dwa pola obok.
4. **Domyślny rozmiar komentarza** — $144 \times 79$ px to za mało na zdanie.

**Wzorzec: jedna funkcja = jedna czynność.** Analogia: **stanowiska w fabryce**. Nie robisz całego produktu na jednym stole — masz osobne stanowisko do cięcia, do montażu, do pakowania. Każde ma **jedno** zadanie i **jasny kontrakt** na wejściu i wyjściu.

```python
# DOBRZE: kazda czynnosc ma swoja funkcje z jawnym kontraktem

MINIATURA_CM = 3.5          # szerokosc miniatury w cm - DECYZJA WZGLEDNA, nie w pikselach
MINIATURA_DPI = 96

def dodaj_miniature(ws, bajty_png: bytes, anchor: str, szerokosc_cm: float) -> None:
    """Wstawia obraz o zadanej SZEROKOSCI w cm, z zachowaniem proporcji."""
    img = Image(BytesIO(bajty_png))                     # swiezy obiekt - sekcja 2.5
    img.width = round(szerokosc_cm / 2.54 * MINIATURA_DPI)
    skala = img.width / Image(BytesIO(bajty_png)).width  # proporcja z zrodla
    img.height = round(Image(BytesIO(bajty_png)).height * skala)
    ws.add_image(img, anchor)

def ustaw_wysokosc_wiersza_pod_miniature(ws, wiersz: int, wysokosc_cm: float) -> None:
    """Wysokosc wiersza w PUNKTACH (modul 11), policzona z cm."""
    ws.row_dimensions[wiersz].height = round(wysokosc_cm / 2.54 * 72)

def dodaj_link_z_napisem(ws, komorka: str, napis: str, adres: str) -> None:
    """Napis NAJPIERW, potem link - inaczej openpyxl wpisze adres (sekcja 2.8)."""
    ws[komorka] = napis
    ws[komorka].hyperlink = adres
    ws[komorka].style = "Hyperlink"

def dodaj_przypis(ws, komorka: str, tekst: str, autor: str,
                  szer_px: int = 340, wys_px: int = 110) -> None:
    """Komentarz-czytelnikowi z rozmiarem, ktory pomiesci zdanie."""
    komentarz = Comment(tekst, autor)
    komentarz.width = szer_px
    komentarz.height = wys_px
    ws[komorka].comment = komentarz
```

Zwróć uwagę na trzy rzeczy, które ten refaktor naprawił:

- **`szerokosc_cm` zamiast pikseli w kodzie głównym.** Piksele to jednostka implementacyjna; **centymetry to jednostka intencji**. Mówisz „miniatura ma 3,5 cm", a funkcja przelicza. To ta sama zasada co `FORMAT["waluta_pln"]` w module 10 i `STYL` w module 15: **jedno źródło decyzji**.
- **Skalowanie proporcjonalne w jednym miejscu.** Naprawiając „rozciągnięte zdjęcia" w jednym miejscu, naprawiasz je wszędzie.
- **Kolejność `napis` → `hyperlink`** jest teraz **wymuszona przez interfejs funkcji**. Nie da się jej pomylić. To jest właśnie ta „architektura, która nie pozwala popełnić błędu" z modułu 22.

I najważniejsze: **`dodaj_miniature` przyjmuje `bytes`, nie `Image`.** To nie przypadek — to **bariera zapobiegająca pułapce z 2.5**. Funkcja sama decyduje, jak długo żyje obiekt `Image`. W module 24 nazwiemy to **enkapsulacją cyklu życia** i będzie to jeden z fundamentów Fasady.

## 3. Przykłady krok po kroku

### Przykład 1 — logo: jednostki, skalowanie, kalibracja (🟢)

Ten przykład jest **samowystarczalny**: generuje plik PNG za pomocą Pillow, więc nie musisz niczego przygotowywać.

```python
"""Modul 16 - obraz: wstawianie, jednostki, kalibracja rozmiaru.

Uruchom:  python examples/16_obraz_podstawy.py
Wymaga:   pip install openpyxl pillow
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.units import EMU_to_pixels

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
ASSETS = ROOT / "assets"
OUTPUT.mkdir(parents=True, exist_ok=True)
ASSETS.mkdir(parents=True, exist_ok=True)

CEL = OUTPUT / "16_obraz_podstawy.xlsx"
LOGO = ASSETS / "logo.png"

DPI_EXCEL = 96            # rozdzielczosc bitmap w Excelu


# ----------------------------------------------------------------------
# 0. Pomocnicze: konwersje, ktorych openpyxl NIE ma (sekcja 2.3)
# ----------------------------------------------------------------------
def cm_to_pixels(cm: float, dpi: int = DPI_EXCEL) -> int:
    """Centymetry -> piksele (jednostka img.width / img.height)."""
    return round(cm / 2.54 * dpi)


def pixels_to_cm(px: float, dpi: int = DPI_EXCEL) -> float:
    return round(px * 2.54 / dpi, 2)


def zrob_logo_png(sciezka: Path, szer: int = 320, wys: int = 120) -> Path:
    """Generuje prosty prostokatny 'logotyp' - zeby przyklad byl samodzielny."""
    from PIL import Image as PILImage, ImageDraw

    obraz = PILImage.new("RGB", (szer, wys), (255, 255, 255))
    rysuj = ImageDraw.Draw(obraz)

    # ramka + kwadrat jako "znak firmowy"
    rysuj.rectangle([0, 0, szer - 1, wys - 1], outline=(91, 92, 240), width=5)
    rysuj.rectangle([14, 14, 14 + wys - 28, wys - 15], fill=(91, 92, 240))
    # napis domyslna czcionka bitmapowa (bez truetype - dziala wszedzie)
    rysuj.text((wys + 8, wys // 2 - 6), "ACME Sp. z o.o.", fill=(20, 20, 20))

    obraz.save(sciezka, format="PNG")
    return sciezka


def zbuduj(cel: Path) -> None:
    if not LOGO.exists():
        zrob_logo_png(LOGO)

    # --- 1. Poznaj ORYGINALNE wymiary (zawsze w pikselach!) ----------
    img = Image(str(LOGO))
    print("=" * 78)
    print(f"Obraz zrodlowy: {LOGO.name}")
    print(f"  img.width  = {img.width} px   ({pixels_to_cm(img.width)} cm)")
    print(f"  img.height = {img.height} px  ({pixels_to_cm(img.height)} cm)")
    print(f"  img.format = {img.format!r}")
    print("=" * 78)
    proporcja = img.height / img.width
    print(f"Proporcja wysokosc/szerokosc = {proporcja:.4f}  "
          f"(zapamietaj ja - bez niej obraz sie rozciagnie)")
    print()

    # --- 2. Chcemy obraz o szerokosci 4 cm. Ile to pikseli? ----------
    docelowa_szer_cm = 4.0
    img.width = cm_to_pixels(docelowa_szer_cm)
    img.height = round(img.width * proporcja)      # ZACHOWANIE PROPORCJI
    print(f"Przeskalowane na {docelowa_szer_cm} cm szerokosci:")
    print(f"  img.width  = {img.width} px")
    print(f"  img.height = {img.height} px  "
          f"({pixels_to_cm(img.height)} cm)")

    # --- 3. Skoroszyt -------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Raport"

    ws["B2"] = "Raport miesięczny"
    ws["B2"].font = __import__("openpyxl").styles.Font(bold=True, size=14)
    ws["B4"] = "Logo poniżej jest OBRAZEM, nie wartością komórki:"
    ws["B6"] = "(tu nie ma nic w B7 - obraz 'wisi' nad siatką)"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 40

    # obraz kotwiczony w B7
    img.anchor = "B7"
    ws.add_image(img)

    # WAŻNE: wysokosc wiersza 7 musi "zmiescic" obraz.
    # img.height jest w PIKSLACH, row_dimensions[..].height w PUNKTACH!
    from openpyxl.utils.units import pixels_to_points
    ws.row_dimensions[7].height = pixels_to_points(img.height)

    # --- 4. DRUGI obraz - celowo "zly" (rozciagniety) ----------------
    img_rozciagniety = Image(str(LOGO))
    img_rozciagniety.width = 400          # sama szerokosc...
    img_rozciagniety.height = 60          # ...i sama wysokosc = DEFORMACJA
    ws["B10"] = "Zły przykład (wymiary ustawione niezależnie):"
    ws.add_image(img_rozciagniety, "B11")

    # --- 5. TRZECI obraz - bez anchor -> wpadnie w A1! ---------------
    # Obraz "samosiadajacy" - zobacz, gdzie wyladuje (sekcja 1.3!)
    # (celowo zakomentowany, aby nie zaslonic danych - odkomentuj i sprawdz)
    # ws.add_image(Image(str(LOGO)))   # -> wyladuje w A1

    wb.save(cel)
    wb.close()

    print(f"\nZapisano: {cel}")


def weryfikuj(cel: Path) -> None:
    """Wczytuje plik i pokazuje, co NAPRAWDE zostalo zapisane."""
    wb = load_workbook(cel)             # NIE read_only! (sekcja 2.11)
    try:
        ws = wb["Raport"]
        obrazy = list(ws._images)       # atrybut PRYWATNY - do diagnostyki
        print("=" * 78)
        print(f"Obrazow w arkuszu '{ws.title}': {len(obrazy)}")
        print("=" * 78)

        for i, img in enumerate(obrazy, 1):
            print(f"[{i}] {img.path}   format={img.format}")
            # UWAGA: po wczytaniu img.width to znowu ORYGINALNY rozmiar!
            print(f"    img.width/height (z pliku obrazu) = "
                  f"{img.width} x {img.height} px")

            anchor = img.anchor
            if hasattr(anchor, "_from") and hasattr(anchor, "ext"):
                print(f"    OneCellAnchor od (col={anchor._from.col}, "
                      f"row={anchor._from.row})   <- 0-based!")
                wyswietlane = (EMU_to_pixels(anchor.ext.cx),
                               EMU_to_pixels(anchor.ext.cy))
                print(f"    WYSWIETLANY rozmiar z anchor.ext = "
                      f"{wyswietlane[0]} x {wyswietlane[1]} px")
            else:
                print(f"    anchor: {anchor!r}")

        print()
        print("KLUCZOWY WNIOSEK: po wczytaniu img.width pokazuje ORYGINAL,")
        print("a faktyczny rozmiar wyswietlania siedzi w anchor.ext (EMU).")
        print("Wczytujac cudzy plik, nie polegaj na img.width!")
    finally:
        wb.close()

    print()
    print("=" * 78)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 78)
    print("  1. Kliknij obraz -> pasek nazwy pokaze 'Obraz 1'. Nacisnij strzalke,")
    print("     aby 'wyjsc' z obrazu na komorke - przekonasz sie, ze obraz")
    print("     NIE jest wartoscia zadnej komorki.")
    print("  2. Kliknij B7. W pasku formuly zobaczysz '(tu nie ma nic...)' -")
    print("     czyli tekst, a obraz jest NAD nim.")
    print("  3. Porownaj pierwszy obraz (proporcjonalny) z drugim (spłaszczony).")
    print("  4. Posortuj zakres B4:B6 (Dane -> Sortuj). Obrazy ZOSTANA na miejscu!")
    print("     Kotwica sie nie przesuwa razem z wartosciami (sekcja 1.3).")
    print("  5. Zaznacz obraz w B11 i przecignij go nad kolumne D - nic sie nie")
    print("     zmieni w danych. Obraz plywa nad siatka.")


def main() -> None:
    zbuduj(CEL)
    print()
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Po `Image(str(LOGO))` openpyxl **otworzył plik, odczytał rozmiar i format, i zamknął uchwyt** (`mark_to_close` z 2.2). W pamięci zostały: ścieżka (`ref`), wymiary w pikselach i format. **Bajty obrazu jeszcze nie zostały wczytane.** Przypisanie `img.anchor = "B7"` to tylko napis — konwersja na kotwicę z indeksami od zera nastąpi przy zapisie.

**Co trafi do pliku.** Nowa część `xl/media/image1.png` (kopia bajtów!), nowa część `xl/drawings/drawing1.xml` z **jednym** `oneCellAnchor` na obraz (uwaga: `_id` obu obrazów jest `1`, więc oba trafiają do jednego rysunku jako kolejne kotwice), oraz relacje wiążące arkusz → rysunek → `media/`. Sama pozycja w `drawing1.xml` wygląda tak:

```xml
<xdr:oneCellAnchor>
  <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
            <xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:ext cx="3781425" cy="1419225"/>       <!-- EMU: piksele * 9525 -->
  <xdr:pic>
    <xdr:nvPicPr>
      <xdr:cNvPr id="1" name="Image 1" descr="Picture"/>   <!-- ← to "Picture"! -->
```

**Trzy rzeczy do przemyślenia:**

1. **`<xdr:col>1</xdr:col>` to kolumna B, a `<xdr:row>6</xdr:row>` to wiersz 7.** Zera w XML-u, jedynki w kodzie (sekcja 2.4). Jeśli kiedyś będziesz debugować kotwice, to jest miejsce, w którym się pogubisz.
2. **`descr="Picture"` jest tu na Twoich oczach.** Nie musisz mi wierzyć — otwórz `drawing1.xml` z wygenerowanego pliku i zobacz. To najlepszy dowód z całego modułu, że dostępność obrazów nie jest obsługiwana (2.7).
3. **`img.width` po ponownym wczytaniu pokazuje $320$ px, nie $151$ px.** Zwróć uwagę na wynik `weryfikuj()`: otworzyliśmy plik, a `img.width` pokazuje **oryginalny** rozmiar obrazu, bo czytnik tworzy nowy obiekt `Image(BytesIO(bajty_obrazu))`, a skalę trzyma w `anchor.ext`. **Wniosek: przy czytaniu cudzych plików faktyczny rozmiar wyświetlania musisz odczytać z `anchor.ext` i przeliczyć przez `EMU_to_pixels`.** To bardzo praktyczna wiedza — bez niej zbudujesz narzędzie, które zaniża/zawyża rozmiary obrazów w raportach klienta.

### Przykład 2 — obraz generowany w locie: matplotlib → bajty → arkusz (🟡)

Ten przykład realizuje wzorzec z 2.6 — i jest to **odpowiedź na wszystkie ograniczenia „chartex" z modułu 15**.

```python
"""Modul 16 - obraz generowany w locie: matplotlib -> PNG w pamieci -> arkusz.

To obejscie dla typow wykresow, ktorych openpyxl NIE umie rysowac
(modul 15): treemap, waterfall, histogram, Pareto, mapa, lejek...

Uruchom:  python examples/16_obraz_z_matplotlib.py
Wymaga:   pip install openpyxl pillow matplotlib
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.units import pixels_to_points

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "16_obraz_z_matplotlib.xlsx"

PALETA = {"akcent": "#5B5CF0", "drugi": "#F0705B", "trzeci": "#3BA55C"}


def _fig_to_bytes(fig, dpi: int) -> bytes:
    """Zamienia figure matplotlib na BAJTY PNG. Nic nie zapisuje na dysk."""
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    import matplotlib.pyplot as plt
    plt.close(fig)                  # KLUCZOWE: zwolnij pamiec (sekcja 2.6)
    return buf.getvalue()


def wykres_waterfall(labels, deltas, tytul: str, dpi: int = 120) -> bytes:
    """Wykres kaskadowy (waterfall) - openpyxl go NIE ma. Matplotlib ma."""
    import matplotlib
    matplotlib.use("Agg")            # backend bez GUI - dziala na serwerze
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(7.2, 3.6))

    narastajaco = 0
    for i, (label, delta) in enumerate(zip(labels, deltas)):
        kolor = PALETA["trzeci"] if delta >= 0 else PALETA["drugi"]
        ax.bar(i, delta, bottom=narastajaco, color=kolor, width=0.62)
        narastajaco += delta
    ax.plot(range(len(labels)), [0] * len(labels), color="none")
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=0)
    ax.set_title(tytul)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", alpha=0.25)
    return _fig_to_bytes(fig, dpi)


def wykres_treemap(etykiety, wartosci, tytul: str, dpi: int = 120) -> bytes:
    """Prosty treemap: rekurencyjne dzielenie prostokata. openpyxl NIE umie."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle

    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    sumaw = sum(wartosci)
    x = 0.0
    kolory = [PALETA["akcent"], PALETA["drugi"], PALETA["trzeci"], "#8B8CF5", "#F5A88B"]
    for i, (etykieta, wartosc) in enumerate(zip(etykiety, wartosci)):
        szer = wartosc / sumaw
        ax.add_patch(Rectangle((x, 0), szer, 1, facecolor=kolory[i % len(kolory)]))
        ax.text(x + szer / 2, 0.5, f"{etykieta}\n{wartosc}",
                ha="center", va="center", color="white", fontsize=9, weight="bold")
        x += szer
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.set_title(tytul)
    return _fig_to_bytes(fig, dpi)


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    ws["B2"] = "Obrazy generowane w locie (matplotlib → PNG → add_image)"
    ws["B4"] = "Wykres kaskadowy (waterfall) — niedostępny w openpyxl:"

    # --- 1. waterfall jako obraz ------------------------------------
    bajty_wf = wykres_waterfall(
        labels=["Start", "Q1", "Q2", "Q3", "Q4"],
        deltas=[100, 25, -40, 60, -15],
        tytul="Przepływ gotówki 2026 (tys. zł)",
    )
    img_wf = Image(BytesIO(bajty_wf))
    print(f"Waterfall: {img_wf.width} x {img_wf.height} px, "
          f"format={img_wf.format!r}")
    img_wf.anchor = "B5"
    ws.add_image(img_wf)
    ws.row_dimensions[5].height = pixels_to_points(img_wf.height)

    # --- 2. treemap jako obraz --------------------------------------
    ws["B26"] = "Treemap — też niedostępny w openpyxl:"
    bajty_tm = wykres_treemap(
        etykiety=["Polska", "Niemcy", "Czechy", "Słowacja", "Inne"],
        wartosci=[42, 28, 14, 9, 7],
        tytul="Udział rynków (%)",
    )
    img_tm = Image(BytesIO(bajty_tm))
    img_tm.anchor = "B27"
    ws.add_image(img_tm)
    ws.row_dimensions[27].height = pixels_to_points(img_tm.height)

    # --- 3. to samo zrodlo -> DWA arkusze (wzorzec z 2.5) -----------
    # Trzymamy BAJTY, nie obiekt Image. Kazdy arkusz dostaje SWIEZY Image.
    for nazwa in ("Wariant A", "Wariant B"):
        ws2 = wb.create_sheet(nazwa[:31])
        ws2["B2"] = f"Ten sam obraz w arkuszu: {nazwa}"
        ws2.add_image(Image(BytesIO(bajty_wf)), "B4")   # ← swiezy obiekt!
        ws2.row_dimensions[4].height = pixels_to_points(img_wf.height)

    wb.save(cel)
    wb.close()
    print(f"Zapisano: {cel}")


def main() -> None:
    try:
        import matplotlib  # noqa: F401
    except ImportError:
        print("Brak matplotlib. Zainstaluj:  pip install matplotlib")
        return
    zbuduj(CEL)
    print()
    print("=" * 78)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 78)
    print("  1. To sa OBRAZY, nie wykresy. Kliknij -> pasek formuly jest PUSTY,")
    print("     bo obraz nie odnosi sie do zadnych komorek.")
    print("  2. ZMIEN wartosc w komorce, ktora byla zrodlem danych (jesli")
    print("     wpiszesz ja obok) - obraz sie NIE zmieni. To migawka!")
    print("  3. Oba obrazy maja descr='Picture' - tekst alternatywny, ktorego")
    print("     openpyxl nie potrafi ustawic (sekcja 2.7).")
    print("  4. Przejdz na arkusze 'Wariant A' i 'Wariant B' - ten sam obraz")
    print("     w dwoch arkuszach, mimo ze copy_worksheet go NIE kopiuje")
    print("     (sekcja 2.12). Zostal dodany JAWNIE dla kazdego arkusza.")
    print()
    print("WAZNE - GDYBY ROBIC TEN SAM OBRAZ DLA WIELU SKOROSZYTOW:")
    print("  NIGDY nie uzywaj tego samego obiektu Image dla drugiego save()!")
    print("  _data() zamyka strumien -> ValueError: I/O operation on closed file")
    print("  Trzymaj BAJTY (jak wyzej) i tworz Image(BytesIO(bajty)) od nowa.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `wykres_waterfall` tworzy figurę matplotlib, zapisuje ją do `BytesIO`, **zamyka figurę** i zwraca **bajty**. Wszystko, co jest w pamięci, to teraz nieprzezroczysty `bytes`. Potem `Image(BytesIO(bajty))` — Pillow otwiera te bajty tylko po to, żeby poznać rozmiar i format (bo `BytesIO` nie jest `str`, więc `mark_to_close=False` — nic nie zostanie zamknięte **w konstruktorze**).

**Co trafi do pliku.** Trzy osobne kopie PNG w `xl/media/` (dwa razy waterfall dla dwóch arkuszy + treemap), trzy kotwice w dwóch plikach rysunków, relacje. Zwróć uwagę, że **ten sam obraz w dwóch arkuszach to dwa oddzielne pliki w `media/`** — żadna deduplikacja się nie dzieje. Analogia: **drukujesz tę samą ulotkę dwa razy i wkładasz do dwóch teczek**. Teczki są lżejsze, gdyby dzieliły jeden oryginał, ale openpyxl tak nie robi.

**Trzy rzeczy do przemyślenia:**

1. **Ten wzorzec jest uniwersalnym „bezpiecznikiem" na braki openpyxl.** Nie umiesz w openpyxl? Narysuj w `matplotlib` albo innym narzędziu i wstaw jako obraz. Traci się interaktywność (2.6), ale zyskuje **dowolność**. To dokładnie ten sam mechanizm, co „aplikacja webowa generuje PDF": narzędzie do prezentacji nie musi umieć wszystkiego, jeśli umiesz **wstrzyknąć gotowy obraz**.

2. **`plt.close(fig)` to nie kosmetyka.** Bez niego pętla po 200 raportach zjesz RAM. To ta sama lekcja co `wb.close()` w `read_only` (moduł 19): **zasoby trzeba zwalniać jawnie**.

3. **Trzy arkusze, jeden zbiór bajtów, trzy świeże `Image`.** To jest w praktyce test reguły z 2.5. Gdybyśmy napisali `img_wf` i dodali go do trzech arkuszy, **trzy zapisy tego samego pliku** — pierwszy przeszedłby, drugi by się wywalił. Zasada „bajty są źródłem prawdy, `Image` jest jednorazowy" uchroniła nas przed błędem, który pojawiłby się dopiero przy testach integracyjnych.

### Przykład 3 — hiperlinki: trzy rodzaje, styl, tooltip (🟡)

```python
"""Modul 16 - hiperlinki: zewnetrzne, wewnetrzne, mailto, plik, tooltip.

Uruchom:  python examples/16_hiperlinki.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import quote_sheetname
from openpyxl.worksheet.hyperlink import Hyperlink

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "16_hiperlinki.xlsx"


def zbuduj(cel: Path) -> None:
    wb = Workbook()

    # --- arkusz z danymi, do ktorego bedziemy linkowac -------------
    dane = wb.active
    dane.title = "Dane 2026"
    dane["A1"] = "Region"
    dane["B1"] = "Sprzedaż"
    for i, (region, wartosc) in enumerate(
        [("Północ", 120), ("Południe", 98), ("Wschód", 145), ("Zachód", 77)], start=2
    ):
        dane[f"A{i}"] = region
        dane[f"B{i}"] = wartosc

    # --- arkusz "Spis" z linkami -----------------------------------
    ws = wb.create_sheet("Spis", 0)         # jako pierwszy
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 46

    ws["A1"] = "Rodzaj linku"
    ws["B1"] = "Kliknij"
    for komorka in ("A1", "B1"):
        ws[komorka].font = __import__("openpyxl").styles.Font(bold=True)

    wiersz = 3

    # ================= 1. LINK ZEWNETRZNY (https) ==================
    # KOLEJNOSC: napis NAJPIERW, potem hyperlink (sekcja 2.8!)
    ws.cell(row=wiersz, column=1, value="Zewnętrzny (https)")
    ws.cell(row=wiersz, column=2, value="Dokumentacja openpyxl").hyperlink = (
        "https://openpyxl.readthedocs.io/"
    )
    ws.cell(row=wiersz, column=2).style = "Hyperlink"
    wiersz += 1

    # ================= 2. MAILTO ===================================
    ws.cell(row=wiersz, column=1, value="E-mail (mailto)")
    ws.cell(row=wiersz, column=2, value="Napisz do zespołu raportów").hyperlink = (
        "mailto:raporty@example.com?subject=Raport%202026"
    )
    ws.cell(row=wiersz, column=2).style = "Hyperlink"
    wiersz += 1

    # ================= 3. LINK DO PLIKU (URI!) =====================
    ws.cell(row=wiersz, column=1, value="Plik lokalny (URI)")
    ws.cell(row=wiersz, column=2, value="Otwórz plik pomocniczy").hyperlink = (
        "file:///C:/raporty/slownik.xlsx"       # UWAGA: URI, nie C:\\...
    )
    ws.cell(row=wiersz, column=2).style = "Hyperlink"
    wiersz += 1

    # ================= 4. LINK WEWNETRZNY (do arkusza) =============
    # TU JEST POPRAWKA: uzywamy location, NIE "#Dane 2026!A1"!
    ws.cell(row=wiersz, column=1, value="Wewnętrzny (inny arkusz)")
    ws.cell(row=wiersz, column=2).value = "Przejdź do 'Dane 2026'!A2"
    ws.cell(row=wiersz, column=2).hyperlink = Hyperlink(
        ref="B6",                                        # nadpisane, ale wymagane
        location=f"{quote_sheetname('Dane 2026')}!A2",   # → "'Dane 2026'!A2"
        tooltip="Skok do pierwszego rekordu danych",
    )
    ws.cell(row=wiersz, column=2).style = "Hyperlink"
    wiersz += 1

    # ================= 5. LINK WEWNETRZNY (ten sam arkusz) =========
    ws.cell(row=wiersz, column=1, value="Wewnętrzny (ten arkusz)")
    ws.cell(row=wiersz, column=2).value = "Wróć na początek (A1)"
    ws.cell(row=wiersz, column=2).hyperlink = Hyperlink(ref="B7", location="A1")
    wiersz += 1

    # ================= 6. LINK Z TOOLTIPEM =========================
    ws.cell(row=wiersz, column=1, value="Z tooltipem")
    ws.cell(row=wiersz, column=2).value = "Najedź i poczytaj"
    ws.cell(row=wiersz, column=2).hyperlink = Hyperlink(
        ref="B8",
        target="https://example.com/raport",
        tooltip="Otworzy się w nowej karcie przeglądarki",
    )
    ws.cell(row=wiersz, column=2).style = "Hyperlink"
    wiersz += 2

    # ================= DEMONSTRACJA PULAPKI ========================
    ws.cell(row=wiersz, column=1, value="⚠️ PUŁAPKA: link na PUSTEJ komórce")
    komorka = ws.cell(row=wiersz, column=2)
    komorka.hyperlink = "https://example.com/auto"     # komorka byla PUSTA!
    wiersz += 1
    ws.cell(row=wiersz, column=1, value="   ↑ co openpyxl wpisał jako napis?")
    ws.cell(row=wiersz, column=2, value="(sprawdź obok w kodzie weryfikacji)")

    wb.save(cel)
    wb.close()
    print(f"Zapisano: {cel}")

    # pokaz pulapki od razu:
    print()
    print("PUŁAPKA 'auto-value': po przypisaniu hyperlinku do pustej komórki")
    print("openpyxl WPISAŁ do niej adres jako napis:")
    print(f"  wartość komórki B{wiersz - 2} = "
          f"{komorka.value!r}")


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Spis"]
        print()
        print("=" * 96)
        print("INWENTARZ HIPERLINKOW (arkusz 'Spis')")
        print("=" * 96)
        print(f"{'Kom.':<6}{'Napis':<34}{'target':<38}{'location':<22}")
        print("-" * 96)

        for row in ws.iter_rows():
            for c in row:
                if c.hyperlink is None:
                    continue
                h = c.hyperlink
                napis = (str(c.value) or "")[:32]
                target = (h.target or "-")[:36]
                location = (h.location or "-")[:20]
                print(f"{c.coordinate:<6}{napis:<34}{target:<38}{location:<22}")
    finally:
        wb.close()

    print()
    print("=" * 96)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 96)
    print("  1. Kliknij każdy link i sprawdź, dokąd prowadzi.")
    print("  2. Link WEWNĘTRZNY (wiersz 6) przenosi do arkusza 'Dane 2026'.")
    print("     NIE zamyka pliku i nie pokazuje ostrzeżenia o zewnętrznym")
    print("     źródle - bo to nawigacja w skoroszycie (atrybut location).")
    print("  3. Kliknij link w wierszu 4 (plik lokalny). Jeśli ścieżka nie")
    print("     istnieje, Excel pokaże błąd.")
    print("  4. Najedź na link z tooltipem - zobaczysz podpowiedź.")
    print("  5. Sprawdź, że po kliknięciu link ma styl 'Followed Hyperlink'")
    print("     (Excel sam zmienia kolor po pierwszym użyciu).")


def main() -> None:
    zbuduj(CEL)
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Dla każdego linku z `target` openpyxl trzyma obiekt `Hyperlink` w `cell._hyperlink`. **Relacja jeszcze nie istnieje** — powstanie dopiero w `write_hyperlinks()` przy zapisie. Dla linków z `location` żadna relacja nie powstanie wcale.

**Co trafi do pliku.** I tu najciekawsze — zobacz, jak **ten sam plik wygląda inaczej** w zależności od rodzaju linku:

```xml
<!-- xl/worksheets/_rels/sheet1.xml.rels  (tylko linki ZEWNETRZE) -->
<Relationship Id="rId1" Type=".../hyperlink"
              Target="https://openpyxl.readthedocs.io/" TargetMode="External"/>
<Relationship Id="rId2" Type=".../hyperlink"
              Target="mailto:raporty@example.com?subject=Raport%202026"
              TargetMode="External"/>

<!-- xl/worksheets/sheet1.xml  (WSZYSTKIE linki, ale roznie opisane) -->
<hyperlinks>
  <hyperlink ref="B3" r:id="rId1"/>                         <!-- zewnetrzny -->
  <hyperlink ref="B4" r:id="rId2"/>                         <!-- zewnetrzny -->
  <hyperlink ref="B6" location="'Dane 2026'!A2"
             tooltip="Skok do pierwszego rekordu danych"/>   <!-- WEWNETRZNY! -->
  <hyperlink ref="B7" location="A1"/>                       <!-- WEWNETRZNY! -->
</hyperlinks>
```

**Trzy rzeczy do przemyślenia:**

1. **Linki wewnętrzne nie mają `r:id`.** Nie ma go, bo nie ma relacji. Ich `location` to zwykły atrybut — **dlatego prefiks `#` byłby błędem**: zamiast `location` trafiłby do `target`, powstałaby relacja `External`, a Excel szukałby pliku `#Dane 2026!A2` poza skoroszytem. To dokładnie ta poprawka z 2.8, teraz widoczna w XML-u.

2. **`ref` zawsze wskazuje komórkę.** Zwróć uwagę, że w kodzie podałem `ref="B6"` w konstruktorze — a setter i tak nadpisał to prawdziwym koordynatem. To nieużyteczny parametr, ale **wymagany technicznie** (deskryptor `String()`).

3. **Napis i link to dwie osobne rzeczy** — i widać to także tutaj: `r:id`/`location` siedzi w `sheet1.xml`, a **napis** komórki też. Ale pochodzą z dwóch różnych miejsc w kodzie (`c.value` vs `c.hyperlink`) i mogą się rozjechać, gdy zapomnisz kolejności.

### Przykład 4 — komentarze: wskazówki dla użytkownika (🟡)

```python
"""Modul 16 - komentarze (notatki): tresc, autor, rozmiar, scalone komorki.

Uruchom:  python examples/16_komentarze.py
"""

from __future__ import annotations

import warnings
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import units

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "16_komentarze.xlsx"


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Formularz"

    # --- naglowki --------------------------------------------------
    ws["A1"] = "Pole"
    ws["B1"] = "Wartość"
    ws["C1"] = "Uwagi"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 34

    # --- 1. PROSTY komentarz ------------------------------------
    ws["A3"] = "Kwota netto"
    komentarz = Comment(
        "Wpisz kwotę w groszach jako liczbę całkowitą (np. 12345 = 123,45 zł). "
        "Nie używaj przecinków ani spacji — ułatwia to import do systemu.",
        "Zespół raportowy",
    )
    # domyslny rozmiar to 144 x 79 px - za maly na to zdanie. Ustawiamy wiekszy.
    komentarz.width = 360
    komentarz.height = 120
    ws["B3"].comment = komentarz

    # --- 2. komentarz z rozmiarem w PUNKTACH (typografia!) -------
    ws["A4"] = "Data wystawienia"
    komentarz2 = Comment("Format RRRR-MM-DD. Puste pole = data dzisiejsza.",
                         "Zespół raportowy")
    komentarz2.width = units.points_to_pixels(240)      # punkty -> piksele
    komentarz2.height = units.points_to_pixels(60)
    ws["B4"].comment = komentarz2

    # --- 3. TEN SAM komentarz w DWOCH komorkach = KOPIA ----------
    wspolny = Comment("Pole wymagane", "Walidadcja")
    ws["C3"].comment = wspolny
    ws["C4"].comment = wspolny
    print("Czy C3.comment to TA SAMA zmienna co C4? ",
          ws["C3"].comment is wspolny, ws["C4"].comment is wspolny)
    print("  (pierwsza True, druga False - openpyxl robi KOPIE, sekcja 1.4)")

    # --- 4. PULAPKA: komentarz w SCALONEJ komorce ----------------
    ws["A6"] = "Scalony nagłówek"
    ws.merge_cells("A6:C6")

    # 4a. LEWY GORNY rog scalonego zakresu -> DZIALA
    ws["A6"].comment = Comment("Ten komentarz przetrwa (lewy górny róg).",
                               "Zespół raportowy")

    # 4b. POD-KOMORKA scalenia -> AttributeError
    ws["A7"] = "To zostanie scalone"
    ws.merge_cells("A7:C7")
    try:
        ws["B7"].comment = Comment("Ten komentarz NIE zadziala.", "Zespół")
        print("B7: udalo sie (nieoczekiwane!)")
    except AttributeError as e:
        print(f"B7: AttributeError -> {e}")
        print("    (MergedCell nie ma atrybutu 'comment' - sekcja 2.10)")

    ws["A9"] = "Komentarze w C3/C4 są IDENTYCZNE (kopia), a w B3/B4 różne."
    ws["A10"] = "Najedź na B3, B4, C3, C4 i sprawdź rozmiary okienek."

    wb.save(cel)
    wb.close()
    print(f"\nZapisano: {cel}")


def weryfikuj(cel: Path) -> None:
    """Wczytuje plik i wypisuje WSZYSTKIE komentarze + nasluchuje ostrzezen."""
    print()
    print("=" * 88)
    print("WCZYTYWANIE Z NASLUCHEM OSTRZEZEN (sekcja 2.10)")
    print("=" * 88)

    # Ostrzezenie o komentarzu w scalonej komorce to warnings.warn,
    # NIE wyjatek. Zeby je zobaczyc, trzeba je PRZECHWYCIC.
    with warnings.catch_warnings(record=True) as zlapane:
        warnings.simplefilter("always")

        wb = load_workbook(cel)
        try:
            ws = wb["Formularz"]
            print(f"Komentarzy z licznikiem arkusza: "
                  f"{len(ws._comments)}   (atrybut prywatny)")
            print()
            # UWAGA: przy czytaniu ROZMIARY sa tracone (dokumentacja, 2.9)
            for row in ws.iter_rows():
                for c in row:
                    if c.comment is None:
                        continue
                    k = c.comment
                    print(f"{c.coordinate:<5} autor={k.author!r:<20} "
                          f"rozmiar={k.width}x{k.height} szer/wys")
                    print(f"      treść: {k.text[:64]}...")
        finally:
            wb.close()

        if zlapane:
            print()
            print("OSTRZEZENIA PRZY WCZYTANIU:")
            for w in zlapane:
                print("  ⚠️ ", w.message)
        else:
            print()
            print("Brak ostrzezen przy wczytaniu tego pliku (nie bylo komentarzy")
            print("w scalonych pod-komorkach - te nie daly sie zapisac).")

    print()
    print("=" * 88)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 88)
    print("  1. Najedź na B3 - okienko z instrukcją (360x120 px).")
    print("  2. Najedź na B4 - Mniejsze okienko (z punktów, ~320x80 px).")
    print("  3. Najedź na C3 i C4 - identyczna treść, ale to DWIE komórki")
    print("     z DWOMA niezależnymi komentarzami.")
    print("  4. ZMIEN treść komentarza w C3 w Excelu - C4 się NIE zmieni.")
    print("     Bo to kopia, nie referencja.")
    print()
    print("EKSPERYMENT (obrazuje utratę danych):")
    print("  1. W Excelu dodaj notatkę do komórki w środku SCALONEGO zakresu")
    print("     (np. do B7). Excel to pozwoli.")
    print("  2. Zapisz plik.")
    print("  3. Wczytaj go ponownie openpyxl i ZOBACZ OSTRZEŻENIE:")
    print("     '...is part of a merged range but has a comment which will")
    print("     be removed because merged cells cannot contain any data.'")
    print("  4. Zapisz i otwórz w Excelu - komentarza NIE MA.")


def main() -> None:
    zbuduj(CEL)
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `Comment` to zwykły obiekt z `content`, `author`, `height`, `width`. Po `ws["B3"].comment = komentarz` openpyxl zapamiętuje go w `cell._comment` i dodaje do `ws._comments`. W trzecim przypadku (wspólny komentarz w dwóch komórkach) wywoływany jest `__copy__` — **i to jest cała magia „automatycznych kopii"** z 1.4. Widzisz to w wydruku: pierwsze `is` daje `True`, drugie `False`.

**Co trafi do pliku.** **Trzy części**, nie jedna:

- `xl/comments1.xml` — treść (`<comment ref="B3" authorId=".."><text><t>...</t></text></comment>`),
- `xl/vmlDrawing1.vml` — **kształt i pozycja okienka** (tu żyją `width` i `height`; dlatego dokumentacja mówi „dimensions are lost upon reading" — openpyxl czyta tylko `comments1.xml`),
- `xl/drawings/` + relacje — spinanie tego razem.

**Trzy rzeczy do przemyślenia:**

1. **Utrata wymiarów przy odczycie to nieuprzejmość, a konsekwencja architektury.** Wymiary są w VML, a openpyxl VML nie parsuje. Gdybyś chciał to obejść, to dokładnie ten sam typ zadania co tekst alternatywny obrazu (2.7) — ingerencja w XML. **Nie rób tego, chyba że musisz.**

2. **`ws._comments` to atrybut prywatny.** Używaj go do diagnostyki (jak `ws._images`). Do normalnej pracy posługuj się `cell.comment`.

3. **`warnings.catch_warnings(record=True)` to Twoje narzędzie w module 18.** Ten przykład pokazuje, że **openpyxl komunikuje utratę danych przez ostreżenia**, a nie wyjątki. Jeśli ich nie przechwycisz — znikną w konsoli i nigdy się o nich nie dowiesz. W module 18 przerobisz to na pełny **raport ryzyka** pliku wejściowego.

### Przykład 5 — inwentarz: obrazy, hiperlinki, komentarze w cudzym pliku (🔴)

To narzędzie do uruchomienia na **każdym** pliku, który dostałeś od kogoś i zamierzasz przepuścić przez openpyxl.

```python
"""Modul 16 - inwentarz: co openpyxl WIDZI w cudzym pliku?

Uruchom:  python examples/16_inwentarz.py [sciezka.xlsx]

Porownaj wynik z tym, co widzisz w Excelu. Roznice = utrata danych.
"""

from __future__ import annotations

import sys
import warnings
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.units import EMU_to_pixels, EMU_to_cm

ROOT = Path(__file__).resolve().parent.parent
DOMYSLNY = ROOT / "output" / "16_obraz_podstawy.xlsx"


def opis_anchora(img) -> str:
    """Bezpiecznie opisuje kotwice obrazu bez importowania typow prywatnych."""
    a = img.anchor
    if isinstance(a, str):
        return f"kotwica tekstowa: {a}"

    if hasattr(a, "_from") and hasattr(a, "to"):
        f, t = a._from, a.to
        return (f"TwoCellAnchor: (c{f.col},r{f.row}) -> (c{t.col},r{t.row}) "
                f"editAs={getattr(a, 'editAs', '?')}")
    if hasattr(a, "_from"):
        ext = getattr(a, "ext", None)
        rozmiar = ""
        if ext is not None:
            px = (EMU_to_pixels(ext.cx), EMU_to_pixels(ext.cy))
            cm = (EMU_to_cm(ext.cx), EMU_to_cm(ext.cy))
            rozmiar = f" | wyswietlane {px[0]}x{px[1]} px ({cm[0]}x{cm[1]} cm)"
        return (f"OneCellAnchor od (c{a._from.col}, r{a._from.row})  "
                f"<- 0-based!{rozmiar}")
    return f"{type(a).__name__}"


def inwentarz(sciezka: Path) -> None:
    print("=" * 100)
    print(f"INWENTARZ: {sciezka.name}")
    print("=" * 100)

    # NASLUCH OSTRZEZEN - openpyxl w ten sposob komunikuje utrate danych!
    with warnings.catch_warnings(record=True) as zlapane:
        warnings.simplefilter("always")

        # KRYTYCZNE: NIE read_only - tam obrazy i komentarze nie istnieja (2.11)
        wb = load_workbook(sciezka)
        try:
            for ws in wb.worksheets:
                obrazy = list(ws._images)
                # komentarze i hiperlinki: iterujemy po komorkach
                komentarze = []
                linki = []
                for row in ws.iter_rows():
                    for c in row:
                        if c.comment is not None:
                            komentarze.append(c)
                        if c.hyperlink is not None:
                            linki.append(c)

                if not (obrazy or komentarze or linki):
                    continue

                print(f"\nARKUSZ '{ws.title}'")
                print(f"  obrazy: {len(obrazy)} | komentarze: {len(komentarze)} "
                      f"| hiperlinki: {len(linki)}")

                # ---------- OBRAZY -------------------------------------
                for i, img in enumerate(obrazy, 1):
                    print(f"    [O{i}] {img.path}  format={img.format}  "
                          f"zrodlo={img.width}x{img.height} px")
                    print(f"         {opis_anchora(img)}")

                # ---------- KOMENTARZE ---------------------------------
                for i, c in enumerate(komentarze, 1):
                    k = c.comment
                    print(f"    [K{i}] {c.coordinate}  autor={k.author!r}  "
                          f"rozmiar={k.width}x{k.height}")
                    print(f"         treść: {k.text[:70]!r}")

                # ---------- HIPERLINKI ---------------------------------
                for i, c in enumerate(linki, 1):
                    h = c.hyperlink
                    rodzaj = "WEWNETRZNY" if h.target is None else "zewnetrzny"
                    print(f"    [H{i}] {c.coordinate}  [{rodzaj}]  "
                          f"target={h.target!r}  location={h.location!r}  "
                          f"tooltip={h.tooltip!r}")
        finally:
            wb.close()

        # ---------- OSTRZEZENIA ------------------------------------
        print()
        print("=" * 100)
        print("OSTRZEZENIA openpyxl przy wczytaniu")
        print("=" * 100)
        if zlapane:
            for w in zlapane:
                print(f"  ⚠️  {w.category.__name__}: {w.message}")
            print()
            print("  KAŻDE ostrzeżenie = potencjalna utrata danych przy zapisie!")
        else:
            print("  Brak. Dobra wiadomość - ale to jeszcze nie gwarancja")
            print("  (np. sparkline'y i kształty nie generują ostrzeżeń - modul 18).")

    print()
    print("=" * 100)
    print("CHECKLISTA ROZBIEZNOSCI (porownaj z tym, co widzisz w Excelu)")
    print("=" * 100)
    print("  [ ] Czy liczba OBRAZÓW się zgadza?")
    print("      Nie -> sprawdź: czy masz pillow? Czy obraz jest WMF?")
    print("      Czy obraz nie jest 'spoza pakietu' (linked, nie embedded)?")
    print("  [ ] Czy obraz ma ZNACZĄCY tekst alternatywny w Excelu?")
    print("      Jeśli tak -> openpyxl go ZGUBI (wpisze 'Picture', sekcja 2.7).")
    print("  [ ] Czy liczba KOMENTARZY się zgadza?")
    print("      Nie -> sprawdź ostrzeżenia wyżej (scalone komórki, 2.10).")
    print("      Komentarze wątkowe (Excel 365) też mogą nie być widoczne.")
    print("  [ ] Czy komentarze mają niestandardowe formatowanie tekstu?")
    print("      Jeśli tak -> będzie zgubione (2.9).")
    print("  [ ] Czy HIPERLINKI się zgadzają?")
    print("      Ta kategoria jest solidna, ale sprawdź linki do plików")
    print("      (ścieżki mogą być nieaktualne) i wewnętrzne (location).")
    print("  [ ] ZAPISZ plik pod nową nazwą i uruchom ten inwentarz PONOWNIE.")
    print("      Porównanie PRZED i PO to najtańszy test utraty danych.")


def main() -> None:
    sciezka = Path(sys.argv[1]) if len(sys.argv) > 1 else DOMYSLNY
    if not sciezka.exists():
        print(f"Brak pliku: {sciezka}")
        print("Uruchom najpierw: python examples/16_obraz_podstawy.py")
        return
    inwentarz(sciezka)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `load_workbook` bez `read_only` **wczytuje całe rysunki** (kotwice + bajty obrazów → obiekty `Image`) i **całą część komentarzy**. Przy dużym pliku to realny koszt pamięci — to cena kompletności (moduł 03, moduł 19).

**Co trafi do pliku.** Nic — to narzędzie tylko czyta. Ale jego **wartość** jest w porównaniu: uruchom je na pliku oryginalnym i po przepuszczeniu przez openpyxl. **Różnica w liczbach to Twoja utrata danych.**

**Trzy rzeczy do przemyślenia:**

1. **`warnings.catch_warnings(record=True)` to najważniejsza linia tego przykładu.** openpyxl komunikuje utratę danych ostrzeżeniami — a domyślnie Python pokazuje każde ostrzeżenie **tylko raz na sesję** (filtruje duplikaty!). `simplefilter("always")` wyłącza to filtrowanie. Bez tego przy dwóch takich samych ostrzeżeniach zobaczysz tylko pierwsze. To subtelna, ale krytyczna pułapka.

2. **`opis_anchora` używa `hasattr`, nie importu typów prywatnych.** To celowe: `_AnchorBase` zaczyna się od podkreślnika, więc importowanie go byłoby zależnością od API wewnętrznego. `hasattr(a, "to")` odróżnia `TwoCellAnchor` od `OneCellAnchor` bez tego.

3. **`img.width` w inwentarzu pokazuje źródło, nie skalę.** Ten sam szczegół, o którym mówiłem w Przykładzie 1 — i dlatego `opis_anchora` **osobno** wypisuje `wyswietlane` z `anchor.ext`. Gdyby narzędzie pokazywało tylko `img.width`, oszukałoby Cię przy każdym przeskalowanym obrazie.

## 4. Anatomia API

| Klasa / metoda / atrybut | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Image(img)` | Obiekt obrazu | `img`: ścieżka `str`, `PIL.Image` albo strumień | **Wymaga `pillow`**; `ImportError` bez niego |
| `img.width`, `img.height` | Rozmiar obrazu | `int`/`float` | **PIKSELE** (z `PIL.size`). Brak `cm_to_pixels` — policz sam |
| `img.format` | Rozszerzenie obrazu | tylko odczyt | `'png'`, `'jpeg'`, `'gif'`; inne → konwersja do PNG |
| `img.ref` | Źródło obrazu | — | Ścieżka, obiekt PIL albo strumień. **Otwierane ponownie przy `save()`** |
| `img.anchor` | Pozycja obrazu | `str` („B2") albo obiekt anchora | Domyślnie `"A1"`! |
| `img.path` | Ścieżka w archiwum | tylko odczyt | `/xl/media/image{id}.{format}` |
| `ws.add_image(img, anchor=None)` | Dodaje obraz do arkusza | `anchor`: komórka | Tylko dopisuje do `ws._images`; geometria przy zapisie |
| `ws._images` | Lista obrazów | — | **Prywatne** — do diagnostyki |
| `OneCellAnchor(_from, ext)` | Kotwica: komórka + stały rozmiar | `AnchorMarker`, `XDRPositiveSize2D` | Tworzona domyślnie |
| `TwoCellAnchor(editAs, _from, to)` | Kotwica: dwie komórki | `AnchorMarker` ×2 | Obraz **rośnie z komórkami** |
| `AbsoluteAnchor(pos, ext)` | Kotwica absolutna | `XDRPoint2D`, `XDRPositiveSize2D` | Bez komórek |
| `AnchorMarker(col, colOff, row, rowOff)` | Współrzędne kotwicy | wszystkie `int` | **0-BASED!** Domyślnie `col=0, row=0` |
| `XDRPositiveSize2D(cx, cy)` | Rozmiar w EMU | `cx`, `cy` w EMU | `cx = pixels_to_EMU(px)` |
| `units.pixels_to_EMU(px)` | px → EMU | — | $1$ px $= 9525$ EMU |
| `units.EMU_to_pixels(emu)` | EMU → px | — | `round(v / 9525)` |
| `units.points_to_pixels(pt)` | pt → px | `dpi=96` | Do rozmiarów komentarzy i wierszy |
| `units.pixels_to_points(px)` | px → pt | `dpi=96` | Do `row_dimensions[h].height` |
| `Comment(text, author, height=79, width=144)` | Komentarz | `text`, `author` obowiązkowe | Domyślnie **144 × 79 px** — za małe na zdanie |
| `comment.text` / `comment.content` | Treść | aliasy | `text` czytelniejszy, `content` w XML-u |
| `comment.author` | Autor | `str` | Wymagany |
| `comment.width`, `comment.height` | Rozmiar okienka | `int` (**piksele**) | **Zapisywane, ale tracone przy odczycie** |
| `cell.comment` | Komentarz komórki | `Comment` albo `None` | Ten sam komentarz w 2 komórkach → **kopia** |
| `ws._comments` | Lista komentarzy | — | **Prywatne** — do diagnostyki |
| `Hyperlink(ref, location, tooltip, display, id, target)` | Hiperlink | `ref` wymagane technicznie | `target` = zewnętrzny, `location` = wewnętrzny |
| `cell.hyperlink = "https://..."` | Link zewnętrzny | `str` albo `Hyperlink` | `str` → `Hyperlink(target=...)` |
| `cell.hyperlink = Hyperlink(location=...)` | Link **wewnętrzny** | — | **NIE używaj prefiksu `#`!** |
| `cell.hyperlink = None` | Usuwa link | — | — |
| `cell.hyperlink.target` | Adres zewnętrzny | `str` albo `None` | `None` ⇒ link wewnętrzny |
| `cell.hyperlink.location` | Cel wewnętrzny | `str` albo `None` | `'Arkusz'!A1` albo nazwa zdefiniowana |
| `cell.hyperlink.tooltip` | Podpowiedź | `str` | Pokazywana przy najechaniu |
| `cell.style = "Hyperlink"` | Styl linku | wbudowany styl nazwany | **Nadpisuje cały styl komórki** (moduł 09) |
| `quote_sheetname(nazwa)` | Cytuje nazwę arkusza | — | `'Dane 2026'` — wstaw do `location` |
| `ws.copy_worksheet(ws)` | Kopiuje arkusz | — | **Kopiuje komentarze i hiperlinki, NIE obrazy** |
| `load_workbook(..., read_only=True)` | Tryb leniwy | — | **Obrazy i komentarze NIEDOSTĘPNE** |
| `Workbook(write_only=True)` | Tryb strumieniowy | — | Obrazy i wykresy **działają** |
| `wb.create_chartsheet(...)` | Arkusz wykresów | — | **Brak `add_image`** (`AttributeError`) |
| `warnings.catch_warnings(record=True)` | Przechwytywanie ostrzeżeń | — | Dodaj `simplefilter("always")` |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — logo i kalibracja rozmiaru.**

Napisz `examples/16_cw1.py`, który tworzy `output/16_cw1.xlsx`:

1. Arkusz `Raport` z nagłówkiem `"Raport kwartalny"` (pogrubiony, rozmiar 14) w `B2`.
2. Wygeneruj prosty obraz PNG za pomocą Pillow ($400 \times 160$ px) — np. kolorowy prostokąt z napisem.
3. Wstaw **trzy wersje tego samego obrazu**, każdą o innej szerokości:
   - `1,5` cm w `B5` (podpis: `"Mała"`),
   - `4` cm w `B8` (podpis: `"Średnia"`),
   - `8` cm w `B11` (podpis: `"Duża"`).
4. **Za każdym razem licz wysokość z proporcji** (użyj funkcji `cm_to_pixels` z Przykładu 1).
5. **Ustaw wysokość wiersza** pod każdą wersją, tak aby obraz się mieścił — pamiętaj, że `row_dimensions[].height` jest w **punktach** (`pixels_to_points`).
6. W czwartej wersji w `B14` **celowo zniekształć** obraz (ustaw `width` i `height` niezależnie) i podpisz `"Zła (zniekształcona)"`.

Dodaj funkcję `weryfikuj()`, która wczytuje plik i dla każdego obrazu wypisuje: `img.path`, źródłowe `img.width/height` oraz **wyświetlany** rozmiar z `anchor.ext` przeliczony przez `EMU_to_pixels`.

W pliku `output/16_cw1_wnioski.md`:

1. **Ile pikseli ma obraz o szerokości $4$ cm przy $96$ DPI?** Policz to ręcznie i porównaj z kodem.
2. **Dlaczego `img.width` po ponownym wczytaniu pokazuje inny rozmiar niż ten, który ustawiłeś?** Gdzie jest faktyczna skala i jak ją odczytać?
3. **Co się dzieje w pamięci i w pliku w momencie `ws.add_image(img, "B5")`?** Czy bajty obrazu są już wtedy wczytane (2.2, 2.5)?
4. **Otwórz `output/16_cw1.xlsx` jako ZIP i znajdź `xl/drawings/drawing1.xml`.** Ile jest tam kotwic? Jaka jest wartość `descr`? Przelicz jedną wartość `cx` na piksele i sprawdź, czy zgadza się z Twoim `width`.
5. **Co się stanie, gdy usuniesz podpis `"Duża"` z `B11` (albo wstawisz tam inny tekst)?** Czy obraz zostanie? Uzasadnij, odnosząc się do modelu „pinezki" z 1.3.

### 🟡 Warsztat

**Zadanie 2 — obraz generowany w locie (matplotlib → PNG → arkusz).**

Napisz `examples/16_cw2.py`, który tworzy `output/16_cw2.xlsx` z dashboardem złożonym z **obrazów**, nie natywnych wykresów:

1. Trzy obrazy wygenerowane w `matplotlib` (backend `Agg`), każdy jako **bajty** (`BytesIO`, `fig.savefig(...)`, `plt.close(fig)`):
   - **histogram** rozkładu jakichś danych (np. liczb losowych) — *openpyxl nie ma histogramu*,
   - **wykres Pareto** (słupki + narastająco linia na drugiej osi) — *openpyxl nie ma Pareto*,
   - **wykres kołowy w wersji matplotlib** — dla porównania z natywnym `PieChart` z modułu 15.
2. Każdy obraz wstaw na arkusz `Dashboard` w odstępach kolumnowych, z podpisem (zwykły tekst w komórce **obok**, nie tylko opcjonalnie).
3. **Ustaw wysokości wierszy** pod obrazy, aby się mieściły.
4. Ten sam histogram wstaw **także do drugiego arkusza** `Załącznik` — z zachowaniem zasady **świeżego `Image` per użycie** (2.5).
5. Dodaj `funkcję sprawdz()` która weryfikuje, że wszystkie obrazy są na swoich kotwicach i wypisuje ostrzeżenie, jeśli któryś zaczyna się w `A1` (bo tam startuje domyślna kotwica).

W pliku `output/16_cw2_wnioski.md`:

1. **Dlaczego `matplotlib.use("Agg")` jest konieczne przy generowaniu na serwerze?** Co by się działo bez tego?
2. **Dlaczego `plt.close(fig)` jest istotne przy pętli generującej 200 raportów?** Policz, ile pamięci zużyłbyś bez tego (przyjmij $\approx 6$ MB na figurę).
3. **Czym różni się ten dashboard od dashboardu z Ćwiczenia 3 w module 15?** Wymień trzy różnice w zachowaniu w Excelu (żywotność, interaktywność, rozmiar pliku).
4. **Dlaczego funkcja zwraca `bytes`, a nie obiekt `Image`?** Odnieś się do pułapki zamkniętego strumienia (2.5).
5. **Znajdź w Excelu funkcję, której openpyxl nie potrafi, a którą mogłeś wygenerować jako obraz.** Wypisz trzy przykłady i uzasadnij, dlaczego „obraz + podpis" jest tu rozsądną odpowiedzią, a „natywny wykres" nie.

### 🔴 Wyzwanie

**Zadanie 3 — hiperlinki nawigacyjne, przypisy i dostępność.**

To zadanie z briefu, doprowadzone do końca. Zbuduj **spis treści** raportu wieloarkuszowego z pełną nawigacją.

```python
"""Raport wieloarkuszowy ze spisem tresci, nawigacja i przypisami.

Wymagania:
  1. Arkusze: "Spis", "Dane 2026", "Podsumowanie", "Metodyka".

  2. NAWIGACJA (arkusz "Spis"):
     - do KAZDEGO arkusza link wewnetrzny (Hyperlink + location),
     - z KAZDEGO arkusza link "Powrot do spisu" na gorze,
     - z arkusza "Dane 2026" link "Zobacz podsumowanie" w prawej kolumnie.

  3. HIPERLINKI ZEWNETRZNE:
     - z arkusza "Metodyka" link do dokumentacji HTTP,
     - link mailto do wlasciciela raportu,
     - link do pliku lokalnego (URI, NIE sciezka Windows).

  4. PRZYPISY (komentarze) jako wskazowki interpretacyjne:
     - do kazdej kolumny z liczbami komentarz z jednostka i formatem,
     - rozmiar okienka DOSTOSOWANY do dlugosci tekstu (nie domyslny!),
     - jeden komentarz wspoldzielony dla kolumn o tym samym znaczeniu,
     - komentarz do scalonego naglowka - TYLKO w lewym gornym rogu.

  5. STRUKTURA:
     - funkcje jednoczynnosciowe (wzorzec z 2.13):
       naglowek_raportu(), ustaw_kolumny(), dodaj_spis(),
       dodaj_link_powrotu(), dodaj_przypisy_kolumn()
     - ZADNA funkcja nie przyjmuje typu openpyxl w sygnaturze tam,
       gdzie mozna przyjac prostszy typ (napis, lista, dict)

  6. WERYFIKACJA:
     - inwentarz linkow (rodzaj, target, location, tooltip)
     - sprawdzenie, ze KAZDY arkusz ma link powrotu
     - ostrzezenie, jesli jakas komorka hiperlinku ma pusta wartosc

  7. DOSTEPNOSC:
     - dodaj arkusz tekstowy "Opis dostepnosci" opisujacy tryb pracy,
     - zapisz w nim, czego openpyxl NIE potrafi (alt text obrazow)
"""

# Szkic sygnatur - trzymaj sie ich:
def naglowek_raportu(ws, tytul: str, podtytul: str, ostatnia_kol: int) -> None: ...
def ustaw_kolumny(ws, szerokosci: dict[str, float]) -> None: ...
def dodaj_spis(ws, arkusze: list[str]) -> None: ...
def dodaj_link_powrotu(ws, docelowy_arkusz: str) -> None: ...
def dodaj_link_z_napisem(ws, komorka: str, napis: str,
                         target: str | None = None,
                         location: str | None = None,
                         tooltip: str | None = None) -> None: ...
def dodaj_przypis(ws, komorka: str, tekst: str, autor: str) -> None: ...
def inwentarz_linkow(sciezka) -> None: ...
```

**Wymagania szczegółowe:**

1. **`dodaj_link_z_napisem` musi wymuszać kolejność „napis → link".** To jest wymóg bezpieczeństwa z 2.8 — funkcja ma **wypisać ostrzeżenie**, jeśli wywołano ją na komórce, która już ma wartość różną od `None` i różną od napisu (żeby uniknąć nadpisania danych).
2. **`dodaj_przypis` ma liczyć rozmiar okienka z długości tekstu.** Przyjmij: $\text{szerokość} = \min(480, \max(220, \lceil \text{len(tekst)} / 40 \rceil \times 320))$ i $\text{wysokość} = \max(90, \lceil \text{len} / 40 \rceil \times 40)$. Uzasadnij wzór w komentarzu.
3. **Arkusz `Spis` ma być ustawiony jako aktywny** (`wb.active = 0`) — moduł 08.
4. **Nawigacja ma obejmować wszystkie arkusze**, także `Metodyka`. Weryfikacja ma to sprawdzać programowo (a nie Twoje „na pewno dodałem").
5. **Żaden link nie może mieć pustej wartości komórki** — weryfikacja ma to zgłosić jako błąd.
6. **Brak `#` w żadnym `location`.** Jeśli użyjesz prefiksu `#`, link będzie zewnętrzny i weryfikacja to wykaże.

W pliku `output/16_cw3_wnioski.md`:

1. **Dlaczego link wewnętrzny ma `location`, a nie `target`?** Odnieś się do `__attrs__` klasy `Hyperlink` i do funkcji `write_hyperlinks` (2.8).
2. **Co się stało z komórką, gdybyś ustawił link PRZED napisem?** Pokaż to eksperymentalnie i wklej wynik.
3. **Jak `quote_sheetname` wpłynął na link do arkusza `"Dane 2026"`?** Co by się stało bez niego? Sprawdź w Excelu.
4. **Dlaczego `dodaj_przypis` liczy rozmiar okienka samodzielnie?** Co się dzieje z domyślnym `144 × 79`, gdy tekst ma 200 znaków?
5. **Dlaczego komentarz do scalonego nagłówka dodajesz TYLKO w lewym górnym rogu?** Przekonaj się eksperymentalnie (2.10).
6. **Co byś musiał zrobić, żeby obrazy w raporcie miały prawdziwy tekst alternatywny?** Wymień **trzy** drogi (jedna: Excel ręcznie; druga: hack XML z 2.7; trzecia: Twoja własna) i oceń każdą pod kątem kosztu i ryzyka.
7. **Który wzorzec z modułu 24 zapowiada ten kod?** Zauważ, że funkcje przyjmują `ws` (typ openpyxl) w sygnaturach. Jak zmieniłbyś sygnatury, żeby **żaden typ openpyxl nie przeciekał na zewnątrz?**

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Raport wieloarkuszowy - kluczowe fragmenty."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import quote_sheetname
from openpyxl.worksheet.hyperlink import Hyperlink

AUTOR_PRZYPISOW = "Zespół raportowy"
DOMYSLNA_WYS_OKIENKA = 79


# ======================================================================
# 1. STRUKTURA: nazwy arkuszy w jednym miejscu (modul 09 - rejestr)
# ======================================================================
ARKUSZE = {
    "spis": "Spis",
    "dane": "Dane 2026",
    "podsumowanie": "Podsumowanie",
    "metodyka": "Metodyka",
    "dostepnosc": "Opis dostępności",
}


def naglowek_raportu(ws, tytul: str, podtytul: str, ostatnia_kol: int) -> None:
    """Naglowek scalony w wierszach 1-2. Komentarz TYLKO w lewym gornym rogu."""
    ostatnia_letter = ws.cell(row=1, column=ostatnia_kol).column_letter
    ws.merge_cells(f"A1:{ostatnia_letter}2")
    komorka = ws["A1"]                       # LEWY GORNY ROG scalenia
    komorka.value = tytul
    komorka.font = Font(bold=True, size=16)
    komorka.alignment = _wysrodkuj()
    # UWAGA: komentarz dodajemy do A1 (działa), NIE do B1 (2.10)

    ws.cell(row=3, column=1, value=podtytul).font = Font(italic=True, size=10)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22


def _wysrodkuj():
    from openpyxl.styles import Alignment
    return Alignment(horizontal="center", vertical="center")


def ustaw_kolumny(ws, szerokosci: dict) -> None:
    """{ 'A': 22.0, 'B': 14.0 } - szerokosci w ZNAKACH (modul 11)."""
    for litera, szer in szerokosci.items():
        ws.column_dimensions[litera].width = szer


# ======================================================================
# 2. NAWIGACJA - jedno miejsce, ktore WYMUSZA kolejnosc napis->link
# ======================================================================
def dodaj_link_z_napisem(ws, komorka: str, napis: str,
                         target: str | None = None,
                         location: str | None = None,
                         tooltip: str | None = None) -> None:
    """Bezpieczne dodanie hiperlinku.

    WYMUSZA kolejnosc napis->link (sekcja 2.8): gdyby link byl pierwszy,
    openpyxl wpisalby do komorki surowy adres.

    Zabezpieczenie: nie nadpisujemy istniejacej, INNEJ wartosci.
    """
    c = ws[komorka]
    if c.value is not None and c.value != napis:
        raise ValueError(
            f"{komorka} ma już wartość {c.value!r}; nadpisanie jej linkiem "
            f"jest prawdopodobnie błędem."
        )
    if target is None and location is None:
        raise ValueError("Podaj target (zewnętrzny) albo location (wewnętrzny).")

    c.value = napis                                  # 1) NAJPIERW napis
    if location is not None:
        c.hyperlink = Hyperlink(ref=komorka, location=location,
                                tooltip=tooltip)      # 2a) WEWNETRZNY
    else:
        c.hyperlink = Hyperlink(ref=komorka, target=target,
                                tooltip=tooltip)      # 2b) ZEWNETRZNY
    c.style = "Hyperlink"


def link_do_arkusza(nazwa_arkusza: str) -> str:
    """Buduje location dla arkusza - z cytowaniem, gdy trzeba (modul 07)."""
    return f"{quote_sheetname(nazwa_arkusza)}!A1"


def dodaj_link_powrotu(ws, spis_nazwa: str) -> None:
    dodaj_link_z_napisem(
        ws, "A4", "← Powrót do spisu",
        location=link_do_arkusza(spis_nazwa),
        tooltip="Wróć do spisu treści",
    )


def dodaj_spis(ws, arkusze: list[str]) -> None:
    ws["A1"] = "Spis treści"
    ws["A1"].font = Font(bold=True, size=16)

    for i, nazwa in enumerate(arkusze, start=3):
        dodaj_link_z_napisem(
            ws, f"A{i}", f"→ {nazwa}",
            location=link_do_arkusza(nazwa),
            tooltip=f"Przejdź do arkusza {nazwa}",
        )


# ======================================================================
# 3. PRZYPISY - rozmiar okienka liczony z dlugosci tekstu
# ======================================================================
def rozmiar_okienka(tekst: str) -> tuple[int, int]:
    """Rozmiar okienka komentarza (PIKSELE) z dlugosci tekstu.

    Wzor: przyjmujemy ~40 znakow w wierszu.
      szerokosc: max(220, min(480, wiersze * 320 / 1.5))?
    Uproszczenie, ktore trzymamy w zadaniu:
      wiersze   = ceil(len / 40)
      szerokosc = min(480, max(220, wiersze * 320))
      wysokosc  = max(90, wiersze * 40)

    Dlaczego tak: domyslne 144x79 px miesci ~2 wiersze po 25 znakow.
    Skalujemy OBYDWA wymiary, bo inaczej dlugi tekst "ucieka" w dol,
    a Excel nie rozszerza okienka automatycznie.
    """
    import math
    wiersze = max(1, math.ceil(len(tekst) / 40))
    szerokosc = min(480, max(220, wiersze * 320))
    wysokosc = max(90, wiersze * 40)
    return szerokosc, wysokosc


def dodaj_przypis(ws, komorka: str, tekst: str, autor: str = AUTOR_PRZYPISOW) -> None:
    """Komentarz z DOSTOSOWANYM rozmiarem okienka (sekcja 2.9)."""
    k = Comment(tekst, autor)
    k.width, k.height = rozmiar_okienka(tekst)
    ws[komorka].comment = k


def dodaj_przypisy_kolumn(ws, wiersz_naglowka: int, przypisy: dict) -> None:
    """{ 'B': 'Kwota w groszach...', 'C': '...' } - po jednym na kolumne."""
    for litera, tekst in przypisy.items():
        # UWAGA: komentarz idzie do wiersza z DANYMI, nie do naglowka,
        # jesli naglowek jest czescia scalonego zakresu.
        dodaj_przypis(ws, f"{litera}{wiersz_naglowka}", tekst)


# ======================================================================
# 4. WERYFIKACJA - program, nie "na pewno dodalem"
# ======================================================================
def inwentarz_linkow(sciezka: Path) -> None:
    wb = load_workbook(sciezka)
    bledy = []
    try:
        for ws in wb.worksheets:
            if ws.title == ARKUSZE["spis"]:
                continue
            ma_powrot = False
            for row in ws.iter_rows():
                for c in row:
                    h = c.hyperlink
                    if h is None:
                        continue
                    rodzaj = "WEW" if h.target is None else "ZEW"
                    print(f"  {ws.title:<18}{c.coordinate:<6}[{rodzaj}] "
                          f"{h.target or h.location}")

                    # niepusta wartosc?
                    if c.value is None:
                        bledy.append(f"{ws.title}!{c.coordinate}: pusty napis linku")

                    # czy jest link powrotu?
                    if h.location and h.location.endswith("!A1") \
                            and ARKUSZE["spis"] in str(h.location):
                        ma_powrot = True

            if not ma_powrot:
                bledy.append(f"{ws.title}: BRAK linku powrotu do spisu")
    finally:
        wb.close()

    print()
    if bledy:
        print("PROBLEMY:")
        for b in bledy:
            print("  ✗", b)
    else:
        print("✓ Nawigacja kompletna, wszystkie linki mają napisy.")
```

**Kluczowe decyzje i uzasadnienia:**

- **`dodaj_link_z_napisem` jako jedyna droga do tworzenia linku.** Nie ma „drugiego sposobu" na dodanie hiperlinku w tym projekcie. To jest dokładnie ta sama dyscyplina co `dodaj_wykres_sprzedazy` w module 15 i `dodaj_miniature` w 2.13. **Jedna funkcja = jedna czynność, a kolejność operacji jest wymuszona przez API funkcji.** W module 24 nazwiesz to **Fasadą**.

- **`link_do_arkusza` używa `quote_sheetname`.** Nazwa `"Dane 2026"` ze spacją **musi** być w apostrofach w `location`, inaczej link nie zadziała. To ten sam problem, który rozwiązywałeś w module 07 (i który w module 15 dawał błąd `#1190 Cannot create charts for worksheets with quotes in the title`).

- **`rozwiazanie` z `raise ValueError` zamiast cichego nadpisania.** To jest **architektura, która nie pozwala popełnić błędu** (moduł 22). Gdybyś w pętli generującej spis przypadkiem trafił w komórkę z danymi, funkcja **krzyknie** zamiast zamienić raport w śmietnik. To ta sama filozofia co `ValueError` przy duplikującej się nazwie tabeli (moduł 12) i `AttributeError` przy komentarzu w scalonej komórce (2.10) — **openpyxl krzyczy, gdy coś jest nie w porządku; Twój kod też powinien.**

- **`rozmiar_okienka` liczy oba wymiary.** Zwróć uwagę na uzasadnienie w docstringu: domyślne $144 \times 79$ mieści około dwóch wierszy. Bez skalowania szerokości długi tekst wychodzi poza okienko, a **Excel nie rozszerza okienka automatycznie** — więc trzeba to policzyć. To ten sam rodzaj decyzji co „wybierz `gapWidth = 80`" z modułu 15: **drobiazg, który decyduje o wrażeniu estetycznym.**

- **Weryfikacja jest programem.** Sedno: `ma_powrot` ustawiane **w trakcie iteracji** i sprawdzane **po jej zakończeniu**. To „asercja struktury" — nie sprawdzasz „czy napisałem kod", tylko „czy efekt końcowy ma pożądaną własność". W module 20 nazwiesz to **testem regresji na inwariantach** i będzie to jeden z fundamentów testowania plików binarnych.

- **Gdzie tu jest przeciek abstrakcji?** Funkcje przyjmują `ws` — czyli typ openpyxl. To jest **świadoma decyzja na tym etapie kursu**. W module 24 zmienisz sygnatury na `dodaj_link_z_napisem(raport, arkusz_nazwa, komorka, ...)`, gdzie `raport` to Twój obiekt, a nie `Worksheet`. **Wtedy będzie można przetestować logikę nawigacji bez openpyxl w ogóle.** To jest cały sens Fasady — i to jest zadanie do przemyślenia w punkcie 7 refleksji.

- **Arkusz „Opis dostępności" nie jest ozdobą.** Jest **dokumentacją ograniczenia** (2.7). To wzorzec z modułu 26: *co system wie o sobie samym, powinien móc powiedzieć.* Zapisanie „obrazy nie mają tekstu alternatywnego, uzupełnij ręcznie" w samym raporcie to akt **inżynierskiej uczciwości** — i jednocześnie informacja dla użytkownika, że to nie jego błąd.

</details>

## 6. Typowe błędy i pułapki

**1. „Obraz ma złe wymiary — albo mikroskopijny, albo gigantyczny" (objaw) → **myślisz w centymetrach, a `img.width` jest w pikselach.** Wykresy są w cm (moduł 15), obrazy w px (2.3). Dla obrazu $800$ px „ustawienie `width = 8`" daje obraz szeroki na 8 pikseli (czyli $0{,}2$ cm — kropkę) (przyczyna) → **używaj `cm_to_pixels(cm)`.** Zapisz funkcję raz w projekcie i nie wpisuj pikseli „na oko". Sprawdź też `img.width` **przed** ustawieniem, żeby znać punkt odniesienia (naprawa).**

**2. „Wszystkie obrazy w raporcie są rozciągnięte / spłaszczone" (objaw) → **ustawiłeś `img.width` i `img.height` niezależnie.** `Image` **nie ma** `resize_proportional` (to `Drawing`, nie `Image` — 2.3). Obraz rozciąga się do podanych wymiarów (przyczyna) → **licz proporcję z oryginału:** `skala = nowa_szer / img.width; img.height = round(img.height * skala)` (naprawa).**

**3. „`ValueError: I/O operation on closed file` przy drugim zapisie" (objaw) → **użyłeś tego samego obiektu `Image` dla dwóch `wb.save()`**, a `_data()` **zamyka strumień** po pierwszym odczycie. Dotyczy obrazów ze `BytesIO`; obrazy ze **ścieżki** są odporne, bo plik jest otwierany od nowa (2.5) (przyczyna) → **Trzymaj BAJTY, nie obiekt `Image`.** Twórz `Image(BytesIO(bajty))` **świeżo** dla każdego zapisu. To reguła odporna na wersję biblioteki (naprawa).**

**4. „`FileNotFoundError` przy `wb.save()` — a plik przecież istniał!" (objaw) → **usunąłeś/przeniosłeś plik obrazu między `add_image` a `save()`.** `_data()` otwiera źródło **ponownie** w momencie zapisu (2.5) (przyczyna) → **Trzymaj pliki źródłowe do końca zapisu.** W programach przetwarzających pliki tymczasowe: kopiuj obraz do trwałego katalogu albo **wczytaj go do `BytesIO`** zaraz przy wejściu (naprawa).**

**5. „`ImportError: You must install Pillow to fetch image objects`" (objaw) → **Pillow nie jest zależnością obowiązkową openpyxl** (2.2). Analogicznie: `module 'PIL' has no attribute 'Image'` przy konflikcie instalacji (przyczyna) → **`pip install pillow`** i sprawdź `python -c "import PIL; print(PIL.__version__)"`. W CI i Dockerze dodaj `pillow` do zależności (moduł 19, 20) (naprawa).**

**6. „Wstawiam obraz, a on ląduje na moich danych w A1" (objaw) → **`Image.anchor` ma klasowy domyślny `"A1"`.** `ws.add_image(img)` **bez** drugiego argumentu wstawia obraz na same dane (1.3, 2.4) (przyczyna) → **Zawsze podawaj kotwicę: `ws.add_image(img, "F2")`.** Traktuj obraz jak wykres z modułu 15 — planuj układ „dane po lewej, nakładki po prawej" (naprawa).**

**7. „Sortuję dane, a obraz zostaje na miejscu i zasłania coś innego" (objaw) → **obraz jest przypięty do komórki-kotwicy, nie do danych.** Domyślnie `OneCellAnchor` trzyma pozycję, a nie wartość (1.3) (przyczyna) → **Traktuj obrazy jako element layoutu, nie danych.** Jeśli obraz ma „jechać" razem z danymi — użyj `TwoCellAnchor` (2.4), ale pamiętaj, że to nadal pozycja komórek, a nie „powiązanie z wierszem" (naprawa).**

**8. „Po wczytaniu pliku obraz ma inny rozmiar niż ten, który zapisałem" (objaw) → **po odczycie `img.width` pokazuje **oryginalny** rozmiar z bajtów obrazu.** Czytnik robi `Image(BytesIO(bajty))` i osobno przypisuje kotwicę — a **faktyczna skala siedzi w `anchor.ext`** (Przykład 1, weryfikacja) (przyczyna) → **Czytaj rozmiar wyświetlania z kotwicy:** `EMU_to_pixels(img.anchor.ext.cx)`. **Nigdy nie polegaj na `img.width` przy czytaniu cudzych plików** — to najczęstszy powód błędnych „rozmiarów obrazów" w narzędziach audytowych (naprawa).**

**9. „Chcę dodać obraz na arkusz wykresów i dostaję `AttributeError: 'Chartsheet' object has no attribute 'add_image'`" (objaw) → **`Chartsheet` nie ma `add_image`** (2.11). To nie jest luka do obejścia (przyczyna) → **Wstaw obraz na zwykły arkusz.** Jeśli potrzebujesz „dużego, czystego miejsca na obraz" — normalny arkusz z `sheet_view.showGridLines = False` (moduł 11) daje ten sam efekt (naprawa).**

**10. „Wczytuję plik w `read_only` i nie widzę ani obrazów, ani komentarzy" (objaw) → **w trybie `read_only` obrazy nie są w ogóle przetwarzane przez czytnik, a komentarze nie są obsługiwane** (moduł 15, 2.11) (przyczyna) → **Do odczytu obrazów i komentarzy użyj zwykłego `load_workbook`.** Zaakceptuj koszt pamięci. **Nigdy nie buduj narzędzia diagnostycznego na `read_only`** — dojdziesz do fałszywego wniosku „plik nie ma obrazów" (naprawa).**

**11. „Komentarz w komórce scalonego zakresu znika / rzuca `AttributeError`" (objaw) → **`MergedCell` (pod-komórka scalenia) nie ma atrybutu `comment`.** Przy wczytywaniu openpyxl wypisuje wtedy ostrzeżenie i **pomija** komentarz (2.10) (przyczyna) → **Komentarz umieszczaj wyłącznie w LEWYM GÓRNYM rogu scalonego zakresu** albo w osobnej komórce obok. **Słuchaj ostrzeżeń** — to jeden z niewielu sygnałów utraty danych, które openpyxl podaje wprost (naprawa).**

**12. „Komentarz ma domyślny rozmiar i tekst się w nim nie mieści" (objaw) → **domyślnie `144 × 79` px** (`Comment.__init__`), co mieści ~2 wiersze. Excel **nie rozszerza okienka** automatycznie (2.9) (przyczyna) → **Ustaw `comment.width` i `comment.height` jawnie** — najlepiej funkcją liczącą rozmiar z długości tekstu (Przykład 4, szkic zadania 3). Pamiętaj: **piksele**, więc z punktów przez `points_to_pixels` (naprawa).**

**13. „Formatowanie tekstu w komentarzu zniknęło po przepuszczeniu przez openpyxl" (objaw) → **openpyxl obsługuje tylko TEKST i AUTORA komentarza.** Formatowanie i wymiary okna są w VML, którego nie parsuje (2.9) (przyczyna) → **Traktuj komentarze jako „tekst + autor" i nic więcej.** Jeśli formatowanie ma znaczenie — nie przepuszczaj pliku przez openpyxl, albo zaakceptuj utratę (moduł 18) (naprawa).**

**14. „Hiperlink wewnętrzny nie działa — klikam i nic się nie dzieje" (objaw) → **użyłeś `cell.hyperlink = "#Arkusz!A1"`.** Prefiks `#` sprawia, że tekst staje się `target` (relacja **zewnętrzna**), a nie `location` (nawigacja w skoroszycie) (2.8) (przyczyna) → **Użyj `Hyperlink(ref=..., location=...)`**, bez `#`. Dla nazw arkuszy ze spacjami użyj `quote_sheetname` (naprawa).**

**15. „Hiperlink zniknął — komórka pokazuje tylko zwykły adres tekstem" (objaw) → **to nie zniknął — to zachowanie settera.** Przy pustej komórce openpyxl **wpisuje do niej adres** jako napis (`self.value = val.target or val.location`), a **stylu `"Hyperlink"` nie nadaje automatycznie** (2.8) (przyczyna) → **Ustaw napis NAJPIERW**, potem link; i dodaj `cell.style = "Hyperlink"` dla niebieskiego podkreślenia. Sprawdź kolejność: jeśli `cell.value` było `None` przed przypisaniem, Twój napis został nadpisany adresem (naprawa).**

**16. „Obraz nie drukuje się, choć jest w arkuszu" (objaw) → **obraz nie należy do zakresu wydruku.** Jeśli kotwica leży **poza** `print_area` (moduł 11), obraz nie trafi na wydruk ani do PDF (przyczyna) → **Ustaw `print_area` tak, by obejmował obszar z obrazami**, albo przeciągnij obraz w obręb zakresu. Inne przyczyny: obraz jest poza `fitToPage` (moduł 11) albo stukrotnie „wypełzł" poza obszar z powodu złego przeliczenia pikseli (naprawa).**

**17. „Ostrzeżenia openpyxl znikają — widzę tylko pierwsze" (objaw) → **domyślny filtr Pythona pokazuje każde **unikalne** ostrzeżenie tylko raz na sesję.** Przy dziesięciu komentarzach w scalonych komórkach zobaczysz jedno ostrzeżenie (Przykład 5) (przyczyna) → **Użyj `warnings.catch_warnings(record=True)` + `warnings.simplefilter("always")`.** Zbieraj ostrzeżenia i **wypisuj je jako raport ryzyka** przed zapisem (moduł 18) (naprawa).**

**18. „Dodałem hiperlink i wykres jednocześnie i plik się popsuł / linki się podwoiły" (objaw) → **historyczne błędy kolizji relacji** — `#88 Charts break hyperlinks`, `#462 Vestigial rId conflicts when adding charts, images or comments`, `#1496 Hyperlinks duplicated on multiple saves`. To obszar, w którym relacje rysunków i hiperlinków współdzielą przestrzeń `rId` (2.1) (przyczyna) → **Trzymaj się aktualnej wersji 3.1.5** — te konkretne problemy są naprawione. Jeśli używasz starszej gałęzi (2.x), **unikaj zapisu wielokrotnego** plików z komentarzami (`#1330 Workbooks with comments cannot be saved multiple times`, naprawione w 2.6.4) i **sprawdzaj plik okiem po każdym typie zmiany** (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Obraz, hiperlink i komentarz to NAKŁADKI, nie dane.** Nie są wartościami komórki. Obraz to karteczka na pinesce (przypięta do kotwicy, nie do danych), hiperlink to drzwi pod napisem, komentarz to przypis na marginesie. **Konsekwencja:** nie odczytasz ich przez `cell.value`, nie zachowają się jak dane w sortowaniu, a `modelem` do nich jest `ws._images`, `cell.hyperlink`, `cell.comment` — nie zawartość siatki.

2. **Obraz mierzy się w PIKSLACH, wykres w CENTYMETRACH, a pod spodem wszystko jest w EMU.** `Image` **nie ma** `resize_proportional` — proporcje liczysz sam. openpyxl **nie ma** `cm_to_pixels` — napisz tę funkcję raz dla całego projektu. A po **wczytaniu** pliku `img.width` pokazuje oryginał, a faktywna skala siedzi w `anchor.ext` — dlatego rozmiary obrazów czytaj z kotwicy, nie z atrybutu.

3. **Obraz to KOPIA BAJTÓW, a nie adres — i to zmienia wszystko.** Plik trafia do `xl/media/`, czyli „wklejony" do archiwum. Dlatego obraz nie reaguje na zmianę danych (przeciwnie niż wykres z modułu 15), dlatego zwiększa rozmiar pliku, i dlatego **`_data()` musi otwierać źródło ponownie przy zapisie** — a po odczycie **zamyka strumień**. **Reguła: trzymaj bajty, nie obiekt `Image`.** Świeży `Image` na każdy `save()`.

4. **Hiperlinki są solidne, ale mają dwa tryby, których nie wolno mylić.** `target` = **relacja zewnętrzna** (`https`, `mailto`, plik), `location` = **nawigacja w skoroszycie**; to drugie **nie tworzy relacji**, tylko atrybut w `sheet1.xml`. Dlatego **link wewnętrzny to `Hyperlink(location=...)`, a NIE `"#Arkusz!A1"`.** I dlatego **ustawiaj napis przed hiperlinkiem** — inaczej openpyxl wpisze do komórki surowy adres, a nie Twój napis.

5. **openpyxl ma trzy dziury, o których trzeba wiedzieć zawczasu.** (a) **Tekst alternatywny obrazów nie istnieje** — openpyxl wpisuje dosłownie `descr="Picture"` dla każdego obrazu i nie ma API, żeby to zmienić (zaplanuj to w projekcie, nie na końcu). (b) **Komentarze w scalonych pod-komórkach są gubione** — z ostrzeżeniem, które **musisz przechwycić**, bo to sygnał utraty danych. (c) **Obrazy i komentarze są niedostępne w `read_only`**, a **obrazy nie przechodzą przez `copy_worksheet`** — mimo że **hiperlinki i komentarze przechodzą**.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image                      # wymaga pillow!
from openpyxl.drawing.spreadsheet_drawing import (
    OneCellAnchor, TwoCellAnchor, AbsoluteAnchor, AnchorMarker,
)
from openpyxl.drawing.xdr import XDRPositiveSize2D, XDRPoint2D
from openpyxl.comments import Comment
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.utils import quote_sheetname, units
from openpyxl.utils.units import (
    pixels_to_EMU, EMU_to_pixels,      # 1 px = 9525 EMU
    points_to_pixels, pixels_to_points,# okienka komentarzy, wysokosci wierszy
    cm_to_EMU, EMU_to_cm,              # wykresy (modul 15)
)

# ==================================================================
# 2. JEDNOSTKI - czego BRAKUJE w openpyxl (napisz sam!)
# ==================================================================
def cm_to_pixels(cm, dpi=96):  return round(cm / 2.54 * dpi)
def pixels_to_cm(px, dpi=96):  return round(px * 2.54 / dpi, 2)
# ⚠️ NIE MA units.cm_to_pixels / pixels_to_cm!

# ==================================================================
# 3. OBRAZ - WSTAWIANIE (jednostka: PIKSELE!)
# ==================================================================
imgg = Image("logo.png")                   # sciezka, PIL.Image albo BytesIO
print(imgg.width, imgg.height, imgg.format) # np. 320 120 png

# skalowanie Z ZACHOWANIEM PROPORCJI (Image NIE ma resize_proportional!)
skala = cm_to_pixels(4.0) / imgg.width     # skala z zadanej szerokosci w cm
imgg.height = round(imgg.height * skala)
imgg.width = cm_to_pixels(4.0)

ws.add_image(imgg, "E2")                   # ALBO imgg.anchor = "E2"; ws.add_image(imgg)
# ⚠️ BEZ anchor OBRAZ LADUJE W A1!

# wysokość wiersza MIESZCZĄCA obraz (w PUNKTACH, nie pikselach!)
ws.row_dimensions[2].height = units.pixels_to_points(imgg.height)

# ==================================================================
# 4. OBRAZ - STRUMIEN I GENEROWANIE W LOCIE
# ==================================================================
ws.add_image(Image(BytesIO(bajty_png)), "B2")

# matplotlib -> bajty (NIC nie zapisujemy na dysk)
import matplotlib
matplotlib.use("Agg")                      # backend bez GUI - serwery!
import matplotlib.pyplot as plt

fig, ax = plt.subplots(figsize=(7, 3.5), dpi=120)
ax.bar(["A", "B", "C"], [1, 3, 2])
buf = BytesIO()
fig.savefig(buf, format="png", bbox_inches="tight")
plt.close(fig)                             # ⚠️ KLUCZOWE: zwolnij pamiec!
bajty_png = buf.getvalue()                 # bytes - zrodlo prawdy
ws.add_image(Image(BytesIO(bajty_png)), "F2")   # swiezy Image!

# ⚠️⚠️ REGUŁA: trzymaj BAJTY, nie obiekt Image!
# _data() ZAMYKA strumien -> drugi save() da:
#    ValueError: I/O operation on closed file
for sciezka in ("a.xlsx", "b.xlsx"):
    wb = Workbook(); ws = wb.active
    ws.add_image(Image(BytesIO(bajty_png)), "B2")   # swiezy obiekt!
    wb.save(sciezka); wb.close()

# ⚠️ Nie usuwaj pliku zrodlowego przed wb.save() - _data() otwiera go PONOWNIE

# ==================================================================
# 5. OBRAZ - KOTWICE (indeksy 0-BASED!)
# ==================================================================
# domyslnie: OneCellAnchor, rozmiar staly w EMU
imgg.anchor = OneCellAnchor(
    _from=AnchorMarker(col=4, row=1),                       # col=4 -> E, row=1 -> wiersz 2
    ext=XDRPositiveSize2D(cx=pixels_to_EMU(300), cy=pixels_to_EMU(150)),
)
# obraz rozciagajacy sie z komorkami:
imgg.anchor = TwoCellAnchor(editAs="twoCell",
                            _from=AnchorMarker(col=4, row=10),
                            to=AnchorMarker(col=8, row=16))

# ==================================================================
# 6. OBRAZ - ODCZYT I DIAGNOSTYKA
# ==================================================================
wb = load_workbook(path)          # ⚠️ NIE read_only! (obrazy niedostepne)
for img in ws._images:            # atrybut PRYWATNY - tylko diagnostyka
    print(img.path, img.format)   # /xl/media/image1.png  png
    print(img.width, img.height)  # ⚠️ to ORYGINAL, nie rozmiar wyswietlania!
    a = img.anchor
    if hasattr(a, "_from") and hasattr(a, "ext"):
        print("kotwica:", a._from.col, a._from.row)                 # 0-based!
        print("wyswietlane:", EMU_to_pixels(a.ext.cx), EMU_to_pixels(a.ext.cy))

# ⚠️ DESKRPCJA OBRAZU = "Picture" (na sztywno, brak API do zmiany):
#    pic.nvPicPr.cNvPr.descr = "Picture"      w SpreadsheetDrawing._picture_frame

# ==================================================================
# 7. HIPERLINKI
# ==================================================================
# ZEWNETRZNY (relacja External) - NAPIS NAJPIERW (sekcja 2.8!)
ws["B2"] = "Dokumentacja"
ws["B2"].hyperlink = "https://openpyxl.readthedocs.io/"
ws["B2"].style = "Hyperlink"           # wbudowany styl nazwany
ws["B3"].hyperlink = "mailto:a@b.com?subject=Raport"
ws["B4"].hyperlink = "file:///C:/raporty/x.xlsx"     # URI, nie C:\...

# WEWNETRZNY (location, BEZ relacji) - NIE uzywaj "#Arkusz!A1"!
ws["B5"].hyperlink = Hyperlink(ref="B5", location="'Dane 2026'!A1")
ws["B5"].hyperlink = Hyperlink(ref="B5", location=f"{quote_sheetname('Dane 2026')}!A1")
ws["B6"].hyperlink = Hyperlink(ref="B6", location="MojaNazwaZdefiniowana")

# Z TOOLTIPEM
ws["B7"].hyperlink = Hyperlink(ref="B7", target="https://example.com",
                               tooltip="Otworzy sie w przegladarce")

ws["B8"].hyperlink = None             # USUNIECIE linku

# ODCZYT
h = ws["B2"].hyperlink
h.target      # 'https://...'  (None => link WEWNETRZNY)
h.location    # "'Dane 2026'!A1" (None => link zewnetrzny)
h.tooltip     # 'Otworzy sie...'

# ⚠️ PULAPKA: link na PUSTEJ komorce -> openpyxl WPISZE ADRES jako napis!
c = ws["D9"]                          # pusta
c.hyperlink = "https://example.com"
print(c.value)                        # 'https://example.com'  <- sam sie wpisal!

# ==================================================================
# 8. KOMENTARZE (notatki)
# ==================================================================
k = Comment("Wpisz kwote w groszach.", "Zespol raportowy")
k.width = 340                          # PIKSELE (domyslnie 144)
k.height = 110                         # PIKSELE (domyslnie 79)
ws["B3"].comment = k

# rozmiar z PUNKTOW (typografia)
k.width = units.points_to_pixels(240)
k.height = units.points_to_pixels(60)

# odczyt
k2 = ws["B3"].comment
k2.text, k2.content, k2.author         # text == content (aliasy)
# ⚠️ przy ODCZYCIE rozmiary sa TRACONE (zyja w VML)

# ten sam komentarz w 2 komorkach => openpyxl robi KOPIE (__copy__)
ws["C3"].comment = k; ws["C4"].comment = k
# ws["C3"].comment is k  -> True
# ws["C4"].comment is k  -> False

# ⚠️ SCALONE KOMORKI: tylko LEWY GORNY rog!
ws.merge_cells("A6:C6")
ws["A6"].comment = Comment("OK", "A")       # ✅
# ws["B6"].comment = Comment("NIE", "A")   # ❌ AttributeError (MergedCell)

# ==================================================================
# 9. NASLUCHIWANIE OSTRZEZEN (openpyxl komunikuje utrate danych!)
# ==================================================================
import warnings
with warnings.catch_warnings(record=True) as zlapane:
    warnings.simplefilter("always")          # ⚠️ bez tego tylko PIERWSZE!
    wb = load_workbook(path)
    ...
for w in zlapane:
    print("⚠️", w.message)
# Komunikat o komentarzach w scalonych:
#  "Cell 'Arkusz':C2 is part of a merged range but has a comment which
#   will be removed because merged cells cannot contain any data."

# ==================================================================
# 10. TRYBY PRACY - NIESYMETRYCZNIE!
# ==================================================================
# read_only=True   -> OBR AZY ✗, KOMENTARZE ✗, HIPERLINKI ✓
# write_only=True  -> OBR AZY ✓, (wykresy ✓, tabele ✓)
# Chartsheet       -> OBR AZY ✗ (AttributeError: no attribute 'add_image')
# copy_worksheet   -> kopiuje: WARTOŚCI, STYLE, HIPERLINKI, KOMENTARZE
#                     NIE kopiuje: OBRAZY, WYKRESY, walidacji, reguł CF

# ==================================================================
# 11. UNIT - PODSUMOWANIE JEDNOSTEK
# ==================================================================
# img.width / img.height ................... PIKSELE
# comment.width / comment.height ........... PIKSELE
# chart.width / chart.height ............... CENTYMETRY (modul 15)
# row_dimensions[h].height ................. PUNKTY
# column_dimensions["A"].width ............. ZNAKI
# AnchorMarker.col / .row .................. 0-BASED (reszta openpyxl 1-based!)
# anchor.ext.cx / .cy ...................... EMU (1 px = 9525 EMU)
```

## 9. Co dalej

Zamknęliśmy trzeci z czterech „światów obiektów" w kursie: dane (moduły 05–07), formatowanie (09–14), wykresy (15) i **nakładki (16)**. I to ten ostatni pokazał Ci coś, czego poprzednie nie mogły: **te same elementy zachowują się zupełnie inaczej w zależności od trybu pracy i od tego, czy plik tworzysz, czy czytasz.** Obraz „działa" w `write_only`, ale nie istnieje w `read_only`. Komentarz kopiuje się przez `copy_worksheet`, ale obraz nie. Hiperlink wewnętrzny nie ma relacji, a zewnętrzny ma.

Trzy rzeczy, które zabierasz do dalszej pracy:

- **Trzy jednostki i jedna reguła.** Piksele dla obrazów i komentarzy, punkty dla wierszy, centymetry dla wykresów, a pod spodem EMU. I reguła nadrzędna: **piksele w kodzie to jednostka implementacyjna — mów intencjami (centymetry) i przeliczaj w jednej funkcji.** To ten sam wzorzec co `FORMAT["waluta_pln"]` (moduł 10) i `STYL` (moduł 15).

- **Cykl życia zamiast „obiektu, który działa".** Pułapka zamkniętego strumienia, plik źródłowy otwierany ponownie przy zapisie, `mark_to_close` — to wszystko uczy Cię czegoś ogólniejszego niż openpyxl: **obiekt może być ważny „przez chwilę", a źródłem prawdy są bajty.** W module 19 (duże pliki) zobaczysz to samo w skali makro: `write_only` to „bajty płyną, obiekt umiera".

- **Ostrzeżenia jako kanał informacji.** openpyxl **mówi**, gdy traci dane (komentarz w scalonej komórce, sparkline, WMF) — ale tylko wtedy, gdy go słuchasz. `warningS.catch_warnings(record=True)` + `simplefilter("always")` to Twoje narzędzie. **W module 18 zamienimy ten nasłuch w pełny raport ryzyka pliku wejściowego.**

**Zadanie do zrobienia teraz, zanim ruszysz dalej.** Otwórz `output/16_obraz_podstawy.xlsx`, a potem wykonaj **test utraty danych** z tego modułu w praktyce:

1. W Excelu dodaj **notatkę** do komórki w środku scalonego zakresu (nie ma takiego w pliku — **najpierw sam scal** `A14:C14`, potem dodaj notatkę do `B14`). Excel to pozwoli.
2. Dodaj też **tekst alternatywny** do jednego z obrazów (`Formatuj obraz → Tekst alternatywny`).
3. Zapisz plik.
4. Wczytaj go `examples/16_inwentarz.py` i przeczytaj ostrzeżenia.
5. Zapisz ten plik pod nową nazwą przez openpyxl i **porównaj z oryginałem** — notatka w scaleniu zniknęła, a tekst alternatywny obrazu to znowu `descr="Picture"`.

To jest cała wiedza z tego modułu w jednym doświadczeniu: **dwa elementy, dwie różne przyczyny utraty danych, jedna wspólna lekcja — sprawdzaj okiem, słuchaj ostrzeżeń i nie przepuszczaj cudzych plików bez inwentarza.**

W **module 17** zajmiemy się czymś zupełnie innym: **walidacją danych i ochroną** — czyli **ograniczaniem użytkownika**. Zobaczysz listy rozwijane (`DataValidation`, w tym listę **z innego arkusza** — najczęstsze pytanie na forach i jedno z niewielu miejsc, gdzie openpyxl ma realne ograniczenie techniczne), ograniczenia liczbowe i tekstowe, a potem ochronę arkusza (`Protection(locked=False)` jako mechanizm „formularza wejściowego") i skoroszytu. Zobaczysz też **uczciwą prawdę o haśle w OOXML**: to kłódka na szufladzie w biurku, nie sejf — chroni przed pomyłką, nie przed intencją. To będzie ostatni moduł, zanim w module 18 wejdziemy w **najważniejszy temat całego kursu: modyfikację istniejących plików bez utraty zawartości.**