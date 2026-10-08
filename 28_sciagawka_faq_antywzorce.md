# Moduł 28 — Ściągawka, FAQ i antywzorce

> **Część:** V — Projekt i domknięcie · **Poziom:** ⭐ (moduł do wracania) · **Wymaga:** modułów 00–27

## 0. W tym module nauczysz się

- **Znajdziesz odpowiedź w ciągu minuty, nie godziny.** Ten moduł nie uczy nowych rzeczy — porządkuje wszystko, co już wiesz, w formę, którą da się **przeszukać wzrokiem**. Trzynaście tabel ściągawki, mapa „problem → moduł", dwadzieścia najczęstszych błędów.
- **Zrozumiesz „trzy kategorie prawdy"** o API (`potrafi` / `potrafi częściowo` / `gubi przy zapisie`) jako **szkielet** całej ściągawki, a nie jako jedno z ostrzeżeń. Każda tabela w tym module jest zbudowana tak, żebyś **od razu wiedział, w której kategorii jesteś**.
- **Poznasz **kolejność diagnostyki** — sześć kroków, które wykonuje się zawsze, gdy „plik wygląda dobrze, ale coś nie liczy". Bez tej kolejności tracisz godziny na sprawdzanie tego, co akurat przyszło Ci do głowy.
- **Otrzymasz narzędzie `xlsxcheck`** — kompletny, uruchamialny skrypt, który robi inwentarz części pliku, podsumowanie struktury, wykrycie ryzyk utraty danych, porównanie dwóch skoroszytów i audyt przed wysyłką. To praktyczny owoc całego kursu.
- **Zobaczysz osiem antywzorców projektowych** — nie „złych praktyk" w sensie ogólnym, ale konkretnych konstrukcji, które **widuje się w niemal każdym** skrypcie openpyxl w firmach, wraz z tym, czym dokładnie za nie płacisz.
- **Dostaniesz tabele migracji** z `xlrd`, `xlwt`, `xlsxwriter` i `openpyxl 2.x` — jeżeli pracujesz z kodem, który ma pięć lat, to jest Twoja tabela tłumaczeń.
- **Dowiesz się, gdzie szukać dalej** — dokumentacja 3.1.x, changelog (czytany **przed** użyciem egzotycznego parametru), wzorce GoF/Fowler, specyfikacja OOXML i uczciwe porównanie z pandas.

## 1. Intuicja i analogia

### 1.1. Tablica z hasłami na budowie, nie rozdział podręcznika

Wyobraź sobie budowę. Ekipa zna swoją pracę. Ale przy wejściu na teren wisi **tablica z hasłami**: gdzie jest prąd, gdzie jest woda, gdzie gaśnica, kto ma klucze, co zrobić, gdy zatnie się winda, numer do kierownika. Nikt tej tablicy nie czyta od góry do dołu podczas przerwy na kawę. Ludzie **podchodzą do niej** w jednym momencie — gdy coś się dzieje albo gdy potrzebują jednej konkretnej informacji.

Ten moduł jest taką tablicą. Przez dwadzieścia siedem modułów budowałeś — i budowałeś dobrze. Teraz masz tablicę:

- **Wchodzisz z problemem** („chcę listę rozwijaną z innego arkusza") → idziesz do `§3.11`, sekcja Walidacja.
- **Wchodzisz z objawem** („plik się otwiera, ale sumy pokazują zero") → idziesz do `§6`, lista dwudziestu błędów, a potem do `§3.16` (FAQ).
- **Wchodzisz z decyzją** („chcę zrobić moduł, który zapisuje raport z szablonu") → idziesz do `§6.2` (antywzorce) i `§9` (mapa problem → moduł).

Trzy wejścia. To nie jest książka, którą się czyta. To jest **miejsce, do którego się przychodzi**. Jeżeli przeczytasz ten moduł raz, zapamiętasz z niego może pięć rzeczy. Jeżeli będziesz do niego wracał dziesięć razy, zapamiętasz większość — i to jest dokładnie ten sposób, w jaki zapamiętuje się narzędzia, a nie teorie.

### 1.2. Dlaczego ściągawka jest trudniejsza do napisania niż rozdział

Powiem Ci coś, co brzmi jak paradoks: **napisać ściągawkę jest trudniej niż napisać rozdział**. Rozdział może być długi — wszystko zmieści się w środku. Ściągawka musi być **krótka i kompletna jednocześnie**, a to są wymagania sprzeczne.

W praktyce oznacza to, że trzeba **wybierać**, co pominąć. I tu jest ryzyko: jeżeli pomijasz coś, co czytelnik naprawdę potrzebuje, ściągawka go **zawodzi w najgorszym momencie** — o drugiej nad ranem, przy awarii produkcyjnej, gdy nie ma czasu czytać trzech modułów.

Dlatego zbudowałem ten moduł na jednej zasadzie: **każdy wpis w ściągawce jest inny niż pozostałe, a te, które są podobne, są oznaczone jako „to samo, inaczej zapisane"**. Nie ma tu dwóch wierszy mówiących to samo innymi słowami. Za to każdy wiersz mówi **coś, czego nie ma w innym wierszu** — najczęściej to, **co się zepsuje**, gdy użyjesz go źle.

### 1.3. „Trzy kategorie prawdy" — kręgosłup tego modułu

Przez cały kurs wracało jedno rozróżnienie. Teraz je nazwę wprost, bo w ściągawce jest **strukturą**, nie dygresją:

**Kategoria P — „potrafi".** Openpyxl robi to w pełni. Możesz na tym polegać. Przykłady: zapis wartości, daty, style komórek, tabele, formatowanie warunkowe, scalanie, ochrona struktury, walidacja danych, ustawienia druku, strumieniowy zapis.

**Kategoria C — „potrafi częściowo".** Openpyxl robi to, ale albo nie w pełni, albo nie tak, jak myślisz, albo wczytuje i zapisuje z uproszczeniem. Przykłady: **wykresy** (obecne, ale **odbudowywane z modelu**, więc nietypowe formatowanie bywa upraszczane), **tabele przestawne** (odczyt definicji, brak tworzenia), **rich text** (od 3.1, z ograniczeniami), **komentarze w scalonych komórkach** (ostrzeżenie przy wczytywaniu, utrata przy zapisie), **własne właściwości dokumentu** (od 3.1), **kształty na rysunkach**.

**Kategoria L — „gubi przy zapisie" (loss).** Openpyxl **nie modeluje** tego, więc przy wczytaniu i ponownym zapisie **to znika**. To najważniejsza kategoria w całym kursie, bo jest **cicha** — nikt nie dostaje błędu, tylko traci dane. Przykłady: sparkline'y (rozszerzenie `x14` wewnątrz XML arkusza), slicery i timeline'y, model danych Power Pivot, zapytania Power Query, formanty ActiveX i kontrolki, niestandardowe części `customXml`, komentarze nowego typu (threaded comments), części rozszerzeń nieznane bibliotece.

**Analogia, która to porządkuje — przeprowadzka.** Wynajmujesz firmę do przepakowania mieszkania (openpyxl robi **przepisanie**, nie edycję — moduł 03). Trzy kategorie to:

- **P** — firma pakuje krzesła, talerze i ubrania. Wszystko dojedzie.
- **C** — firma pakuje żyrandol, ale **odkręca go i wkłada w karton bez osłonek**. Dojedzie, ale ryzykujesz, że coś się odkształci. Sprawdź po przyjeździe.
- **L** — firma mówi „płyty szachulcowe, szlifowane sklepienie i instalacja alarmowa: tego nie pakujemy". **Po przeprowadzce ich nie ma.** Nie ma błędu. Po prostu ich nie ma.

Jeżeli zapamiętasz jedną rzecz z tego modułu, niech to będzie: **przed każdym przepisaniem pliku przez openpyxl zapytaj, co w nim jest z kategorii L.** Reszta to szczegóły.

## 2. Teoria

### 2.1. Graf obiektów w trzydzieści sekund (przypomnienie z modułu 03)

```text
Workbook (wb)
├── properties          -> metadane dokumentu (tytuł, autor, wersja)
├── defined_names       -> słownik definiowanych nazw (DefinedName)
├── security            -> lockStructure, workbookPassword
├── calculation         -> fullCalcOnLoad (wymuszenie przeliczenia)
├── _named_styles       -> nazwane style skoroszytu
└── Worksheet[] (ws)
    ├── title, sheet_state, sheet_properties, sheet_view
    ├── protection          -> ochrona arkusza
    ├── _cells              -> słownik "A1" -> Cell  (TRYB NORMALNY)
    ├── tables              -> słownik displayName -> Table
    ├── conditional_formatting -> reguły warunkowe (zakres -> reguły)
    ├── data_validations    -> walidacje danych
    ├── merged_cells        -> scalone zakresy
    ├── column_dimensions / row_dimensions
    ├── page_setup / page_margins / print_area / print_title_rows
    ├── _images             -> obrazy (drawing)
    └── _charts             -> wykresy (drawing)
        └── Cell
            ├── value       -> wartość Pythona (albo literał formuły!)
            ├── data_type   -> 'n' 's' 'd' 'b' 'f' 'e'
            ├── _style      -> INDEKSY do wspólnej tabeli stylów
            │   (fontId, fillId, borderId, numFmtId, alignmentId, protectionId)
            ├── number_format -> kod reprezentacji liczby
            ├── hyperlink
            └── comment
```

Trzy konsekwencje, które trzeba mieć pod ręką:

1. **Indeksowanie od 1** (`ws.cell(row=1, column=1)` to `A1`) i **notacja A1 vs `(row, column)`** — obie istnieją, obie są poprawne, `$A$1` to już **tekst formuły**, nie adres Pythona.
2. **Komórka nie ma formatowania — ma indeksy.** Dlatego `cell.font.bold = True` nie działa: nie modyfikujesz „pola komórki", tylko obiekt współdzielony, a indeks w komórce zostaje ten sam. Analogia: **pieczątka leży w szufladzie, a komórka trzyma tylko numer pieczątki** — poprawienie pieczątki w szufladzie nie zmienia tego, co już odbite.
3. **`wb` w pamięci i plik na dysku to dwa światy.** `save()` to **przepisanie całości**, nie dopisanie. Trzecia konsekwencja jest źródłem wszystkich problemów z kategorii L.

### 2.2. Dlaczego openpyxl nie policzy formuły (i co z tego wynika)

Openpyxl **nie ma silnika obliczeniowego**. Nigdy nie będzie. To nie jest decyzja marketingowa, tylko architektoniczna: implementuje **format pliku OOXML**, a nie **program Excel**. Formuła to dla niego dowolny tekst zaczynający się od `=`, zapisany w komórce z atrybutem `t="str"` albo bez niego, jeśli w pliku jest cache.

Trzy fakty praktyczne:

- **Nowa formuła w pliku bez cache** → Excel policzy ją dopiero przy otwarciu (i tylko wtedy, gdy plik ma ustawione przeliczenie). Otwierając taki plik w Pythonie (`data_only=True`) dostaniesz `None`.
- **`data_only=True` czyta cache, nie formułę.** Jeżeli plik nigdy nie był otwarty w Excelu, cache może nie istnieć → `None`. Jeżeli był — cache bywa nieaktualny.
- **Zapisu po `data_only=True` nie robimy nigdy.** Taka kombinacja zamienia formuły na ich ostatnie wartości. To jest najczęstszy sposób „zjedzenia" cudzego pliku.

**Co zrobić, gdy naprawdę potrzebujesz liczby w Pythonie?** Cztery drogi, w kolejności od najlepszej:

1. **Policz w Pythonie.** Marża, agregaty, dynamika — to logika biznesowa, jej miejsce jest w domenie (moduł 26, 27).
2. **Ustaw `wb.calculation.fullCalcOnLoad = True`** i oddaj plik użytkownikowi. To działa, gdy plik ma w formule coś, co musi być widoczne dla człowieka.
3. **Przelicz przez LibreOffice headless:** `soffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir out plik.xlsx`. LibreOffice ma silnik obliczeń i po konwersji zapisze świeży cache.
4. **Użyj biblioteki z silnikiem formuł** (`formulas`, `pycel`) albo dostawcy chmurowego. To działa dla prostych przypadków i **zawodzi** na nowoczesnych funkcjach, tabelach i formułach tablicowych.

### 2.3. Kolejność diagnostyki — sześć kroków

Gdy „plik wygląda dobrze, ale coś nie liczy", **zawsze** wykonujesz te kroki w tej kolejności. Nie dlatego, że są magiczne — dlatego, że każdy następny jest droższy od poprzedniego, a każdy poprzedni eliminuje całe klasy przyczyn. To jest **sortowanie diagnozy od najtańszej do najdroższej**.

| Krok | Co robisz | Ile trwa | Co wyklucza |
|---|---|---|---|
| **1** | Otwórz plik w Excelu i **włącz widok formuł** (`Ctrl+`` ` ``) | 10 s | Czy w komórce jest wartość, czy formuła |
| **2** | `podsumowanie_skoroszytu()` — arkusze, wymiary, liczba komórek, reguł | 30 s | Czy to plik, o którym myślisz (nazwy, liczba wierszy) |
| **3** | Sprawdź `cell.data_type` i `type(cell.value)` dla trzech podejrzanych komórek | 1 min | „Liczba wygląda jak liczba" — najczęstsza przyczyna |
| **4** | `inwentarz(sciezka)` — części ZIP i **ich rozmiary** | 1 min | Czy plik ma to, co powinien (media, wykresy, tabele) |
| **5** | Porównaj z plikiem wzorcowym (`porownaj_skoroszyty`) | 5 min | Który **konkretnie** element się różni |
| **6** | Otoczenie: wersja openpyxl, obecność `lxml`/`pillow`, locale, wersja Excela | 5 min | Błędy środowiskowe, nie kodu |

**Dlaczego krok 1 nie może być ostatni.** Bo jeśli w komórce jest **formuła**, to żadna analiza wartości w Pythonie nie pomoże — a jeśli jest **wartość**, to żaden problem z formułami Cię nie dotyczy. To jedna obserwacja, która **połowi** przestrzeń poszukiwań. Ludzie pomijają ten krok, bo „przecież wiem, co tam jest" — i to jest najdroższe założenie w całej diagnostyce.

**Dlaczego krok 3 jest przed krokiem 4.** Bo błąd typu (tekst zamiast liczby) jest **najczęstszy** i **najtańszy** do naprawienia. Sprawdzanie zawartości archiwum przed sprawdzeniem typu wartości to pomyłka w kolejności — najpierw wykluczamy to, co prawdopodobne.

### 2.4. Anatomia zapisu: co dokładnie trafia do pliku

Wiesz z modułu 02, że `.xlsx` to ZIP z XML-em. Dla ściągawki potrzebna jest **mapa części**, którą da się porównać z `zipfile.namelist()`:

| Część pliku | Co zawiera | Kategoria |
|---|---|---|
| `[Content_Types].xml` | spis typów części pakietu | P |
| `_rels/.rels` | powiązania pakietu | P |
| `docProps/core.xml`, `docProps/app.xml` | właściwości dokumentu (`wb.properties`) | P |
| `docProps/custom.xml` | własne właściwości dokumentu | C (od 3.1) |
| `xl/workbook.xml` | lista arkuszy, definiowane nazwy, `calcPr` | P |
| `xl/_rels/workbook.xml.rels` | powiązania arkuszy i stylów | P |
| `xl/worksheets/sheetN.xml` | komórki, wymiary, reguły warunkowe, walidacje, ochrona, **sparkline'y w `extLst`** | P / **L** (sparkline'y) |
| `xl/sharedStrings.xml` | współdzielone łańcuchy (present gdy są stringi) | P |
| `xl/styles.xml` | `fonts`, `fills`, `borders`, `cellXfs`, `dxfs` | P |
| `xl/tables/tableN.xml` | tabele Excela | P |
| `xl/drawings/drawingN.xml` | rysunki: obrazy, wykresy, **kształty i pola tekstowe** | C / **L** (kształty) |
| `xl/drawings/_rels/drawingN.xml.rels` | powiązania rysunku z mediami | P |
| `xl/charts/chartN.xml` | definicje wykresów | C (odbudowywane z modelu) |
| `xl/media/*` | obrazy, wideo | P |
| `xl/pivotTables/`, `xl/pivotCache/` | tabele przestawne i ich cache | C (odczyt definicji) |
| `xl/slicers/`, `xl/timelineCaches/` | slicery, timeline'y | **L** |
| `xl/queryTables/` | zapytania Power Query | **L** |
| `xl/model/` | model danych (Power Pivot) | **L** |
| `xl/ctrlProps/`, `xl/activeX/` | kontrolki formularza, ActiveX | **L** |
| `xl/vbaProject.bin` | projekt VBA (makra) | C (tylko `keep_vba=True`) |
| `customXml/` | własne części XML | **L** |
| `xl/calcChain.xml` | łańcuch obliczeń | P (ignorowany/odtwarzany) |
| `xl/comments*.xml`, `xl/threadedComments/` | komentarze klasyczne / **nowe wątki** | P / **L** |

**Trzy wnioski z tej tabeli.**

1. **Większość części nazywa się w sposób, który pozwala wykryć kategorię L samą nazwą** — ale **NIE wszystkie**. Sparkline'y siedzą **wewnątrz** `sheetN.xml`, więc skan po nazwach ich nie znajdzie. Musisz zajrzeć do treści (kod w `§3.13` i `§8`). To jest dokładnie ten rodzaj subtelności, dla którego warto mieć ściągawkę: **pierwsza wersja narzędzia diagnostycznego zawsze szuka po nazwach i zawsze przegapia sparkline'y**.
2. **`xl/styles.xml` rośnie w rozmiarze liniowo z liczbą unikalnych stylów.** Jeżeli plik ma 2 GB i `styles.xml` jest wielki, problemem są style, nie dane (pułapka w `§6`).
3. **Brak `sharedStrings.xml` w pliku zapisanym strumieniowo** to nie błąd. `write_only` nie buduje tabeli współdzielonych łańcuchów.

### 2.5. Glosariusz

Słownik pojęć, którymi posługuje się cały kurs i cała dokumentacja openpyxl. Ułóż go sobie **według hierarchii** (skoroszyt → arkusz → komórka), bo tak zapamiętuje się strukturę.

| Pojęcie | Definicja | W praktyce |
|---|---|---|
| **skoroszyt** (workbook) | plik `.xlsx` jako całość; obiekt `Workbook` | „zeszyt" z analogii modułu 03 |
| **arkusz** (worksheet) | jedna kartka w skoroszycie; `Worksheet` | nazwa = zakładka u dołu ekranu |
| **Chartsheet** | arkusz zawierający wyłącznie wykres | rzadko używany w automatyzacji |
| **komórka** (cell) | jedna kratka; `Cell` | ma **wartość** i **indeks stylu** |
| **zakres** (range) | prostokąt komórek; `A1:C10`; `CellRange` | tylko adnotacja i wygoda obsługi |
| **styl** (style) | komplet: font + wypełnienie + obramowanie + wyrównanie + format liczby + ochrona | 6 pieczątek razem; **niemutowalny** z perspektywy komórki |
| **styl nazwany** (named style) | styl zarejestrowany w skoroszycie pod nazwą; `NamedStyle` | `cell.style = "RaportWaluta"` |
| **dxf** (differential style) | **różnica** stylu używana w regułach warunkowych | reguła nie opisuje całego stylu, a odchylenie od niego |
| **format liczby** (`number_format`) | kod reprezentacji liczby (`#,##0.00`) | **nie zmienia wartości** — tylko to, co widzi człowiek |
| **tabela** (Table, ListObject) | obiekt arkusza: nazwa, nagłówki, filtr, styl | kontrakt z użytkownikiem, nie tylko estetyka |
| **definiowana nazwa** (defined name) | nazwa odnosząca się do adresu lub stałej | `wb.defined_names`; API zmienione w 3.1 |
| **anchor** | komórka lub punkt, do którego **przypięty** jest obraz albo wykres | obraz „pływa" nad siatką; nie jest w komórce |
| **OOXML / SpreadsheetML** | standard formatu pliku (`ISO/IEC 29500`, `ECMA-376`) | `.xlsx` = ZIP z częściami XML |
| **część** (part) | pojedynczy plik XML lub binarny wewnątrz ZIP | `xl/worksheets/sheet1.xml` |
| **relacja** (relationship) | powiązanie między częściami (`_rels/`) | bez relacji część istnieje, ale nikt do niej nie sięga |
| **sharedStrings** | tabela współdzielonych łańcunków | ten sam tekst raz, wskazywany przez wiele komórek |
| **cache** | zapisana wartość wyniku formuły albo indeks wartości | czytany przez `data_only=True` |
| **streaming** | przetwarzanie wiersz po wierszu, ze stałą pamięcią | `read_only` i `write_only`; „czytanie przez szczelinę" |
| **formula injection** | wstrzyknięcie formuły przez dane wejściowe (`=`, `+`, `-`, `@`) | sanitacja **na wejściu do domeny** |
| **zip bomb** | archiwum o małym rozmiarze i ogromnej zawartości | limit rozpakowania dla plików od użytkowników |

## 3. Przykłady krok po kroku

Trzynaście tabel ściągawki, cztery przepisy i sesja diagnostyczna. Uwaga na konwencję: **`kategoria`** w ostatniej kolumnie, tam gdzie ma to znaczenie, oznacza `P` / `C` / `L` z sekcji 1.3.

### 3.1. Tworzenie, otwieranie, zapis

```python
from io import BytesIO
from pathlib import Path
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904
```

| Wywołanie | Co robi | Uwagi / kategoria |
|---|---|---|
| `Workbook()` | nowy skoroszyt w pamięci | ma **jeden domyślny arkusz** `Sheet` → `wb.remove(wb.active)` |
| `Workbook(write_only=True)` | skoroszyt strumieniowy | brak dostępu losowego; `ws.append()` tylko w przód |
| `load_workbook("plik.xlsx")` | wczytanie w tryb normalny | **cały model w RAM** — uwaga przy dużych plikach |
| `load_workbook(..., read_only=True)` | odczyt strumieniowy | `wb.close()` obowiązkowo; brak `ws["A1"]` |
| `load_workbook(..., data_only=True)` | wartości z cache zamiast formuł | **NIGDY nie zapisuj tak wczytanego pliku** — zabijesz formuły |
| `load_workbook(..., keep_vba=True)` | zachowanie `vbaProject.bin` | wymagane dla `.xlsm`; makr **nie** edytujesz |
| `load_workbook(..., keep_links=False)` | odcięcie linków zewnętrznych | odwołania zewnętrzne mogą pokazać `#REF!` |
| `load_workbook(..., rich_text=True)` | tekst sformatowany (`CellRichText`) | nowość 3.1; zobacz `§3.2` |
| `wb.save("out.xlsx")` | zapis na dysk | **nadpisuje bez ostrzeżenia**; to nie „zapisz jako" |
| `wb.save(buf)` gdzie `buf = BytesIO()` | zapis do pamięci | potem `buf.seek(0)` przed odczytem/wysyłką |
| `wb.close()` | zwolnienie zasobów | w `read_only` **obowiązkowe** |
| `wb.properties.title = ...` | metadane dokumentu | `creator`, `description`, `keywords`, `category`, `version`, `created` |
| `wb.calculation.fullCalcOnLoad = True` | wymuszenie przeliczenia przy otwarciu | ratunek dla formuł bez cache |
| `wb.epoch = CALENDAR_WINDOWS_1900` | system dat (1900/1904) | `CALENDAR_MAC_1904` tylko dla plików z Maca |
| `wb.iso_dates = True` | daty jako ISO, nie jako liczba | **sprawdź w dokumentacji 3.1.x** — zmienia sposób przechowywania dat |
| `wb.defined_names.add(DefinedName("SUMA", attr_text="Dane!$B$2:$B$100"))` | definiowana nazwa | API 3.1; stare `add_named_range` **usunięte** |
| `wb.add_named_style(NamedStyle(name="X", ...))` | rejestracja nazwanego stylu | **przed** pierwszym `cell.style = "X"` |
| `wb.security.lockStructure = True` | ochrona struktury skoroszytu | hasło ustawiasz osobno: `wb.security.workbookPassword` |
| `wb.sheetnames` | lista nazw arkuszy | **`wb.get_sheet_names()` usunięte** w 3.x |
| `wb.active = 0` / `wb.active` | ustawienie / odczyt aktywnego arkusza | po zapisie użytkownik widzi ten arkusz |
| `wb.remove(wb["Nazwa"])` / `del wb["Nazwa"]` | usunięcie arkusza | odwołania w nazwach i tabelach **nie** są aktualizowane |
| `wb.copy_worksheet(ws)` | kopia arkusza w tym samym skoroszycie | **nie kopiuje obrazów i wykresów**; brak kopiowania między skoroszytami |
| `wb.move_sheet(ws, offset=2)` | przestawienie arkusza | `wb._sheets` to implementacja prywatna — nie ruszaj |
| `wb.index(ws)` | indeks arkusza | do porządkowania i `wb.active` |
| `wb.read_only`, `wb.template` | atrybuty stanu | `read_only` przydatne w adapterze do kontroli trybu |

**Co dzieje się w pamięci, a co w pliku.** `Workbook()` tworzy w RAM graf obiektów **bez** żadnej części XML. Dopiero `save()` uruchamia serializację: tworzy ZIP, wypisuje `[Content_Types].xml`, `_rels/`, `xl/workbook.xml`, `xl/styles.xml`, tyle części arkuszy, ile jest arkuszy, oraz `sharedStrings.xml`, jeżeli są stringi. **Kolejność arkuszy w pliku to kolejność utworzenia** (chyba że użyjesz `index` albo `move_sheet`).

### 3.2. Komórki i wartości

```python
from datetime import date, datetime, time, timedelta
from decimal import Decimal
from openpyxl.utils.datetime import from_excel, to_excel
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `ws["A1"] = 5` | zapis przez adres | czytelny dla człowieka |
| `ws.cell(row=1, column=1, value=5)` | zapis przez współrzędne | **1-based**; wygodny w pętlach |
| `ws["A1"].value` | odczyt wartości | `None` dla pustej komórki |
| `ws["A1"] = "=SUM(B1:C1)"` | zapis **formuły** | literał tekstowy; openpyxl nie policzy |
| `ws["A1"] = datetime(2026, 10, 8, 12, 30)` | data i czas | zapisane jako liczba z formatem; **bez strefy czasowej** |
| `ws["A1"] = time(14, 30)` / `timedelta(hours=27)` | czas / czas trwania | `timedelta` wymaga formatu `[h]:mm:ss` dla >24 h |
| `ws["A1"] = Decimal("1234.56")` | kwota | wewnątrz pliku i tak `float`; trzymaj groszach jako `int` |
| `cell.value = [1, 2, 3]` | **błąd** | `ValueError` — list/dict/set nie są konwertowane; użyj `str()` lub `json.dumps()` |
| `cell.data_type` | typ zawartości | `'n'` liczba, `'s'` string, `'d'` data, `'b'` bool, `'f'` formuła, `'e'` błąd |
| `cell.is_date` | czy liczba jest datą | pomocne przy odczycie plików obcego pochodzenia |
| `cell.number_format` | kod reprezentacji | nie zmienia `value` |
| limit 32 767 znaków | długość tekstu w komórce | dłuższe teksty trzeba dzielić |
| `ws.values` | generator wierszy wartości | szybkie czytanie całego arkusza |
| `from_excel(43831)`, `to_excel(dt)` | konwersja serial ↔ `datetime` | serial = liczba dni od epoki |
| `#N/A`, `#DIV/0!`, `#VALUE!` | wartości błędów | przychodzą jako **stringi**, nie wyjątki |

**Pusta komórka ≠ pusty string ≠ zero.** Dla Excela to trzy różne rzeczy: `COUNTA` liczy puste stringi, `ISBLANK` tylko prawdziwie puste; `AVERAGE` pomija puste, a traktuje `0` jako zero. Dla Ciebie: `cell.value is None` jest jedynym wiarygodnym testem pustości. **Analogia: puste krzesło na sali** — nie jest ani osobą, ani brakiem sali.

### 3.3. Wiersze, kolumny, zakresy

```python
from openpyxl.utils import (
    get_column_letter, column_index_from_string,
    quote_sheetname, absolute_coordinate, get_column_interval,
)
from openpyxl.utils.cell import (
    range_boundaries, coordinate_from_string, coordinate_to_tuple,
    rows_from_range, cols_from_range, CellRange,
)
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `ws.iter_rows(min_row, max_row, min_col, max_col, values_only=...)` | iteracja wierszami | najszybsza forma czytania |
| `ws.iter_cols(...)` | iteracja kolumnami | rzadziej potrzebna |
| `for row in ws:` | iteracja po wierszach | `row` to **krotka komórek** |
| `ws.append([...])` | dopisanie wiersza na końcu | przyjmuje **iterowalne**: `append("ABC")` wpisze `A`,`B`,`C`! |
| `ws.insert_rows(1, amount=3)` | wstawienie wierszy | **nie przesuwa referencji w formułach** |
| `ws.delete_rows(2, amount=1)` | usunięcie wierszy | formatowanie i reguły: sprawdź po fakcie |
| `ws.insert_cols`, `ws.delete_cols` | analogicznie dla kolumn | jak wyżej |
| `ws.move_range("A1:C10", rows=2, cols=3, translate=True)` | przesunięcie zakresu | `translate=True` **przesuwa referencje** — jedyna „inteligentna" operacja |
| `ws.max_row`, `ws.max_column` | wymiary „z pomiaru" | **rosną od samego formatowania**; kłamią |
| `ws.dimensions` | wymiar z metadanych pliku | może być `None` w `read_only` |
| `ws.calculate_dimension(force=True)` | przeliczenie wymiaru z danych | `force=True` omija cache wymiaru |
| `ws.reset_dimensions()` | reset zapisanych wymiarów | po modyfikacjach przez `delete_rows` |
| `get_column_letter(28)` → `"AB"` | numer → litera | 1-based |
| `column_index_from_string("AB")` → `28` | litera → numer | |
| `get_column_interval("A", "D")` → `["A","B","C","D"]` | zakres liter | |
| `quote_sheetname("Dane 2026")` → `"'Dane 2026'"` | bezpieczne odwołanie do arkusza | używaj przy budowaniu formuł i `ref` |
| `absolute_coordinate("A1:B5")` → `"$A$1:$B$5"` | adresy absolutne | potrzebne w nazwach i regułach |
| `range_boundaries("A1:C3")` → `(1, 1, 3, 3)` | `(min_col, min_row, max_col, max_row)` | **uwaga na kolejność!** — to nie `(x1,y1,x2,y2)` |
| `coordinate_from_string("AB12")` → `("AB", 12)` | adres → litera i wiersz | |
| `coordinate_to_tuple("AB12")` → `(12, 28)` | adres → `(row, column)` | **odwrotna kolejność niż w `range_boundaries`** |
| `rows_from_range("A1:C3")` | wiersze jako listy adresów | do generatorów |
| `cols_from_range("A1:C3")` | kolumny jako listy adresów | |
| `CellRange("A1:C3").rows` | obiekt zakresu z pomocnikami | `range_boundaries` używa go pod spodem |

**Najczęstsza pomyłka w tym miejscu:** `range_boundaries` zwraca `(min_col, min_row, max_col, max_row)`, a `coordinate_to_tuple` zwraca `(row, column)`. Obie funkcje są w tym samym pakiecie i **mają odwrotną kolejność argumentów**. Zapisz to sobie: `boundaries` = **kolumna najpierw** (bo tak jest w adresie `A1`), `tuple` = **wiersz najpierw** (bo tak jest w `ws.cell`).

### 3.4. Arkusze

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `wb.create_sheet(title="Dane", index=0)` | nowy arkusz | `index=0` wstawia na początek |
| `del wb["Nazwa"]` / `wb.remove(ws)` | usunięcie | odwołania w nazwach, tabelach i regułach **nie** są poprawiane |
| `ws.title = "Podsumowanie"` | zmiana nazwy | max **31 znaków**, zakaz `: \ / ? * [ ]`, unikalność bez względu na wielkość liter |
| `ws.sheet_state = "hidden"` | ukrycie zakładki | także `"visible"`, `"veryHidden"` (ukryty przed użytkownikiem, ale dostępny z kodu) |
| `ws.sheet_properties.tabColor = "1F4E78"` | kolor zakładki | tanie, bardzo poprawia nawigację |
| `ws.sheet_view.showGridLines = False` | odcięcie linii siatki | tabela ma mieć własne obramowanie |
| `ws.sheet_view.zoomScale = 90` | zbliżenie widoku | razem z `tabSelected` dla wielu arkuszy |
| `ws.freeze_panes = "B2"` | zamrożenie okien | co dokładnie się zamrozi, zależy od pozycji (od `B2` zamraża kolumnę A i wiersz 1) |
| `Chartsheet("Wykresy")` | arkusz wykresów | rzadko używany |
| `wb.active = wb.sheetnames.index("Podsumowanie")` | aktywny arkusz po otwarciu | pierwsze wrażenie użytkownika |

**Wzorzec, który zawsze się opłaca: jeden arkusz = jedna odpowiedzialność.** `Dane`, `Podsumowanie`, `Dynamika`, `Metodyka`, `_meta`. Użytkownik czyta jak dokument, a Ty wiesz, gdzie czego szukać po nazwie — bez czytania kodu.

### 3.5. Style i formaty

```python
from copy import copy
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection, NamedStyle, Color
from openpyxl.styles.differential import DifferentialStyle
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `cell.font = Font(bold=True, size=12, color="FF1F4E78")` | nadanie fontu | kolor jako **8 znaków ARGB** |
| `cell.fill = PatternFill(fill_type="solid", fgColor="FFFFC000")` | wypełnienie | dla `solid` **widoczny jest `fgColor`** |
| `cell.border = Border(left=Side(style="thin"), bottom=Side(style="medium"))` | obramowanie | `Side(style=...)`: `thin`, `medium`, `thick`, `double`, `dashed`, `dotted`, `hair` |
| `cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)` | wyrównanie i zawijanie | `wrap_text=True` **nie** ustawia wysokości wiersza |
| `cell.protection = Protection(locked=False, hidden=True)` | odblokowanie / ukrycie formuły | `hidden=True` ukrywa formułę w pasku |
| `cell.number_format = '#,##0.00 "zl"'` | format liczby | kod **nie jest walidowany** przez openpyxl |
| `nowy = copy(cell.font); nowy.bold = True; cell.font = nowy` | **modyfikacja** istniejącego stylu | jedyna poprawna droga |
| `cell.font.bold = True` | **nie działa** | obiekt jest współdzielony; indeks w komórce się nie zmienia |
| `wb.add_named_style(NamedStyle(name="RaportWaluta", number_format="..."))` | rejestracja nazwanego stylu | potem `cell.style = "RaportWaluta"` |
| `DifferentialStyle(font=..., fill=..., border=..., numFmt=..., alignment=...)` | styl dla reguły warunkowej | opisuje **różnicę**, nie komplet |
| `Color(rgb="FFFF0000")`, `Color(theme=4)`, `Color(indexed=10)` | kolor | mieszanie `theme` i `name`/`rgb` daje niespójności |
| `from openpyxl.styles import GradientFill` | wypełnienie gradientowe | na komórkach działa kapryśnie — sprawdź efekt w Excelu |

**Hierarchia kosztów**, gdy masz 100 000 komórek do sformatowania (od najtańszego):

1. **Styl nazwany** (`cell.style = "X"`) — jedno odwołanie, zero nowych obiektów.
2. **Współdzielony obiekt** tworzony raz poza pętlą (`cel.font = _FONT_NAGLOWKA`).
3. **Nowy obiekt w pętli** — marnujesz czas i pamięć; Excel i tak zdeduplikuje style w `styles.xml`.

### 3.6. Wymiary i druk

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `ws.column_dimensions["B"].width = 18` | szerokość kolumny | jednostka to „znaki" przybliżone |
| `ws.row_dimensions[1].height = 22` | wysokość wiersza | `wrap_text` nie ustawi jej automatycznie |
| `ws.column_dimensions.group("B", "D", hidden=True)` | grupowanie / zwijanie kolumn | tworzy „outline" |
| `ws.merge_cells("A1:D1")` | scalanie | **wartości z komórek poza górną-lewą są tracone** |
| `ws.unmerge_cells("A1:D1")` | rozdzielenie | nie przywraca utraconych wartości |
| `ws.merged_cells.ranges` | lista scalonych zakresów | sprawdź przed zapisem i przed komentarzami |
| `ws.freeze_panes` | zamrożenie | patrz `§3.4` |
| `ws.auto_filter.ref = "A1:H100"` | autofiltr na zakresie | alternatywa dla tabeli (tabela jest lepsza) |
| `ws.auto_filter.add_filter_column(0, ["Kraków"])` | filtr wstępny | `0` = pierwsza kolumna zakresu |
| `ws.auto_filter.add_sort_condition("D2:D100")` | sortowanie wstępne | |
| `ws.print_area = "A1:H50"` | obszar wydruku | poza nim nic się nie wydrukuje |
| `ws.print_title_rows = "1:1"` | powtarzanie nagłówka na każdej stronie | **format `"1:1"`**, nie `"1"` |
| `ws.print_title_cols = "A:A"` | powtarzanie kolumny | |
| `ws.page_setup.orientation = "landscape"` | orientacja | |
| `ws.page_setup.paperSize = ws.PAPERSIZE_A4` | format papieru | |
| `ws.sheet_properties.pageSetUpPr.fitToPage = True` | **włączenie** dopasowania | bez tego `fitToWidth` jest ignorowane |
| `ws.page_setup.fitToWidth = 1` / `fitToHeight = 0` | dopasowanie do szerokości / brak limitu stron | **trzy ustawienia razem** |
| `ws.page_margins.left = 0.5` | marginesy (cale) | `top`, `right`, `bottom`, `header`, `footer` |
| `ws.oddHeader.center.text = "Raport &D"` | nagłówek wydruku | kody: `&P` strona, `&N` liczba stron, `&D` data, `&"Arial,Bold"` czcionka |
| `ws.oddFooter.right.text = "&P / &N"` | stopka | |
| `ws.print_options.horizontalCentered = True` | wyśrodkowanie wydruku | |
| `ws.page_breaks.append(Break(id=20))` | ręczne łamanie strony | `from openpyxl.worksheet.pagebreak import Break` |

**Reguła praktyczna:** raport, który ma być drukowany, ustawia **trzy rzeczy jednocześnie** — `print_title_rows` (żeby nagłówek był na każdej stronie), `fitToPage` + `fitToWidth` (żeby kolumny się nie ucinały) i `print_area` (żeby nie drukować pustych obszarów). Każda z osobna daje dokument, który „prawie" wygląda dobrze.

### 3.7. Tabele

```python
from openpyxl.worksheet.table import Table, TableStyleInfo
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `Table(displayName="TabelaRegiony", ref="A4:D40")` | definicja tabeli | `displayName` **bez spacji i znaków specjalnych**, unikalne w **skoroszycie** |
| `tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False, showFirstColumn=False, showLastColumn=False)` | styl tabeli | nazwy stylów z listy wbudowanych Excela |
| `ws.add_table(tabela)` | dodanie tabeli do arkusza | dodaj **po** wypełnieniu danych i nagłówków |
| `ws.tables` | słownik tabel arkusza | klucz to `displayName` |
| `tabela.ref` | zakres tabeli | rozszerzasz **ręcznie** przy dopisywaniu wierszy |
| `tabela.autoFilter` | filtr tabeli | w 3.1 filtry tabeli nie są nadpisywane (naprawione w 3.1.0) |
| `tabela.headerRowCount = 1`, `tabela.totalsRowCount = 1` | wiersze nagłówka i sum | `totalsRowShown` osobno |
| ograniczenia | brak nakładania tabel, brak scalonych komórek w zakresie, `ref` z nagłówkiem | inaczej Excel zgłosi uszkodzenie lub odrzuci plik |

**Dlaczego tabela, a nie zakres z ładnym obramowaniem.** Bo tabela to **kontrakt z użytkownikiem**: ma nazwę, po której można się do niej odwołać (`=SUMA(TabelaRegiony[Sprzedaż netto])`), sama się rozszerza przy dopisywaniu wierszy, ma filtry i ma styl, który użytkownik może zmienić w dwa kliknięcia. Analogia: **zakres to stos kartek, tabela to segregator z przegrodami i etykietami.**

### 3.8. Formatowanie warunkowe

```python
from openpyxl.formatting.rule import (
    CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule,
    Top10Rule, AboveAverageRule, BelowAverageRule, DuplicateValuesRule,
    TextRule, DateRule, Rule,
)
from openpyxl.styles.differential import DifferentialStyle
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `ws.conditional_formatting.add("D2:D40", reguła)` | dodanie reguły dla zakresu | `sqref` w formacie `"D2:D40"` lub `"D2:D40 F2:F40"` dla wielu obszarów |
| `CellIsRule(operator="lessThan", formula=["0"], fill=..., font=...)` | porównanie z wartością | operatory: `lessThan`, `greaterThan`, `between`, `notBetween`, `equal`, `notEqual`, `lessThanOrEqual`, `greaterThanOrEqual` |
| `FormulaRule(formula=["$D2=\"Pilne\""], fill=...)` | reguła z formuły | **formuła BEZ znaku `=`** — najczęstsza pułapka |
| `ColorScaleRule(start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84", end_type="max", end_color="63BE7B")` | skala kolorów | typy progów: `num`, `percent`, `percentile`, `min`, `max`, `formula` |
| `DataBarRule(start_type="num", start_value=-0.5, end_type="num", end_value=0.5, color="638EC6", showValue=True)` | pasek danych | **dla wartości ujemnych ustaw oba końce jawne** |
| `IconSetRule("3TrafficLights1", "percent", [0, 33, 67], showValue=True)` | zestaw ikon | `"3Arrows"`, `"3TrafficLights1"`, `"4Rating"`, `"5Quarters"` itd. |
| `Top10Rule(rank=10, percent=False, bottom=False)` | górne/dolne wartości | `Criteria` ustawiane przez parametry |
| `AboveAverageRule()` / `BelowAverageRule()` | powyżej/poniżej średniej | |
| `DuplicateValuesRule()` / `UniqueValuesRule()` | duplikaty / unikaty | |
| `TextRule(operator, formula, text)` / `DateRule` / `TimePeriod` | tekst, daty, okresy | lista wariantów zależna od wersji — sprawdź w dokumentacji |
| `Rule(type="expression", dxf=..., priority=1)` + `DifferentialStyle` | pełna kontrola | priorytet **mniejsza liczba = ważniejsza** |
| `reguła.stopIfTrue = True` | przerwij dalsze sprawdzanie | bez tego „pierwsza pasująca" nie znaczy „jedyna" |
| `for cf in ws.conditional_formatting: print(cf.sqref, [...])` | przegląd reguł arkusza | `cf.sqref` to zakres, `cf.rules` lista |
| **Kategoria C/L** | — | openpyxl **gubi** reguły z rozszerzeń (`x14`) i część niestandardowych reguł innych narzędzi |

**W komórce nie ma koloru.** Kolor z reguły warunkowej pojawia się **dopiero w Excelu** — dlatego `cell.fill` po wczytaniu pliku pokaże `None` albo bezsensowny kolor, a Test „czy komórka jest czerwona" jest **niemożliwy** do napisania po stronie openpyxl. Analogia: **lampa z czujnikiem ruchu** — nie malujesz podłogi, tylko instalujesz urządzenie.

### 3.9. Wykresy

```python
from openpyxl.chart import (
    BarChart, LineChart, PieChart, DoughnutChart, ScatterChart, AreaChart,
    RadarChart, BubbleChart, StockChart, SurfaceChart, Reference, Series,
)
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.marker import Marker
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `Reference(ws, min_col=2, min_row=1, max_col=4, max_row=12)` | wskazanie danych | `min_col`/`min_row` **1-based** |
| `Series(values, xvalues, title_from_data=True)` | pojedyncza seria | alternatywa dla `add_data` |
| `chart.add_data(ref, titles_from_data=True)` | dodanie serii z zakresu | `titles_from_data=True` bierze nagłówek jako nazwę serii |
| `chart.set_categories(ref)` | kategorie osi X | **bez tego oś X pokaże 1, 2, 3…** |
| `BarChart(type="col", grouping="clustered")` | kolumny pionowe | `type="bar"` = poziome; `grouping`: `clustered`, `stacked`, `percentStacked`, `standard` |
| `LineChart()`, `PieChart()`, `DoughnutChart(holeSize=50)`, `ScatterChart()`, `AreaChart()`, `RadarChart()`, `BubbleChart()` | pozostałe typy | `StockChart` i `SurfaceChart` — rzadko, sprawdź w dokumentacji |
| `chart.title`, `chart.x_axis.title`, `chart.y_axis.title` | tytuły | czcionki w tytułach to obiekt `RichText` — `openpyxl.chart.text` |
| `chart.legend.position = "b"` / `chart.legend.delete = True` | legenda | pozycje: `"b"`, `"l"`, `"r"`, `"t"`, `"tr"` |
| `chart.y_axis.numFmt`, `chart.y_axis.majorGridlines = ChartLines()` | format osi i siatka | `majorGridlines = None` usuwa siatkę |
| `chart.y_axis.scaling.min/max` | zakres osi | zadziała przy `scaling.orientation` i `delete=False` |
| `chart.dataLabels = DataLabelList(); chart.dataLabels.showVal = True` | etykiety danych | `showCatName`, `showPercent`, `number_format`, `dLblPos` |
| `chart.series[0].graphicalProperties = GraphicalProperties(solidFill="1F4E78")` | kolor serii | **nie istnieje** `series.color` |
| `chart.series[0].marker = Marker(symbol="circle", size=7)` | znacznik linii | symbole: `circle`, `square`, `diamond`, `triangle`, `x`, `star`, `none` |
| `ws.add_chart(chart, "F2")` | zakotwiczenie wykresu | adres = lewa górna krawędź; wykres „pływa" nad siatką |
| `chart.width = 20`, `chart.height = 9.5` | wymiary | w centymetrach (!) |
| `drugi.y_axis.axId = 200`, `pierwszy.y_axis.crosses = "max"`, `pierwszy += drugi` | **wykres dwuosiowy** | trzy linie, w tej kolejności |
| **Kategoria C/L** | — | brak sparkline'ów; wykresy **odbudowywane z modelu**; nietypowe kombinacje serii bywają upraszczane; wykresy kombinowane z nietypową kolejnością serii mogą dać „uszkodzony plik" (problem `idx`/`order`) |

**Zapamiętaj jedną rzecz o dwuosiowości.** `axId` to **identyfikator osi**, nie „numer wykresu". Każda nowa oś wartości musi mieć **unikalny** identyfikator (typowe wartości: `100` dla pierwszej, `200` dla drugiej), a `crosses="max"` przesuwa oś pierwszej serii na krawędź wykresu. Bez tego obie serie dzielą jedną skalę, a linia marży (0,05) leży płasko na dole osi sprzedaży (500 000) — i **nie ma żadnego błędu**, tylko bezużyteczny wykres.

### 3.10. Obrazy, hiperlinki, komentarze

```python
from openpyxl.drawing.image import Image as XLImage
from openpyxl.comments import Comment
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `ws.add_image(XLImage("logo.png"), "B2")` | wstawienie obrazu | wymaga **`pillow`**; brak → `ImportError` przy imporcie `Image` |
| `img.width = 180`, `img.height = 44` | rozmiar | jednostka to **piksele obrazu przy jego DPI**, nie piksele ekranu; ustawiaj **z zachowaniem proporcji** |
| `img.anchor` | zakotwiczenie | obraz nie jest „w komórce", a **przypięty** do siatki — sortowanie go nie przenosi |
| `img.ranges` / `AnchorMarker` | zaawansowane kotwice | `OneCellAnchor`/`TwoCellAnchor` — sprawdź w dokumentacji, API niższego poziomu |
| `ws._images` | lista obrazów arkusza | użyteczne w diagnostyce, ale to atrybut wewnętrzny |
| `BytesIO` + `Image` | obraz wygenerowany w locie | np. `matplotlib` → PNG → `BytesIO` → `XLImage(buf)` |
| `cell.hyperlink = "https://..."`, `cell.value = "Otwórz"` | hiperlink zewnętrzny | sam adres bez `value` daje **puste pole** |
| `cell.hyperlink = "#Regiony!A1"` | hiperlink wewnętrzny | odwołanie do arkusza i komórki |
| `cell.style = "Hyperlink"` | wygląd linku | styl wbudowany w openpyxl |
| `cell.value = "www.firma.pl"` | **to nie hiperlink** | tylko tekst, dopóki nie ustawisz `cell.hyperlink` |
| `Comment("Treść", "Autor")`, `cell.comment = komentarz` | komentarz klasyczny | `comment.width`, `comment.height` |
| **Utrata przy zapisie** | komentarz w komórce **scalonego zakresu** | openpyxl ostrzega przy wczytywaniu i gubi przy zapisie |
| **Kategoria L** | nowe komentarze wątkowe (threaded comments) | poza modelem — sprawdź w swojej wersji |
| `wb.copy_worksheet(ws)` | kopia arkusza | **obrazy i wykresy nie są kopiowane** |

**Jednostki obrazów to najczęstsza pomyłka tego działu.** Rozmiar obrazu w openpyxl jest wyrażony w pikselach **przy DPI zapisanym w pliku obrazu**. Jeżeli Twój obraz ma 300 DPI, a ekran 96 DPI, ustawienie `img.width = 300` da obraz o innej szerokości na ekranie, niż myślisz. **Jedyna niezawodna metoda: licz proporcje.**

```python
from PIL import Image as PILImage
from openpyxl.drawing.image import Image as XLImage

with PILImage.open("assets/logo.png") as obraz:
    oryginal_w, oryginal_h = obraz.size        # szerokosc, wysokosc w pikselach pliku

docelowa_wysokosc = 44
skala = docelowa_wysokosc / oryginal_h
img = XLImage("assets/logo.png")
img.height = docelowa_wysokosc
img.width = int(oryginal_w * skala)            # TA SAMA skala dla obu wymiarow!
ws.add_image(img, "A1")
```

### 3.11. Walidacja danych i ochrona

```python
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Protection
```

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `DataValidation(type="list", formula1='"Tak,Nie"', allow_blank=True)` | lista rozwijana z literałów | **limit 255 znaków**; separator to przecinek |
| `DataValidation(type="list", formula1="=Regiony!$A$2:$A$20", allow_blank=True)` | lista z **zakresu** | **zalecana forma** dla długich list i polskich nazw |
| `DataValidation(type="whole", operator="between", formula1=1, formula2=100)` | liczby całkowite | typy: `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`, `none` |
| `DataValidation(type="custom", formula1="=LEN(A1)<=20")` | reguła własna | formuła Excela |
| `dv.promptTitle`, `dv.prompt`, `dv.errorTitle`, `dv.error`, `dv.showErrorMessage`, `dv.errorStyle` | komunikaty | `errorStyle`: `stop`, `warning`, `information` |
| `ws.add_data_validation(dv)` | **rejestracja** walidacji | bez tego obiekt istnieje i nic nie robi |
| `dv.add("B2:B100")` | przypisanie zakresu | można wołać wielokrotnie |
| `ws.data_validations` | kolekcja walidacji arkusza | iteracja: `for dv in ws.data_validations.dataValidation` |
| `ws.protection.sheet = True` | ochrona arkusza | |
| `ws.protection.password = "haslo"` | hasło ochrony | **to nie szyfrowanie** |
| `ws.protection.enable()` | ochrona + wszystkie ograniczenia | potem **wyłącz** te, które mają zostać dostępne (`autoFilter`, `sort`, `insertRows`, …) |
| `cell.protection = Protection(locked=False)` | odblokowanie komórki (formularz) | domyślnie **wszystkie komórki są zablokowane** |
| `Protection(hidden=True)` | ukrycie formuły | użytkownik nie zobaczy formuły w pasku |
| `wb.security.lockStructure = True`, `wb.security.workbookPassword = "..."` | ochrona struktury skoroszytu | |
| **Kategoria L** | szyfrowanie pliku | **openpyxl nie potrafi zapisać zaszyfrowanego `.xlsx`** — użyj narzędzia zewnętrznego |

**Trzy fakty, które trzeba znać o ochronie.** (1) To **kłódka na szufladzie, nie sejf** — chroni przed pomyłką, nie przed intencją. (2) Hasło w OOXML to słaby hash, łamany w sekundy. (3) Do poufnych danych potrzebujesz **szyfrowania pliku** (`msoffcrypto-tool` przy odczycie; LibreOffice albo osobne narzędzie przy zapisie) — a nie hasła na arkuszu.

### 3.12. Tryby wydajności

| Tryb | Jak utworzyć | Co zyskujesz | Czego nie możesz |
|---|---|---|---|
| **normalny** | `Workbook()` / `load_workbook()` | pełny dostęp losowy, wszystkie funkcje | pamięć rośnie z rozmiarem; duże pliki → GB RAM |
| **`read_only`** | `load_workbook(p, read_only=True)` | stała pamięć przy czytaniu | brak `ws["A1"]`, brak `cell()`, brak wykresów/obrazów, wymiary mogą być `None`, **brak kopiowania arkusza**, `wb.close()` obowiązkowe |
| **`write_only`** | `Workbook(write_only=True)` | stała pamięć przy pisaniu, duża szybkość | brak dostępu losowego, brak `ws["A1"]`, tabele i wykresy na tym samym arkuszu niedostępne, style przez `WriteOnlyCell` przed `append` |

```python
from openpyxl.cell import WriteOnlyCell

# Zapis 200k+ wierszy przy stalej pamieci
wb = Workbook(write_only=True)
ws = wb.create_sheet("Dane")

naglowek = []
for tekst in ("Data", "Region", "Kwota"):
    c = WriteOnlyCell(ws, value=tekst)
    c.font = _FONT_NAGLOWKA         # obiekt utworzony RAZ, poza petla
    c.fill = _FILL_NAGLOWKA         # obiekt utworzony RAZ
    naglowek.append(c)
ws.append(naglowek)

for wiersz in generator_wierszy:     # GENERATOR, nie lista!
    ws.append(list(wiersz))

wb.save("output/dane.xlsx")
```

**Techniki przyspieszenia, w kolejności zysku:**

| # | Technika | Typowy zysk |
|---|---|---|
| 1 | `write_only` zamiast normalnego przy generowaniu | 3–10× mniej pamięci, 2–5× szybciej |
| 2 | `lxml` zamiast `xml.etree` | 2–4× szybciej przy dużych plikach |
| 3 | `iter_rows(values_only=True)` zamiast komórka po komórce | 2–5× szybciej |
| 4 | Generator zamiast listy wierszy | stała pamięć |
| 5 | Brak formatowania per komórka (styli nazwane / `write_only` bez stylów) | 2–10× szybciej |
| 6 | `append` zamiast przypisywania komórek po jednej | 1,5–3× szybciej |
| 7 | Reguła warunkowa na zakresie, nie na komórkach | rozmiar pliku mniejszy o rzędy wielkości |
| 8 | Podział na pliki miesięczne/kwartalne | pamięć i czas liniowe per plik |
| 9 | Równoległość **na osobnych plikach** (`ProcessPoolExecutor`) | skalowanie; **nie współdziel `wb` między procesami** |

**Jak sprawdzić, czy `lxml` jest obecny:**

```python
import importlib.util
print("lxml:", importlib.util.find_spec("lxml") is not None)
print("pillow:", importlib.util.find_spec("PIL") is not None)
# w openpyxl istnieje znacznik wyboru parsera - sprawdz w swojej wersji:
try:
    import openpyxl
    print("openpyxl.LXML =", getattr(openpyxl, "LXML", "brak atrybutu"))
except Exception as exc:            # pragma: no cover
    print("blad importu:", exc)
```

**Kiedy openpyxl przestaje być właściwym narzędziem.** Powyżej ~1–2 mln wierszy, które trzeba **analizować**, a nie tylko przepisać, przenosisz dane do bazy albo do Parquet/CSV, a Excel zostaje **warstwą prezentacji** (agregat, wykres, podsumowanie). Analogia z modułu 19: **nie wozi się węgla ferrari**.

### 3.13. Narzędzia

| Wywołanie | Co robi | Uwagi |
|---|---|---|
| `get_column_letter(n)` / `column_index_from_string("AB")` | numer ↔ litera kolumny | 1-based |
| `quote_sheetname("Dane 2026")` | bezpieczne odwołanie do arkusza | używaj zawsze przy budowaniu formuł |
| `absolute_coordinate("A1:B5")` | adresy absolutne `$A$1:$B$5` | w definiowanych nazwach i regułach |
| `range_boundaries("A1:C3")` | `(min_col, min_row, max_col, max_row)` | **kolejność kolumna-pierwsza** |
| `coordinate_from_string` / `coordinate_to_tuple` | rozbiór adresu | `to_tuple` zwraca `(row, column)` |
| `rows_from_range` / `cols_from_range` | zakres → wiersze/kolumny adresów | |
| `get_column_interval("A", "D")` | lista liter | |
| `dataframe_to_rows(df, index=False, header=True)` | DataFrame → generator wierszy | `from openpyxl.utils.dataframe import dataframe_to_rows` |
| `from_excel` / `to_excel` | serial daty ↔ `datetime` | `openpyxl.utils.datetime` |
| `openpyxl.utils.escape` | escapowanie tekstu „formułopodobnego" | **sprawdź w dokumentacji 3.1.x** — API i zachowanie bywają zmieniane; nie polegaj na nim jako jedynej ochronie |
| `zipfile.ZipFile(p).namelist()` | inwentarz części pliku | podstawa diagnostyki i wykrywania kategorii L |
| `zipfile.ZipFile(p).getinfo(nazwa).file_size` | rozmiar części **po rozpakowaniu** | wykrywa zip bomb i „puchnięcie” stylów |
| `importlib.util.find_spec("lxml")` | obecność opcjonalnych zależności | diagnostyka środowiska |
| `os.replace(tmp, cel)` | **atomowa** podmiana pliku | tylko **ten sam wolumen** |
| `tempfile.mkstemp(dir=cel.parent)` | plik tymczasowy obok celu | `dir=` jest **kluczowe** dla atomowości |
| `shutil.copy2(oryginał, backup)` | backup z metadanymi | **`copy2`**, nie `copy` |

**Dwa narzędzia własne, których będziesz używać zawsze.** Poniżej gotowe, uruchamialne funkcje — wchodzą też do ściągawki `§8`.

```python
"""Dwa narzedzia diagnostyczne, ktore warto miec pod reka."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

from openpyxl import load_workbook


def inwentarz(sciezka: str | Path) -> dict:
    """Zwraca slownik czesci pliku i ich rozpakowanych rozmiarow.

    To pierwszy krok kazdej diagnozy: 'co w ogole jest w tym pliku?'.
    """
    sciezka = Path(sciezka)
    czesci: dict[str, int] = {}
    with zipfile.ZipFile(sciezka) as archiwum:
        for nazwa in archiwum.namelist():
            czesci[nazwa] = archiwum.getinfo(nazwa).file_size
    return {
        "plik": str(sciezka),
        "bajtow": sciezka.stat().st_size,
        "czesci": len(czesci),
        "rozpakowane": sum(czesci.values()),
        "spis": dict(sorted(czesci.items())),
    }


# Znane prefiksy czesci - pierwsza, tania warstwa detekcji.
PREFIKSY_RYZYKA = {
    "xl/slicers/":         "UTRATA: slicery",
    "xl/timelines/":       "UTRATA: timeline'y",
    "xl/queryTables/":     "UTRATA: zapytania Power Query",
    "xl/model/":           "UTRATA: model danych (Power Pivot)",
    "xl/ctrlProps/":       "UTRATA: kontrolki formularza",
    "xl/activeX/":         "UTRATA: kontrolki ActiveX",
    "xl/threadedComments/": "RYZYKO: komentarze nowego typu (sprawdz w swojej wersji)",
    "customXml/":          "UTRATA: wlasne czesci XML",
}
PREFIKSY_CZESCIOWE = {
    "xl/charts/":   "CZESCIOWO: wykresy (odbudowywane z modelu openpyxl)",
    "xl/drawings/": "CZESCIOWO: rysunki (ksztalty i pola tekstowe moga zginac)",
    "xl/pivotTables/": "CZESCIOWO: tabele przestawne (odczyt definicji)",
    "xl/media/":    "OBECNE: obrazy (sa kopiowane do pliku wynikowego)",
}


def wykryj_sparkline(sciezka: str | Path) -> bool:
    """Sparkline'y siedza WEWNATRZ XML arkusza (rozszerzenie x14).

    Dlatego NIE da sie ich wykryc po nazwie czesci - trzeba zajrzec do tresci.
    To najczestszy blad pierwszych wersji narzedzi diagnostycznych.
    """
    with zipfile.ZipFile(sciezka) as archiwum:
        for nazwa in archiwum.namelist():
            if nazwa.startswith("xl/worksheets/") and nazwa.endswith(".xml"):
                tresc = archiwum.read(nazwa).decode("utf-8", "ignore")
                if "sparklineGroups" in tresc:
                    return True
    return False


def ryzyka(sciezka: str | Path) -> tuple[str, ...]:
    """Pelny raport ryzyka utraty danych."""
    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()
        tresc_arkuszy = "".join(
            archiwum.read(n).decode("utf-8", "ignore")
            for n in nazwy
            if n.startswith("xl/worksheets/") and n.endswith(".xml")
        )

    ostrzezenia: list[str] = []
    for prefiks, opis in PREFIKSY_RYZYKA.items():
        if any(n.startswith(prefiks) for n in nazwy):
            ostrzezenia.append(opis)
    for prefiks, opis in PREFIKSY_CZESCIOWE.items():
        if any(n.startswith(prefiks) for n in nazwy):
            ostrzezenia.append(opis)
    if "xl/vbaProject.bin" in nazwy:
        ostrzezenia.append("MAKRA: obecne (wczytuj z keep_vba=True, inaczej znikna)")
    if "sparklineGroups" in tresc_arkuszy:
        ostrzezenia.append("UTRATA: sparkline'y (openpyxl nie ma dla nich API)")
    return tuple(ostrzezenia)


def podsumowanie_skoroszytu(sciezka: str | Path) -> dict:
    """Krotki opis struktury - do testow regresji i szybkiej diagnozy."""
    wb = load_workbook(sciezka, read_only=True, data_only=False)
    try:
        return {
            "arkusze": tuple(wb.sheetnames),
            "ukryte": tuple(n for n in wb.sheetnames
                            if wb[n].sheet_state != "visible"),
            "wymiary": {n: wb[n].calculate_dimension() for n in wb.sheetnames},
            "z_formulami": {
                n: sum(
                    1
                    for rzad in wb[n].iter_rows()
                    for c in rzad
                    if c.data_type == "f"
                )
                for n in wb.sheetnames
            },
        }
    finally:
        wb.close()


def _czesci_do_porownania(sciezka: Path) -> dict[str, str]:
    """Czesci XML ze stabilna trescia - do porownania dwoch plikow."""
    istotne = ("xl/worksheets/", "xl/styles.xml", "xl/tables/", "xl/workbook.xml")
    wynik: dict[str, str] = {}
    with zipfile.ZipFile(sciezka) as archiwum:
        for nazwa in archiwum.namelist():
            if nazwa.endswith(".xml") and nazwa.startswith(istotne):
                tresc = archiwum.read(nazwa).decode("utf-8", "ignore")
                # normalizacja: usun biale znaki miedzy znacznikami
                wynik[nazwa] = hashlib.sha256(
                    "".join(tresc.split()).encode("utf-8")
                ).hexdigest()
    return wynik


def porownaj_skoroszyty(a: str | Path, b: str | Path) -> dict:
    """Porownanie po CZESCIACH XML, nie bajt w bajt.

    UWAGA: hasz calego pliku .xlsx NIE jest deterministyczny miedzy
    uruchomieniami, bo ZIP zapisuje znaczniki czasu wpisow. Dlatego
    deterministycznosc mierzymy na zawartosci, nie na bajtach pliku.
    """
    czesci_a = _czesci_do_porownania(Path(a))
    czesci_b = _czesci_do_porownania(Path(b))
    wspolne = set(czesci_a) & set(czesci_b)
    return {
        "tylko_w_a": sorted(set(czesci_a) - set(czesci_b)),
        "tylko_w_b": sorted(set(czesci_b) - set(czesci_a)),
        "rozne": sorted(n for n in wspolne if czesci_a[n] != czesci_b[n]),
        "identyczne": sorted(n for n in wspolne if czesci_a[n] == czesci_b[n]),
    }
```

### 3.14. Cztery przepisy, których nie trzeba wymyślać

**Przepis 1 — bezpieczny zapis z backupem i atomowością.**

```python
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path


class BladZapisu(RuntimeError):
    """Blad zapisu pliku wynikowego."""


def bezpieczny_zapis(wb, cel: str | Path, *, backup: bool = True) -> Path:
    """Zapisuje skoroszyt atomowo: tmp -> weryfikacja -> backup -> replace."""
    cel = Path(cel)
    cel.parent.mkdir(parents=True, exist_ok=True)

    # Plik tymczasowy MUSI byc w tym samym katalogu co cel (ten sam wolumen).
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{cel.stem}.", suffix=".tmp.xlsx", dir=cel.parent
    )
    os.close(fd)
    tmp = Path(tmp_name)

    try:
        wb.save(tmp)

        # Tania weryfikacja: czy plik sie otwiera i ma tyle arkuszy, ile ma miec?
        from openpyxl import load_workbook
        sprawdzony = load_workbook(tmp, read_only=True)
        try:
            if len(sprawdzony.sheetnames) != len(wb.sheetnames):
                raise BladZapisu("weryfikacja tmp: niezgodna liczba arkuszy")
        finally:
            sprawdzony.close()

        if backup and cel.exists():
            znacznik = datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(cel, cel.with_name(f"{cel.stem}.{znacznik}.bak{cel.suffix}"))

        os.replace(tmp, cel)          # COMMIT - jeden ruch systemowy
        return cel

    except PermissionError as blad:
        raise BladZapisu(
            f"Plik {cel.name} jest otwarty w innym programie (najczesciej Excel). "
            f"Zamknij go i uruchom ponownie albo zapisz pod inna nazwa."
        ) from blad
    finally:
        tmp.unlink(missing_ok=True)   # ZAWSZE, takze przy KeyboardInterrupt
```

**Przepis 2 — inwentarz i raport przed modyfikacją cudzego pliku.**

```python
from pathlib import Path


def raport_przed_modyfikacja(sciezka: str | Path) -> str:
    """Krotki raport: co jest w pliku i co moze zginac. Wypisz PRZED zapisem."""
    sciezka = Path(sciezka)
    inv = inwentarz(sciezka)
    ostrzezenia = ryzyka(sciezka)

    linie = [
        f"PLIK: {sciezka.name} ({inv['bajtow'] / 1024:.0f} KB, "
        f"{inv['czesci']} czesci, {inv['rozpakowane'] / 1024:.0f} KB po rozpakowaniu)",
        "-" * 60,
    ]
    if ostrzezenia:
        linie.append("RYZYKO UTRATY / OGRANICZEN:")
        linie += [f"  * {o}" for o in ostrzezenia]
    else:
        linie.append("Nie wykryto typowych czesci zagrozonych utrata.")

    linie.append("-" * 60)
    linie.append("WNIOSEK:")
    linie.append(
        "  Jezeli plik byl tworzony recznie w Excelu i widzisz tu pozycje "
        "oznaczone jako UTRATA - NIE przepuszczaj go przez openpyxl. "
        "Wygeneruj nowy plik albo oddaj modyfikacje Excelowi."
        if ostrzezenia
        else "  Plik wyglada na technicznie czysty - modyfikacja jest "
             "akceptowalna, ale zachowaj kopie i sprawdz wynik okiem."
    )
    return "\n".join(linie)
```

**Przepis 3 — plik do pamięci i do odpowiedzi HTTP.**

```python
from io import BytesIO

from openpyxl import Workbook


def skoroszyt_do_bajtow(wb: Workbook) -> bytes:
    """Serializuje skoroszyt do bajtow. Cel: HTTP, e-mail, test, kolejka."""
    bufor = BytesIO()
    wb.save(bufor)              # openpyxl przyjmuje obiekt plikopodobny
    return bufor.getvalue()


# --- Django ---
# from django.http import HttpResponse
# odpowiedz = HttpResponse(
#     skoroszyt_do_bajtow(wb),
#     content_type=(
#         "application/vnd.openxmlformats-officedocument"
#         ".spreadsheetml.sheet"
#     ),
# )
# odpowiedz["Content-Disposition"] = 'attachment; filename="raport.xlsx"'

# --- Flask ---
# from flask import send_file
# return send_file(BytesIO(skoroszyt_do_bajtow(wb)),
#                  mimetype="application/vnd.openxmlformats-officedocument"
#                           ".spreadsheetml.sheet",
#                  as_attachment=True, download_name="raport.xlsx")

# --- FastAPI ---
# from fastapi.responses import StreamingResponse
# return StreamingResponse(iter([skoroszyt_do_bajtow(wb)]),
#                          media_type="application/vnd.openxmlformats-"
#                                     "officedocument.spreadsheetml.sheet",
#                          headers={"Content-Disposition":
#                                   'attachment; filename="raport.xlsx"'})
```

**Uwaga, która ratuje pół dnia:** jeżeli używasz **jednego** `BytesIO` do zapisu i potem do odczytu (np. w teście), **zawsze `buf.seek(0)`** przed `load_workbook(buf)`. Bez tego dostaniesz `zipfile.BadZipFile: File is not a zip file`, mimo że plik jest doskonały — bo wskaźnik stoi na końcu.

**Przepis 4 — scalenie wielu plików w jeden arkusz.**

```python
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell


def scal_arkusze(pliki: list[Path], cel: Path, nazwa_arkusza: str = "Dane") -> int:
    """Scala wiersze danych z wielu plikow w jeden arkusz, strumieniowo.

    Zalozenie: pierwszy wiersz kazdego arkusza to naglowek (pomijamy go
    poza pierwszym plikiem). Pamięć pozostaje stala.
    """
    wb_out = Workbook(write_only=True)
    ws_out = wb_out.create_sheet(nazwa_arkusza)
    naglowek_zapisany = False
    licznik = 0

    for plik in pliki:
        wb_in = load_workbook(plik, read_only=True, data_only=True)
        try:
            for arkusz in wb_in.worksheets:
                for numer, rzad in enumerate(arkusz.iter_rows(values_only=True)):
                    if numer == 0:
                        if naglowek_zapisany:
                            continue          # naglowek tylko raz
                        naglowek_zapisany = True
                    ws_out.append(list(rzad))
                    licznik += 1
        finally:
            wb_in.close()                     # OBOWIAZKOWE w read_only

    wb_out.save(cel)
    return licznik
```

**Dwie uwagi.** (1) `data_only=True` jest tu **bezpieczne**, bo nic nie zapisujemy z powrotem — czytamy tylko wartości. (2) Kolumny muszą mieć **tę samą kolejność** we wszystkich plikach. Jeżeli nie mają, to nie jest problem openpyxl, tylko problem **kontraktu danych** — i wymaga mapowania po **nazwach nagłówków**, nie po pozycjach.

### 3.15. Sesja diagnostyczna — „plik wygląda dobrze, a sumy pokazują zero"

Realny scenariusz, krok po kroku, **w kolejności z sekcji 2.3**. Zakładamy, że dostałeś plik `sprzedaz_2026-09.xlsx` wygenerowany przez Twój skrypt, klientka mówi: „otworzyłam, sumy na dole pokazują 0,00 zł".

**Krok 1 — Excel, widok formuł (`Ctrl+`` ` ``).**

Widzisz w komórkach D2:D200 coś takiego: `1 234,56` **bez** znaku `=` na początku. Gdyby były tam formuły, widziałbyś `=B2*C2`. **Wniosek: w komórkach są wartości, nie formuły.** Problem nie leży w cache ani w silniku obliczeń. Zawężenie: coś jest nie tak z **wartościami** albo z **ich typem**.

**Krok 2 — podsumowanie struktury.**

```python
>>> podsumowanie_skoroszytu("sprzedaz_2026-09.xlsx")
{'arkusze': ('Regiony', 'Dynamika', '_meta'),
 'ukryte': ('_meta',),
 'wymiary': {'Regiony': 'A4:D48', 'Dynamika': 'A1:D13', '_meta': 'A1:B7'},
 'z_formulami': {'Regiony': 46, 'Dynamika': 0, '_meta': 0}}
```

Zwróć uwagę: `Regiony` ma **46 komórek z formułami** — czyli suma na dole **jest** formułą. To zmienia diagnozę: suma jest formułą, ale liczy zero. To znaczy, że **zakres sumy jest pusty albo zawiera tekst**. Idziemy dalej.

**Krok 3 — typy wartości.**

```python
>>> from openpyxl import load_workbook
>>> wb = load_workbook("sprzedaz_2026-09.xlsx")
>>> ws = wb["Regiony"]
>>> for adres in ("C5", "C6", "C7"):
...     c = ws[adres]
...     print(adres, repr(c.value), type(c.value).__name__, c.data_type)
C5 '1 234,56' str s
C6 '987,00' str s
C7 1234.56 float n
```

**Znalazłeś.** Kolumny C5 i C6 to **tekst**, nie liczby — bo w CSV były zapisane jako `"1 234,56"` (z separatorem tysięcy i polskim przecinkiem dziesiętnym), a Twój kod zapisał je bez konwersji. Wiersz C7 działa, bo w tym wierszu dane w CSV były „czyste". `SUM` w Excelu **ignoruje tekst** i zwraca zero. **To jest przyczyna, a nie objaw.**

**Krok 4 — inwentarz (na wypadek, gdyby problem był innego rodzaju).**

```python
>>> for k, v in inwentarz("sprzedaz_2026-09.xlsx")["spis"].items():
...     print(f"{v:>10} {k}")
      ...
   1934 xl/styles.xml
  15221 xl/worksheets/sheet1.xml
```

Brak `xl/media/`, brak `xl/charts/` → **brak ryzyka utraty danych** przy zapisie. Ten plik jest „technicznie czysty" i można go bezpiecznie modyfikować. Dobrze wiedzieć — ale to **nie była** przyczyna.

**Krok 5 — porównanie z plikiem wzorcowym** (jeżeli masz plik z poprzedniego miesiąca, który działał):

```python
>>> porownaj_skoroszyty("sprzedaz_2026-08.xlsx", "sprzedaz_2026-09.xlsx")
{'tylko_w_a': [], 'tylko_w_b': [],
 'rozne': ['xl/worksheets/sheet1.xml'],
 'identyczne': ['xl/styles.xml', 'xl/workbook.xml']}
```

Różnica jest **w danych arkusza**, a nie w stylach czy strukturze. To potwierdza diagnozę z kroku 3.

**Krok 6 — środowisko.** Wersja openpyxl, `lxml`, `pillow`, locale — nieistotne w tym przypadku, bo problem jest w danych. Wykonujemy ten krok **mimo to**, bo jeśli pominiemy go „bo wiemy", to przy następnej awarii zaczniemy od kroku 6 zamiast od kroku 1.

**Naprawa — i uwaga, gdzie jest prawdziwe miejsce naprawy.** Konwersja należy do **domeny** (moduł 26/27), nie do adaptera openpyxl:

```python
def na_kwote(tekst: object) -> float:
    """'1 234,56' -> 1234.56. Miejsce: domain/reguly.py."""
    if tekst is None or tekst == "":
        raise BladDomeny("pusta kwota")
    if isinstance(tekst, (int, float)):
        return float(tekst)
    czysty = (
        str(tekst)
        .replace("\u00a0", "")          # twarda spacja
        .replace(" ", "")
        .replace("zl", "")
        .replace("PLN", "")
        .strip()
        .replace(",", ".")
    )
    try:
        return float(czysty)
    except ValueError as blad:
        raise BladDomeny(f"nie umiem odczytac kwoty: {tekst!r}") from blad
```

**Trzy wnioski z tej sesji.** (1) Kolejność kroków **naprawdę** oszczędza czas — krok 1 wykluczył połowę możliwych przyczyn w dziesięć sekund. (2) Krok 4 nie był potrzebny diagnostycznie, ale był potrzebny **ryzykowo** — i to jest wartość, którą zauważasz dopiero po fakcie. (3) Naprawa trafiła do **domeny**, nie do adaptera — bo to reguła biznesowa („jak wygląda kwota w polskim pliku CSV"), a nie szczegół Excela.

### 3.16. FAQ — dwadzieścia pytań, na które wraca się najczęściej

**1. Jak policzyć formuły bez Excela?**
Nie policzysz ich openpyxl. Cztery drogi (w kolejności od najlepszej): (a) policz w Pythonie — to logika biznesowa; (b) `wb.calculation.fullCalcOnLoad = True` i oddaj plik użytkownikowi; (c) przelicz przez LibreOffice headless: `soffice --headless --convert-to xlsx --outdir out plik.xlsx` — to zapisze świeży cache; (d) biblioteka z silnikiem formuł (`formulas`, `pycel`) — działa dla prostych przypadków, zawodzi na nowoczesnych funkcjach, tabelach i formułach tablicowych.

**2. Jak zmienić istniejący plik i nic nie stracić?**
Nie ma jednej odpowiedzi — jest **rytuał**. (a) `raport_przed_modyfikacja()` i przeczytaj ostrzeżenia. (b) Nigdy nie zapisuj wczytanego z `data_only=True`. (c) Zapisz do **nowego** pliku, nigdy nie nadpisuj wejścia. (d) Jeżeli plik był tworzony ręcznie w Excelu i ma elementy z kategorii L — **nie przepuszczaj go przez openpyxl**, tylko generuj nowy. (e) Otwórz wynik w Excelu i **obejrzyj**. (f) Zachowaj kopię oryginału.

**3. Jak przyspieszyć zapis 1 000 000 wierszy?**
`Workbook(write_only=True)` + `lxml` + generator + brak formatowania per komórka + `append` zamiast komórek pojedynczo. Jeżeli to nie wystarcza — **zmień narzędzie**: CSV albo Parquet do danych, Excel tylko do agregatów i wykresów. Trzy tysiące razy „ale to tylko Excel" nie zastąpi bazy danych.

**4. Czy openpyxl działa bez Excela?**
Tak — jest czysto pythonowy i **nie potrzebuje ani Excela, ani LibreOffice**. To największa zaleta na serwerze: działa w kontenerze, w CI, na Linuksie bez GUI. Konsekwencja jest druga strona tego medalu: **nie ma silnika obliczeń**.

**5. Czy mogę tworzyć sparkline'y?**
Nie. openpyxl nie ma API do sparkline'ów i **gubi je przy zapisie** (są zapisane w rozszerzeniu `x14` wewnątrz XML arkusza). Jeżeli sparkline'y są wymaganiem, użyj `xlsxwriter` (ma `worksheet.add_sparkline()`, ale **nie czyta i nie modyfikuje** istniejących plików) albo zamień wymaganie na pasek danych (`DataBarRule`).

**6. Jak zrobić listę rozwijaną z innego arkusza?**
Walidacja z zakresu, nie z literałów:
```python
dv = DataValidation(
    type="list",
    formula1="=Regiony!$A$2:$A$20",   # nazwa arkusza BEZ apostrofow, chyba ze ma spacje
    allow_blank=True, showErrorMessage=True,
    errorTitle="Nieznany region", error="Wybierz region z listy.",
)
ws.add_data_validation(dv)             # PAMIETAJ o rejestracji!
dv.add("B5:B100")
```
Jeśli nazwa arkusza ma spacje: `formula1="='Moje dane'!$A$2:$A$20"` (z apostrofami). Lista literałowa ma **limit 255 znaków**.

**7. Jak wstawić wykres z dwóch osi?**
```python
linia.y_axis.axId = 200          # unikalne ID drugiej osi wartosci
linia.y_axis.title = "Marza (%)"
linia.y_axis.delete = False      # upewnij sie, ze os jest widoczna
slupki.y_axis.crosses = "max"    # os pierwszej serii przy krawedzi
slupki += linia                  # sklej serie w jeden wykres
ws.add_chart(slupki, "F2")
```

**8. Jak zapisać plik do pamięci / odpowiedzi HTTP?**
`BytesIO` + `wb.save(bufor)` + `bufor.getvalue()`. Wzorce dla Django/Flask/FastAPI w `§3.14`, przepis 3. **Zawsze `seek(0)`** przed ponownym odczytem tego samego bufora.

**9. Jak połączyć arkusze z wielu plików?**
`load_workbook(..., read_only=True)` w pętli + `wb_out.append(...)` do `write_only`. Kod w `§3.14`, przepis 4. Uwaga: kolumny muszą mieć tę samą kolejność — mapuj po **nazwach nagłówków**.

**10. Jak zachować makra?**
`load_workbook("plik.xlsm", keep_vba=True)`, zapisz **pod tym samym rozszerzeniem** `.xlsm`. Makr **nie** odczytasz ani nie zmodyfikujesz (openpyxl traktuje `vbaProject.bin` jako nieprzezroczystą część binarną). Jeżeli zapiszesz plik z makrami jako `.xlsx`, Excel go otworzy, ale makra znikną.

**11. Jak nadać hasło?**
Dwa różne mechanizmy, o których trzeba wiedzieć: **ochrona struktury** (`ws.protection.password`, `wb.security.workbookPassword` — słaba, nie szyfruje) i **szyfrowanie pliku** (openpyxl **nie potrafi** zapisać zaszyfrowanego `.xlsx`). Do szyfrowania użyj zewnętrznego narzędzia; do odczytu zaszyfrowanego pliku — `msoffcrypto-tool`. Nie obiecuj klientowi, że hasło na arkuszu chroni dane. **Nie chroni.**

**12. Jak walidować pliki od użytkowników?**
Trzy warstwy: (a) **rozmiar i typ** — sprawdź rozszerzenie, `mimetype` i **limit rozpakowania** (zip bomb: `sum(getinfo(n).file_size for n in namelist())`, zanim cokolwiek wczytasz); (b) **treść** — waliduj każdą wartość tekstową sanitizerem z `§6` (objaw `=`/`+`/`-`/`@`), nigdy nie ufaj typom; (c) **kontekst** — nie przetwarzaj plików od użytkowników w tym samym katalogu, w którym zapisujesz wyniki, i nie loguj surowych danych.

**13. Jak zaktualizować plik, gdy jest otwarty w Excelu?**
Na Windows dostaniesz `PermissionError` przy `os.replace`. Trzy opcje: (a) zapisz pod **inną nazwą** (użytkownik zobaczy wynik i zamknie stary plik); (b) spróbuj ponownie z odstępem i **skończonym** limitem prób; (c) **jawny komunikat** z instrukcją (wzorzec w `§3.14`, przepis 1). Na Linuksie podmiana „przejdzie", ale użytkownik nadal patrzy na **starą treść** w otwartym oknie — więc „się udało" nie znaczy „użytkownik to widzi".

**14. Jak porównać dwa skoroszyty?**
Nie bajt w bajt — **nigdy**, bo ZIP zapisuje znaczniki czasu wpisów, więc dwa pliki z identyczną treścią mają różne hasze. Porównuj po **częściach XML** z normalizacją białych znaków (`porownaj_skoroszyty()` z `§3.13`) albo po **modelu**: wczytaj oba i porównaj komórki, zakresy, liczbę reguł, `ref` tabel.

**15. Jak migrować z `xlrd`/`xlwt`/`xlsxwriter`/`openpyxl 2.x`?**
Tabele w `§4.4`. W skrócie: `xlrd` czyta tylko `.xls` (od 2.0 **nie** czyta `.xlsx`) — jeżeli go używasz do `.xlsx`, to Twój kod „działa" przez pomyłkę albo wcale; `xlwt` pisze tylko `.xls`; `xlsxwriter` **nie czyta** i **nie modyfikuje**; `openpyxl 2.x` różni się nazwami wielu metod (usunięte `get_sheet_names`, `add_named_range`, `save_virtual_workbook`, `openpyxl.pandas`).

**16. Formuły pokazują `0` albo stare wartości po otwarciu. Co zrobić?**
Trzy przyczyny w kolejności prawdopodobieństwa: (a) **brak cache** — nowa formuła nigdy nie była policzona → ustaw `wb.calculation.fullCalcOnLoad = True`, żeby Excel przeliczył przy otwarciu; (b) **tryb automatycznego przeliczania wyłączony** w Excelu (Ręczne) — to po stronie użytkownika; (c) **stary cache** w pliku, który wczytałeś i zapisałeś — jeżeli robisz read-modify-write, cache wartości nie jest odświeżany przez openpyxl, bo **nie ma czym** go odświeżyć.

**17. Jak scalić komórki bez utraty danych?**
Najpierw **przepisz** wartości do komórki górnej-lewej, **potem** scalaj. Przy scalaniu openpyxl zachowuje wartość górnej-lewej, a pozostałe zostaną odrzucone przez Excela. Jeżeli czytasz plik ze scalonymi komórkami: `ws.merged_cells.ranges` daje listę zakresów, a wartości szukaj **tylko** w górnej-lewej komórce każdego zakresu — w pozostałych będzie `None`.

**18. Jak znaleźć ostatni wiersz z danymi, skoro `max_row` kłamie?**
Skanuj **w górę** od `max_row` i sprawdzaj, czy wiersz ma jakąkolwiek niepustą wartość:
```python
def ostatni_wiersz_danych(ws, kolumny: range | None = None) -> int:
    """Ignoruje komorki puste, ale sformatowane (max_row ich nie ignoruje)."""
    kolumny = kolumny or range(1, ws.max_column + 1)
    for numer in range(ws.max_row, 0, -1):
        if any(ws.cell(row=numer, column=k).value is not None for k in kolumny):
            return numer
    return 0
```
W trybie `read_only` nie ma dostępu losowego — musisz **przejść przez `iter_rows()` do końca** i zapamiętać ostatni niepusty.

**19. Jak usunąć formatowanie warunkowe albo styl z komórki?**
Reguły:
```python
del ws.conditional_formatting["D2:D40"]      # usun zakres reguł
```
Styl komórki „resetujesz", przypisując nowy, pusty obiekt: `cell.font = Font()`, `cell.fill = PatternFill()`, `cell.border = Border()`, `cell.number_format = "General"`. **Nie da się** „odkręcić" stylu przez mutację. Przy nazwanym stylu: `cell.style = "Normal"`.

**20. Jak odróżnić pustą komórkę od zera i od pustege stringa?**
```python
c = ws["A1"]
if c.value is None:            # prawdziwie pusta
    ...
elif c.value == "":            # pusty string - dla Excela to TEZ tekst
    ...
elif c.value == 0:             # zero - wartosc liczbowa
    ...
```
Dla formuł `COUNTA`, `AVERAGE` i `COUNT` te trzy przypadki zachowują się **różnie**. Jeżeli Twoja logika biznesowa tego nie rozróżnia, to **dane nie mają znaczenia**, a tylko wyglądają.

## 4. Anatomia API

### 4.1. Sygnatury, które warto znać na pamięć

| Wywołanie | Co robi | Najważniejsze parametry | Uwagi |
|---|---|---|---|
| `load_workbook(filename, read_only=False, keep_vba=False, data_only=False, keep_links=True, rich_text=False)` | wczytanie pliku | wszystkie pięć flag zmienia zachowanie **fundamentalnie** | `data_only=True` → **nie zapisuj**; `keep_vba=True` przy `.xlsm` |
| `Workbook(write_only=False)` / `Workbook()` | nowy skoroszyt | — | `write_only=True` → tylko `append` |
| `wb.save(filename)` | zapis | ścieżka lub obiekt plikopodobny | **COMMIT**; nadpisuje bez ostrzeżenia |
| `ws.cell(row=, column=, value=)` | dostęp do komórki | `row`, `column` **1-based** | `ws["A1"]` = to samo, inaczej zapisane |
| `ws.iter_rows(min_row=, max_row=, min_col=, max_col=, values_only=False)` | iteracja wierszami | `values_only=True` → krotki wartości | najszybszy sposób czytania |
| `ws.append(iterable)` | dopisanie wiersza | przyjmuje **iterowalne** | `append("ABC")` = trzy komórki! |
| `ws.insert_rows(idx, amount=1)` / `ws.delete_rows(idx, amount=1)` | modyfikacja struktury | — | **nie aktualizuje referencji w formułach** |
| `ws.move_range(ref, rows=0, cols=0, translate=False)` | przesunięcie zakresu | `translate=True` | jedyna operacja „inteligentna" |
| `ws.merge_cells(ref)` / `ws.unmerge_cells(ref)` | scalanie | zakres `"A1:D1"` | traci wartości poza górną-lewą |
| `ws.freeze_panes = "B2"` | zamrożenie okien | adres | co się zamrozi, zależy od pozycji |
| `ws.add_data_validation(dv)` + `dv.add(sqref)` | walidacja danych | — | **bez `add_data_validation` obiekt nic nie robi** |
| `ws.conditional_formatting.add(sqref, rule)` | reguła warunkowa | zakres i reguła **osobno** | `FormulaRule` bez znaku `=` |
| `ws.add_table(table)` | tabela | `displayName`, `ref` | nazwa unikalna w skoroszycie |
| `ws.add_chart(chart, anchor)` / `ws.add_image(img, anchor)` | wykres / obraz | `anchor` = adres komórki | oba „pływają" nad siatką |
| `wb.create_sheet(title=, index=)` | nowy arkusz | `index=0` = na początek | max 31 znaków nazwy |
| `wb.copy_worksheet(ws)` | kopia arkusza | — | **bez obrazów i wykresów** |
| `wb.add_named_style(style)` | nazwany styl | `NamedStyle` | **przed** użyciem nazwy |
| `wb.defined_names.add(DefinedName(name, attr_text=...))` | definiowana nazwa | `attr_text` = zakres | API 3.1 |
| `wb.close()` | zwolnienie zasobów | — | obowiązkowe w `read_only` |
| `WriteOnlyCell(ws, value=)` | komórka w `write_only` | — | nadaj style **przed** `append` |

### 4.2. Inwentarz: trzy kategorie dla każdego obszaru

| Obszar | **P — potrafi** | **C — częściowo** | **L — gubi przy zapisie** |
|---|---|---|---|
| Wartości i typy | tekst, liczby, bool, daty, `timedelta`, `Decimal` | `Decimal` → `float` (utrata precyzji), brak `tzinfo` | — |
| Formuły | zapis literału formuły | odczyt tylko cache | cache wartości (nie jest odświeżany) |
| Style | font, fill, border, alignment, ochrona, `number_format`, nazwane style | gradienty na komórkach, motywy (`theme`) | — |
| Formatowanie warunkowe | wszystkie podstawowe typy reguł | reguły z rozszerzeń `x14` | nierozpoznane reguły innych narzędzi |
| Tabele | tworzenie, styl, filtr | filtry przy modyfikacji (naprawione w 3.1.0) | — |
| Wykresy | tworzenie wszystkich głównych typów | **odbudowywane z modelu** — nietypowe formatowanie uproszczone | sparkline'y |
| Obrazy | wstawianie, `BytesIO` | pozycjonowanie zaawansowane (anchory), alt text | — |
| Komentarze | klasyczne | rozmiar/popozycja | komentarze w scalonych komórkach, komentarze wątkowe |
| Ochrona | hasło na arkuszu i strukturze | lista uprawnień per operacja | szyfrowanie pliku (brak obsługi) |
| Wydajność | `read_only`, `write_only`, `WriteOnlyCell` | brak `ws["A1"]` w `read_only` | — |
| Makra | zachowanie przez `keep_vba=True` | brak edycji VBA | makra przy braku `keep_vba` |
| Metadane | `wb.properties`, własne właściwości (3.1) | `created`/`modified` zależne od systemu | — |
| Rozszerzenia | — | — | slicery, timeline'y, Power Query, model danych, kontrolki, `customXml` |

### 4.3. Narzędzia `openpyxl.utils` — kompletny przegląd najczęściej używanych

| Funkcja | Wejście | Wyjście | Typowy użytek |
|---|---|---|---|
| `get_column_letter(28)` | `int` | `"AB"` | budowanie `ref` i formuł |
| `column_index_from_string("AB")` | `str` | `28` | rozbiór adresu z konfiguracji |
| `get_column_interval("A", "D")` | `str, str` | `["A","B","C","D"]` | iteracja po kolumnach konfiguracji |
| `quote_sheetname("Dane 2026")` | `str` | `"'Dane 2026'"` | **zawsze** przy budowaniu formuł |
| `absolute_coordinate("A1:B5")` | `str` | `"$A$1:$B$5"` | definiowane nazwy, reguły |
| `range_boundaries("A1:C3")` | `str` | `(1, 1, 3, 3)` = `(min_col, min_row, max_col, max_row)` | odczyt zakresu z konfiguracji |
| `range_to_tuple("'Ark 1'!A1:C3")` | `str` | `(sheetname, (min_col, min_row, max_col, max_row))` | zakres z nazwą arkusza |
| `coordinate_from_string("AB12")` | `str` | `("AB", 12)` | rozbiór na kolumnę i wiersz |
| `coordinate_to_tuple("AB12")` | `str` | `(12, 28)` = `(row, column)` | gotowe do `ws.cell(*tuple)` |
| `rows_from_range("A1:C3")` | `str` | generator list adresów, wierszami | pętle po adresach |
| `cols_from_range("A1:C3")` | `str` | generator list adresów, kolumnami | j.w. |
| `CellRange("A1:C3")` | `str` | obiekt z `.rows`, `.cols`, `.bounds` | elastyczna praca z zakresem |
| `dataframe_to_rows(df, index=False, header=True)` | `DataFrame` | generator list | pandas → arkusz |
| `from_excel(43831)`, `to_excel(dt)` | liczba / `datetime` | odwrotność | diagnostyka dat |
| `openpyxl.LXML` (jeśli istnieje w Twojej wersji) | — | `bool` | sprawdzenie parsera XML |

### 4.4. Migracja — tabela odpowiedników

**Z `xlrd` (odczyt `.xls`, historycznie też `.xlsx`):**

| `xlrd` | `openpyxl` | Uwaga |
|---|---|---|
| `xlrd.open_workbook(path)` | `load_workbook(path)` | `xlrd >= 2.0` **nie czyta** `.xlsx` |
| `wb.sheet_names()` | `wb.sheetnames` | atrybut, nie metoda |
| `wb.sheet_by_index(0)` / `sheet_by_name("X")` | `wb.worksheets[0]` / `wb["X"]` | |
| `sh.nrows`, `sh.ncols` | `ws.max_row`, `ws.max_column` | `max_row` rośnie od formatowania |
| `sh.cell_value(r, c)` | `ws.cell(row=r+1, column=c+1).value` | **uwaga na przesunięcie o 1!** |
| `sh.row_values(r)` | `[c.value for c in ws[r+1]]` | |
| `sh.cell_type(...)` | `ws.cell(...).data_type` | inne oznaczenia typów |

**Z `xlwt` (zapis `.xls`):**

| `xlwt` | `openpyxl` | Uwaga |
|---|---|---|
| `xlwt.Workbook()` | `Workbook()` | |
| `wb.add_sheet("Nazwa")` | `wb.create_sheet("Nazwa")` | |
| `sheet.write(r, c, val)` | `ws.cell(row=r+1, column=c+1, value=val)` | przesunięcie o 1! |
| `easypxf` / `XFStyle` | `Font`, `PatternFill`, `NamedStyle` | inny model stylu |
| `wb.save("plik.xls")` | `wb.save("plik.xlsx")` | formaty **nie są** zamienne |

**Z `xlsxwriter` (zapis `.xlsx`, brak odczytu):**

| `xlsxwriter` | `openpyxl` | Uwaga |
|---|---|---|
| `xlsxwriter.Workbook("f.xlsx")` | `Workbook()` + `wb.save("f.xlsx")` | inne cykl życia |
| `wb.add_worksheet("Dane")` | `wb.create_sheet("Dane")` | |
| `ws.write(r, c, val)` | `ws.cell(row=r+1, column=c+1, value=val)` | przesunięcie o 1! |
| `ws.write_formula("A1", "=...")` | `ws["A1"] = "=..."` | |
| `wb.add_format({"num_format": "..."})` | `NamedStyle(number_format=...)` / `cell.number_format` | |
| `ws.conditional_format(range, {...})` | `ws.conditional_formatting.add(range, Rule)` | |
| `ws.add_chart(chart, "F2")` | `ws.add_chart(chart, "F2")` | API **podobne**, ale nie identyczne |
| `ws.add_sparkline(...)` | **brak odpowiednika** | kluczowa różnica funkcjonalna |
| `ws.data_validation(...)` | `DataValidation` + `add_data_validation` | |
| `ws.autofilter(range)` | `ws.auto_filter.ref = range` | |
| `wb.close()` | `wb.save(...)` | „close" w xlsxwriter **zapisuje** |
| — | **odczyt istniejących plików** | xlsxwriter **nie potrafi**; to jego główne ograniczenie |

**Z `openpyxl 2.x` na 3.1.x (usunięte i zmienione):**

| 2.x | 3.1.x | Uwaga |
|---|---|---|
| `wb.get_sheet_names()` | `wb.sheetnames` | usunięte w 3.0 |
| `wb.get_sheet_by_name("X")` | `wb["X"]` | usunięte w 3.0 |
| `wb.add_named_range(...)`, `get_named_range(...)`, `remove_named_range(...)` | `wb.defined_names` + `DefinedName` | usunięte w 3.1 |
| `openpyxl.writer.excel.save_virtual_workbook(wb)` | `wb.save(BytesIO())` + `getvalue()` | usunięte w 3.1 |
| `openpyxl.pandas` | `openpyxl.utils.dataframe.dataframe_to_rows` | usunięte w 3.1 |
| `wb.guess_types = True` | `cell.data_type` / jawne typy | usunięte w 3.0 |
| `ws.get_squared_range(...)` | `ws.iter_rows(...)` | usunięte w 3.0 |
| `Table(name="X", ...)` | `Table(displayName="X", ...)` | zmiana parametru w 3.0 |
| `Style(...)` (globalny) | `NamedStyle(...)` | dawno usunięte |
| `ws.cell(...).value = ...` po `read_only` | niedostępne | było i jest niedostępne |
| `Worksheet.get_highest_row/column()` | `max_row` / `max_column` | usunięte dawno |

## 5. Ćwiczenia

### 🟢 Rozgrzewka — sześć krótkich zadań (30–60 minut całości)

**1. Instalacja i kontekst.** Utwórz środowisko, zainstaluj `openpyxl==3.1.5`, `pillow`, `lxml`, wypisz wersję i sprawdź obecność opcjonalnych zależności kodem z `§3.12`.

**2. Twoja pierwsza ściągawka.** Utwórz plik `openpyxl_helpers.py` w katalogu `output/` i wklej do niego cztery narzędzia z `§3.13` (`inwentarz`, `ryzyka`, `podsumowanie_skoroszytu`, `porownaj_skoroszyty`) plus przepis 1 z `§3.14` (`bezpieczny_zapis`). To jest **Twój** plik — będzie Ci służył dłużej niż ten kurs.

**3. Inwentarz trzech plików.** Weź trzy `.xlsx` z pracy: jeden wygenerowany przez system, jeden tworzony ręcznie, jeden przysłany przez klienta. Dla każdego wypisz `inwentarz()` i `ryzyka()`. Zapisz wynik w `output/28_inwentarz.md`.

**4. Osiem asercji na osiem pułapek.** Napisz plik z testami, które **wymuszają** zachowanie z `§6`:

```python
def test_append_ze_stringiem_wpisuje_trzy_komorki():
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.append("ABC")
    assert ws["A1"].value == "A" and ws["C1"].value == "C"   # to NIE jest oczywiste!

def test_mutacja_stylu_nie_trafia_do_komorki():
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook(); ws = wb.active
    ws["A1"] = "x"
    ws["A1"].font.bold = True
    assert ws["A1"].font.bold in (None, False)   # sprawdz, jak jest w Twojej wersji

def test_formula_rule_bez_znaku_rowna():
    from openpyxl.formatting.rule import FormulaRule
    regula = FormulaRule(formula=['$D2="Pilne"'], fill=None)
    assert not regula.formula[0].startswith("=")
```

Dopisz pięć kolejnych: `fgColor` vs `bgColor`, `range_boundaries` vs `coordinate_to_tuple`, limit 32 767 znaków, `max_row` po formatowaniu, `data_only=True` + `save()`.

**5. Przejrzyj dokumentację.** Otwórz dokumentację openpyxl 3.1.x i przejdź **spis treści**. Zaznacz rozdziały, których nie znasz. To jest Twoja lista dalszej nauki — konkretna, nie „przeczytam kiedyś całość".

**6. Zmierz koszt stylów.** Wygeneruj trzy pliki po 20 000 komórek: (a) bez stylów, (b) styl nazwany, (c) nowy `Font(bold=True)` w każdym wywołaniu. Porównaj czas i rozmiar `xl/styles.xml`. To jest doświadczenie, którego nie da się zastąpić opisem.

### 🟡 Warsztat — trzy zadania diagnostyczne (2–4 godziny każde)

**Zadanie A — diagnoza z opisem objawów.**

Poniżej pięć objawów. Dla każdego przeprowadź **kolejność z `§2.3`** i zapisz w `output/28_diagnozy.md`: który krok dał odpowiedź, jaka jest przyczyna, gdzie jest **właściwe** miejsce naprawy (domena? aplikacja? adapter?) i jak to zabezpieczyć na przyszłość.

1. „Raport otwiera się, ale wykres jest pusty."
2. „Wszystkie komórki z kwotami po prawej są **wyrównane do lewej**."
3. „Podsumowanie pokazuje 250 000 zł, wiersze szczegółowe dają 249 998 zł." (różnica 2 zł)
4. „Po zapisie pliku klient zgłasza, że zniknęły mu dwie zakładki." (w pliku są 4 zamiast 6)
5. „Reguła warunkowa widoczna w `_meta` jako dodana, ale na arkuszu nic się nie podświetla."

**Zadanie B — porównywarka skoroszytów.**

Rozbuduj `porownaj_skoroszyty()` tak, żeby oprócz różnic w częściach XML pokazywała **różnice semantyczne**: które arkusze mają inną liczbę wierszy, które komórki mają różne wartości, ile reguł warunkowych ma każdy arkusz, czy tabele mają inne `ref`. Napisz na to test: wygeneruj dwa pliki różniące się **jedną** komórką i sprawdź, że funkcja wskaże **dokładnie jedną** różnicę.

**Zadanie C — sanitizer z testami.**

Napisz `domain/reguly.py::przyjmij_tekst_zewnetrzny()` (wzorzec z modułu 27) i pokryj ją testami parametryzowanymi:

```python
import pytest

@pytest.mark.parametrize("wartosc", ["=1+1", "+1", "-1", "@x", "\tx", "\rx",
                                     "=HYPERLINK(\"http://x\",\"klik\")"])
def test_odrzuca(wartosc):
    with pytest.raises(BladDomeny):
        przyjmij_tekst_zewnetrzny(wartosc, pole="klient")

@pytest.mark.parametrize("wartosc", ["Kowalski", "Kowalski-Nowak", "35-001", "",
                                     "Zażółć gęślą jaźń"])
def test_przyjmuje(wartosc):
    assert przyjmij_tekst_zewnetrzny(wartosc, pole="klient") == wartosc.strip()
```

**Dlaczego akurat te dwa przypadki.** `"Kowalski-Nowak"` **musi** przejść (łącznik nie na początku), a `"-1"` **musi** zostać odrzucone. To jest różnica między „filtruję znak wszędzie" (psuje dane) a „filtruję znak na początku" (chroni). Napisz też test na puste wejście i na tekst 40 000 znaków.

### 🔴 Wyzwanie — narzędzie `xlsxcheck` z pełną diagnostyką

**Cel:** jedno narzędzie CLI, jedno wywołanie, komplet informacji o pliku. To zwieńczenie kursu: łączy `zipfile`, model openpyxl w trybie `read_only`, sanitację, porównanie i raportowanie.

**Wymagania:**

```bash
python -m xlsxcheck inwentarz raport.xlsx
python -m xlsxcheck struktura raport.xlsx
python -m xlsxcheck ryzyka raport.xlsx
python -m xlsxcheck porownaj a.xlsx b.xlsx
python -m xlsxcheck audyt raport.xlsx            # komplet przed wysyłką
python -m xlsxcheck audyt raport.xlsx --json     # maszynowo, do CI
```

- Wyjście czytelne dla człowieka **i** `--json` dla maszyny (kod wyjścia ≠ 0, gdy wykryto ryzyko).
- **Skanowanie treści** arkuszy, nie tylko nazw części (sparkline'y!).
- Wykrywanie: ukrytych arkuszy, komentarzy, linków zewnętrznych, makr, `xl/media/`, formuł na ukrytych arkuszach, zewnętrznych odwołań w formułach.
- **Brak** modyfikowania plików wejściowych. Tylko odczyt.
- Bez `openpyxl` tam, gdzie wystarczy `zipfile` (inwentarz, ryzyka) — narzędzie ma działać także wtedy, gdy `openpyxl` nie jest zainstalowany w danym etapie CI.

<details>
<summary><strong>Szkic rozwiązania — struktura, decyzje i pełny kod `xlsxcheck`</strong></summary>

### Struktura

```text
xlsxcheck/
├── __main__.py          <- CLI: argparse, podkomendy, kody wyjscia
├── inventory.py         <- zipfile-only: czesci, rozmiary, kategorie ryzyka
├── structure.py         <- openpyxl read_only: arkusze, wymiary, typy, reguly
├── security.py          <- audyt przed wysylka (ukryte, linki, komentarze)
├── compare.py           <- porownanie po czesciach XML
└── report.py            <- formatowanie tekst/JSON
```

### Kluczowa decyzja 1: dwa poziomy zależności

Podkomendy `inwentarz` i `ryzyka` **nie importują openpyxl** — działają na `zipfile`. Dzięki temu:

- działają w środowisku bez biblioteki (i to jest **testowalne** w CI — job bez openpyxl),
- są **szybkie** (nie parsują XML-a),
- mogą być uruchomione na pliku, który openpyxl **odrzuca** przy wczytaniu (np. uszkodzone relacje) — co jest dokładnie sytuacją, w której narzędzie diagnostyczne jest najbardziej potrzebne.

### Kluczowa decyzja 2: wykrywanie sparkline'ów przez treść, nie przez nazwę

Pokazane w `§3.13`. To jest ten szczegół, który odróżnia narzędzie **działające** od narzędzia **wyglądającego na działające**.

### Pełny kod: `xlsxcheck/__main__.py`

```python
"""xlsxcheck - narzedzie diagnostyczne do plikow .xlsx / .xlsm.

Filozofia:
  - TYLKO ODCZYT. Narzedzie nigdy nie modyfikuje pliku wejsciowego.
  - Dwa poziomy zaleznosci: inwentarz/ryzyka dzialaja bez openpyxl.
  - Kod wyjscia != 0, gdy wykryto ryzyko (uzyteczne w CI).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from xlsxcheck.compare import porownaj
from xlsxcheck.inventory import inwentarz, ryzyka
from xlsxcheck.report import formatuj_inwentarz, formatuj_ryzyka
from xlsxcheck.security import audyt
from xlsxcheck.structure import struktura


def _wypisz(dane: object, jako_json: bool) -> None:
    if jako_json:
        print(json.dumps(dane, ensure_ascii=False, indent=2, default=str))
    else:
        print(formatuj_inwentarz(dane) if isinstance(dane, dict) else dane)


def _inwentarz(args) -> int:
    dane = inwentarz(args.plik)
    _wypisz(dane, args.json)
    return 0


def _ryzyka(args) -> int:
    lista = ryzyka(args.plik)
    if args.json:
        print(json.dumps({"ryzyka": list(lista)}, ensure_ascii=False, indent=2))
    else:
        print(formatuj_ryzyka(lista))
    return 1 if lista else 0        # RYZYKO = kod wyjscia 1 (dla CI)


def _struktura(args) -> int:
    _wypisz(struktura(args.plik), args.json)
    return 0


def _porownaj(args) -> int:
    wynik = porownaj(args.a, args.b)
    _wypisz(wynik, args.json)
    return 0 if not wynik["rozne"] else 1


def _audyt(args) -> int:
    wynik = audyt(args.plik)
    _wypisz(wynik, args.json)
    return 1 if wynik["ryzyka"] else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="xlsxcheck",
        description="Diagnostyka plikow .xlsx: inwentarz, struktura, ryzyka, porownanie.",
    )
    parser.add_argument("--json", action="store_true",
                        help="wyjscie maszynowe (JSON)")
    pod = parser.add_subparsers(dest="komenda", required=True)

    for nazwa, funkcja, opis in (
        ("inwentarz", _inwentarz, "lista czesci pliku i rozmiary"),
        ("ryzyka", _ryzyka, "co moze zginac przy zapisie przez openpyxl"),
        ("struktura", _struktura, "arkusze, wymiary, formuly, regul warunkowe"),
        ("audyt", _audyt, "kompletny audyt przed wyslaniem pliku na zewnatrz"),
    ):
        p = pod.add_parser(nazwa, help=opis)
        p.add_argument("plik", type=Path)
        p.set_defaults(func=funkcja)

    p_por = pod.add_parser("porownaj", help="porownanie dwoch plikow")
    p_por.add_argument("a", type=Path)
    p_por.add_argument("b", type=Path)
    p_por.set_defaults(func=_porownaj)

    args = parser.parse_args(argv)

    if not getattr(args, "plik", None) or not args.plik.exists():
        pliki = [getattr(args, "plik", None), getattr(args, "a", None),
                 getattr(args, "b", None)]
        brakujacy = next((p for p in pliki if p is not None and not p.exists()), None)
        if brakujacy is not None:
            print(f"BLAD: plik nie istnieje: {brakujacy}", file=sys.stderr)
            return 2
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
```

### Pełny kod: `xlsxcheck/inventory.py`

```python
"""Skanowanie archiwum - BEZ openpyxl. Dziala zawsze i szybko."""

from __future__ import annotations

import zipfile
from pathlib import Path

# Kategoria L: czesci, ktorych openpyxl nie modeluje (wykrywane po nazwie).
UTRATA = {
    "xl/slicers/":     "slicery",
    "xl/timelines/":   "timeline'y",
    "xl/queryTables/": "zapytania Power Query",
    "xl/model/":       "model danych (Power Pivot)",
    "xl/ctrlProps/":   "kontrolki formularza",
    "xl/activeX/":     "kontrolki ActiveX",
    "customXml/":      "wlasne czesci XML",
    "xl/threadedComments/": "komentarze nowego typu",
}
# Kategoria C: czesci przetwarzane, ale z ograniczeniami.
CZESCIOWO = {
    "xl/charts/":      "wykresy (odbudowywane z modelu)",
    "xl/drawings/":    "rysunki (ksztalty moga zginac)",
    "xl/pivotTables/": "tabele przestawne (odczyt definicji)",
    "xl/pivotCache/":  "cache tabel przestawnych",
}


def inwentarz(sciezka: str | Path) -> dict:
    """Czesci pliku, ich rozpakowane rozmiary i wspolczynnik sprezania."""
    sciezka = Path(sciezka)
    if not zipfile.is_zipfile(sciezka):
        return {"plik": str(sciezka), "blad": "to nie jest poprawny plik ZIP/OOXML"}

    czesci: dict[str, int] = {}
    with zipfile.ZipFile(sciezka) as archiwum:
        for info in archiwum.infolist():
            czesci[info.filename] = info.file_size

    rozpakowane = sum(czesci.values())
    bajtow = sciezka.stat().st_size
    return {
        "plik": str(sciezka),
        "bajtow": bajtow,
        "rozpakowane": rozpakowane,
        "wspolczynnik": round(rozpakowane / bajtow, 1) if bajtow else 0,
        "czesci": len(czesci),
        "spis": dict(sorted(czesci.items(), key=lambda kv: -kv[1])),
    }


def _tresc_arkuszy(archiwum: zipfile.ZipFile) -> str:
    """Tresc XML wszystkich arkuszy - potrzebna do detekcji sparkline'ow."""
    return "".join(
        archiwum.read(n).decode("utf-8", "ignore")
        for n in archiwum.namelist()
        if n.startswith("xl/worksheets/") and n.endswith(".xml")
    )


def ryzyka(sciezka: str | Path) -> tuple[str, ...]:
    """Co moze zginac albo zostac uproszczone przy zapisie przez openpyxl."""
    sciezka = Path(sciezka)
    if not zipfile.is_zipfile(sciezka):
        return ("PLIK: to nie jest poprawny plik ZIP/OOXML",)

    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()
        tresc = _tresc_arkuszy(archiwum)

    ostrzezenia: list[str] = []
    for prefiks, opis in UTRATA.items():
        if any(n.startswith(prefiks) for n in nazwy):
            ostrzezenia.append(f"UTRATA: {opis}")

    # Sparkline'y NIE maja wlasnej czesci - siedza w XML arkusza (x14).
    if "sparklineGroups" in tresc:
        ostrzezenia.append("UTRATA: sparkline'y (brak API w openpyxl)")

    for prefiks, opis in CZESCIOWO.items():
        if any(n.startswith(prefiks) for n in nazwy):
            ostrzezenia.append(f"CZESCIOWO: {opis}")

    if "xl/vbaProject.bin" in nazwy:
        ostrzezenia.append("MAKRA: obecne (wczytuj z keep_vba=True)")

    # Zip bomb - orientacyjny wskaznik.
    rozpakowane = sum(archiwum.getinfo(n).file_size for n in nazwy)
    bajtow = sciezka.stat().st_size
    if bajtow and rozpakowane / bajtow > 100:
        ostrzezenia.append(
            f"RYZYKO: wysokie sprezenie ({rozpakowane / bajtow:.0f}x) - "
            f"sprawdz limit rozpakowania przed przetwarzaniem"
        )

    return tuple(ostrzezenia)
```

### Pełny kod: `xlsxcheck/security.py`

```python
"""Audyt przed wyslaniem pliku na zewnatrz (modul 21 w praktyce)."""

from __future__ import annotations

import zipfile
from pathlib import Path

from openpyxl import load_workbook


def audyt(sciezka: str | Path) -> dict:
    """Wykrywa to, czego nie widac na pierwszy rzut oka.

    Sprawdza: ukryte arkusze, komentarze, hiperlinki, odwolania zewnetrzne,
    makra, media, wlasciwosci dokumentu (metadane autora i sciezki).
    """
    from xlsxcheck.inventory import ryzyka

    sciezka = Path(sciezka)
    wynik: dict = {
        "plik": str(sciezka),
        "ryzyka": list(ryzyka(sciezka)),
        "ukryte_arkusze": [],
        "komentarze": 0,
        "hiperlinki": [],
        "odwolania_zewnetrzne": [],
        "media": [],
        "metadane": {},
    }

    wb = load_workbook(sciezka, read_only=True)
    try:
        for nazwa in wb.sheetnames:
            ws = wb[nazwa]
            if ws.sheet_state != "visible":
                wynik["ukryte_arkusze"].append(f"{nazwa} ({ws.sheet_state})")
            for rzad in ws.iter_rows():
                for c in rzad:
                    if c.hyperlink is not None:
                        cel = getattr(c.hyperlink, "target", None) or c.hyperlink.location
                        wynik["hiperlinki"].append(f"{nazwa}!{c.coordinate} -> {cel}")
                    if c.comment is not None:
                        wynik["komentarze"] += 1
                    if (c.data_type == "f" and isinstance(c.value, str)
                            and ("[" in c.value and "]" in c.value)):
                        # Odwolanie do innego skoroszytu w formule.
                        wynik["odwolania_zewnetrzne"].append(
                            f"{nazwa}!{c.coordinate} -> {c.value[:80]}"
                        )
        props = wb.properties
        wynik["metadane"] = {
            "title": props.title,
            "creator": props.creator,
            "lastModifiedBy": props.lastModifiedBy,
            "description": props.description,
            "created": props.created,
        }
    finally:
        wb.close()      # OBOWIAZKOWE w read_only

    with zipfile.ZipFile(sciezka) as archiwum:
        wynik["media"] = [n for n in archiwum.namelist() if n.startswith("xl/media/")]

    # Dodatkowe pozycje do listy ryzyk na podstawie audytu.
    if wynik["ukryte_arkusze"]:
        wynik["ryzyka"].append(
            f"UKRYTE ARKUSZE: {', '.join(wynik['ukryte_arkusze'])} - "
            f"sprawdz, czy na pewno maja wyjsc na zewnatrz"
        )
    if wynik["odwolania_zewnetrzne"]:
        wynik["ryzyka"].append(
            f"ODWOLANIA ZEWNETRZNE: {len(wynik['odwolania_zewnetrzne'])} "
            f"- moga ujawnic sciezki lub nazwy serwerow"
        )
    if wynik["komentarze"]:
        wynik["ryzyka"].append(
            f"KOMENTARZE: {wynik['komentarze']} - moga zawierac notatki "
            f"nieprzeznaczone dla odbiorcy"
        )
    return wynik
```

### Pełny kod: `xlsxcheck/compare.py`

```python
"""Porownanie dwoch skoroszytow - po CZESCIACH, nie bajt w bajt."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

ISTOTNE = ("xl/worksheets/", "xl/styles.xml", "xl/tables/",
           "xl/workbook.xml", "xl/sharedStrings.xml", "xl/_rels/")


def _hasze(sciezka: Path) -> dict[str, str]:
    """Hasze tresci czesci XML po normalizacji bialych znakow.

    Normalizacja jest konieczna, bo openpyxl i Excel inaczej formatuja XML
    (wciecia, kolejnosc atrybutow) - bez niej porownanie zawsze pokaze roznice.
    """
    wynik: dict[str, str] = {}
    with zipfile.ZipFile(sciezka) as archiwum:
        for nazwa in archiwum.namelist():
            if nazwa.endswith(".xml") and nazwa.startswith(ISTOTNE):
                tresc = archiwum.read(nazwa).decode("utf-8", "ignore")
                wynik[nazwa] = hashlib.sha256(
                    "".join(tresc.split()).encode("utf-8")
                ).hexdigest()
    return wynik


def porownaj(a: str | Path, b: str | Path) -> dict:
    """Zwraca czesci tylko w A, tylko w B, rozne i identyczne."""
    hasze_a = _hasze(Path(a))
    hasze_b = _hasze(Path(b))
    wspolne = set(hasze_a) & set(hasze_b)
    return {
        "tylko_w_a": sorted(set(hasze_a) - set(hasze_b)),
        "tylko_w_b": sorted(set(hasze_b) - set(hasze_a)),
        "rozne": sorted(n for n in wspolne if hasze_a[n] != hasze_b[n]),
        "identyczne": sorted(n for n in wspolne if hasze_a[n] == hasze_b[n]),
        "uwaga": (
            "Porownanie po czesciach XML. Hasz calego pliku .xlsx NIE jest "
            "deterministyczny (ZIP zapisuje znaczniki czasu) - nie uzywaj go "
            "do testow regresji."
        ),
    }
```

### Pełny kod: `xlsxcheck/report.py`

```python
"""Formatowanie wynikow dla czlowieka."""

from __future__ import annotations


def _kb(bajtow: int) -> str:
    return f"{bajtow / 1024:.0f} KB"


def formatuj_inwentarz(dane: dict) -> str:
    if "blad" in dane:
        return f"[!] {dane['blad']}: {dane['plik']}"

    linie = [
        f"PLIK      {dane['plik']}",
        f"ROZMIAR   {_kb(dane['bajtow'])} "
        f"(po rozpakowaniu {_kb(dane['rozpakowane'])}, "
        f"{dane['wspolczynnik']}x)",
        f"CZESCI    {dane['czesci']}",
        "",
        "NAJWIEKSZE CZESCI:",
    ]
    for nazwa, rozmiar in list(dane["spis"].items())[:12]:
        linie.append(f"  {_kb(rozmiar):>10}  {nazwa}")
    return "\n".join(linie)


def formatuj_ryzyka(lista: tuple[str, ...]) -> str:
    if not lista:
        return "[OK] Nie wykryto czesci zagrozonych utrata danych.\n" \
               "     Plik wyglada na technicznie czysty dla openpyxl.\n" \
               "     To NIE zwalnia z weryfikacji na oko po zapisie."
    linie = ["[!] WYKRYTO RYZYKA - przeczytaj przed zapisem przez openpyxl:", ""]
    linie += [f"  * {pozycja}" for pozycja in lista]
    linie += [
        "",
        "WNIOSEK:",
        "  Jezeli plik byl tworzony recznie w Excelu i widzisz UTRATA -",
        "  NIE przepuszczaj go przez openpyxl w trybie 'wczytaj i zapisz'.",
        "  Wygeneruj nowy plik albo oddaj modyfikacje Excelowi.",
    ]
    return "\n".join(linie)
```

### Test — to najważniejsza część tego wyzwania

```python
"""tests/test_xlsxcheck.py"""

import pytest
from openpyxl import Workbook

from xlsxcheck.compare import porownaj
from xlsxcheck.inventory import inwentarz, ryzyka


def _zapisz(tmp_path, nazwa, wypelnij):
    wb = Workbook()
    ws = wb.active
    wypelnij(ws)
    cel = tmp_path / nazwa
    wb.save(cel)
    return cel


def test_inwentarz_pustego_skoroszytu(tmp_path):
    sciezka = _zapisz(tmp_path, "a.xlsx", lambda ws: ws.__setitem__("A1", "x"))
    dane = inwentarz(sciezka)
    assert "xl/workbook.xml" in dane["spis"]
    assert dane["rozpakowane"] > 0


def test_ryzyka_czystego_pliku_puste(tmp_path):
    sciezka = _zapisz(tmp_path, "b.xlsx", lambda ws: ws.__setitem__("A1", 1))
    assert ryzyka(sciezka) == ()


def test_ryzyka_wykrywaja_wykres(tmp_path):
    from openpyxl.chart import BarChart, Reference

    def dodaj(ws):
        ws.append(["Rok", "Wartosc"])
        ws.append([2026, 100]); ws.append([2027, 120])
        w = BarChart()
        w.add_data(Reference(ws, min_col=2, min_row=1, max_row=3),
                   titles_from_data=True)
        ws.add_chart(w, "E2")

    sciezka = _zapisz(tmp_path, "c.xlsx", dodaj)
    ostrzezenia = ryzyka(sciezka)
    assert any("wykresy" in o for o in ostrzezenia)


def test_porownaj_znajduje_jedna_rozna_czesc(tmp_path):
    a = _zapisz(tmp_path, "a.xlsx", lambda ws: ws.__setitem__("A1", 1))
    b = _zapisz(tmp_path, "b.xlsx", lambda ws: ws.__setitem__("A1", 2))
    wynik = porownaj(a, b)
    assert wynik["rozne"] == ["xl/worksheets/sheet1.xml"]
    assert wynik["tylko_w_a"] == [] and wynik["tylko_w_b"] == []


def test_porownaj_identycznych_ma_zero_roznic(tmp_path):
    a = _zapisz(tmp_path, "a.xlsx", lambda ws: ws.__setitem__("A1", 1))
    b = _zapisz(tmp_path, "b.xlsx", lambda ws: ws.__setitem__("A1", 1))
    wynik = porownaj(a, b)
    # UWAGA: porownanie dziala mimo roznych znacznikow czasu w ZIP!
    assert wynik["rozne"] == []


def test_porownanie_bajt_w_bajt_zawodzi(tmp_path):
    """Dowod, dlaczego NIE porownujemy plikow .xlsx bajt w bajt."""
    a = _zapisz(tmp_path, "a.xlsx", lambda ws: ws.__setitem__("A1", 1))
    b = _zapisz(tmp_path, "b.xlsx", lambda ws: ws.__setitem__("A1", 1))
    assert a.read_bytes() != b.read_bytes()      # rozne bajty...
    assert porownaj(a, b)["rozne"] == []         # ...ta sama tresc
```

**Ten ostatni test jest esencją tego wyzwania.** Dwa pliki o **identycznej** zawartości mają **różne bajty**. Kto porównuje skoroszyty bajt w bajt, ten ma test regresji, który jest czerwony zawsze — i dlatego po tygodniu zostaje wyłączony. Kto porównuje **części po normalizacji**, ten ma test regresji, który naprawdę wykrywa zmiany.

### Test bez openpyxl (do joba CI bez biblioteki)

```python
"""tests/test_inventory_bez_openpyxl.py - ten plik NIE importuje openpyxl."""

import sys
import zipfile
import pytest

from xlsxcheck.inventory import inwentarz, ryzyka


def test_inwentarz_nie_wymaga_openpyxl(tmp_path):
    assert "openpyxl" not in sys.modules, (
        "Ten test ma dzialac bez openpyxl - sprawdz importy w xlsxcheck.inventory"
    )
    # Minimalny poprawny archiwum OOXML
    cel = tmp_path / "minimal.xlsx"
    with zipfile.ZipFile(cel, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", "<workbook/>")
    assert "xl/workbook.xml" in inwentarz(cel)["spis"]
```

</details>

## 6. Typowe błędy i pułapki

### 6.1. Top 20 błędów początkującego (objaw → przyczyna → naprawa)

**1. `cell.font.bold = True` nie działa.**
→ *Przyczyna:* obiekt stylu jest **współdzielony**, a komórka trzyma tylko **indeks** do wspólnej tabeli stylów. Mutujesz pieczątkę w szufladzie, a nie odcisk na papierze.
→ *Naprawa:*
```python
from copy import copy
nowy = copy(cell.font)
nowy.bold = True
cell.font = nowy
```
Albo, lepiej dla powtarzalnych stylów: `NamedStyle` + `cell.style = "RaportNaglowek"`.

**2. `ws.append("ABC")` wpisuje trzy komórki zamiast jednej.**
→ *Przyczyna:* `append` przyjmuje **iterowalne** i rozkłada je na komórki. String jest iterowalny.
→ *Naprawa:* `ws.append(["ABC"])` — lista jednoelementowa. Zasada: **`append` zawsze dostaje listę/krotkę**, nigdy goły string.

**3. Nadpisanie pliku źródłowego.**
→ *Przyczyna:* `wb.save(ta_sama_sciezka)` — zapis nadpisuje **bez ostrzeżenia**, a Ty masz w pamięci model wczytany z tego pliku, więc „zapisujesz to, co było" **plus** zmiany. Jeżeli cokolwiek zostało zgubione przy wczytywaniu, właśnie to utrwaliłeś.
→ *Naprawa:* `wb.save("raport_v2.xlsx")` albo `bezpieczny_zapis()` z backupem. Zasada z modułu 00: **nigdy nie nadpisujemy pliku źródłowego.** Jeżeli klient nalega na tę samą nazwę — użyj `os.replace` z tmp i **backupu**, ale wiedz, że nadal nadpisujesz.

**4. `data_only=True` + `wb.save()`.**
→ *Przyczyna:* `data_only=True` wczytuje wartości zamiast formuł. Zapis takiego modelu **zapisuje wartości**. Formuły są nieodwracalnie zamienione na liczby.
→ *Naprawa:* nigdy nie łącz tych dwóch rzeczy. Do odczytu danych: `data_only=True`, ale **tylko czytasz**. Do modyfikacji: `data_only=False` (domyślnie). Jeżeli potrzebujesz obu — wczytaj plik **dwa razy**.

**5. Brak `keep_vba` przy pliku `.xlsm`.**
→ *Przyczyna:* bez `keep_vba=True` część `xl/vbaProject.bin` nie trafia do pliku wynikowego. Excel otwiera plik, ale makra znikają — a jeżeli plik ma rozszerzenie `.xlsm`, Excel może zgłosić problemy.
→ *Naprawa:* `load_workbook("plik.xlsm", keep_vba=True)` i zapis **pod tym samym rozszerzeniem** `.xlsm`. Makr nie edytujesz — tylko zachowujesz.

**6. `max_row` zwraca 1 048 576 (albo 4 000 przy 200 wierszach danych).**
→ *Przyczyna:* `max_row` **rośnie od samego formatowania**, nie tylko od wartości. Kliknięcie „pogrub" w kolumnie do końca arkusza ustawia wymiar. Po `delete_rows` wymiar może zostać zapisany w metadanych.
→ *Naprawa:* skanuj **w górę** i sprawdzaj `value is not None` (kod w `§3.16`, pytanie 18) albo `ws.reset_dimensions()` / `ws.calculate_dimension(force=True)`, jeżeli kontrolujesz cały cykl życia pliku.

**7. Wykres ma osie z podpisami `1, 2, 3` zamiast nazw regionów.**
→ *Przyczyna:* brak `set_categories()`. Bez kategorii Excel używa numerów punktów.
→ *Naprawa:* `chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ostatni))`. Pamiętaj: `min_row=2` (bez nagłówka), `min_col=1` (kolumna z nazwami).

**8. Scalenie komórek zabrało dane.**
→ *Przyczyna:* scalanie zachowuje **tylko** wartość górnej-lewej komórki. Reszta jest odrzucana przez Excela (a openpyxl zapisze zakres bez ich wartości).
→ *Naprawa:* **najpierw przepisz** wartości do górnej-lewej, **potem** `ws.merge_cells(...)`. Przy odczycie: `ws.merged_cells.ranges` i pamiętaj, że wartość jest **tylko** w górnej-lewej.

**9. Wypełnienie komórki nie działa mimo ustawionego koloru.**
→ *Przyczyna:* dla `fill_type="solid"` **widoczny jest `fgColor`**, a nie `bgColor`. Intuicja mówi odwrotnie („tło = background").
→ *Naprawa:* `PatternFill(fill_type="solid", fgColor="FFFFC000")`. `bgColor` ma znaczenie **tylko dla wypełnień wzorkiem**.

**10. Reguła warunkowa z `FormulaRule` nie działa.**
→ *Przyczyna:* formuła w regule **nie zaczyna się od `=`** — a Ty napisałeś `"=A1>100"`. Albo użyłeś adresu względnego tam, gdzie potrzebny jest absolutny (`A1` zamiast `$A1`).
→ *Naprawa:* `FormulaRule(formula=["$D2=\"Pilne\""], ...)` — bez `=`, a adresy **względem lewej-górnej komórki zakresu**. Reguła to **wzór na kalkę**: piszesz go raz dla pierwszego pola, Excel przesuwa na resztę.

**11. `fitToWidth` nie działa, wydruk ucina kolumny.**
→ *Przyczyna:* samo `fitToWidth`/`fitToHeight` jest ignorowane bez włączenia trybu dopasowania.
→ *Naprawa:* trzy linie razem:
```python
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
```

**12. Wyciek pamięci przy `read_only`.**
→ *Przyczyna:* brak `wb.close()`. W trybie `read_only` openpyxl trzyma otwarty uchwyt do archiwum i strumieni.
→ *Naprawa:* `try: ... finally: wb.close()`. Przy wielu plikach w pętli to jest **obowiązkowe**, nie „zalecane" — bez tego proces rośnie w pamięci z każdym plikiem.

**13. `ws["A1"]` w trybie `read_only` zgłasza wyjątek.**
→ *Przyczyna:* `ReadOnlyWorksheet` nie ma dostępu losowego — czyta **tylko w przód**, wiersz po wierszu.
→ *Naprawa:* `for rzad in ws.iter_rows(values_only=True): ...` i zbieraj potrzebne kolumny licznikami. Jeżeli naprawdę potrzebujesz dostępu do konkretnej komórki — wczytaj plik w trybie normalnym (i zapłać pamięcią) albo dwukrotnie przejdź plik w trybie `read_only` (raz dla nagłówków, raz dla danych).

**14. `ImportError: You must install Pillow to fetch image objects` przy `add_image`.**
→ *Przyczyna:* `openpyxl.drawing.image.Image` wymaga `pillow` — i błąd pojawia się **przy imporcie**, więc kodu nie da się nawet wczytać.
→ *Naprawa:* `pip install pillow`. Jeżeli obrazów używasz opcjonalnie — **importuj `Image` wewnątrz funkcji** i obsłuż `ImportError`, żeby brak `pillow` nie wywracał całego modułu.

**15. Logo jest rozciągnięte w pionie.**
→ *Przyczyna:* ustawiłeś **tylko jeden** wymiar (`img.height = 44`) i liczyłeś, że drugi „zachowa proporcje". Nie zachowa. Dodatkowo jednostki obrazu zależą od DPI pliku obrazu.
→ *Naprawa:* licz **obie** wartości z **jednej skali** (kod w `§3.10`).

**16. Komentarz zniknął — był w scalonej komórce.**
→ *Przyczyna:* openpyxl nie obsługuje komentarzy w komórkach należących do scalonych zakresów; przy wczytywaniu ostrzega, a przy zapisie komentarz nie trafia do pliku.
→ *Naprawa:* umieszczaj komentarze w komórkach **poza** scalonymi zakresami (np. w komórce obok). Sprawdź przed zapisem: `for rng in ws.merged_cells.ranges: assert not any(...)`.

**17. Lista rozwijana się nie zapisała (albo działa tylko w części komórek).**
→ *Przyczyna:* dwa warianty: (a) zabrakło `ws.add_data_validation(dv)` — obiekt utworzony, ale nie zarejestrowany; (b) `formula1` przekroczył **255 znaków** (lista literałowa).
→ *Naprawa:* zawsze `add_data_validation(dv)` **przed** `dv.add(zakres)`; dla długich list — zakres z innego arkusza: `formula1="=Regiony!$A$2:$A$20"`.

**18. Brak `lxml` i wydajność „nie zgadza się z tutorialem".**
→ *Przyczyna:* bez `lxml` openpyxl używa czystego `xml.etree`, co przy dużych plikach jest **kilkukrotnie** wolniejsze. Tutoriale zakładają `lxml`, bo dla wielu osób jest już zainstalowany jako zależność czegoś innego.
→ *Naprawa:* `pip install lxml` i **dodaj do zależności produkcyjnych** — jeżeli w wymaganiach masz „200 000 wierszy w 60 sekund", to `lxml` jest częścią tego wymagania, a nie optymalizacją. Sprawdź kodem z `§3.12`.

**19. Data wyświetla się jako liczba (43831), albo miesiąc ma angielską nazwę.**
→ *Przyczyna:* dwa różne problemy. (a) **Brak formatu daty** — komórka trzyma serial, więc Excel pokazuje liczbę. (b) **Format z nazwą miesiąca** (`mmmm`) renderuje się w **języku ustawień systemowych**, nie w języku Twojego kodu; format z `[$-xxx]` wskazuje locale, ale i tak decyduje system odbiorcy.
→ *Naprawa:* (a) `cell.number_format = "yyyy-mm-dd"` **lub** styl nazwany z formatem daty. (b) Używaj formatów **numerycznych** (`yyyy-mm-dd`, `dd.mm.yyyy`) w raportach międzynarodowych; nazwy miesięcy zostaw, gdy wiesz, że odbiorca ma polskie ustawienia. **Nigdy nie licz na to, że `mmmm` da „październik" na każdym komputerze.**

**20. `PermissionError` (Windows) albo „użytkownik widzi starą treść" (Linux) przy zapisie.**
→ *Przyczyna:* Windows blokuje podmianę otwartego pliku; Linux podmienia „pod spodem", ale otwarte okno Excela dalej pokazuje starą treść z pamięci. Dwie różne platformy, **jeden problem**: plik jest otwarty.
→ *Naprawa:* jawny komunikat dla użytkownika (wzorzec w `§3.14`, przepis 1) + zapis pod nową nazwą jako opcja + retry z ograniczoną liczbą prób. **Nie** rozwiązuj tego `try/except: pass` — użytkownik nie może się dowiedzieć o awarii dopiero po tygodniu.

### 6.2. Antywzorce projektowe

**1. „Bogata fasada", przez którą i tak trzeba znać openpyxl.**
*Objaw:* klasa `RaportExcel` ma metody `dodaj_arkusz(ws)`, `sformatuj(cell, font)`, `dodaj_wykres(chart, anchor)`. Publiczne API przyjmuje typy openpyxl.
*Dlaczego to antywzorzec:* to nie jest fasada, to **przepakowanie**. Nie kupiłeś nic — nadal musisz znać openpyxl, a dodatkowo musisz znać **drugie** API. Test sprawdzający „brak przecieku abstrakcji" nie przechodzi, bo fasada **jest** przeciekiem.
*Napraw:* sygnatury w języku domeny: `dodaj_arkusz_regionow(wiersze: list[AgregatRegionu])`, `sformatuj_kwoty(kolumna: str)`. **Zero `Workbook`, `Cell`, `Font`** w publicznym API (moduł 26).

**2. Cały program w 400-linijkowej funkcji `main()`.**
*Objaw:* `main()` robi: parsowanie argumentów, wczytanie CSV, czyszczenie danych, obliczenia, tworzenie arkuszy, formatowanie, wykresy, zapis, logowanie. Wszystko liniowo, zmienne `ws1`, `ws2`, `i`, `j`, `dane2`.
*Dlaczego to antywzorzec:* nie da się **przetestować** żadnej części osobno, a więc w praktyce nie testuje się nic. Każda zmiana wymaga uruchomienia całości. Debugowanie wymaga czytania 400 linii, bo nie ma granic.
*Napraw:* podział na warstwy i funkcje o jednej odpowiedzialności (moduły 22–26): `wczytaj_dane()`, `zbuduj_model()`, `eksportuj()`, `main()` tylko komponuje. Test na każdą funkcję osobno, bez plików.

**3. Style wpisane w czterdziestu miejscach.**
*Objaw:* `Font(bold=True, color="FFC000")` w 40 miejscach, każdy z innym odcieniem żółtego. Zmiana „firmowego złotego" to 40 zmian i tydzień pracy.
*Dlaczego to antywzorzec:* kolor to **decyzja o prezentacji**, a jest rozsiany po kodzie **logiki**. Żaden test tego nie złapie, a pierwszy raport po zmianie marki wygląda niejednolicie.
*Napraw:* jeden słownik `KOLORY` i jeden rejestr `FORMATY` (moduł 09, 23). Test: `grep -rn '"F[0-9A-F]\{5\}"' src/` musi zwracać **tylko** plik `styl.py`.

**4. Brak rejestru stylów — everything by hand.**
*Objaw:* każda komórka dostaje nowy obiekt stylu tworzony w pętli. Plik 5 MB, `xl/styles.xml` 2 MB, generowanie trwa cztery razy dłużej niż powinno.
*Dlaczego to antywzorzec:* marnujesz czas i pamięć na coś, co Excel i tak zdeduplikuje przy zapisie. Flyweight (moduł 23) jest tu **wprost** użyteczny.
*Napraw:* `NamedStyle` albo współdzielone obiekty tworzone **poza** pętlą. Mierz: `getinfo("xl/styles.xml").file_size` przed i po.

**5. Raport jako źródło prawdy.**
*Objaw:* ktoś poprawia liczbę bezpośrednio w pliku `.xlsx` i wysyła dalej. Wersja „prawdziwa" i wersja „wysłana" rozjeżdżają się. Za kwartał nikt nie wie, która jest aktualna.
*Dlaczego to antywzorzec:* raport jest **projekcją** danych, a nie ich magazynem. Każda ręczna poprawka w pliku wynikowym jest **nieodtwarzalna** — nie ma jej w systemie źródłowym, nie przejdzie testów, nie pojawi się w kolejnym generowaniu.
*Napraw:* dane w systemie źródłowym (baza, pliki wejściowe), raport generowany **zawsze od zera**. Jeżeli potrzebne są korekty — wprowadzaj je w **danych**, nie w pliku. Arkusz `_meta` z `hash_wsadu` pozwala wykryć, że plik został zmieniony ręcznie (moduł 26, 27).

**6. Brak testów („to tylko generowanie Excela").**
*Objaw:* jedyny test to „uruchomiłem i otworzyłem plik, wygląda dobrze".
*Dlaczego to antywzorzec:* brak testów oznacza, że **nie masz odwagi nic zmienić**. Każda modyfikacja jest ryzykiem, więc kod zamarza — a wymagania się zmieniają. Testy przez `BytesIO` (moduł 20) są **takie same** jak dla zwykłej logiki: nie potrzebują Excela ani dysku.
*Napraw:* testy domeny bez openpyxl (job CI), testy struktury pliku na `BytesIO`, test regresji na `podsumowanie_skoroszytu`.

**7. Kopiowanie szablonu przez `copy_worksheet` z nadzieją, że przeniesie obrazki.**
*Objaw:* szablon ma logo w nagłówku; skrypt kopiuje arkusz szablonu i generuje z niego raport. W wyniku logo jest **tylko** w pierwszym arkuszu (a jeśli szablon kopiujesz do innego pliku — w żadnym).
*Dlaczego to antywzorzec:* `copy_worksheet` kopiuje **komórki i część atrybutrów arkusza**, ale **nie obrazy ani wykresy**. Nie da się go użyć między skoroszytami. Oparcie szablonu na rzeczach, których nie da się skopiować, to projekt, który **musi** się zepsuć przy pierwszej zmianie.
*Napraw:* albo **`load_workbook(szablon)` i zapis do nowego pliku** (wtedy wszystko zostaje, bo nie kopiujemy — przepisujemy), albo generuj arkusz w kodzie i **wstawiaj logo programowo** (`add_image`), traktując szablon jako źródło **wzorców** (stylów, układów), a nie gotowy dokument.

**8. Logika biznesowa w widoku / CLI.**
*Objaw:* w `views.py` albo w `main()` widzisz `marza = (netto - koszt) / netto if netto else 0` i obliczenia progów dla formatowania warunkowego.
*Dlaczego to antywzorzec:* reguła biznesowa jest **nietestowalna** (trzeba odtworzyć cały request/CLI), jest w dwóch miejscach (bo raport i API liczą inaczej), a zmiana progu wymaga wdrożenia całej aplikacji.
*Napraw:* reguły w **domenie** (`domain/reguly.py`), a warstwy prezentacji tylko **delegują**. Test: `grep -rn "openpyxl" src/` nie zwraca nic z katalogu domeny i aplikacji (moduł 26, 27).

### 6.3. Checklista „przed wysłaniem pliku na zewnątrz"

Krótka lista, którą można przejść w trzy minuty. Odpowiednik `xlsxcheck audyt`, ale w wersji do zapamiętania:

- [ ] **1. Ukryte arkusze** — czy nie wychodzą z plikiem? (`ws.sheet_state != "visible"`)
- [ ] **2. Metadane** — `wb.properties.creator` / `lastModifiedBy` nie zawierają nazwiska i ścieżki sieciowej, których nie chcesz ujawniać?
- [ ] **3. Komentarze** — czy nie ma notatek roboczych i „do sprawdzenia"?
- [ ] **4. Odwołania zewnętrzne** — czy formuły nie wskazują lokalnych ścieżek albo nazw serwerów?
- [ ] **5. Media** — czy w `xl/media/` nie ma przypadkowych zrzutów ekranu ze spotkania?
- [ ] **6. Definiowane nazwy** — czy nie ma nazw typu `cena_kosztowa_2025` ujawniających strukturę wewnętrzną?
- [ ] **7. Zawartość `_meta`** — czy nie ma wewnętrznych identyfikatorów, które nie powinny wyjść?
- [ ] **8. Formuły** — czy wszystkie wciąż są formułami (a nie wartościami z `data_only`)?
- [ ] **9. Makra** — czy plik ma być `.xlsm`? Jeżeli tak — czy na pewno?
- [ ] **10. Ochrona** — czy arkusze wynikowe są chronione, żeby odbiorca nie „poprawił" danych?
- [ ] **11. Rozmiar i liczba wierszy** — czy wysyłany plik nie jest o rząd wielkości większy niż powinien?
- [ ] **12. Otworzenie w Excelu** — **i naprawdę obejrzenie.**

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Ściągawka nie zastępuje zrozumienia, ale porządkuje to, co zrozumiałeś.** Ten moduł nie nauczył Cię niczego nowego — pokazał Ci **mapę** tego, co już wiesz. Jego wartość nie tkwi w treści, a w **czasie dostępu**: odpowiedź w minutę zamiast godziny szukania w modułach. Dlatego warto mieć własną wersję tego pliku: **przepisane, a nie skopiowane** — bo to, co przepiszesz, zapamiętasz.

2. **Trzy kategorie prawdy są jedynym naprawdę trwałym elementem wiedzy o openpyxl.** Nie „jak wywołać `add_table`" (to znajdziesz w dokumentacji), ale **„czy ta funkcja przejdzie, przejdzie częściowo, czy zginie po cichu"**. Biblioteka się zmieni, API się rozszerzy, ale to rozróżnienie zostanie — bo wynika z architektury (openpyxl implementuje **format**, nie **program**), nie z implementacji. Dlatego `§4.2` (inwentarz P/C/L) jest najważniejszą tabelą w tym module.

3. **Kolejność diagnostyki to skrócenie cierpienia, nie procedura.** Sześć kroków, od najtańszego do najdroższego: Excel → struktura → typy → inwentarz → porównanie → środowisko. Krok pierwszy (widok formuł w Excelu) wyklucza **połowę** możliwych przyczyn w dziesięć sekund i jest regularnie pomijany, bo „przecież wiem, co tam jest". Ta wiedza to najdroższe założenie w całej diagnostyce.

4. **Osiem antywzorców z `§6.2` powtarza się w niemal każdym firmowym skrypcie openpyxl.** Style rozsypane po kodzie, `main()` na 400 linii, brak testów, raport jako źródło prawdy, fasada przeciekająca typy. **Każdy z nich ma jedno rozwiązanie** i nie jest to rozwiązanie techniczne, a organizacyjne: jedno miejsce na kolory, jedno miejsce na orkiestrację, jedno miejsce na reguły biznesowe, jedno miejsce na decyzję o zapisie. Kiedy zastanawiasz się „gdzie to umieścić", to prawie zawsze odpowiedź brzmi: **tam, gdzie jest już wszystko, co podobne** — a nie obok.

5. **Narzędzia, które sam napiszesz, są ważniejsze od wiedzy, którą przyswoisz.** `inwentarz()`, `ryzyka()`, `podsumowanie_skoroszytu()`, `porownaj_skoroszyty()`, `bezpieczny_zapis()` i `xlsxcheck` — to siedem rzeczy z tego modułu, które zostaną z Tobą na lata. Wiedza o API się zmienia (i zawsze wymaga sprawdzenia w dokumentacji wersji, którą masz zainstalowaną), ale **dobre narzędzie diagnostyczne nigdy nie traci wartości** — bo problem „nie wiem, co jest w tym pliku" istnieje niezależnie od tego, którą wersję biblioteki zainstalujesz.

## 8. Ściągawka modułu

```python
# ==================================================================
# A. IMPORTY, KTORYCH NAPRAWDE UZYWA SIE NA CO DZIEN
# ==================================================================
from io import BytesIO
from copy import copy
from pathlib import Path
from datetime import datetime, date, time, timedelta
from decimal import Decimal
import zipfile, os, shutil, tempfile, hashlib

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import (
    CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule,
    Top10Rule, AboveAverageRule, BelowAverageRule, DuplicateValuesRule, Rule,
)
from openpyxl.styles import (
    Alignment, Border, Font, NamedStyle, PatternFill, Protection, Side,
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import (
    get_column_letter, column_index_from_string, quote_sheetname,
    absolute_coordinate,
)
from openpyxl.utils.cell import range_boundaries, coordinate_to_tuple
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties


# ==================================================================
# B. MINIMALNY, POPRAWNY SKOROSZYT (wzorzec do kopiowania)
# ==================================================================
wb = Workbook()
wb.remove(wb.active)                       # USUN domyslny arkusz "Sheet"!

ws = wb.create_sheet("Dane")
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = "1F4E78"

ws.append(["Data", "Region", "Sprzedaz netto"])
for wiersz in dane_generator:              # GENERATOR, nie lista
    ws.append(list(wiersz))

# Format liczbowy zamiast "na oko wygladajacej liczby":
for wiersz in ws.iter_rows(min_row=2, min_col=3, max_col=3):
    wiersz[0].number_format = '#,##0.00 "zl"'

wb.active = 0                              # co widzi uzytkownik po otwarciu
wb.calculation.fullCalcOnLoad = True       # Excel przeliczy formuly przy otwarciu
wb.properties.title = "Raport"
wb.properties.description = f"schemat=v1; wsad={hash_wsadu}"
wb.save("output/raport.xlsx")


# ==================================================================
# C. MAPA "CHCE ZROBIC X" -> "WOLAM Y"
# ==================================================================
# nowy skoroszyt bez domyslnego arkusza ............ Workbook(); wb.remove(wb.active)
# otworzyc i zmodyfikowac .......................... load_workbook(p)  [NIE z data_only!]
# czytac milion wierszy ............................ load_workbook(p, read_only=True) + close()
# zapisac milion wierszy ........................... Workbook(write_only=True) + append()
# zachowac makra ................................... load_workbook(p, keep_vba=True) + .xlsm
# zapisac do pamieci / HTTP ........................ wb.save(BytesIO()) -> getvalue()
# policzyc ostatni wiersz danych ................... skan w gore od max_row (value is not None)
# wstawic wiersz / usunac wiersz ................... ws.insert_rows / ws.delete_rows
# przesunac zakres i referencje .................... ws.move_range(ref, rows=, cols=, translate=True)
# scalic komorki BEZ utraty danych ................. najpierw przepisz wartosci, potem merge_cells
# zamrozic naglowek ................................ ws.freeze_panes = "B2"
# autofiltr ........................................ ws.auto_filter.ref = "A1:H100"
# prawdziwa tabela z stylem ........................ Table(displayName=, ref=) + TableStyleInfo
# format liczby (waluta, procent) .................. cell.number_format = '...'
# zmienic JEDNA ceche istniejacego stylu ........... nowy = copy(cell.font); nowy.bold = True; cell.font = nowy
# styl nazwany, uzywany wszedzie ................... wb.add_named_style(...); cell.style = "Nazwa"
# podswietlic wartosci warunkowo ................... ws.conditional_formatting.add(zakres, CellIsRule(...))
# podswietlic CALY wiersz warunkowo ................ FormulaRule(formula=['$D2="Pilne"']) [BEZ znaku =]
# pasek danych dla wartosci ujemnych ............... DataBarRule(start_value=-0.5, end_value=0.5)
# skala kolorow .................................... ColorScaleRule(start_type="min", ..., end_type="max")
# wykres slupkowy .................................. BarChart() + add_data + set_categories + add_chart
# wykres z DWOCH osi ............................... drugi.y_axis.axId = 200; pierwszy += drugi
# kolor slupka ..................................... series.graphicalProperties = GraphicalProperties(solidFill="...")
# wstawic obraz / logo ............................. XLImage(p); img.width/height z JEDNEJ skali; add_image
# hiperlink ........................................ cell.hyperlink = "https://..."; cell.value = "Otworz"
# komentarz ........................................ cell.comment = Comment("tresc", "autor")
# lista rozwijana z innego arkusza ................. DataValidation(type="list", formula1="=Regiony!$A$2:$A$20")
# zarejestrowac walidacje .......................... ws.add_data_validation(dv)  [POTEM dv.add(zakres)]
# ochrona arkusza .................................. ws.protection.sheet = True; password = ...; enable()
# odblokowac komorki do wpisywania ................. cell.protection = Protection(locked=False)
# przygotowac do druku ............................. print_title_rows="1:1" + pageSetUpPr.fitToPage=True
# porownac dwa skoroszyty .......................... czesci XML po normalizacji [NIE bajt w bajt!]
# sprawdzic, co zginie przy zapisie ................ zipfile.namelist() + skan tresci arkuszy


# ==================================================================
# D. CZTERY ZDANIA, KTORE MUSISZ PAMIETAC
# ==================================================================
# 1. openpyxl NIE liczy formul - zapisuje je jako tekst. Cache czyta data_only=True.
# 2. Zapisu po data_only=True NIE ROBI SIE NIGDY - zamienia formuly na wartosci.
# 3. save() to PRZEPISANIE calego pliku - wszystko, czego openpyxl nie zna, znika.
# 4. Styl komorki jest WSPOLDZIELONY i niemutowalny - kopiuj i przypisuj nowy.


# ==================================================================
# E. NARZEDZIA DIAGNOSTYCZNE (skopiuj do wlasnego helpers.py)
# ==================================================================
def inwentarz(sciezka):
    """Czesci pliku i rozpakowane rozmiary."""
    sciezka = Path(sciezka)
    with zipfile.ZipFile(sciezka) as z:
        spis = {n: z.getinfo(n).file_size for n in z.namelist()}
    return {"plik": str(sciezka), "bajtow": sciezka.stat().st_size,
            "rozpakowane": sum(spis.values()), "spis": spis}


PREFIKSY_UTRATA = ("xl/slicers/", "xl/timelines/", "xl/queryTables/",
                   "xl/model/", "xl/ctrlProps/", "xl/activeX/", "customXml/")
PREFIKSY_CZESCIOWO = ("xl/charts/", "xl/drawings/", "xl/pivotTables/")


def ryzyka(sciezka):
    """Co moze zginac albo zostac uproszczone przy zapisie."""
    sciezka = Path(sciezka)
    with zipfile.ZipFile(sciezka) as z:
        nazwy = z.namelist()
        tresc = "".join(z.read(n).decode("utf-8", "ignore")
                        for n in nazwy
                        if n.startswith("xl/worksheets/") and n.endswith(".xml"))
    wynik = [f"UTRATA: {p}" for p in PREFIKSY_UTRATA if any(n.startswith(p) for n in nazwy)]
    wynik += [f"CZESCIOWO: {p}" for p in PREFIKSY_CZESCIOWO if any(n.startswith(p) for n in nazwy)]
    if "sparklineGroups" in tresc:                 # sparkline'y sa W XML ARKUSZA!
        wynik.append("UTRATA: sparkline'y")
    if "xl/vbaProject.bin" in nazwy:
        wynik.append("MAKRA: wczytuj z keep_vba=True")
    return tuple(wynik)


def podsumowanie_skoroszytu(sciezka):
    """Arkusze, wymiary, liczba formul - do testow regresji."""
    wb = load_workbook(sciezka, read_only=True)
    try:
        return {
            "arkusze": tuple(wb.sheetnames),
            "ukryte": tuple(n for n in wb.sheetnames if wb[n].sheet_state != "visible"),
            "wymiary": {n: wb[n].calculate_dimension() for n in wb.sheetnames},
        }
    finally:
        wb.close()                                  # OBOWIAZKOWE!


def bezpieczny_zapis(wb, cel, *, backup=True):
    """tmp -> weryfikacja -> backup -> os.replace. Atomowo."""
    cel = Path(cel); cel.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=f".{cel.stem}.", suffix=".tmp.xlsx",
                                    dir=cel.parent)     # dir= !!
    os.close(fd); tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        sprawdzony = load_workbook(tmp, read_only=True)
        try:
            if len(sprawdzony.sheetnames) != len(wb.sheetnames):
                raise RuntimeError("weryfikacja tmp: niezgodna liczba arkuszy")
        finally:
            sprawdzony.close()
        if backup and cel.exists():
            zn = datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(cel, cel.with_name(f"{cel.stem}.{zn}.bak{cel.suffix}"))
        os.replace(tmp, cel)
        return cel
    except PermissionError as blad:
        raise RuntimeError(
            f"Plik {cel.name} jest otwarty w innym programie. Zamknij go."
        ) from blad
    finally:
        tmp.unlink(missing_ok=True)


# ==================================================================
# F. TRZY ZASADY SANITACJI (modul 21 w trzech liniach)
# ==================================================================
# 1. Znaki wyzwalajace to nie tylko "=": rowniez "+", "-", "@", TAB, CR.
# 2. Waliduj na WEJSCIU do domeny (jedna funkcja), nie przy zapisie.
# 3. ODRZUCAJ z bledem, nie modyfikuj cicho - czlowiek ma zdecydowac.
ZNAKI_WYZWALAJACE = ("=", "+", "-", "@", "\t", "\r")


# ==================================================================
# G. CHECKLISTA PRZED WYSLANIEM PLIKU (12 punktow, sekcja 6.3)
# ==================================================================
# ukryte arkusze | metadane | komentarze | odwolania zewnetrzne | media
# definiowane nazwy | zawartosc _meta | formuly | makra | ochrona
# rozmiar | OTWORZENIE W EXCELU I OBEJRZENIE
```

## 9. Co dalej

### 9.1. Mapa kursu: problem → moduł

To jest najważniejsza tabela tego modułu, bo zamienia kurs w **podręcznik**. Szukaj po **problemie**, nie po nazwie modułu.

| Problem / pytanie | Moduł | Co konkretnie przeczytać |
|---|---|---|
| „Chcę zacząć od zera, mam działające środowisko" | 00 | instalacja, konwencje, zasada nie-nadpisywania źródła |
| „Którą bibliotekę wybrać?" | 01 | tabela porównawcza, kiedy nie openpyxl |
| „Co jest w środku pliku .xlsx?" | 02 | mapa części ZIP, `sharedStrings`, indeksy stylów |
| „Dlaczego `cell.font.bold = True` nie działa?" | 03, 09 | model mentalny, niemutowalność stylów |
| „Pierwszy skoroszyt: tworzę i zapisuję" | 04 | `Workbook`, `save`, `load_workbook`, ścieżki, `BytesIO` |
| „Kwota zapisana jako tekst / typy w komórkach" | 05 | dopuszczalne typy, `None` vs `0` vs `""`, daty i strefy |
| „Jak policzyć formułę bez Excela?" | 06 | brak silnika obliczeń, `data_only` i jego pułapki |
| „`max_row` kłamie / ostatni wiersz danych" | 07 | wymiary, iteracja, `insert_rows` bez przesuwania referencji |
| „Zarządzanie arkuszami, kopiowanie, ukrywanie" | 08 | `create_sheet`, `copy_worksheet` i czego nie kopiuje |
| „Kolor, obramowanie, wyrównanie, style nazwane" | 09 | model stylu, `NamedStyle`, rejestr stylów |
| „Waluta, procent, format daty, dziwny wygląd liczby" | 10 | anatomia kodu formatu, `[$-xxx]`, locale |
| „Raport do druku / PDF / scalanie / zamrażanie" | 11 | `print_title_rows`, `fitToPage`, `merge_cells` |
| „Chcę prawdziwą tabelę Excel z filtrem" | 12 | `Table`, `TableStyleInfo`, structured references |
| „Podświetlić wiersz warunkowo / reguły kolidują" | 13 | `dxf`, priorytety, `FormulaRule` bez `=` |
| „Pasek danych dla wartości ujemnych / ikony / skala" | 14 | `DataBarRule`, `ColorScaleRule`, `IconSetRule` |
| „Wykres dwuosiowy / dlaczego linia jest płaska" | 15 | `axId`, `crosses`, `+=`, `graphicalProperties` |
| „Logo rozciągnięte / obrazki / hiperlinki / komentarze" | 16 | jednostki obrazów, anchory, komentarze w scalonych |
| „Lista rozwijana z innego arkusza / ochrona" | 17 | `DataValidation`, `protection`, hasło ≠ szyfrowanie |
| „Zmodyfikować istniejący plik i nic nie stracić" | **18** | inwentarz utraty, `load_workbook` flagi, atomowy zapis |
| „200 000 wierszy / brak pamięci / wolno" | 19 | `read_only`, `write_only`, `lxml`, generator |
| „Jak to przetestować? CI bez Excela" | 20 | `BytesIO`, `pytest`, testy regresji struktury |
| „Bezpieczeństwo: dane od użytkownika" | 21 | formula injection, zip bomb, audyt przed wysyłką |
| „Dlaczego mój skrypt ma 900 linii i nikt go nie rusza?" | 22 | poziomy abstrakcji, YAGNI, trzy światy |
| „Fabryka, Builder, szablon, rejestr stylów" | 23 | wzorce konstrukcyjne — konkretne przykłady |
| „Fasada, adapter, dekorator, proxy, sekcje raportu" | 24 | wzorce strukturalne — brak przecieku abstrakcji |
| „Za dużo `if`-ów: strategie, komendy, reguły" | 25 | wzorce behawioralne, `StosKomend`, Specification |
| „Gdzie umieścić openpyxl w architekturze systemu?" | 26 | warstwy, Ports & Adapters, transakcyjność, DI |
| „Zbudować prawdziwy generator raportów" | 27 | projekt końcowy, kamienie milowe, rubryka oceny |
| „Wracam po miesiącach — gdzie to było?" | **28** | ten moduł |

### 9.2. Trzy drogi wejścia do kursu (jak używać go dalej)

**Droga pierwsza — „mam problem".** Wejdź w tabelę `§9.1`, znajdź problem, otwórz **jeden** moduł, przeczytaj sekcje 2 i 3, sprawdź sekcję 6 (pułapki). Nie czytaj modułu od początku. Kurs jest zbudowany tak, żeby dało się wejść w środek — każdy moduł ma na górze listę „Wymaga: moduły…" i sekcję „Co dalej", więc **odwołania są obustronne**.

**Droga druga — „mam objaw".** Wejdź w `§6.1` (Top 20) albo `§3.16` (FAQ). Jeżeli objaw nie pasuje do żadnego wpisu — idź do `§2.3` (kolejność diagnostyki) i przejdź sześć kroków. W **90 %** przypadków krok 1 albo krok 3 daje odpowiedź. Jeżeli nie — użyj `xlsxcheck` z `§5`.

**Droga trzecia — „uczę kogoś / piszę standard w firmie".** Wtedy kolejność modułów ma znaczenie: 00–03 (fundamenty), 04–08 (podstawy), 09–14 (formatowanie), 15–17 (wzbogacanie), **18** (modyfikacja istniejących — obowiązkowe dla każdego), 19–21 (jakość), 22–26 (architektura), 27 (projekt), 28 (referencja). Minimum dla zespołu, który „robi raporty w Excelu z Pythona": **03, 06, 09, 13, 18, 20, 21**.

### 9.3. Źródła i dalsza nauka

**Dokumentacja openpyxl 3.1.x** — punkt wyjścia i punkt powrotu. Cztery sekcje, które warto przeczytać **w całości**, bo w kursie tylko je zasygnalizowałem: *Working with Styles*, *Conditional Formatting*, *Charts*, *Performance*. Czytaj **dokumentację wersji, którą masz zainstalowaną** — dokumentacja „latest" opisuje to, czego w Twojej wersji może jeszcze nie być.

**Changelog openpyxl (`.rst` w repozytorium albo wydania)** — **czytaj przed użyciem egzotycznego parametru**. Reguła praktyczna, którą warto przyjąć: jeżeli parametr nie jest wspomniany w samouczku, a chcesz go użyć, **sprawdź w changelogu, kiedy został dodany**. To oszczędza sytuacji „u mnie działa, u kolegi `TypeError: unexpected keyword argument`", a potem dwóch godzin ustalania, kto ma starszą wersję.

**Specyfikacja OOXML** (`ECMA-376`, `ISO/IEC 29500`) — nie do czytania od początku, tylko jako **słownik**. Gdy openpyxl czegoś nie obsługuje, mapa części pliku (`§2.4`) pozwala znaleźć nazwę elementu, a specyfikacja mówi, co ten element znaczy. W praktyce wystarcza wyszukiwarka i sekcja o SpreadsheetML.

**Wzorce projektowe: GoF („Design Patterns") i Martin Fowler („Patterns of Enterprise Application Architecture")** — konkretnie: Factory, Builder, Prototype (konstrukcyjne), Facade, Adapter, Decorator (strukturalne), Strategy, Template Method, Command, Visitor, Specification (behawioralne) oraz Repository, Unit of Work, Data Transfer Object (Fowler). W kursie pokazałem je na openpyxl — w książkach zobaczysz je **bez kontekstu Excela**, co jest potrzebne, żeby zrozumieć, kiedy **nie** ich używać.

**Porównanie z pandas** — warto znać granicę: pandas jest do **analizy danych** (agregacje, przekształcenia, łączenie zbiorów), openpyxl do **formy prezentacji** (wygląd, wykresy, ochrona, druk). Praktyczna reguła: **gdy wszystko, czego potrzebujesz, to liczby i tabele — pandas + `dataframe_to_rows`; gdy potrzebujesz dokumentu, który ktoś otworzy i wydrukuje — openpyxl.** Nie mieszaj: wczytywanie przez pandas, a potem formatowanie przez openpyxl jest sensowne tylko wtedy, gdy wyraźnie **rozgraniczysz** etap czyszczenia od etapu prezentacji.

**Specyfika merytoryczna pozostałych bibliotek** — jeżeli natrafisz na wymaganie, którego openpyxl nie spełnia: `xlsxwriter` (sparkline'y, bogatsze wykresy, brak odczytu), `xlwings`/COM (interakcja z otwartym Excelem, przeliczanie), LibreOffice headless (konwersja i przeliczanie), `msoffcrypto-tool` (odczyt i szyfrowanie plików). Każda z nich jest **uzupełnieniem**, nie zastępstwem.

**I ostatnia rzecz — źródło, którego nie znajdziesz w żadnej dokumentacji: pliki, które sam wygenerowałeś.** Trzy miesiące po wdrożeniu otwórz `output/`, weź **losowy** plik z tego katalogu i uruchom na nim `podsumowanie_skoroszytu()` oraz `ryzyka()`. Zobacz, czy `_meta` odpowiada na pytanie „skąd to i czy jest kompletne", czy arkusze mają to, co powinny, i czy nie ma tam rzeczy, których **nie** chciałeś zapisywać. To jest test, który sprawdza nie bibliotekę i nie Twój kod, a **decyzje projektowe** — i to on mówi najwięcej o tym, czy kurs przyniósł efekt.