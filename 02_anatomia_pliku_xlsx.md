# Moduł 02 — Anatomia pliku .xlsx: co naprawdę jest w środku

> **Część:** 0 — Fundamenty i model mentalny · **Poziom:** ⭐ · **Wymaga:** modułów 00–01

## 0. W tym module nauczysz się

- Rozumiesz, że plik `.xlsx` to **archiwum ZIP** wypełnione plikami XML, i umiesz to sprawdzić jednym poleceniem — bez openpyxl i bez Excela.
- Znasz mapę pakietu: co siedzi w `[Content_Types].xml`, `_rels/`, `xl/workbook.xml`, `xl/worksheets/sheetN.xml`, `xl/styles.xml`, `xl/sharedStrings.xml`, `xl/media/` — i jaką rolę odgrywa każda część.
- Rozumiesz, dlaczego **komórka nie „ma" formatowania** — ma tylko numer stylu, a cały wygląd siedzi w `xl/styles.xml`. Umiesz prześledzić ten łańcuch ręcznie, krok po kroku.
- Wiesz, że w pliku zapisanym przez openpyxl **nie ma części `sharedStrings.xml`** — openpyxl zapisuje teksty wprost w komórkach — a w pliku z Excela ta część jest. Umiesz sprawdzić to samodzielnie i wiesz, co z tego wynika dla rozmiaru pliku i wydajności.
- Odróżniasz **koordynaty 1-indeksowane** (adres `B2`, wiersz `r="4"`) od **indeksów 0-indeksowanych** (styl `s="3"`, łańcuch `t="s"`). Ten sam plik używa obu systemów jednocześnie — i to jest źródło wielu pomyłek w dalszej części kursu.
- Rozpoznajesz po nazwie części pakietu **co openpyxl zgubi przy zapisie**: `vbaProject.bin`, `slicer*`, `timeline*`, `pivotCache*` oraz bloki rozszerzeń `extLst` (sparkline'y).
- Umiesz napisać własne narzędzie diagnostyczne, które „prześwietla" skoroszyt bez otwierania Excela — i które w module 18 uratuje Ci godziny pracy.

## 1. Intuicja i analogia

Zacznij od najprostszego eksperymentu, jaki możesz wykonać natychmiast, bez ani jednej linii Pythona.

Weź dowolny plik `.xlsx`. Zmień jego rozszerzenie na `.zip`. Spróbuj go rozpakować — Windows, macOS, Linux, dowolny program do archiwów. **Rozpakuje się bez protestu.**

To nie jest sztuczka. To nie jest „hack". To jest fakt o naturze pliku. `.xlsx` **jest** archiwum ZIP. Rozszerzenie `.xlsx` mówi tylko dwie rzeczy: „to archiwum ZIP" oraz „w środku jest zestaw dokumentów XML opisanych wspólnym standardem OOXML (Office Open XML), a dokładniej jego częścią tabelaryczną — SpreadsheetML".

**Analogia centralna tego modułu: skoroszyt to karton z teczką aktową.**

Wyobraź sobie, że dostajesz zamknięty karton. Otwierasz go i widzisz dobrze zorganizowaną teczkę z dokumentami w segregatorze. W tej teczce:

- Na wieczku teczki przyklejona jest **metryka**: lista wszystkich dokumentów w środku wraz z informacją, jakiego typu każdy z nich jest (`[Content_Types].xml`). Bez tej metryki nikt nie wie, co ma w rękach.
- Pod metryką jest **spinka wskazująca główny dokument** (`.rels/.rels`) — mówi „wszystko zaczyna się od tego pliku, idź do `xl/workbook.xml`".
- Otwierasz główny dokument i znajdujesz **spis treści** (`xl/workbook.xml`): lista kartek, które tworzą skoroszyt, każda z nazwą i numerem identyfikacyjnym.
- Przy każdej pozycji spisu treści jest **pinezka** prowadząca do fizycznej kartki (`xl/_rels/workbook.xml.rels`). Pozycja „Dane" ma pinezkę `rId1`, która prowadzi do `worksheets/sheet1.xml`. Pozycja „Podsumowanie" ma pinezkę `rId2`, która prowadzi do `worksheets/sheet2.xml`.
- Są **kartki** (`xl/worksheets/sheetN.xml`). Kartka to siatka kratek. **Na kartce nie ma jednak tekstu** — są tylko numery. Tekst (a raczej informacja, gdzie go szukać) jest osobno.
- Jest **katalog tekstów** (`xl/sharedStrings.xml`). Tutaj, w jednym miejscu, spisane są wszystkie napisy użyte w całym skoroszycie, każdy z własnym numerem. Na kartce w kratce `A2` widnieje liczba `7` — i to znaczy „siódma pozycja w katalogu tekstów". Dlaczego tak? Bo jeśli słowo „Warszawa" pojawia się w 400 komórkach, zapisanie go 400 razy byłoby marnotrawstwem. Zamiast tego katalog ma je raz, a kartki odsyłają do katalogu.
- Jest **legenda wyglądu** (`xl/styles.xml`). I tu rzecz najważniejsza: **na kartce nie ma ani jednego koloru, ani jednej grubości czcionki, ani jednego obramowania.** W kratce jest tylko odsyłacz: „styl nr 4". Legenda wyglądu mówi, że styl nr 4 to pogrubiona czcionka Calibri 12 punktów na szarym tle z cienką linią na dole.
- W osobnej **koszulce na dokumenty** są załączniki (`xl/media/` — obrazy), **rysunki techniczne** (`xl/drawings/`) i **diagramy** (`xl/charts/`).
- Wreszcie, w zapieczętowanej kopercie przyklejonej do wewnętrznej strony okładki, jest **projekt VBA** (`xl/vbaProject.bin`) — makra, do których openpyxl nawet nie zagląda, a które przy zapisie albo się przeniosą (jeśli poprosisz), albo przepadną.

Karton wokół tego wszystkiego to ZIP. Cała zawartość — XML. Zero plików binarnych w warstwie danych.

Dwie analogie uzupełniające, które będą wracać w dalszych modułach:

> **Styl komórki to bilet z numerem miejsca.** Wyobraź sobie widza w teatrze, który trzyma kartkę z napisem „miejsce 12". Ta kartka nie opisuje, jak wygląda miejsce 12 — nie mówi, czy fotel jest zielony czy czerwony, czy jest wyżej czy niżej. Mówi tylko „idź do miejsca 12". Dopiero bileter (plik `styles.xml`) wie, gdzie to miejsce jest i jak wygląda. Dlatego kolor tła **nie jest** cechą komórki. Jest cechą stylu, na który komórka wskazuje.

> **Katalog tekstów to cennik w sklepie.** Nie drukujesz opisu towaru na każdym paragonie — drukujesz **kod towaru**. pełny opis jest raz, w katalogu. Właśnie tak działa `sharedStrings.xml`.

I jeszcze jedna uwaga, która jest sednem całego modułu:

> Obie techniki — katalog tekstów i zapis wprost w komórce — są **poprawnym OOXML**. Format dopuszcza jedno i drugie. Excel wybiera katalog. openpyxl wybiera zapis wprost. Konsekwencje praktyczne (rozmiar pliku, wydajność, kompatybilność z innymi narzędziami) są różne — i tym zajmiemy się w tym module.

Skoro plik jest archiwum ZIP, to najciekawsze pytanie brzmi: **jakie części openpyxl czyta, które zapisuje, a których nie zna wcale?** To pytanie będzie osią całego kursu. Dzisiaj zdobędziesz narzędzia, żeby odpowiedzieć na nie samodzielnie dla dowolnego pliku.

## 2. Teoria

### 2.1. ZIP pierwszy — sprawdź sam, zanim uwierzysz

Nie wierz mi — sprawdź. Oto kompletny program, który to zadanie wykonuje:

```python
"""Dowód nr 1: .xlsx to archiwum ZIP."""
import zipfile
from pathlib import Path

sciezka = Path("output/02_inwentarz_demo.xlsx")

# is_zipfile czyta tylko nagłówek archiwum. Jest natychmiastowe,
# nawet dla pliku wielkości gigabajta.
print("Czy to ZIP?", zipfile.is_zipfile(sciezka))

with zipfile.ZipFile(sciezka) as archiwum:
    for nazwa in archiwum.namelist():
        print(" ", nazwa)
```

Jeśli uruchomisz to na pliku utworzonym przez openpyxl, zobaczysz coś takiego:

```text
Czy to ZIP? True
  [Content_Types].xml
  _rels/.rels
  docProps/core.xml
  docProps/app.xml
  xl/workbook.xml
  xl/_rels/workbook.xml.rels
  xl/worksheets/sheet1.xml
  xl/styles.xml
  xl/theme/theme1.xml
```

Zwróć uwagę na jedną rzecz: **nie ma tam `xl/sharedStrings.xml`.** Zapamiętaj tę obserwację — wrócimy do niej z pełnym wyjaśnieniem w sekcji 2.7. To jedna z tych informacji, które odróżniają osobę, która „używa openpyxl", od osoby, która *rozumie* openpyxl.

### 2.2. Pełna mapa pakietu

Poniższa tabela to najważniejsza rzecz w tym module. To legenda do wszystkich plików `.xlsx`, z jakimi się zetkniesz. Kolumna „Kategoria" używa języka z modułu 01.

| Część pakietu | Rola | Kategoria (openpyxl) |
|---|---|---|
| `[Content_Types].xml` | Spis zawartości — deklaruje typ każdej części | potrafi |
| `_rels/.rels` | Punkt startowy — wskazuje główny dokument pakietu | potrafi |
| `docProps/core.xml` | Metadane podstawowe: autor, tytuł, data utworzenia | potrafi |
| `docProps/app.xml` | Metadane aplikacji: nazwa programu, wersja, tytuły arkuszy | potrafi częściowo (generowane od nowa) |
| `docProps/custom.xml` | Własne właściwości dokumentu (nowość 3.1) | potrafi |
| `xl/workbook.xml` | Spis treści: lista arkuszy, nazwy zdefiniowane, ustawienia obliczeń | potrafi |
| `xl/_rels/workbook.xml.rels` | Pinezki z skoroszytu do arkuszy, stylów, katalogu tekstów, motywu | potrafi |
| `xl/worksheets/sheetN.xml` | Kartka: komórki, wymiary kolumn, scalenia, reguły, ustawienia druku | potrafi |
| `xl/worksheets/_rels/sheetN.xml.rels` | Pinezki z kartki do rysunków, tabel, tabel przestawnych | potrafi |
| `xl/styles.xml` | Legenda wyglądu: czcionki, wypełnienia, obramowania, formaty liczb, style warunkowe | potrafi |
| `xl/sharedStrings.xml` | Katalog tekstów (używany przez Excela; **openpyxl go nie tworzy**) | potrafi odczytać |
| `xl/theme/theme1.xml` | Motyw: kolory i czcionki tematyczne | potrafi |
| `xl/calcChain.xml` | Kolejność przeliczania formuł (podpowiedź optymalizacyjna dla Excela) | potrafi częściowo / ryzykowne |
| `xl/drawings/drawingN.xml` | Warstwa rysunkowa: rozmieszczenie obrazów, wykresów, kształtów | potrafi częściowo |
| `xl/media/imageN.png` | Binarny plik obrazu wstawionego do arkusza | potrafi |
| `xl/charts/chartN.xml` | Definicja wykresu: seria, osie, tytuł, formatowanie | potrafi częściowo |
| `xl/tables/tableN.xml` | Tabela Excela (`Table` z nagłówkami i filtrem) | potrafi |
| `xl/commentsN.xml`, `xl/threadedComments/` | Komentarze klasyczne i nowoczesne | potrafi |
| `xl/pivotTables/`, `xl/pivotCache/` | Tabela przestawna i jej pamięć podręczna | potrafi częściowo (odczyt i zapis, bez tworzenia od zera) |
| `xl/externalLinks/` | Odwołania do innych skoroszytów | potrafi częściowo (flaga `keep_links`) |
| `xl/vbaProject.bin` | Projekt makr VBA | potrafi częściowo (tylko z `keep_vba=True`) |
| `xl/ctrlProps/`, `xl/activeX/`, `customUI/` | Właściwości formantów, ActiveX, własna wstążka | gubi przy zapisie (zachowuje tylko w ścieżce VBA) |
| `xl/slicers/`, `xl/slicerCaches/` | Slicery (filtry wizualne) | gubi przy zapisie |
| `xl/timelines/` | Osie czasu | gubi przy zapisie |
| `xl/connections.xml`, `xl/queryTables/`, `xl/model/` | Power Query, Power Pivot, model danych | gubi przy zapisie |
| `customXml/` | Niestandardowe dane XML powiązane z komórkami | gubi przy zapisie |
| bloki `<extLst>` w arkuszach | Rozszerzenia, m.in. sparkline'y | gubi przy zapisie |

Nie musisz tego zapamiętywać. Wróć do tej tabeli po module 18 i przekonasz się, że każdy wiersz ma swoje miejsce w jakiejś historii o utraconych danych.

### 2.3. `[Content_Types].xml` — metryka na wieczku teczki

To plik, od którego zaczyna się czytanie pakietu. W OOXML **każda część musi być zadeklarowana** — jeśli jakaś część leży w archiwum, ale nie ma jej w metryce, to dla odbiorcy (Excela, openpyxl, dowolnego parsera) ona nie istnieje.

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
           ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/theme/theme1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
```

Widzisz dwa rodzaje deklaracji:

- **`<Default>`** — działa po rozszerzeniu pliku. „Każda część kończąca się na `.rels` jest plikiem relacji. Każda część `.xml` to ogólnie XML." To deklaracja zbiorcza, tania w zapisie.
- **`<Override>`** — działa po pełnej nazwie części. „Konkretnie `/xl/workbook.xml` to główny skoroszyt. Konkretnie `/xl/worksheets/sheet1.xml` to arkusz."

Dlaczego typ jest taki ważny? Bo `sheet1.xml` i `styles.xml` wyglądają dla parsującego XML identycznie — oba są dokumentami XML. Dopiero typ z metryki mówi „ten plik to arkusz, tamten to legenda wyglądu". Bez tego Excela nie da się poprowadzić.

**Konsekwencja praktyczna:** jeśli kiedykolwiek zbudujesz pakiet ręcznie (a zrobisz to w ćwiczeniu), zapomnienie o wpisie w `[Content_Types].xml` spowoduje, że Excel nie znajdzie części. Objaw: komunikat „Znaleziono problem z zawartością" i próba naprawy.

### 2.4. `_rels/` — spinki i relacje

OOXML nie lubi odwołań po ścieżkach. Nie lubi też odwołań cyklicznych. Wprowadza więc pośrednika: **relacje**. Każda część, która chce wskazać inną część, robi to przez identyfikator relacji — nigdy bezpośrednio przez nazwę pliku.

Na szczycie pakietu jest `_rels/.rels` (kropka na początku nazwy, tak, to poprawna nazwa):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                Target="xl/workbook.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
                Target="docProps/core.xml"/>
</Relationships>
```

Czytasz to tak: „nie wiem, jaki jest główny dokument tego pakietu — ale sprawdzę relację typu `officeDocument` i pójdę tam, gdzie wskazuje jej `Target`".

I dalej, w `xl/_rels/workbook.xml.rels`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                Target="styles.xml"/>
  <Relationship Id="rId3"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                Target="theme/theme1.xml"/>
</Relationships>
```

A w samym skoroszycie (`xl/workbook.xml`) arkusze odwołują się już tylko przez ID:

```xml
<sheet name="Dane" sheetId="1" r:id="rId1"/>
```

**Po co ta cała warstwa?** Z trzech powodów:

1. **Przenośność.** Część można wstawić w dowolne miejsce w strukturze katalogów, a relacja się nie zepsuje (albo zepsuje się tylko jej `Target`).
2. **Jednoznaczność.** Cztery różne arkusze mogą nazywać się podobnie; ID są unikalne.
3. **Rozszerzalność.** Można dodać nowy typ części bez zmiany formatu reszty — wystarczy nowy typ relacji.

**Analogia:** relacje to spinki w segregatorze. Pozycja w spisie treści mówi „rozdział trzeci", a pinezka prowadzi do fizycznej kartki. Jeśli wyjmiesz kartkę, a nie zaktualizujesz spinki — spis treści wskaże donikąd. Dokładnie to się dzieje, gdy usuniesz arkusz, do którego odwołuje się nazwa zdefiniowana albo tabela (moduł 08).

### 2.5. `xl/workbook.xml` — spis treści skoroszytu

To niewielki plik, który mówi wszystko o strukturze na najwyższym poziomie:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr/>
  <bookViews>
    <workbookView activeTab="1"/>
  </bookViews>
  <sheets>
    <sheet name="Dane" sheetId="1" r:id="rId1"/>
    <sheet name="Podsumowanie" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="SumaRoczna">'Dane'!$B$2:$B$13</definedName>
  </definedNames>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>
```

Trzy elementy zasługują na uwagę:

- **`<sheet name=... sheetId=... r:id=...>`** — kolejność elementów `<sheet>` to **kolejność zakładek w Excelu**. `name` to nazwa widoczna dla użytkownika. `sheetId` to historyczny identyfikator. `r:id` to spinka do fizycznego pliku arkusza w `_rels`. Nazwa arkusza, którą widzisz w `wb.sheetnames`, pochodzi stąd.
- **`<definedNames>`** — nazwy zdefiniowane (named ranges). To jeden z obszarów, który w openpyxl 3.1 został przebudowany; stąd nowe API `DefinedName` i `wb.defined_names.add(...)`.
- **`<calcPr calcId="124519" fullCalcOnLoad="1"/>`** — ustawienia obliczeń. I tu ciekawostka z modułu 01: openpyxl domyślnie zapisuje `fullCalcOnLoad=True`. W klasie `openpyxl.workbook.properties.CalcProperties` wartość domyślna to `fullCalcOnLoad=True`, `calcId=124519`. Dzięki temu Excel po otwarciu pliku przeliczy wszystkie formuły od nowa — co jest dokładnie tym, czego potrzebujesz, jeśli openpyxl nie potrafi ich policzyć. (Sprawdź w dokumentacji swojej wersji, czy nazwy parametrów się nie zmieniły.)

### 2.6. `xl/worksheets/sheetN.xml` — kartka

To plik, w którym żyją Twoje dane. Struktura jest hierarchiczna i nazwy elementów mówią same za siebie:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:C3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43"/>
  <cols>
    <col min="1" max="1" width="18.5" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>120000</v></c>
      <c r="C1" s="1"><v>0.18</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>1</v></c>
      <c r="B2"><v>98500</v></c>
      <c r="C2" s="1"><v>0.21</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>2</v></c>
      <c r="B3"><v>76200</v></c>
      <c r="C3" s="1"><v>0.15</v></c>
    </row>
  </sheetData>
  <mergeCells count="0"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
```

Omówmy elementy po kolei.

**`<dimension ref="A1:C3"/>`** — obszar, w którym są dane. To podpowiedź optymalizacyjna: parser wie, jak dużą siatkę ma przygotować. Nie jest to informacja wiążąca.

**`<sheetViews>`** — widok: który arkusz był aktywny, jaki poziom powiększenia, czy linie siatki są widoczne, gdzie jest zaznaczenie.

**`<sheetFormatPr>`** — domyślna wysokość wiersza i domyślna szerokość kolumny.

**`<cols>`** — nadpisania szerokości dla konkretnych kolumn. Uwaga na pułapkę: **`min` i `max` tutaj są 1-indeksowane**, a `width` jest wyrażona w jednostkach „znaków", nie w pikselach.

**`<sheetData>`** — serce pliku. Zawiera `<row>` (z 1-indeksowanym `r`), a każdy wiersz zawiera `<c>` — komórki.

I tu dochodzimy do najważniejszego fragmentu tego modułu. Komórka:

```xml
<c r="C1" s="1" t="s"><v>0</v></c>
```

Cztery kawałki:

| Atrybut/element | Znaczenie | Indeksowanie |
|---|---|---|
| `r="C1"` | Koordynat — adres komórki w notacji A1 | **1-indeksowany** (kolumna A = 1, wiersz 1 = 1) |
| `s="1"` | Numer stylu — odsyłacz do `cellXfs` w `styles.xml` | **0-indeksowany** |
| `t="s"` | Typ danych: `s` = współdzielony łańcuch | — |
| `<v>0</v>` | Wartość: **indeks do katalogu tekstów** | **0-indeksowany** |

**To jest najważniejsza obserwacja w tym module.** W jednym, maleńkim fragmencie XML oba systemy indeksowania sąsiadują ze sobą:

- `C1` — litera C to trzecia kolumna, ale w adresie nie ma żadnej liczby; wiersz 1 to pierwszy wiersz. W notacji A1 liczymy od 1.
- `s="1"` — to **drugi** styl w liście (bo pierwszy ma numer 0).
- `<v>0</v>` — to **pierwszy** tekst w katalogu (bo katalog liczy od 0).

W praktyce: `A1` to `row=1, column=1`. Ale `s="0"` to **pierwszy** styl, a `<v>0</v>` w komórce ze `t="s"` to **pierwszy** tekst. Jeśli pomyślisz „komórka wskazuje styl numer 1, więc to drugi styl" — masz rację. Jeśli pomyślisz „`<v>1</v>` to drugi tekst w katalogu" — też masz rację. A jeśli w tym samym zdaniu powiesz „kolumna B, więc drugi element" — też. Ale gdy zaczniesz przenosić te intuicje na arytmetykę w kodzie, bardzo łatwo o off-by-one. Wrócimy do tego w module 03, gdzie zbiorę to w reguły.

**Typy danych (`t`):**

| Wartość `t` | Znaczenie | Zawartość |
|---|---|---|
| brak (domyślnie `n`) | liczba | `<v>42</v>` lub `<v>3.14</v>` |
| `n` | liczba (jawnie) | jak wyżej |
| `s` | współdzielony łańcuch | `<v>N</v>` — 0-indeksowany odsyłacz do katalogu tekstów |
| `inlineStr` | tekst zapisany na miejscu | `<is><t>Ala</t></is>` |
| `str` | wynik formuły zwracającej tekst | `<v>Ala</v>` |
| `b` | wartość logiczna | `<v>0</v>` lub `<v>1</v>` |
| `e` | błąd | `<v>#DIV/0!</v>` |
| `d` | data w formacie ISO 8601 | rzadko używane |

Dla liczb openpyxl zapisuje typ jawnie (`t="n"`), a Excel często pomija ten atrybut i polega na domyślnej wartości. **Oba warianty są poprawne.** Dlatego w kodzie parsującym zawsze używaj `element.get("t", "n")` — z wartością domyślną.

**Formuły.** Formuła to element `<f>` wewnątrz komórki, bez wiodącego znaku `=`:

```xml
<c r="D1"><f>SUM(B1:B3)</f><v>294700</v></c>
```

Zwróć uwagę na `<v>` obok `<f>`. To jest **wartość z pamięci podręcznej** — liczba, którą Excel zapisał przy ostatnim przeliczeniu. Jeśli plik powstał w Excelu, `<v>` tam będzie. Jeśli plik powstał w openpyxl, `<v>` **nie będzie** (bo openpyxl nie ma czym jej policzyć). To dokładnie wyjaśnia różnicę w zachowaniu `data_only=True`, którą widziałeś w module 01 — i zaraz zobaczymy to jeszcze raz w przykładzie 4 tego modułu.

### 2.7. `xl/styles.xml` — legenda wyglądu, czyli dlaczego komórka nie ma koloru

To najczęściej źle rozumiana część pakietu. Zbudujmy ją od podstaw.

Wyobraź sobie plik z 10 000 komórek, z których 3 000 ma pogrubioną czcionkę. Gdyby każda komórka opisywała swój wygląd samodzielnie, w pliku XML pojawiłoby się 3 000 kopii tego samego opisu. Marnotrawstwo. Dlatego OOXML robi to, co bazy danych: **normalizuje**. Wszystkie unikalne czcionki, wypełnienia i obramowania są zdefiniowane raz, w `styles.xml`, każda z własnym numerem. Komórka nie trzyma opisu czcionki — trzyma **numer stylu**.

A oto struktura `xl/styles.xml` (uproszczona, ale zgodna z prawdą):

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00 &quot;zł&quot;"/>
  </numFmts>

  <fonts count="3">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>

  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid">
      <fgColor rgb="FF1F4E78"/><bgColor indexed="64"/>
    </patternFill></fill>
  </fills>

  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFBFBFBF"/></left>
      <right style="thin"><color rgb="FFBFBFBF"/></right>
      <top style="thin"><color rgb="FFBFBFBF"/></top>
      <bottom style="thin"><color rgb="FFBFBFBF"/></bottom>
      <diagonal/>
    </border>
  </borders>

  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>

  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>

  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>

  <dxfs count="1">
    <dxf>
      <font><color rgb="FF9C0006"/></font>
      <fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>
    </dxf>
  </dxfs>

</styleSheet>
```

Prześledźmy to od końca do początku, bo tak działa odczyt.

**`<dxfs>`** — „differential styles". To style dla **formatowania warunkowego**. Zamiast pełnego stylu niosą tylko *różnicę* względem stylu komórki. Reguła formatowania warunkowego (która żyje w arkuszu, w `<conditionalFormatting>`) odsyła do `dxfs[N]` przez atrybut `dxfId`. **Konsekwencja: kolor z formatowania warunkowego nie jest w komórce, ani nawet w stylu komórki — jest w regule w arkuszu plus w `dxf` w `styles.xml`.** Wrócimy do tego w modułach 13 i 14.

**`<cellStyles>`** — style nazwane widoczne na liście w Excelu (Normalny, Obliczenia, Dobry, Zły...). Odsyłają do `cellStyleXfs`.

**`<cellXfs>`** — **to jest ten element, do którego odsyła `s="N"` w komórce.** Każdy `<xf>` to jeden pełny zestaw: „format liczby numer X, czcionka numer Y, wypełnienie numer Z, obramowanie numer W".

**`<cellStyleXfs>`** — style bazowe, na których opierają się style nazwane. Zwykle jeden, domyślny.

**`<fonts>`, `<fills>`, `<borders>`, `<numFmts>`** — magazyny. Każdy wpis ma swój numer wynikający z **kolejności na liście**, liczonej od zera.

Teraz najważniejsza rzecz w tym module — **rozwiązywanie łańcucha**. Załóżmy, że w komórce widzisz `s="2"`. Jak dojść do tego, jak ta komórka wygląda?

1. `s="2"` → bierzemy **trzeci** element z listy `<cellXfs>` (bo indeks 0 = pierwszy, 1 = drugi, 2 = trzeci).
2. W tym `<xf>` czytamy `fontId="2"`, `fillId="2"`, `borderId="1"`, `numFmtId="0"`.
3. `fontId="2"` → **trzeci** element z listy `<fonts>`: `<b/><sz val="14"/><color rgb="FFFFFFFF"/>` — pogrubiona Calibri 14 punktów, biała.
4. `fillId="2"` → **trzeci** element z listy `<fills>`: jednolity kolor `FF1F4E78` (ciemnoniebieski).
5. `borderId="1"` → **drugi** element z listy `<borders>`: cienka szara obwódka ze wszystkich czterech stron.
6. `numFmtId="0"` → format wbudowany `General`.

I teraz kluczowa informacja, którą wyjaśnia cała ta sekcja: **`FF1F4E78` i „pogrubiona Calibri 14" nie istnieją w komórce.** Komórka ma tam literę `C`, cyfrę `1` i cyfrę `2`. Cały wygląd siedzi gdzie indziej i jest współdzielony ze wszystkimi innymi komórkami, które wskazują ten sam styl.

**Analogia:** komórka trzyma **bilet z numerem miejsca**, a nie opis fotela. Zmieniasz „miejsce 12" ze zielonego na czerwone — i wszystkie bilety z numerem 12 nagle wskazują na czerwony fotel. To jest jednocześnie siła tego modelu (jeden wpis zmienia tysiąc komórek) i źródło niespodzianek (zmiana, której się nie spodziewałeś, bo nie wiedziałeś, że styl jest współdzielony).

I jeszcze dwie pułapki tego pliku, o których musisz wiedzieć, bo wracają w module 09:

**Pułapka 1: `fills` mają zarezerwowane dwa pierwsze miejsca.** Excel rezerwuje `fills[0]` (patternType="none") i `fills[1]` (patternType="gray125") niezależnie od tego, czy są używane. **Pierwsze niestandardowe wypełnienie, jakie dodasz, trafi na indeks 2, nie 0.** To dlatego w powyższym przykładzie fillId="2" oznacza dosłownie „trzecie wypełnienie", czyli pierwsze, które ma sens. Wielu ludzi zakłada, że „jeśli nie używam żadnego wypełnienia, to pierwsze, które dodam, będzie numerem 0" — nie będzie.

**Pułapka 2: `numFmtId` poniżej 164 to formaty wbudowane.** Numery 0–163 mają znaczenie narzucone przez standard (0 = General, 2 = `0.00`, 4 = `#,##0.00`, 9 = `0%`, 10 = `0.00%`, 14 = `mm-dd-yy`...). **Formaty niestandardowe zaczynają się od 164** i tylko one pojawiają się w liście `<numFmts>`. Jeśli zobaczysz `numFmtId="164"`, musisz zajrzeć do `<numFmts>`. Jeśli zobaczysz `numFmtId="4"`, nie ma czego szukać — to znany format wbudowany.

### 2.8. Dwa światy tekstu: katalog (`sharedStrings`) kontra zapis na miejscu (`inlineStr`)

Wróćmy do obserwacji z sekcji 2.1. Plik utworzony przez openpyxl **nie miał części `xl/sharedStrings.xml`**. Dlaczego? I dlaczego plik utworzony przez Excela ją ma?

Format OOXML dopuszcza dwa sposoby zapisania tekstu w komórce:

**Sposób A — katalog tekstów (używa go Excel).**

```xml
<!-- xl/worksheets/sheet1.xml -->
<c r="A1" t="s"><v>0</v></c>
<c r="A2" t="s"><v>1</v></c>
<c r="A3" t="s"><v>1</v></c>
```

```xml
<!-- xl/sharedStrings.xml -->
<sst xmlns="..." count="3" uniqueCount="2">
  <si><t>Region</t></si>
  <si><t>Warszawa</t></si>
</sst>
```

Trzy komórki, dwa unikalne teksty. „Warszawa" pojawia się dwa razy, ale jest zapisane **raz**, a obie komórki odsyłają do tego samego indeksu `1` (drugi wpis — indeksy od zera!).

**Sposób B — tekst wprost w komórce (używa go openpyxl).**

```xml
<!-- xl/worksheets/sheet1.xml -->
<c r="A1" t="inlineStr"><is><t>Region</t></is></c>
<c r="A2" t="inlineStr"><is><t>Warszawa</t></is></c>
<c r="A3" t="inlineStr"><is><t>Warszawa</t></is></c>
```

Każda komórka nosi swój tekst. Nie ma części `sharedStrings.xml`. Plik jest samowystarczalny.

**Skąd to wiemy?** Nie z plotek — z kodu openpyxl. W `openpyxl/cell/_writer.py` funkcja `_set_attributes()` ustawia dla tekstu typ `"inlineStr"`, a `lxml_write_cell` / `etree_write_cell` zapisują tekst w elemencie `<is>`. Natomiast w `openpyxl/writer/excel.py` wiersz, który kiedyś zapisywał część `sharedStrings`, jest **zakomentowany**. To nie błąd w wersji 3.1.5 — to świadoma decyzja projektowa, obecna od bardzo dawna: budowanie katalogu tekstów wymaga trzymania w pamięci wszystkich unikalnych napisów i generuje narzut, a openpyxl obrał kurs na prostotę i szybkość.

**Co to znaczy dla Ciebie w praktyce? Pięć konkretnych konsekwencji:**

1. **Rozmiar pliku.** Jeśli Twoje dane to głównie powtarzalne teksty (kategoria, region, status), plik z openpyxl będzie **większy** niż ten sam plik z Excela, bo katalog tekstów daje kompresję przez deduplikację. Przy 1 000 000 wierszy z pięcioma powtarzającymi się wartościami różnica może być wyraźna. Odwrotnie: przy danych unikalnych (UUID, numery zamówień) różnicy prawie nie ma.

2. **Pamięć przy zapisie.** openpyxl nie musi trzymać katalogu tekstów w pamięci — to zaleta inline strings. Ale w trybie `write_only` i tak gromadzi sporo danych, więc nie traktuj tego jako lekarstwa na wszystko (moduł 19).

3. **Kompatybilność z innymi narzędziami.** Niektóre parsery OOXML zakładają, że tekst w skoroszycie jest zawsze w katalogu. Zdarzają się biblioteki, które czytają `t="s"` perfekcyjnie, a `t="inlineStr"` pomijają — i wtedy wczytają arkusz bez tekstów, bez żadnego błędu. Jeśli generujesz pliki dla innego systemu, **przetestuj to na jego odbiorcy**, zanim uznasz, że wszystko działa. To realne ryzyko integracyjne.

4. **Odczytywanie.** openpyxl czyta oba sposoby bez problemu. Ale zwróć uwagę na flagę `rich_text`: tekst zapisany inline może zawierać nie jeden element `<t>`, a kilka „przebiegów" (`<r>`) z różnym formatowaniem fragmentów. Domyślnie openpyxl scala to do zwykłego tekstu; z `rich_text=True` (nowość 3.1) oddaje obiekt `CellRichText`.

5. **Diagnostyka.** Jeśli piszesz narzędzie, które ma porównywać pliki różnego pochodzenia (a będziesz — moduł 20), musisz wiedzieć, że plik z Excela i plik z openpyxl o tej samej zawartości **będą się różnić strukturą XML**. Nie porównuj ich bajt w bajt. Porównuj *model*, nie *zapis*.

**Uwaga o spornych doniesieniach.** W sieci krążą raporty, że pewne bardzo rygorystyczne implementacje Excela (podobno wersja na macOS) reagują komunikatem o problemie z zawartością na pliki z inline strings, podczas gdy na Windows otwierają się bez protestu. Nie traktuj tego jako ustalonego faktu — źródła są niejednorodne, a problem może dotyczyć konkretnych wersji i konkretnych kombinacji. Traktuj to jako **listę rzeczy do sprawdzenia**: jeśli tworzysz pliki dla szerokiego grona odbiorców na różnych platformach, otwórz wygenerowany plik na kilku z nich, zanim ogłosisz sukces. To i tak dobra praktyka (moduł 18).

### 2.9. Rysunki, wykresy, obrazy: `xl/drawings/`, `xl/charts/`, `xl/media/`

To trójka, która pokazuje, że komórka to nie wszystko, co żyje w arkuszu.

- **`xl/media/image1.png`** — binarny plik obrazu. Tak, plik `.xlsx` może zawierać prawdziwe PNG-i i JPEG-i. To jedna z niewielu części binarnych w pakiecie.
- **`xl/drawings/drawing1.xml`** — warstwa rysunkowa. To tutaj jest **pozycja** każdego obrazu: „obraz `rId1` zaczyna się w komórce B2 i rozciąga na dwie komórki w prawo i osiem w dół". Obraz nie jest „w komórce" — jest na warstwie, która **pływa nad siatką** i jest zakotwiczona do komórek. Dlatego sortowanie i filtrowanie nie przesuwa obrazów (moduł 16).
- **`xl/charts/chart1.xml`** — definicja wykresu. Są tu serie danych, zakresy odwołań, typ, osie, tytuł. Wykres **nie jest obrazkiem** — jest przepisem, który Excel odtwarza przy każdym otwarciu (moduł 15).

Powiązania między nimi są znów przez relacje: arkusz → `drawing1.xml` → `image1.png` i `chart1.xml`. Trzy poziomy spinek, żeby przejść od kartki do obrazka na ekranie.

**Konsekwencja dla modułu 18:** skoro obraz to plik w `xl/media/` plus opis w `xl/drawings/`, to przy zapisie openpyxl musi odtworzyć **oba**. Obrazy są jednym z elementów, które openpyxl obsługuje — ale wczytywanie i zapisywanie wykresów działa inaczej: openpyxl **odbudowuje wykres z własnego modelu**. Jeśli model nie obejmuje jakiejś ozdoby, ozdoba zniknie.

### 2.10. Metadane: `docProps/`

Trzy pliki, o których łatwo zapomnieć, a które bywają istotne:

- **`docProps/core.xml`** — autor, tytuł, temat, data utworzenia i modyfikacji, ostatni autor. Tu trafiają dane z `wb.properties` (moduł 04). W audycie to złoto: „kto i kiedy wygenerował ten raport".
- **`docProps/app.xml`** — metadane aplikacji: nazwa programu, wersja, nazwy arkuszy. I tu ciekawostka: openpyxl **generuje ten plik od nowa**, tworząc świeży obiekt `ExtendedProperties()`, zamiast kopiować część z pliku źródłowego. Wniosek praktyczny: nie zakładaj, że metadane aplikacji w pliku źródłowym przetrwają zapis w niezmienionej postaci. Sprawdź to na własnym pliku — przykład 1 tego modułu pokaże Ci, co dokładnie tam jest.
- **`docProps/custom.xml`** — własne właściwości dokumentu. W openpyxl 3.1 możesz je tworzyć programowo. Świetne miejsce na „wersja szablonu", „hash danych wejściowych", „identyfikator uruchomienia" (moduły 18 i 26).

### 2.11. Części ryzykowne: co openpyxl gubi przy zapisie

I wreszcie lista, dla której powstał cały ten moduł. Te części pakietu istnieją w prawdziwych plikach Excela i **nie mają odpowiednika w modelu openpyxl**. Przy zapisie, który jest przepisaniem modelu, nie mają z czego powstać.

- **`xl/vbaProject.bin`** — projekt makr. Zachowany **tylko** z `keep_vba=True`. I nawet wtedy: zachowany, nie edytowalny.
- **`xl/slicers/`, `xl/slicerCaches/`, `xl/timelines/`** — slicery i osie czasu.
- **`xl/connections.xml`, `xl/queryTables/`, `xl/model/`, `customXml/`** — Power Query, Power Pivot, model danych.
- **`xl/pivotCache/`** — pamięć podręczna tabel przestawnych. Uwaga na niuans: openpyxl potrafi **odczytać i zapisać z powrotem** tabelę przestawną, którą wczytał. Ale nie potrafi jej **utworzyć od zera** ani **odświeżyć** jej danych. Tabela przestawna to trzy współzależne zapisy: pamięć podręczna z danymi, definicja tabeli i literalna siatka wyników — wszystkie trzy muszą się zgadzać. Openpyxl modeluje to na tyle, by nie zepsuć, ale nie na tyle, by zbudować.
- **`xl/ctrlProps/`, `xl/activeX/`, `customUI/`** — właściwości formantów, kontrolki ActiveX, własna wstążka. Ciekawostka z kodu openpyxl: te części są kopiowane z pliku źródłowego **tylko wtedy, gdy wczytasz plik z `keep_vba=True`** — kopiowanie odbywa się w ścieżce „merge VBA". Bez tej flagi przepadają.
- **Bloki `<extLst>`** — i to jest temat, który zasługuje na osobne omówienie.

### 2.12. Części rozszerzeń (`extLst`) — dlaczego sparkline'y giną

Format OOXML został zaprojektowany z myślą o rozbudowie. Specyfikacja przewidziała miejsce na „różne rzeczy wymyślone później": blok `<extLst>` (extension list), który można wstawić w wielu miejscach pakietu — w skoroszycie, w arkuszu, w stylach.

Mechanizm: każdy wpis `<ext>` ma atrybut `uri` — unikalny identyfikator typu rozszerzenia. Parser, który zna dany `uri`, przetwarza zawartość. Parser, który go nie zna, **powinien go zostawić w spokoju i zachować przy zapisie**. Taka jest teoria.

**Praktyka openpyxl:** openpyxl nie modeluje bloków rozszerzeń jako danych do zachowania. Przy wczytaniu arkusza trafia na `<extLst>`, sprawdza `uri` i wypisuje na standardowe wyjście błędu ostrzeżenie w rodzaju:

```text
UserWarning: Sparkline Group extension is not supported and will be removed
```

Ostrzeżenie mówi dokładnie to, co się stanie: przy zapisie tego nie będzie.

Oto jak wygląda prawdziwy sparkline w pliku Excela — grupa sparkline siedzi wewnątrz `<extLst>`, w przestrzeni nazw `x14` (rozszerzenie Excela 2010), a identyfikator URI to `{05C60535-1F16-4fd2-B633-F4F36F0B64E0}`:

```xml
<extLst>
  <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
       xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
    <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
      <x14:sparklineGroup type="line">
        <x14:sparklines>
          <x14:sparkline>
            <xm:f>Arkusz!B2:E2</xm:f>
            <xm:sqref>F2</xm:sqref>
          </x14:sparkline>
        </x14:sparklines>
      </x14:sparklineGroup>
    </x14:sparklineGroups>
  </ext>
</extLst>
```

Czytasz to tak: „w komórce F2 narysuj miniwykres liniowy na podstawie danych z B2:E2". Bez bloku `extLst` komórka F2 jest pusta — i taka będzie po przepuszczeniu pliku przez openpyxl.

**Ale uwaga na subtelność, która wielu ludzi myli.** openpyxl potrafi ostrzec o kilku różnych rozszerzeniach. Lista znanych typów w kodzie źródłowym zawiera między innymi: `Sparkline Group`, `Slicer List`, `Timeline Ref`, `Protected Range`, `Ignored Error`, `Web Extension` — ale też **`Conditional Formatting`** i **`Data Validation`**. Czy to znaczy, że openpyxl gubi formatowanie warunkowe i walidację danych?

Nie. I to jest ważne rozróżnienie. openpyxl **potrafi** tworzyć i odczytywać podstawowe reguły formatowania warunkowego oraz walidację danych — to zwykłe elementy arkusza (`<conditionalFormatting>`, `<dataValidations>`), a nie rozszerzenia. Ostrzeżenie dotyczy **rozszerzonych wariantów** tych funkcji, zapisanych w `x14:...` obok zwykłych. Oznacza to, że plik może zawierać **dwie kopie** tych samych reguł: zwykłą (którą openpyxl rozumie) i rozszerzoną (którą gubi). To się nie zdarza w każdym pliku, ale się zdarza — szczególnie gdy plik był edytowany w nowych wersjach Excela.

**Konsekwencja praktyczna:** gdy zobaczysz w konsoli ostrzeżenie „... extension is not supported and will be removed", nie ignoruj go i nie wyciszaj globalnie (`warnings.filterwarnings("ignore")` to antywzorzec — ukryje też ostrzeżenia, które naprawdę mają znaczenie). **Odczytaj, którego rozszerzenia dotyczy, i oceń, czy w Twoim pliku coś od niego zależy.**

### 2.13. Dwie arytmetyki w jednym pliku — przygotowanie do modułu 03

Podsumujmy to, co będzie wracało przez cały kurs:

| Rzecz | System |
|---|---|
| Adres komórki (`A1`, `C7`) | litery kolumn od A, wiersze od 1 |
| `<row r="...">` | od 1 |
| `<col min="..." max="...">` | od 1 |
| Wymiar arkusza (`max_row`, `max_column`) | od 1 |
| Indeks stylu w komórce (`s="..."`) | **od 0** |
| Indeks wpisu w `cellXfs`, `fonts`, `fills`, `borders` | **od 0** |
| Indeks w katalogu tekstów (`<v>` przy `t="s"`) | **od 0** |
| `dxfId` w regule formatowania warunkowego | **od 0** |
| `numFmtId` — wbudowane (0–163) i niestandardowe (od 164) | liczby, nie indeksy listy |

W tym samym pliku, w tym samym fragmencie XML, możesz jednocześnie mieć `C1` (1-indeksowane), `s="1"` (0-indeksowane) i `<v>0</v>` (0-indeksowane). Trzy różne „jedynki" w jednej linijce. To najczęstsze źródło błędów w narzędziach pracujących na OOXML — i dlatego module 03 poświęcimy temu osobną uwagę.

### 2.14. Dlaczego to wszystko ma znaczenie dla modułu 18

Moduł 18 to „modyfikacja istniejących plików bez utraty zawartości". Bez dzisiejszego modułu byłby dla Ciebie zbiorem niepowiązanych ostrzeżeń. Z dzisiejszą wiedzą jest oczywistością:

> openpyxl **wczytuje pakiet do modelu** i **zapisuje model do nowego pakietu**. Wszystko, co ma odpowiednik w modelu — przetrwa. Wszystko, co nie ma — nie przetrwa, bo nie było czego przepisać. Nie jest to usterka. Jest to bezpośrednia konsekwencja tego, jak zbudowany jest pakiet i jak zbudowany jest openpyxl.

Dlatego zanim kiedykolwiek pozwolisz openpyxl zapisać cudzy plik, zajrzysz do jego wnętrza i sprawdzisz, czy nie ma tam `slicer*`, `timeline*`, `vbaProject.bin` albo `extLst` ze sparkline'ami. Dziś nauczysz się to robić.

## 3. Przykłady krok po kroku

Wszystkie przykłady zapisują do `output/` i czytają z `data/`, zgodnie z konwencją modułu 00.

### Przykład 1 — Inwentarz pakietu

Program, który bierze dowolny plik `.xlsx` i wypisuje wszystkie jego części, pogrupowane według roli, z rozmiarami i stopniem kompresji. To pierwsze narzędzie, jakie powinieneś uruchomić na każdym nieznanym pliku.

```python
"""Wypisuje inwentarz części dowolnego pliku .xlsx.

Uruchom:  python examples/02_inwentarz.py [plik1.xlsx plik2.xlsx ...]
Bez argumentów: generuje plik demonstracyjny.
"""

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

# Przypisanie części do roli. Kolejność ma znaczenie: sprawdzamy od góry,
# bo niektóre wzorce się nakładają (np. "xl/worksheets/" vs "xl/workbook.xml").
KATEGORIE: dict[str, tuple[str, ...]] = {
    "Spis i powiązania": ("[Content_Types].xml", "_rels/"),
    "Metadane": ("docProps/",),
    "Szkielet skoroszytu": ("xl/workbook.xml", "xl/_rels/"),
    "Arkusze (kartki)": ("xl/worksheets/",),
    "Style i motyw": ("xl/styles.xml", "xl/theme/"),
    "Katalog tekstów": ("xl/sharedStrings.xml",),
    "Rysunki i wykresy": ("xl/drawings/", "xl/charts/", "xl/media/"),
    "Tabele i filtry": ("xl/tables/",),
    "Komentarze": ("xl/comments", "xl/threadedComments/", "xl/persons/"),
    "Kolejność obliczeń": ("xl/calcChain.xml",),
    "Odwołania zewnętrzne": ("xl/externalLinks/",),
    "Tabele przestawne": ("xl/pivotTables/", "xl/pivotCache/"),
    "Makra i formanty": (
        "xl/vbaProject.bin", "xl/ctrlProps/", "xl/activeX/", "customUI/",
    ),
    "Slicery i osie czasu": ("xl/slicers/", "xl/slicerCaches/", "xl/timelines/"),
    "Power Query / model danych": (
        "xl/connections.xml", "xl/queryTables/", "xl/model/", "customXml/",
    ),
}


def zaklasyfikuj(nazwa: str) -> str:
    """Przypisuje nazwę części do roli. Zwraca 'NIEZNANE' gdy nic nie pasuje."""
    for kategoria, wzorce in KATEGORIE.items():
        for wzorzec in wzorce:
            if nazwa == wzorzec or nazwa.startswith(wzorzec) or wzorzec in nazwa:
                return kategoria
    return "NIEZNANE (openpyxl prawdopodobnie tego nie modeluje)"


def pokaz(sciezka: Path) -> None:
    print("=" * 74)
    print(f"PLIK: {sciezka}")
    print("=" * 74)

    # Test deterministyczny z modułu 01: .xlsx JEST ZIP-em, .xls nie jest.
    if not zipfile.is_zipfile(sciezka):
        print("  To NIE jest archiwum ZIP -> to nie .xlsx (może .xls albo .xlsb).\n")
        return

    with zipfile.ZipFile(sciezka) as archiwum:
        infos = archiwum.infolist()   # ZipInfo: nazwa, rozmiar, rozmiar po kompresji
        print(f"  Części: {len(infos)}   "
              f"rozmiar na dysku: {sciezka.stat().st_size / 1024:.1f} kB\n")

        grupy: dict[str, list[tuple[str, int, int]]] = {}
        for info in infos:
            grupy.setdefault(zaklasyfikuj(info.filename), []).append(
                (info.filename, info.file_size, info.compress_size)
            )

        for kategoria in sorted(grupy):
            wpisy = grupy[kategoria]
            rozpakowane = sum(w[1] for w in wpisy)
            print(f"  [{kategoria}]  ({len(wpisy)} części, "
                  f"{rozpakowane / 1024:.1f} kB po rozpakowaniu)")
            for nazwa, rozp, skompr in sorted(wpisy):
                procent = (skompr / rozp * 100) if rozp else 0.0
                print(f"      {nazwa:<50} {rozp:>9} B -> {skompr:>8} B ({procent:4.1f}%)")
            print()

        # Kluczowe pytanie dla zrozumienia openpyxl vs Excel:
        nazwy = archiwum.namelist()
        ma_katalog = "xl/sharedStrings.xml" in nazwy
        print(f"  xl/sharedStrings.xml obecne: {'TAK' if ma_katalog else 'NIE'}")
        if ma_katalog:
            print("     -> teksty są prawdopodobnie w katalogu (t=\"s\"). "
                  "Tak robi Excel i xlsxwriter.")
        else:
            print("     -> teksty są prawdopodobnie zapisane w komórkach (t=\"inlineStr\").")
            print("        Tak domyślnie robi openpyxl.")

        # Szybkie ostrzeżenia o elementach, które openpyxl gubi.
        ryzyka = {
            "vbaProject.bin": "makra VBA (przetrwają tylko z keep_vba=True)",
            "slicer": "slicery",
            "timeline": "osie czasu",
            "pivotCache": "pamięć podręczna tabeli przestawnej",
            "activeX": "formanty ActiveX",
            "ctrlProps": "właściwości formantów",
        }
        znalezione = []
        for nazwa in nazwy:
            for fragment, opis in ryzyka.items():
                if fragment in nazwa:
                    znalezione.append(f"{nazwa} -> {opis}")
        if znalezione:
            print("\n  UWAGA - części, które openpyxl gubi przy zapisie:")
            for wpis in sorted(set(znalezione)):
                print(f"     ! {wpis}")
        print()


def wygeneruj_demo() -> Path:
    """Plik demonstracyjny, gdy użytkownik nie podał żadnego."""
    from openpyxl import Workbook

    OUTPUT.mkdir(parents=True, exist_ok=True)
    sciezka = OUTPUT / "02_inwentarz_demo.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    for i, wartosc in enumerate(["Region", "Warszawa", "Krakow", "Gdansk"], start=1):
        ws.cell(row=i, column=1, value=wartosc)
    wb.save(sciezka)
    print(f"(nie podano pliku — wygenerowałem {sciezka})\n")
    return sciezka


def glowny() -> None:
    pliki = [Path(p) for p in sys.argv[1:]]
    if not pliki:
        pliki = [wygeneruj_demo()]
    for plik in pliki:
        if plik.exists():
            pokaz(plik)
        else:
            print(f"Pomijam (nie ma pliku): {plik}\n")


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** `zipfile.ZipFile` czyta tylko **katalog centralny** archiwum (spis części na końcu pliku), a nie treść części. Dlatego inwentarz jest natychmiastowy nawet dla pliku 200 MB. Same dane części nie są wczytywane.

**Co trafi do pliku:** nic, poza ścieżką „brak argumentu", która tworzy `output/02_inwentarz_demo.xlsx`.

**Oczekiwany fragment wyniku dla pliku z openpyxl:**

```text
  [Katalog tekstów]  (0 części, 0.0 kB po rozpakowaniu)

  xl/sharedStrings.xml obecne: NIE
     -> teksty są prawdopodobnie zapisane w komórkach (t="inlineStr").
        Tak domyślnie robi openpyxl.
```

Zwróć uwagę, że kategoria „Katalog tekstów" pojawia się pusta — to celowy zabieg dydaktyczny. Pusta kategoria mówi więcej niż brak kategorii.

### Przykład 2 — Rentgen: czytelny wydruk XML z wnętrza pakietu

Teraz, gdy umiesz już wypisać części, zajrzyjmy do ich wnętrza. Ten program pokazuje cztery najważniejsze strukturalne części w czytelnej, wciętej formie — plus arkusz.

```python
"""Pokazuje XML najważniejszych części pakietu, z wcięciami.

Uruchom:  python examples/02_rentgen.py output/02_inwentarz_demo.xlsx
"""

import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.dom import minidom

CZESCI_STRUKTURALNE = [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
]


def sformatuj(dane: bytes, limit: int = 2200) -> str:
    """Zamienia jednoliniowy XML w czytelne, wcięte drzewo.

    minidom jest wygodny do PODGLĄDANIA (nie do parsowania danych):
    umie zwrócić drzewo z wcięciami. Uwaga na limit - dla kilkusetmegabajtowych
    części ten sposób zjada pamięć. Do danych używaj ElementTree (patrz przykład 3).
    """
    dom = minidom.parseString(dane)
    surowy = dom.toprettyxml(indent="    ")
    # minidom dokłada sporo pustych linii - usuwamy je.
    linie = [linia for linia in surowy.splitlines() if linia.strip()]
    tekst = "\n".join(linie)
    if len(tekst) > limit:
        tekst = tekst[:limit] + "\n    <!-- ... obcięto ... -->"
    return tekst


def bez_ns(tag: str) -> str:
    """Z 'http://...}worksheet' robi 'worksheet' (usuwa przestrzeń nazw)."""
    return tag.rsplit("}", 1)[-1]


def pokaz_czesc(archiwum: zipfile.ZipFile, nazwa: str) -> None:
    try:
        dane = archiwum.read(nazwa)
    except KeyError:
        print(f"\n--- {nazwa}: BRAK TEJ CZĘŚCI ---")
        return
    print(f"\n--- {nazwa}  ({len(dane)} B) ---")
    print(sformatuj(dane))


def podsumuj_arkusz(dane: bytes, nazwa: str) -> None:
    """Dla arkusza nie drukujemy całości - wypisujemy strukturę i komórki."""
    print(f"\n--- {nazwa}  ({len(dane)} B) ---")

    root = ET.fromstring(dane)
    print("  Elementy najwyższego poziomu:")
    for dziecko in root:
        atrybuty = " ".join(
            f'{bez_ns(k)}="{v}"' for k, v in dziecko.attrib.items()
        )
        sufiks = f"  {atrybuty}" if atrybuty else ""
        print(f"      <{bez_ns(dziecko.tag)}{sufiks}>")

    print("\n  Komórki w surowej postaci:")

    def tekst_komorki(c: ET.Element) -> str:
        """Wyciąga tekst z <is><t> (inlineStr) albo wartość z <v>."""
        inline = next((x for x in c if bez_ns(x.tag) == "is"), None)
        if inline is not None:
            return "".join(
                (t.text or "") for t in inline.iter() if bez_ns(t.tag) == "t"
            )
        wartosc = next((x for x in c if bez_ns(x.tag) == "v"), None)
        if wartosc is not None:
            return wartosc.text or ""
        formula = next((x for x in c if bez_ns(x.tag) == "f"), None)
        if formula is not None:
            return f"=<formuła> {formula.text or ''}"
        return ""

    licznik = 0
    for c in root.iter():
        if bez_ns(c.tag) != "c":
            continue
        atrybuty = " ".join(
            f'{bez_ns(k)}="{v}"' for k, v in c.attrib.items()
        )
        typ = c.get("t", "n")            # DOMYŚLNIE 'n' gdy brak atrybutu!
        indeks_stylu = c.get("s", "0")   # DOMYŚLNIE '0' gdy brak atrybutu!
        print(f"      <c {atrybuty}>")
        print(f"          t={typ!r}  s={indeks_stylu!r}  zawartość={tekst_komorki(c)!r}")
        licznik += 1
        if licznik >= 10:
            print("      ... (dalej pominięto)")
            break


def glowny() -> None:
    if len(sys.argv) < 2:
        print("Podaj ścieżkę do pliku .xlsx")
        print("np.:  python examples/02_rentgen.py output/02_inwentarz_demo.xlsx")
        return

    sciezka = Path(sys.argv[1])
    if not zipfile.is_zipfile(sciezka):
        print(f"{sciezka} nie jest archiwum ZIP.")
        return

    with zipfile.ZipFile(sciezka) as archiwum:
        for nazwa in CZESCI_STRUKTURALNE:
            pokaz_czesc(archiwum, nazwa)

        for nazwa in archiwum.namelist():
            if nazwa.startswith("xl/worksheets/") and nazwa.endswith(".xml"):
                podsumuj_arkusz(archiwum.read(nazwa), nazwa)


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** `minidom.parseString` buduje pełne drzewo DOM w pamięci — wygodne, ale pamięciożerne. `ET.fromstring` też buduje drzewo, ale lżejsze i szybsze. Dla pojedynczych części demonstracyjnych to nie ma znaczenia; dla pliku z milionem wierszy `minidom` jest złym wyborem i dlatego w kodzie danych używamy `ElementTree`.

**Co trafi do pliku:** nic. To narzędzie wyłącznie diagnostyczne.

**Oczekiwany fragment (arkusz z openpyxl):**

```text
--- xl/worksheets/sheet1.xml  (1024 B) ---
  Elementy najwyższego poziomu:
      <sheetPr>
      <dimension  ref="A1:A4">
      <sheetViews>
      <sheetFormatPr  defaultRowHeight="15">
      <sheetData>
      ...

  Komórki w surowej postaci:
      <c r="A1" t="inlineStr">
          t='inlineStr'  s='0'  zawartość='Region'
      <c r="A2" t="inlineStr">
          t='inlineStr'  s='0'  zawartość='Warszawa'
```

I to jest ten moment, w którym teoria z sekcji 2.8 staje się faktem widzianym na własne oczy: **`t='inlineStr'`** i brak jakiegokolwiek katalogu tekstów.

### Przykład 3 — Tropiciel stylu: od `s="2"` do konkretnej czcionki

Ten program robi dokładnie to, co robiłeś ręcznie w sekcji 2.7 — ale automatycznie i na prawdziwym pliku. To najbardziej pouczający przykład tego modułu.

```python
"""Śledzi łańcuch: komórka -> s="N" -> cellXfs[N] -> fontId/fillId/... -> wpis.

Uruchom:  python examples/02_tropiciel_stylu.py [plik.xlsx]
Bez argumentu tworzy plik ze stylem, żeby było co tropić.
"""

import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

# Wybrane WBUDOWANE formaty liczb (numFmtId < 164 to formaty standardowe).
# Wartości od 164 wzwyż są NIESTANDARDOWE i muszą być zadeklarowane w <numFmts>.
WBUDOWANE_FORMATY = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}


def bez_ns(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def dziecko(el: ET.Element, nazwa: str):
    """Pierwsze dziecko o danej nazwie lokalnej (lub None).

    UWAGA na pułapkę ElementTree: bool pustego elementu <b/> to False!
    Dlatego ZAWSZE porównuj z "is not None", nigdy "if dziecko(...):".
    """
    for d in el:
        if bez_ns(d.tag) == nazwa:
            return d
    return None


def dzieci(el: ET.Element, nazwa: str) -> list[ET.Element]:
    return [d for d in el if bez_ns(d.tag) == nazwa]


def opisz_font(font: ET.Element) -> str:
    czesci: list[str] = []
    # Celowo używamy "is not None" - patrz komentarz w funkcji dziecko().
    if dziecko(font, "b") is not None:
        czesci.append("pogrubiona")
    if dziecko(font, "i") is not None:
        czesci.append("kursywa")
    if dziecko(font, "u") is not None:
        czesci.append("podkreślona")
    nazwa = dziecko(font, "name")
    if nazwa is not None:
        czesci.append(f"czcionka={nazwa.get('val')!r}")
    rozmiar = dziecko(font, "sz")
    if rozmiar is not None:
        czesci.append(f"rozmiar={rozmiar.get('val')}")
    kolor = dziecko(font, "color")
    if kolor is not None:
        opis = []
        if kolor.get("rgb"):
            opis.append(f"rgb={kolor.get('rgb')}")
        if kolor.get("theme"):
            opis.append(f"theme={kolor.get('theme')}")
        if kolor.get("indexed"):
            opis.append(f"indexed={kolor.get('indexed')}")
        czesci.append("kolor(" + ", ".join(opis) + ")" if opis else "kolor(domyślny)")
    return ", ".join(czesci) if czesci else "(domyślna czcionka)"


def opisz_wypelnienie(fill: ET.Element) -> str:
    wzor = dziecko(fill, "patternFill")
    if wzor is None:
        return "(brak wypełnienia)"
    typ = wzor.get("patternType", "none")
    if typ == "none":
        return "(brak wypełnienia)"
    fg = dziecko(wzor, "fgColor")
    bg = dziecko(wzor, "bgColor")
    elementy = [f"typ={typ!r}"]
    if fg is not None:
        elementy.append(f"fgColor={fg.get('rgb') or fg.get('indexed') or fg.get('theme')}")
    if bg is not None:
        elementy.append(f"bgColor={bg.get('rgb') or bg.get('indexed')}")
    return ", ".join(elementy)


def opisz_obramowanie(border: ET.Element) -> str:
    boki = []
    for bok in ("left", "right", "top", "bottom"):
        el = dziecko(border, bok)
        if el is not None and el.get("style"):
            boki.append(f"{bok}={el.get('style')}")
    return ", ".join(boki) if boki else "(brak obramowania)"


def opisz_format(numfmt_id: int, wlasne: dict[int, str]) -> str:
    if numfmt_id in wlasne:
        return f"{wlasne[numfmt_id]!r} (niestandardowy, id={numfmt_id})"
    if numfmt_id in WBUDOWANE_FORMATY:
        return f"{WBUDOWANE_FORMATY[numfmt_id]!r} (wbudowany, id={numfmt_id})"
    return f"? (nieznany id={numfmt_id})"


def przeanalizuj(sciezka: Path) -> None:
    print("=" * 74)
    print(f"PLIK: {sciezka}")
    print("=" * 74)

    with zipfile.ZipFile(sciezka) as archiwum:
        if "xl/styles.xml" not in archiwum.namelist():
            print("Ten plik nie ma części xl/styles.xml (nie ma zdefiniowanych stylów).")
            return

        style = ET.fromstring(archiwum.read("xl/styles.xml"))

        # 1. Magazyny - każdy wpis ma numer wynikający z KOLEJNOŚCI, liczonej od 0.
        czcionki = dzieci(dziecko(style, "fonts") or ET.Element("x"), "font")
        wypelnienia = dzieci(dziecko(style, "fills") or ET.Element("x"), "fill")
        obramowania = dzieci(dziecko(style, "borders") or ET.Element("x"), "border")
        xfs = dzieci(dziecko(style, "cellXfs") or ET.Element("x"), "xf")

        wlasne_formaty: dict[int, str] = {}
        numfmts = dziecko(style, "numFmts")
        if numfmts is not None:
            for nf in dzieci(numfmts, "numFmt"):
                wlasne_formaty[int(nf.get("numFmtId"))] = nf.get("formatCode")

        print(f"\nMAGAZYNY W xl/styles.xml:")
        print(f"   <fonts>    -> {len(czcionki)} wpisów (indeksy 0..{len(czcionki) - 1})")
        print(f"   <fills>    -> {len(wypelnienia)} wpisów (indeksy 0..{len(wypelnienia) - 1})")
        print(f"      ^ uwaga: wpisy 0 i 1 są ZAREZERWOWANE przez Excel")
        print(f"   <borders>  -> {len(obramowania)} wpisów")
        print(f"   <cellXfs>  -> {len(xfs)} wpisów <- na te odsyła atrybut s= w komórce")

        # 2. Surowe magazyny czcionek - żeby zobaczyć, co jest pod jakim indeksem.
        print(f"\nWSZYSTKIE CZCIONKI (indeks -> opis):")
        for i, font in enumerate(czcionki):
            print(f"   fonts[{i}] = {opisz_font(font)}")

        # 3. Komórki z arkusza i rozwiązanie łańcucha.
        for nazwa in archiwum.namelist():
            if not (nazwa.startswith("xl/worksheets/sheet") and nazwa.endswith(".xml")):
                continue
            print(f"\nKOMPÓRKI Z {nazwa}:")
            arkusz = ET.fromstring(archiwum.read(nazwa))
            licznik = 0
            for c in arkusz.iter():
                if bez_ns(c.tag) != "c":
                    continue
                adres = c.get("r", "?")
                indeks_stylu = int(c.get("s", "0"))   # DOMYŚLNIE 0 gdy brak atrybutu

                if indeks_stylu >= len(xfs):
                    print(f"   {adres}: s={indeks_stylu} - POZA ZAKRESEM cellXfs!")
                    continue

                xf = xfs[indeks_stylu]
                font_id = int(xf.get("fontId", "0"))
                fill_id = int(xf.get("fillId", "0"))
                border_id = int(xf.get("borderId", "0"))
                numfmt_id = int(xf.get("numFmtId", "0"))

                print(f"\n   {adres}:   s={indeks_stylu}")
                print(f"      cellXfs[{indeks_stylu}] -> "
                      f"fontId={font_id}, fillId={fill_id}, "
                      f"borderId={border_id}, numFmtId={numfmt_id}")
                if font_id < len(czcionki):
                    print(f"      fonts[{font_id}]    -> {opisz_font(czcionki[font_id])}")
                if fill_id < len(wypelnienia):
                    print(f"      fills[{fill_id}]    -> {opisz_wypelnienie(wypelnienia[fill_id])}")
                if border_id < len(obramowania):
                    print(f"      borders[{border_id}] -> {opisz_obramowanie(obramowania[border_id])}")
                print(f"      format liczby   -> {opisz_format(numfmt_id, wlasne_formaty)}")

                licznik += 1
                if licznik >= 5:
                    print("\n   ... (dalej pominięto)")
                    break


def wygeneruj_demo() -> Path:
    """Plik ze zróżnicowanymi stylami, żeby łańcuch był widoczny."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    OUTPUT.mkdir(parents=True, exist_ok=True)
    sciezka = OUTPUT / "02_tropiciel_demo.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Style"

    # Komórka A1: domyślna (s prawdopodobnie nie pojawi się wcale).
    ws["A1"] = "Domyślna"

    # Komórka A2: pogrubiona czcionka (nowy wpis w <fonts>).
    ws["A2"] = "Pogrubiona"
    ws["A2"].font = Font(bold=True)

    # Komórka A3: pełny zestaw - czcionka, wypełnienie, obramowanie, format.
    ws["A3"] = 1234.5
    ws["A3"].font = Font(name="Arial", size=14, bold=True, italic=True)
    ws["A3"].fill = PatternFill("solid", fgColor="FF1F4E78")
    ws["A3"].border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    ws["A3"].alignment = Alignment(horizontal="center")
    ws["A3"].number_format = '#,##0.00 "zł"'

    wb.save(sciezka)
    print(f"(nie podano pliku — wygenerowałem {sciezka})\n")
    return sciezka


def glowny() -> None:
    if len(sys.argv) > 1:
        sciezka = Path(sys.argv[1])
        if not sciezka.exists():
            print(f"Nie ma takiego pliku: {sciezka}")
            return
    else:
        sciezka = wygeneruj_demo()
    przeanalizuj(sciezka)


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** `ET.fromstring` buduje drzewo `styles.xml` (mały plik, nawet dla rozbudowanych stylów) oraz drzewo arkusza. Wszystkie odwołania rozwiązujemy przez listy Pythona — dokładnie tak, jakby `cellXfs` było zwykłą listą, a `s="2"` indeksem.

**Co trafi do pliku:** `output/02_tropiciel_demo.xlsx` ze zróżnicowanymi stylami.

**Oczekiwany fragment wyniku:**

```text
MAGAZYNY W xl/styles.xml:
   <fonts>    -> 3 wpisów (indeksy 0..2)
   <fills>    -> 3 wpisów (indeksy 0..2)
      ^ uwaga: wpisy 0 i 1 są ZAREZERWOWANE przez Excel
   <borders>  -> 2 wpisów
   <cellXfs>  -> 3 wpisów <- na te odsyła atrybut s= w komórce

WSZYSTKIE CZCIONKI (indeks -> opis):
   fonts[0] = czcionka='Calibri', rozmiar=11
   fonts[1] = pogrubiona, czcionka='Calibri', rozmiar=11
   fonts[2] = pogrubiona, kursywa, czcionka='Arial', rozmiar=14

KOMPÓRKI Z xl/worksheets/sheet1.xml:

   A3:   s=2
      cellXfs[2] -> fontId=2, fillId=2, borderId=1, numFmtId=164
      fonts[2]    -> pogrubiona, kursywa, czcionka='Arial', rozmiar=14
      fills[2]    -> typ='solid', fgColor=FF1F4E78, bgColor=64
      borders[1] -> left=thin, right=thin, top=thin, bottom=thin
      format liczby   -> '#,##0.00 "zł"' (niestandardowy, id=164)
```

To jest dokładnie ten łańcuch, o którym mówiliśmy w teorii: `s=2` → `cellXfs[2]` → `fontId=2` → `fonts[2]`. I zwróć uwagę: **format niestandardowy dostał id 164** — pierwszą wolną liczbę powyżej wbudowanych. Zgadza się z tym, co pisaliśmy w sekcji 2.7.

### Przykład 4 — Dwa światy tekstu: `inlineStr` kontra katalog

Ten program buduje **ręcznie** dwa pliki `.xlsx` o identycznej zawartości — jeden z katalogiem tekstów (jak Excel), drugi z tekstami w komórkach (jak openpyxl) — a potem je rozbiera i pokazuje różnicę. Budując te pliki ręcznie, dowiesz się więcej o formacie niż z dziesięciu artykułów.

```python
"""Buduje dwa .xlsx ręcznie (bez openpyxl) i porównuje sposób zapisu tekstu.

Uruchom:  python examples/02_dwa_swiaty_textu.py
Efekt:    output/02_katalog.xlsx  i  output/02_inline.xlsx
"""

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
GŁÓWNY = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# ----------------------------------------------------------------------
# Składniki pakietu. Zmieniamy tylko dwie rzeczy między wariantami:
# sposób zapisu tekstu w arkuszu i obecność części sharedStrings.
# ----------------------------------------------------------------------

CONTENT_TYPES_KATALOG = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>'''

CONTENT_TYPES_INLINE = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''

ROOT_RELS = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKGREL}">
  <Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WORKBOOK = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{GŁÓWNY}" xmlns:r="{REL}">
  <sheets><sheet name="Dane" sheetId="1" r:id="rId1"/></sheets>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>'''

# Wariant z katalogiem: relacja typu sharedStrings do części katalogu.
WB_RELS_KATALOG = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKGREL}">
  <Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''

WB_RELS_INLINE = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKGREL}">
  <Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''

# Arkusz, wariant A: t="s" + indeksy do katalogu. "Warszawa" zapisana RAZ.
ARKUSZ_KATALOG = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{GŁÓWNY}">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>120000</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>98500</v></c></row>
    <row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>99000</v></c></row>
    <row r="4"><c r="A4" t="s"><v>1</v></c><c r="B4"><v>101200</v></c></row>
  </sheetData>
</worksheet>'''

# Arkusz, wariant B: tekst wprost w komórce. "Warszawa" zapisana TRZY RAZY.
ARKUSZ_INLINE = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{GŁÓWNY}">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1"><v>120000</v></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Warszawa</t></is></c><c r="B2"><v>98500</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Warszawa</t></is></c><c r="B3"><v>99000</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Warszawa</t></is></c><c r="B4"><v>101200</v></c></row>
  </sheetData>
</worksheet>'''

KATALOG = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="{GŁÓWNY}" count="4" uniqueCount="2">
  <si><t>Region</t></si>
  <si><t>Warszawa</t></si>
</sst>'''


def zbuduj(sciezka: Path, wariant: str) -> None:
    """Zapisuje minimalny pakiet .xlsx. wariant: 'katalog' albo 'inline'."""
    if wariant == "katalog":
        czesci = {
            "[Content_Types].xml": CONTENT_TYPES_KATALOG,
            "_rels/.rels": ROOT_RELS,
            "xl/workbook.xml": WORKBOOK,
            "xl/_rels/workbook.xml.rels": WB_RELS_KATALOG,
            "xl/worksheets/sheet1.xml": ARKUSZ_KATALOG,
            "xl/sharedStrings.xml": KATALOG,
        }
    else:
        czesci = {
            "[Content_Types].xml": CONTENT_TYPES_INLINE,
            "_rels/.rels": ROOT_RELS,
            "xl/workbook.xml": WORKBOOK,
            "xl/_rels/workbook.xml.rels": WB_RELS_INLINE,
            "xl/worksheets/sheet1.xml": ARKUSZ_INLINE,
        }

    # [Content_Types].xml zapisujemy pierwszy - niektóre narzędzia tego oczekują.
    with zipfile.ZipFile(sciezka, "w", zipfile.ZIP_DEFLATED) as archiwum:
        for nazwa, tresc in czesci.items():
            archiwum.writestr(nazwa, tresc.encode("utf-8"))


def rozbierz(sciezka: Path) -> None:
    print(f"\n{'=' * 74}\nPLIK: {sciezka.name}\n{'=' * 74}")
    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()
        print(f"Części ({len(nazwy)}): {', '.join(sorted(nazwy))}")
        print(f"sharedStrings.xml obecne: "
              f"{'TAK' if 'xl/sharedStrings.xml' in nazwy else 'NIE'}")

        arkusz = ET.fromstring(archiwum.read("xl/worksheets/sheet1.xml"))
        print("Komórki w arkuszu:")

        katalog: list[str] = []
        if "xl/sharedStrings.xml" in nazwy:
            sst = ET.fromstring(archiwum.read("xl/sharedStrings.xml"))
            for si in sst:
                if si.tag.rsplit("}", 1)[-1] != "si":
                    continue
                katalog.append("".join(
                    (t.text or "") for t in si.iter()
                    if t.tag.rsplit("}", 1)[-1] == "t"
                ))

        for c in arkusz.iter():
            if c.tag.rsplit("}", 1)[-1] != "c":
                continue
            adres = c.get("r")
            typ = c.get("t", "n")
            if typ == "inlineStr":
                is_el = next(x for x in c if x.tag.rsplit("}", 1)[-1] == "is")
                tekst = "".join(
                    (t.text or "") for t in is_el.iter()
                    if t.tag.rsplit("}", 1)[-1] == "t"
                )
                print(f"   {adres}: t='inlineStr'  tekst={tekst!r}  <- tekst W KOMÓRCE")
            elif typ == "s":
                v = next(x for x in c if x.tag.rsplit("}", 1)[-1] == "v")
                indeks = int(v.text)
                print(f"   {adres}: t='s'  v={indeks}  -> katalog[{indeks}]="
                      f"{katalog[indeks]!r}   <- to są INDEKSY (od 0!)")
            else:
                v = next(x for x in c if x.tag.rsplit("}", 1)[-1] == "v")
                print(f"   {adres}: t={typ!r}  v={v.text}  <- liczba wprost")


def glowny() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    a = OUTPUT / "02_katalog.xlsx"
    b = OUTPUT / "02_inline.xlsx"

    zbuduj(a, "katalog")
    zbuduj(b, "inline")

    rozbierz(a)
    rozbierz(b)

    print(f"\n{'=' * 74}")
    print("WNIOSEK")
    print("=" * 74)
    print(f"Rozmiar z katalogiem: {a.stat().st_size:>5} B")
    print(f"Rozmiar z inlineStr : {b.stat().st_size:>5} B")
    print()
    print("Ten sam plik, ta sama zawartość, dwie poprawne reprezentacje.")
    print("Excel buduje katalog tekstów. openpyxl pisze teksty w komórkach.")
    print()
    print("Oba pliki otwórz w Excelu/LibreOffice - wyglądają identycznie.")


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** budujemy słowniki `nazwa części → treść XML` i pakujemy je do ZIP. Zero openpyxl — to dowód, że sam format nie potrzebuje żadnej biblioteki.

**Co trafi do pliku:** dwa minimalne, ale **poprawne** skoroszyty `.xlsx`. Otwórz je w Excelu — zobaczysz dane i liczby, bez żadnych ostrzeżeń.

**Konsekwencja praktyczna:** zobaczysz na własne oczy, że `t="s"` w wierszach 2–4 wskazuje ten sam indeks `1`, a katalog ma „Warszawa" tylko raz. A plik inline ma „Warszawa" trzy razy. Przy danych o wysokiej powtarzalności to różnica w rozmiarze — i dokładnie ta różnica występuje między plikiem z Excela a plikiem z openpyxl.

### Przykład 5 — Detektor rozszerzeń: złap ostrzeżenie openpyxl i udowodnij utratę

Ostatni przykład jest pomostem do modułu 18. Wstrzykniemy do pliku blok rozszerzeń `extLst` z URI sparkline'a, wczytamy go openpyxl-em, złapiemy ostrzeżenie, zapiszemy — i pokażemy, że extension zniknął.

```python
"""Wstrzykuje blok extLst (sparkline) do pliku, a potem pokazuje, że openpyxl go gubi.

Uruchom:  python examples/02_detektor_rozszerzen.py
Efekt:    output/02_ze_sparkline.xlsx   (ma extension)
          output/02_po_openpyxl.xlsx    (extension już nie ma)
"""

import shutil
import warnings
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

# URI grupy sparkline w rozszerzeniu x14 (Excel 2010+).
URI_SPARKLINE = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"

# UPROSZCZONY blok <ext>. openpyxl reaguje na sam URI, więc do demonstracji
# wystarczy samo <ext uri="..."/>. W prawdziwym pliku Excela w środku jest
# grupa sparkline w przestrzeni nazw x14 (patrz teoria, sekcja 2.12).
WSTRZYKNIECIE = (
    f'<extLst><ext uri="{URI_SPARKLINE}"/></extLst>'
).encode("utf-8")


def zbuduj_bazowy() -> Path:
    """Zwykły plik openpyxl, do którego wstrzykniemy extension."""
    sciezka = OUTPUT / "02_baza.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sparkline"
    ws["A1"] = "Produkt"
    for i, wart in enumerate(["Q1", "Q2", "Q3", "Q4"], start=2):
        ws.cell(row=1, column=i, value=wart)
    ws["A2"] = "Widżet"
    for i, wart in enumerate([10, 40, 20, 55], start=2):
        ws.cell(row=2, column=i, value=wart)
    wb.save(sciezka)
    return sciezka


def wstrzyknij_extension(zrodlo: Path, cel: Path) -> None:
    """Dodaje <extLst> na koniec pierwszego arkusza i zapisuje nowy ZIP."""
    with zipfile.ZipFile(zrodlo) as wejscie:
        infos = wejscie.infolist()
        dane = {info.filename: wejscie.read(info.filename) for info in infos}

    # Znajdź pierwszy arkusz.
    nazwa_arkusza = next(
        n for n in dane
        if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
    )
    xml = dane[nazwa_arkusza]

    # Wstawiamy przed zamykającym tagiem arkusza. rfind jest odporny na to,
    # że w pliku mogą wystąpić inne wystąpienia tekstu "</worksheet>".
    pozycja = xml.rfind(b"</worksheet>")
    if pozycja == -1:
        raise RuntimeError("Nie znalazłem zamykającego tagu arkusza.")
    dane[nazwa_arkusza] = xml[:pozycja] + WSTRZYKNIECIE + xml[pozycja:]

    with zipfile.ZipFile(cel, "w", zipfile.ZIP_DEFLATED) as wyjscie:
        for info in infos:
            wyjscie.writestr(info.filename, dane[info.filename])


def czy_ma_extension(sciezka: Path) -> bool:
    with zipfile.ZipFile(sciezka) as archiwum:
        for nazwa in archiwum.namelist():
            if nazwa.startswith("xl/worksheets/sheet") and nazwa.endswith(".xml"):
                if URI_SPARKLINE.encode() in archiwum.read(nazwa):
                    return True
    return False


def glowny() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)

    baza = zbuduj_bazowy()
    ze_sparkline = OUTPUT / "02_ze_sparkline.xlsx"
    po_openpyxl = OUTPUT / "02_po_openpyxl.xlsx"

    wstrzyknij_extension(baza, ze_sparkline)
    print(f"1. Zbudowałem plik i wstrzyknąłem extension: {ze_sparkline.name}")
    print(f"   Czy plik MA extension? {czy_ma_extension(ze_sparkline)}")

    # KLUCZOWY MOMENT: wczytujemy i przechwytujemy ostrzeżenia.
    print("\n2. Wczytuję plik openpyxl-em i przechwytuję ostrzeżenia:")
    with warnings.catch_warnings(record=True) as zebrane:
        warnings.simplefilter("always")     # bez tego Python pokaże tylko raz
        wb = load_workbook(ze_sparkline)

    if not zebrane:
        print("   (brak ostrzeżeń - sprawdź wersję openpyxl)")
    for ostrzezenie in zebrane:
        print(f"   ! {type(ostrzezenie.message).__name__}: {ostrzezenie.message}")

    # Zapis "przepisuje" model - a w modelu nie ma extLst.
    wb.save(po_openpyxl)
    print(f"\n3. Zapisałem ten sam skoroszyt jako: {po_openpyxl.name}")
    print(f"   Czy plik MA extension? {czy_ma_extension(po_openpyxl)}")

    print(f"\n{'=' * 74}")
    print("WNIOSEK")
    print("=" * 74)
    print("Wczytanie -> ostrzeżenie. Zapis -> utrata. Dokładnie tak, jak mówi teoria:")
    print("openpyxl nie modeluje bloków extLst, więc przy przepisaniu pakietu")
    print("nie ma z czego ich odtworzyć. To samo dotyczy sparkline'ów")
    print("w prawdziwych plikach Excela.")
    print()
    print("Otwórz oba pliki w Excelu: w pierwszym zakładka jest 'bogatsza',")
    print("w drugim - wygląda tak samo, bo uproszczony <ext/> nic nie rysuje.")
    print("W prawdziwym pliku ze sparkline'ami różnica byłaby widoczna.")
    print()
    print("Uwaga: NIE wyciszaj tych ostrzeżeń przez warnings.filterwarnings('ignore').")
    print("To jedyny sygnał, że tracisz dane.")


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** wczytujemy plik do modelu openpyxl. Blok `extLst` **nie ma odpowiednika** w tym modelu — parser go widzi, ostrzega, i zapomina. `wb` w pamięci nie wie o istnieniu sparkline'a.

**Co trafi do pliku:** `output/02_po_openpyxl.xlsx` — pakiet bez bloku `extLst`. To nie jest „uszkodzenie" ani „błąd". To konsekwencja przepisania modelu.

**Oczekiwany wynik:**

```text
1. Zbudowałem plik i wstrzyknąłem extension: 02_ze_sparkline.xlsx
   Czy plik MA extension? True

2. Wczytuję plik openpyxl-em i przechwytuję ostrzeżenia:
   ! UserWarning: Sparkline Group extension is not supported and will be removed

3. Zapisałem ten sam skoroszyt jako: 02_po_openpyxl.xlsx
   Czy plik MA extension? False
```

To jest cała esencja modułu 18, zademonstrowana w jednym skrypcie: **wczytanie → ostrzeżenie; zapis → utrata.** Ostrzeżenie nie jest ozdobą konsoli. Jest ostatnim momentem, w którym możesz zareagować.

## 4. Anatomia API

Tabela zbiera wszystko, czego użyliśmy. Uwaga: pierwsza połowa to **biblioteka standardowa Pythona**, nie openpyxl. To celowe — anatomia pakietu to wiedza o formacie, a nie o bibliotece.

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `zipfile.is_zipfile(path)` | Szybki test: czy plik jest archiwum ZIP | ścieżka | Czytanie nagłówka, natychmiastowe. Rozdziela `.xlsx` od `.xls` |
| `zipfile.ZipFile(path)` | Otwiera archiwum do czytania/pisania | `mode`, `compression` | Używaj jako kontekst menedżera (`with`) |
| `ZipFile.namelist()` | Lista nazw części | — | Kolejność **nie jest** gwarantowana semantycznie |
| `ZipFile.infolist()` | Lista obiektów `ZipInfo` | — | Zawiera `filename`, `file_size`, `compress_size` |
| `ZipFile.read(nazwa)` | Zwraca bajty części | nazwa | Rzuca `KeyError`, gdy części nie ma |
| `ZipFile.writestr(info_lub_nazwa, dane)` | Zapisuje część do archiwum | — | Tak zbudujesz własny pakiet |
| `xml.etree.ElementTree.fromstring(bytes)` | Buduje drzewo XML (lekkie, szybkie) | bajty lub tekst | Do pracy z danymi |
| `ElementTree.Element.iter()` | Iteruje po wszystkich elementach poddrzewa | — | Tak znajdziesz wszystkie `<c>` |
| `Element.get(klucz, domyślna)` | Atrybut z wartością domyślną | — | **Zawsze dawaj domyślną**: `get("t", "n")`, `get("s", "0")` |
| `Element.find`, `Element.findall` | Wyszukiwanie dzieci po tagu | tag z przestrzenią nazw | Wymaga pełnego `{ns}tag` |
| `ElementTree.tostring(el, encoding="unicode")` | Serializacja elementu do tekstu | — | Poddrzewo dostaje własne deklaracje `xmlns` |
| `xml.dom.minidom.parseString(bytes)` | Buduje drzewo DOM (wygodne, cięższe) | bajty | Do **podglądania**, nie do danych |
| `Document.toprettyxml(indent=...)` | Ładnie sformatowany XML | — | Dokłada puste linie — trzeba je odfiltrować |
| `Document.getElementsByTagName(nazwa)` | Wyszukiwanie po nazwie kwalifikowanej | — | `toxml()` zwraca fragment z deklaracjami |
| `openpyxl.utils.get_column_letter(n)` | `28` → `"AB"` | liczba ≥ 1 | 1-indeksowane |
| `openpyxl.utils.column_index_from_string("AB")` | `"AB"` → `28` | tekst | 1-indeksowane |
| `openpyxl.utils.cell.coordinate_to_tuple("B2")` | `"B2"` → `(2, 2)` (wiersz, kolumna) | tekst | 1-indeksowane |
| `openpyxl.utils.cell.coordinate_from_string("B2")` | `"B2"` → `("B", "2")` | tekst | |
| `openpyxl.utils.cell.range_boundaries("A1:C3")` | `(1, 1, 3, 3)` | tekst | Kolejność: min_col, min_row, max_col, max_row |
| `openpyxl.utils.cell.quote_sheetname("Dane 2026")` | `"'Dane 2026'"` | tekst | Do formuł i odwołań |
| `openpyxl.utils.datetime.from_excel(wartość, epoch)` | Liczba → `datetime` | wartość, epoka | Excel trzyma daty jako liczby |
| `openpyxl.utils.datetime.to_excel(dt, epoch)` | `datetime` → liczba | — | |
| `openpyxl.utils.datetime.CALENDAR_WINDOWS_1900` | Epoka systemu Windows | — | Domyślna |
| `openpyxl.utils.datetime.CALENDAR_MAC_1904` | Epoka systemu Mac | — | Spotykana w starych plikach |
| `load_workbook(path, keep_vba=True)` | Zachowuje projekt VBA i części zależne | — | Bez tego `vbaProject.bin` i formanty przepadają |
| `load_workbook(path, rich_text=True)` | Teksty bogate jako `CellRichText` | — | Nowość 3.1; dotyczy też `inlineStr` |
| `openpyxl.packaging.relationship` | Model relacji pakietu | — | **API wewnętrzne** — zmieni się bez ostrzeżenia |
| `openpyxl.xml.constants` | Stałe ścieżek (`ARC_*`, `PACKAGE_*`) | — | **API wewnętrzne**; użyteczne do nauki, nie do produkcji |

Zasada dotycząca ostatnich dwóch wierszy: możesz je **czytać**, żeby zrozumieć bibliotekę (tak jak my to robiliśmy w teorii), ale **nie importuj ich w kodzie produkcyjnym**. Podobnie rzecz się ma z `openpyxl.cell._writer` — to świetne źródło wiedzy o formacie, ale nie kontrakt API.

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — inwentarz własnego pliku.** Uruchom `examples/02_inwentarz.py` na pliku `output/01_formula_dowod.xlsx` z modułu 01. Następnie na **innym** pliku `.xlsx`, który masz na dysku — najlepiej takim, który powstał w Excelu, nie w Pythonie. Zanotuj w pliku `output/02_uwagi.md`:

- Ile części ma każdy plik?
- Czy któryś ma `xl/sharedStrings.xml`?
- Czy któryś ma części, które Twój inwentarz zakwalifikował jako „NIEZNANE"?
- Które kategorie w drugim pliku są puste, a które pojawiły się jako nowe?

**Zadanie 2 — gdzie jest tekst, a gdzie liczba.** Otwórz dowolny plik `.xlsx` (choćby z zadania 1) i odnajdź w `xl/worksheets/sheet1.xml` jedną komórkę z tekstem i jedną z liczbą. Odpowiedz pisemnie:

- Jakiego atrybutu używa komórka z tekstem, a jakiego z liczbą?
- Jeśli w pliku nie ma `sharedStrings.xml`, gdzie fizycznie leży tekst?

### 🟡 Warsztat

**Zadanie 3 — znajdź indeks stylu i rozstrzygnij, którego `Font` używa.** To ćwiczenie jest sedno modułu. Postępuj tak:

1. Utwórz plik skryptem: sformatuj komórkę `B2` tak, żeby miała pogrubioną czcionkę **i** wypełnienie **i** obramowanie. Nadaj też komórce `C3` tylko pochyloną czcionkę.
2. Zapisz plik i **nie używaj** `02_tropiciel_stylu.py` — najpierw zrób to ręcznie. Rozpakuj plik (albo użyj `zipfile`) i znajdź `xl/worksheets/sheet1.xml`. Odczytaj `s="..."` z komórki `B2`.
3. Wejdź do `xl/styles.xml`. Znajdź wpis numer `s` na liście `<cellXfs>`. Zapisz jego `fontId`.
4. Wejdź na listę `<fonts>` pod indeksem `fontId`. **Której czcionki używa komórka B2?** Zapisz pełny opis (nazwa, rozmiar, pogrubienie, kolor).
5. Powtórz dla `C3`. Czy `C3` wskazuje na ten sam wpis w `<fonts>`, czy na inny? Dlaczego?
6. **Dopiero teraz** uruchom `02_tropiciel_stylu.py` na swoim pliku i sprawdź, czy Twój ręczny łańcuch się zgadza.

Zapisz cały łańcuch w pliku `output/02_lancuch_stylu.md` w formacie: `B2 → s=N → cellXfs[N] → fontId=M → fonts[M] → opis`.

**Zadanie 4 — zbuduj minimalny `.xlsx` ręcznie.** Uruchom `examples/02_dwa_swiaty_textu.py`, a potem **zmodyfikuj go tak**, żeby zbudować plik z trzema arkuszami o nazwach `Dane`, `Podsumowanie` i `Konfiguracja`. Będziesz musiał:

- dodać trzy części `xl/worksheets/sheetN.xml`,
- dodać trzy wpisy `<Override>` w `[Content_Types].xml`,
- dodać trzy relacje w `xl/_rels/workbook.xml.rels` (typu `worksheet`),
- dodać trzy elementy `<sheet>` w `xl/workbook.xml`, każdy z własnym `r:id`.

Nie używaj openpyxl. Otwórz wynik w Excelu. Jeśli Excel zgłosi problem z zawartością — **to jest część ćwiczenia**: przeczytaj komunikat, znajdź brakujący wpis i napraw.

**Zadanie 5 — policz oszczędność katalogu tekstów.** Zmodyfikuj `examples/02_dwa_swiaty_textu.py` tak, aby wygenerował plik z **50 000 komórek** zawierających jeden z pięciu powtarzających się tekstów — raz z katalogiem tekstów, raz z `inlineStr`. Porównaj rozmiary plików na dysku. Zapisz wynik w `output/02_oszczednosc.md` i odpowiedz: przy jakiej powtarzalności danych katalog tekstów opłaca się najmocniej, a kiedy różnica praktycznie znika? (Podpowiedź: pomyśl o danych z unikalnymi identyfikatorami.)

### 🔴 Wyzwanie

**Zadanie 6 — porównaj plik z Excela i plik z openpyxl.** To zadanie jest wprost wskazane w celach modułu. Wykonaj:

1. Utwórz w Excelu (ręcznie, nie kodem) skoroszyt z:
   - dwoma arkuszami,
   - nagłówkiem z pogrubioną czcionką i kolorowym tłem,
   - kolumną z formatem walutowym,
   - wykresem słupkowym,
   - obrazkiem wstawionym z pliku.
2. Zapisz go do `data/`.
3. Utwórz **programowo** (openpyxl-em) plik o **możliwie najbardziej zbliżonej** zawartości i strukturze.
4. Napisz skrypt `examples/02_porownaj.py`, który:
   - wypisuje inwentarz obu plików,
   - liczbę części w każdej kategorii,
   - **różnice w zbiorach nazw części** (co jest w jednym, a czego nie ma w drugim),
   - obecność i długość `xl/sharedStrings.xml`,
   - rozmiary obu plików.
5. Odpowiedz pisemnie w `output/02_porownanie.md` na cztery pytania:
   - Których części ma plik z Excela, a nie ma pliku z openpyxl?
   - Które z nich są **opcjonalne** (można ich nie mieć), a które oznaczają **utratę funkcji**?
   - Gdybyś teraz wziął plik z Excela, wczytał go openpyxl-em i zapisał — **co dokładnie by zniknęło**? Odpowiedz, wymieniając konkretne nazwy części.
   - Który plik jest większy i dlaczego? (Pomyśl o katalogu tekstów, o liczbie definicji stylów i o tym, że openpyxl zawsze zapisuje część `theme1.xml`.)
6. Na koniec uruchom `examples/02_detektor_rozszerzen.py` na pliku z Excela (nie na wygenerowanym przez openpyxl). Zobacz, jakie rozszerzenia openpyxl zgłasza w Twoim konkretnym pliku.

<details>
<summary><strong>Szkic rozwiązania zadania 6 — szkiclet skryptu `porownaj.py`</strong></summary>

```python
"""Porównuje dwa pliki .xlsx na poziomie CZĘŚCI PAKIETU.

Uruchom:  python examples/02_porownaj.py data/raport_z_excela.xlsx output/raport_openpyxl.xlsx
"""

import sys
import zipfile
from pathlib import Path


def kategorie() -> dict[str, tuple[str, ...]]:
    """Ta sama klasyfikacja, co w przykładzie 1."""
    return {
        "Spis i powiązania": ("[Content_Types].xml", "_rels/"),
        "Metadane": ("docProps/",),
        "Szkielet": ("xl/workbook.xml", "xl/_rels/"),
        "Arkusze": ("xl/worksheets/",),
        "Style i motyw": ("xl/styles.xml", "xl/theme/"),
        "Katalog tekstów": ("xl/sharedStrings.xml",),
        "Rysunki i wykresy": ("xl/drawings/", "xl/charts/", "xl/media/"),
        "Tabele": ("xl/tables/",),
        "Komentarze": ("xl/comments", "xl/threadedComments/"),
        "Obliczenia": ("xl/calcChain.xml",),
        "Tabele przestawne": ("xl/pivotTables/", "xl/pivotCache/"),
        "Makra i formanty": ("xl/vbaProject.bin", "xl/ctrlProps/", "xl/activeX/"),
        "Slicery": ("xl/slicers/", "xl/slicerCaches/"),
        "Power Query": ("xl/connections.xml", "xl/queryTables/", "xl/model/"),
    }


def czesci(sciezka: Path) -> set[str]:
    with zipfile.ZipFile(sciezka) as archiwum:
        return set(archiwum.namelist())


def zgrupuj(nazwy: set[str]) -> dict[str, set[str]]:
    grupy: dict[str, set[str]] = {}
    for nazwa in nazwy:
        for kat, wzorce in kategorie().items():
            if any(w in nazwa for w in wzorce):
                grupy.setdefault(kat, set()).add(nazwa)
                break
        else:
            grupy.setdefault("NIEZNANE", set()).add(nazwa)
    return grupy


def main() -> None:
    if len(sys.argv) != 3:
        print("Użycie: python examples/02_porownaj.py A.xlsx B.xlsx")
        return

    a, b = Path(sys.argv[1]), Path(sys.argv[2])
    ca, cb = czesci(a), czesci(b)

    print(f"A = {a.name}  ({a.stat().st_size} B, {len(ca)} części)")
    print(f"B = {b.name}  ({b.stat().st_size} B, {len(cb)} części)\n")

    ga, gb = zgrupuj(ca), zgrupuj(cb)

    print(f"{'Kategoria':<28} {'A':>5} {'B':>5}")
    for kat in sorted(set(ga) | set(gb)):
        print(f"{kat:<28} {len(ga.get(kat, ())):>5} {len(gb.get(kat, ())):>5}")

    print("\n--- TYLKO W A (potencjalna utrata przy przepisaniu przez openpyxl):")
    for nazwa in sorted(ca - cb):
        print(f"   ! {nazwa}")

    print("\n--- TYLKO W B:")
    for nazwa in sorted(cb - ca):
        print(f"   + {nazwa}")

    print("\n--- WSPÓLNE:")
    print(f"   {len(ca & cb)} części")

    # Specjalna uwaga na katalog tekstów.
    sst_a = "xl/sharedStrings.xml" in ca
    sst_b = "xl/sharedStrings.xml" in cb
    print(f"\nKatalog tekstów:  A={'TAK' if sst_a else 'NIE'}  "
          f"B={'TAK' if sst_b else 'NIE'}")


if __name__ == "__main__":
    main()
```

Wnioski, do których powinieneś dojść:

- **Opcjonalne, ale nie oznaczają utraty funkcji:** `xl/sharedStrings.xml` (to tylko inny sposób zapisu tego samego tekstu), `xl/calcChain.xml` (podpowiedź optymalizacyjna).
- **Oznaczają utratę funkcji:** wszystko z kategorii „Makra i formanty", „Slicery", „Power Query", a także bloki `<extLst>` (sparkline'y) — których **nie widać** w liście nazw części, bo siedzą wewnątrz XML arkusza. Dlatego sam inwentarz nazw nie wystarczy; potrzebujesz jeszcze skanera zawartości (przykład 5 i moduł 18).
- **Rozmiar:** plik z Excela zwykle wygrywa przy powtarzalnych tekstach (katalog), a przegrywa przy danych unikalnych. openpyxl zawsze dopisuje `xl/theme/theme1.xml` (motyw domyślny) nawet jeśli plik wyszedł z Excela — to jeden ze sposobów na wykrycie, że plik „przeszedł przez openpyxl".
- **Najważniejszy wniosek:** jeśli w zestawieniu „TYLKO W A" pojawi się cokolwiek z kategorii ryzykownych, **nie zapisuj tego pliku openpyxl-em**. To jest dokładnie ta decyzja, którą moduł 18 zamieni w procedurę.

</details>

## 6. Typowe błędy i pułapki

**1. „`load_workbook("dane.xls")` rzuca wyjątek, choć Excel otwiera plik bez problemu" (objaw) → `.xls` to binarny format BIFF, nie archiwum ZIP; openpyxl obsługuje wyłącznie rodzinę OOXML (objaw wtórny: `zipfile.is_zipfile()` zwraca `False`) → przekonwertuj plik do `.xlsx` (LibreOffice headless, Excel), a do samego odczytu `.xls` w Pythonie użyj `xlrd` (naprawa).**

**2. „Edytowałem `sheet1.xml` bezpośrednio w archiwum i Excel mówi, że plik jest uszkodzony" (objaw) → trzy możliwe przyczyny: (a) podmieniłeś bajty w archiwum bez przeliczenia sumy kontrolnej CRC32 wpisu ZIP, (b) rozjechałeś spójność relacji (`r:id` wskazuje na nieistniejącą relację), (c) nie zaktualizowałeś `[Content_Types].xml` po dodaniu nowej części → nie edytuj plików w archiwum „w miejscu"; rozpakuj, zmień, zapakuj od nowa używając `zipfile`, i po każdej zmianie struktury zaktualizuj `[Content_Types].xml` oraz `_rels` (naprawa).**

To jest najczęstsza przyczyna „magicznego" uszkodzenia plików `.xlsx`. ZIP nie jest formatem, który znosi edycję bajtową — każdy wpis ma sumę kontrolną, a OOXML ma własne reguły spójności ponad ZIP-em.

**3. „Moje narzędzie czyta tekst z arkusza i nie widzi nic, choć w Excelu tekst jest" (objaw) → narzędzie obsługuje tylko `t="s"` (odsyłacz do katalogu), a plik używa `t="inlineStr"` (tekstu w komórce) → rozszerz parser o obsługę obu wariantów: `t="s"` wymaga rozwiązania indeksu w katalogu, `t="inlineStr"` wymaga odczytania `<is><t>` (naprawa).**

Pamiętaj, że to działa w obie strony: narzędzie, które obsługuje tylko `inlineStr`, też zawiedzie na pliku z Excela. A openpyxl czyta **oba** warianty poprawnie — to właśnie jedna z jego mocniejszych stron.

**4. „Wziąłem indeks tekstu z komórki i wziąłem `katalog[indeks]` — dostałem sąsiedni tekst" (objaw) → `<v>` w komórce ze `t="s"` jest indeksem **0-indeksowanym**, a Ty przeoczyłeś, że indeksy w katalogu liczą od 0 → pamiętaj: `<v>0</v>` to **pierwszy** tekst w katalogu, `<v>1</v>` to drugi (naprawa).**

To samo dotyczy `s="0"` w komórce (pierwszy styl) i `fontId="0"` (pierwsza czcionka). Ale adres `A1` w tym samym pliku liczy od 1. W module 03 zbierzemy to w tabelę reguł.

**5. „Napisałem funkcję, która sprawdza, czy element XML istnieje, i zawsze zwraca `False`, choć element tam jest" (objaw) → w `xml.etree.ElementTree` wartość logiczna `Element` zależy od tego, czy ma dzieci: `bool(Element("b"))` to `False`, bo `<b/>` jest pusty → **nigdy** nie pisz `if element:` dla elementów; zawsze `if element is not None:` (naprawa).**

To realna pułapka Pythona, na którą trafisz przy parsowaniu `styles.xml` (elementy `<b/>`, `<i/>`, `<u/>` są puste, ale znaczą „pogrubiona", „kursywa", „podkreślona"!). Dlatego w przykładzie 3 użyliśmy `is not None` konsekwentnie.

**6. „Dodałem nową część do archiwum, ale Excel jej nie widzi" (objaw) → OOXML wymaga, aby każda część była zadeklarowana w `[Content_Types].xml`; część bez deklaracji jest dla odbiorcy niewidzialna → dodaj odpowiedni `<Override PartName="/..." ContentType="..."/>` (albo `<Default Extension="..."/>`), a jeśli część jest odwoływana z innej, dodaj jeszcze relację (naprawa).**

**7. „Wyciszyłem ostrzeżenia przez `warnings.filterwarnings("ignore")` i wszystko działa" (objaw) → ostrzeżenia openpyxl o rozszerzeniach to **jedyny** sygnał, że przy zapisie stracisz dane → nie wyciszaj ich globalnie; filtruj konkretne ostrzejenie po treści i tylko wtedy, gdy naprawdę wiesz, że utrata jest zamierzona i udokumentowana (naprawa).**

Antywzorzec, który kosztuje najwięcej w długiej perspektywie: cisza w konsoli jest wygodna przez tydzień i katastrofalna w momencie, gdy ktoś zauważy brakujące dane w raporcie.

**8. „Wczytałem plik `.xlsm` openpyxl-em i mój formularz z makrami stracił kontrolki ActiveX" (objaw) → części `xl/activeX/`, `xl/ctrlProps/` i `customUI/` są kopiowane tylko wtedy, gdy wczytasz plik z `keep_vba=True` (kopiowanie odbywa się w ścieżce scalania projektu VBA) → wczytaj z `keep_vba=True` i **mimo to** przetestuj wynik, bo ścieżka VBA nie gwarantuje zachowania wszystkich powiązanych części (naprawa).**

**9. „Zapisałem plik openpyxl-em i na Macu Excel zgłasza problem z zawartością, na Windowsie jest OK" (objaw) → zgłaszane są problemy z bardzo rygorystycznymi parserami na niektórych platformach w związku z zapisem tekstów jako `inlineStr` i stanem `calcPr` bez `calcChain`; źródła nie są jednorodne → **nie zakładaj winy biblioteki ani jej niewinności**: otwórz wygenerowany plik na docelowych platformach i wersjach Excela przed wdrożeniem; jeśli problem wystąpi, rozważ bibliotekę zapisującą katalog tekstów albo narzędzie normalizujące pakiet (naprawa).**

To piękny przykład sytuacji, w której prawda nie jest zerojedynkowa. Nie prezentuj tego jako „potwierdzonego błędu openpyxl", ale też nie zamiataj — sprawdź na swoim przypadku.

**10. „Zrobiłem kopię pliku przez `copy_worksheet` i w kopii nie ma ani obrazka, ani wykresu" (objaw) → `copy_worksheet` kopiuje komórki i część atrybutów arkusza, ale **nie kopiuje obrazów ani wykresów** — one żyją w warstwie rysunkowej, której ta metoda nie przenosi (przyczyna) → jeśli potrzebujesz kopii z grafiką, kopiuj **plik**, nie arkusz (naprawa).**

Szczegóły w module 08, ale warto wiedzieć już teraz — to jeden z tych mitów, które „wyglądają jakby działały", dopóki ktoś nie otworzy kopii.

**11. „Chciałem odczytać wartość formuły przez `data_only=True` i dostałem `None`" (objaw) → komórka z formułą ma `<f>` bez `<v>` — nie ma zapisanej wartości z pamięci podręcznej, bo openpyxl jej nie policzył, a żaden Excel tego pliku nie otwierał → odczytaj formułę bez `data_only` (dostaniesz tekst `"=SUM(...)"`), albo przelicz plik w Excelu/LibreOffice, albo policz wartość w Pythonie (naprawa).**

Zwróć uwagę na jedną rzecz: **plik zapisany w Excelu ma `<v>` obok `<f>`, a plik zapisany w openpyxl nie ma.** To kolejna namacalna różnica między tymi dwoma światami — i kolejny powód, dla którego „openpyxl a Excel" to nie jedna biblioteka o dwóch twarzach, tylko dwie różne implementacje wspólnego formatu.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **`.xlsx` to archiwum ZIP z plikami XML.** Rozszerzenie na `.zip` i rozpakowanie to pełnoprawna metoda diagnostyczna. Sprawdzenie `zipfile.is_zipfile()` deterministycznie rozdziela obsługiwane pliki od `.xls`/`.xlsb`.

2. **Pakiet to teczka z metryką, spisem treści i spinkami.** `[Content_Types].xml` deklaruje typ każdej części, `_rels/` prowadzi od części do części przez identyfikatory, `xl/workbook.xml` to spis treści skoroszytu, a same kartki (`sheetN.xml`), legenda wyglądu (`styles.xml`) i katalog tekstów (`sharedStrings.xml`) to oddzielne części.

3. **Komórka nie ma formatowania — ma numer stylu.** `s="2"` to odsyłacz do `cellXfs[2]` w `styles.xml`, a ten z kolei do `fonts[fontId]`, `fills[fillId]`, `borders[borderId]`. Zmiana jednego wpisu w `<fonts>` zmienia wygląd tysięcy komórek. To jest jednocześnie mechanizm deduplikacji i źródło niespodzianek.

4. **Dwa systemy indeksowania współistnieją w jednym pliku.** Adresy i wiersze liczą od 1; indeksy stylów, wpisów w magazynach, katalogu tekstów i `dxfs` — od 0. W jednej linijce XML możesz mieć `C1` (od 1), `s="1"` (od 0) i `<v>0</v>` (od 0).

5. **openpyxl używa `inlineStr`, Excel używa katalogu tekstów — oba warianty są poprawne.** openpyxl domyślnie nie tworzy części `sharedStrings.xml`; zapisuje teksty w komórkach. Konsekwencje: inny rozmiar pliku przy danych powtarzalnych, inna struktura XML do porównywania w testach, potencjalne problemy z parserami zakładającymi katalog. I najważniejsze: **wszystko, co nie ma odpowiednika w modelu openpyxl — `vbaProject.bin` bez `keep_vba`, slicery, osie czasu, Power Query, bloki `extLst` — przy zapisie ginie.** Ostrzeżenie w konsoli to nie hałas, to jedyny komunikat o utracie danych.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. ROZPOZNANIE FORMATU (biblioteka standardowa, bez openpyxl)
# ==================================================================
import zipfile
from pathlib import Path

sciezka = Path("output/raport.xlsx")

zipfile.is_zipfile(sciezka)          # True dla .xlsx, False dla .xls / .xlsb

with zipfile.ZipFile(sciezka) as a:
    a.namelist()                     # lista nazw części
    a.infolist()                     # ZipInfo: filename, file_size, compress_size
    a.read("xl/workbook.xml")        # bajty części (KeyError gdy brak)

# ==================================================================
# 2. STRUKTURA PAKIETU - mapa najważniejszych części
# ==================================================================
#   [Content_Types].xml              typ każdej części (bez wpisu = część nie istnieje)
#   _rels/.rels                      punkt startowy -> xl/workbook.xml
#   docProps/core.xml                autor, tytuł, daty
#   docProps/app.xml                 metadane aplikacji (openpyxl generuje od nowa)
#   docProps/custom.xml              własne właściwości (openpyxl 3.1)
#   xl/workbook.xml                  lista arkuszy, nazwy zdefiniowane, <calcPr/>
#   xl/_rels/workbook.xml.rels       spinki: arkusz, style, katalog tekstów, motyw
#   xl/worksheets/sheetN.xml         komórki, kolumny, scalenia, wydruk
#   xl/styles.xml                    fonts / fills / borders / cellXfs / dxfs
#   xl/sharedStrings.xml             katalog tekstów (Excel: jest; openpyxl: NIE MA)
#   xl/theme/theme1.xml              kolory i czcionki motywu
#   xl/drawings/, xl/charts/, xl/media/    warstwa rysunkowa, wykresy, obrazy
#   xl/tables/tableN.xml             tabele Excela
#   xl/calcChain.xml                 kolejność przeliczania (podpowiedź Excela)
#   xl/pivotTables/, xl/pivotCache/  tabele przestawne
#   xl/vbaProject.bin                makra (tylko keep_vba=True)

# ==================================================================
# 3. KOMÓRKA - co znaczy każdy kawałek XML
# ==================================================================
#   <c r="C1" s="1" t="s"><v>0</v></c>
#
#     r="C1"   adres w notacji A1                     -> 1-INDEKSOWANE
#     s="1"    numer stylu -> cellXfs[1]              -> 0-INDEKSOWANE
#     t="s"    typ: brak/n=liczba, s=sharedString,
#              inlineStr=tekst w komórce, str=wynik
#              formuły tekstowej, b=bool, e=błąd, d=data
#     <v>0</v> wartość; przy t="s" to INDEKS do katalogu -> 0-INDEKSOWANE
#     <f>...</f> formuła BEZ znaku "="
#
#   Domyślne wartości przy braku atrybutu (jak w parserze openpyxl):
#     t -> "n"        s -> 0
#   Dlatego zawsze: c.get("t", "n")   oraz   c.get("s", 0)

# ==================================================================
# 4. ŁAŃCUCH STYLU - jak dojść od komórki do wyglądu
# ==================================================================
import xml.etree.ElementTree as ET

def bez_ns(tag: str) -> str:
    """'http://...}font' -> 'font'."""
    return tag.rsplit("}", 1)[-1]

def dziecko(el, nazwa):
    """Patrz uważnie: puste <b/> ma bool() == False! Zawsze 'is not None'."""
    for d in el:
        if bez_ns(d.tag) == nazwa:
            return d
    return None

def dzieci(el, nazwa):
    return [d for d in el if bez_ns(d.tag) == nazwa]

style = ET.fromstring(archiwum.read("xl/styles.xml"))
czcionki    = dzieci(dziecko(style, "fonts"), "font")      # indeksy od 0
wypelnienia = dzieci(dziecko(style, "fills"), "fill")      # 0 i 1 ZAREZERWOWANE
obramowania = dzieci(dziecko(style, "borders"), "border")
xfs         = dzieci(dziecko(style, "cellXfs"), "xf")      # tu wskazuje s="..."

#   s=N  ->  xf = xfs[N]
#            fontId   -> czcionki[fontId]
#            fillId   -> wypelnienia[fillId]
#            borderId -> obramowania[borderId]
#            numFmtId -> <164 wbudowany, >=164 w <numFmts>

# ==================================================================
# 5. DWA SPOSOBY ZAPISU TEKSTU
# ==================================================================
#   Excel (katalog):     <c r="A1" t="s"><v>0</v></c>
#                        + xl/sharedStrings.xml: <si><t>Region</t></si>
#
#   openpyxl (inline):   <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
#                        + BRAK xl/sharedStrings.xml
#
#   Odczyt inlineStr:
def tekst_inline(c) -> str:
    is_el = dziecko(c, "is")
    if is_el is None:
        return ""
    return "".join((t.text or "") for t in is_el.iter() if bez_ns(t.tag) == "t")

#   Odczyt t="s" (pamiętaj o indeksowaniu od 0 i obsłudze braku katalogu):
#     v = int(dziecko(c, "v").text); tekst = katalog[v]

# ==================================================================
# 6. OSTRZEŻENIA, KTÓRYCH NIE WOLNO WYCISZAĆ
# ==================================================================
import warnings

with warnings.catch_warnings(record=True) as zebrane:
    warnings.simplefilter("always")      # 'ignore' to ANTYWZORZEC
    wb = load_workbook(sciezka)

for w in zebrane:
    print(f"{type(w.message).__name__}: {w.message}")

#   Typowe komunikaty:
#     "Sparkline Group extension is not supported and will be removed"
#     "Slicer List extension is not supported and will be removed"
#     "Timeline Ref extension is not supported and will be removed"
#     "Conditional Formatting extension is not supported ..."  <- wariant x14,
#         NIE zwykłe reguły formatowania warunkowego
#     "Data Validation extension is not supported ..."          <- wariant x14
#
#   Każdy z nich = utrata tej funkcji przy najbliższym zapisie.

# ==================================================================
# 7. NAJCZĘSTSZE MIEJSCA UTRATY (skanuj nazwy części)
# ==================================================================
RYZYKO = {
    "vbaProject.bin": "makra (tylko keep_vba=True, bez edycji)",
    "slicer": "slicery",
    "timeline": "osie czasu",
    "pivotCache": "pamięć podręczna tabeli przestawnej",
    "activeX": "formanty ActiveX",
    "ctrlProps": "właściwości formantów",
    "connections.xml": "Power Query / model danych",
    "queryTables": "Power Query",
    "model/": "model danych Power Pivot",
    "customXml/": "niestandardowe części XML",
}
#   Uwaga: sparkline'y NIE mają osobnej części - siedzą w <extLst> wewnątrz
#   XML arkusza. Sam inwentarz nazw ich nie wykryje. Trzeba zajrzeć do treści.

# ==================================================================
# 8. NARZĘDZIA DO ODCZYTU (przypomnienie)
# ==================================================================
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import (
    coordinate_to_tuple, coordinate_from_string,
    range_boundaries, quote_sheetname,
)
from openpyxl.utils.datetime import from_excel, to_excel

get_column_letter(28)                 # 'AB'         (1-indeksowane)
column_index_from_string("AB")        # 28
coordinate_to_tuple("B2")             # (2, 2)  -> (wiersz, kolumna)
coordinate_from_string("B2")          # ('B', '2')
range_boundaries("A1:C3")             # (1, 1, 3, 3)
quote_sheetname("Dane 2026")          # "'Dane 2026'"
```

## 9. Co dalej

Masz teraz komplet narzędzi do myślenia o pliku `.xlsx` jak o strukturze danych, a nie jak o czarnej skrzynce. Wiesz, co w nim jest, jak jest zapisane i po czym poznać, że openpyxl czegoś nie obsłuży.

W **module 03** zbudujemy z tego **model mentalny openpyxl**. To będzie najważniejszy moduł całego kursu, bo bez niego każdy późniejszy błąd jest zagadką, a z nim — oczywistością. Zobaczymy:

- pełny **graf obiektów**: `Workbook` → `Worksheet` → `Cell` → współdzielone obiekty stylów — i zrozumiemy, dlaczego obiekty stylów w openpyxl są **niemutowalne** (a więc dlaczego `cell.font.bold = True` nie działa — to bezpośrednia konsekwencja tego, że styl to współdzielony wpis w `styles.xml`, który przed chwilą rozebraliśmy na części).
- **cykl życia**: dlaczego `load_workbook` składa cały model w pamięci, dlaczego `save` przepisuje **cały** pakiet od nowa, i dlaczego to jest fundamentalna różnica między „edycją" a „przepisaniem" — analogia fotokopii stanie się wtedy dosłowna, a nie metaforyczna.
- **trzy tryby pracy** openpyxl: normalny, `read_only` i `write_only` — czym się różnią na poziomie dostępu do modelu i dlaczego `write_only` nie pozwoli Ci wrócić do komórki `B7`.
- tabelę **„potrafi / potrafi częściowo / gubi przy zapisie"** z odsyłaczami do konkretnych modułów, żebyś w dowolnym momencie kursu mógł sprawdzić, na jakim gruncie stoisz.

Przygotuj do modułu 03 trzy rzeczy: plik utworzony przez openpyxl (ten z przykładu 1 tego modułu), plik utworzony ręcznie w Excelu oraz kartkę do narysowania grafu obiektów. Będziemy rysować.