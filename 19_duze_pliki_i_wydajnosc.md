Zweryfikuję kluczowe fakty techniczne dla tego modułu: zachowanie trybów `read_only` / `write_only`, wykrywanie `lxml`, `WriteOnlyCell` oraz to, co faktycznie działa (a co nie) w trybie strumieniowym.

Świetnie. Zweryfikuję jeszcze dokładne liczby z dokumentacji wydajności oraz to, jak faktycznie wygląda obsługa `conditional_formatting` i `data_validations` w źródle trybu `write_only`.

Zweryfikuję ostatnie szczegóły: wykrywanie `lxml` w `openpyxl/__init__.py`, zachowanie `read_only` przy `conditional_formatting` oraz czy `WriteOnlyWorksheet` ma natywne atrybuty `conditional_formatting`/`data_validations` w 3.1.x.

Ostatnia weryfikacja dwóch subtelnych punktów, w których brief wymaga ostrożności: zachowanie dostępu losowego w `read_only` oraz czy `WriteOnlyWorksheet` w 3.1.x ma natywne atrybuty `conditional_formatting`/`data_validations`.

# Moduł 19 — Duże pliki i wydajność

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–18

## 0. W tym module nauczysz się

- **Zrozumiesz bilans „czas kontra pamięć"** i to, dlaczego openpyxl w trybie domyślnym potrzebuje ok. **50× rozmiar pliku** w RAM — czyli 2,5 GB na plik 50 MB.
- **Opanujesz dwa tryby optymalizowane**: `read_only=True` (czytanie „przez szczelinę", strona po stronie) i `Workbook(write_only=True)` (prasa drukarska) — wraz z **pełną listą tego, co w nich nie działa**.
- **Nauczysz się mierzyć, a nie zgadywać**: `time.perf_counter`, `tracemalloc`, `cProfile`, `memory_profiler` na realnym pliku 100 000 wierszy.
- **Poznasz dziewięć konkretnych technik przyspieszenia** — od `values_only=True` i `append`, przez rejestr stylów, aż po podział monolitu na pliki miesięczne.
- **Zrozumiesz, kiedy przyspieszać równolegle, a kiedy to nie ma sensu** — i dlaczego `ProcessPoolExecutor` działa na **osobnych plikach**, ale nie na współdzielonym `wb`.
- **Dowiesz się, jak włączyć `lxml`** i dlaczego to jedna linia instalacji daje duży zysk; oraz jak sprawdzić, czy openpyxl faktycznie go używa.
- **Nauczysz się streamować dane z bazy danych prosto do arkusza** przez `yield`, zamiast trzymać milion wierszy w liście.
- **Zapamiętasz, kiedy openpyxl przestaje być właściwym narzędziem** — bo nie wozi się węgla ferrari.

## 1. Intuicja i analogia

### 1.1. Czytanie książki — dwie strategie, dwa koszty

Wyobraź sobie książkę — powiedzmy **„Wojna i pokój"**. Masz ją przeczytać. Masz dwie drogi, a **każda kosztuje coś innego**:

- **Strategia A — rozkładasz całą książkę na kolanach.** Możesz skakać między stronami do woli: sprawdzić przypis na stronie 8, wrócić do rozdziału 3, porównać dwie sceny. Wygodnie. Ale musisz **mieć miejsce na całą książkę** — na stole, w rękach, w głowie. Im grubsza książka, tym większy stół potrzebujesz.
- **Strategia B — czytasz przez szczelinę.** Książka jest przyklejona do ściany, a Ty widzisz jedną stronę naraz przez wąską szczelinę w drzwiach. Otwierasz na pierwszej stronie, czytasz, druga strona się przesuwa, trzecia... **Nie możesz wrócić**. Nie możesz zajrzeć na koniec. Ale **nie potrzebujesz stołu** — potrzebujesz jedynie szczeliny; pamiętasz tylko to, co sam chcesz zapisać na kartce.

To jest **dokładnie wybór**, który daje Ci openpyxl:

| Strategia | Odpowiednik w openpyxl | Zyskujesz | Płacisz |
|---|---|---|---|
| Cała książka na kolanach | `load_workbook()` / `Workbook()` | pełny dostęp losowy, wszystkie funkcje | pamięć ≈ 50× rozmiar pliku |
| Czytanie przez szczelinę | `load_workbook(read_only=True)` | stała, mała pamięć | brak dostępu losowego, brak formatowania |
| Dyktowanie skrybie | `Workbook(write_only=True)` | stała, mała pamięć, dowolnie dużo danych | nie widzisz, co napisałeś; jednorazowy zapis |

Zauważ, proszę, co **nie jest** w tym zestawieniu: **nie ma opcji „wszystko naraz i tanio"**. To jest fundamentalne. W informatyce to się nazywa **trade-off** — kompromis, w którym nie ma darmowego obiadu. Każdy, kto Ci powie, że jego biblioteka „jest po prostu szybka i nie zjada pamięci", pomija połowę zdania. Prawda brzmi:

> **openpyxl w trybie domyślnym jest wygodny i funkcjonalny, ale drogi pamięciowo. Tryby optymalizowane są tanie pamięciowo, ale okrojone funkcjonalnie.**

### 1.2. Dlaczego pamięć, nie czas, jest tu prawdziwym problemem

Wydaje się, że „wolno" to problem. Nie. **Prawdziwym problemem jest pamięć**, i to z dwóch powodów:

1. **Czas można przeczekać.** Skrypt, który liczy 40 sekund zamiast 4, nadal się kończy. Możesz mu dać kolejkę zadań, kawę i poczekać.
2. **Brak pamięci kończy się awarią.** Gdy RAM się skończy, program albo leci na swap (i spowalnia się **setki razy**), albo zostaje zabity przez system (`MemoryError`, albo gorzej — `OOMKilled` w kontenerze, bez ani jednego śladu w logach).

A liczby są brutalne. Dokumentacja openpyxl podaje je wprost:

> *„Memory use is fairly high in comparison with other libraries and applications and is approximately **50 times the original file size**, e.g. **2.5 GB for a 50 MB Excel file**."*

Przeczytaj to jeszcze raz, bo to sedno modułu. **Pięćdziesięciokrotność.** Plik, który na dysku zajmuje 20 MB, w pamięci zajmie **gigabajt**. Plik 100 MB — **pięć gigabajtów**. Jeśli Twój serwer ma 4 GB i piszesz w Django, to nie jest ostrzeżenie; to jest zapowiedź katastrofy.

Analogia, która to tłumaczy: plik `.xlsx` na dysku to **ściśnięta walizka z próżniowym workiem** — ZIP (moduł 02). Gdy openpyxl wczytuje plik, to tak, jakbyś **wyjął wszystko z walizki i rozłożył na podłodze**. XML-owa treść jest gadatliwa (`<row r="150000"><c r="A150000" t="n"><v>42</v></c></row>`), a do tego **każda komórka to osobny obiekt Pythona** z kilkudziesięcioma atrybutami — i ten obiekt ma swój narzut. Dlatego 7 bajtów danych w komórce potrafi kosztować **kilkaset bajtów w RAM**.

### 1.3. Gdzie jest granica — i dlaczego musisz ją znać

Zanim zaczniesz optymalizować, musisz wiedzieć, **na jaką skalę pracujesz**. Poniższa tabela to Twój „przewodnik po gabarytach" — zapamiętaj rząd wielkości, nie dokładne liczby:

| Rozmiar arkusza | Tryb domyślny | Rekomendacja |
|---|---|---|
| do ~5 000 wierszy | bez problemu | tryb domyślny — pełna funkcjonalność |
| 5 000 – 50 000 wierszy | działa, pamięć zauważalna | tryb domyślny, `values_only=True` w pętlach, **`lxml`** |
| 50 000 – 500 000 wierszy | **ryzyko** | `read_only` do odczytu, `write_only` do zapisu |
| 500 000 – 1 048 576 wierszy | **prawie zawsze awaria** | **wyłącznie** tryby optymalizowane |
| powyżej 1 048 576 | **Excel tego nie otworzy** | baza danych / Parquet / CSV (sekcja 2.11) |

Ten **ostatni wiersz** jest twardą granicą samego Excela, nie openpyxl. Arkusz `.xlsx` ma **1 048 576 wierszy × 16 384 kolumn** (XFD). Nie da się tego obejść — to limit formatu. A skoro openpyxl **potrafi** wygenerować plik z większą liczbą wierszy (dokumentacja mówi: *„It is able to export unlimited amount of data (even more than Excel can handle actually)"*), to powstaje pytanie, po co. Odpowiedź: **po nic**. Plik, którego Excel nie potrafi otworzyć, nie jest raportem — jest dowodem na błąd w projekcie.

### 1.4. Nie wozi się węgla ferrari

To analogia, którą zapamiętasz na dłużej niż wszystkie liczby.

Ferrari to **świetne auto**. Szybkie, precyzyjne, dopracowane. Ale gdyby ktoś zaproponował Ci **przewóz tony węgla ferrari**, uznałbyś, że coś z nim nie tak — nie dlatego, że ferrari jest złe, ale dlatego, że **to nie jest zadanie dla ferrari**.

openpyxl jest jak ferrari w świecie Excela: gdy chcesz złożyć **staranny dokument z formatowaniem, wykresami i stylami**, nie ma lepszego narzędzia. Ale gdy chcesz **przetworzyć 20 milionów rekordów i policzyć agregaty**, openpyxl jest **niewłaściwym wyborem** — i żadna optymalizacja tego nie zmieni, bo problem nie leży w optymalizacji, tylko w **doborze narzędzia**.

Praktyczna zasada, którą zapamiętaj i stosuj w projektach:

> **Excel jest warstwą prezentacji, nie magazynem ani silnikiem obliczeniowym.**
>
> Jeśli dane są duże i wielokrotnie przetwarzane — trzymaj je w bazie danych albo w Parquet. Excel dostaje **gotowy, zagregowany wynik** do obejrzenia i wysłania dalej.

Innymi słowy: **agreguj w bazie, prezentuj w Excelu**. To nie jest ograniczenie openpyxl — to jest dobra architektura, którą rozwiniesz w module 26 (wzorzec „raport jako projekcja").

### 1.5. Prasa drukarska — jak myśleć o `write_only`

Do trybu `read_only` wystarczy obraz ze szczeliną z 1.1. Ale `write_only` zasługuje na własną analogię, bo jest **znacznie bardziej restrykcyjny**.

Wyobraź sobie **prasę drukarską**. Podajesz jej kartki jedna po drugiej — z prędkością, z jaką chcesz. Prasa drukuje każdą i **odkłada na stos, którego Ty nie widzisz**. Nie możesz wrócić do kartki piątej i coś na niej dopisać, bo kartka piąta jest **już wydrukowana i leży w zamkniętym stosie**. Nie możesz też podejrzeć, co jest na kartce, którą właśnie wsunąłeś — prasa po prostu ją przyjęła i poszła dalej.

Co z tego wynika:

1. **Można drukować nieskończenie dużo** — bo stos rośnie na dysku, nie w pamięci. Pamięć zajęta to jedna kartka naraz. Dokumentacja podaje: **poniżej 10 MB** niezależnie od liczby wierszy.
2. **Nie ma dostępu losowego.** Zero. `ws["A1"]`, `ws.cell(...)`, `ws.iter_rows()` — nic z tego nie istnieje na `WriteOnlyWorksheet`.
3. **Zapis jest jednorazowy.** Gdy raz puścisz `wb.save()`, stos jest zabindowany i **nie dopiszesz już ani jednej kartki**. Próba kończy się `WorkbookAlreadySaved`.
4. **To, co ma trafić PRZED kartkami, musi być ustawione PRZED pierwszym `append`.** Bo prasa **wypisuje nagłówek książki zanim zacznie drukować strony**. Taki `freeze_panes` (moduł 11) musisz ustawić **na początku**, nie na końcu.

To ostatnie to najczęstsze źródło frustracji w `write_only` — i dlatego wrócimy do niego w 2.4.

### 1.6. Trzy kategorie prawdy dla tego modułu

Zgodnie z obietnicą z modułu 03 — tabela, wokół której zbudowany jest cały moduł:

| Kategoria | Przykłady w kontekście „dużych plików" | Twoja reakcja |
|---|---|---|
| **openpyxl potrafi** | `iter_rows(values_only=True)`, `append`, `write_only` dla dowolnej liczby wierszy, `lxml`, `ProcessPoolExecutor` na osobnych plikach, liczenie sum w Pythonie podczas strumieniowego czytania | używaj śmiało. To ten dobry świat. |
| **openpyxl potrafi częściowo** | `read_only` (tylko wiersze, tylko w przód, wymiary bywają złe), `write_only` (tylko `append`, jednorazowy zapis), formatowanie warunkowe i walidacja w `write_only` (działa, ale ścieżka jest węższa — sprawdź w swojej wersji) | **przetestuj na swoim pliku**, zanim zbudujesz na tym produkcję |
| **openpyxl gubi/nie ma** | w `read_only`: wykresy, obrazy, komentarze, scalenia, `iter_cols`, `column_dimensions`; w `write_only`: tabele (`Table`), odczyt i modyfikacja, drugi zapis; brak `lxml` = brak przyspieszenia | nie zakładaj, że „jakoś działa". Sprawdź albo obejdź. |

## 2. Teoria

### 2.1. Ile dokładnie kosztuje komórka?

Warto raz zobaczyć, **skąd** bierze się ta pięćdziesięciokrotność, bo dzięki temu przestaniesz się dziwić.

W trybie domyślnym **każda** komórka, do której sięgniesz, staje się obiektem `openpyxl.cell.cell.Cell`. Ten obiekt trzyma:

- wartość (`_value`) i typ danych (`data_type`),
- koordynaty: `row`, `column`, `coordinate` (`"BC1234"`),
- referencję do arkusza (`parent`),
- `_style` — **indeks stylu** (nie obiekt stylu — to ważne, moduł 09),
- hiperlink, komentarz (jeśli są),
- flagi: `_comment`, `_hyperlink`, `_value_is_date`…

Do tego dochodzi słownik `_cells: dict[tuple[int, int], Cell]` arkusza, obiekty `DimensionHolder`, `MultiCellRange`, nazwy, style w `styles.xml`… To wszystko **rośnie liniowo z liczbą komórek**, które dotknąłeś.

I tu jest pułapka, którą trzeba nazwać wprost, bo wraca w tym module jak bumerang:

> **Samo „zajrzenie" do komórki ją tworzy.**

W 3.1 w `Worksheet.cell()` jest jawne sprawdzenie: jeśli komórka nie istnieje, **tworzy ją pusta** i zapamiętuje. Dlatego pętla typu:

```python
for row in range(1, 100_000):
    for col in range(1, 100):
        ws.cell(row=row, column=col)   # ⚠️ nawet bez przypisania wartości!
```

alokuje **10 milionów obiektów** `Cell`, których w większości nie potrzebujesz. To dlatego dokumentacja ostrzega:

> *„Because of this feature, scrolling through cells instead of accessing them directly will create them all in memory, even if you don't assign them a value."*

**Pierwsza technika przyspieszenia** (z listy w 2.8) rodzi się właśnie z tego zdania: **`iter_rows(values_only=True)` nie tworzy obiektów `Cell` wcale.** Zwraca krotki wartości. Koniec. To nie jest mikrooptymalizacja — to różnica rzędu wielkości.

### 2.2. Tryb domyślny — co daje, za co płaci

**Daje:** pełny dostęp losowy. `ws["BC1234"]`, skakanie po arkuszach, modyfikowanie komórek w dowolnej kolejności, wykresy, obrazy, tabele, formatowanie warunkowe, walidacja, scalenia, ochrona, `copy_worksheet`, `insert_rows`, `merge_cells`. Wszystko z modułów 00–18.

**Płaci:** pamięć. Model **całego** skoroszytu siedzi w RAM, aż do `wb.save()`.

```python
from openpyxl import Workbook

wb = Workbook()                     # model w pamięci, jeszcze pusty
ws = wb.active
ws.title = "Dane"

# Kazde append tworzy obiekty Cell i trzyma je w _cells
for i in range(1, 100_001):
    ws.append([i, f"Klient {i}", i * 1.5])

wb.save("output/normalny.xlsx")     # tutaj caly model idzie na dysk
wb.close()
```

**Co się dzieje w pamięci.** Po pętli `ws._cells` ma **300 000 obiektów `Cell`** (100 000 wierszy × 3 kolumny). Każdy z nich zajmuje od kilkudziesięciu do kilkuset bajtów. To jest te 50× — komórka 7-bajtowej wartości waży w RAM kilkaset bajtów.

**Co trafi do pliku.** Wszystkie 300 000 komórek, ale **spakowane** — ZIP z XML-em. Plik na dysku będzie **kilkaset razy mniejszy** niż zajętość RAM-u.

**Kiedy używać trybu domyślnego:** zawsze, gdy **potrzebujesz czegokolwiek poza surowymi danymi** — a więc gdy piszesz raport z formatowaniem, wykresami, tabelami (moduły 09–17). Dla raportu 5 000 wierszy z wykresem i formatowaniem warunkowym **tryb domyślny jest właściwym wyborem** i żadne „optymalizacje" nie są potrzebne.

### 2.3. `read_only=True` — czytanie przez szczelinę

Włączasz go jedną flagą:

```python
from openpyxl import load_workbook

wb = load_workbook("big_data.xlsx", read_only=True)
try:
    ws = wb["Dane"]
    for row in ws.iter_rows(values_only=True):
        ...
finally:
    wb.close()          # ⚠️ OBOWIĄZKOWE
```

**Co się dzieje w pamięci.** openpyxl **nie parsuje całego arkusza**. Otwiera plik ZIP i **celowo zostawia go otwartego**. Gdy prosisz o kolejny wiersz, parser czyta z XML-a **tylko tyle, ile trzeba**, buduje krotkę i zapomina poprzednią. Pamięć jest **stała**, niezależna od liczby wierszy.

Dokumentacja openpyxl opisuje to tak: *„a read-only workbook will use lazy loading. The workbook must be explicitly closed with the `close()` method."*

I jeszcze jedno zdanie z dokumentacji wydajności, które jest tu kluczowe i które wyjaśnia, dlaczego ten tryb jest tak dobry w automatyzacji:

> *„openpyxl's read-only mode opens a workbook almost immediately making it suitable for multiple processes, this also reduces memory use significantly."*

„Otwiera skoroszyt **prawie natychmiast**" — bo nie ma czego parsować na starcie. To nie jest drobiazg. To jest fundament równoległości (2.10).

#### 2.3.1. Co konkretnie **nie działa** w `read_only`

Ta lista jest długa i warto ją znać **całą**, bo połowa „tajemniczych błędów" z openpyxl to po prostu próba użycia zwykłego API w trybie optymalizowanym.

**Nie ma tych metod i właściwości w ogóle** (dostaniesz `AttributeError`):

| Brak w `read_only` | Dlaczego |
|---|---|
| `ws.iter_cols()`, `ws.columns` | wymagają dostępu losowego do całej siatki; strumień tego nie da. Dokumentacja: *„For performance reasons the `Worksheet.iter_cols()` method is not available in read-only mode."* |
| `ws.merged_cells` | scalenia są zapisane **na końcu** pliku, po danych |
| `ws.column_dimensions`, `ws.row_dimensions` | parsowane w innym miejscu pliku |
| `ws.page_setup`, `ws.page_margins`, `ws.print_options` | to metadane druku, ładowane tylko w pełnym modelu |
| `ws.protection`, `ws.sheet_properties`, `ws.views` | jak wyżej — dostępne w pełnym parserze |

**Tego nie rób, bo się wywali:**

| Próba | Efekt |
|---|---|
| `ws["A1"] = 42` | `TypeError: 'ReadOnlyWorksheet' object does not support item assignment` |
| `wb.save(...)` | `TypeError: Workbook is read-only` |
| `wb.copy_worksheet(ws)` | nie działa — brak modelu do skopiowania |
| `cell.value = ...` | `AttributeError: Cell is read only` |

**Tego nie ma w komórkach** (atrybuty zwracają `None` albo nie istnieją):

| Atrybut | Uwaga |
|---|---|
| `cell.font`, `cell.border`, `cell.fill`, `cell.alignment` | **nie są ładowane**. To konsekwencja tego, że `read_only` czyta tylko wartości. |
| `cell.comment` | brak |
| `cell.hyperlink` | brak |
| `cell.number_format` | nie w tej ścieżce |

**Uwaga o `ws["A1"]` — czytanie.** Teoretycznie `ReadOnlyWorksheet` **ma** zmapowane `__getitem__` i `cell`, ale w praktyce jest to **pułapka**: implementacja `_get_cell` skanuje strumień od początku, więc jednorazowe `ws["A1"]` może wymagać przejścia przez tysiące wierszy, a przy **nieskalowanym arkuszu** dostaniesz `ValueError: Worksheet is unsized`. Zasada praktyczna brzmi:

> **W `read_only` NIE MA dostępu losowego. Czytasz wierszami, w przód, raz. Jeśli potrzebujesz czegoś z „środka" — najpierw zbierz to podczas przejścia.**

#### 2.3.2. Wymiary mogą kłamać — `calculate_dimension` i `reset_dimensions`

To drugi „klasyk" tego trybu i trzeba go zrozumieć, nie tylko zapamiętać.

`read_only` **nie czyta całego arkusza**, więc nie wie, ile ma wierszy. Wierzy w to, co mu powie **metadana w pliku** (element `<dimension>`). A metadane bywają błędne — dokumentacja mówi to wprost:

> *„Read-only mode relies on applications and libraries that created the file providing correct information about the worksheets, specifically the used part of it, known as the dimensions. Some applications set this incorrectly."*

Objaw: `ws.max_row` jest `None` albo `ws.calculate_dimension()` zwraca `"A1:A1"`, mimo że w pliku jest 200 000 wierszy.

Lekarstwo:

```python
ws = wb["Dane"]

# (a) Sprawdz, co biblioteka sadzi o rozmiarze
print("max_row =", ws.max_row, " max_column =", ws.max_column)

# (b) Forsuj skanowanie, jesli metadane sa zle
#     Uwaga: to JEDNORAZOWE przejscie przez caly arkusz i ustawienie wymiarow.
if ws.max_row is None or ws.max_column is None:
    print("Wymiary nieustalone - forsujemy skan...")
    print("calculate_dimension(force=True) =", ws.calculate_dimension(force=True))

# (c) Alternatywa: wyzeruj wymiary i pozwól im sie policzyc przy iteracji
# ws.reset_dimensions()
```

**Zasada, którą zapamiętaj:** `calculate_dimension()` **bez** `force=True` rzuci `ValueError("Worksheet is unsized, use calculate_dimension(force=True)")`, jeśli wymiary są nieznane. A `force=True` zrobi pełne przejście przez arkusz — **raz**. Po tym `ws.max_row` i `ws.max_column` będą już ustawione.

I jeszcze pułapka, która jest subtelna: **`reset_dimensions()` usuwa** wymiary, nie liczy ich. Po wywołaniu `ws.max_row` znów będzie `None`, a openpyxl policzy je przy okazji iteracji.

#### 2.3.3. `EmptyCell` — komórka, która nie istnieje

W `read_only` czytasz **strumień**, a w pliku XML komórki mogą być **pominięte** (puste komórki nie są zapisywane). openpyxl wypełnia te luki „wypełniaczem", ale to **nie jest** prawdziwa komórka:

```python
# ReadOnlyWorksheet._cells_by_row
filler = EMPTY_CELL
if values_only:
    filler = None
```

Konsekwencja, która wywala kod:

```python
for row in ws.iter_rows():
    for cell in row:
        print(cell.coordinate)   # ⚠️ EmptyCell nie ma .coordinate!
```

`EmptyCell` ma `value = None`, ale **nie ma** `coordinate`, `row` ani `column`. Kod, który buduje adresy z komórek, **pęknie przy pierwszej luce**.

**Rozwiązanie:** użyj `values_only=True` (dostajesz `None` w krotce, co jest naturalne) albo **licz numer wiersza i kolumny sam** przez `enumerate`. To drugie rozwiązanie jest tym, które stosują profesjonaliści:

```python
for row_idx, row in enumerate(ws.iter_rows(), start=1):
    for col_idx, cell in enumerate(row, start=1):
        ...

# Albo z values_only - najprosciej i najszybciej:
for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
    for col_idx, value in enumerate(row, start=1):
        if value is not None:
            ...
```

Zdanie z praktyki branżowej, warte zapamiętania: *„Track row and column numbers yourself with `enumerate` (starting at the `min_row` you asked for), or use `values_only=True` and ignore cell objects altogether."*

### 2.4. `write_only=True` — prasa drukarska

Włączasz go w konstruktorze:

```python
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

wb = Workbook(write_only=True)
ws = wb.create_sheet(title="Dane")     # ⚠️ NIE ma arkusza domyslnego!
```

**Co się dzieje w pamięci.** `WriteOnlyWorksheet` trzyma **tylko bufor na jeden wiersz**. Gdy wołasz `ws.append(...)`, wiersz jest **natychmiast serializowany do XML-a i wypychany** do archiwum. Nie zostaje żadna referencja. Dlatego pamięć jest stała — dokumentacja podaje: **poniżej 10 MB**.

**Co trafi do pliku.** Dokładnie to, co wypchnąłeś przez `append`, w kolejności wypchnięcia — i **nic więcej**, bo nie ma czego odtworzyć.

#### 2.4.1. Pięć twardych zasad trybu `write_only`

Wszystkie pochodzą wprost z dokumentacji i wszystkie są **bez wyjątków**:

**1. Nowy skoroszyt `write_only` nie ma żadnego arkusza.** Musisz go jawnie utworzyć:

```python
wb = Workbook(write_only=True)
print(wb.sheetnames)        # []  <- pusto!
ws = wb.create_sheet(title="Dane")
```

Dlaczego: bo „domyślny arkusz" musiałby powstać w pamięci, a nikt go nie potrzebuje. Pamiętaj, że **`wb.active`** w `write_only` nie działa tak jak w normalnym trybie — nie ma czego „aktywować".

**2. Wiersze dodaje się WYŁĄCZNIE przez `append()`.** Dokumentacja: *„In a write-only workbook, rows can only be added with `append()`. It is not possible to write (or read) cells at arbitrary locations with `cell()` or `iter_rows()`."*

Czyli **nie ma** `ws["A1"] = 1`, **nie ma** `ws.cell(row=1, column=1, value=1)`, **nie ma** `ws["A1"].font = Font(...)`. Jest tylko `ws.append([...])`.

**3. Skoroszyt można zapisać TYLKO RAZ.** Dokumentacja: *„A write-only workbook can only be saved once. After that, every attempt to save the workbook or append() to an existing worksheet will raise an `openpyxl.utils.exceptions.WorkbookAlreadySaved` exception."*

Co to znaczy w praktyce: **nie możesz zrobić `wb.save()`, potem dopisać wiersz i znów `wb.save()`.** Pamięć przepływu jest jednokierunkowa, jak prasa. Jeśli potrzebujesz dwóch wersji pliku — wygeneruj dwa skoroszyty.

**4. Wszystko, co ma się znaleźć PRZED danymi, musi być ustawione PRZED pierwszym `append`.** Cytat z dokumentacji: *„Everything that appears in the file before the actual cell data must be created before cells are added because it must written to the file before then. For example, `freeze_panes` should be set before cells are added."*

To jest **najczęstsza przyczyna „to ustawienie nie działa"** w `write_only`. Wynika to z formatu: w XML-u arkusza elementy takie jak `sheetViews` (freeze), `cols` (szerokości kolumn), `sheetFormatPr` są **przed** `sheetData` (danymi). Skoro openpyxl pisze strumieniowo, to gdy strumień już minął ten punkt, nie ma jak wrócić.

```python
wb = Workbook(write_only=True)
ws = wb.create_sheet(title="Dane")

# ✅ PRZED danymi - dziala:
ws.freeze_panes = "A2"
ws.sheet_view.showGridLines = False
ws.column_dimensions["B"].width = 25

# ... a teraz dopiero dane:
ws.append(["ID", "Nazwa"])
for i in range(1, 1001):
    ws.append([i, f"Klient {i}"])

# ❌ PO danych - NIE zadziala (zostanie zignorowane lub rzuci blad):
# ws.freeze_panes = "A2"
```

**5. Zamykanie.** `wb.close()` działa w trybie `write_only` (docstring z modułu 18: *„Only affects read-only and write-only modes"*). Dobry nawyk — trzymaj go w `finally`.

#### 2.4.2. `WriteOnlyCell` — jedyny sposób na format w strumieniu

Skoro `append` bierze **wartości**, a nie obiekty komórek, to jak nadać format? Odpowiedź: **nie bierzesz wartości, tylko jawnie tworzysz komórkę**:

```python
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment

wb = Workbook(write_only=True)
ws = wb.create_sheet(title="Dane")

naglowek = WriteOnlyCell(ws, value="Sprzedaż")
naglowek.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
naglowek.fill = PatternFill("solid", fgColor="1F4E78")
naglowek.alignment = Alignment(horizontal="center")

ws.append([naglowek, "Luty", 1200.50])   # mozesz mieszac WriteOnlyCell z wartosciami
```

I teraz **najważniejsza rzecz o `WriteOnlyCell`**, która jest źródłem wielu godzin debugowania:

> **`WriteOnlyCell` to jednorazowy „przekaźnik", nie obiekt do trzymania.** Cykl życia jest taki: tworzysz komórkę → wywołujesz `append` → **openpyxl natychmiast wypisuje ją do XML-a i porzuca**. W źródle write-only writer widać to czarno na białym: po zapisaniu ostylowanej komórki tworzona jest **nowa** pusta `WriteOnlyCell(self)` — bo stara nie ma już znaczenia.

Konsekwencje, które trzeba przyswoić:

1. **Nie trzymaj referencji i nie modyfikuj po `append`.** To, co zrobiłeś po `append`, nie trafi do pliku.
2. **Nie reużywaj tego samego obiektu.** Nowy wiersz = nowa `WriteOnlyCell`.
3. **Styl nadaj PRZED `append`.** Po nim jest już za późno.
4. **Ale: możesz trzymać i reużywać obiekt `Font`/`PatternFill`.** To obiekty stylu (niemutowalne, moduł 09) — one nie są „zużywane". I to jest dokładnie miejsce, gdzie wchodzi technika „styl z rejestru" (2.8, punkt 5).

Trik, który drastycznie zmniejsza alokacje w `write_only`: **twórz `Font`/`Fill`/`Alignment` raz, poza pętlą**, a w pętli tylko „przypinaj" je do `WriteOnlyCell`. Nowa komórka, ten sam styl.

### 2.5. Co działa w którym trybie — tabela rozstrzygająca

To jest tabela, do której będziesz wracał. Zapamiętaj ją — oszczędzi Ci dziesiątki minut debugowania.

| Możliwość | Normalny | `read_only` | `write_only` |
|---|---|---|---|
| `ws["A1"]`, `ws.cell()` | ✅ | ❌ | ❌ |
| `ws.iter_rows()` | ✅ | ✅ (tylko w przód) | ❌ |
| `ws.iter_rows(values_only=True)` | ✅ | ✅ **najszybsza ścieżka** | ❌ |
| `ws.iter_cols()`, `ws.columns` | ✅ | ❌ | ❌ |
| `ws.append()` | ✅ | ❌ | ✅ **jedyny sposób zapisu** |
| `ws.max_row` / `max_column` | ✅ | ⚠️ może być `None` | ❌ |
| `ws.calculate_dimension(force=True)` | ✅ | ✅ | ❌ |
| Odczyty wartości | ✅ | ✅ | ❌ |
| Modyfikacja istniejącej wartości | ✅ | ❌ | ❌ |
| `cell.font`, `cell.fill`, `cell.border` (odczyt) | ✅ | ❌ | ❌ |
| `WriteOnlyCell(ws, value=...)` do zapisu ze stylem | ✅ | ❌ | ✅ |
| `cell.number_format` | ✅ | ❌ | ✅ (przez `WriteOnlyCell`) |
| Wykresy, obrazy, komentarze, hiperlinki | ✅ | ❌ | ⚠️ komentarze tak (`WriteOnlyCell`); reszta nie |
| Scalanie, freeze panes, szerokości kolumn | ✅ (w dowolnym momencie) | ❌ | ⚠️ **tylko PRZED danymi** |
| Tabele (`Table`) | ✅ | ✅ (odczyt `ws.tables`?) | ❌ **nieobsługiwane — trzeba samemu** |
| Formatowanie warunkowe | ✅ | ⚠️ bardzo wolne (jest na końcu pliku) | ⚠️ technicznie tak — patrz 2.6 |
| Walidacja danych | ✅ | ❌ | ⚠️ technicznie tak — patrz 2.6 |
| Ochrona arkusza / skoroszytu | ✅ | ❌ | ⚠️ ograniczona |
| `copy_worksheet()` | ✅ | ❌ | ❌ |
| `wb.save()` | ✅ | ❌ (`TypeError`) | ✅ **tylko raz** |
| Stała pamięć | ❌ | ✅ | ✅ |

Legenda: ✅ działa · ❌ nie działa · ⚠️ działa częściowo lub z zastrzeżeniem.

### 2.6. Formatowanie warunkowe i walidacja w `write_only` — ostrożnie

To punkt, w którym brief prosi o ostrożność — i słusznie, bo sytuacja jest **naprawdę subtelna**. Wyjaśnię ją od podstaw, bez upraszczania.

**Dlaczego to w ogóle może działać?** Bo w XML-u arkusza kolejność elementów jest **ściśle określona** i wygląda tak (w uproszczeniu):

```xml
<worksheet>
  <sheetPr/>            <!-- wlasciwosci arkusza      -->
  <dimension/>
  <sheetViews/>         <!-- freeze panes             -->
  <sheetFormatPr/>
  <cols/>               <!-- szerokosci kolumn        -->
  <sheetData>           <!-- ⬅ DANE (tu pisze strumien) -->
    <row>...</row>
  </sheetData>
  <conditionalFormatting/>   <!-- ⬅ PO danych!         -->
  <dataValidations/>         <!-- ⬅ PO danych!         -->
  <hyperlinks/>
  <mergeCells/>
  ...
</worksheet>
```

Zauważ: **formatowanie warunkowe i walidacja leżą PO `sheetData`, nie przed.** A to znaczy, że autor openpyxl (Charlie Clark) mógł o nich powiedzieć:

> *„When it comes to write-only mode, pretty much anything that is stored in the XML **after the cell data** should be possible. Conversely, things like column formats have to be written **before** cell data and so are more difficult to deal."*

Czyli **odwrotnie niż `freeze_panes`**: formatowanie warunkowe i walidacja wpisują się w filozofię strumienia, bo przychodzą **na końcu**. To dlatego writer `write_only` ma ten fragment:

```python
# openpyxl/worksheet/_write_only.py (fragment zapisu)
if self.sort_state.ref:
    xf.write(self.sort_state.to_tree())

if self.conditional_formatting:
    cfs = write_conditional_formatting(self)
    for cf in cfs:
        xf.write(cf)

if self.data_validations.count:
    xf.write(self.data_validations.to_tree())
```

**Ale problem jest gdzie indziej.** Sam writer **umie zapisać** te rzeczy — pytanie brzmi, czy `WriteOnlyWorksheet` **wystawia** atrybuty `conditional_formatting` i `data_validations`, do których sięgasz tytułem `ws.conditional_formatting.add(...)`. I tu jest właśnie miejsce, gdzie trzeba być ostrożnym.

W historii openpyxl bywało tak, że `WriteOnlyWorksheet` **nie miał** tych atrybutów i próba użycia kończyła się komunikatem w rodzaju *„WriteOnlyWorksheet does not allow such definition"*. Społeczność znalazła na to obejście — **jawne utworzenie brakującego atrybutu**:

```python
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidationList
from openpyxl.formatting.formatting import ConditionalFormattingList

wb = openpyxl.Workbook(write_only=True)
ws = wb.create_sheet(title="Dane")

# Obejscie dla wersji, ktore nie maja tych atrybutow na WriteOnlyWorksheet:
if not hasattr(ws, "data_validations"):
    ws.data_validations = DataValidationList()
if not hasattr(ws, "conditional_formatting"):
    ws.conditional_formatting = ConditionalFormattingList()
```

**Jak to ugryźć w 3.1.x — praktyczna instrukcja płynowa:**

1. **Najpierw sprawdź, czy Twoja wersja ma te atrybuty:**

   ```python
   from openpyxl import Workbook
   wb = Workbook(write_only=True)
   ws = wb.create_sheet(title="T")
   print("data_validations:", hasattr(ws, "data_validations"))
   print("conditional_formatting:", hasattr(ws, "conditional_formatting"))
   ```

2. **Jeśli `True`** — używaj normalnego API (`ws.conditional_formatting.add(...)`, `ws.add_data_validation(...)`). W 3.x wiele typów arkuszy zostało **zharmonizowanych**, więc prawdopodobieństwo, że działa, jest wysokie.

3. **Jeśli `False`** — użyj jawnego przypisania list powyżej i **przetestuj plik w Excelu** (obowiązkowa weryfikacja „na oko" z modułu 18).

4. **Zawsze** — sprawdź w dokumentacji swojej wersji i **zapisz wynik testu** w komentarzu w kodzie. Nie zgaduj.

**Czego natomiast `write_only` NIE obsługuje i nie obsłuży** (to nie jest kwestia wersji):

- **Tabele (`Table` / `ListObject`).** Cytat od autora: *„if you want to use tables then you'll need to do some work yourself that openpyxl normally does."* Tabela wymaga metadanych, których strumień nie ma jak wytworzyć. Obejście: tworzysz zwykły zakres danych, a tabelę zakładasz **osobnym przebiegiem** — czyli otwierasz plik normalnie, dodajesz tabelę, zapisujesz (i płacisz pamięcią tylko za ten etap).
- **Odczyt czegokolwiek.** Z definicji.
- **Drugi `save()`.** `WorkbookAlreadySaved`.
- **`merge_cells`, `insert_rows`, modyfikacja po fakcie.** Nie ma czego modyfikować.

**Wniosek, który zapamiętaj:** `write_only` jest **wystarczający dla 90% raportów „dużych i płaskich"** — dużo wierszy, proste nagłówki, kilka formatów, może walidacja i pasek danych. Gdy potrzebujesz tabel, wykresów, scalonych nagłówków czy obrazów — **nie kombinuj ze strumieniem**, tylko przemyśl architekturę (moduł 26: warstwa prezentacji osobno).

### 2.7. `lxml` — darmowe przyspieszenie, które trzeba włączyć świadomie

To najbardziej „niewidzialna" optymalizacja w całym module, bo **sprowadza się do jednej instalacji**:

```bash
pip install lxml
```

**Jak to działa od środka.** openpyxl nie ma własnego parsera XML. Wybiera między dwiema implementacjami:

- **`xml.etree.ElementTree`** — biblioteka standardowa Pythona. Zawsze dostępna, ale **czysty Python** w dużej części.
- **`lxml`** — wiązania do **libxml2** napisanego w C. Znacznie szybszy.

W `openpyxl/xml/functions.py` wybór wygląda tak:

```python
from openpyxl import DEFUSEDXML, LXML

if LXML is True:
    from lxml.etree import (
        Element, SubElement, register_namespace, QName, xmlfile, XMLParser,
    )
    from lxml.etree import fromstring, tostring
    # bezpieczeństwo: nie rozwiazuj encji zewnetrznych
    safe_parser = XMLParser(resolve_entities=False)
    fromstring = partial(fromstring, parser=safe_parser)
else:
    from xml.etree.ElementTree import (
        Element, SubElement, fromstring, tostring, QName, register_namespace,
    )
    from et_xmlfile import xmlfile
    if DEFUSEDXML is True:
        from defusedxml.ElementTree import ...
```

**Dwie rzeczy warte uwagi w tym kodzie:**

1. **`LXML` to flaga, nie import.** openpyxl decyduje **raz, przy imporcie**, i od tego momentu cały pakiet używa wybranego backendu. Dlatego **nie da się przełączyć w trakcie działania programu**.
2. **`resolve_entities=False`** — to zabezpieczenie przed atakiem typu „billion laughs" / XXE. Wrócimy do tego w module 21, ale zapamiętaj: **używanie `lxml` przez openpyxl jest bezpieczne**, bo openpyxl jawnie wyłącza rozwiązywanie encji.

**Jak sprawdzić, czy openpyxl faktycznie używa `lxml`:**

```python
from openpyxl import LXML, DEFUSEDXML, NUMPY, PANDAS
import openpyxl

print("openpyxl :", openpyxl.__version__)
print("LXML     :", LXML)          # True = uzywamy lxml
print("DEFUSEDXML:", DEFUSEDXML)   # True = chronieni przez defusedxml
print("NUMPY    :", NUMPY)
print("PANDAS   :", PANDAS)

# Co DOKLADNIE jest pod "Element" - rozstrzygnie spór:
from openpyxl.xml.functions import Element
print("Element  :", Element)
print("Backend  :", Element.__module__)   # 'lxml.etree' albo 'xml.etree.ElementTree'
```

**Ale uwaga — flaga `LXML` nie zależy tylko od tego, czy pakiet jest zainstalowany.** W źródle openpyxl decyzja jest koniunkcją dwóch warunków:

```python
LXML = lxml_available() and lxml_env_set()
```

- **`lxml_available()`** — sprawdza, czy `lxml` da się zaimportować **oraz** czy wersja jest dostatecznie nowa. Konkretnie openpyxl wymaga **`LXML_VERSION >= (3, 3, 1, 0)`**. Starsza wersja dostanie ostrzeżenie:
  > *„The installed version of lxml is too old to be used with openpyxl"* — i backend pozostanie `etree`.
- **`lxml_env_set()`** — sprawdza **zmienną środowiskową `OPENPYXL_LXML`**. Jeśli ustawisz ją na cokolwiek innego niż `True`, openpyxl **wyłączy `lxml`** nawet gdy jest zainstalowany.

To ostatnie jest ważne w dwie strony:

- **Praktyczna korzyść:** jeśli `lxml` sprawia Ci problem (np. w testach z `pyfakefs`, albo przy nietypowych plikach), możesz go **wyłączyć jedną zmienną**: `OPENPYXL_LXML=False` — bez odinstalowywania czegokolwiek. To bardzo wygodne narzędzie diagnostyczne („Czy to `lxml` psuje mój plik?").
- **Pułapka:** jeśli w Twoim środowisku (CI, kontener, orkiestrator) ktoś ustawił tę zmienną globalnie, przyspieszenie może być **nieme**, a Ty się dziwisz, dlaczego „na serwerze jest wolniej". Diagnostyka: wypisz `openpyxl.LXML` i `os.environ.get("OPENPYXL_LXML")` w logu startowym.

**Sprawdzenie, jak `lxml` przekłada się na zapis.** W `openpyxl/cell/_writer.py` jest analogiczny wybór, tym razem **funkcji zapisującej komórkę**:

```python
def lxml_write_cell(xf, worksheet, cell, styled=False):
    value, attributes = _set_attributes(cell, styled)
    ...

if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell
```

Czyli `lxml` przyspiesza **zarówno czytanie, jak i pisanie** — i to właśnie dlatego dokumentacja przy trybie `write_only` mówi wprost:

> *„When you want to dump large amounts of data make sure you have lxml installed."*

**Realny zysk.** Dokumentacja wydajności podaje porównanie (dla 1 000 wierszy × 50 kolumn, Python 3.8): openpyxl **1,10 s** → openpyxl z trybem optymalizowanym **0,57 s**. To prawie dwukrotnie. Przy setkach tysięcy wierszy ten sam mechanizm daje zysk, który czujesz w minutach, nie milisekundach.

#### 2.7.1. `lxml` w Dockerze

To pytanie wraca w każdym projekcie produkcyjnym, bo **`pip install lxml` w kontenerze wymaga czasem kompilatora**. Dwa podejścia:

**Podejście A — koła binarne (zalecane, najprostsze).** Współczesne `lxml` publikuje „wheele" (`manylinux`) na PyPI. Instalacja przez `pip install lxml` **pobiera gotowy binarny pakiet**, bez kompilacji:

```dockerfile
FROM python:3.12-slim

WORKDIR /app

# Bez apt-get install gcc! Wheele lxml dzialaja na manylinux.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["python", "-m", "raporty", "--config", "config.yaml"]
```

gdzie `requirements.txt` to:

```text
openpyxl==3.1.5
lxml>=4.9
pillow>=10
```

**Podejście B — obraz z kompilatorem (gdy wheeli brak dla Twojej architektury).** Na `arm64` albo dla starych wersji Pythona `lxml` może wymagać kompilacji — wtedy potrzebne są biblioteki systemowe:

```dockerfile
FROM python:3.12-slim

# tylko jesli musimy kompilowac lxml (np. architektura bez wheeli)
RUN apt-get update && apt-get install -y --no-install-recommends \
        gcc libxml2-dev libxslt1-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
```

**Zasada:** nie dodawaj `gcc` i bibliotek „na wszelki wypadek". **Sprawdź, czy `pip install lxml` przechodzi bez nich** — na `x86_64` z niemal zawsze przechodzi, a każdy zbędny pakiet w obrazie to większa powierzchnia ataku (moduł 21) i większy obraz.

**Weryfikacja w kontenerze** — jeden krok w CI, który ratuje od „dlaczego u mnie szybciej":

```bash
docker run --rm my-image python -c "from openpyxl import LXML; assert LXML, 'lxml NIE aktywny!'; print('lxml aktywny')"
```

Ten test w CI jest wart więcej niż wszystkie optymalizacje razem, bo chroni przed **cichą regresją wydajności** — gdy obraz zmieni się tak, że `lxml` przestanie być aktywny, test się wywali, zamiast pozwolić, żeby raporty nagle robiły się trzy razy wolniej bez śladu w logach.

### 2.8. Dziewięć technik przyspieszenia

Poniżej lista **konkretnych, mierzalnych** technik — uporządkowana od najbardziej opłacalnych.

**1. `iter_rows(values_only=True)` zamiast `ws.cell(...)` w pętli.**
Największy zysk pojedynczej zmiany. Nie tworzysz obiektów `Cell` — dostajesz krotki wartości. `ws.values` to skrót do tego samego.

```python
# ❌ wolno: 100k * 10 = 1 000 000 obiektow Cell w RAM
total = 0
for row in range(1, ws.max_row + 1):
    for col in range(1, 11):
        v = ws.cell(row=row, column=col).value

# ✅ szybko: brak obiektow Cell, brak alokacji
total = 0
for row in ws.iter_rows(values_only=True):
    for v in row:
        ...
```

**2. `append()` zamiast przypisywania komórek po jednej.**
W trybie normalnym `append` jest szybszy niż `ws.cell(r, c, value=...)` w podwójnej pętli, bo wewnętrznie operuje na wierszu jako całości.

**3. `write_only` przy generowaniu od zera.**
Gdy piszesz **nowy** plik i nie potrzebujesz niczego odczytać — `write_only`. Pamięć spada do <10 MB, a czas zapisu także.

**4. Brak formatowania warunkowego dla setek tysięcy komórek.**
Reguła warunkowa z formułą na 200 000 komórek to **koszt, który Excel płaci przy każdym otwarciu i przeliczeniu**. openpyxl tego nie liczy, ale użytkownik czeka. Ogranicz reguły do **kolumny wynikowej** albo do **próbki** — nigdy do całej kolumny (`A:A`) w dużym pliku.

**5. Style z rejestru (moduł 09) — mniej obiektów.**
Nie twórz `Font(bold=True)` w każdym obiegu pętli. Utwórz raz, używaj wszędzie.

```python
# ❌ wolno: 100 000 obiektow Font
for i in range(100_000):
    cell = WriteOnlyCell(ws, value=i)
    cell.font = Font(bold=True)          # nowy obiekt za kazdym razem!

# ✅ szybko: jeden obiekt Font
POGRUBIONY = Font(bold=True)
for i in range(100_000):
    cell = WriteOnlyCell(ws, value=i)
    cell.font = POGRUBIONY               # ten sam obiekt
```

**6. Unikanie `max_row` w warunku pętli.**
`ws.max_row` to **właściwość**, która w trybie normalnym wykonuje `calculate_dimension()`. Jeśli wołasz ją **w każdym obiegu pętli**, robisz to 100 000 razy. Wyciągnij do zmiennej **przed** pętlą.

```python
# ❌ wolno: max_row liczone przy KAZDEJ iteracji
for row in range(1, ws.max_row + 1):
    ...

# ✅ szybko: raz, przed petla
ostatni = ws.max_row
for row in range(1, ostatni + 1):
    ...
```

**7. `gc` i zwalnianie dużych struktur.**
Jeśli musiałeś zebrać dane w dużą listę, **zwolnij ją**, gdy przestaje być potrzebna. `del duza_lista` i — w razie potrzeby — `gc.collect()`. W długo działającym procesie (moduł 26) to zabezpieczenie przed narastaniem pamięci.

```python
import gc

wiersze = list(pobierz_wiersze())     # duza lista
przetworz(wiersze)
del wiersze                           # zwolnij referencje
gc.collect()                          # wymus cykl zbierania (ostroznie!)
```

Uwaga: `gc.collect()` **nie jest** lekarstwem na wszystko i ma swój koszt (zatrzymuje wykonanie). Stosuj **rzadko** — po zwolnieniu naprawdę dużej struktury, nie w każdej iteracji.

**8. Podział na pliki miesięczne/kwartalne zamiast jednego monolitu.**
Technika **architektoniczna**, nie mikrooptymalizacyjna — a często najskuteczniejsza. Zamiast jednego pliku 1M wierszy generujesz 12 plików po ~83 000 wierszy. Zyski:

- pamięć szczytowa spada **12×**, bo każdy plik jest osobny;
- nadają się do **równoległości** (2.10);
- użytkownik otwiera tylko to, co potrzebuje;
- **odporność na awarię** — błąd w jednym miesiącu nie zabija całego raportu.

**9. Równoległość na osobnych plikach (2.10).**
Nie przyspiesza jednego pliku, ale przyspiesza **generowanie wielu** — i to wielokrotnie.

### 2.9. Profilowanie — mierzyć, nie zgadywać

To najważniejsza rada inżynierska tego modułu. **Optymalizacja bez pomiaru to zgadywanie**, a zgadywanie w wydajności zwykle prowadzi do tego, że optymalizujesz 3% kosztu i dziwisz się, że nic nie przyspieszyło.

**Zasada żelazna:** najpierw **zmierz**, potem **zmień jedną rzecz**, potem **zmierz ponownie**.

#### 2.9.1. Czas — `time.perf_counter`

`perf_counter()` to **najdokładniejszy zegar dostępny w Pythonie** do mierzenia czasu trwania. Nie mierzy „czasu systemowego" — mierzy czas rzeczywisty z najwyższą dostępną rozdzielczością.

```python
import time

t0 = time.perf_counter()
zrob_cos()
t1 = time.perf_counter()
print(f"Zajelo: {t1 - t0:.3f} s")
```

Czego **nie** używać: `time.time()` (podatny na korekty zegara systemowego, mniejsza rozdzielczość).

#### 2.9.2. Pamięć — `tracemalloc`

`tracemalloc` to **wbudowany w Pythona** profiler alokacji. Mówi Ci, **ile pamięci zaalokował Twój kod Pythona** — dokładnie to, czego potrzebujesz do porównania trybów.

```python
import tracemalloc

tracemalloc.start()

# ... kod do zmierzenia (np. wczytanie skoroszytu) ...

current, peak = tracemalloc.get_traced_memory()
print(f"Aktualnie: {current / 1_048_576:.1f} MB")
print(f"Szczytowo: {peak / 1_048_576:.1f} MB")

tracemalloc.stop()
```

- **`current`** — ile pamięci zajmują **żywe** obiekty w tej chwili.
- **`peak`** — **najwyższy** poziom w trakcie pomiaru. To jest liczba, która decyduje o tym, czy Twój kontener przeżyje.

Dodatkowa moc: **`tracemalloc.take_snapshot()`** i porównywanie śladów (`snapshot1.compare_to(snapshot2, "lineno")`) pozwala znaleźć **które linie kodu** alokują najwięcej. To jest „gdzie" — nie tylko „ile".

#### 2.9.3. Co się dzieje — `cProfile`

`cProfile` to **profiler wywołań**. Mówi Ci, **w których funkcjach** spędzasz czas.

```python
import cProfile
import pstats
import io

pr = cProfile.Profile()
pr.enable()

# ... kod do profilowania ...

pr.disable()

s = io.StringIO()
ps = pstats.Stats(pr, stream=s).sort_stats("cumulative")
ps.print_stats(15)          # 15 najdrozszych funkcji
print(s.getvalue())
```

Sortowanie po `"cumulative"` pokazuje **łączny** czas funkcji wraz z tym, co wywołują — czyli „gdzie naprawdę idzie czas". Sortowanie po `"tottime"` pokazuje czas **własny** funkcji. Oba są przydatne, ale **zacznij od `cumulative`**.

**Ostrzeżenie praktyczne:** `cProfile` **znacząco zwalnia** kod (narzut na każdą funkcję). Profiluj na **mniejszej próbce** (np. 10 000 wierszy zamiast 1 000 000), żeby zrozumieć **rozkład kosztów**, a potem mierz realny czas na pełnym pliku już bez profilera.

#### 2.9.4. Pamięć „od systemu" — `memory_profiler`

`tracemalloc` widzi alokacje **Pythona**. Ale część pamięci (np. bufory `lxml` czy `zipfile`) może być poza nim. `memory_profiler` (pakiet zewnętrzny, `pip install memory_profiler`) czyta **rzeczywiste zużycie RSS** procesu:

```bash
pip install memory_profiler
```

```python
from memory_profiler import profile

@profile
def generuj_raport():
    ...
    return "ok"

if __name__ == "__main__":
    generuj_raport()
```

Wynik to linia-po-linii wykres zużycia pamięci procesu. To narzędzie jest **wolne** (mierzy przez odpytywanie systemu), więc traktuj je jak „szybkie zdjęcie RTG", nie jak stały monitoring. Jest jednak niezastąpione, gdy `tracemalloc` pokazuje 200 MB, a kontener jest zabijany przy 2 GB — bo wtedy wiesz, że **pamięć idzie gdzieś poza Pythona**.

**Interpretacja w praktyce — jak czytać te trzy liczby:**

| Narzędzie | Mówi Ci | Kiedy użyć |
|---|---|---|
| `time.perf_counter` | „jak długo" | zawsze, przy każdej zmianie |
| `tracemalloc` | „ile zaalokował mój kod i gdzie" | porównanie trybów, szukanie wycieku |
| `cProfile` | „w których funkcjach idzie czas" | gdy jest wolno i nie wiesz dlaczego |
| `memory_profiler` | „ile RSS zjada proces" | gdy `tracemalloc` nie tłumaczy awarii |

### 2.10. Równoległość — osobne pliki, nie jedno `wb`

Tu trzeba być precyzyjnym, bo to temat, w którym łatwo o **nieprawdziwe** porady.

**Fakt pierwszy, z dokumentacji wydajności:**

> *„Reading worksheets is fairly CPU-intensive which limits any benefits to be gained by parallelisation. However, if you are mainly interested in dumping the contents of a workbook then you can use openpyxl's read-only mode and open multiple instances of a workbook and take advantage of multiple CPUs."*

Rozbierzmy to zdanie na czynniki:

1. **Czytanie arkusza jest „CPU-intensive"** — czyli pochłania procesor (`lxml`, parsowanie XML, konwersje typów). To **dobra wiadomość dla równoległości**, bo praca na procesorach zwykle da się rozdzielić.
2. **Ale zysk jest ograniczony**, gdy potrzebujesz **jednego** pliku w całości — bo wtedy musisz zebrać wyniki z powrotem i scalić.
3. **`read_only` czyni równoległość realną**, bo — jak już wiemy — *„opens a workbook almost immediately"*, więc **każdy proces może otworzyć ten sam plik oddzielnie** i czytać swoją część.

**Fakt drugi — i to jest źródło najczęstszego błędu:**

> **Nie dziel jednego obiektu `Workbook` ani `Worksheet` między procesy/wątki.** openpyxl nie jest w tym sensie bezpieczny wątkowo. Model trzyma mutowalny stan (słownik `_cells`, liczniki, bufory strumienia), a `write_only` **wypycha wiersze sekwencyjnie** — równoległe `append` do tego samego arkusza to gwarantowany bałagan.

**Co więc wolno robić równolegle:**

| Wzorzec | Czy OK | Dlaczego |
|---|---|---|
| Wiele procesów czyta **ten sam plik** w `read_only`, każdy inny zakres | ✅ **TAK** | każdy ma własny, niezależny uchwyt ZIP i własny parser |
| Wiele procesów pisze **osobne pliki** (`write_only`) | ✅ **TAK** | każdy `Workbook` jest niezależny |
| Kilka wątków dzieli **jeden `wb`** i pisze do niego | ❌ **NIE** | wspólny, mutowalny stan |
| Kilka procesów pisze do **jednego arkusza** | ❌ **NIE** | `write_only` jest sekwencyjne; sama z siebie kolejność może się pomieszać |
| `ThreadPoolExecutor` na CPU-bound `openpyxl` | ⚠️ **bez sensu** | GIL — wątki nie równoleglą pracy procesora |

Klucz w kodzie: **`concurrent.futures.ProcessPoolExecutor`** (procesy, nie wątki) i **osobne pliki** albo **osobne zakresy tego samego pliku**:

```python
from concurrent.futures import ProcessPoolExecutor

def _przetworz_czesc(args):
    """Ta funkcja uruchamia sie w OSOBNYM PROCESIE."""
    sciezka, min_row, max_row = args
    from openpyxl import load_workbook     # import w procesie potomnym!
    wb = load_workbook(sciezka, read_only=True)
    try:
        ws = wb["Dane"]
        suma = 0
        for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                values_only=True):
            if row[2] is not None:
                suma += row[2]
        return suma
    finally:
        wb.close()
```

**Dwie pułapki, o których nikt nie mówi:**

1. **Funkcja robocza musi być na najwyższym poziomie modułu.** `ProcessPoolExecutor` na „spawn" (Windows, macOS z nowszym Pythonem) **musi ją zaimportować**, więc nie może to być funkcja zagnieżdżona ani lambda. I cały program musi mieć `if __name__ == "__main__":`.
2. **Argumenty i wyniki muszą być „picklowalne".** Nie przekazuj `wb`, `ws` ani `Cell` — przekazuj **ścieżki, liczby i krotki**. To nie jest ograniczenie openpyxl; to ograniczenie przekazywania danych między procesami.

**Zasada kciuka dla doboru `max_workers`:** zacznij od liczby **rdzeni CPU** (`os.cpu_count()`), nie od 50. Przy pracy obciążającej pamięć (duże pliki) zbyt wiele procesów **zwiększy** zużycie i może doprowadzić do swapu — czyli paradoksalnie **spowolni**. Ustaw `max_workers=4` dla 8 rdzeni to rozsądny punkt startowy i zawsze zmierz.

### 2.11. Kiedy openpyxl przestaje być właściwym narzędziem

Wracamy do „nie wozi się węgla ferrari" — ale teraz z **konkretnymi kryteriami**.

**Sygnały, że to już nie zadanie dla openpyxl:**

- Musisz **wielokrotnie** przetwarzać te same dane (agregować, filtrować, łączyć). Powtarzalność = miejsce dla bazy danych.
- Plik ma **miliony wierszy** albo rośnie w czasie. Excel i tak nie otworzy więcej niż 1 048 576 wierszy.
- Dane są **przetwarzane przy każdym uruchomieniu** od nowa. Jeśli tak — nie trzymaj ich w `.xlsx`, trzymaj je raz w bazie.
- Potrzebujesz **zapytań** („znajdź wszystkie zamówienia klienta X w Q3"), a nie tylko wyświetlenia.
- Wielu użytkowników **edytuje** te same dane jednocześnie.

**Co wtedy — tabela alternatyw:**

| Potrzeba | Właściwe narzędzie | Dlaczego |
|---|---|---|
| Dane do przetworzenia, wielokrotnie | **SQLite** / PostgreSQL | zapytania, indeksy, transakcje, jeden plik (SQLite) |
| Duże zbiory do analizy w Pythonie | **Parquet** + pandas/pyarrow | kolumnowy, skompresowany, **bardzo szybki** odczyt, czyta wiersze selektywnie |
| Wymiana danych, prosty format | **CSV** | czytelny, strumieniowy, brak narzutu |
| Excel **jako prezentacja** | **openpyxl** (próbka + podsumowanie) | mały, ładny plik wychodzący do ludzi |

**Wzorzec docelowy — „pipeline", nie „monolit":**

```text
DANE ŹRÓDŁOWE (baza / Parquet / CSV)
        │
        ▼
     AGREGACJA w Pythonie / SQL     <- tu dzieje się cała praca
        │
        ▼
   MAŁY ZESTAW WYNIKOWY (tysiące, nie miliony wierszy)
        │
        ▼
 openpyxl -> ŁADNY PLIK .XLSX       <- tu openpyxl jest na właściwym miejscu
```

To nie jest „poddanie się". To jest **umieszczenie każdego narzędzia tam, gdzie jest dobre**. openpyxl nie konkuruje z bazą danych — on jest **warstwą prezentacji** (moduł 26 rozwinie to jako „raport jako projekcja").

### 2.12. Współpraca z bazą — strumień, nie lista

To realizacja wzorca z 2.11 w praktyce, więc warto go pokazać konkretnie.

**Antywzorzec — „pobierz wszystko, potem zapisz":**

```python
# ❌ PAMIEC: 1 000 000 wierszy w liscie (Python) + caly model (openpyxl) = awaria
wiersze = kursor.fetchall()               # lista 1M krotek
wb = Workbook()                           # drugie tyle w modelu!
ws = wb.active
for w in wiersze:
    ws.append(w)
wb.save("raport.xlsx")
```

Tu masz **dwa pełne zbiory danych w RAM jednocześnie**: listę `wiersze` i model openpyxl. Jeśli plik jest duży, awaria jest kwestią czasu.

**Wzorzec — „strumień od kursora prosto do arkusza":**

```python
from openpyxl import Workbook

def main():
    wb = Workbook(write_only=True)        # pamiec stala!
    ws = wb.create_sheet(title="Sprzedaż")
    ws.append(["ID", "Region", "Kwota"])  # naglowek

    # Kursor NIE trzyma wszystkiego - oddaje wiersz po wierszu.
    # W zaleznosci od sterownika: fetchone() w petli albo iterator kursora.
    while True:
        wiersz = kursor.fetchone()
        if wiersz is None:
            break
        ws.append(list(wiersz))           # od razu do pliku, nie do listy

    wb.save("output/raport_bazy.xlsx")
    wb.close()
```

**Dlaczego to działa i co jest w tym pięknego:**

- **Zużycie pamięci jest O(1)** — jeden wiersz od bazy, jeden wiersz w buforze openpyxl. Reszta jest już na dysku.
- **Dane nie „przechodzą" przez Python w całości.** Nie ma momentu, w którym cały zbiór istnieje w RAM.
- **Nazywa się to strumieniowaniem** — i jest standardowym podejściem przy eksporcie z bazy.

Wariant z generatorem (`yield`), który jest jeszcze elegancki i czyni kod testowalnym (moduł 20):

```python
def wiersze_z_bazy(kursor):
    """Generator: oddaje wiersz po wierszu, nie trzymajac niczego."""
    while True:
        wiersz = kursor.fetchone()
        if wiersz is None:
            return
        yield list(wiersz)

def zapisz(zrodlo, sciezka):
    """zrodlo: dowolny iterowalny zbiór wierszy (baza, plik, generator...)."""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Dane")
    for wiersz in zrodlo:
        ws.append(wiersz)
    wb.save(sciezka)
    wb.close()

# Uzycie:
with polacz_baze() as conn:
    kursor = conn.cursor()
    kursor.execute("SELECT id, region, kwota FROM sprzedaz")
    zapisz(wiersze_z_bazy(kursor), "output/raport.xlsx")
```

**Uwaga o generatorze**, bo to pojęcie z modułu 04 i warto je przypomnieć jednym zdaniem: **generator** to funkcja z `yield`, która oddaje wartości **na żądanie**, jedną po drugiej, zamiast budować całą listę. Analogia: **kurier dostarczający paczki jeden po drugim** vs **hurtownia wynajmująca cały magazyn**. W `append(wiersz)` z generatorem masz „jedną paczkę naraz" — dokładnie to, czego chce strumień.

To rozwiązanie ma jeszcze jedną zaletę: **`zapisz` nie wie, skąd są dane.** Dostaje „zbiór wierszy" — może to być baza, CSV, generator, lista. To jest dokładnie ten **rozdział źródeł i zapisu**, którego pełną formę zobaczysz w module 26 (Ports & Adapters).

### 2.13. Trzy kategorie prawdy — powtórka operacyjna

| Kategoria | W kontekście wydajności |
|---|---|
| **potrafi** | `iter_rows(values_only=True)`, `append`, `write_only` dla dowolnej liczby wierszy, `lxml`, równoległość na osobnych plikach, `calculate_dimension(force=True)`, strumieniowanie z bazy |
| **potrafi częściowo** | `read_only` (tylko wiersze w przód, wymiary bywają `None`, brak formatowania w komórkach), `write_only` (tylko `append`, jeden zapis, `freeze_panes` przed danymi), `conditional_formatting`/`data_validations` w `write_only` (zależy od wersji — sprawdź) |
| **nie ma / gubi** | `iter_cols`/`columns` w `read_only`; tabele w `write_only`; modyfikacja istniejącego arkusza w `write_only`; drugi `save()` → `WorkbookAlreadySaved`; brak `lxml` = brak przyspieszenia |

## 3. Przykłady krok po kroku

### Przykład 1 — 100 000 wierszy: normalny vs `write_only` (🟢)

Punkt odniesienia dla całego modułu. Zmierzymy **czas i pamięć** dwóch trybów zapisu.

```python
"""Modul 19 - zapis 100k wierszy: tryb normalny vs write_only.

Uruchom:  python examples/19_porownanie_zapisu.py
"""

from __future__ import annotations

import time
import tracemalloc
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

LICZBA_WIERSZY = 100_000

# Style tworzone RAZ - technika "styl z rejestru" (sekcja 2.8, punkt 5)
STYL_NAGLOWKA = {
    "font": Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
    "fill": PatternFill("solid", fgColor="1F4E78"),
    "align": Alignment(horizontal="center", vertical="center"),
}


def _naglowek_write_only(ws) -> WriteOnlyCell:
    """Buduje naglowek jako WriteOnlyCell ze stylem."""
    c = WriteOnlyCell(ws, value="ID")
    c.font = STYL_NAGLOWKA["font"]
    c.fill = STYL_NAGLOWKA["fill"]
    c.alignment = STYL_NAGLOWKA["align"]
    return c


def zapis_normalny(sciezka: Path, n: int) -> tuple[float, int]:
    """Zapis w trybie domyslnym. Zwraca (czas_s, szczyt_pamiec_B)."""
    tracemalloc.start()
    t0 = time.perf_counter()

    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"

    # naglowek ze stylem - w trybie normalnym komorka jest trwala
    ws.append(["ID", "Klient", "Kwota"])
    for komorka in ws[1]:
        komorka.font = STYL_NAGLOWKA["font"]
        komorka.fill = STYL_NAGLOWKA["fill"]

    for i in range(1, n + 1):
        ws.append([i, f"Klient {i}", i * 1.5])

    wb.save(sciezka)
    wb.close()

    t1 = time.perf_counter()
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return t1 - t0, peak


def zapis_write_only(sciezka: Path, n: int) -> tuple[float, int]:
    """Zapis w trybie strumieniowym. Zwraca (czas_s, szczyt_pamiec_B)."""
    tracemalloc.start()
    t0 = time.perf_counter()

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Dane")

    # ⚠️ freeze_panes PRZED danymi - inaczej nie zadziala!
    ws.freeze_panes = "A2"

    # naglowek: WriteOnlyCell bo potrzebujemy stylu
    ws.append([
        _naglowek_write_only(ws),
        WriteOnlyCell(ws, value="Klient"),
        WriteOnlyCell(ws, value="Kwota"),
    ])

    for i in range(1, n + 1):
        # zwykle wartosci - bez stylu, wiec bez alokacji WriteOnlyCell
        ws.append([i, f"Klient {i}", i * 1.5])

    wb.save(sciezka)
    wb.close()

    t1 = time.perf_counter()
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return t1 - t0, peak


def main() -> None:
    print("=" * 74)
    print(f"ZAPIS {LICZBA_WIERSZY:,} WIERSZY - PORÓWNANIE TRYBÓW".replace(",", " "))
    print("=" * 74)

    cel_a = OUTPUT / "19_normalny.xlsx"
    cel_b = OUTPUT / "19_write_only.xlsx"

    print("\n[1/2] Tryb domyslny (Workbook)...")
    czas_a, pamiec_a = zapis_normalny(cel_a, LICZBA_WIERSZY)
    print(f"      czas:  {czas_a:6.2f} s")
    print(f"      szczyt pamięci (Python): {pamiec_a / 1_048_576:7.1f} MB")
    print(f"      plik:  {cel_a.stat().st_size / 1_048_576:.2f} MB")

    print("\n[2/2] Tryb strumieniowy (Workbook(write_only=True))...")
    czas_b, pamiec_b = zapis_write_only(cel_b, LICZBA_WIERSZY)
    print(f"      czas:  {czas_b:6.2f} s")
    print(f"      szczyt pamięci (Python): {pamiec_b / 1_048_576:7.1f} MB")
    print(f"      plik:  {cel_b.stat().st_size / 1_048_576:.2f} MB")

    print("\n" + "=" * 74)
    print("PODSUMOWANIE")
    print("=" * 74)
    print(f"  Czas:   {czas_a:6.2f} s  ->  {czas_b:6.2f} s   "
          f"({czas_b / czas_a:.2f}x)")
    print(f"  Pamięć: {pamiec_a / 1_048_576:7.1f} MB  ->  "
          f"{pamiec_b / 1_048_576:7.1f} MB   "
          f"({pamiec_b / pamiec_a:.3f}x)")
    print()
    print("  Uwaga: tracemalloc mierzy alokacje PYTHONA, nie caly RSS procesu.")
    print("  Realna różnica w RSS bywa jeszcze wieksza (bufory XML/ZIP).")
    print()
    print("  WERYFIKACJA OBOWIĄZKOWA: otwórz OBA pliki w Excelu.")
    print("    - w obu musi być nagłówek z niebieskim tłem i białym tekstem,")
    print("    - w obu musi być zamrożony wiersz 1 (freeze panes),")
    print("    - liczba wierszy danych musi być identyczna.")
    print("  Jesli write_only nie ma stylu nagłówka -> zapomniałeś o")
    print("  WriteOnlyCell i przekazales zwykly string.")
    print("  Jesli brak freeze panes -> ustawiles go PO pierwszych danych.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W trybie normalnym po pętli `ws._cells` zawiera **300 000 obiektów `Cell`** (100 000 wierszy × 3 kolumny) — każdy ze swoim zestawem atrybutów. W `write_only` po pętli w pamięci jest **co najwyżej bufor jednego wiersza** — wszystko inne zostało już wypchnięte do archiwum. Dlatego szczyt pamięci dla trybu domyślnego będzie rzędu **setek megabajtów**, a dla strumieniowego — **kilku megabajtów**: różnica jednego-dwóch rzędów wielkości.

**Co trafi do pliku.** Identyczna liczba komórek w obu plikach. Ale uwaga na **dwie różnice techniczne**, które wynikają z natury trybów:

1. **Styl nagłówka.** W trybie normalnym zmodyfikowałem komórki po `append` — to działa, bo komórki są trwałe. W `write_only` **musiałem** zbudować `WriteOnlyCell` **przed** `append`, bo po nim komórka jest już wydrukowana. To dokładnie ta różnica z 2.4.2.
2. **`freeze_panes`.** W trybie normalnym mógłbym ustawić go **na końcu**, przed `save`. W `write_only` **musiałem** to zrobić **przed pierwszym `append`** — inaczej byłoby ignorowane.

**Trzy rzeczy do przemyślenia:**

1. **Uruchom to na swoim komputerze.** Czasy i pamięć zależą od sprzętu, Pythona i obecności `lxml`. **Twoje liczby są ważniejsze niż moje** — bo to one opisują Twój system produkcyjny.
2. **Sprawdź, czy `lxml` jest aktywny** (`from openpyxl import LXML; print(LXML)`) i uruchom skrypt **z nim i bez niego** (`OPENPYXL_LXML=False python ...`). Zobaczysz różnicę na własne oczy — i to jest lekcja, której nie da się zapomnieć.
3. **Dlaczego w `write_only` używam `WriteOnlyCell` TYLKO w nagłówku?** Bo styl kosztuje. Każda `WriteOnlyCell` to obiekt do zbudowania i wypchnięcia. Dla 100 000 wierszy danych **bez formatowania** przekazuję zwykłe wartości — szybciej i prościej. **Formatuj to, co ma znaczenie dla człowieka** (nagłówki, podsumowania), nie każdą z 100 000 komórek danych.

### Przykład 2 — `read_only`: 200 000 wierszy i sumy warunkowe (🟡)

Teraz strona odczytu. Pokażemy strumieniowe czytanie, sprawdzanie wymiarów i bezpieczne obsłużenie `None`/`EmptyCell`.

```python
"""Modul 19 - read_only: czytanie 200k wierszy i agregacja strumieniowa.

Uruchom:  python examples/19_read_only.py
"""

from __future__ import annotations

import time
import tracemalloc
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
DATA.mkdir(parents=True, exist_ok=True)
OUTPUT.mkdir(parents=True, exist_ok=True)

ZRODLO = DATA / "19_duze_dane.xlsx"
LICZBA_WIERSZY = 200_000


def przygotuj_plik() -> None:
    """Tworzy duzy plik testowy w write_only (bo 200k wierszy w normalnym boli)."""
    if ZRODLO.exists():
        return
    print(f"Tworzę plik testowy ({LICZBA_WIERSZY:,} wierszy) w write_only..."
          .replace(",", " "))
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Dane")
    ws.append(["ID", "Region", "Kwota", "Status"])
    regiony = ["Północ", "Południe", "Wschód", "Zachód"]
    for i in range(1, LICZBA_WIERSZY + 1):
        # prowizja: None w co 37. wierszu -> test obslugi brakow!
        kwota = None if i % 37 == 0 else float(i % 1000) * 1.5
        status = "OK" if i % 11 else "REKLAMACJA"
        ws.append([i, regiony[i % 4], kwota, status])
    wb.save(ZRODLO)
    wb.close()
    print(f"  gotowe: {ZRODLO.name}  ({ZRODLO.stat().st_size / 1_048_576:.1f} MB)")


def analiza_read_only(sciezka: Path) -> dict:
    """Strumieniowa analiza: suma kwot per region + liczba brakow."""
    tracemalloc.start()
    t0 = time.perf_counter()

    wb = load_workbook(sciezka, read_only=True)
    try:
        ws = wb["Dane"]

        # --- 1. Sprawdz wymiary (moga byc None!) ---
        if ws.max_row is None or ws.max_column is None:
            print("  ⚠️  Wymiary nieustalone - forsujemy skan (jednorazowo)...")
            wymiar = ws.calculate_dimension(force=True)
            print(f"      calculate_dimension(force=True) = {wymiar}")
        print(f"  max_row={ws.max_row}  max_column={ws.max_column}")

        # --- 2. Iteracja wierszami, TYLKO WARTOŚCI (najszybsza sciezka) ---
        sumy: dict[str, float] = {}
        licznik: dict[str, int] = {}
        braki = 0
        wierszy = 0

        iterator = ws.iter_rows(min_row=2, values_only=True)
        for wiersz in iterator:
            identyfikator, region, kwota, status = wiersz

            if kwota is None:
                braki += 1
                continue

            sumy[region] = sumy.get(region, 0.0) + kwota
            licznik[region] = licznik.get(region, 0) + 1
            wierszy += 1
    finally:
        wb.close()          # ⚠️ OBOWIAZKOWE - inaczej wyciek uchwytu

    t1 = time.perf_counter()
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    return {
        "sumy": sumy,
        "licznik": licznik,
        "braki": braki,
        "wierszy_z_danymi": wierszy,
        "czas_s": t1 - t0,
        "szczyt_MB": peak / 1_048_576,
    }


def main() -> None:
    przygotuj_plik()

    print("=" * 74)
    print("READ_ONLY - AGREGACJA STRUMIENIOWA")
    print("=" * 74)

    wynik = analiza_read_only(ZRODLO)

    print(f"\n  Przetworzono {wynik['wierszy_z_danymi']:,} wierszy z danymi"
          .replace(",", " "))
    print(f"  Pominięto braków (None): {wynik['braki']}")
    print(f"  Czas:   {wynik['czas_s']:.2f} s")
    print(f"  Pamięć: szczyt {wynik['szczyt_MB']:.1f} MB "
          f"(stała, niezależna od liczby wierszy!)")

    print("\n  Sumy per region:")
    for region in sorted(wynik["sumy"]):
        suma = wynik["sumy"][region]
        n = wynik["licznik"][region]
        print(f"    {region:<10} {n:>7} wierszy   suma = {suma:>14,.2f}")

    print()
    print("=" * 74)
    print("KLUCZOWA OBSERWACJA")
    print("=" * 74)
    print("  Szczyt pamięci NIE rośnie z liczbą wierszy.")
    print("  Sprawdź to: zmień LICZBA_WIERSZY na 400_000 i uruchom ponownie.")
    print("  Szczyt pozostanie niemal identyczny - o to chodzi w strumieniu!")
    print()
    print("  Gdybyś użył pominiętego trybu (load_workbook bez read_only),")
    print(f"  szczyt byłby rzędu {LICZBA_WIERSZY * 4 * 250 / 1_048_576:.0f} MB+,")
    print("  bo powstałby obiekt Cell dla każdej z 800 000 komórek.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Kluczowa jest pętla: bierzemy **jeden wiersz naraz** i natychmiast go porzucamy — w pamięci żyje tylko krotka z 4 wartościami i dwa słowniki agregatów (`sumy`, `licznik`), które mają po 4 wpisy. Szczyt pamięci będzie **kilka–kilkanaście megabajtów** niezależnie od tego, czy w pliku jest 200 000 czy 2 000 000 wierszy. To jest **cała istota strumienia**.

**Co trafi do pliku.** Nic — to przykład tylko do odczytu. Ale gdybyś chciał zapisać wynik (4 wiersze sum per region), zrobiłbyś to zwykłym `Workbook()` — i to jest wzorzec z 2.11: **czytaj strumieniowo, zapisuj mały wynik**.

**Trzy rzeczy do przemyślenia:**

1. **Sprawdzenie wymiarów na starcie nie jest przesadą.** Metadane `<dimension>` w plikach tworzonych przez różne narzędzia bywają błędne. `calculate_dimension(force=True)` robi jednorazowe przejście i ustawia `max_row`/`max_column` — i to jest **dokładnie** to, co dokumentacja opisuje jako lekarstwo: *„resetting the `max_row` and `max_column` attributes should allow you to work with the file"*.
2. **`kwota is None`, nie `kwota == 0`.** Pamiętasz z modułu 05: **pusta komórka to nie zero**. W tym pliku co 37. wiersz ma `None` — i gdybyś zrobił `sumy[region] += kwota`, dostałbyś `TypeError: unsupported operand type(s) for +=: 'float' and 'NoneType'`. To jest właśnie ta pułapka, o której mówiliśmy — teraz widzisz ją w akcji.
3. **`finally: wb.close()` jest tu krytyczne.** W `read_only` openpyxl trzyma **otwarty uchwyt do archiwum ZIP** (2.3). Bez `close()` nie zwalnia go — a w pętli generującej tysiące raportów to wyciek, który skończy się `Too many open files`. W historii openpyxl były konkretne błędy w tym obszarze (#2122 „File handlers not always released in read-only mode", #2149 „Workbook files not properly closed on Python ≥ 3.11.8 and Windows") — naprawione w 3.1.3, ale zasada „zawsze `close()` w `finally`" chroni Cię niezależnie od wersji.

### Przykład 3 — generator + `WriteOnlyCell`: eksport strumieniowy ze stylowaniem (🟡)

To wzorzec łączący 2.8, 2.12 i 2.4: **dane płyną generatorem, arkusz je wypycha, style z rejestru**.

```python
"""Modul 19 - generator + write_only: strumieniowy eksport ze stylowaniem.

Uruchom:  python examples/19_generator_stream.py
"""

from __future__ import annotations

import time
from collections.abc import Iterator
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.datetime import to_excel     # do konwersji na liczbe Excela

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

LICZBA = 150_000

# --- REJESTR STYLOW (modul 09) - kazdy obiekt tworzony RAZ ---
REGISTER = {
    "naglowek": {
        "font": Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
        "fill": PatternFill("solid", fgColor="2F5496"),
        "align": Alignment(horizontal="center", vertical="center"),
    },
    "kwota": {"fmt": "#,##0.00"},
    "data": {"fmt": "yyyy-mm-dd"},
}


def _komorka_naglowka(ws, tekst: str) -> WriteOnlyCell:
    c = WriteOnlyCell(ws, value=tekst)
    for k, v in REGISTER["naglowek"].items():
        if k == "font":
            c.font = v
        elif k == "fill":
            c.fill = v
        elif k == "align":
            c.alignment = v
    return c


def _komorka_kwoty(ws, wartosc: float) -> WriteOnlyCell:
    c = WriteOnlyCell(ws, value=wartosc)
    c.number_format = REGISTER["kwota"]["fmt"]
    return c


def _komorka_daty(ws, wartosc: date) -> WriteOnlyCell:
    c = WriteOnlyCell(ws, value=wartosc)
    c.number_format = REGISTER["data"]["fmt"]
    return c


def strumien_danych(n: int) -> Iterator[tuple]:
    """Generator: oddaje wiersz po wierszu, NIE budujac listy.

    W realnym systemie to bylby kursor bazy albo wiersz z Parquet/CSV.
    Tutaj symulujemy - ale interfejs jest identyczny.
    """
    start = date(2026, 1, 1)
    for i in range(1, n + 1):
        yield (
            i,
            f"Faktura/{i:08d}",
            start + timedelta(days=i % 365),
            round((i % 5000) * 1.37, 2),
        )


def eksportuj(sciezka: Path, n: int) -> None:
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Faktury")

    # ⚠️ Wszystko PRZED danymi (2.4, zasada 4):
    ws.freeze_panes = "A2"
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    # Naglowek - ze stylem, wiec przez WriteOnlyCell
    ws.append([
        _komorka_naglowka(ws, "ID"),
        _komorka_naglowka(ws, "Numer"),
        _komorka_naglowka(ws, "Data"),
        _komorka_naglowka(ws, "Kwota"),
    ])

    # Dane ze strumienia - zgodnie ze stylem tylko tam, gdzie ma sens
    for identyfikator, numer, data_wyst, kwota in strumien_danych(n):
        ws.append([
            identyfikator,
            numer,
            _komorka_daty(ws, data_wyst),     # format daty
            _komorka_kwoty(ws, kwota),        # format kwoty
        ])

    wb.save(sciezka)
    wb.close()


def main() -> None:
    cel = OUTPUT / "19_faktury_stream.xlsx"

    print("=" * 74)
    print(f"EKSPORT STRUMIENIOWY: {LICZBA:,} wierszy".replace(",", " "))
    print("=" * 74)

    t0 = time.perf_counter()
    eksportuj(cel, LICZBA)
    t1 = time.perf_counter()

    print(f"  czas:  {t1 - t0:.2f} s")
    print(f"  plik:  {cel.stat().st_size / 1_048_576:.2f} MB")
    print()
    print("  Uwaga: NIGDZIE nie powstała lista ze wszystkimi wierszami.")
    print("  Strumien -> append -> plik, wiersz po wierszu.")
    print()
    print("  WERYFIKACJA 'NA OKO' (obowiazkowa!):")
    print(f"    otwórz {cel.name} w Excelu i sprawdź:")
    print("      - kolumna C formatuje się jako data (2026-XX-XX),")
    print("      - kolumna D ma dwa miejsca po przecinku i separator tysiecy,")
    print("      - wiersz nagłówka jest zamrożony (przewiń w dół!),")
    print("      - nagłówek ma granatowe tło i biały tekst.")
    print("    Potem przejdź na koniec (Ctrl+End) - ostatni ID =", LICZBA)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `strumien_danych` to **generator** — nie ma w nim listy 150 000 krotek. Każde `next()` tworzy jedną krotkę, którą `eksportuj` natychmiast pakuje do arkusza. Po `append` wiersz jest serializowany do XML-a i porzucony. Szczyt pamięci: kilka megabajtów.

**Co trafi do pliku.** 150 000 wierszy z dwiema kolumnami sformatowanymi (`yyyy-mm-dd` i `#,##0.00`) plus nagłówek ze stylem i zamrożony wiersz. Zwróć uwagę, że **każda komórka z datą i kwotą to osobna `WriteOnlyCell`** — bo tylko tak da się nadać `number_format` w strumieniu. Ale **komórki `ID` i `Numer` to zwykłe wartości** — bo nie potrzebują formatu. To jest ta oszczędność z przykładu 1.

**Trzy rzeczy do przemyślenia:**

1. **`to_excel` nie jest tu użyte, choć zaimportowane.** Zwróć uwagę: `WriteOnlyCell` z `value=date(2026, 1, 1)` **sam** konwertuje datę na liczbę Excela (bo openpyxl rozpoznaje typy `datetime`). Import `to_excel` zostawiłem w kodzie jako **wskazówkę** — jest potrzebny, gdy konwertujesz daty ręcznie (moduł 05). W produkcji usuń nieużywany import (moduł 20: `ruff` to złapie).
2. **`column_dimensions` DZIAŁA, ale tylko dlatego, że jest przed danymi.** To jest dokładnie ten punkt, który autor openpyxl nazwał „trudniejszym do obsłużenia" — szerokości kolumn leżą w XML-ie **przed** `sheetData`. Zmieniłbyś kolejność i szerokości by nie zadziałały.
3. **Ten wzorzec jest gotowy na bazę danych.** Podmień `strumien_danych` na `wiersze_z_bazy(kursor)` z 2.12 **i nic więcej nie zmieniaj** — bo `eksportuj` nie wie, skąd płyną dane. To jest zapowiedź modułu 26: **funkcja zapisu nie może wiedzieć, skąd dane pochodzą.**

### Przykład 4 — `ProcessPoolExecutor`: 1M wierszy rozdzielone na 12 plików (🔴)

Kulminacja modułu: podział monolitu (2.8, punkt 8) + równoległość (2.10).

```python
"""Modul 19 - rownolegle generowanie 12 plikow (1M wierszy razem).

Uruchom:  python examples/19_rownolegle.py
UWAGA: wymaga uruchomienia jako plik (nie w notatniku) - spawn na Win/macOS.
"""

from __future__ import annotations

import os
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output" / "miesieczne"

# Po ~83 334 na miesiac daje ok. 1 000 000 wierszy razem
WIERSZY_NA_MIESIAC = 83_334


def _zapisz_miesiac(zadanie: tuple[str, int, int]) -> tuple[str, int, float]:
    """Ta funkcja dziala w OSOBNYM PROCESIE.

    Dostaje TYLKO proste typy (str, int) - nie obiekty openpyxl!
    Import jest lokalny, bo proces potomny nie dziedziczy importow.
    """
    katalog, rok, miesiac = zadanie

    from openpyxl import Workbook     # import w procesie potomnym (spawn!)
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font

    t0 = time.perf_counter()

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title=f"{rok}-{miesiac:02d}")

    # wszystko przed danymi
    ws.freeze_panes = "A2"

    # jeden styl, tworzony raz (technika z 2.8, punkt 5)
    naglowek_font = Font(bold=True)
    def _naglowek(tekst: str) -> WriteOnlyCell:
        c = WriteOnlyCell(ws, value=tekst)
        c.font = naglowek_font
        return c

    ws.append([
        _naglowek("ID"),
        _naglowek("Miesiąc"),
        _naglowek("Kwota"),
    ])

    for i in range(1, WIERSZY_NA_MIESIAC + 1):
        ws.append([i, miesiac, round(i * 0.5, 2)])

    sciezka = Path(katalog) / f"raport_{rok}-{miesiac:02d}.xlsx"
    wb.save(sciezka)
    wb.close()

    t1 = time.perf_counter()
    return str(sciezka), WIERSZY_NA_MIESIAC, t1 - t0


def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    rok = 2026

    rdzenie = os.cpu_count() or 1
    # Nie przesadzaj: przy pracy obciazajacej pamiec zbyt wiele procesow
    # zwieksza zuzycie i moze doprowadzic do swapu (sekcja 2.10).
    max_workers = min(4, rdzenie)

    zadania = [(str(OUTPUT), rok, m) for m in range(1, 13)]

    print("=" * 74)
    print(f"RÓWNOLEGŁE GENEROWANIE: 12 plików × {WIERSZY_NA_MIESIAC:,} wierszy"
          .replace(",", " "))
    print(f"Rdzanie CPU: {rdzenie}   max_workers: {max_workers}")
    print("=" * 74)

    t0 = time.perf_counter()

    wyniki: list[tuple[str, int, float]] = []
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        for przyszly in as_completed(
            [executor.submit(_zapisz_miesiac, z) for z in zadania]
        ):
            sciezka, n, czas = przyszly.result()
            wyniki.append((sciezka, n, czas))
            print(f"  ✓ {Path(sciezka).name}: {n:,} wierszy w {czas:.2f} s"
                  .replace(",", " "))

    t1 = time.perf_counter()

    wyniki.sort()
    suma_wierszy = sum(n for _, n, _ in wyniki)
    suma_sciezek = sum(Path(s).stat().st_size for s, _, _ in wyniki)

    print()
    print("=" * 74)
    print("PODSUMOWANIE")
    print("=" * 74)
    print(f"  Plików:        {len(wyniki)}")
    print(f"  Wierszy razem: {suma_wierszy:,}".replace(",", " "))
    print(f"  Rozmiar razem: {suma_sciezek / 1_048_576:.1f} MB")
    print(f"  Czas ZUPEŁNY:  {t1 - t0:.2f} s  (zmierzony na zegarze sciennym)")
    print(f"  Suma czasów:   {sum(c for _, _, c in wyniki):.2f} s")
    print()
    print("  Jesli 'czas zupełny' << 'suma czasów', procesy naprawde")
    print("  pracowaly rownolegle. Jesli sa podobne - pracowaly szeregowo")
    print("  (za malo rdzeni albo praca byla I/O-bound).")
    print()
    print("  WERYFIKACJA: otwórz 2-3 pliki w Excelu. Kazdy musi miec:")
    print("    nagłówek pogrubiony, zamrożony wiersz 1, nazwę arkusza = rok-miesiąc.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Tu nie ma już jednego „modelu" — jest **12 niezależnych procesów**, każdy z własnym `Workbook(write_only=True)` i własnym buforem jednego wiersza. Pamięć całego systemu to 12 × (kilka MB) = kilkadziesiąt megabajtów, zamiast jednej monstrualnej struktury.

**Co trafi do pliku.** 12 osobnych plików `.xlsx` — każdy z własnym nagłówkiem, własnym arkuszem i swoim zakresem danych. Sumarycznie ~1 000 000 wierszy. **Każdy plik pojedynczo otwiera się w Excelu normalnie.**

**Trzy rzeczy do przemyślenia:**

1. **Import `openpyxl` jest WEWNĄTRZ funkcji `_zapisz_miesiac`.** To nie kaprys — to wymóg modelu „spawn": proces potomny **startuje od zera** i nie dziedziczy importów. Bez tego dostaniesz `NameError`. Zasada: **w funkcji roboczej importuj to, czego używasz, lokalnie.**
2. **`if __name__ == "__main__":` jest obowiązkowe.** Bez tego na Windows i macOS procesy potomne **ponownie zaimportują Twój moduł** i dojdzie do nieskończonej rekurencji tworzenia procesów. To najczęstsza przyczyna „program zamraża komputer" z `ProcessPoolExecutor`.
3. **`max_workers = min(4, rdzenie)` — świadomie nisko.** Kusi, żeby ustawić 12 i „mieć 12× szybciej". Nie zadziała: przy pracy obciążającej pamięć więcej procesów = większy sumaryczny RAM i ryzyko swapu. Zasada: **ustaw konserwatywnie, zmierz, potem zwiększaj.** Miara jest w wydruku: porównaj „czas zupełny" z „sumą czasów". Jeśli są zbliżone, równoległość nie działa — i wiesz to **z liczb**, nie z przeczucia.

## 4. Anatomia API

| Klasa / funkcja / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `load_workbook(path, read_only=True)` | Wczytuje leniwie, strumieniowo | — | Brak dostępu losowego, DV, CF, obrazów, wykresów, scaleń |
| `Workbook(write_only=True)` | Skoroszyt strumieniowy | — | **Brak arkusza domyślnego** — trzeba `create_sheet()` |
| `ReadOnlyWorksheet` | Odpowiednik `Worksheet` w `read_only` | — | `iter_rows`/`rows`/`values` tak; `iter_cols`/`columns` **nie** |
| `WriteOnlyWorksheet` | Odpowiednik `Worksheet` w `write_only` | — | Tylko `append()`; nie ma `cell()`/`iter_rows()` |
| `ReadOnlyCell` | Komórka w `read_only` | — | Brak `font`/`fill`/`border`/`comment`/`hyperlink`; przypisanie → `AttributeError` |
| `WriteOnlyCell(ws, value=...)` | Komórka do nadania stylu w strumieniu | `value` | **Jednorazowa** — styl ustaw przed `append`; import: `from openpyxl.cell import WriteOnlyCell` |
| `ws.append(...)` | Dodaje wiersz | iterowalne wartości/`WriteOnlyCell` | **Jedyny sposób zapisu w `write_only`** |
| `ws.iter_rows(min_row, max_row, min_col, max_col, values_only=False)` | Wiersze jako krotki komórek/wartości | `values_only=True` = **najszybsza ścieżka** | W `read_only` tylko w przód |
| `ws.values` | Skrót: wiersze wartościami | — | Odpowiednik `iter_rows(values_only=True)` |
| `ws.rows` | Generator wierszy | — | Dostępne w `read_only`; `ws.columns` — **nie** |
| `ws.max_row`, `ws.max_column` | Wymiary | — | W `read_only` **mogą być `None`** |
| `ws.calculate_dimension(force=False)` | Zwraca zakres np. `A1:M24` | `force=True` wymusza skan | Bez `force` przy nieznanych wymiarach → `ValueError` |
| `ws.reset_dimensions()` | Zeruje wymiary | — | **Usuwa**, nie liczy; przydatne przy błędnych metadanych |
| `ws.freeze_panes = "A2"` | Zamraża wiersz/kolumnę | — | W `write_only` **tylko PRZED danymi** |
| `ws.column_dimensions["B"].width` | Szerokość kolumny | — | W `write_only` **tylko PRZED danymi** |
| `wb.save(path)` | Zapisuje | — | W `read_only` → `TypeError`; w `write_only` **tylko raz** |
| `wb.close()` | Zamyka uchwyt | — | **Obowiązkowe** w `read_only`/`write_only` |
| `WorkbookAlreadySaved` | Wyjątek | — | `from openpyxl.utils.exceptions import WorkbookAlreadySaved` |
| `openpyxl.LXML` | Czy używamy `lxml`? | — | `True`/`False`; zależy od instalacji **i** `OPENPYXL_LXML` |
| `openpyxl.DEFUSEDXML` | Czy chroni `defusedxml`? | — | Aktywne, gdy brak `lxml` |
| `OPENPYXL_LXML` (env var) | Włącza/wyłącza `lxml` | `"True"` = włącz | Ustaw na inną wartość, by **wyłączyć** bez odinstalowywania |
| `write_conditional_formatting(ws)` | Wypisuje CF do XML | — | W `write_only` wołane **po** danych (bo tam leży w XML) |
| `write_validations(ws)` / `ws.data_validations` | Walidacja | — | W `write_only` **zweryfikuj `hasattr`** — patrz 2.6 |
| `DataValidationList` | Lista walidacji | — | `from openpyxl.worksheet.datavalidation import DataValidationList` |
| `ConditionalFormattingList` | Lista reguł CF | — | `from openpyxl.formatting.formatting import ConditionalFormattingList` |
| `time.perf_counter()` | Zegar wysokiej rozdzielczości | — | Do mierzenia czasu (nie `time.time()`) |
| `tracemalloc.start/get_traced_memory/stop` | Profil pamięci Pythona | `(current, peak)` w bajtach | `peak` = liczba decydująca o przeżyciu kontenera |
| `tracemalloc.take_snapshot()` | Ślad alokacji | `.compare_to(...)` | Znajduje **linie**, które alokują |
| `cProfile.Profile()` | Profil wywołań | `.enable()`, `.disable()` | Sortuj po `cumulative` |
| `ProcessPoolExecutor(max_workers=N)` | Procesy robocze | funkcja **na najwyższym poziomie** | Argumenty picklowalne; importy lokalnie |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — zbuduj własną tabelę „operacja → czas".**

Napisz `examples/19_cw1.py`, który generuje pliki o trzech rozmiarach (**10 000**, **100 000**, **1 000 000** wierszy) w dwóch trybach i zbiera **czasy** w jedną tabelę.

Wymagania:

1. Trzy osobne funkcje pomocnicze: `zapisz_normalny(n)`, `zapisz_write_only(n)`, `wczytaj_read_only(n)`.
2. Każda funkcja mierzy się **własnym** `time.perf_counter()` i zwraca czas w sekundach.
3. Porządek pomiarów: najpierw wszystkie **10k**, potem wszystkie **100k**, potem **1M** — żeby porównywać w podobnych warunkach.
4. Wynik wypisz jako prostą tabelę:

   ```text
   Rozmiar      normalny   write_only   read_only
   10 000       ...  s      ...  s        ...  s
   100 000      ...  s      ...  s        ...  s
   1 000 000    ...  s      ...  s        ...  s
   ```
5. Na końcu wypisz **liczbę aktywnych backendów**:

   ```python
   from openpyxl import LXML
   print("lxml aktywny:", LXML)
   ```

6. Uruchom **dwa razy**: raz normalnie, raz z `OPENPYXL_LXML=False` (w Linux/macOS: `OPENPYXL_LXML=False python examples/19_cw1.py`). Wypisz oba wyniki.

W `output/19_cw1_wnioski.md` odpowiedz:

1. **Jak zmienia się czas w trybie normalnym przy 10× więcej danych?** Czy liniowo? Podaj konkretne liczby i stosunek.
2. **W którym momencie `read_only` zaczyna wyraźnie wygrywać?** Dlaczego przy 10k różnica jest mała, a przy 1M — ogromna?
3. **Ile wyniosła różnica `lxml` vs bez `lxml`?** Dla której operacji była największa? Dlaczego (podpowiedź: `lxml` przyspiesza **oba** kierunki, ale nie w tym samym stopniu)?
4. **Która operacja jest najdroższa i dlaczego?** Nie zgaduj — oprzyj się na swoich liczbach.
5. **Co by się stało z plikiem 1 000 000 wierszy w trybie normalnym na maszynie z 2 GB RAM?** Oszacuj zużycie, korzystając z „50× rozmiar pliku" z dokumentacji i ze zmierzonego rozmiaru pliku.

### 🟡 Warsztat

**Zadanie 2 — analiza 200 000 wierszy w `read_only`.**

Napisz `examples/19_cw2.py`, który:

1. Tworzy plik `data/19_cw2_duze.xlsx` (**200 000** wierszy) w trybie `write_only`, z kolumnami: `ID`, `Data` (`datetime`), `Kategoria` (4 wartości cyklicznie), `Kwota` (`float`, co 41. wiersz = `None`), `Kraj` (5 wartości cyklicznie).
2. Wczytuje go **tylko** przez `load_workbook(..., read_only=True)` i liczy:

   - **sumę i średnią** kwoty **per kategoria** (pamiętając, że `None` to nie `0`!),
   - **liczbę braków** kwot,
   - **liczbę wierszy per kraj**,
   - **miesiąc z największą sumą** (konwertując `datetime` → miesiąc),
   - **minimalną i maksymalną** kwotę **per kategoria**.
3. Na start wypisuje `ws.max_row`/`ws.max_column` i — jeśli są `None` — używa `calculate_dimension(force=True)`.
4. Mierzy **czas** i **szczyt pamięci** przez `tracemalloc` i oba wypisuje.
5. Wypisuje wyniki jako czytelny raport tekstowy.
6. **Dodatkowo** — w osobnym przebiegu — próbuje (w `try/except`) i **wypisuje dokładny komunikat błędu** dla:

   - `ws.iter_cols()`,
   - `wb.save("cokolwiek.xlsx")`,
   - `ws["A1"] = 5`,
   - odczytu `ws["A1"].font`.

W `output/19_cw2_wnioski.md` odpowiedz:

1. **Wklej cztery komunikaty błędów** z punktu 6 i do każdego dopisz **jedno zdanie wyjaśnienia**, dlaczego taki błąd jest właściwym zachowaniem.
2. **Ile wyniósł szczyt pamięci przy 200 000 wierszy?** Porównaj z rozmiarem pliku na dysku (skoroszyt w trybie normalnym zająłby ok. 50× tyle). Jaki jest **stosunek oszczędności**?
3. **Co by się stało, gdybyś policzył średnią kwoty bez sprawdzania `None`?** Pokaż, jakiego rodzaju błąd by wystąpił i gdzie dokładnie w kodzie.
4. **Jak obsłużyłeś konwersję `datetime` na miesiąc?** Czy `cell.value` zwraca obiekt `datetime` czy `int` (serial Excela)? Sprawdź eksperymentalnie i wyjaśnij, odwołując się do modułu 05.
5. **Czy `calculate_dimension(force=True)` zmienił wynik analizy?** Sprawdź, czy `max_row` przed i po był taki sam. Jeśli różny — co to oznacza o metadanych pliku?
6. **Czy dałoby się policzyć „min/max per kategoria" bez trzymania wszystkich wartości?** Opisz (nie implementuj) dwa podejścia i porównaj ich zużycie pamięci.

### 🔴 Wyzwanie

**Zadanie 3 — generator raportów miesięcznych w `ProcessPoolExecutor`.**

Napisz `examples/19_cw3.py` — produkcyjny generator 12 plików z 1 000 000 wierszy **łącznie**, w pełni równoległy.

**Wymagania funkcjonalne:**

1. Funkcja `_zapisz_miesiac(zadanie)` — **na najwyższym poziomie modułu**, przyjmuje krotkę `(katalog, rok, miesiac, liczba_wierszy, specyfikacja_stylow)` i:

   - tworzy `Workbook(write_only=True)` i `create_sheet(title=f"{rok}-{miesiac:02d}")`,
   - **przed danymi** ustawia `freeze_panes`, szerokości kolumn, `sheet_view.showGridLines = False`,
   - tworzy nagłówek przez `WriteOnlyCell` z **jednym** obiektem `Font` utworzonym raz,
   - dopisuje wiersze z **generatora** wewnętrznego (nie z listy!),
   - w kolumnie `Kwota` używa `WriteOnlyCell` z `number_format="#,##0.00"`,
   - zapisuje plik i zwraca krotkę `(sciezka, liczba_wierszy, czas_s, rozmiar_B)`.

2. Funkcja `main()`, która:

   - tworzy katalog wyjściowy,
   - buduje `zadania` dla 12 miesięcy,
   - uruchamia `ProcessPoolExecutor` z `max_workers = min(4, os.cpu_count() or 1)`,
   - zbiera wyniki przez `as_completed` i wypisuje postęp na bieżąco,
   - na końcu wypisuje: liczbę plików, sumę wierszy, sumę rozmiaru, **czas zupełny**, **sumę czasów** i **współczynnik równoległości** (`suma czasów / czas zupełny`).

3. Ochrona `if __name__ == "__main__":` — **obowiązkowo**, z komentarzem wyjaśniającym dlaczego.

4. **Dodatkowo** funkcja `weryfikuj_plik(sciezka)`, która w trybie `read_only` sprawdza:

   - czy arkusz ma poprawną nazwę (`rok-miesiąc`),
   - czy `max_row` jest zgodny z liczbą wierszy + nagłówek,
   - czy nagłówek jest **pogrubiony** (uwaga: w `read_only` stylów **nie ma** — więc ten test wymaga otwarcia w trybie normalnym **tylko dla jednego, małego pliku kontrolnego**; opisz to w komentarzu jako limit `read_only`).

**Wymagania dotyczące pomiarów:**

- `time.perf_counter()` dla czasów,
- wypisz, ile `lxml` jest aktywny (`from openpyxl import LXML`),
- wypisz `os.cpu_count()`.

**Testy/procedura:**

1. Uruchom z `max_workers=1` → zapisz czas.
2. Uruchom z `max_workers=min(4, cpu)` → zapisz czas.
3. Porównaj w tabeli: czasy i współczynnik równoległości dla obu przebiegów.

W `output/19_cw3_wnioski.md` odpowiedz:

1. **Jaki był współczynnik równoległości dla 4 workerów?** Czy zbliżony do 4, czy mniejszy? Jeśli mniejszy — dlaczego (dyski, pamięć, rdzenie, GIL)?
2. **Dlaczego funkcja robocza musi być na najwyższym poziomie modułu?** Wyjaśnij, odwołując się do „spawn" i tego, co robi proces potomny przy starcie.
3. **Co by się stało bez `if __name__ == "__main__":` na Windows?** Opisz krok po kroku, dlaczego doszłoby do rekurencji.
4. **Dlaczego argumenty zadania to `str` i `int`, a nie `wb`/`ws`?** Co by się stało, gdybyś spróbował przekazać `ws`? Wskaż wyjątek, jakiego byś oczekiwał.
5. **Czy `read_only` wystarczył do weryfikacji nagłówka?** Odpowiedz po wykonaniu i wskaż, który test musiałeś zrobić inaczej.
6. **Jak zmieniłby się projekt, gdyby danych było 20 000 000 wierszy (20× więcej)?** Opisz dwie–trzy zmiany architektoniczne (podpowiedź: 2.8 punkt 8, 2.11, 2.12).

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Generator raportow miesiecznych - kluczowe fragmenty."""

from __future__ import annotations

import os
import time
from collections.abc import Iterator
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WYJSCIE = ROOT / "output" / "cw3"
ROK = 2026
WIERSZY_NA_MIESIAC = 83_334


def _generator_wierszy(miesiac: int, n: int) -> Iterator[list]:
    """Generator: nie buduje listy, oddaje wiersz po wierszu (sekcja 2.12)."""
    for i in range(1, n + 1):
        # co 53. wiersz ma None - test obslugi brakow
        kwota = None if i % 53 == 0 else round(i * 0.37, 2)
        yield [f"FV/{ROK}/{miesiac:02d}/{i:07d}", i, kwota]


def _zapisz_miesiac(zadanie: tuple[str, int, int, int]) -> tuple[str, int, float, int]:
    """Dziala w OSOBNYM PROCESIE. Importy lokalnie (spawn!)."""
    katalog, rok, miesiac, n = zadanie

    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font

    t0 = time.perf_counter()

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title=f"{rok}-{miesiac:02d}")

    # === WSZYSTKO PRZED DANYMI (inaczej nie zadziala!) ===
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["C"].width = 14

    # JEDEN obiekt Font, tworzony raz (technika 2.8, punkt 5)
    naglowek_font = Font(bold=True, color="FFFFFF")
    naglowek_fill = None  # opcjonalnie PatternFill, ale budujmy minimalnie

    def _naglowek(tekst: str) -> WriteOnlyCell:
        c = WriteOnlyCell(ws, value=tekst)
        c.font = naglowek_font
        if naglowek_fill is not None:
            c.fill = naglowek_fill
        return c

    ws.append([_naglowek("Numer"), _naglowek("ID"), _naglowek("Kwota")])

    # === DANE ZE STRUMIENIA ===
    for numer, ident, kwota in _generator_wierszy(miesiac, n):
        if kwota is None:
            # brak kwoty -> zwykla wartosc None (modul 05: to nie zero!)
            ws.append([numer, ident, None])
        else:
            k = WriteOnlyCell(ws, value=kwota)
            k.number_format = "#,##0.00"
            ws.append([numer, ident, k])

    sciezka = Path(katalog) / f"raport_{rok}-{miesiac:02d}.xlsx"
    wb.save(sciezka)
    wb.close()

    t1 = time.perf_counter()
    return str(sciezka), n, t1 - t0, sciezka.stat().st_size


def _weryfikuj(sciezka: Path, rok: int, miesiac: int, n: int) -> list[str]:
    """Weryfikacja w trybie read_only. Zwraca liste problemow."""
    from openpyxl import load_workbook

    problemy: list[str] = []
    wb = load_workbook(sciezka, read_only=True)
    try:
        oczekiwana = f"{rok}-{miesiac:02d}"
        if wb.sheetnames != [oczekiwana]:
            problemy.append(f"zła nazwa arkusza: {wb.sheetnames}")

        ws = wb[oczekiwana]
        # ⚠️ W read_only max_row MOZE byc None - sprawdzamy!
        if ws.max_row is None or ws.max_column is None:
            ws.calculate_dimension(force=True)
        if ws.max_row != n + 1:      # +1 na naglowek
            problemy.append(f"max_row={ws.max_row}, oczekiwano {n + 1}")
        if ws.max_column != 3:
            problemy.append(f"max_column={ws.max_column}, oczekiwano 3")

        # Naglowek: w read_only NIE odczytamy stylu (ograniczenie!)
        # Wiec sprawdzamy tylko TRESC, nie pogrubienie:
        for wiersz in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            if list(wiersz) != ["Numer", "ID", "Kwota"]:
                problemy.append(f"zły nagłówek: {wiersz}")
    finally:
        wb.close()
    return problemy


def main() -> None:
    WYJSCIE.mkdir(parents=True, exist_ok=True)

    from openpyxl import LXML
    rdzenie = os.cpu_count() or 1

    print(f"lxml aktywny: {LXML}   rdzenie: {rdzenie}")

    zadania = [(str(WYJSCIE), ROK, m, WIERSZY_NA_MIESIAC)
               for m in range(1, 13)]

    for max_workers in (1, min(4, rdzenie)):
        print(f"\n{'=' * 70}")
        print(f"PRZEBIEG: max_workers={max_workers}")
        print("=" * 70)

        t0 = time.perf_counter()
        wyniki = []
        with ProcessPoolExecutor(max_workers=max_workers) as ex:
            for fut in as_completed([ex.submit(_zapisz_miesiac, z) for z in zadania]):
                wyniki.append(fut.result())

        t1 = time.perf_counter()
        suma_czasow = sum(c for _, _, c, _ in wyniki)
        czas_zup = t1 - t0

        print(f"  plików:          {len(wyniki)}")
        print(f"  wierszy razem:   {sum(n for _, n, _, _ in wyniki):,}")
        print(f"  rozmiar razem:   {sum(s for *_, s in wyniki) / 1_048_576:.1f} MB")
        print(f"  czas zupełny:    {czas_zup:.2f} s")
        print(f"  suma czasów:     {suma_czasow:.2f} s")
        print(f"  współczynnik:    {suma_czasow / czas_zup:.2f}x")

        # weryfikacja wszystkich plikow
        wszystkie: list[str] = []
        for i, (sciezka, n, _, _) in enumerate(sorted(wyniki), start=1):
            wszystkie += _weryfikuj(Path(sciezka), ROK, i, n)
        print(f"  problemy:        {wszystkie if wszystkie else 'brak ✓'}")

    print("\n  WERYFIKACJA 'NA OKO' (obowiazkowa):")
    print("    Otwórz 2-3 pliki i sprawdź: nagłówek pogrubiony i biały,")
    print("    wiersz 1 zamrożony, brak linii siatki, kwoty z 2 miejscami.")


if __name__ == "__main__":
    main()
    # ⚠️ if __name__ == "__main__" jest OBOWIAZKOWE:
    # Na Windows i macOS procesy potomne wystartuja z modelu "spawn",
    # czyli ZAIMPORTUJA ten modul od nowa. Bez tej ochrony kazdy proces
    # potomny uruchomilby main() -> tworzyl kolejne procesy -> rekurencja
    # bez konca (efekt: zamrozenie komputera).
```

**Kluczowe decyzje:**

- **`_generator_wierszy` jako generator, nie lista.** To sedno oszczędności. Dla 83 334 wierszy lista krotek zajęłaby ~5–8 MB (nie katastrofa), ale dla 20 mln wierszy — gigabajty. **Wzorzec musi być skalowalny od początku**, żeby nie trzeba go było przepisywać pod większe dane.

- **`None` przekazywane jawnie dla braków kwot.** Alternatywa (`ws.append([..., None])` vs pominięcie wiersza) to decyzja **biznesowa**, nie techniczna: brak kwoty to „wiersz istnieje, ale kwota nieznana". Pominięcie całego wiersza zafałszowałoby liczbę faktur. To jest dokładnie ten „pusty wiersz ≠ pusta komórka" z modułu 05.

- **Weryfikacja w `read_only` sprawdza tylko treść nagłówka, nie styl.** To jest **celowo** — i jest ilustracją ograniczenia z 2.3.1. Gdybyśmy chcieli sprawdzić `bold`, musielibyśmy otworzyć plik w trybie normalnym (i zapłacić pamięcią). Praktyczna zasada: **sprawdzaj strukturę i wartości w `read_only`, a warstwę wizualną — osobnym testem na małym pliku albo okiem.** To dokładnie to, co moduł 18 nazywał „weryfikacją obowiązkową".

- **Współczynnik równoległości liczony jawnie.** To najważniejsza metryka tego zadania. Jeśli `suma_czasów / czas_zupełny ≈ 1`, procesy **nie** pracowały równolegle (za mało rdzeni, praca I/O-bound, narzut picklowania). Jeśli `≈ 4` — działa jak trzeba. **Mierzenie tego jest obowiązkiem, nie ozdobą** — bez tego nie wiesz, czy `ProcessPoolExecutor` Ci cokolwiek dał.

- **Pętla po dwóch wartościach `max_workers`** (1 i 4) — bo **pomiar szeregowy jest punktem odniesienia**. Bez niego liczba „35 sekund" nic nie znaczy. To zasada z 2.9: **najpierw zmierz, potem zmień jedną rzecz, potem zmierz ponownie.**

- **Czego tu świadomie nie ma:** logów audytowych z modułu 18 (bo to inny temat — choć w produkcji połączyłbyś oba) i `gc.collect()` (bo przy strumieniu nie ma czego zwalniać). **Optymalizacja, która nie jest potrzebna, to złożoność, za którą ktoś zapłaci — tyle tylko, że nie Ty, a Twój następca.**

</details>

## 6. Typowe błędy i pułapki

**1. „Program zjadł cały RAM i został zabity" (objaw) → **tryb domyślny na dużym pliku: model trzyma ok. 50× rozmiaru pliku w pamięci** (2.1) (przyczyna) → **Policz budżet: rozmiar pliku × 50. Jeśli przekracza dostępny RAM — użyj `read_only`/`write_only`** (2.2–2.4). To pierwsza decyzja, którą podejmuj przy każdym zadaniu (naprawa).**

**2. „`AttributeError: ReadOnlyWorksheet object has no attribute 'iter_cols'" (objaw) → **próba iteracji po kolumnach w `read_only`.** Strumień nie ma dostępu losowego, więc `iter_cols`/`columns` **nie istnieją** (2.3.1) (przyczyna) → **Czytaj wierszami (`iter_rows(values_only=True)`) i — jeśli naprawdę potrzebujesz kolumny — zbierz ją podczas przejścia albo przekształć dane w Pythonie** (`zip(*dane)`, choć to trzyma wszystko w pamięci). Dla wybranych kolumn użyj `iter_rows(min_col=3, max_col=3, values_only=True)` (naprawa).**

**3. „`TypeError: Workbook is read-only`" (objaw) → **`wb.save()` na skoroszycie wczytanym z `read_only=True`** (2.3.1) (przyczyna) → **`read_only` jest WYŁĄCZNIE do czytania.** Jeśli chcesz zapisywać — wczytaj w trybie normalnym **albo** zapisz w **osobnym** `Workbook`. Nie ma trzeciej drogi (naprawa).**

**4. „`max_row` = `None`, a w pliku jest 200 000 wierszy" (objaw) → **`read_only` polega na metadanych `<dimension>`, a te mogą być błędne** (2.3.2) (przyczyna) → **`ws.calculate_dimension(force=True)`** (jednorazowe przejście i ustawienie wymiarów) **albo `ws.reset_dimensions()`.** Nigdy nie zakładaj, że metadane są poprawne (naprawa).**

**5. „`AttributeError: 'EmptyCell' object has no attribute 'coordinate'" (objaw) → **w `read_only` pominięte komórki są wypełniane przez `EmptyCell`, które nie ma `coordinate`/`row`/`column`** (2.3.3) (przyczyna) → **Używaj `values_only=True`** albo **licz numer wiersza i kolumny sam** przez `enumerate`. Nie buduj adresów z obiektów komórek w trybie strumieniowym (naprawa).**

**6. „`write_only` nie zapisuje nagłówka ze stylem — cały jest czarny" (objaw) → **przekazano `append([...])` ze zwykłymi stringami**, a w `write_only` **tylko `WriteOnlyCell`** może nieść styl (2.4.2) (przyczyna) → **Zbuduj komórki przez `WriteOnlyCell(ws, value=...)`, ustaw `font`/`fill` PRZED `append`.** Pamiętaj: styl nadany po `append` nie trafi do pliku (naprawa).**

**7. „W `write_only` `freeze_panes` nic nie robi" (objaw) → **ustawiono po pierwszych danych.** W XML-u `sheetViews` leży **przed** `sheetData`, a strumień już ten punkt minął (2.4.1, zasada 4) (przyczyna) → **Ustawiaj `freeze_panes`, `column_dimensions`, `sheet_view` PRZED pierwszym `append`.** Ogólna zasada: **wszystko „przed danymi" ustaw przed danymi** (naprawa).**

**8. „`WorkbookAlreadySaved` przy drugim zapisie" (objaw) → **`write_only` można zapisać TYLKO RAZ.** Po `save()` archiwum jest zamknięte i nie dopiszesz wiersza (2.4.1, zasada 3) (przyczyna) → **Zbuduj cały skoroszyt i zapisz raz.** Jeśli potrzebujesz dwóch plików — zbuduj dwa `Workbook`. Nie próbuj „dopisywać po zapisie" (naprawa).**

**9. „`ws['A1'] = 1` w `write_only` nic nie robi / rzuca błąd" (objaw) → **w `write_only` nie ma zapisu po komórkach** — tylko `append()` (2.4.1, zasada 2) (przyczyna) → **Używaj wyłącznie `ws.append([...])`.** Jeśli potrzebujesz dostępu losowego — `write_only` jest niewłaściwym trybem; przejdź na tryb normalny i zaakceptuj koszt pamięci (naprawa).**

**10. „Na serwerze wolniej niż u mnie, choć ten sam kod" (objaw) → **`lxml` nieaktywny** (nie zainstalowany **albo** wyłączony przez `OPENPYXL_LXML`) (2.7) (przyczyna) → **Testuj jawnie: `from openpyxl import LXML; assert LXML`.** Dodaj ten assert do logu startowego albo do CI. Sprawdź `os.environ.get("OPENPYXL_LXML")`. Jeśli `lxml` jest zainstalowany, ale wyłączony — znajdź, kto ustawia tę zmienną (naprawa).**

**11. „Wolno, a mam 8 rdzeni i ustawiłem `max_workers=50`" (objaw) → **za dużo procesów przy pracy obciążającej pamięć → swap → paradoksalne spowolnienie.** Do tego `ProcessPoolExecutor` ma narzut na tworzenie procesów i picklowanie (2.10) (przyczyna) → **Zacznij od `min(4, os.cpu_count())`, zmierz, zwiększaj ostrożnie.** Mierz **współczynnik równoległości** (`suma czasów / czas zupełny`), a nie tylko czas (naprawa).**

**12. „`ProcessPoolExecutor` zamroził komputer / wciąż tworzy procesy" (objaw) → **brak `if __name__ == "__main__":`.** Proces potomny w modelu „spawn" **ponownie importuje moduł**, co uruchamia `main()` i tworzy kolejne procesy — rekurencja (2.10) (przyczyna) → **Zawsze chroń punkt wejścia `if __name__ == "__main__":`.** Dodatkowo: funkcja robocza musi być **na najwyższym poziomie modułu** (naprawa).**

**13. „`AttributeError`/`NameError` w funkcji roboczej: `openpyxl` nie jest zdefiniowany" (objaw) → **proces potomny nie dziedziczy importów** (model „spawn") (2.10, przykład 4) (przyczyna) → **Importuj lokalnie, WEWNĄTRZ funkcji roboczej** (naprawa).**

**14. „`TypeError: cannot pickle ...` przy wysyłaniu zadania do procesu" (objaw) → **próbujesz przekazać `wb`, `ws` albo `Cell` do procesu potomnego.** Te obiekty nie są „picklowalne" i nie mogą przejść granicy procesu (2.10) (przyczyna) → **Przekazuj tylko proste typy: ścieżki, liczby, krotki, dataclass.** Wzorzec: proces dostaje **dane wejściowe**, nie „stan biblioteki" (naprawa).**

**15. „Mój kod zamiast czytać dane tworzy 10 milionów komórek" (objaw) → **`ws.cell(row, column)` w pętli skanującej** tworzy komórkę, nawet jeśli nie przypisujesz wartości (2.1) (przyczyna) → **`ws.iter_rows(values_only=True)`** — żadnych obiektów `Cell`. Jeśli musisz po komórkach — ogranicz `min_col`/`max_col` do realnego zakresu (naprawa).**

**16. „`ws.max_row` w warunku pętli → strasznie wolno" (objaw) → **`max_row` to właściwość licząca wymiary**; w pętli wywołujesz ją 100 000 razy (2.8, punkt 6) (przyczyna) → **Wyciągnij do zmiennej przed pętlą:** `ostatni = ws.max_row` (naprawa).**

**17. „Raport otwiera się 3 minuty" (objaw) → **formatowanie warunkowe na setkach tysięcy komórek** (np. `A:A` dla 1M wierszy). openpyxl tego nie liczy, ale **Excel tak** — przy każdym otwarciu (2.8, punkt 4) (przyczyna) → **Ogranicz `sqref` do zakresu faktycznych danych lub do kolumny wynikowej.** Jeśli chcesz wygląd na całej kolumnie — rozważ inną prezentację (naprawa).**

**18. „Analiza pokazała brak walidacji, a w Excelu je widzę" (objaw) → **otwarto plik z `read_only=True` w narzędziu diagnostycznym.** Reader pomija DV, CF, obrazy i tabele (2.3.1) (przyczyna) → **Do diagnostyki używaj trybu normalnego.** `read_only` jest do **czytania danych**, nie do inspekcji zawartości. To powtórka pułapki z modułu 18 (naprawa).**

**19. „Nie wiem, czy mój kod jest wolny, czy tylko mi się wydaje" (objaw) → **brak pomiaru.** „Wydaje się wolne" to nie jest informacja (2.9) (przyczyna) → **`time.perf_counter()` + `tracemalloc` na realnym pliku.** Zbuduj tabelę „operacja → czas" (ćwiczenie 1) raz i traktuj ją jako punkt odniesienia (naprawa).**

**20. „Optymalizowałem pętlę przez tydzień, a program jest tak samo wolny" (objaw) → **optymalizacja bez profilowania.** Poprawiłeś 3% czasu, a 97% idzie gdzie indziej (`cProfile` by to pokazał w 20 sekund) (2.9) (przyczyna) → **Profiluj przed optymalizacją.** `cProfile` → znajdź najdroższą funkcję → zmień **jedną** rzecz → zmierz ponownie (naprawa).**

**21. „Wygenerowałem plik 3 000 000 wierszy i Excel go nie otwiera" (objaw) → **przekroczyłeś limit formatu: 1 048 576 wierszy × 16 384 kolumn** (2.11) (przyczyna) → **Podziel na pliki** albo — lepiej — **nie trzymaj takich danych w Excelu.** To sygnał z 2.11: dane należą do bazy/Parquet, a Excel dostaje próbkę i podsumowanie (naprawa).**

**22. „Zapomniałem `wb.close()` i po tysiącu raportów padł system" (objaw) → **w `read_only`/`write_only` otwarte są uchwyty do archiwum ZIP**; brak `close()` = `Too many open files` (2.3, przykład 2) (przyczyna) → **Zawsze `try/finally: wb.close()`.** W historii openpyxl były konkretne błędy zwalniania uchwytów (#2122, #2149 — naprawione w 3.1.3), ale zasada chroni niezależnie od wersji (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Nie ma opcji „wszystko naraz i tanio".** openpyxl w trybie domyślnym daje pełny dostęp losowy **i** wszystkie funkcje (wykresy, tabele, formatowanie), ale kosztuje **ok. 50× rozmiar pliku w RAM** — 2,5 GB na plik 50 MB. Tryby optymalizowane dają **stałą, minimalną pamięć**, ale **odbierają funkcje**. Każdy wybór jest kompromisem, a pierwsze pytanie przy każdym zadaniu brzmi: **„jaki jest mój budżet pamięci?"**.

2. **`read_only` to szczelina w drzwiach, `write_only` to prasa drukarska.** Czytasz **wierszami, w przód, raz** (`iter_rows(values_only=True)`) — nie ma `iter_cols`, nie ma `ws["A1"]`, nie ma stylów w komórkach, a wymiary mogą kłamać (`calculate_dimension(force=True)`). Piszesz **tylko przez `append`**, styl nadajesz **`WriteOnlyCell` przed** `append`, wszystko co „przed danymi" (freeze, szerokości) ustawiasz **przed pierwszym `append`**, a `save()` działa **tylko raz**. Oba tryby wymagają **`close()`**.

3. **`lxml` to jedna instalacja i dwa razy szybciej — ale trzeba to zweryfikować.** `pip install lxml` wystarcza, bo openpyxl wybiera backend **przy imporcie**. Ale flaga `LXML` zależy też od `OPENPYXL_LXML` i wersji `lxml (>= 3.3.1)`. **Testuj jawnie** (`from openpyxl import LXML; assert LXML`) w CI i w logu startowym, bo wyłączony `lxml` to **cicha regresja wydajności** — program działa, tylko wolniej, i nikt tego nie zauważy.

4. **Mierz, nie zgaduj — i równolegle tylko na osobnych plikach.** `time.perf_counter` (czas), `tracemalloc` (`peak`), `cProfile` (gdzie idzie czas), `memory_profiler` (RSS procesu). **`ProcessPoolExecutor` na `read_only`/osobnych plikach** — tak; **współdzielenie jednego `wb` między procesy** — nie. Pamiętaj o trzech warunkach: funkcja **na najwyższym poziomie modułu**, `if __name__ == "__main__":`, argumenty **picklowalne**. Mierz **współczynnik równoległości**, bo „35 sekund" nic nie znaczy bez punktu odniesienia.

5. **Nie wozi się węgla ferrari.** Excel to **warstwa prezentacji**, nie magazyn ani silnik obliczeń. Gdy danych jest dużo i są przetwarzane wielokrotnie → baza danych albo Parquet. Excel dostaje **mały, zagregowany wynik**. Wzorzec docelowy to **strumień od źródła do pliku** (`yield` z kursora → `append` → dysk), bez momentu, w którym cały zbiór istnieje w RAM. **Dane zostają tam, gdzie należą, a openpyxl robi to, w czym jest dobry — ładny plik wychodzący do ludzi.**

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY I DIAGNOSTYKA
# ==================================================================
import os, time, tracemalloc, cProfile, pstats, io
from pathlib import Path

from openpyxl import Workbook, load_workbook, LXML, DEFUSEDXML
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.exceptions import WorkbookAlreadySaved

print("openpyxl :", __import__("openpyxl").__version__)
print("lxml     :", LXML)          # True = przyspieszenie aktywne!
print("OPENPYXL_LXML env:", os.environ.get("OPENPYXL_LXML"))

# ==================================================================
# 2. CZYTANIE DUZYCH PLIKOW - read_only
# ==================================================================
wb = load_workbook("duzy.xlsx", read_only=True)
try:
    ws = wb["Dane"]

    # wymiary moga byc None - forsuj skan, jesli trzeba
    if ws.max_row is None or ws.max_column is None:
        ws.calculate_dimension(force=True)     # jednorazowe przejscie
        # alternatywa: ws.reset_dimensions()

    # NAJSZYBSZA SCIEZKA: wierszami, tylko wartosci
    for row in ws.iter_rows(min_row=2, values_only=True):
        identyfikator, nazwa, kwota = row
        if kwota is None:               # pusta komorka != 0!
            continue
        ...
finally:
    wb.close()                          # OBOWIAZKOWE (uchwyt ZIP)

# CZEGO NIE MA W read_only:
#   ws.iter_cols(), ws.columns        -> AttributeError
#   ws["A1"] = 5                      -> TypeError: does not support item assignment
#   wb.save(...)                      -> TypeError: Workbook is read-only
#   ws.merged_cells, ws.column_dimensions, ws.page_setup, ws.protection
#   cell.font / cell.fill / cell.comment / cell.hyperlink

# ==================================================================
# 3. PISANIE DUZYCH PLIKOW - write_only
# ==================================================================
wb = Workbook(write_only=True)
ws = wb.create_sheet(title="Dane")      # NIE ma arkusza domyslnego!

# ⚠️ WSZYSTKO PRZED DANYMI (inaczej zignorowane!):
ws.freeze_panes = "A2"
ws.sheet_view.showGridLines = False
ws.column_dimensions["B"].width = 25

# STYL: tworz obiekt Font RAZ, nie w petli
NAGLOWEK_FONT = Font(bold=True, color="FFFFFF")
NAGLOWEK_FILL = PatternFill("solid", fgColor="1F4E78")

def naglowek(tekst):
    c = WriteOnlyCell(ws, value=tekst)
    c.font = NAGLOWEK_FONT
    c.fill = NAGLOWEK_FILL
    return c

ws.append([naglowek("ID"), naglowek("Nazwa"), naglowek("Kwota")])

def strumien(n):                        # generator, nie lista!
    for i in range(1, n + 1):
        yield [i, f"Klient {i}", i * 1.5]

for wiersz in strumien(1_000_000):
    ws.append(wiersz)                   # tylko append()!

wb.save("output/duzy.xlsx")             # TYLKO RAZ!
wb.close()
# Po save() -> WorkbookAlreadySaved przy kazdej kolejnej probie.

# ==================================================================
# 4. STREAMING Z BAZY (wzorzec z 2.12)
# ==================================================================
def wiersze_z_bazy(kursor):
    while True:
        w = kursor.fetchone()
        if w is None:
            return
        yield list(w)

# zapisz(wiersze_z_bazy(kursor), "raport.xlsx")  <- pamiec O(1)

# ==================================================================
# 5. POMIAR: CZAS + PAMIEC
# ==================================================================
tracemalloc.start()
t0 = time.perf_counter()

# ... mierzony kod ...

t1 = time.perf_counter()
current, peak = tracemalloc.get_traced_memory()
tracemalloc.stop()

print(f"czas: {t1 - t0:.2f} s   szczyt: {peak / 1_048_576:.1f} MB")

# ==================================================================
# 6. PROFILOWANIE WYWOŁAN (gdy nie wiesz, gdzie idzie czas)
# ==================================================================
pr = cProfile.Profile(); pr.enable()
# ... kod ...
pr.disable()
s = io.StringIO()
pstats.Stats(pr, stream=s).sort_stats("cumulative").print_stats(15)
print(s.getvalue())

# ==================================================================
# 7. ROWNOLEGLOSC (osobne pliki / osobne zakresy!)
# ==================================================================
from concurrent.futures import ProcessPoolExecutor, as_completed

def _pracuj(zadanie):                   # NAJWYZSZY poziom modulu!
    sciezka, min_row, max_row = zadanie
    from openpyxl import load_workbook   # import LOKALNIE (spawn!)
    wb = load_workbook(sciezka, read_only=True)
    try:
        suma = 0
        for row in wb["Dane"].iter_rows(min_row=min_row, max_row=max_row,
                                        values_only=True):
            if row[2] is not None:
                suma += row[2]
        return suma
    finally:
        wb.close()

if __name__ == "__main__":              # OBOWIAZKOWE!
    zadania = [("plik.xlsx", 2, 100_000), ("plik.xlsx", 100_001, 200_000)]
    with ProcessPoolExecutor(max_workers=min(4, os.cpu_count() or 1)) as ex:
        for fut in as_completed([ex.submit(_pracuj, z) for z in zadania]):
            print(fut.result())
    # MIERZ wspolczynnik: suma_czasow / czas_zupelny
    # ~1 = szeregowo, ~4 = rownolegle

# ==================================================================
# 8. LXML: WYLACZENIE BEZ ODINSTALOWANIA (diagnostyka)
# ==================================================================
# bash:  OPENPYXL_LXML=False python skrypt.py
# -> openpyxl.LXML == False, backend = xml.etree.ElementTree

# ==================================================================
# 9. DOCKER (wheele lxml dzialaja bez gcc na x86_64)
# ==================================================================
# FROM python:3.12-slim
# WORKDIR /app
# COPY requirements.txt .
# RUN pip install --no-cache-dir -r requirements.txt   # openpyxl==3.1.5, lxml
# COPY . .
# Weryfikacja: python -c "from openpyxl import LXML; assert LXML"

# ==================================================================
# 10. PIEC ZDAN DO ZAPAMIETANIA
# ==================================================================
# 1) Pamiec ~50x rozmiar pliku w trybie domyslnym. Policz budzet PRZED.
# 2) read_only: wiersze w przod, values_only=True, close() w finally.
# 3) write_only: create_sheet, freeze/szerokosci PRZED danymi, save RAZ.
# 4) Mierz (perf_counter + tracemalloc), potem zmien JEDNA rzecz.
# 5) Nie woz sie wegla ferrari: duze dane -> baza/Parquet, Excel = prezentacja.
```

## 9. Co dalej

Do tego momentu kurs przebiegał przez **coraz większą skalę**: od pierwszej komórki (moduł 04), przez formatowanie (09–14), wzbogacanie (15–17), aż do modyfikacji istniejących plików (18) i wydajności (19). Zaczynałeś od „jak zapisać jeden wyraz", a skończyłeś na „jak wygenerować milion wierszy w dwunastu plikach, nie zabijając serwera".

I tu pojawia się pytanie, które wcześniej nie miało sensu, a teraz jest najważniejsze:

> **Skąd wiesz, że to wszystko działa? I skąd wiesz, że po zmianie w kodzie nadal działa?**

Bo przecież w module 18 nauczyłeś się, że **utrata danych jest cicha** — plik się otwiera, nic nie krzyczy, a połowa warstwy wizualnej zniknęła. A w tym module widziałeś, że **regresja wydajności też jest cicha** — wyłączony `lxml` nie daje błędu, tylko „na serwerze jest wolniej". Dwa rodzaje cichych awarii, których nie złapie żaden `try/except`. **Trzeba je złapać testem.**

W **module 20** zajmiemy się właśnie tym — i będzie to moduł, który zmieni Twój sposób pracy bardziej niż jakikolwiek inny. Zobaczysz:

- **`BytesIO` jako Twój najlepszy przyjaciel.** Pamiętasz, że `wb.save()` przyjmuje **obiekt plikopodobny**, nie tylko ścieżkę? To znaczy, że możesz **zapisać skoroszyt do pamięci, wczytać go z powrotem z pamięci i sprawdzić — bez ani jednego pliku na dysku.** Testy plików binarnych przestaną wyglądać groźnie.
- **Testy w `pytest`** z fixtures (`tmp_path`, fabryka skoroszytu), asercjami na **wartościach, typach, formatach i liczbie reguł** — a nie na bajtach.
- **„Odcisk skoroszytu"** z modułu 18 jako w pełni automatyczny test regresji: raz nagrany punkt odniesienia pilnuje, żeby kolejna zmiana w kodzie nie zgubiła części pliku.
- **Testy wydajności jako zabezpieczenie**, nie jako ciekawostka. Skoro wiesz już, jak mierzyć (`perf_counter`, `tracemalloc`), to teraz zamienimy pomiar w **twardy budżet**: „100 000 wierszy musi powstać w mniej niż 10 sekund" — i test, który to egzekwuje na każdym commicie.
- **Testy niezmienników dla czyścicieli danych** (`hypothesis`): dowolne dane wejściowe, jedna obietnica — „żadna wartość nie jest tekstem, jeśli miała być liczbą".
- **CI/CD bez Excela.** Zobaczysz, że **Excel nie jest do niczego potrzebny** do uruchomienia całego zestawu testów — bo openpyxl nie uruchamia Excela (moduł 01). Testy działają w kontenerze, na serwerze buildowym, na GitHub Actions.

I najważniejsza rzecz, którą moduł 20 nazwie wprost: **oddzielenie „co ma być w raporcie" od „jak zapisać plik".** Bo tylko wtedy, gdy logika jest oddzielona, da się ją przetestować — bez uruchamiania openpyxl, bez plików, bez czekania. To będzie też pierwszy krok do modułów 22–26, gdzie cały ten `Workbook`, `Worksheet` i `WriteOnlyCell` schowają się za **jedną, własną fasadą** — i staną się szczegółem implementacji, a nie centrum Twojego programu.

Zaczynamy od pytania praktycznego, które słyszałeś pewnie nie raz: **„Skąd mam wiedzieć, że ten raport jest poprawny?"** Moduł 20 odpowie: *„Bo go przetestowałeś. Osiemnastoma testami. W pamięci. W trzy sekundy."*