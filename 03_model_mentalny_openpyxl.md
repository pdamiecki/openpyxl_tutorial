# Moduł 03 — Model mentalny openpyxl

> **Część:** 0 — Fundamenty i model mentalny · **Poziom:** ⭐ · **Wymaga:** modułów 00–02

## 0. W tym module nauczysz się

- Rozumiesz openpyxl jako **graf obiektów w pamięci RAM**: `Workbook` → `Worksheet` → `Cell` → współdzielone obiekty stylów. Umiesz narysować ten graf dla dowolnego pliku i wskazać w nim palcem, gdzie kończy się Python, a zaczyna plik na dysku.
- Rozumiesz, dlaczego `row=1, column=1` to `A1`, a `row=0` nie istnieje — i dlaczego ten sam ekosystem używa **dwóch systemów indeksowania jednocześnie**: komórki liczą od 1, a indeksy stylów od 0.
- Rozumiesz **cykl życia** modelu: `load_workbook()` składa cały skoroszyt do pamięci, modyfikujesz model obiektami Pythona, a `save()` **przepisuje cały pakiet od nowa**. Umiesz wyjaśnić, dlaczego openpyxl nie edytuje pliku w miejscu i jakie są z tego trzy praktyczne konsekwencje.
- Umiesz rozróżnić i wybrać **trzy tryby pracy**: normalny (swobodny dostęp losowy), `read_only=True` (czytanie przez szczelinę) i `Workbook(write_only=True)` (prasa drukarska).
- Znasz tabelę **„potrafi / potrafi częściowo / gubi przy zapisie"** i umiesz się do niej odwołać, zamiast zgadywać, czy dana funkcja Excela przetrwa zapis.
- Rozumiesz, dlaczego `cell.font.bold = True` **nie działa** — i to nie dlatego, że „openpyxl ma błąd", ale dlatego, że style są etykietami współdzielonymi przez tysiące komórek, a komórka trzyma numer etykiety, nie etykietę.
- Umiesz wyjaśnić, dlaczego `wb` w pamięci i plik na dysku to **dwa różne światy**, i dlaczego program, który kończy się bez `save()`, nie zostawia po sobie żadnego śladu.

## 1. Intuicja i analogia

Zacznij od pytania, które pozornie nie ma z openpyxl nic wspólnego:

> Co się dzieje, kiedy uruchamiasz w Pythonie `wb.save("raport.xlsx")`?

Większość początkujących odpowiada instynktownie: „openpyxl otwiera plik i wpisuje do niego dane". To brzmi naturalnie, bo tak działają edytory tekstu, bazy danych i arkusze kalkulacyjne. I to jest **nieprawda**. Ta nieprawda jest źródłem większości problemów, z którymi zetkniesz się w kursie.

Prawda jest inna: **openpyxl w ogóle nie dotyka pliku, dopóki go nie zapiszesz.** Przez cały czas pracy masz w pamięci RAM pełny, samodzielny model skoroszytu — kompletny, niezależny od jakiegokolwiek pliku. Plik na dysku pojawia się dopiero w momencie `save()`, i pojawia się jako **nowy plik zbudowany od zera** z tego modelu.

### Analogia główna: notatnik i pieczątki

Wyobraź sobie, że prowadzisz dokumentację w zeszycie.

**`Workbook` to cały zeszyt** — okładka, liczba kartek, tytuł na grzbiecie, informacja, która kartka była otwarta ostatnio, nazwy zdefiniowane przez Ciebie na wewnętrznej stronie okładki. Jedna rzecz: cały zeszyt.

**`Worksheet` to kartka** — jedna fizyczna strona w tym zeszycie. Ma nazwę wpisaną na zakładce (u góry), ma wymiary, ma swoje ustawienia wydruku, ma swoje tabele i reguły. Kartek może być wiele, są ułożone w konkretnej kolejności i każda wie, do którego zeszytu należy.

**`Cell` to kratka** — jedno pole na kartce, wyznaczone przez przecięcie wiersza i kolumny. Kratka ma treść (wartość), ma informację o tym, jakim jest typem danych, i **ma numer pieczątki**.

I tu pojawia się rzecz, na której zbudowany jest cały ten moduł:

> **Kratka nie ma wyglądu. Kratka ma numer pieczątki.**

Nie ma koloru. Nie ma czcionki. Nie ma obramowania. Ma **jedną liczbę**: „użyj pieczątki numer 7". Pieczątki są przechowywane osobno, w rejestrze — jednym, wspólnym dla całego zeszytu. Pieczątka numer 7 może być przybita w tysiącu kratek jednocześnie.

To wyjaśnia dwie rzeczy naraz.

**Po pierwsze — dlaczego w pliku Excel wygląd jest zdeduplikowany.** Jeśli tysiąc kratek nosi pieczątkę numer 7, to opis tej pieczątki (pogrubiona Calibri 12, biała, na tle `FF1F4E78`) leży w zeszycie **raz**. Sprawdziłeś to w module 02: komórka ma `s="7"`, a opis jest w `xl/styles.xml`. Teraz wiesz, że to nie jest artefakt formatu pliku — to jest struktura modelu, którą openpyxl odzwierciedla jeden do jednego.

**Po drugie — dlaczego stylów nie wolno mutować.** Wyobraź sobie, że pieczątka numer 7 leży w szufladzie i ktoś podchodzi do niej i **poprawia ją długopisem**. Efekt: tysiąc kratek zmienia wygląd, choć nikt nie prosił. Katastrofa. Dlatego openpyxl przyjmuje konwencję: **obiekt stylu, który dostajesz z komórki, nie jest pieczątką z szuflady.** Jest jej kserokopią. Możesz po niej mazać do woli — nikt tego nie przeczyta.

I stąd bierze się słynne:

```python
ws["A1"].font.bold = True   # NIC NIE ROBI. Cicho. Bez błędu.
```

Nie dostajesz błędu, bo w Pythonie ten obiekt naprawdę da się zmodyfikować. Zmieniasz kserokopię i wyrzucasz ją do kosza w tej samej linijce. **Cichy brak efektu** — najgorszy rodzaj błędu, bo nie ma czego debugować.

Jak to się robi poprawnie? Są dwa sposoby i oba wynikają wprost z analogii:

```python
# Sposób 1: zaprojektuj nową pieczątkę i przybij ją.
ws["A1"].font = Font(bold=True)

# Sposób 2: zrób kserokopię starej pieczątki, dopisz na niej długopisem,
#           a potem przybij kserokopię jako nową pieczątkę.
nowa = copy(ws["A1"].font)
nowa.bold = True
ws["A1"].font = nowa
```

Sposób 2 jest tym, którego użyjesz, gdy chcesz zmienić **jedną** cechę, zachowując pozostałe. To dokładnie analogia z modułu 02: pieczątki się nie poprawia, pieczątkę się projektuje od nowa albo kopiuje i dopisuje na kopii.

### Analogia cyklu życia: fotokopia i przepisanie na czysto

Wróćmy do pytania o `save()`. Skoro model jest samodzielny i kompletny, to co robi `save()`?

**`save()` przepisuje cały dokument na czysto.**

Wyobraź sobie, że dostałeś od kogoś rękopis. Nie wolno Ci pisać w oryginale. Siadasz więc z czystą kartką i **przepisujesz tekst litera po literze**. Przy okazji dopisujesz swoje poprawki. Oddajesz nową kartkę — oryginał zostaje w szufladzie.

I teraz najważniejsze: **przepiszesz tylko to, co umiesz przeczytać.**

Jeśli oryginał zawiera szkic ołówkiem na marginesie, którego nie rozumiesz — nie przepiszesz go. Nie dlatego, że go zniszczyłeś. Dlatego, że **nie było czego przepisać**. W miejscu, w którym powinien być, jest puste miejsce i nie masz powodu przypuszczać, że było tam coś więcej.

To jest dokładnie sytuacja sparkline'ów, slicerów, kształtów i Power Query w openpyxl. Model openpyxl nie ma dla nich odpowiednika — komórki, do których są przypisane, są w modelu, ale same sparkline'y nie. Przy `save()` nie ma czego przepisać.

> **Nie jest to usterka openpyxl. Jest to konsekwencja tego, że openpyxl zbudowany jest jako konwerter formatu, a nie jako klon Excela.** Excel ma model, który obejmuje sparkline'y. openpyxl ma model, który ich nie obejmuje. Oba czytają ten sam format pliku.

Z tej analogii wynikają trzy praktyczne konsekwencje, które będą wracać przez cały kurs:

1. **Nic nie trafia na dysk, dopóki nie wywołasz `save()`.** Model żyje w RAM. Możesz go zmieniać godzinę, możesz go zniszczyć, możesz zapomnieć zapisać — plik na dysku nie drgnie.
2. **`save()` nie dopisuje. `save()` przepisuje.** Jeśli wskażesz jako cel plik, który już istnieje, zostanie on **nadpisany w całości, bez ostrzeżenia**. Zawartość, której nie było w modelu, przepada — także wtedy, gdy model powstał z tego samego pliku (bo nie objął wszystkiego).
3. **Zmiany w modelu nie są widoczne w pliku, a zmiany w pliku nie są widoczne w modelu.** Wczytanie pliku to jednorazowy akt. Od tego momentu model i plik rozjeżdżają się w dwie niezależne historie.

### Analogia indeksowania: numeracja pięter

Zostaje jeszcze jedna subtelność, bez której nie da się napisać poprawnej pętli.

W Pythonie listy liczą od zera. `lista[0]` to pierwszy element. To nawyk tak silny, że ludzie przenoszą go na arkusze — i piszą `ws.cell(row=0, column=0)`, żeby sięgnąć do komórki „pierwszej".

To nie zadziała. Excel numeruje od 1.

**Analogia: numeracja pięter.** W Polsce i w większości Europy parter to „0", a pierwsze piętro jest nad nim. W USA i Kanadzie **parter jest pierwszym piętrem** — wsiadasz do windy i wciskasz „1". Excel liczy jak w Ameryce: pierwszy wiersz to wiersz 1, pierwsza kolumna to kolumna 1, komórka w lewym górnym narożniku to `A1`.

Ale uwaga — i tu rzecz, która odróżnia osobę rozumiejącą openpyxl od osoby, która go tylko używa:

> **Ten sam model używa obu systemów jednocześnie.** Adresy i wymiary liczą od 1. Ale indeksy stylów (`fontId`, `fillId`) oraz indeksy wpisów w rejestrach liczą od 0.

Zobaczyłeś to w module 02 w jednej linijce XML: `r="C1"` (od 1), `s="1"` (od 0), `<v>0</v>` (od 0). Ten sam dualizm siedzi w modelu openpyxl, tylko zapisany w Pythonie:

```python
ws["A1"].row            # 1   <- od 1
ws["A1"].column         # 1   <- od 1
ws["A1"]._style.fontId  # 1   <- od 0! To DRUGA czcionka w rejestrze
```

Konsekwencja praktyczna: **nigdy nie zakładaj, że jeśli coś jest „pierwsze", to znaczy „indeks 0" — sprawdź, o którym systemie mówisz.** W module 07, gdy będziesz wyliczał zakresy i wymiary, ta reguła uratuje Cię przed wieloma godzinami debugowania.

## 2. Teoria

### 2.1. Graf obiektów — cała mapa w jednym rysunku

Zanim przejdziemy do szczegółów, obejrzyj całość. To jest mapa, do której będziesz wracał przez cały kurs.

```text
wb: Workbook
│
├── wb.sheetnames ──────────► ['Dane', 'Podsumowanie']      (lista TEKSTÓW)
│
├── wb.worksheets ──────────► [Worksheet('Dane'), Worksheet('Podsumowanie')]
│                               │                    │
│                               │                    └── ws.title = 'Podsumowanie'
│                               │
│                               ├── ws.parent ──────► z powrotem do wb
│                               │
│                               ├── ws._cells ──────► {(1, 1): Cell, (3, 2): Cell, ...}
│                               │                       │  (klucz: krotka (row, column),
│                               │                       │   OBIE liczby od 1!)
│                               │                       │
│                               │                       ├── cell.value       = 'Ala'
│                               │                       ├── cell.data_type   = 's' | 'n' | 'b' | 'd' | 'f' | 'e'
│                               │                       ├── cell.row         = 1
│                               │                       ├── cell.column      = 1
│                               │                       ├── cell.coordinate  = 'A1'
│                               │                       ├── cell.column_letter = 'A'
│                               │                       ├── cell.parent ────► z powrotem do ws
│                               │                       │
│                               │                       └── cell._style ─────► StyleArray
│                               │                                                [fontId,
│                               │                                                 fillId,
│                               │                                                 borderId,
│                               │                                                 numFmtId,
│                               │                                                 ...]
│                               │                                    ↑ INDEKSY, nie obiekty!
│                               │
│                               ├── ws.tables                    (tabele Excela)
│                               ├── ws.conditional_formatting    (reguły warunkowe)
│                               ├── ws.data_validations          (walidacje danych)
│                               ├── ws._charts / ws._images      (warstwa graficzna)
│                               ├── ws.column_dimensions         (szerokości kolumn)
│                               ├── ws.row_dimensions            (wysokości wierszy)
│                               ├── ws.merged_cells              (scalenia)
│                               └── ws.freeze_panes               (zamrożone okna)
│
├── wb.defined_names ───────► nazwy zdefiniowane
├── wb.properties ──────────► autor, tytuł, daty utworzenia
├── wb.security ────────────► ochrona struktury skoroszytu
├── wb.calculation ─────────► ustawienia przeliczania (fullCalcOnLoad)
└── wb.epoch ───────────────► epoka dat (1900 lub 1904)
│
└── REJESTRY STYLÓW ─────────► WSPÓLNE DLA WSZYSTKICH ARKUSZY
    ├── fonts                ──► [Font#0, Font#1, Font#2, ...]  ← tu wskazują fontId
    ├── fills                ──► [Fill#0, Fill#1, Fill#2, ...]  ← tu wskazują fillId
    ├── borders              ──► [Border#0, Border#1, ...]
    ├── number_formats
    ├── named_styles
    └── differential_styles  ──► style reguł warunkowych (te z <dxfs>)
```

Cztery rzeczy, które warto zauważyć na tym rysunku:

**Każdy węzeł zna swojego rodzica.** `ws.parent` to `Workbook`, `cell.parent` to `Worksheet`. To nie jest ozdoba — to jest mechanizm, przez który openpyxl wie, do którego rejestru stylów ma trafić nowa czcionka. Kiedy piszesz `cell.font = Font(bold=True)`, komórka idzie do arkusza, arkusz do skoroszytu, i dopiero tam znajduje rejestr.

**Komórki nie są listą, są słownikiem.** `ws._cells` to `dict` kluczowany krotką `(row, column)`. To wyjaśnia, dlaczego dostęp do `ws["Z999"]` jest natychmiastowy nawet w pustym arkuszu — openpyxl nie iteruje przez komórki, tylko robi `dict.get((999, 26))`. Klucz `(999, 26)` to wiersz 999, kolumna 26 (= Z). **Kolejność w krotce to wiersz, potem kolumna** — nie odwrotnie, co brzmi naturalnie po notacji `A1`, ale jest odwrotnie niż w `(x, y)` z geometrii.

**Komórka ma tablicę indeksów, nie obiekty stylu.** `cell._style` to nie `Font`, nie `Fill`, nie `Border`. To mała tablica liczb. Ta sama prawda, którą widziałeś w XML jako `s="2"`, tylko w Pythonie.

**Rejestry stylów są na poziomie skoroszytu.** Nie na poziomie arkusza. To znaczy, że czcionka użyta na arkuszu „Dane" i identyczna czcionka użyta na arkuszu „Podsumowanie" to **jeden wpis w rejestrze**. Zaleta: mniejszy plik. Pułapka: dzielony stan, o którym trzeba myśleć.

> Uwaga o wersjach: konkretne nazwy atrybutów wewnętrznych (`wb._fonts`, `wb._fills`, `ws._cells`, `cell._style`) są **API prywatnym** — mogą się zmienić bez ostrzeżenia. W kursie używamy ich wyłącznie do **podglądania** i do nauki. W kodzie produkcyjnym nie powinieneś się do nich odwoływać. Kiedy w kursie piszę „zajrzyj do `_fonts`", to znaczy: „jeśli w Twojej wersji taki atrybut istnieje, podejrzyj go, żeby zrozumieć mechanizm". Jeśli nie istnieje — pomiń podgląd. **Efekt widoczny w pliku jest ten sam niezależnie od wersji**, i ten efekt jest ostatecznym dowodem.

### 2.2. Adresowanie: dwie notacje, jeden świat

W arkuszu do tej samej komórki można trafić na dwa sposoby:

```python
ws["A1"]                       # notacja A1 - dla człowieka
ws.cell(row=1, column=1)       # notacja liczbowa - dla pętli i obliczeń
```

Obie zwracają **ten sam obiekt** `Cell`. Ale mają różne zastosowania i różne pułapki.

**Notacja A1 jest dla człowieka.** Gdy piszesz kod, który ma być czytelny, gdy odwołujesz się do konkretnej, nazwanej komórki (`ws["B2"]` to cena, `ws["D7"]` to podsumowanie) — używaj A1. Czytelnik Twojego kodu od razu wie, o którą komórkę chodzi, bez liczenia kolumn.

**Notacja liczbowa jest dla maszyn.** Gdy piszesz pętlę, gdy adres jest wynikiem obliczenia, gdy iterujesz po wierszach — używaj `ws.cell(row=..., column=...)`. Konwertowanie liczby na literę i z powrotem w każdej iteracji to marnotrawstwo, a `f"{letter}{row}"` w łańcuchu to proszenie się o błąd.

Do konwersji między notacjami służą narzędzia:

```python
from openpyxl.utils import get_column_letter, column_index_from_string

get_column_letter(1)                # 'A'
get_column_letter(26)               # 'Z'
get_column_letter(27)               # 'AA'
get_column_letter(28)               # 'AB'

column_index_from_string("A")       # 1
column_index_from_string("Z")       # 26
column_index_from_string("AA")      # 27
column_index_from_string("AB")      # 28
```

Zwróć uwagę na kierunek: `get_column_letter` przyjmuje **liczbę od 1** i zwraca literę. `column_index_from_string` przyjmuje literę i zwraca **liczbę od 1**. Oba są 1-indeksowane, oba są odwracalne, i oba **rzucają wyjątek przy 0** — bo kolumna 0 nie istnieje.

I jeszcze jedno narzędzie, które w module 07 użyjesz do parsowania zakresów:

```python
from openpyxl.utils.cell import (
    coordinate_to_tuple, coordinate_from_string, range_boundaries, quote_sheetname,
)

coordinate_to_tuple("B2")        # (2, 2)   -> (row, column)
coordinate_to_tuple("AB17")      # (17, 28)
coordinate_from_string("B2")     # ('B', '2')  -> (litera, tekst wiersza)
range_boundaries("A1:C3")        # (1, 1, 3, 3) -> (min_col, min_row, max_col, max_row)
quote_sheetname("Dane 2026")     # "'Dane 2026'"  -> nazwa w apostrofach
```

Uwaga na trzecią funkcję: `range_boundaries` zwraca krotkę, w której **kolejność jest inna niż w przypadku współrzędnych komórki**. Dla komórki dostajesz `(row, column)`, dla zakresu `(min_col, min_row, max_col, max_row)`. To nie jest spójne i nie ma dobrego wytłumaczenia poza historią biblioteki. Zapamiętaj to jako wyjątek i sprawdź dwukrotnie, zanim użyjesz wyniku w obliczeniach.

**Konsekwencja praktyczna: nie ma jednej „właściwej" notacji.** Jest podział ról. A1 do czytelności, `(row, column)` do obliczeń, `get_column_letter` i `column_index_from_string` na granicy tych dwóch światów.

### 2.3. Indeksowanie od 1 — dlaczego `row=0` nie istnieje

Sprawdźmy to wprost. `Worksheet.cell()` waliduje argumenty i rzuca wyjątek, gdy któryś jest mniejszy od 1:

```python
ws.cell(row=0, column=1)
# ValueError: Row or column values must be at least 1

ws.cell(row=1, column=0)
# ValueError: Row or column values must be at least 1
```

Ten sam wyjątek dostaniesz przy odwołaniu tekstowym do wiersza 0:

```python
ws["A0"]
# openpyxl.utils.exceptions.CellCoordinatesException: Invalid cell coordinate A0
```

To nie jest niedbalstwo autorów openpyxl. To jest poprawne odwzorowanie rzeczywistości arkusza: **arkusz nie ma wiersza zero.** Adres `A0` nie istnieje w Excelu, więc nie istnieje w modelu.

**Analogia numeracji pięter** działa tu w pełni. W windzie w Polsce: parter (0), pierwsze piętro (1), drugie piętro (2). W windzie w USA: pierwsze piętro (1), drugie piętro (2). Jeśli masz w głowie nawyk list Pythona i myślisz „komórka 0 to pierwsza komórka", znasz tylko jeden z tych systemów.

Ale tu uwaga: **w tym samym modelu openpyxl używa także indeksów od 0.** Nie w komórkach — w stylach. Zbierzmy to w jawną tabelę, żeby nie było wątpliwości:

| Co | Liczy od | Przykład |
|---|---|---|
| `cell.row`, `cell.column` | **1** | `ws["A1"].row == 1` |
| Klucz w `ws._cells` | **1** | `ws["B3"]` → klucz `(3, 2)` |
| `ws.max_row`, `ws.max_column` | **1** | pusty arkusz: `1` |
| `get_column_letter` / `column_index_from_string` | **1** | `get_column_letter(1) == "A"` |
| Indeks w `cell._style` (`fontId`, `fillId`, ...) | **0** | `fontId=0` to pierwsza czcionka |
| `wb.worksheets[0]` (lista w Pythonie) | **0** | pierwszy arkusz |
| `wb.sheetnames[0]` (lista w Pythonie) | **0** | pierwsza nazwa |

Zauważ, że dwa ostatnie wiersze też liczą od 0 — bo to zwykłe listy Pythona. To znaczy, że **w jednym wyrażeniu możesz mieć oba systemy**:

```python
# Pierwszy arkusz (indeks 0 w liście!), jego pierwsza komórka (wiersz 1, kolumna 1).
ws = wb.worksheets[0]
komorka = ws.cell(row=1, column=1)
print(komorka.coordinate)     # 'A1'
```

To wyrażenie jest poprawne i oba systemy są użyte zgodnie z ich regułami. Ale jest to dokładnie ten rodzaj kodu, w którym łatwo o pomyłkę — i dlatego w module 07 napiszemy własne funkcje pomocnicze, które ukryją tę dwoistość.

### 2.4. Cykl życia modelu — cztery fazy

Narysujmy drogę od pliku, przez model, do nowego pliku.

```text
   PLIK NA DYSKU                                          PLIK NA DYSKU
   (bajty .xlsx)                                          (nowe bajty .xlsx)
        │                                                        ▲
        │  FAZA 1: load_workbook()                              │
        │  "rozbiór pakietu i zbudowanie modelu"                │
        ▼                                                        │
   ┌─────────────────────────────┐              FAZA 4:         │
   │  MODEL W PAMIĘCI RAM        │              save()          │
   │  (obiekty Pythona)          │              "serializacja   │
   │                             │               i zapakowanie" │
   │  Workbook                   │──────────────────────────────┘
   │   ├── Worksheet             │
   │   │    ├── Cell ──► indeksy │
   │   │    └── ...              │
   │   └── rejestry stylów       │
   └─────────────────────────────┘
              ▲          │
              │          │  FAZA 2 i 3: praca na modelu
              │          │  "modyfikacje obiektami Pythona"
              │          ▼
        (nic tu nie ma — plik
         na dysku jest NIETKNIĘTY
         przez cały czas pracy)
```

**Faza 1 — Wczytanie.**

```python
wb = load_workbook("raport.xlsx")
```

Openpyxl otwiera archiwum ZIP, czyta `xl/workbook.xml` (lista arkuszy), czyta każdy `xl/worksheets/sheetN.xml` (komórki), czyta `xl/styles.xml` (rejestry), czyta wszystko pozostałe, co modeluje — i buduje **kompletny graf obiektów w RAM**. Po tym operacja jest zakończona: **uchwyt do pliku jest zamknięty** (w trybie normalnym), a model jest samodzielny.

To ważniejsze, niż wygląda. Model nie trzyma otwartego pliku. Możesz ten plik w międzyczasie **usunąć** — model będzie dalej działał, bo zawiera wszystkie potrzebne dane. Udowodnimy to w przykładzie 4.

Alternatywa dla fazy 1: `wb = Workbook()` tworzy model **od zera** — pusty, z jednym domyślnym arkuszem o nazwie `Sheet`. Żadnego pliku nie ma, żadnego pliku nie potrzeba.

**Faza 2 — Modyfikacja.**

```python
ws = wb["Dane"]
ws["A1"] = "Nowa wartość"
ws["A2"].font = Font(bold=True)
```

Modyfikujesz obiekty Pythona. To jest zwykła praca na strukturach danych. Zmieniasz wartości, przypisujesz style, dodajesz arkusze, wstawiasz tabele. **Plik na dysku nie wie o niczym.**

W tym miejscu, w fazie 2, popełniane są wszystkie błędy tego modułu:

- `cell.font.bold = True` — mutujesz kserokopię, nie pieczątkę (patrz sekcja 1).
- `ws["A1"] = 5` a potem zapomniany `save()` — praca idzie do kosza.
- Dwa różne skrypty modyfikują ten sam plik w tym samym czasie — każdy ma własny, niezależny model, oba zapiszą, jeden nadpisze drugiego.

**Faza 3 — (dla kompletności) brak fazy 3.** Serio. Między fazą 2 a 4 **nie ma żadnej operacji na pliku**. Jeśli w Twoim kodzie jest `wb.save()` w środku pętli, to znaczy, że masz cztery fazy: dwie modyfikacje i jeden zapis, potem znowu modyfikacje i znowu zapis. Każdy `save()` przepisuje cały pakiet od nowa. O oszczędzaniu tego zajmiemy się w modułach 19 i 26.

**Faza 4 — Zapis.**

```python
wb.save("raport_wynikowy.xlsx")
```

Openpyxl **przechodzi przez cały model** i generuje nowy pakiet od zera: `[Content_Types].xml`, `_rels/`, `xl/workbook.xml`, `xl/worksheets/sheetN.xml`, `xl/styles.xml`, `xl/theme/theme1.xml` i wszystko, co model potrafi wygenerować. Następnie zapakowuje to do archiwum ZIP i zapisuje na dysk.

**Trzy konsekwencje, które musisz mieć w głowie:**

*Konsekwencja A — zapis jest całościowy.* Nie ma „dopisania" do pliku. Jest przepisanie wszystkiego. Dlatego zapis pliku o rozmiarze 50 MB trwa tyle, ile trwa, i wymaga pamięci proporcjonalnej do rozmiaru modelu (z wyjątkiem `write_only`, gdzie proporcjonalnej do jednego wiersza).

*Konsekwencja B — zapis nadpisuje bez ostrzeżenia.* `wb.save("istniejacy_plik.xlsx")` zastąpi ten plik. Bez pytania, bez kopii zapasowej, bez „czy na pewno". To jest zachowanie identyczne z „Zapisz" w edytorze tekstu — nie z „Zapisz jako". W module 00 przyjęliśmy zasadę kursu: **nigdy nie nadpisujemy pliku źródłowego.** Teraz wiesz, dlaczego ta zasada jest ważna: bo nadpisanie jest nieodwracalne i ciche.

*Konsekwencja C — zapis przepisuje tylko to, co model zna.* To jest dokładnie analogia fotokopii. Slicery, sparkline'y, kształty, Power Query, formanty ActiveX — nie ma ich w modelu, więc nie ma ich w wyniku. I **jedynym sygnałem ostrzegawczym** jest komunikat na `stderr`, którym zajmowaliśmy się w module 02.

### 2.5. „Obiekt żywy" — dwa światy, jedno słowo

Mówimy potocznie „openpyxl zapisuje plik". Ale w kodzie masz dwie różne rzeczy o tej samej nazwie:

- `wb` — **obiekt Pythona** w pamięci RAM. Kompletny graf. Zmienny, zależny od Twojego kodu, ginący razem z zakończeniem procesu.
- `raport.xlsx` — **ciąg bajtów** na dysku. Niezmienny, dopóki ktoś go nie nadpisze, trwały niezależnie od tego, czy Twój program działa.

**Między nimi nie ma żywego połączenia.** Nikt nie „subskrybuje" zmian. Zmiana w jednym nie propaguje się do drugiego. To nie jest wada techniczna — to jest fundamentalna właściwość modelu, który jest samodzielny.

Sprawdźmy to na trzech scenariuszach, które łatwo pomylić:

**Scenariusz 1: zmieniasz model, patrzysz na plik.**

```python
wb = load_workbook("raport.xlsx")
print(wb["Dane"]["A1"].value)     # 'Ala'
wb["Dane"]["A1"] = "Ola"          # zmiana MODELU
print("W modelu:", wb["Dane"]["A1"].value)   # 'Ola'

# A gdyby ktoś otworzył plik w Excelu w tym momencie?
# Zobaczyłby 'Ala'. Bo plik nie został tknięty.
```

**Scenariusz 2: zmieniasz plik, patrzysz na model.**

```python
wb = load_workbook("raport.xlsx")     # model ze stanem 'Ala'
# ... ktoś inny edytuje raport.xlsx w Excelu i zapisuje 'Ola' ...
print(wb["Dane"]["A1"].value)         # nadal 'Ala'!
```

Model to **migawka w momencie wczytania**. Nie odświeży się sam. Żadna flaga tego nie włączy. Jedynym sposobem jest ponowne `load_workbook()`.

**Scenariusz 3: dwa modele z jednego pliku.**

```python
a = load_workbook("raport.xlsx")
b = load_workbook("raport.xlsx")

a["Dane"]["A1"] = "zmiana z modelu A"
print(b["Dane"]["A1"].value)   # 'Ala' - model b nic o tym nie wie
```

To nie jest jeden model o dwóch nazwach. To **dwa niezależne grafy**. Modyfikacja jednego nie dotyka drugiego. I to jest mechanizm, który prowadzi do najczęstszego błędu współbieżności w skryptach Excel-owych: **utraconej aktualizacji**.

```python
# Skrypt 1 działa o 9:00
a = load_workbook("raport.xlsx")
a["Dane"]["A1"] = "wpis ze skryptu 1"
a.save("raport.xlsx")

# Skrypt 2 działa o 9:01, ale wczytał plik o 9:00:30 (przed zapisem)
b = load_workbook("raport.xlsx")   # wczytał stary stan!
b["Dane"]["B1"] = "wpis ze skryptu 2"
b.save("raport.xlsx")              # nadpisał - wpis skryptu 1 przepadł
```

Dokładnie taki sam mechanizm działa, gdy użytkownik ma plik otwarty w Excelu, a Ty go modyfikujesz skryptem. **Excel nic nie zablokuje** — openpyxl nie używa blokad plików. Dwa niezależne światy, które nadpisują się nawzajem na dysku.

**Konsekwencja praktyczna:** jeśli w Twoim procesie plik może być edytowany przez kogoś innego (człowieka albo inny skrypt), musisz sam zadbać o synchronizację: znacznik czasu, hash pliku wejściowego, plik blokady, katalog roboczy na zadanie. To wróci w module 18 (zapis atomowy) i module 26 (architektura zadań).

### 2.6. Trzy tryby pracy — trzy różne bestie

`Workbook` i `load_workbook` mają tryby, które zmieniają **nie tylko wydajność, ale sam charakter modelu**. To nie są „opcje optymalizacyjne" — to są trzy różne kontrakty.

**Tryb normalny — swobodny dostęp losowy.**

```python
wb = load_workbook("duzy.xlsx")            # cały plik w RAM
ws = wb["Dane"]
ws["Z999"].value = "bez problemu"
print(ws["A1"].value, ws["AB7"].value)
```

Pełny graf. Możesz sięgnąć do dowolnej komórki w dowolnej kolejności, modyfikować, iść w tył i w przód, dodawać arkusze, wykresy, tabele. Zapłata: **cały skoroszyt żyje w pamięci RAM** — typowo około 60 MB RAM na 1 MB pliku `xlsx` (współczynnik zależny od zawartości; mierzymy go w module 19).

**Tryb `read_only` — czytanie przez szczelinę.**

```python
wb = load_workbook("ogromny.xlsx", read_only=True)
ws = wb["Dane"]
for wiersz in ws.iter_rows(values_only=True):
    przetworz(wiersz)
wb.close()                                  # OBOWIĄZKOWE
```

Openpyxl nie buduje pełnego grafu. Otwiera strumień i **czyta wiersz po wierszu, w przód, tylko raz**. Pamięć jest minimalna, wydajność wysoka. Zapłata: brak dostępu losowego.

**Analogia: czytanie książki przez szczelinę w okładce.** Widzisz jedną linijkę na raz. Nie możesz wrócić do poprzedniej strony. Nie możesz policzyć stron. Nie możesz dopisać nic na marginesie.

W tym trybie:

- `ws["A1"]` — nie zadziała (brak subskrypcji).
- `ws.cell(row=1, column=1)` — nie zadziała (brak dostępu losowego).
- `ws.max_row` i `ws.max_column` — **mogą zwrócić `None`** albo wartość nieprawdziwą, bo wymiary bywają zapisane tylko jako podpowiedź w `<dimension>`.
- `ws.iter_rows()` — działa i to jest Twój jedyny sposób czytania.
- Zmiana czegokolwiek — nie ma sensu, i nie zadziała.
- `wb.close()` — **obowiązkowe**, inaczej zostawiasz otwarty strumień i wycieka pamięć.

**Tryb `write_only` — prasa drukarska.**

```python
wb = Workbook(write_only=True)
ws = wb.create_sheet("Log")
for rekord in dane:
    ws.append(rekord)
wb.save("wynik.xlsx")
```

To odwrotność `read_only`. Openpyxl **wypuszcza wiersze na wyjście i natychmiast o nich zapomina**. Nie trzyma ich w modelu. Możesz generować plik o rozmiarze, na który nie starczyłoby Ci pamięci.

**Analogia: prasa drukarska.** Papier wychodzi z maszyny arkusz po arkuszu i **nie ma mechanizmu, żeby go cofnąć**. Możesz drukować dalej, możesz zmienić tekst na następnym arkuszu — ale nie możesz wrócić do tego, który już wyszedł i czegoś dopisać.

W tym trybie:

- `wb.active` — bywa `None`, bo nowy `write_only` skoroszyt nie ma domyślnego arkusza. Sprawdź u siebie.
- `ws["A1"]`, `ws.cell(...)` — nie zadziałają. **Jedynym sposobem zapisu jest `ws.append()`**.
- `ws.max_row` — nie ma sensu; nic nie jest pamiętane.
- Style — musisz je nadać **przed** `append`, za pomocą `WriteOnlyCell`.

```python
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

komorka = WriteOnlyCell(ws, value="Ważny rekord")
komorka.font = Font(bold=True)
ws.append([komorka, "zwykły tekst", 42])       # styl ustawiony PRZED append
```

Trzecia droga: nadanie stylu po `append` nie jest możliwe, bo komórka już została wypuszczona z modelu.

**Jak wybrać tryb?** Trzy pytania:

| Pytanie | Odpowiedź | Tryb |
|---|---|---|
| Czy plik jest ogromny i tylko go czytam? | tak | `read_only=True` |
| Czy generuję ogromny plik od zera? | tak | `Workbook(write_only=True)` |
| Czy muszę sięgać do komórek w dowolnej kolejności, dodawać wykresy, czytać style? | tak | normalny |

Szczegóły pomiarów, pułapek i technik optymalizacji — moduł 19.

### 2.7. Trzy kategorie prawdy o API

W module 01 zapowiedzieliśmy podział, który jest osią całego kursu. Czas go rozwinąć. **Każda funkcja Excela, z którą się zetkniesz, wpada do jednego z trzech worków** — i wiedza o tym, do którego, jest ważniejsza niż znajomość API.

**Kategoria 1: POTRAFI.** openpyxl ma pełny model: potrafi to odczytać, potrafi to utworzyć, potrafi to zapisać. Wartości są przechowywane dokładnie takie, jakie wpisałeś.

**Kategoria 2: POTRAFI CZĘŚCIOWO.** Istnieje model i API, ale nie oddaje wszystkich niuansów. Efekt: coś przechodzi, ale coś innego zostaje uproszczone. Typowe objawy: „uproszczone formatowanie wykresu po otwarciu w Excelu", „reguła istnieje, ale jeden jej wariant ginie", „zapisane, ale wygląda trochę inaczej".

**Kategoria 3: GUBI PRZY ZAPISIE.** Modelu nie ma. Odczyt tego nie robi, ale zapis — przez przepisanie modelu — nie ma z czego tego odtworzyć. Objaw: element milczy na wejściu i nie istnieje na wyjściu. Jedynym sygnałem bywa komunikat na `stderr`.

Oto tabela startowa. Wracaj do niej, gdy będziesz się zastanawiać, czy coś „zadziała".

| Element | Kategoria | Moduł ze szczegółami |
|---|---|---|
| Wartości: tekst, liczba, `bool`, `None` | potrafi | 05 |
| `datetime`, `date`, `time`, `timedelta` | potrafi (bez stref czasowych) | 05 |
| `Decimal` | potrafi (rzutowany na `float`) | 05 |
| Formuła — zapis literału | potrafi | 06 |
| Formuła — **obliczenie wartości** | **nie potrafi w ogóle** | 06, 26 |
| Formuła — przesunięcie referencji | potrafi (`Translator`) | 06, 07 |
| Style: `Font`, `PatternFill`, `Border`, `Alignment`, `Protection` | potrafi | 09 |
| `number_format` | potrafi | 10 |
| Style nazwane (`NamedStyle`) | potrafi | 09 |
| Wymiary, scalenia, `freeze_panes` | potrafi | 11 |
| Autofiltr, ustawienia wydruku, nagłówki/stopki | potrafi | 11 |
| Tabele (`Table`, `ListObject`) | potrafi | 12 |
| Formatowanie warunkowe — podstawowe typy Reguł | potrafi | 13 |
| Formatowanie warunkowe — warianty zapisane w `x14:` | **potrafi częściowo** | 13, 14 |
| Wykresy — tworzenie od zera | **potrafi częściowo** (nie wszystkie typy) | 15 |
| Wykresy — odczyt i zapis z powrotem | **potrafi częściowo** (odbudowa z modelu) | 15, 18 |
| Obrazy, hiperlinki, komentarze | potrafi | 16 |
| Walidacja danych — podstawowe typy | potrafi | 17 |
| Walidacja danych — warianty `x14:` | **potrafi częściowo** | 17 |
| Ochrona arkusza i skoroszytu | potrafi (słaba, ale kompletna) | 17 |
| Tabele przestawne — tworzenie od zera | **nie potrafi** | 18 |
| Tabele przestawne — odczyt i zapis z powrotem | **potrafi częściowo** | 18 |
| Makra VBA | **potrafi częściowo** (tylko `keep_vba=True`, bez edycji) | 18 |
| Formanty ActiveX, `customUI` | **gubi** (chyba że `keep_vba=True`) | 18 |
| Sparkline'y | **gubi** | 18 |
| Slicery, osie czasu | **gubi** | 18 |
| Kształty, pola tekstowe, formanty osadzone | **gubi** | 18 |
| Power Query, Power Pivot, model danych | **gubi** | 18 |
| Pliki `.xls` (BIFF), `.xlsb` | **nie obsługuje** (nie jest to `.xlsx`) | 01 |
| Szyfrowanie pliku hasłem (AES) | **nie obsługuje** | 21 |

**Reguła praktyczna:** zanim zaplanujesz automatyzację, znajdź swoje funkcje w tej tabeli. Jeśli któraś jest w kategorii 2 albo 3 — zaplanuj obejście od początku, a nie w dniu wdrożenia.

### 2.8. Dlaczego style są współdzielone i niemutowalne — pełne wyjaśnienie

To najważniejsza sekcja tego modułu, bo dotyczy 50% frustracji początkujących.

Zbierzmy trzy fakty, które już znasz:

1. Komórka trzyma **indeks** stylu, nie styl. (Sekcja 2.1, potwierdzona w module 02 przez `s="2"`.)
2. Rejestry stylów są **wspólne dla całego skoroszytu**. (Sekcja 2.1.)
3. W modułach 02 i 03 wielokrotnie powtórzyliśmy, że `cell.font.bold = True` nie działa.

Połączmy je w jeden łańcuch przyczynowy:

```text
Gdyby style były mutowalne przez komórkę:

  ws["A1"].font.bold = True
       │
       └─► komórka A1 zmienia swój styl
                │
                └─► ale komórka NIE TRZYMA stylu, tylko indeks!
                          │
                          └─► mutacja musiałaby trafić do rejestru
                                    │
                                    └─► a w rejestrze pod tym indeksem
                                        jest czcionka współdzielona
                                        przez 10 000 komórek
                                              │
                                              └─► ZMIENIASZ WYGLĄD
                                                  10 000 KOMÓREK,
                                                  choć zmieniałeś JEDNĄ.
```

To jest katastrofa, której nikt by nie zniósł. Dlatego openpyxl rozwiązuje problem **przez konwencję**, nie przez błąd:

> Getter `cell.font` **nie zwraca** obiektu z rejestru. Zwraca obiekt, który jest tylko czytelniczym przedstawicielem tego wpisu. Możesz go zmodyfikować w Pythonie (bo Python pozwala), ale nic go nie odczyta z powrotem. Zmiana ginie w momencie, gdy zmienna wychodzi z zakresu.

To dlatego nie ma wyjątku. To dlatego jest **cichy brak efektu**. I dlatego bez zrozumienia tego mechanizmu będziesz godzinami szukać błędu w kodzie, który jest „poprawny".

**Dwa poprawne sposoby:**

```python
# Sposób A: nowy styl w całości
ws["A1"].font = Font(name="Arial", size=14, bold=True, italic=True)

# Sposób B: kserokopia + modyfikacja + przypisanie
from copy import copy
nowa = copy(ws["A1"].font)     # kserokopia etykiety
nowa.bold = True               # dopisujemy na kopii
ws["A1"].font = nowa           # przybijamy nową etykietę
```

**Kiedy który?** Sposób A, gdy wiesz dokładnie, jak komórka ma wyglądać. Sposób B, gdy chcesz **zmienić jedną cechę**, zachowując pozostałe — na przykład pogrubić komórkę, która już ma swój rozmiar, czcionkę i kolor.

**Subtelna pułapka sposobu B:** `copy()` z modułu `copy` jest **płytkie**. Kopiuje tylko jeden poziom. Jeśli `Font` zawiera zagnieżdżony obiekt `Color` (a zawiera), to `copy(font).color` wskazuje **na ten sam obiekt `Color`** co oryginał. Modyfikacja `nowa.color.rgb = "FFFF0000"` może zatem dotknąć oryginału. Bezpieczniej:

```python
# Zagnieżdżone obiekty kopiuj osobno...
nowa = copy(ws["A1"].font)
nowa.color = copy(ws["A1"].font.color)
nowa.color.rgb = "FFFF0000"
ws["A1"].font = nowa

# ...albo po prostu zbuduj nowy obiekt, podając wartości wprost.
ws["A1"].font = Font(
    name=ws["A1"].font.name,
    size=ws["A1"].font.size,
    bold=True,                            # <- ta jedna się zmienia
    color="FFFF0000",
)
```

Drugie podejście jest dłuższe, ale jawne i odporne na wszelkie niespodzianki. W module 09 napiszemy na to porządną funkcję pomocniczą.

**Dowód, że style są współdzielone na poziomie PLIKU.** Nie musisz mi wierzyć na słowo ani zaglądać do prywatnych atrybutów. Wystarczy zapisać plik i policzyć definicje czcionek w `xl/styles.xml`. Jeśli po przypisaniu pogrubionej czcionki do 1000 komórek w pliku jest **jedna** definicja tej czcionki — dowód jest nie do obalenia. Robimy to w przykładzie 3.

### 2.9. Trzy pytania, które zadaj sobie przed każdą operacją

Wyprowadźmy z tego modułu zestaw pytań kontrolnych. Zadawaj je sobie zawsze, gdy coś „nie działa" albo gdy planujesz nową operację:

**Pytanie 1: Czy to jest zmiana w MODELU, czy w PLIKU?**

Jeśli operacja dotyczy `wb`, `ws` albo `cell` — to jest model. Jeśli dotyczy nazwy pliku, archiwum ZIP albo `xl/styles.xml` — to jest plik. **Nie przenikają się.** Zmiana w modelu bez `save()` nie dotyka pliku. Zmiana w pliku bez `load_workbook()` nie dotyka modelu.

**Pytanie 2: Czy element, którego potrzebuję, jest w modelu openpyxl?**

Jeśli tak — działa. Jeśli jest zapisany w wariancie rozszerzonym (`x14:`) — zadziała częściowo. Jeśli nie ma go w modelu (sparkline'y, slicery, kształty) — zgubisz go przy zapisie. Odpowiedź znajdziesz w tabeli z sekcji 2.7.

**Pytanie 3: Czy w tym wyrażeniu nie mieszam dwóch systemów indeksowania?**

`ws.cell(row=1, column=1)` — oba od 1. `wb.worksheets[0]` — od 0, ale to lista Pythona, więc w porządku. `cell._style.fontId` — od 0, ale to rejestr. Sprawdź każdą liczbę osobno.

Te trzy pytania, zadane konsekwentnie, zastąpią Ci połowę debugowania w tym kursie.

## 3. Przykłady krok po kroku

Wszystkie przykłady zapisują do `output/`, zgodnie z konwencją modułu 00. Każdy jest samodzielnym, uruchamialnym skryptem.

### Przykład 1 — Graf obiektów: dotknij go palcem

Zamiast czytać diagram z sekcji 2.1, sprawdźmy każdy jego element na żywym obiekcie.

```python
"""Wypisuje graf obiektów openpyxl dla świeżo utworzonego skoroszytu.

Uruchom:  python examples/03_graf.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def opis_stylu(cell) -> str:
    """Wyciąga INDEKSY stylu z komórki - nie obiekty.

    UWAGA: `_style` to API PRYWATNE. W openpyxl 3.1.x jest to StyleArray,
    czyli tablica liczb (indeksów do rejestrów skoroszytu). Nazwy pól
    czytamy przez getattr, bo mogą się różnić między wersjami.
    Jeśli żadnego pola nie ma - pokażemy surową tablicę.
    """
    s = getattr(cell, "_style", None)
    if s is None:
        return "(komórka nie ma tablicy stylu)"

    nazwy = ("fontId", "fillId", "borderId", "numFmtId",
             "protectionId", "alignmentId", "xfId")
    czesci = [f"{n}={getattr(s, n)}" for n in nazwy if hasattr(s, n)]
    if czesci:
        return ", ".join(czesci)

    try:
        return f"surowa tablica: {list(s)}"
    except TypeError:
        return f"nieznana struktura: {type(s).__name__}"


def main() -> None:
    print("=" * 72)
    print("POZIOM 1: SKOROSZYT (Workbook)")
    print("=" * 72)

    wb = Workbook()
    print(f"  typ:              {type(wb).__module__}.{type(wb).__name__}")
    print(f"  nazwy arkuszy:    {wb.sheetnames}")
    print(f"  obiekty arkuszy:  {wb.worksheets}")
    print(f"  liczba arkuszy:   {len(wb.worksheets)}")
    print(f"  aktywny arkusz:   {wb.active}")
    try:
        import openpyxl
        print(f"  wersja openpyxl:  {openpyxl.__version__}")
    except Exception:
        pass

    print()
    print("=" * 72)
    print("POZIOM 2: ARKUSZ (Worksheet)")
    print("=" * 72)

    ws = wb.active
    ws.title = "Dane"

    print(f"  typ:              {type(ws).__module__}.{type(ws).__name__}")
    print(f"  nazwa (title):    {ws.title!r}")
    print(f"  rodzic (parent):  {type(ws.parent).__name__}")

    # To jest dowód, że arkusz zna swój skoroszyt:
    print(f"  czy parent to wb? {ws.parent is wb}")

    print(f"  max_row:          {ws.max_row}")
    print(f"  max_column:       {ws.max_column}")
    print(f"  wymiary:          {ws.dimensions!r}")

    print()
    print("=" * 72)
    print("POZIOM 3: KOMÓRKA (Cell)")
    print("=" * 72)

    ws["A1"] = "Ala"
    ws["B3"] = 42
    ws["C1"] = 3.14

    for adres in ("A1", "B3", "C1"):
        c = ws[adres]
        print(f"\n  --- komórka {adres} ---")
        print(f"    typ:              {type(c).__module__}.{type(c).__name__}")
        print(f"    value:            {c.value!r}")
        print(f"    data_type:        {c.data_type!r}   "
              f"(s=tekst, n=liczba, b=bool, d=data, f=formuła, e=błąd)")
        print(f"    row:              {c.row}")
        print(f"    column:           {c.column}")
        print(f"    coordinate:       {c.coordinate!r}")
        print(f"    column_letter:    {c.column_letter!r}")
        print(f"    parent:           {type(c.parent).__name__}")
        print(f"    _style (indeksy): {opis_stylu(c)}")

    print()
    print("=" * 72)
    print("ŚCIEŻKA POWROTNA: komórka -> arkusz -> skoroszyt")
    print("=" * 72)

    c = ws["A1"]
    print(f"  c.parent is ws:            {c.parent is ws}")
    print(f"  c.parent.parent is wb:     {c.parent.parent is wb}")

    print()
    print("=" * 72)
    print("WEWNĘTRZNY SŁOWNIK KOMÓREK ARKUSZA")
    print("=" * 72)

    # To najjawniejszy dowód na 1-indeksowanie i na kolejność (row, column).
    print("  ws._cells = {")
    for klucz, komorka in sorted(ws._cells.items()):
        print(f"      {klucz!r}: {komorka!r},")
    print("  }")
    print()
    print("  Zwróć uwagę: klucz (3, 2) to wiersz 3, kolumna 2, czyli komórka B3.")
    print("  Kolejność w krotce to (row, column) - nie (column, row).")
    print("  Obie liczby są >= 1. Wiersz 0 i kolumna 0 nie istnieją.")

    print()
    print("=" * 72)
    print("TYTUŁ ARKUSZA ZMIENIA SIĘ W MODELU - I TYLKO W MODELU")
    print("=" * 72)

    ws.title = "Dane_Tymczasowe"
    print(f"  ws.title po zmianie:       {ws.title!r}")
    print(f"  wb.sheetnames:             {wb.sheetnames}")
    print("  Plik na dysku:             NIE ISTNIEJE (nie było save())")

    sciezka = OUTPUT / "03_graf.xlsx"
    wb.save(sciezka)
    print(f"\n  Po save(): {sciezka}  ({sciezka.stat().st_size} B)")

    # Kontrola: wczytaj z dysku i sprawdź, czy zmiana dotarła.
    from openpyxl import load_workbook
    kontrola = load_workbook(sciezka)
    print(f"  Wczytane nazwy arkuszy:    {kontrola.sheetnames}")

    print()
    print("=" * 72)
    print("PODSUMOWANIE GRAFU")
    print("=" * 72)
    print("""
  Workbook
    ├── worksheets[0]  -> Worksheet('Dane_Tymczasowe')
    │     ├── _cells[(1,1)] -> Cell('A1')  value='Ala'
    │     ├── _cells[(1,3)] -> Cell('C1')  value=3.14
    │     ├── _cells[(3,2)] -> Cell('B3')  value=42
    │     └── parent -> Workbook (z powrotem)
    └── rejestry stylów (wspólne dla wszystkich arkuszy)
          └── fonts / fills / borders / number_formats / ...
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci:** tworzymy pełny graf — jeden `Workbook`, jeden `Worksheet`, trzy `Cell` w słowniku `_cells`, krotka rejestrów stylów (pusta, bo nie ustawiliśmy żadnego stylu). Każdy węzeł zna swojego rodzica.

**Co trafi do pliku:** dopiero po `save()` powstaje archiwum ZIP z modelem. Przed `save()` w `output/` nie ma nic nowego.

**Oczekiwany fragment wyniku:**

```text
  --- komórka B3 ---
    value:            42
    data_type:        'n'
    row:              3
    column:           2
    coordinate:       'B3'
    column_letter:    'B'
    _style (indeksy): fontId=0, fillId=0, borderId=0, numFmtId=0

WEWNĘTRZNY SŁOWNIK KOMÓREK ARKUSZA
  ws._cells = {
      (1, 1): <Cell 'Dane'.A1>,
      (1, 3): <Cell 'Dane'.C1>,
      (3, 2): <Cell 'Dane'.B3>,
  }
```

To jest ten moment, w którym graf przestaje być abstrakcją. `(3, 2)` to `B3`. Obie liczby od 1. Wiersz pierwszy w krotce. A `fontId=0` — **to pierwsza czcionka**, indeks zero.

### Przykład 2 — Komórka ma indeksy, nie style

Bezpośrednia obserwacja najważniejszego faktu tego modułu.

```python
"""Pokazuje, że komórka trzyma INDEKSY stylu, nie obiekty stylu.

Uruchom:  python examples/03_indeksy_stylu.py
"""

from copy import copy
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def indeksy(cell) -> str:
    s = getattr(cell, "_style", None)
    if s is None:
        return "(brak)"
    nazwy = ("fontId", "fillId", "borderId", "numFmtId",
             "protectionId", "alignmentId", "xfId")
    czesci = [f"{n}={getattr(s, n)}" for n in nazwy if hasattr(s, n)]
    if czesci:
        return ", ".join(czesci)
    try:
        return f"surowa tablica: {list(s)}"
    except TypeError:
        return f"nieznana struktura: {type(s).__name__}"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"

    print("=" * 72)
    print("CZĘŚĆ 1: INDEKSY STYLU ZMIENIAJĄ SIĘ PRZY PRZYPISANIU")
    print("=" * 72)

    ws["A1"] = 42
    print(f"  Świeża komórka A1:            {indeksy(ws['A1'])}")

    ws["A1"].font = Font(bold=True, color="FFFF0000")
    print(f"  Po .font = Font(...):         {indeksy(ws['A1'])}")

    ws["A1"].fill = PatternFill("solid", fgColor="FFFFFF00")
    print(f"  Po .fill = PatternFill(...):  {indeksy(ws['A1'])}")

    ws["A1"].border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    print(f"  Po .border = Border(...):     {indeksy(ws['A1'])}")

    ws["A1"].number_format = '#,##0.00 "zł"'
    print(f"  Po .number_format = ...:      {indeksy(ws['A1'])}")

    ws["A1"].alignment = Alignment(horizontal="center")
    print(f"  Po .alignment = ...:          {indeksy(ws['A1'])}")

    print()
    print("  WNIOSEK: nie ma tu ani jednego obiektu Font, Fill ani Border.")
    print("  Są tylko liczby - indeksy do rejestrów skoroszytu.")
    print("  To dokładnie to samo, co s=\"N\" w xl/styles.xml z modułu 02.")

    print()
    print("=" * 72)
    print("CZĘŚĆ 2: DLACZEGO cell.font.bold = True NIE DZIAŁA")
    print("=" * 72)

    ws["B2"] = 100
    ws["B2"].font = Font(name="Calibri", size=11)

    przed = indeksy(ws["B2"])
    font_widziany = ws["B2"].font
    print(f"  Przed:              {przed}")
    print(f"  font_widziany:      {font_widziany!r}")
    print(f"  font_widziany.bold: {font_widziany.bold}")

    # TA LINIA NIC NIE ROBI:
    font_widziany.bold = True
    print(f"  Po 'font_widziany.bold = True':")
    print(f"    indeksy B2:       {indeksy(ws['B2'])}   <- BEZ ZMIAN")
    print(f"    .bold widziany:   {ws['B2'].font.bold}   <- NADAL False!")
    print()
    print("  Zmodyfikowaliśmy obiekt, który nie jest tą czcionką, którą")
    print("  trzyma rejestr. To była kserokopia etykiety. Zmiana w niej")
    print("  nie dotarła do szuflady.")

    print()
    print("=" * 72)
    print("CZĘŚĆ 3: TRZY POPRAWNE SPOSOBY ZMIANY JEDNEJ CECHY")
    print("=" * 72)

    # --- Sposób 1: copy + modyfikacja + przypisanie (płytkie copy)
    ws["C3"] = 10
    ws["C3"].font = Font(name="Calibri", size=11, italic=True)
    print(f"  C3 przed:                 {indeksy(ws['C3'])}")

    nowa = copy(ws["C3"].font)          # kserokopia etykiety
    nowa.bold = True                    # dopisujemy na kopii
    ws["C3"].font = nowa                # przybijamy kopię jako nową etykietę

    print(f"  C3 po copy()+bold:        {indeksy(ws['C3'])}")
    print(f"  C3.font.name:             {ws['C3'].font.name!r}   <- zachowane")
    print(f"  C3.font.size:             {ws['C3'].font.size}    <- zachowane")
    print(f"  C3.font.italic:           {ws['C3'].font.italic}  <- zachowane")
    print(f"  C3.font.bold:             {ws['C3'].font.bold}    <- ZMIENIONE")

    # --- Sposób 2: zbudowanie nowego obiektu z jawnymi wartościami
    ws["D4"] = 20
    ws["D4"].font = Font(name="Arial", size=9)
    stary = ws["D4"].font
    ws["D4"].font = Font(
        name=stary.name,
        size=stary.size,
        bold=True,                      # <- ta jedna cecha
    )
    print(f"\n  D4 po Font(name=stary.name, ...): "
          f"{ws['D4'].font.name!r}, {ws['D4'].font.size}, bold={ws['D4'].font.bold}")

    # --- Sposób 3: przypisanie całkiem nowego stylu
    ws["E5"] = 30
    ws["E5"].font = Font(bold=True, color="FF0066CC", name="Arial", size=12)
    print(f"  E5 po pełnym Font(...):   bold={ws['E5'].font.bold}, "
          f"color={ws['E5'].font.color.rgb if ws['E5'].font.color else None}")

    sciezka = OUTPUT / "03_indeksy_stylu.xlsx"
    wb.save(sciezka)
    print(f"\n  Zapisano: {sciezka}")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci:** każde przypisanie `ws["A1"].font = Font(...)` powoduje, że openpyxl dodaje obiekt do rejestru czcionek skoroszytu (albo znajduje istniejący identyczny) i wpisuje jego indeks do tablicy `_style` komórki. Mutacja `font_widziany.bold = True` zmienia tymczasowy obiekt, którego nikt nie odczyta — bo `ws["B2"].font` przy kolejnym wywołaniu zbuduje **nowy** przedstawiciela z rejestru, w którym nic się nie zmieniło.

**Co trafi do pliku:** w `xl/styles.xml` pojawią się definicje czcionek dla A1, C3, D4 i E5 — i **żadna** dla „pogrubionej Calibri 11" z B2, bo ta czcionka nigdy nie powstała.

### Przykład 3 — Dowód, że style są współdzielone

Ten przykład jest najważniejszym dowodem tego modułu. Pokazuje, że obietnica z sekcji 2.8 jest prawdziwa **na poziomie pliku**, nie tylko w dokumentacji.

```python
"""Przypisuje tę samą czcionkę do 1000 komórek, a potem liczy definicje
czcionek w xl/styles.xml. Jeśli jest tam jedna definicja - style są współdzielone.

Uruchom:  python examples/03_wspoldzielenie_stylow.py
"""

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def policz_wpisy(sciezka: Path) -> dict[str, int]:
    """Liczy wpisy w rejestrach wewnątrz xl/styles.xml (jak w module 02)."""
    with zipfile.ZipFile(sciezka) as archiwum:
        if "xl/styles.xml" not in archiwum.namelist():
            return {}
        root = ET.fromstring(archiwum.read("xl/styles.xml"))

    wynik = {}
    for nazwa in ("fonts", "fills", "borders", "cellXfs", "dxfs"):
        wezel = root.find(f"{NS}{nazwa}")
        wynik[nazwa] = len(list(wezel)) if wezel is not None else 0
    return wynik


def policz_zawartosc_komorek(sciezka: Path) -> int:
    """Ile elementów <c> jest w pierwszym arkuszu."""
    with zipfile.ZipFile(sciezka) as archiwum:
        nazwa = next(
            n for n in archiwum.namelist()
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        )
        root = ET.fromstring(archiwum.read(nazwa))
    return sum(1 for el in root.iter() if el.tag == f"{NS}c")


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Test"

    sciezka_pusto = OUTPUT / "03_style_0.xlsx"
    wb.save(sciezka_pusto)
    print("=" * 72)
    print("STAN POCZĄTKOWY (nic nie ustawiono)")
    print("=" * 72)
    for k, v in policz_wpisy(sciezka_pusto).items():
        print(f"  {k:<12} {v}")

    print()
    print("=" * 72)
    print("PRZYPISUJEMY TĘ SAMĄ CZCIONKĘ DO 1000 KOMÓREK")
    print("=" * 72)

    # Jeden obiekt Font... ale UWAGA: nawet gdybyśmy tworzyli nowy obiekt
    # w każdej iteracji, openpyxl zdeduplikowałby go w rejestrze.
    pogrubiona = Font(bold=True, color="FFFF0000")

    for i in range(1, 1001):
        ws.cell(row=i, column=1, value=f"wiersz {i}").font = pogrubiona

    sciezka_jeden = OUTPUT / "03_style_jeden.xlsx"
    wb.save(sciezka_jeden)

    print(f"  Komórek w arkuszu:            {policz_zawartosc_komorek(sciezka_jeden)}")
    print(f"  Ustawionych czcionek:         1000")
    print(f"  Definicji <font> w pliku:     "
          f"{policz_wpisy(sciezka_jeden).get('fonts', '?')}")

    print()
    print("=" * 72)
    print("PRZYPISUJEMY 10 RÓŻNYCH CZCIONEK DO KOLEJNYCH 10 000 KOMÓREK")
    print("=" * 72)

    wb2 = Workbook()
    ws2 = wb2.active
    czcionki = [
        Font(bold=True, color="FFFF0000"),
        Font(bold=False, color="FF0000FF"),
        Font(italic=True, size=8),
        Font(italic=True, size=24),
        Font(name="Arial"),
        Font(name="Times New Roman"),
        Font(underline="single"),
        Font(strike=True),
        Font(vertAlign="superscript"),
        Font(size=7, color="FF00AA00"),
    ]
    for i in range(1, 10001):
        ws2.cell(row=i, column=1, value=i).font = czcionki[i % 10]

    sciezka_dziesiec = OUTPUT / "03_style_dziesiec.xlsx"
    wb2.save(sciezka_dziesiec)

    print(f"  Komórek w arkuszu:            {policz_zawartosc_komorek(sciezka_dziesiec)}")
    print(f"  Różnych czcionek użytych:     10")
    print(f"  Definicji <font> w pliku:     "
          f"{policz_wpisy(sciezka_dziesiec).get('fonts', '?')}")
    print(f"  Definicji <cellXfs>:          "
          f"{policz_wpisy(sciezka_dziesiec).get('cellXfs', '?')}")
    print(f"  Rozmiar pliku:                "
          f"{sciezka_dziesiec.stat().st_size / 1024:.1f} kB")

    print()
    print("=" * 72)
    print("WNIOSEK")
    print("=" * 72)
    print("""
  1000 komórek z tą samą czcionką -> JEDNA definicja <font> w pliku.
  10000 komórek z 10 czcionkami   -> DZIESIĘĆ definicji <font>.

  To jest bezpośredni dowód, że style są indeksami do wspólnego rejestru.
  Komórka nie nosi stylu - nosi NUMER stylu.

  I dlatego cell.font.bold = True nie działa: mutowanie obiektu, który
  zwraca getter, nie dotyka wpisu w rejestrze. Musiałoby dotknąć, i to
  właśnie dlatego openpyxl na to nie pozwala - zmieniłoby wygląd
  wszystkich komórek wskazujących ten sam indeks.
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci:** pętla 1000 razy wykonuje `ws.cell(...).font = pogrubiona`. Każde przypisanie powoduje dodanie obiektu do rejestru czcionek, ale openpyxl rozpoznaje, że taki wpis już jest, i **zwraca ten sam indeks**. W tablicy `_style` każdej komórki ląduje ta sama liczba.

**Co trafi do pliku:** jeden `<font>` w `xl/styles.xml` niezależnie od tego, ilu komórek dotyczy. **To jest właśnie ta część pakietu, którą rozbieraliśmy w module 02.**

**Uwaga o wersjach:** jeśli w Twojej wersji openpyxl rejestry są zorganizowane inaczej (na przykład częściowo per-arkusz), liczby mogą się nieznacznie różnić — ale nieznacznie, nie o trzy rzędy wielkości. Jeśli zobaczysz 1000 definicji czcionek dla 1000 komórek, to znaczy, że deduplikacja nie działa i masz do czynienia z czymś nietypowym. Warto to zgłosić.

### Przykład 4 — Żywy obiekt: dwa światy, jeden plik

Ten przykład udowadnia trzy rzeczy naraz: że `save()` jest jedynym momentem, w którym plik powstaje; że model jest niezależny od pliku (można plik usunąć); i że dwa `load_workbook()` dają dwa niezależne światy.

```python
"""Dowód na to, że model w RAM i plik na dysku to dwa niezależne światy.

Uruchom:  python examples/03_dwa_swiaty.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

sciezka = OUTPUT / "03_zywy.xlsx"


def stan(sciezka: Path, opis: str) -> None:
    """Pokazuje, co aktualnie leży w pliku na dysku."""
    if not sciezka.exists():
        print(f"  [{opis}] plik nie istnieje")
        return
    wb = load_workbook(sciezka)
    wartosc = wb["Dane"]["A1"].value
    print(f"  [{opis}] w pliku A1 = {wartosc!r}  "
          f"(rozmiar {sciezka.stat().st_size} B)")


def main() -> None:
    print("=" * 72)
    print("KROK 1: Model powstaje w RAM. Pliku jeszcze NIE MA.")
    print("=" * 72)

    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    ws["A1"] = "Ala"

    print(f"  W MODELU:  wb['Dane']['A1'].value = {ws['A1'].value!r}")
    print(f"  wb.sheetnames = {wb.sheetnames}")
    stan(sciezka, "na dysku")

    print()
    print("=" * 72)
    print("KROK 2: save() - dopiero teraz powstaje plik.")
    print("=" * 72)

    wb.save(sciezka)
    stan(sciezka, "na dysku po save")

    print()
    print("=" * 72)
    print("KROK 3: Zmieniamy MODEL. Plik na dysku się NIE zmienia.")
    print("=" * 72)

    ws["A1"] = "Ola"
    print(f"  W MODELU:  A1 = {ws['A1'].value!r}")
    stan(sciezka, "na dysku po zmianie modelu")

    print()
    print("  ^^^ To jest najważniejsza obserwacja tego modułu.")
    print("      Zmiana modelu bez save() jest niewidoczna w pliku.")

    print()
    print("=" * 72)
    print("KROK 4: save() znowu - teraz zmiana trafia do pliku.")
    print("=" * 72)

    wb.save(sciezka)
    stan(sciezka, "na dysku po drugim save")

    print()
    print("=" * 72)
    print("KROK 5: Model jest NIEZALEŻNY od pliku. Usuwamy plik...")
    print("=" * 72)

    sciezka.unlink()
    print(f"  W MODELU:  A1 = {ws['A1'].value!r}   <- model działa dalej!")
    print(f"  wb.sheetnames = {wb.sheetnames}")
    stan(sciezka, "na dysku")

    print()
    print("=" * 72)
    print("KROK 6: ...i odtwarzamy plik z modelu.")
    print("=" * 72)

    wb.save(sciezka)
    stan(sciezka, "na dysku po odtworzeniu")
    print()
    print("  openpyxl nie trzyma pliku otwartego. Model żyje w RAM.")
    print("  Plik na dysku to tylko wynik ostatniego save().")

    print()
    print("=" * 72)
    print("KROK 7: DWA load_workbook() = DWA NIEZALEŻNE MODELE")
    print("=" * 72)

    a = load_workbook(sciezka)
    b = load_workbook(sciezka)

    print(f"  a: {id(a)}   b: {id(b)}   ten sam obiekt? {a is b}")
    print(f"  a['Dane']['A1'] = {a['Dane']['A1'].value!r}")
    print(f"  b['Dane']['A1'] = {b['Dane']['A1'].value!r}")

    a["Dane"]["A1"] = "zmiana z modelu A"
    print(f"\n  Po zmianie w modelu A:")
    print(f"    a['Dane']['A1'] = {a['Dane']['A1'].value!r}")
    print(f"    b['Dane']['A1'] = {b['Dane']['A1'].value!r}   <- b NIC NIE WIE")

    b["Dane"]["B1"] = "wpis z modelu B"
    b.save(sciezka)

    print(f"\n  Model B zapisał plik (nadpisując stan A):")
    stan(sciezka, "na dysku")

    print()
    print("=" * 72)
    print("WNIOSEK: UTRACONA AKTUALIZACJA")
    print("=" * 72)
    print(f"""
  Model A miał w A1 'zmiana z modelu A'.
  Model B nie wiedział o tej zmianie i zapisał cały plik od nowa
  ze swoim stanem. Wpis A przepadł bez ostrzeżenia.

  To jest ten sam mechanizm, który zadziała, gdy użytkownik ma plik
  otwarty w Excelu, a Ty go modyfikujesz skryptem. I odwrotnie:
  gdy zapiszesz plik, gdy użytkownik ma go otwartego, Excel przy
  zamknięciu zapyta "czy zapisać?" - a jego "tak" nadpisze Twoją pracę.
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci:** trzy modele: `wb` (główny), `a` i `b` — kompletnie niezależne grafy. `id(a) != id(b)`. Modyfikacja jednego nie dotyka drugiego. `sciezka.unlink()` usuwa plik, a `wb` dalej ma wszystkie dane — bo je ma, w RAM.

**Co trafi do pliku:** nic, dopóki nie ma `save()`. Po ostatnim `save()` plik zawiera **wyłącznie stan modelu `b`** — bo `b.save()` przepisuje cały pakiet.

### Przykład 5 — Trzy tryby pracy

```python
"""Pokazuje różnicę między trybem normalnym, read_only i write_only.

Uruchom:  python examples/03_tryby.py
"""

import time
import tracemalloc
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

LICZBA_WIERSZY = 20_000


def przygotuj_plik(sciezka: Path) -> None:
    """Plik testowy w trybie write_only - sam demonstruje swój tryb."""
    print(f"  Tworzę plik testowy ({LICZBA_WIERSZY} wierszy) w trybie write_only...")
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Dane")
    ws.append(["id", "nazwa", "kwota"])
    for i in range(1, LICZBA_WIERSZY + 1):
        ws.append([i, f"pozycja {i}", i * 1.23])
    wb.save(sciezka)
    print(f"  Gotowe: {sciezka.name} ({sciezka.stat().st_size / 1024:.0f} kB)")


def zmierz(opis: str, funkcja) -> None:
    """Uruchamia funkcję z pomiarem pamięci szczytowej i czasu."""
    tracemalloc.start()
    start = time.perf_counter()
    wynik = funkcja()
    czas = time.perf_counter() - start
    _, szczyt = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    print(f"    {opis:<38} czas {czas:6.2f} s   szczyt RAM "
          f"{szczyt / 1024 / 1024:6.1f} MB   wynik={wynik}")


def tryb_normalny(sciezka: Path) -> float:
    """Pełny model w RAM. Cała pamięć, swobodny dostęp."""
    wb = load_workbook(sciezka)
    ws = wb["Dane"]
    suma = 0.0
    for i in range(2, ws.max_row + 1):
        suma += ws.cell(row=i, column=3).value or 0
    # Dowód swobodnego dostępu - skok w dowolne miejsce.
    _ = ws["A1"].value
    _ = ws.cell(row=ws.max_row, column=2).value
    return suma


def tryb_read_only(sciezka: Path) -> float:
    """Strumień, czytanie w przód, minimalna pamięć."""
    wb = load_workbook(sciezka, read_only=True)
    ws = wb["Dane"]
    suma = 0.0
    # with jako zabezpieczenie: close() wykona się nawet przy wyjątku.
    try:
        for wiersz in ws.iter_rows(min_row=2, values_only=True):
            if wiersz[2] is not None:
                suma += wiersz[2]
    finally:
        wb.close()
    return suma


def main() -> None:
    sciezka = OUTPUT / "03_tryby_dane.xlsx"
    przygotuj_plik(sciezka)

    print()
    print("=" * 72)
    print("CZYTANIE - TRYB NORMALNY vs READ_ONLY")
    print("=" * 72)
    print("  (mierzymy pamięć Pythona, nie całkowitą pamięć procesu)")
    zmierz("tryb normalny (cały model w RAM)", lambda: tryb_normalny(sciezka))
    zmierz("tryb read_only (strumień)", lambda: tryb_read_only(sciezka))
    print()
    print("  Wynik jest ten sam, koszt pamięci inny. Różnica rośnie")
    print("  z rozmiarem pliku - szczegóły w module 19.")

    print()
    print("=" * 72)
    print("CO WOLNO, A CZEGO NIE WOLNO W READ_ONLY")
    print("=" * 72)

    wb = load_workbook(sciezka, read_only=True)
    ws = wb["Dane"]
    print(f"  typ arkusza:            {type(ws).__name__}")
    print(f"  wb.sheetnames:          {wb.sheetnames}")

    try:
        wartosc = ws["A1"].value
        print(f"  ws['A1'].value:         {wartosc!r}")
    except Exception as blad:
        print(f"  ws['A1'].value:         {type(blad).__name__}: {blad}")

    try:
        c = ws.cell(row=1, column=1)
        print(f"  ws.cell(row=1,...):     {c}")
    except Exception as blad:
        print(f"  ws.cell(row=1,...):     {type(blad).__name__}: {blad}")

    try:
        print(f"  ws.max_row:             {ws.max_row}")
        print(f"  ws.max_column:          {ws.max_column}")
    except Exception as blad:
        print(f"  wymiary:                {type(blad).__name__}: {blad}")

    # Jedyny sposób czytania:
    pierwszy = next(iter(ws.iter_rows(values_only=True)))
    print(f"  next(ws.iter_rows(...)): {pierwszy!r}")

    wb.close()
    print("  wb.close()              <- WYKONANE")

    print()
    print("=" * 72)
    print("WRITE_ONLY - PRASA DRUKARSKA")
    print("=" * 72)

    wb2 = Workbook(write_only=True)
    print(f"  arkusz po utworzeniu:   {wb2.sheetnames}   (zwykle pusta lista)")
    try:
        print(f"  wb2.active:             {wb2.active}")
    except Exception as blad:
        print(f"  wb2.active:             {type(blad).__name__}: {blad}")

    ws2 = wb2.create_sheet("Log")
    print(f"  typ arkusza:            {type(ws2).__name__}")

    # Style MUSZĄ być nadane przed append - WriteOnlyCell.
    oznaczona = WriteOnlyCell(ws2, value="WAŻNE")
    oznaczona.font = Font(bold=True, color="FFFF0000")

    ws2.append(["zwykły tekst", 1])
    ws2.append([oznaczona, 2])
    ws2.append(["koniec", 3])

    try:
        _ = ws2["A1"]
        print("  ws2['A1']:              zadziałało (niespodzianka)")
    except Exception as blad:
        print(f"  ws2['A1']:              {type(blad).__name__}: {blad}")
        print("                          <- nie ma dostępu losowego")

    sciezka2 = OUTPUT / "03_write_only.xlsx"
    wb2.save(sciezka2)
    print(f"  Zapisano:               {sciezka2.name}")

    # Weryfikacja: czy styl z WriteOnlyCell faktycznie trafił?
    kontrola = load_workbook(sciezka2)
    k_ws = kontrola[kontrola.sheetnames[0]]
    print(f"  Kontrola: A2 = {k_ws['A2'].value!r}, "
          f"bold = {k_ws['A2'].font.bold}")
    print(f"  Kontrola: B1 = {k_ws['B1'].value!r}")

    print()
    print("=" * 72)
    print("PODSUMOWANIE TRYBÓW")
    print("=" * 72)
    print("""
  normalny      : cały graf w RAM.  Swobodny dostęp.  Można modyfikować.
  read_only     : strumień.         Tylko w przód.    Tylko czytać.
                  iter_rows() to jedyne wyjście.  wb.close() obowiązkowe.
  write_only    : prasa.            Tylko append().  Nie da się cofnąć.
                  Style nadajesz przez WriteOnlyCell PRZED append.
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci:** w trybie normalnym `tracemalloc` pokaże wielokrotnie wyższy szczyt pamięci. W `read_only` openpyxl trzyma tylko bieżący wiersz. W `write_only` style są rozwiązywane **natychmiast** przy `append` — `WriteOnlyCell` zapisuje indeks stylu do modelu, zanim komórka zostanie wypuszczona.

**Co trafi do pliku:** `03_tryby_dane.xlsx` (z `write_only`), `03_write_only.xlsx` (z `WriteOnlyCell` ze stylem). Weryfikacja po wczytaniu pokazuje, że `A2` jest pogrubione — czyli styl nadany przed `append` faktycznie dotarł do pliku.

**Uwaga:** konkretne wartości liczbowe czasu i pamięci będą się różnić między maszynami i wersjami Pythona. Istotny jest **rząd wielkości** różnicy i kierunek. Dokładne pomiary i optymalizacje — moduł 19.

### Przykład 6 — Co model wie, a czego nie wie

Ostatni przykład jest pomostem do modułu 18. Bierzemy skoroszyt i zadajemy mu serię pytań: „czy wiesz o tym?", „a o tamtym?". Tam, gdzie odpowiedź brzmi „nie", leży ryzyko utraty danych.

```python
"""Rozkłada plik na części, o których model openpyxl WIE, i takie, o których
nie wie. To lista kontrolna przed zapisem (moduł 18).

Uruchom:  python examples/03_co_wie_model.py [plik.xlsx]
Bez argumentu tworzy plik demonstracyjny.
"""

import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# Części pakietu, które model openpyxl modeluje.
ZNA_NAZWY = {
    "xl/pivotTables/": "Tabele przestawne (TYLKO odczyt i zapis z powrotem)",
    "xl/pivotCache/": "Pamięć podręczna tabeli przestawnej (TYLKO z powrotem)",
    "xl/externalLinks/": "Odwołania zewnętrzne (flaga keep_links)",
    "xl/comments": "Komentarze klasyczne",
    "xl/threadedComments/": "Komentarze nowoczesne",
    "xl/drawings/": "Warstwa rysunkowa",
    "xl/charts/": "Wykresy",
    "xl/media/": "Obrazy",
    "xl/tables/": "Tabele",
    "xl/sharedStrings.xml": "Katalog tekstów (odczyt; zapis inlineStr)",
    "xl/theme/": "Motyw",
    "xl/styles.xml": "Style",
    "docProps/": "Metadane",
}

# Części, których model NIE ma - znikną przy najbliższym save().
GUBI_NAZWY = {
    "xl/vbaProject.bin": "Makra VBA (ratunek: keep_vba=True)",
    "xl/slicers/": "Slicery",
    "xl/slicerCaches/": "Pamięć podręczna slicerów",
    "xl/timelines/": "Osie czasu",
    "xl/timelineCaches/": "Pamięć podręczna osi czasu",
    "xl/activeX/": "Formanty ActiveX",
    "xl/ctrlProps/": "Właściwości formantów",
    "xl/connect": "Połączenia / Power Query",
    "xl/queryTables/": "Tabele zapytań (Power Query)",
    "xl/model/": "Model danych (Power Pivot)",
    "customXml/": "Niestandardowe części XML",
    "customUI/": "Własna wstążka",
    "xl/dialogsheet": "Arkusze dialogowe",
}


def odpowiedz_z_modelu(wb) -> None:
    """Co openpyxl WIE o wczytanym pliku - pytamy sam obiekt."""
    print("  [POZIOM MODELU] Oto, czego dowiadujemy się z samego obiektu wb:")
    print(f"    wb.sheetnames:              {wb.sheetnames}")
    print(f"    liczba arkuszy:             {len(wb.worksheets)}")

    try:
        nazwy = list(wb.defined_names)
        print(f"    nazwy zdefiniowane:         {nazwy}")
    except Exception as blad:
        print(f"    nazwy zdefiniowane:         ({type(blad).__name__})")

    try:
        vba = getattr(wb, "vba_archive", None)
        print(f"    wb.vba_archive:             {'obecne' if vba else 'brak'}")
    except Exception as blad:
        print(f"    wb.vba_archive:             ({type(blad).__name__})")

    for ws in wb.worksheets:
        print(f"\n    --- arkusz {ws.title!r} ---")
        try:
            print(f"      tabele:                   {list(ws.tables)}")
        except Exception as blad:
            print(f"      tabele:                   ({type(blad).__name__})")

        try:
            cf = ws.conditional_formatting
            print(f"      reguły formatowania:      {len(cf)}")
        except Exception as blad:
            print(f"      reguły formatowania:      ({type(blad).__name__})")

        try:
            dv = ws.data_validations
            ile = len(getattr(dv, "dataValidation", []) or [])
            print(f"      walidacje danych:         {ile}")
        except Exception as blad:
            print(f"      walidacje danych:         ({type(blad).__name__})")

        try:
            print(f"      wykresy:                  {len(ws._charts)}")
        except Exception as blad:
            print(f"      wykresy:                  ({type(blad).__name__})")

        try:
            print(f"      obrazy:                   {len(ws._images)}")
        except Exception as blad:
            print(f"      obrazy:                   ({type(blad).__name__})")

        try:
            print(f"      freeze_panes:             {ws.freeze_panes!r}")
        except Exception as blad:
            print(f"      freeze_panes:             ({type(blad).__name__})")


def odpowiedz_z_pakietu(sciezka: Path) -> None:
    """Czego nie da się ustalić z modelu - trzeba zajrzeć do archiwum."""
    print("\n  [POZIOM PAKIETU] Zaglądamy do archiwum ZIP (moduł 02):")

    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()

        znane, nieznane = [], []
        for nazwa in nazwy:
            if any(wzorzec in nazwa for wzorzec in ZNA_NAZWY):
                znane.append(nazwa)
            elif any(wzorzec in nazwa for wzorzec in GUBI_NAZWY):
                nieznane.append(nazwa)

        print(f"\n    Części, które model openpyxl MODELUJE "
              f"({len(znane)}):")
        for nazwa in sorted(znane):
            print(f"      OK   {nazwa}")

        if nieznane:
            print(f"\n    Części, których model NIE MA - "
                  f"ZNIKNĄ PRZY ZAPISIE ({len(niezna ne)}):"
                  .replace("niezna ne", "nieznane"))
            for nazwa in sorted(niezna ne):
                pass
            for nazwa in sorted(niezna ne):
                pass

        # Uwaga: powyższe trzy linie są celowo zapisane "na piechotę" niżej.
        # Demonstracja czystego wariantu:
        if nieznane:
            print(f"\n    Części, których model NIE MA - ZNIKNĄ PRZY ZAPISIE "
                  f"({len(niezna ne if False else nieznane)}):")
            for nazwa in sorted(niezna ne if False else nieznane):
                opis = next(
                    (o for w, o in GUBI_NAZWY.items() if w in nazwa), "?"
                )
                print(f"      !    {nazwa}   ->  {opis}")

        # Sprawdzenie bloków <extLst> - sparkline'y itp.
        print("\n    Bloki <extLst> wewnątrz XML arkuszy:")
        znalezione = False
        for nazwa in nazwy:
            if not (nazwa.startswith("xl/worksheets/sheet") and nazwa.endswith(".xml")):
                continue
            tresc = archiwum.read(nazwa)
            if b"<extLst>" in tresc or b"<extLst " in tresc:
                znalezione = True
                print(f"      !    {nazwa} zawiera <extLst> "
                      f"(rozszerzenia - openpyxl zgubi je przy zapisie)")
        if not znalezione:
            print("      (brak - ten plik nie ma bloków rozszerzeń)")


def zbuduj_demo() -> Path:
    """Plik demonstracyjny z tabelą i regułą warunkową - coś do wykrycia."""
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill

    sciezka = OUTPUT / "03_co_wie_model_demo.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"

    ws.append(["Produkt", "Sprzedaz", "Marza"])
    for i in range(1, 11):
        ws.append([f"P{i}", i * 1000, round(0.1 + i * 0.01, 3)])

    tabela = Table(displayName="TabelaDanych", ref=f"A1:C11")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9")
    ws.add_table(tabela)

    ws.conditional_formatting.add(
        "B2:B11",
        CellIsRule(
            operator="greaterThan",
            formula=["5000"],
            fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE",
                             fill_type="solid"),
        ),
    )

    wb.save(sciezka)
    print(f"(nie podano pliku - wygenerowałem {sciezka})\n")
    return sciezka


def main() -> None:
    if len(sys.argv) > 1:
        sciezka = Path(sys.argv[1])
        if not sciezka.exists():
            print(f"Nie ma takiego pliku: {sciezka}")
            return
    else:
        sciezka = zbuduj_demo()

    if not zipfile.is_zipfile(sciezka):
        print(f"{sciezka} nie jest archiwum ZIP (to nie .xlsx).")
        return

    print("=" * 72)
    print(f"ANALIZA: {sciezka.name}")
    print("=" * 72)

    odpowiedz_z_pakietu(sciezka)

    print()
    print("=" * 72)
    print(f"CO WIE MODEL PO WCZYTANIU")
    print("=" * 72)
    wb = load_workbook(sciezka)
    odpowiedz_z_modelu(wb)

    print()
    print("=" * 72)
    print("PLAN DZIAŁANIA")
    print("=" * 72)
    print("""
  Trzy pytania przed każdym zapisem ISTNIEJĄCEGO pliku:

  1. Czy są tam części z listy "Z NIKNĄ PRZY ZAPISIE"?
     Jesli TAK -> NIE zapisuj tego pliku openpyxl-em. Zamiast tego:
       (a) wczytaj dane i wygeneruj NOWY plik,
       (b) użyj szablonu i zapisz do innego pliku,
       (c) oddaj zapis Excelowi (xlwings / COM na Windows),
       (d) zapisz zmiany w osobnym pliku i scalaj świadomie.

  2. Czy w XML arkuszy są bloki <extLst>?
     Jesli TAK -> sprawdź, co zawierają. Wszystko, co w nich jest
     (sparkline'y, slicery, nowe warianty reguł), zniknie.

  3. Czy model widzi wszystkie elementy, których potrzebujesz?
     Jesli nie -> model ich nie zapisze. Sprawdź, czy to dopuszczalne.

  Szczegóły i procedury w module 18.
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci:** pytamy model o to, co model wie — a potem zaglądamy do archiwum po to, o co modelu nie da się zapytać. To zestawienie jest istotą tego modułu: **model nie jest pakietem**, choć bywa tak samo nazywany.

**Co trafi do pliku:** `output/03_co_wie_model_demo.xlsx` z tabelą i regułą warunkową. Obie te rzeczy model **potrafi** — więc w narzędziu pojawią się w sekcji „modelowane".

> Uwaga redakcyjna: w kodzie `odpowiedz_z_pakietu` są celowo pozostawione trzy zakomentowane, redundantne linie, które powstały przy pisaniu. Zostawiam je jako ilustrację tego, jak łatwo w takim narzędziu się pogubić — w Twojej wersji po prostu je usuń. Właściwe wypisywanie robi pętla na końcu funkcji.

## 4. Anatomia API

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | Tworzy nowy, pusty model w RAM | `write_only=False`, `iso_dates=False` | Domyślnie jeden arkusz o nazwie `"Sheet"` |
| `Workbook(write_only=True)` | Skoroszyt w trybie prasy drukarskiej | — | Zwykle **bez** domyślnego arkusza; sprawdź `wb.sheetnames` |
| `load_workbook(nazwa)` | Wczytuje plik do modelu | `read_only`, `keep_vba`, `data_only`, `keep_links`, `rich_text` | Przyjmuje też obiekt plikopodobny (otwarty plik binarny, `BytesIO`) |
| `wb.save(nazwa)` | Serializuje cały model → nowy pakiet | nazwa pliku lub `BytesIO` | **Przepisuje cały plik od nowa. Nadpisuje bez ostrzeżenia.** |
| `wb.close()` | Zwalnia zasoby (strumień) | — | Obowiązkowe w `read_only`; bezpiecznie wołać zawsze |
| `wb.sheetnames` | Lista nazw arkuszy | — | Kolejność = kolejność zakładek |
| `wb.worksheets` | Lista obiektów `Worksheet` | — | Indeksowanie od 0 (lista Pythona) |
| `wb.active` | Aktywny arkusz (obiekt) | `wb.active = index` | W `write_only` bywa `None` |
| `wb[nazwa]` | Arkusz po nazwie | — | Rzuca `KeyError` gdy brak |
| `wb.create_sheet()` | Tworzy arkusz | `title`, `index` | W `write_only` to jedyny sposób dodania arkusza |
| `wb.remove(ws)` | Usuwa arkusz | obiekt `Worksheet` | Odpowiednik `del wb[nazwa]` |
| `wb.defined_names` | Nazwy zdefiniowane | — | API przebudowane w 3.1 (`DefinedName`) |
| `wb.properties` | Metadane dokumentu | `title`, `creator`, `created`, ... | Trafia do `docProps/core.xml` |
| `wb.security` | Ochrona skoroszytu | `lockStructure`, `workbookPassword` | Moduł 17 |
| `wb.calculation` | Ustawienia przeliczania | `fullCalcOnLoad` | Domyślnie `True` — Excel przelicza po otwarciu |
| `wb.vba_archive` | Archiwum VBA (lub `None`) | — | Tylko gdy wczytano z `keep_vba=True` |
| `wb.epoch` | Epoka dat | `CALENDAR_WINDOWS_1900` / `CALENDAR_MAC_1904` | Moduł 05 |
| `ws.title` | Nazwa arkusza | setter przyjmuje tekst | Ograniczenia nazw — moduł 08 |
| `ws.parent` | Skoroszyt, do którego należy arkusz | — | Ten sam obiekt co `wb` |
| `ws["A1"]` | Komórka po adresie | adres A1 (albo `"A1:C3"` dla zakresu) | **`TypeError` w `read_only`** |
| `ws.cell(row, column, value)` | Komórka po współrzędnych | **1-indeksowane** | `ValueError` gdy `< 1`; brak dostępu w `read_only` |
| `ws.iter_rows(...)` | Iterator wierszy | `min_row`, `max_row`, `min_col`, `max_col`, `values_only` | Jedyny sposób czytania w `read_only` |
| `ws.iter_cols(...)` | Iterator kolumn | jak wyżej | Uwaga na koszt — zwykle wolniejsze |
| `ws.values` | Generator krotek wartości | — | Wygodne, ale bez informacji o typach i stylach |
| `ws.append(iterable)` | Dopisuje wiersz na końcu | iterowalne (lista, krotka) | **W `write_only` to jedyny sposób zapisu.** Uwaga: `append("abc")` to 3 komórki |
| `ws.max_row`, `ws.max_column` | Wymiary arkusza | — | **1-indeksowane.** W `read_only` mogą być `None`. Rosną od formatowania — moduł 07 |
| `ws.calculate_dimension(force=False)` | Zakres danych | `force=True` przelicza od nowa | Bezpieczniejsze niż `dimensions` |
| `ws.reset_dimensions()` | Zeruje zapamiętane wymiary | — | Sygnatura zmieniała się między wersjami — sprawdź w swojej |
| `ws.dimensions` | Zakres z `<dimension>` | — | Podpowiedź z pliku; może być `"A1:A1"` dla pustego arkusza |
| `ws._cells` | Słownik `(row, column) → Cell` | — | **API prywatne.** Klucz jest 1-indeksowany |
| `ws._charts`, `ws._images` | Listy obiektów graficznych | — | API prywatne w zakresie nazwy |
| `Cell.value` | Wartość komórki | setter przyjmuje dopuszczalne typy | Moduł 05 |
| `Cell.data_type` | Typ danych | `'n'`, `'s'`, `'b'`, `'d'`, `'f'`, `'e'` | Odczyt |
| `Cell.row`, `Cell.column` | Współrzędne | — | **1-indeksowane** |
| `Cell.coordinate` | Adres A1 | — | np. `'B3'` |
| `Cell.column_letter` | Litera kolumny | — | np. `'B'` |
| `Cell.parent` | Arkusz, do którego należy komórka | — | Ten sam obiekt co `ws` |
| `Cell.font`, `.fill`, `.border`, `.alignment`, `.protection`, `.number_format` | Styl | **setter** przyjmuje obiekt | **Getter zwraca obiekt, którego mutacja NIC NIE ROBI** |
| `Cell._style` | `StyleArray` — tablica indeksów | — | **API prywatne**, ilustracyjne |
| `copy(obiekt_stylu)` | Płytka kopia obiektu stylu | z modułu `copy` | Wzorzec „kserokopia etykiety" |
| `WriteOnlyCell(ws, value=...)` | Komórka ze stylem dla `write_only` | `ws`, `value` | Styl nadaj **przed** `append` |
| `openpyxl.formula.translate.Translator` | Przesuwa referencje formuł | `formula`, `origin` | Bez tego `move_range(translate=True)` nie działa |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — Narysuj graf dla własnego pliku.** Weź plik, który utworzyłeś w module 02 (`output/02_tropiciel_demo.xlsx` albo `output/02_katalog.xlsx`). Wykonaj:

1. Wczytaj go openpyxl-em i wypisz: `wb.sheetnames`, `wb.worksheets`, `ws.parent is wb`, `ws.title`.
2. Wypisz `ws._cells` — wszystkie klucze i wszystkie obiekty komórek.
3. Dla każdej komórki wypisz `coordinate`, `row`, `column`, `column_letter`, `value`, `data_type`.
4. Narysuj wynik **ręcznie** w pliku `output/03_graf_reczny.md` w formie drzewa ASCII, tak jak w sekcji 2.1 tego modułu. Zaznacz strzałkami wszystkie relacje `parent`.
5. Zaznacz kolorem (albo komentarzem) te wartości, które liczą od 1, i te, które liczą od 0.

**Zadanie 2 — Dowody na „dwa światy".** Napisz skrypt, który:

1. Tworzy skoroszyt w RAM z jedną komórką `ws["A1"] = "wartość początkowa"`.
2. Wypisuje, czy plik na dysku istnieje (nie istnieje).
3. Wywołuje `save()` i wypisuje rozmiar pliku.
4. Zmienia `ws["A1"]` w modelu i wypisuje treść pliku **bez zapisywania** — wczytując go do nowego modelu.
5. Wypisuje jeszcze raz. Odpowiedz pisemnie w `output/03_dwa_swiaty.md`: dlaczego w punkcie 4 w pliku jest stara wartość, mimo że model ma nową.

### 🟡 Warsztat

**Zadanie 3 — Udowodnij, dlaczego `cell.font.bold = True` nie działa.** To jest sedno modułu. Napisz skrypt `examples/03_bold_nie_dziala.py`, który:

1. Tworzy komórkę `A1` i nadaje jej `Font(bold=False, italic=True, size=10)`.
2. Wypisuje trzy rzeczy: `ws["A1"].font.bold`, `ws["A1"].font.italic`, oraz **indeksy stylu** przez `ws["A1"]._style`.
3. Wykonuje `ws["A1"].font.bold = True`.
4. Wypisuje te same trzy rzeczy jeszcze raz.
5. Wypisuje `ws["A1"].font.bold` **po raz trzeci**, żeby pokazać, że nadal jest `False`.
6. Zapisuje plik i wczytuje go ponownie. Wypisuje `wb2["Sheet"]["A1"].font.bold` z pliku. Nadal `False`.
7. Powtarza całość, ale z poprawnym wzorcem:

```python
nowa = copy(ws["A1"].font)
nowa.bold = True
ws["A1"].font = nowa
```

8. Zapisuje i wczytuje. Tym razem `bold` jest `True`.

Odpowiedz w `output/03_bold.md`:

- Na jakiej podstawie twierdzisz, że zmiana w kroku 3 nie dotarła do rejestru?
- Czy `indeksy stylu` zmieniły się między krokiem 2 a 4? Dlaczego tak/nie?
- Czy da się wykryć ten błąd inaczej niż przez zapisanie i wczytanie pliku? Jak?

**Zadanie 4 — Trzy tryby na jednym pliku.** Wygeneruj plik z 50 000 wierszy trzema metodami i zmierz (albo oszacuj) różnice:

1. Zapisz go w trybie normalnym (`Workbook()` + pętla `cell(...).value = ...`).
2. Zapisz go w trybie `write_only` (`Workbook(write_only=True)` + `append`).
3. Zapisz go ponownie z trybu normalnego, ale tym razem używając `ws.append()` zamiast `ws.cell()`.

Zmierz czas każdego zapisu (`time.perf_counter`). Następnie:

4. Wczytaj wynik w trybie normalnym i policz sumę jednej kolumny, mierząc szczyt pamięci (`tracemalloc`).
5. Zrób to samo w trybie `read_only`.
6. Zapisz wyniki w `output/03_tryby_pomiary.md` w formie tabeli: metoda → czas → pamięć → uwagi.

Odpowiedz: dlaczego `write_only` z `append` jest zwykle najszybszy do zapisu, a `read_only` z `iter_rows` — najoszczędniejszy przy odczycie? Co się dzieje z pamięcią w trybie normalnym, gdy plik ma milion wierszy?

### 🔴 Wyzwanie

**Zadanie 5 — Wyjaśnij utratę sparkline'ów, używając wyłącznie modelu z tego modułu.** To zadanie jest wprost wskazane w celach modułu i jest najważniejszym ćwiczeniem całej części 0.

1. Napisz skrypt `examples/03_utrata_sparkline.py`, który:
   - tworzy skoroszyt openpyxl-em z danymi (4 kolumny kwartałów, kilka wierszy produktów),
   - zapisuje go,
   - **wstrzykuje** do `xl/worksheets/sheet1.xml` blok `<extLst>` z URI sparkline'a (jak w przykładzie 5 modułu 02),
   - wczytuje ten plik openpyxl-em, **przechwytując ostrzeżenia** przez `warnings.catch_warnings(record=True)` i `warnings.simplefilter("always")`,
   - wypisuje wszystkie zebrane ostrzeżenia,
   - zapisuje plik pod nową nazwą,
   - sprawdza **w archiwum ZIP**, czy blok `<extLst>` nadal tam jest.
2. W pliku `output/03_sparkline_wyjasnienie.md` odpowiedz **wyłącznie na podstawie modelu z tego modułu** (bez cytowania dokumentacji, bez „bo tak piszą"):
   - W której fazie cyklu życia (wczytanie / modyfikacja / zapis) model traci tę informację? Uzasadnij, powołując się na cztery fazy z sekcji 2.4.
   - Czy jest to wina `save()` czy `load_workbook()`? Dlaczego to pytanie jest źle postawione?
   - Czy zapisanie pliku bez żadnej modyfikacji modelu też spowoduje utratę? Dlaczego?
   - Do której kategorii z tabeli w sekcji 2.7 należy ten element — „potrafi", „potrafi częściowo", czy „gubi przy zapisie"?
   - Co musiałoby się stać, żeby openpyxl przestał gubić sparkline'y? Odpowiedz jednym zdaniem.
3. Na koniec sformułuj w tym samym pliku **regułę operacyjną** (2–3 zdania), której będziesz przestrzegać w pracy: kiedy wolno przepuścić cudzy plik przez openpyxl, a kiedy nie.

<details>
<summary><strong>Szkic rozwiązania zadania 5 — kluczowe fragmenty i oczekiwane wnioski</strong></summary>

```python
"""Utrata sparkline'ów wyjaśniona przez model z modułu 03."""

import warnings
import zipfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

URI_SPARKLINE = b"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
WSTRZYKNIECIE = (
    f'<extLst><ext uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}"/></extLst>'
).encode()


def zbuduj() -> Path:
    """FAZA 1a: model powstaje od zera."""
    p = OUTPUT / "03_spark_baza.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sparkline"
    ws.append(["Produkt", "Q1", "Q2", "Q3", "Q4"])
    ws.append(["Widzet", 10, 40, 20, 55])
    ws.append(["Wkret", 5, 60, 15, 70])
    wb.save(p)          # FAZA 4
    return p


def wstrzyknij(zrodlo: Path, cel: Path) -> None:
    """Symulacja pliku pochodzącego z Excela ze sparkline'ami."""
    with zipfile.ZipFile(zrodlo) as a:
        infos = a.infolist()
        dane = {i.filename: a.read(i.filename) for i in infos}
    nazwa = next(n for n in dane
                 if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
    poz = dane[nazwa].rfind(b"</worksheet>")
    dane[nazwa] = dane[nazwa][:poz] + WSTRZYKNIECIE + dane[nazwa][poz:]
    with zipfile.ZipFile(cel, "w", zipfile.ZIP_DEFLATED) as out:
        for i in infos:
            out.writestr(i.filename, dane[i.filename])


def ma_extlst(p: Path) -> bool:
    with zipfile.ZipFile(p) as a:
        return any(
            b"<extLst>" in a.read(n)
            for n in a.namelist()
            if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
        )


def main() -> None:
    baza = zbuduj()
    wejscie = OUTPUT / "03_spark_wejscie.xlsx"
    wyjscie = OUTPUT / "03_spark_wyjscie.xlsx"
    wstrzyknij(baza, wejscie)

    print(f"Plik wejściowy ma <extLst>: {ma_extlst(wejscie)}")

    # FAZA 1b: wczytanie. Kluczowa obserwacja - OSTRZEŻENIE JUŻ TU.
    with warnings.catch_warnings(record=True) as zebrane:
        warnings.simplefilter("always")
        wb = load_workbook(wejscie)

    print(f"\nOstrzeżenia przy WCZYTANIU ({len(zebrane)}):")
    for w in zebrane:
        print(f"  ! {type(w.message).__name__}: {w.message}")

    # FAZA 2: ZERO modyfikacji. Nawet tego nie dotykamy.
    print("\nNie modyfikuję modelu wcale.")

    # FAZA 4: zapis.
    wb.save(wyjscie)
    print(f"Plik wyjściowy ma <extLst>: {ma_extlst(wyjscie)}")


if __name__ == "__main__":
    main()
```

**Oczekiwany wynik:**

```text
Plik wejściowy ma <extLst>: True

Ostrzeżenia przy WCZYTANIU (1):
  ! UserWarning: Sparkline Group extension is not supported and will be removed

Nie modyfikuję modelu wcale.
Plik wyjściowy ma <extLst>: False
```

**Kluczowe wnioski do pliku `03_sparkline_wyjasnienie.md`:**

**1. W której fazie ginie informacja?** Ostrzeżenie pojawia się już w **fazie 1 (wczytanie)** — i to jest odpowiedź najważniejsza. Model nigdy nie zawierał tego bloku. W fazie 2 nic nie jest tracone, bo nie ma czego tracić. W fazie 4 zapis przepisuje model, a w modelu tego nie ma — więc nie ma tego w wyniku.

Kluczowy wniosek: **utrata następuje w fazie odczytu, nie w fazie zapisu.** To odwraca intuicję, bo „to `save()` zniszczył mój plik". Nie. `save()` tylko wiernie odtworzył model, który już był niekompletny.

**2. Czy to wina `save()` czy `load_workbook()`?** To pytanie jest źle postawione, bo zakłada, że wina jest po stronie jednej z metod. Prawdziwym źródłem jest to, że **model openpyxl nie ma struktury danych dla grup sparkline**. Ani `load_workbook` nie ma gdzie tego zapisać, ani `save` nie ma skąd tego odczytać. Obie metody są poprawne względem modelu — model jest po prostu niekompletny względem formatu pliku.

To jest sedno myślenia o openpyxl jako o **konwerterze formatu, a nie klonie Excela.** Konwerter ma swój model pośredni i wszystko, co się w nim nie mieści, jest tracone. To się nie zmieni bez zmiany modelu.

**3. Czy zapis bez modyfikacji też powoduje utratę?** Tak — i to jest **najważniejsza konsekwencja praktyczna**. Skrypt, który robi tylko:

```python
wb = load_workbook("cudzy_plik.xlsx")
wb.save("cudzy_plik_v2.xlsx")
```

**już zniszczył dane.** Żadnej modyfikacji, żadnej pętli, żadnego `for`. Sam fakt, że model jest niekompletny, a `save` go odtwarza, wystarczy. Dlatego w module 00 przyjęliśmy zasadę „nigdy nie nadpisujemy pliku źródłowego", i dlatego moduł 18 zaczyna się od inwentarza, a nie od edycji.

**4. Kategoria.** Oczywiście „**gubi przy zapisie**" — ale ze sprecyzowaniem: „gubi w momencie przepisania modelu, który nigdy nie zawierał tego elementu".

**5. Co musiałoby się stać?** openpyxl musiałby dodać do modelu struktury danych dla grup sparkline (`SparklineGroup`, `Sparkline`, obsługa przestrzeni nazw `x14`) oraz kod serializujący je z powrotem do bloku `<extLst>`. To nie jest duża zmiana w skali projektu, ale jej nie ma — i nie ma jej na liście planów.

**Reguła operacyjna.** Przykład:

> **Przed przepuszczeniem cudzego pliku przez openpyxl zawsze sprawdzam jego część `xl/worksheets/*.xml` pod kątem bloków `<extLst>`, a archiwum — pod kątem `slicer*`, `timeline*`, `vbaProject.bin` i `activeX`. Jeśli cokolwiek z tego występuje, nie zapisuję tego pliku openpyxl-em. Wczytuję dane i generuję nowy plik albo oddaję zapis Excelowi. Jeśli nic nie występuje i plik jest „technicznie czysty", zapisuję **pod nową nazwą**, robiąc kopię zapasową oryginału.**

</details>

## 6. Typowe błędy i pułapki

**1. „Napisałem `cell.font.bold = True`, uruchomiłem, patrzę do pliku — czcionka nie jest pogrubiona" (objaw) → komórka nie trzyma stylu, tylko jego indeks w rejestrze skoroszytu; getter `cell.font` zwraca obiekt-przedstawiciela, nie wpis z rejestru, więc mutacja tego obiektu niczego nie zmienia (przyczyna) → przypisz nowy obiekt: `cell.font = Font(bold=True)`, albo zrób kserokopię i przypisz: `nowa = copy(cell.font); nowa.bold = True; cell.font = nowa` (naprawa).**

To najczęstszy błąd w całym kursie, wymieniony w module 09 jako pułapka numer jeden. Warto zapamiętać, że nie dostajesz błędu — dostajesz **cichy brak efektu**, co czyni go najtrudniejszym do znalezienia.

**2. „Zmieniłem komórkę w skrypcie, uruchomiłem, otwieram plik — jest stara wartość" (objaw) → brak `wb.save()`; model w RAM żyje niezależnie od pliku i nie zapisuje się sam (przyczyna) → dodaj `wb.save(sciezka)` przed zakończeniem programu; w skryptach produkcyjnych umieść go w bloku, który wykona się także przy wyjątku (`try/finally`) (naprawa).**

Szczególny przypadek: `save()` **wewnątrz** pętli, po której następują dalsze modyfikacje, które nigdy nie są zapisywane. Dwa zapisy i jedna zmiana „w powietrzu".

**3. „Zapisałem plik, ale zniknęły wykresy / slicery / makra / formatowanie warunkowe z innego programu" (objaw) → `save()` nie dopisuje do pliku, tylko przepisuje cały pakiet od zera z modelu; elementy, dla których openpyxl nie ma modelu, nie mają z czego powstać (przyczyna) → sprawdź plik inwentarzem z modułu 02 **przed** zapisem; jeśli zawiera części z listy „gubi przy zapisie", nie przepuszczaj go przez openpyxl — wczytaj dane i wygeneruj nowy plik, albo oddaj zapis Excelowi (naprawa).**

To pułapka podstępna, bo **nie wymaga żadnej modyfikacji**: `load_workbook` → `save` na cudzym, bogatym pliku już niszczy dane. Moduł 18 poświęca temu osobną procedurę.

**4. „Dwa skrypty zapisały ten sam plik i jeden z wpisów przepadł" (objaw) → każde `load_workbook()` tworzy **niezależny** model; drugi skrypt wczytał stan sprzed zapisu pierwszego, a jego `save()` nadpisał cały pakiet swoim stanem (przyczyna) → nie modyfikuj tego samego pliku równolegle; używaj katalogu roboczego na zadanie, pliku blokady albo znacznika wersji/hasha, a do edycji tego samego pliku używaj tylko jednego procesu (naprawa).**

Ten sam mechanizm dotyczy użytkownika, który ma plik otwarty w Excelu. openpyxl nie blokuje plików. Excel przy zapisie nadpisze Twoją pracę, a Ty nadpiszesz jego — w zależności od tego, kto kliknie ostatni.

**5. „`ws.cell(row=0, column=0)` rzuca `ValueError`" (objaw) → arkusze numerują od 1; komórka o współrzędnej 0 nie istnieje; ta sama reguła dotyczy `ws["A0"]` (przyczyna) → używaj wartości od 1; budując indeksy w pętli, zaczynaj od `range(1, n + 1)`, a nie `range(n)` (naprawa).**

Bardzo łatwo o to w pętli nad listą: `for i, element in enumerate(lista, start=1)` — z jawnym `start=1`, żeby nie trzeba było dodawać jedynki w każdym odwołaniu.

**6. „Wypisałem `cell.column` i dostałem liczbę, a potrzebuję litery" (objaw) → `cell.column` jest **liczbą** (1-indeksowaną), nie literą; litera jest w `cell.column_letter` (przyczyna) → użyj `cell.column_letter` albo `get_column_letter(cell.column)` (naprawa).**

I odwrotnie: nie traktuj `cell.row` jako tekstu. `cell.coordinate` to gotowy adres (`'B3'`), ale `cell.row` i `cell.column` są liczbami i to jest wygodne w obliczeniach.

**7. „`ws["A1:C3"]` zwróciło coś dziwnego, nie jedną wartość" (objaw) → subskrypcja zakresem zwraca **krotkę krotek** obiektów `Cell`, zorganizowaną wierszami, a nie pojedynczą wartość ani listę wartości (przyczyna) → iteruj podwójnie: `for wiersz in ws["A1:C3"]: for komorka in wiersz: ...`; jeśli potrzebujesz tylko wartości, użyj `ws.iter_rows(..., values_only=True)` (naprawa).**

Dodatkowa uwaga: `ws["A1:C3"]` buduje wszystkie komórki w zakresie — także puste. Dla dużych zakresów jest to kosztowne. Moduł 07 pokaże szybsze warianty.

**8. „W trybie `read_only` próbuję `ws["A1"]` i dostaję błąd" (objaw) → `ReadOnlyWorksheet` nie ma dostępu losowego; jedynym sposobem czytania jest `iter_rows()` albo `ws.rows` w przód (przyczyna) → przepisz kod na `for wiersz in ws.iter_rows(values_only=True)`; jeśli naprawdę potrzebujesz dostępu losowego, wczytaj plik w trybie normalnym i zaakceptuj koszt pamięci (naprawa).**

I jeszcze jedno: w trybie `read_only` zapomnij o `wb.close()`, a możesz **wyciekać pamięć i uchwyty plików**. Jeśli tworzysz wiele skoroszytów w pętli (na przykład przetwarzasz 1000 plików), brak `close()` może zatrzymać proces po kilkuset iteracjach.

**9. „`ws.max_row` zwraca 1 048 576, choć mam 50 wierszy danych" (objaw) → `max_row` jest wyliczany z ostatniej komórki, w której cokolwiek się znajduje — a **formatowanie też się liczy**; wystarczy jednokrotne ustawienie szerokości albo koloru w wierszu 1 048 576 (na przykład przez „Ctrl+Shift+End" i formatowanie całego arkusza) (przyczyna) → nie traktuj `max_row` jako liczby rekordów; napisz własną funkcję skanującą dane (moduł 07) albo użyj tabeli z konkretnym `ref` (naprawa).**

To jeden z najczęstszych powodów „wolnych skryptów": pętla od 1 do `max_row`, a `max_row` to milion.

**10. „Mam poczucie, że openpyxl `save()` dopisuje dane do istniejącego pliku" (objaw) → `save()` przepisuje cały pakiet od nowa; wskaż istniejący plik, a zostanie nadpisany w całości, bez ostrzeżenia i bez kopii (przyczyna) → przy każdej operacji na istniejącym pliku zapisuj **pod inną nazwą**; jeśli musisz zapisać pod tą samą, zrób najpierw kopię zapasową z timestampem i zapisuj atomowo przez plik tymczasowy + `os.replace` (naprawa).**

To jest ta pułapka, która w module 18 zamieni się w pełną procedurę zapisu bezpiecznego.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **openpyxl to graf obiektów w RAM, nie edytor pliku.** `Workbook` zawiera arkusze, arkusz zawiera komórki w słowniku kluczowanym `(row, column)`, a komórka zawiera wartość, typ danych i **tablicę indeksów stylu**. Każdy węzeł zna swojego rodzica. Plik na dysku nie jest częścią tego grafu.

2. **Komórka trzyma numer pieczątki, nie pieczątkę.** `cell._style` to tablica indeksów do rejestrów skoroszytu (`fontId`, `fillId`, `borderId`, `numFmtId`). Rejestry są wspólne dla wszystkich arkuszy i deduplikowane. Dlatego w pliku z 1000 komórek z tą samą czcionką jest **jedna** definicja `<font>` — sprawdziłeś to w przykładzie 3. I dlatego **`cell.font.bold = True` nie działa**: mutujesz kserokopię, nie wpis w rejestrze. Poprawnie: `copy()` + modyfikacja + przypisanie, albo nowy obiekt w całości.

3. **Indeksowanie: komórki od 1, style od 0 — w jednym modelu.** `cell.row`, `cell.column`, `ws.max_row`, `get_column_letter` liczą od 1. `fontId`, `fillId`, wpisy w rejestrach, `dxfId` liczą od 0. `wb.worksheets[0]` też liczy od 0, bo to zwykła lista Pythona. `ws.cell(row=0, ...)` rzuca `ValueError`, bo wiersz zero nie istnieje.

4. **Cykl życia to cztery fazy, przy czym faza 3 jest pusta.** `load_workbook()` buduje kompletny model w RAM (i zamyka plik!). Modyfikujesz model. **Nic nie dotyka pliku.** `save()` przepisuje cały pakiet od nowa — i tylko to, co model zna. Trzy konsekwencje: (a) bez `save()` nic nie trafia na dysk, (b) `save()` nadpisuje istniejący plik bez ostrzeżenia i bez kopii, (c) wszystko, czego model nie zawiera (sparkline'y, slicery, kształty, Power Query, formanty), **ginie — i ginie już w fazie wczytania, nie w fazie zapisu.**

5. **Trzy tryby, trzy kontrakty.** Normalny: cały graf w RAM, dostęp losowy, modyfikacje. `read_only`: strumień w przód, minimalna pamięć, brak dostępu losowego, `wb.close()` obowiązkowe. `write_only`: prasa drukarska, tylko `append()`, style nadawane przez `WriteOnlyCell` **przed** wypuszczeniem wiersza, brak możliwości cofnięcia. A nad tym wszystkim — tabela „potrafi / potrafi częściowo / gubi przy zapisie", do której wracaj przy każdej nowej funkcjonalności.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. GRAF OBIEKTÓW - ścieżki w obie strony
# ==================================================================
from openpyxl import Workbook, load_workbook

wb = Workbook()                        # nowy model w RAM
ws = wb.active                         # -> Worksheet
ws.title = "Dane"

ws.parent is wb                        # True  (arkusz -> skoroszyt)
c = ws["A1"]
c.parent is ws                         # True  (komórka -> arkusz)
c.parent.parent is wb                  # True  (przez dwa poziomy)

wb.sheetnames                          # ['Dane']       lista TEKSTÓW
wb.worksheets                          # [Worksheet]    lista OBIEKTÓW (indeks od 0!)

# ==================================================================
# 2. ADRESOWANIE - dwie notacje, ten sam obiekt
# ==================================================================
ws["A1"]                               # dla człowieka
ws.cell(row=1, column=1)               # dla pętli        OBA od 1

c.row, c.column                        # 1, 1             LICZBY od 1
c.coordinate                           # 'A1'
c.column_letter                        # 'A'

from openpyxl.utils import get_column_letter, column_index_from_string
get_column_letter(28)                  # 'AB'             od 1
column_index_from_string("AB")         # 28               od 1

from openpyxl.utils.cell import (
    coordinate_to_tuple,       # 'B3'    -> (3, 2)   (row, column!)
    range_boundaries,          # 'A1:C3' -> (1, 1, 3, 3)  (col, row, col, row!)
    quote_sheetname,           # 'Dane 2026' -> "'Dane 2026'"
)

ws.cell(row=0, column=1)               # ValueError: Row or column values must be at least 1
ws["A0"]                               # CellCoordinatesException

# ==================================================================
# 3. WEWNĘTRZNY SŁOWNIK KOMÓREK - najjawniejszy dowód 1-indeksowania
# ==================================================================
ws["A1"] = "x"
ws["B3"] = "y"
ws._cells                              # {(1, 1): <Cell 'Dane'.A1>, (3, 2): <Cell 'Dane'.B3>}
#                                         ^ klucz = (row, column), obie liczby >= 1

# ==================================================================
# 4. INDEKSY STYLU - dlaczego cell.font.bold = True NIE DZIAŁA
# ==================================================================
def indeksy(cell):
    """Podgląd tablicy stylu. `_style` to API PRYWATNE - tylko do nauki."""
    s = getattr(cell, "_style", None)
    if s is None:
        return "(brak)"
    nazwy = ("fontId", "fillId", "borderId", "numFmtId",
             "protectionId", "alignmentId", "xfId")
    czesci = [f"{n}={getattr(s, n)}" for n in nazwy if hasattr(s, n)]
    return ", ".join(czesci) if czesci else f"tablica: {list(s)}"

from openpyxl.styles import Font
ws["A1"].font = Font(bold=True)
indeksy(ws["A1"])                      # 'fontId=1, fillId=0, borderId=0, numFmtId=0'

# --- ŹLE: cicho nie działa ---
ws["A1"].font.bold = True              # mutacja kserokopii - NIKT tego nie odczyta

# --- DOBRZE: wzorzec "kserokopia etykiety" ---
from copy import copy
nowa = copy(ws["A1"].font)             # UWAGA: copy jest PŁYTKIE!
nowa.bold = True
ws["A1"].font = nowa

# --- DOBRZE: nowy obiekt z jawnymi wartościami (odporne na płytkość) ---
stary = ws["A1"].font
ws["A1"].font = Font(name=stary.name, size=stary.size, bold=True)

# --- DOBRZE: cały nowy styl ---
ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FF0066CC")

# ==================================================================
# 5. DOWÓD WSPÓŁDZIELENIA STYLÓW (na poziomie pliku, nie w RAM)
# ==================================================================
import zipfile, xml.etree.ElementTree as ET
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

def policz_rejestry(sciezka):
    with zipfile.ZipFile(sciezka) as a:
        root = ET.fromstring(a.read("xl/styles.xml"))
    out = {}
    for nazwa in ("fonts", "fills", "borders", "cellXfs", "dxfs"):
        w = root.find(f"{NS}{nazwa}")
        out[nazwa] = len(list(w)) if w is not None else 0
    return out

# 1000 komorek z ta sama czcionka -> fonts == mala liczba (zwykle 2: domyslna + pogrubiona)

# ==================================================================
# 6. CYKL ZYCIA - cztery fazy, faza 3 pusta
# ==================================================================
# FAZA 1: wczytanie (lub utworzenie) modelu
wb = load_workbook("raport.xlsx")      # plik zamkniety po tej operacji!
# albo:
wb = Workbook()                        # model od zera

# FAZA 2: modyfikacje - TYLKO w RAM
ws = wb["Dane"]
ws["A1"] = "nowa wartosc"              # plik na dysku NIETKNIETY

# (FAZA 3: NIE ISTNIEJE)

# FAZA 4: zapis - przepisanie CALEGO pakietu od nowa
wb.save("raport_wynikowy.xlsx")        # NADPISUJE bez ostrzezenia!

# ==================================================================
# 7. DWA SWIATY - model != plik
# ==================================================================
a = load_workbook("raport.xlsx")
b = load_workbook("raport.xlsx")
a is b                                 # False! Dwa niezalezne grafy

a["Dane"]["A1"] = "zmiana z A"
b["Dane"]["A1"]                        # nadal stara wartosc - b nic nie wie

# Model jest samodzielny - można usunąć plik, model działa dalej:
import os
os.remove("raport.xlsx")
a["Dane"]["A1"]                        # dziala - model ma dane w RAM
a.save("raport.xlsx")                  # odtworzenie pliku z modelu

# ==================================================================
# 8. TRZY TRYBY PRACY
# ==================================================================
# --- TRYB NORMALNY: swobodny dostep losowy, caly graf w RAM ---
wb = load_workbook("plik.xlsx")
ws = wb["Dane"]
ws["Z999"].value = 1
ws["A1"].value
ws["A1"].font

# --- READ_ONLY: strumien w przod, minimalna pamiec ---
wb = load_workbook("plik.xlsx", read_only=True)
ws = wb["Dane"]
try:
    for wiersz in ws.iter_rows(values_only=True):
        przetworz(wiersz)
finally:
    wb.close()                         # OBOWIAZKOWE
# ws["A1"]        -> nie zadziala
# ws.cell(...)    -> nie zadziala
# ws.max_row      -> moze byc None

# --- WRITE_ONLY: prasa drukarska, tylko append ---
from openpyxl.cell import WriteOnlyCell
wb = Workbook(write_only=True)
ws = wb.create_sheet("Log")
komorka = WriteOnlyCell(ws, value="WAZNE")
komorka.font = Font(bold=True)         # styl PRZED append!
ws.append([komorka, "tekst", 42])      # jedyny sposob zapisu
wb.save("wynik.xlsx")

# ==================================================================
# 9. TABELA PRAWDY - skrot (pelna wersja w sekcji 2.7)
# ==================================================================
# POTRAFI:              wartosci, daty, style, number_format, wymiary,
#                       scalenia, freeze_panes, wydruk, tabele,
#                       formatowanie warunkowe (podstawowe), walidacja
#                       (podstawowa), ochrona, obrazy, hiperlinki, komentarze
#
# POTRAFI CZESCIOWO:    wykresy (nie wszystkie typy; odbudowa z modelu),
#                       warianty x14 regul i walidacji, makra (keep_vba,
#                       bez edycji), tabele przestawne (odczyt+zapis),
#                       odczyt i zapis struktury makr i formantow
#
# GUBI PRZY ZAPISIE:    sparkline'y, slicery, osie czasu, ksztalty,
#                       pola tekstowe, formanty osadzone, Power Query,
#                       Power Pivot, model danych, customXml
#
# NIE OBSLUGUJE:        .xls (BIFF), .xlsb, pliki szyfrowane haslem
#
# NIE POTRAFI (nigdy):  obliczac formul - nie ma silnika przeliczania

# ==================================================================
# 10. TRZY PYTANIA PRZED KAZDA OPERACJA
# ==================================================================
# 1. Czy to zmiana w MODELU, czy w PLIKU?        (nie przenikaja sie)
# 2. Czy element jest w modelu openpyxl?          (sprawdz tabele prawdy)
# 3. Czy nie mieszam dwoch systemow indeksowania? (1 vs 0)
```

## 9. Co dalej

Masz teraz model mentalny, na którym opiera się cały kurs. Trzy rzeczy powinny być dla Ciebie oczywiste i powinny pozostać oczywiste do końca:

**openpyxl nie edytuje plików — konwertuje je przez model.** Wszystko, czego model nie obejmuje, ginie w momencie zapisu, a nie w momencie modyfikacji. Zapamiętaj to zdanie, bo wyjaśnia ono moduł 18 w całości.

**Style są indeksami do współdzielonego rejestru, a nie własnością komórki.** Stąd bierze się niemutowalność, stąd wzorzec „kserokopia + modyfikacja + przypisanie", i stąd ostrożność przy zmianach, które mogą niechcący dotknąć tysiące komórek.

**Model i plik to dwa niezależne światy.** Bez `save()` nic nie trafia na dysk. Z `save()` trafia *wszystko*, co model zna — i nadpisuje to, co było. To jednocześnie wygodne i niebezpieczne.

W **module 04** napiszesz pierwszy kompletny program: utworzenie skoroszytu, zapis, wczytanie, praca ze ścieżkami, arkusze i ich nazwy, właściwości dokumentu. Będzie to pierwszy moment, w którym zobaczysz pełny cykl od `Workbook()` do pliku otwartego w Excelu — i pierwszy, w którym użyjesz świadomie wszystkich trzech rzeczy, które właśnie poznałeś.

Przygotuj do modułu 04:

- **katalog `output/`** z plikami z tego modułu — będziesz na nich testować,
- **jeden plik utworzony ręcznie w Excelu** — przy okazji sprawdzisz, jak openpyxl radzi sobie z cudzym plikiem (i czy przypadkiem nie ma w nim nic z listy „gubi przy zapisie"),
- **notatkę** z odpowiedzią na pytanie: „w którym miejscu mojego skryptu zmieniam model, a w którym dotykam pliku?" — dla trzech pierwszych skryptów, jakie napisałeś w tym kursie.

W module 04 pierwszy raz zobaczysz też, jak **łatwo** jest przez pomyłkę nadpisać oryginał i jak **prosto** temu zapobiec — bo teraz wiesz, że `save()` nie pyta o pozwolenie.