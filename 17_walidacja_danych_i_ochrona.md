Zanim napiszę moduł, zweryfikuję kluczowe API: sygnaturę `DataValidation`, zachowanie `ws.data_validations`, listę atrybutów `SheetProtection` i `Workbook.security` oraz znane ograniczenia (limit 255 znaków, `read_only`).

Teraz zweryfikuję znane pułapki: limit znaków, prefiks `=`, odwróconą semantykę `showDropDown`, zachowanie w `read_only` oraz `Protection(locked=False)`.

Świetny materiał. Teraz zweryfikuję trzy subtelne rzeczy: dokładne brzmienie ostrzeżenia o rozszerzeniu `x14`, zachowanie deskryptora hasła (`set_password` → `enable()`) oraz konwencję `formula1` (prefiks `=`, cudzysłowy).

# Moduł 17 — Walidacja danych i ochrona: bramkarz i kłódka na szufladzie

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–16

## 0. W tym module nauczysz się

- **Zbudujesz listy rozwijane** (`DataValidation`) — od prostej listy wpisanej w kodzie po listę pobieraną z innego arkusza, wraz z komunikatami wejścia i błędów w języku polskim.
- **Poznasz wszystkie siedem typów walidacji** (`list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom`) i zrozumiesz, że **walidacja to warunek, nie obliczenie** — openpyxl zapisuje regułę, ale jej nie sprawdza ani nie ocenia.
- **Zrozumiesz, dlaczego walidacja to tylko UX, a nie bezpieczeństwo** — i dlaczego każdy skrypt czytający Twój arkusz (pandas, openpyxl, baza) przejdzie obok listy rozwijanej jak obok otwartej furtki.
- **Opanujesz ochronę arkusza** — 15 niezależnych uprawnień, mechanizm „odblokuj pola, potem zamknij resztę" i ukrywanie formuł przed wzrokiem użytkownika.
- **Ochronisz strukturę skoroszytu** (`wb.security.lockStructure`) i dowiesz się, dlaczego openpyxl **nie egzekwuje** żadnej ochrony przy czytaniu.
- **Zobaczysz wprost, że hasło w OOXML to nie szyfrowanie** — 16-bitowy hasz, 65 536 możliwości i złamanie w kilka sekund.
- **Poznasz najpoważniejsze ograniczenie tego obszaru:** walidacje zapisane przez Excela w rozszerzeniu `x14` (typowo: listy z innego arkusza) **są gubione przy zapisie przez openpyxl**, z ostrzeżeniem, które można przeoczyć.

## 1. Intuicja i analogia

### 1.1. Bramkarz, który stoi tylko przy jednym wejściu

Wyobraź sobie klub, do którego prowadzą **dwa wejścia**:

- **Wejście główne z bramkarzem.** Bramkarz ma listę gości. Sprawdza każdego wchodzącego, a jeśli ktoś nie jest na liście, **nie wpuszcza go** — albo tylko upomina, albo zatrzymuje, zależnie od tego, jak surowy ma rozkaz.
- **Wejście od kuchni, bez ochrony.** Stoi otwarte. Nikt go nie pilnuje.

Excel ma dokładnie tak samo. **Lista rozwijana to bramkarz na wejściu głównym.** Działa tylko wtedy, gdy człowiek **wpisuje** wartość w komórkę przez interfejs Excela. Ale w pliku **nie ma żadnej blokady danych** — jest tylko **instrukcja dla programu Excel**, żeby sprawdzał dane wejściowe.

Konsekwencja, którą musisz zrozumieć zanim napiszesz pierwszą linię kodu:

```python
# Excel pokaze liste rozwijana z "Tak/Nie". Ale:
wb = load_workbook("formularz.xlsx")
ws = wb["Formularz"]
ws["B2"] = "MOZE BYC COKOLWIEK"      # openpyxl NIE sprawdzi listy rozwijanej!
wb.save("zepsuty.xlsx")
```

Po zapisie w komórce `B2` siedzi `"MOZE BYC COKOLWIEK"`. Excel otworzy plik, **pokoloruje tę komórkę na czerwono** (bo wartość łamie regułę) — ale **nie cofnie jej** i nie zgłosi błędu. Bramkarz nie zauważył, bo nie stał przy kuchennych drzwiach.

To nie jest usterka openpyxl. To jest usterka **Twojego myślenia**, jeśli sądziłeś, że walidacja to kontrola integralności danych. **Walidacja w Excelu to wygoda dla człowieka, nie gwarancja dla programu.** Zapamiętaj to zdanie — wracają do niego moduły 20, 21 i 26.

### 1.2. Walidacja to bilet, nie kłódka

Jeszcze jedno rozróżnienie, bo od niego zależy cała druga połowa modułu.

**Walidacja** mówi: *„ta wartość jest podejrzana"*. **Ochrona** mówi: *„nie możesz tu nic zmienić"*.

Zobacz różnicę na dwóch zdaniach:

- *„Wpisz datę w formacie RRRR-MM-DD"* — to **walidacja**. Użytkownik może wpisać `13/13/2026`, a Excel tylko się oburzy.
- *„To pole jest zablokowane"* — to **ochrona**. Użytkownik **nie może** wpisać niczego, dopóki nie zdejmie ochrony arkusza.

Analogia z formularzem urzędowym: **formularz jest zadrukowany i sztywny** (ochrona), a **w białych kratkach mieści się tylko to, co urząd przewidział** (walidacja). Zadrukowany formularz chroni przed dopisaniem własnych rubryk, ale w białej kratce nadal można napisać bzdurę.

### 1.3. Białe pola i zamalowany wzór

Dwa mechanizmy, które spotkasz w praktyce i które trzeba zrozumieć mechanicznie:

**Białe pola formularza.** W Excelu **każda komórka jest domyślnie zablokowana**. Ta blokada leży i nic nie robi, dopóki nie włączysz ochrony arkusza. Wtedy nagle **wszystko** jest martwe. To największe zaskoczenie w tym module.

Dlatego wzorzec nie brzmi „zablokuj formuły". Brzmi: **„odblokuj pola do wpisywania, potem włącz ochronę"**. Analogia: **nie zaznaczasz na formularzu, które rubryki są urzędowe — zaznaczasz białe kratki dla petenta.** Reszta jest urzędowa z definicji.

**Zamalowany wzór w zeszycie ćwiczeń.** Gdy dziecko ma policzyć `2 + 3 = ...` i wpisać `5`, w zeszycie **nauczyciela** widać też cały rachunek. Żeby uczeń nie podglądał rozwiązania, nauczyciel **zamalowuje wzór**. Uczeń widzi `5`, nie widzi skąd się wzięło.

Dokładnie to robi `Protection(hidden=True)`: użytkownik widzi **wynik** formuły, ale w pasku formuły zamiast `=SUMA(A1:A10)` zobaczy... nic konkretnego. Ale uwaga — **zamalowanie działa tylko wtedy, gdy arkusz jest chroniony.** Bez ochrony wzór jest jawny, w każdej kopii i w każdym pliku, który ktoś zapisze. To szczegół, który za chwilę zobaczysz w kodzie.

### 1.4. Aneks w nieznanym alfabecie

Tu jest najpoważniejsza praktyczna pułapka modułu i musisz ją zrozumieć **teraz**, bo wróci w modułach 18 i 26.

Wyobraź sobie umowę, do której dołączono **aneks spisany alfabetem, którego Twoja kancelaria nie zna**. Prawnik umie przeczytać umowę główną, ale aneksu nie. Co robi, gdy przepisuje umowę na czysto? **Zostawia aneks na boku** — nie przepisze tego, czego nie rozumie.

W OOXML jest tak samo, w dwóch warstwach:

- **Walidacje z 2006 roku** (podstawowy standard OOXML) — openpyxl czyta i zapisuje. Lista wpisana wprost, liczba w zakresie, data, długość tekstu, formuła własna.
- **Walidacje z 2009 roku** (`x14:dataValidation` w `extLst`) — **rozszerzenie Microsoftu**. Excel zapisuje tam **listy rozwijane wskazujące na inny arkusz**, dla zgodności ze starym Excelem 2007. openpyxl tego **nie rozumie**: wypisuje ostrzeżenie i **gubi przy zapisie**.

Ostrzeżenie, które zobaczysz, brzmi dokładnie tak:

```text
UserWarning: Data Validation extension is not supported and will be removed
```

Modyfikator openpyxl powiedział o tym wprost na liście dyskusyjnej: *„There is currently no way around this."* Nie ma magicznego parametru. Są tylko **obejścia** — dwa, i oba pokażę w sekcji 2.8.

Wniosek praktyczny: **jeśli plik przyszedł z Excela i ma listy rozwijane z innego arkusza, to przepuszczenie go przez openpyxl w celu zapisu niemal na pewno zniszczy te listy.** Wykryjesz to po ostrzeżeniu — ale musisz je **przechwycić**, a nie pozwolić mu zniknąć w konsoli.

### 1.5. Trzy kategorie prawdy dla tego modułu

Zgodnie z obietnicą z modułu 03:

| Element | openpyxl **potrafi** | **potrafi częściowo** | **gubi przy zapisie** |
|---|---|---|---|
| **Walidacja (typ 2006)** | utworzyć, odczytać, zapisać, komunikaty | — | — |
| **Walidacja z listą z innego arkusza** | utworzyć (`formula1="Arkusz!$A$1:$A$10"`) | odczytać własną | **cudzą z `x14` — tak, całkowicie** |
| **Limit 255 znaków listy** | zapisze, nie ostrzeże | — | Excel odrzuci/naprawi wartość |
| **Ochrona arkusza** | włączyć, hasło, 15 uprawnień | — | — (solidne) |
| **Ochrona skoroszytu** | `lockStructure`, hasła | — | — (solidne) |
| **„Wymuszenie" ochrony** | **nie istnieje** | — | każdy skrypt ją ignoruje |
| **Prawdziwe bezpieczeństwo** | **nie istnieje** | — | hasło łamane w sekundę (2.9) |

Zwróć uwagę na dwie ostatnie linie. **Ochrona jest solidna jako mechanizm, ale bezużyteczna jako bezpieczeństwo.** To nie sprzeczność — to precyzyjny opis tego, co openpyxl oferuje.

## 2. Teoria

### 2.1. Walidacja: pełna sygnatura i jej siedem twarzy

Zacznijmy od dokładnego źródła — 3.1.5, klasa `DataValidation`:

```python
# openpyxl/worksheet/datavalidation.py
class DataValidation(Serialisable):
    tagname = "dataValidation"

    sqref = Convertible(expected_type=MultiCellRange)
    cells = Alias("sqref")            # ← aliasy z openpyxl 2.x, dla zgodności
    ranges = Alias("sqref")

    showDropDown = Bool(allow_none=True)
    hide_drop_down = Alias('showDropDown')      # ← NAZWA MNIEJ ZWIODLIWA!
    showInputMessage = Bool(allow_none=True)
    showErrorMessage = Bool(allow_none=True)
    allowBlank = Bool(allow_none=True)
    allow_blank = Alias('allowBlank')

    errorTitle = String(allow_none=True)
    error = String(allow_none=True)
    promptTitle = String(allow_none=True)
    prompt = String(allow_none=True)
    formula1 = NestedText(allow_none=True, expected_type=str)
    formula2 = NestedText(allow_none=True, expected_type=str)

    type = NoneSet(values=("whole", "decimal", "list", "date", "time",
                           "textLength", "custom"))
    errorStyle = NoneSet(values=("stop", "warning", "information"))
    imeMode = NoneSet(values=(...))
    operator = NoneSet(values=("between", "notBetween", "equal", "notEqual",
                               "lessThan", "lessThanOrEqual",
                               "greaterThan", "greaterThanOrEqual"))
    validation_type = Alias('type')

    def __init__(self, type=None, formula1=None, formula2=None,
                 showErrorMessage=False, showInputMessage=False,
                 showDropDown=False, allowBlank=False, sqref=(),
                 promptTitle=None, errorStyle=None, error=None, prompt=None,
                 errorTitle=None, imeMode=None, operator=None,
                 allow_blank=None):
        ...
        if allow_blank is not None:
            allowBlank = allow_blank        # ← allow_blank WYGRYWA z allowBlank
        ...
```

Cztery rzeczy, które trzeba z tego wyczytać — wszystkie mają konsekwencje w praktyce:

**1. `sqref` to nie string, a `MultiCellRange`.** Czyli **zbiór zakresów**, nie jeden zakres. Możesz mieć `"A1 B2:B5 C7"` — walidacja obejmuje trzy rozłączne obszary. `dv.add(...)` **dokłada** do tego zbioru, nie zastępuje go. Dlatego `dv.add("B2:B100")` po `dv.add("D2")` daje **dwa** zakresy w jednej regule.

**2. `cells` i `ranges` to aliasy `sqref`.** W openpyxl 2.x `cells` było zbiorem współrzędnych (można było robić `dv.cells.add("B2")`). W 3.1 `dv.cells` **zwraca ten sam `MultiCellRange`**. Stare tutoriale z `dv.cells.add(...)` Cię tu nie zaprowadzą — to punkt migracyjny.

**3. `validation_type` to alias `type`.** Stare `DataValidation(validation_type="list")` **nadal działa**, ale `type` jest poprawną nazwą. (W bardzo starych wersjach wywoływało to `DeprecationWarning`.)

**4. `hide_drop_down` to alias `showDropDown`** — i to jest **pułapka numer jeden** w tym module. Wrócę do niej w 2.4.

Siedem typów walidacji i kiedy ich używać:

| `type` | Co sprawdza | `operator` sensowny? | `formula1` / `formula2` |
|---|---|---|---|
| `list` | wartość musi być jedną z listy | nie | lista w cudzysłowach **albo** zakres **albo** nazwa zdefiniowana |
| `whole` | liczba całkowita | tak | `formula1` (i `formula2` dla `between`) |
| `decimal` | liczba rzeczywista | tak | jak wyżej |
| `date` | data | tak | wzorzec jako data/serial Excela |
| `time` | czas | tak | jak wyżej |
| `textLength` | długość tekstu | tak | `formula1=15` (np. „maksymalnie 15 znaków") |
| `custom` | dowolna formuła Excela, która ma dać `TRUE`/`FALSE` | nie | `formula1="=ISNUMBER(A1)*..."` |

Warto zapamiętać jedną rzecz o `custom`: to **jedyne naprawdę elastyczne narzędzie** — ale i **jedyne, które openpyxl nie sprawdzi nawet w Excelu**, bo formula jest oceniana przez Excela przy wpisywaniu. W praktyce `custom` to najbliższy krewny reguł formatowania warunkowego z modułu 13 — tyle że działa na wejściu, a nie na wyświetlaniu.

### 2.2. Trzy kroki, które zawsze występują razem (i brak któregokolwiek jest cichy)

To najważniejsza procedura w pierwszej połowie modułu:

```python
# KROK 1: zbuduj regule
dv = DataValidation(type="list", formula1='"Tak,Nie"', allow_blank=True)

# KROK 2: ZAREJESTRUJ regule na arkuszu
ws.add_data_validation(dv)

# KROK 3: przypisz zakresy
dv.add("B2:B100")
```

Źródło `Worksheet.add_data_validation` to jedna linijka:

```python
def add_data_validation(self, data_validation):
    """Add a data-validation object to the sheet."""
    self.data_validations.append(data_validation)
```

I tu jest cała przyczyna najczęstszego błędu początkujących. **Pominięcie kroku 2 nie zgłasza błędu** — `dv.add("B2:B100")` zadziała, reguła będzie miała poprawny `sqref`, *„tylko nigdy nie trafi do arkusza"*. W pliku nie ma walidacji, a w kodzie nie ma wyjątku.

A co z pominięciem kroku 3? Też cicho — i to **udokumentowane wprost**:

> **Note** — *„Validations without any cell ranges will be ignored when saving a workbook."*

Potwierdza to źródło `DataValidationList.to_tree`:

```python
def to_tree(self, tagname=None):
    """
    Need to skip validations that have no cell ranges
    """
    ranges = self.dataValidation  # copy
    self.dataValidation = [r for r in self.dataValidation if bool(r.sqref)]
    xml = super(DataValidationList, self).to_tree(tagname)
    self.dataValidation = ranges
    return xml
```

Zwróć uwagę na filtr: `if bool(r.sqref)`. **Reguła bez ani jednego zakresu jest wyrzucana przy serializacji.** Czyli możesz mieć idealnie zbudowaną walidację w pamięci i **pusty efekt w pliku**.

Analogia: **bramkarz z listą gości w kieszeni, któremu nie powiedzieli, przy których drzwiach ma stanąć**. Lista jest, bramkarz jest — nic z tego nie wynika.

### 2.3. `formula1`: trzy postacie, których nie wolno mylić

To źródło największej liczby niejasności w sieci. Rozróżnienie jest proste, gdy raz je zrozumiesz.

**Postać A — lista wpisana wprost.** Wartości oddzielone **przecinkiem**, a **całość w cudzysłowach**, które są **częścią wartości**, nie skladni Pythona:

```python
dv = DataValidation(type="list", formula1='"Tak,Nie,Nie wiem"', allow_blank=True)
#                                          ^                        ^
#                                          te cudzysloby sa W SRODKU napisu!
```

W Pythonie możesz to też zapisać czytelniej, budując napis dynamicznie:

```python
opcje = ["Tak", "Nie", "Nie wiem"]
dv = DataValidation(type="list", formula1='"' + ",".join(opcje) + '"', allow_blank=True)
```

**Postać B — zakres komórek.** **Bez cudzysłowów, bez znaku `=`.** Dokumentacja openpyxl podaje dokładnie ten wzorzec:

```python
from openpyxl.utils import quote_sheetname

dv = DataValidation(type="list",
                    formula1="{0}!$B$1:$B$10".format(quote_sheetname(sheetname)))
```

`quote_sheetname` to znajomy pomocnik z modułów 07 i 16 — wstawia apostrofy wokół nazwy arkusza, gdy zawiera ona spację. Bez tego link do `Dane 2026!$B$1:$B$10` będzie **niepoprawny**, a Excel otworzy plik z błędem reguły.

> **⚠️ O znaku `=`.** Dokumentacja openpyxl podaje zakresy **bez** `=`. Spotkasz też w sieci wersje z `=` (np. `formula1="=$A:$A"`) i one bywają u ludzi działające — Excel toleruje nadmiarowy `=` w tym miejscu. **Trzymaj się formy z dokumentacji** (bez `=` dla zakresu), a `=` dodawaj tylko w typie `custom`, gdzie dokumentacja pokazuje `formula1="=SOMEFORMULA"`. Jeśli nie jesteś pewien — sprawdź w swojej wersji i obejrzyj plik w Excelu.

**Postać C — nazwa zdefiniowana.** Najbardziej odporna i najczęściej pomijana:

```python
from openpyxl.workbook.defined_name import DefinedName

wb.defined_names.add(DefinedName("Regiony", attr_text="'Słownik'!$A$2:$A$10"))
dv = DataValidation(type="list", formula1="Regiony", allow_blank=True)
```

Dlaczego to jest lepsze? Bo gdy lista się przesunie albo wydłuży, **nie musisz ruszać reguł** — poprawiasz nazwę w jednym miejscu. To dokładnie ta sama zasada co `NamedStyle` (moduł 09) i `FORMAT["waluta_pln"]` (moduł 10): **jedno źródło decyzji**. I — co ważniejsze dla tego modułu — **walidacja oparta na nazwie zdefiniowanej jest tą, którą openpyxl potrafi zachować przy modyfikacji cudzego pliku** (2.8).

### 2.4. `showDropDown` — najbardziej zdradliwy parametr w całym module

Popatrz na to jeszcze raz:

```python
showDropDown = Bool(allow_none=True)
hide_drop_down = Alias('showDropDown')      # ← ALIAS NAZYWA SIE "HIDE"!
```

Nazwa `showDropDown` sugeruje: „pokaż listę rozwijaną". **Jest dokładnie odwrotnie.** Dokumentacja openpyxl mówi to wprost:

> **Note** — *„Excel and LibreOffice interpret the parameter `showDropDown=True` as the dropdown arrow should be hidden."*

Co się dzieje od strony pliku? Atrybut w XML nosi nazwę `showDropDown`, ale w specyfikacji OOXML ma **odwrotne znaczenie** niż brzmi: `showDropDown="1"` oznacza *„ukryj strzałkę listy"*. openpyxl przekazuje wartość **bez zmiany**, więc jego parametr dziedziczy to przewrócone znaczenie. Stąd alias `hide_drop_down` — który jest **uczciwą nazwą tego samego pola**.

**Praktyczna reguła:**

```python
# ✅ DOBRZE - domyslna wartosc showDropDown=False, strzalka WIDOCZNA
dv = DataValidation(type="list", formula1='"Tak,Nie"', allow_blank=True)

# ❌ ZLE - showDropDown=True UKRYWA strzalke (lista "nie dziala")
dv = DataValidation(type="list", formula1='"Tak,Nie"', showDropDown=True)
```

Konsekwencja: użytkownik widzi komórkę z walidacją, ale **bez strzałki listy rozwijanej**. Wartość nadal musi być jedną z listy, ale żeby ją wybrać, trzeba trafić w nią z palca. Wygląda jak „walidacja nie działa", a naprawdę działa *za dobrze*.

To najczęstsza przyczyna zgłoszeń „`DataValidation` nie pokazuje drop-downu" — obok pominięcia `ws.add_data_validation(dv)`. **Nie ustawiaj tego parametru, jeśli nie wiesz, po co.**

### 2.5. Limit 255 znaków: cicha katastrofa

Excel ma twardy limit na **długość napisu listy wpisanej wprost**: **255 znaków łącznie z przecinkami**.

Kluczowe: to limit **Excela**, nie openpyxl. XlsxWriter w tym miejscu **ostrzega**:

```text
UserWarning: Length of list items exceeds Excel's limit of 255, use a formula range instead
```

**openpyxl tego nie robi.** Zapisze plik z za długą listą **bez jednego słowa ostrzeżenia**. Excel przy otwarciu zrobi z tym co potrafi — w najlepszym razie pokaże uciętą listę, w gorszym zgłosi problem z plikiem.

Analogia: **wykładowca, który przyjmuje pracę dłuższą niż limit, bo nie sprawdza objętości** — a potem komisja egzaminacyjna odrzuca pracę. Wina nie leży po stronie studenta, ale konsekwencje zbiera on.

Praktyczny test — policz sam:

```python
def dlugosc_listy(opcje: list[str]) -> int:
    """Dlugosc napisu listy tak, jak liczy ja Excel: wartosci + przecinki."""
    return len(",".join(opcje))

opcje_ok = ["Tak", "Nie", "Nie wiem"]
print(dlugosc_listy(opcje_ok))          # 17 - bezpiecznie

opcje_dlugie = [f"Region {i:02d}" for i in range(1, 40)]
print(dlugosc_listy(opcje_dlugie))      # 399 - PRZEKROCZONY LIMIT!
```

**Reguła projektowa:** jeśli lista ma **więcej niż ~15 krótkich pozycji** albo **nie jesteś pewien jej długości** — **zawsze** używaj zakresu komórek albo nazwy zdefiniowanej. Nie licz na to, że „jakoś się zmieści".

I jeszcze jedna, mniej znana konsekwencja: **długa lista wpisana wprost jest powielana w pliku.** Reguła siedzi w `sheet1.xml`, więc 200 pozycji × 20 reguł = 4 000 pozycji tekstu w jednym arkuszu XML. Zakres komórek zajmuje jedną linijkę. To ten sam argument, co „link do pliku zamiast kopii bajtów" z modułu 16.

### 2.6. Ochrona arkusza: 15 przełączników z odwróconą logiką

`SheetProtection` w 3.1.5, z pełnymi wartościami domyślnymi:

```python
# openpyxl/worksheet/protection.py
class SheetProtection(Serialisable, _Protected):
    """
    Information about protection of various aspects of a sheet. True values
    mean that protection for the object or action is active This is the
    **default** when protection is active, ie. users cannot do something
    """
    tagname = "sheetProtection"

    sheet = Bool()
    enabled = Alias('sheet')
    objects = Bool()
    scenarios = Bool()
    formatCells = Bool()
    formatColumns = Bool()
    formatRows = Bool()
    insertColumns = Bool()
    insertRows = Bool()
    insertHyperlinks = Bool()
    deleteColumns = Bool()
    deleteRows = Bool()
    selectLockedCells = Bool()
    selectUnlockedCells = Bool()
    sort = Bool()
    autoFilter = Bool()
    pivotTables = Bool()
    ...

    def __init__(self, sheet=False, objects=False, scenarios=False,
                 formatCells=True, formatRows=True, formatColumns=True,
                 insertColumns=True, insertRows=True, insertHyperlinks=True,
                 deleteColumns=True, deleteRows=True, selectLockedCells=False,
                 selectUnlockedCells=False, sort=True, autoFilter=True,
                 pivotTables=True, password=None, ...):
```

**Kluczowe zdanie z docstringa: „`True` values mean that protection for the object or action is active, ie. users cannot do something."** To znaczy: **każda flaga nazywa to, co jest ZABLOKOWANE.** `True` = zabronione, `False` = dozwolone.

To odwrócenie jest źródłem połowy pomyłek, bo w interfejsie Excela widzisz listę **checkboxów z uprawnieniami** („Allow all users of this worksheet to: Use AutoFilter"), a nie z blokadami. Odwzorowanie:

| Flaga openpyxl | Domyślnie | `True` znaczy | Odpowiednik w UI Excela |
|---|---|---|---|
| `sheet` | `False` | ochrona włączona (główny wyłącznik) | przycisk „Protect Sheet" wciśnięty |
| `objects` | `False` | **nie można** edytować obiektów (wykresy, obrazy, kształty) | checkbox **„Edit objects" ODRZUCONY** |
| `scenarios` | `False` | **nie można** edytować scenariuszy | checkbox **„Edit scenarios" ODRZUCONY** |
| `formatCells` | `True` | **nie można** zmieniać formatowania komórek | „Format cells" ODRZUCONY |
| `formatColumns` | `True` | **nie można** zmieniać szerokości/ukrywania kolumn | „Format columns" ODRZUCONY |
| `formatRows` | `True` | **nie można** zmieniać wysokości/ukrywania wierszy | „Format rows" ODRZUCONY |
| `insertColumns` | `True` | **nie można** wstawiać kolumn | „Insert columns" ODRZUCONY |
| `insertRows` | `True` | **nie można** wstawiać wierszy | „Insert rows" ODRZUCONY |
| `insertHyperlinks` | `True` | **nie można** wstawiać hiperlinków | „Insert hyperlinks" ODRZUCONY |
| `deleteColumns` | `True` | **nie można** usuwać kolumn | „Delete columns" ODRZUCONY |
| `deleteRows` | `True` | **nie można** usuwać wierszy | „Delete rows" ODRZUCONY |
| `selectLockedCells` | `False` | **nie można** zaznaczać zablokowanych komórek | „Select locked cells" ODRZUCONY |
| `selectUnlockedCells` | `False` | **nie można** zaznaczać odblokowanych komórek | „Select unlocked cells" ODRZUCONY |
| `sort` | `True` | **nie można** sortować | „Sort" ODRZUCONY |
| `autoFilter` | `True` | **nie można** używać strzałek filtra | „Use AutoFilter" ODRZUCONY |

**Uwaga na dwa ostatnie.** Dokumentacja Microsoftu dodaje warunki, których **nie da się obejść flagami**:

> *„Users can't sort ranges that contain locked cells on a protected worksheet, regardless of this setting."*
>
> *„Users cannot apply or remove AutoFilters on a protected worksheet, regardless of this setting."*

Czyli: `ws.protection.sort = False` **daje uprawnienie**, ale jeśli w zakresie są **zablokowane komórki** — sortowanie i tak nie zadziała. A filtra **nie da się założyć ani zdjąć** na chronionym arkuszu; można tylko **używać istniejącej strzałki**. Wniosek praktyczny: **filtr i sortowanie zaplanuj przed włączeniem ochrony** (a filtr najlepiej jako Tabelę — moduł 12), a zakres danych zostaw odblokowany, jeśli ma być sortowalny.

### 2.7. Trzy sposoby włączenia ochrony — i jeden haczyk

```python
# SPOSOB 1: flaga
ws.protection.sheet = True

# SPOSOB 2: metoda
ws.protection.enable()          # identyczne: ustawia sheet = True
ws.protection.disable()         # ustawia sheet = False

# SPOSOB 3: haslo
ws.protection.password = "tajne"
```

**Haczyk jest w sposobie 3.** Popatrz na źródło `_Protected` i `SheetProtection`:

```python
class _Protected(object):
    _password = None

    def set_password(self, value='', already_hashed=False):
        """Set a password on this sheet."""
        if not already_hashed:
            value = hash_password(value)
        self._password = value

    @property
    def password(self):
        return self._password

    @password.setter
    def password(self, value):
        self.set_password(value)        # ← property -> set_password

class SheetProtection(Serialisable, _Protected):
    def set_password(self, value='', already_hashed=False):
        super(SheetProtection, self).set_password(value, already_hashed)
        self.enable()                   # ← !!! USTAWIENIE HASLA WŁĄCZA OCHRONĘ
```

**Ustawienie `ws.protection.password` automatycznie włącza ochronę arkusza.** Jeśli myślisz, że „najpierw ustawię hasło, a potem zdecyduję, czy włączyć" — nie masz tej opcji. Hasło i włączenie są jedną operacją.

To samo dotyczy konstruktora:

```python
from openpyxl.worksheet.protection import SheetProtection

# to TEZ wlacza ochrone (bo __init__ robi self.password = password -> set_password -> enable)
ws.protection = SheetProtection(password="tajne")
print(ws.protection.sheet)      # True
```

I jeszcze jedna konsekwencja, którą trzeba znać: `bool(ws.protection)` zwraca `ws.protection.sheet`, więc możesz w kodzie pisać `if ws.protection:` i to działa zgodnie z intuicją (arkusz jest chroniony → `True`).

### 2.8. `locked` i `hidden` — ochrona, która MIESZKA W STYLU

Wróć pamięcią do modułu 02: komórka nie „ma" formatowania, ma **numer stylu**, a wygląd siedzi w `xl/styles.xml`. Ta sama zasada obowiązuje dla stanu blokady.

`Protection` to **szósty składnik stylu komórki**:

```python
from openpyxl.styles import Protection
from copy import copy

# Komorka domyslnie: locked=True, hidden=False
cell.protection = Protection(locked=False)     # odblokuj do wpisywania
cell.protection = Protection(hidden=True)      # ukryj formule w pasku formuly
cell.protection = Protection(locked=False, hidden=False)
```

Dwie konsekwencje, obie z wcześniejszych modułów:

**1. „Niemutowalność" z modułu 09 dotyczy także tego.** To, co w module 09 nazwaliśmy „pieczątką", dla ochrony znaczy dokładnie to samo: `cell.protection.locked = False` **nie zadziała**. Trzeba **przypisać nowy obiekt**. A jeśli chcesz zmienić **tylko** stan blokady, nie gubiąc stanu `hidden`:

```python
nowa_ochrona = copy(cell.protection)      # kserokopia pieczatki
nowa_ochrona.locked = False                # dopisujemy jedna zmiane
cell.protection = nowa_ochrona             # przybijamy nowa pieczatke
```

**2. Stan blokady trafia do `styles.xml`, nie do komórki.** To znaczy, że **zablokowanie/odblokowanie komórek tworzy nowe wpisy w `cellXfs`** — dokładnie tak samo, jak zmiana czcionki. W pliku zobaczysz to tak:

```xml
<!-- xl/styles.xml -->
<cellXfs count="3">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0">
    <protection locked="0"/>          <!-- ← odblokowana komorka -->
  </xf>
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0">
    <protection locked="1" hidden="1"/>   <!-- ← zablokowana + ukryta formuła -->
  </xf>
</cellXfs>
```

A w samym arkuszu `sheet1.xml` pojawia się osobny element:

```xml
<sheetProtection sheet="1" password="XXXX" objects="0" scenarios="0"
                 formatCells="1" formatRows="1" formatColumns="1"
                 insertColumns="1" insertRows="1" insertHyperlinks="1"
                 deleteColumns="1" deleteRows="1" selectLockedCells="0"
                 selectUnlockedCells="0" sort="1" autoFilter="1"
                 pivotTables="1"/>
```

**Zwróć uwagę na `password="XXXX"`.** Cztery znaki heksadecymalne. O tym za chwilę.

**O `hidden=True` — powiedzmy to wprost.** Ukrycie formuły **nie tworzy żadnego zabezpieczenia danych**. Formuła nadal siedzi w `sheet1.xml` w pełnej czytelnej postaci. Ukrycie działa **wyłącznie na pasku formuły w Excelu** i **wyłącznie wtedy, gdy arkusz jest chroniony**. Weź pod uwagę cytat, który znalazłem w dokumentacji: *„on a protected sheet it hides a cell's formula from the formula bar, so while protection is off — and in any unprotected copy you save along the way — those formulas are visible to anyone who opens the file."*

Czyli: **zamalowany wzór w zeszycie ćwiczeń.** Uczeń nie widzi rachunku, ale nauczyciel tak. A każdy, kto zrobi ksero zeszytu nauczyciela, też.

### 2.9. Uczciwa prawda o haśle: kłódka na szufladzie, nie sejf

Teraz rzecz, którą musisz powiedzieć swojemu klientowi zanim obieca „zabezpieczony raport".

Specyfikacja OOXML formułuje to lepiej, niż ja bym potrafił:

> *„Worksheet or workbook element protection should not be confused with file security. It is meant to make your workbook safe from unintentional modification, and cannot protect it from malicious modification."*

Po polsku: **ochrona arkusza czy skoroszytu nie jest bezpieczeństwem pliku. Ma chronić przed przypadkową zmianą, nie przed celowym działaniem.**

I teraz mechanika, dlaczego to nie jest tylko formułka:

- openpyxl używa **przestarzałego algorytmu Excela** (*Legacy Password Hash Algorithm*), chyba że jawnie skonfigurujesz inny.
- Wynik tego algorytmu to **16 bitów**, czyli **cztery znaki heksadecymalne** w pliku — a więc **65 536 możliwych wartości**. Dopasowanie hasła można znaleźć praktycznie natychmiast, a **wiele różnych haseł odblokuje ten sam plik**.
- Nowoczesny Excel, gdy hasło ustawia **człowiek**, zapisuje sól i skrót SHA-512. openpyxl **czyta i zachowuje** takie hasła, ale **zapisuje** postać przestarzałą.

I najważniejsze — **openpyxl nie egzekwuje ochrony przy czytaniu.** Możesz wziąć skoroszyt z `lockStructure=True` i **bez hasła** dodać arkusz:

```python
wb = load_workbook("raport.xlsx")          # struktura "zablokowana" haslem
wb.create_sheet("Archiwum")                 # openpyxl nie protestuje!
wb.save("raport_archived.xlsx")             # ...i zapisuje
```

Po ponownym wczytaniu `reloaded.security.lockStructure` to nadal `True` — **blokada jest na miejscu, ale została zignorowana przez Twój skrypt.** To samo dotyczy każdego innego narzędzia: `pandas.read_excel`, import do bazy, konwersja do CSV — **wszystkie widzą każdy arkusz i każdą komórkę, niezależnie od kłódki**.

Analogia z briefu jest tu najlepsza, jaka może być: **kłódka na szufladzie biurka, nie sejf.** Zatrzyma kolegę z biura, który chciał „tylko poprawić jedną liczbę". Nie zatrzyma nikogo, kto naprawdę chce dostać się do środka — a nawet nie zatrzyma Twojego własnego skryptu.

**Co robić, gdy dane są naprawdę poufne:**

1. **Szyfrowanie pliku** (hasło do otwarcia, nie do edycji) — osobny mechanizm OOXML, którego openpyxl nie realizuje; wymaga innego narzędzia lub Excela.
2. **Uprawnienia na poziomie systemu** — katalogi, udziały sieciowe, kontrola dostępu.
3. **Nie umieszczaj w pliku tego, czego nie chcesz ujawnić** — najlepsza „ochrona".

### 2.10. Ochrona skoroszytu: struktura, okna, zmiany śledzone

`wb.security` to obiekt `WorkbookProtection`:

```python
class WorkbookProtection(Serialisable):
    workbookPassword = ...
    workbookPasswordCharacterSet = String(allow_none=True)
    revisionsPassword = ...
    revisionsPasswordCharacterSet = String(allow_none=True)
    lockStructure = Bool(allow_none=True)
    lock_structure = Alias("lockStructure")
    lockWindows = Bool(allow_none=True)
    lock_windows = Alias("lockWindows")
    lockRevision = Bool(allow_none=True)
    lock_revision = Alias("lockRevision")
    ...
```

Trzy przełączniki i dwa hasła, każde do czego innego:

```python
# Struktura: nie mozna dodawac/usuwac/przenosic/ukrywac/zmieniac nazw arkuszy
wb.security.workbookPassword = "haslo-struktury"
wb.security.lockStructure = True

# Zmiany sledzone (legacy, wspoldzielone skoroszyty)
wb.security.revisionsPassword = "haslo-zmian"
wb.security.lockRevision = True
```

**Dokumentacja dodaje istotne zastrzeżenie:** *„Other properties … will only be enforced if the appropriate password is set."* Czyli `lockStructure = True` **bez** `workbookPassword` **nic nie robi** — nie ma czego wymagać od użytkownika. Analogia: **kartka „Nie ruszać" bez podpisu** — nikt się nią nie przejmie, bo nie ma z czym porównać uprawnień.

O dwóch pozostałych flagach:

- **`lockWindows`** — dotyczy rozmiaru i pozycji okien skoroszytu. Nowsze wersje Excela **nie oferują już tej opcji**, więc **zostaw ją nieustawioną**.
- **`lockRevision`** — dotyczy przestarzałego śledzenia zmian we wspólnych skoroszytach. W generowanych raportach **prawie nigdy nie potrzebne**.

W pliku trafi to do `xl/workbook.xml`:

```xml
<workbookProtection workbookPassword="XXXX" lockStructure="1"/>
```

### 2.11. Walidacja a tryby pracy — asymetria, którą trzeba znać

**`read_only=True` — walidacji NIE MA.** To nie jest luka; to konsekwencja architektury. Popatrz na czytnik:

```python
# openpyxl/reader/excel.py
if self.read_only:
    ws = ReadOnlyWorksheet(self.wb, sheet.name, rel.target, self.shared_strings)
    ws.sheet_state = sheet.state
    self.wb._sheets.append(ws)
    continue                       # ← KONIEC. Zaden parser arkusza nie zostal uruchomiony.
```

`continue` przerywa przetwarzanie arkusza **przed** jakimkolwiek parsowaniem — a więc **przed** walidacjami, formatowaniem warunkowym, tabelami, obrazami i komentarzami. `ReadOnlyWorksheet` to osobna, okrojona klasa.

**Konsekwencja praktyczna, którą trzeba zapamiętać razem z modułem 16:** jeżeli budujesz **narzędzie diagnostyczne** („co ten plik zawiera?"), **nigdy nie używaj `read_only`**. Otrzymasz fałszywy obraz: „plik nie ma walidacji", „plik nie ma obrazów", „plik nie ma reguł warunkowych". Trzy razy się mylisz, trzy razy w tym samym miejscu.

**`write_only=True` — obszar nieudokumentowany.** Formularze buduj w trybie normalnym. Nie opieraj na `write_only` niczego, co ma zawierać walidację lub ochronę; **sprawdź w swojej wersji**, zanim się na tym oprzesz (moduł 19 pokaże pełną tabelę ograniczeń trybu strumieniowego).

### 2.12. `x14` — rozszerzenie, które openpyxl czyta jako ostrzeżenie i wyrzuca

Wracamy do aneksu w nieznanym alfabecie z 1.4, teraz z konkretami.

Excel 2010 wprowadził nowe warianty walidacji — przede wszystkim **listę rozwijaną, której źródło leży na innym arkuszu**. Dla zgodności z Excelem 2007 zapisuje je **podwójnie**: w miejscu standardowym oraz w bloku rozszerzeń na końcu `sheet1.xml`:

```xml
<!-- koniec xl/worksheets/sheet1.xml -->
<extLst>
  <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"
       xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
    <x14:dataValidations count="1"
                         xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
      <x14:dataValidation type="list" allowBlank="1" showErrorMessage="1">
        <x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1>
        <xm:sqref>A2:A100</xm:sqref>
      </x14:dataValidation>
    </x14:dataValidations>
  </ext>
</extLst>
```

openpyxl rozumie **podstawowy** standard OOXML z 2006 roku (tam lista rozwijana **nie mogła** odwoływać się do innego arkusza). Blok `x14` jest dla niego właśnie tym aneksem — **czyta go jako ostrzeżenie i nie przenosi do nowego pliku**.

Rozmowa na liście openpyxl-users pokazuje, że granica jest ostra:

> **Pytanie:** *„I want read the file and update the content and retain the validation as well."*
> **Odpowiedź (Charlie Clark, opiekun openpyxl):** *„The warning is pretty clear: file contains a validation not covered by the original OOXML specification, which is why it will be lost from the file. There is currently no way around this."*

I uzupełnienie, które precyzuje, **gdzie dokładnie** jest problem:

> *„openpyxl does support data validation as covered by the original OOXML specification. However, since then Microsoft has extended the options for data validation and it these that are not supported."*

**Dwa obejścia.** Oba praktyczne, oba wymagają pracy:

**Obejście 1 — nazwy zdefiniowane (preferowane).** Zamiast odwołania bezpośredniego do zakresu na innym arkuszu, użyj **nazwy zdefiniowanej** i wskaż ją w regule:

```python
from openpyxl.workbook.defined_name import DefinedName

wb.defined_names.add(DefinedName("Regiony", attr_text="'Słownik'!$A$2:$A$10"))
dv = DataValidation(type="list", formula1="Regiony", allow_blank=True)
ws.add_data_validation(dv)
dv.add("B2:B500")
```

Taka reguła jest rozumiana przez openpyxl **jako walidacja standardowa** i przechodzi przez wczytanie i zapis. To jest udokumentowane przez społeczność openpyxl jako działające obejście dla plików szablonowych. **Zastrzeżenie:** traktuj to jako „potrafi częściowo" — przetestuj na swojej wersji i obejrzyj plik w Excelu po zapisie.

**Obejście 2 — odbudowa reguł z rozszerzenia.** Skoro reguły **nadal są w oryginalnym pliku**, możesz je wyczytać z `x14` i **odtworzyć jako zwykłe `DataValidation`**:

```python
"""Odczytuje walidacje z bloku x14 i odtwarza je jako zwykle DataValidation."""
import re
import zipfile

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation


def extension_validations(path) -> dict:
    """Mapuje czesc arkusza -> liste regul {type, formula, sqref} z x14."""
    znalezione = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            xml = z.read(name).decode("utf-8")
            reguly = []
            for blok in re.findall(
                r"<x14:dataValidation\b(.*?)</x14:dataValidation>", xml, re.S
            ):
                reguly.append({
                    "type": re.search(r'type="(\w+)"', blok)[1],
                    "formula": re.search(r"<xm:f>(.*?)</xm:f>", blok, re.S)[1],
                    "sqref": re.search(r"<xm:sqref>(.*?)</xm:sqref>", blok, re.S)[1],
                })
            if reguly:
                znalezione[name] = reguly
    return znalezione
```

Dalej każdą regułę odtwarzasz normalnie: `DataValidation(type=..., formula1=f"={formula}", allow_blank=True)` + `add_data_validation` + `add(sqref)`. **Uwaga na przedrostek `=`** — w tej konkretnej rekonstrukcji jest on potrzebny, bo odtwarzasz formułę z surowego XML-a; sprawdź efekt w Excelu.

To dokładnie ten sam wzorzec co „przepisz to, czego narzędzie nie umie przenieść" z modułu 18: **nie polegaj na tym, że biblioteka doniesie wszystko; sprawdź, wypisz braki i odbuduj je jawnie.**

### 2.13. Wzorce: formularz, słownik, arkusz wynikowy

Trzy wzorce, które w tym module są chlebem powszednim. To zapowiedź modułów 23–26.

**Wzorzec A — formularz wejściowy.** Sekwencja, która **musi** być w tej kolejności myślowo (mechanicznie openpyxl jest wyrozumiały, ale bez tego porządku się pogubisz):

1. **Zbuduj strukturę** — nagłówki, opisy, wymiary kolumn.
2. **Utwórz słownik** (patrz wzorzec B) i **walidacje** odwołujące się do niego.
3. **Odblokuj pola do wpisywania** — `Protection(locked=False)` + wyraźne wypełnienie tła (użytkownik ma widzieć, gdzie wpisywać!).
4. **Zablokuj i ukryj formuły** — `Protection(hidden=True)` + domyślne `locked=True`.
5. **Włącz ochronę arkusza** — `ws.protection.sheet = True` + hasło.
6. **Sprawdź w Excelu**, że da się wpisywać w białe pola i że nie da się w pozostałe.

Analogia: **drukujesz formularz urzędowy i zostawiasz białe kratki w miejscach, które wypełnia petent.** Nie da się tego zrobić w złej kolejności — najpierw projekt, potem druk, potem pieczątka „wzór zatwierdzony".

**Wzorzec B — słownik jako arkusz-konfiguracja.** Listy wartości nie mieszkają w kodzie Pythona, tylko w **arkuszu `Słownik`**:

```python
ARKUSZE = {
    "formularz": "Formularz",
    "slownik": "Słownik",
    "meta": "_meta",
}

SŁOWNIK = {
    "Status": ["Nowe", "W toku", "Zamknięte", "Anulowane"],
    "Region": ["Północ", "Południe", "Wschód", "Zachód"],
    "Priorytet": ["Niski", "Normalny", "Wysoki", "Krytyczny"],
}

def zbuduj_slownik(ws) -> dict:
    """Zapisuje slownik i zwraca mapowanie: nazwa -> zakres A1."""
    zakresy = {}
    kolumna = 1
    for nazwa, wartosci in SŁOWNIK.items():
        litera = get_column_letter(kolumna)
        ws.cell(row=1, column=kolumna, value=nazwa)
        for i, w in enumerate(wartosci, start=2):
            ws.cell(row=i, column=kolumna, value=w)
        zakresy[nazwa] = f"$'{ws.title}'".replace("'", "") + \
                         f"!${litera}$2:${litera}${len(wartosci) + 1}"
        kolumna += 1
    return zakresy
```

Zalety: **słownik edytuje człowiek, nie programista**; jedna zmiana w arkuszu zmienia wszystkie listy; nowa wartość nie wymaga wdrożenia. I — co ważne dla tego modułu — jeśli uprzesz się przy odwołaniu do arkusza, to w tej właśnie sytuacji **musi** zadziałać albo nazwa zdefiniowana, albo odbudowa z `x14`.

Arkusz słownika warto **ukryć**: `ws.sheet_state = "hidden"` albo `"veryHidden"` (moduł 08). Analogia: **słownik leży w szufladzie biurka, nie na blacie** — użytkownik z niego korzysta, ale nie musi go widzieć.

**Wzorzec C — arkusz wynikowy.** Wyniki obliczeń wychodzą na osobny arkusz, który jest **w pełni zablokowany** (domyślne `locked=True` wszędzie), z **ukrytymi formułami**, chroniony hasłem. Formularz jest dla człowieka; arkusz wynikowy dla audytu.

### 2.14. Zapowiedź: walidator jako obiekt (Strategy, moduł 25)

Zwróć uwagę, ile powtórzeń już się pojawia:

```python
dv = DataValidation(type="list", formula1='"Tak,Nie"', allow_blank=True)
dv.error = 'Wybierz Tak albo Nie.'
dv.errorTitle = 'Nieprawidłowa wartość'
dv.errorStyle = 'stop'
dv.prompt = 'Wybierz wartość z listy.'
dv.promptTitle = 'Wybór'
ws.add_data_validation(dv)
dv.add("B2:B100")
```

Siedem linii na **jedną** regułę. Przy dwudziestu polach formularza to 140 linii i dwadzieścia miejsc, w których można się pomylić (jedno pominięte `dv.add` i pole jest niechronione).

W module 25 zobaczysz, jak to opisać **deklaratywnie**:

```python
REGULY = [
    {"pole": "B2:B100",  "typ": "list", "zrodlo": "Regiony",
     "komunikat": "Wybierz region z listy."},
    {"pole": "C2:C100",  "typ": "date", "operator": "greaterThanOrEqual",
     "formula": "=DATE(2026,1,1)", "blad": "Data nie może być z 2025 roku."},
    {"pole": "D2:D100",  "typ": "textLength", "operator": "lessThanOrEqual",
     "formula": 200, "blad": "Uwagi mogą mieć maks. 200 znaków."},
]
```

...i **jedna funkcja**, która z tej listy robi `DataValidation` po `DataValidation`. To wzorzec **Strategy** połączony z prostą **Specification**: reguła opisuje *co* ma być sprawdzone, a kod — *jak* to przetłumaczyć na openpyxl. Cała przewaga jest w tym, że logikę biznesową możesz testować **bez openpyxl** (moduły 20 i 26).

## 3. Przykłady krok po kroku

### Przykład 1 — lista rozwijana od zera i odczyt ze pliku (🟢)

```python
"""Modul 17 - lista rozwijana: tworzenie i weryfikacja.

Uruchom:  python examples/17_lista_podstawy.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "17_lista_podstawy.xlsx"

# Stale w jednym miejscu (modul 09: rejestr)
STYL_NAGLOWEK = Font(bold=True)
STYL_POLE_INPUT = PatternFill("solid", fgColor="FFF7E6")   # delikatne tlo "wpisz tu"


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Ankieta"

    ws["A1"] = "Pytanie"
    ws["B1"] = "Odpowiedź"
    ws["A1"].font = STYL_NAGLOWEK
    ws["B1"].font = STYL_NAGLOWEK

    ws["A2"] = "Czy wdrożyć zmiany?"
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 20

    # --------------------------------------------------------------
    # KROK 1: regula. Lista wprost - cudzysloby SA CZESCIA napisu!
    # --------------------------------------------------------------
    opcje = ["Tak", "Nie", "Nie wiem"]
    formula = '"' + ",".join(opcje) + '"'
    print(f"formula1 = {formula!r}   dlugosc napisu = {len(formula)}")
    print(f"(limit Excela dla listy wprost to 255 znakow)")

    dv = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
        # showDropDown NIE ustawiamy - domyslne False = strzalka WIDOCZNA (2.4!)
    )

    # komunikaty: prompt = podpowiedz przy wejsciu, error = blad
    dv.promptTitle = "Wybór"
    dv.prompt = "Wybierz jedną z trzech wartości z listy."
    dv.errorTitle = "Nieprawidłowa wartość"
    dv.error = "Dozwolone są tylko: Tak, Nie, Nie wiem."
    dv.errorStyle = "stop"          # stop | warning | information

    # --------------------------------------------------------------
    # KROK 2: REJESTRACJA na arkuszu. Bez tego plik nie ma walidacji!
    # --------------------------------------------------------------
    ws.add_data_validation(dv)

    # --------------------------------------------------------------
    # KROK 3: zakresy. Mozna podac napis albo obiekt Cell.
    # --------------------------------------------------------------
    dv.add("B2:B100")

    # kontrole "w locie" - to dziala, bo DataValidation ma __contains__
    print(f"'B50' in dv -> {'B50' in dv}")
    print(f"'C50' in dv -> {'C50' in dv}")

    # dekoracja pola wpisywania
    for row in ws.iter_rows(min_row=2, max_row=20, min_col=2, max_col=2):
        for c in row:
            c.fill = STYL_POLE_INPUT

    wb.save(cel)
    wb.close()
    print(f"\nZapisano: {cel}")


def weryfikuj(cel: Path) -> None:
    """Wczytuje plik i wypisuje, co NAPRAWDE jest w srodku.

    UWAGA: NIE read_only! Tam walidacji nie ma (2.11).
    """
    wb = load_workbook(cel)
    try:
        ws = wb["Ankieta"]
        dv_list = ws.data_validations          # DataValidationList

        print()
        print("=" * 78)
        print(f"Reguł walidacji w arkuszu: {len(dv_list)}")
        print("=" * 78)

        for i, dv in enumerate(dv_list.dataValidation, 1):
            print(f"[{i}] type={dv.type!r}  operator={dv.operator!r}")
            print(f"    formula1={dv.formula1!r}")
            print(f"    allowBlank={dv.allowBlank}  "
                  f"showDropDown={dv.showDropDown}  "
                  f"errorStyle={dv.errorStyle!r}")
            print(f"    promptTitle={dv.promptTitle!r}")
            print(f"    errorTitle={dv.errorTitle!r}")
            print(f"    sqref={dv.sqref}   typ={type(dv.sqref).__name__}")
            print(f"    zakresy ({len(dv.sqref)}):")
            for rng in dv.sqref:
                print(f"        {rng}")
    finally:
        wb.close()

    print()
    print("=" * 78)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 78)
    print("  1. Kliknij B2 -> po prawej stronie komorki pojawi sie STRZALKA listy.")
    print("     Jesli strzalki nie ma, sprawdz showDropDown (2.4).")
    print("  2. Kliknij B2 raz jeszcze -> zobaczysz TITLE 'Wybór' i podpowiedz")
    print("     'Wybierz jedną z trzech wartości z listy.' (prompt)")
    print("  3. Wpisz z palca 'Moze' -> Excel pokaze blad z tytulem")
    print("     'Nieprawidłowa wartość' i stylem 'stop' (nie da sie zatwierdzic).")
    print("  4. Wpisz 'nie wiem' malymi literami -> Excel NIE rozroznia wielkosci")
    print("     liter w listach, wiec to przejdzie. Sprawdz!")
    print("  5. Wpisz 'Tak' -> przechodzi.")
    print()
    print("EKSPERYMENT 'OTWARTE DRZWI OD KUCHNI' (sekcja 1.1):")
    print("  uruchom:  python -c \"import openpyxl; wb=openpyxl.load_workbook(")
    print("      'output/17_lista_podstawy.xlsx'); wb['Ankieta']['B2']='BZDURA';")
    print("      wb.save('output/obejscie.xlsx')\"")
    print("  Otwórz output/obejscie.xlsx w Excelu: w B2 jest 'BZDURA', a Excel")
    print("  nie protestowal przy zapisie. Walidacja dziala TYLKO na wejsciu")
    print("  przez interfejs Excela.")


def main() -> None:
    zbuduj(CEL)
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Po `ws.add_data_validation(dv)` reguła trafia do `ws.data_validations.dataValidation` (lista). `dv.add("B2:B100")` **dokłada zakres do `sqref`** — a `sqref` to `MultiCellRange`. `dv.sqref` jest więc iterowalny i możesz po nim chodzić pętlą, jak pokazuje `weryfikuj()`. Zapis pliku następuje dopiero w `wb.save`.

**Co trafi do pliku.** W `xl/worksheets/sheet1.xml` pojawi się blok:

```xml
<dataValidations count="1">
  <dataValidation type="list" allowBlank="1" showErrorMessage="1"
                  showInputMessage="1" errorStyle="stop"
                  promptTitle="Wybór" prompt="Wybierz jedną z trzech..."
                  errorTitle="Nieprawidłowa wartość"
                  error="Dozwolone są tylko: Tak, Nie, Nie wiem.">
    <formula1>"Tak,Nie,Nie wiem"</formula1>
    <sqref>B2:B100</sqref>
  </dataValidation>
</dataValidations>
```

Trzy rzeczy do przemyślenia:

1. **`allowBlank="1"` i `showErrorMessage="1"` pojawiły się same.** Nie ustawiałeś `showErrorMessage`, a ono ma wartość `1` w XML-u. To skutek tego, że w `__init__` `showErrorMessage=False`, ale gdy ustawisz **`error` i `errorTitle`**, Excel potrzebuje je gdzieś pokazać — a i tak najbezpieczniej jest **jawnie** ustawić `showErrorMessage=True`, jeśli chcesz, by błąd blokował wpis.
2. **`sqref` w XML to `B2:B100` — a `sqref` w Pythonie to `MultiCellRange`.** Format XML-owy różni się: przy wielu zakresach zobaczysz `A1 B2:B5 C7` — **spacje**, nie przecinki. To ta sama różnica, którą znasz z `Table.ref` (moduł 12).
3. **Cała reguła siedzi w `sheet1.xml`.** Nie ma osobnego pliku walidacji — w odróżnieniu od komentarzy (moduł 16), które miały osobny `comments1.xml` **i** VML. Walidacja jest lekka, więc jest w miejscu.

### Przykład 2 — wszystkie siedem typów walidacji w jednej tabeli (🟡)

```python
"""Modul 17 - wszystkie typy walidacji + opisy przykladu uzycia.

Uruchom:  python examples/17_typy_walidacji.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import quote_sheetname
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "17_typy_walidacji.xlsx"

STYL_NAGLOWEK = Font(bold=True)

# (pole, type, operator, formula1, formula2, opis)
REGUŁY = [
    ("B2:B50",   "list",       None,             '"Nowe,W toku,Zamknięte"', None,
     "Lista decyzyjna"),

    ("C2:C50",   "whole",      "greaterThan",    0,    None,
     "Liczba całkowita > 0"),

    ("D2:D50",   "whole",      "between",        1,    100,
     "Liczba 1..100"),

    ("E2:E50",   "decimal",    "between",        0,    1,
     "Ułamek 0..1 (np. 0,25 = 25%)"),

    ("F2:F50",   "date",       "greaterThanOrEqual", "=DATE(2026,1,1)", None,
     "Data nie starsza niż 2026-01-01"),

    ("G2:G50",   "time",       "lessThan",       "=TIME(18,0,0)", None,
     "Godzina przed 18:00 (serial czasu)"),

    ("H2:H50",   "textLength", "lessThanOrEqual", 200, None,
     "Maks. 200 znaków"),

    ("I2:I50",   "custom",
     None, "=ISNUMBER(J2)", None,
     "Formuła własna: I2 wymaga liczby w J2"),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Walidacje"

    naglowki = ["Lista", "Całkowita >0", "1..100", "Ułamek 0..1",
                "Data", "Godzina", "Tekst ≤200", "Custom"]
    for i, nazwa in enumerate(naglowki, start=1):
        c = ws.cell(row=1, column=i, value=nazwa)
        c.font = STYL_NAGLOWEK
        ws.column_dimensions[c.column_letter].width = 16

    for pole, typ, operator, f1, f2, opis in REGUŁY:
        dv = DataValidation(
            type=typ,
            operator=operator,
            formula1=f1,
            formula2=f2,
            allow_blank=True,
            showErrorMessage=True,     # jawnie: blad ma byc pokazany
        )
        dv.errorTitle = f"Błąd ({opis})"
        dv.error = f"Pole łamie regułę: {opis}."
        ws.add_data_validation(dv)
        dv.add(pole)
        print(f"{pole:<10} type={typ:<10} operator={operator!s:<18} "
              f"f1={f1!r}  f2={f2!r}")

    # --- dodatkowo: lista z INNEGO arkusza + nazwa zdefiniowana -------
    slownik = wb.create_sheet("Słownik")
    slownik["A1"] = "Region"
    for i, region in enumerate(["Północ", "Południe", "Wschód", "Zachód"], start=2):
        slownik[f"A{i}"] = region
    slownik.sheet_state = "hidden"        # schowaj slownik (modul 08)
    slownik["D1"] = "Status"
    for i, status in enumerate(["Nowe", "W toku", "Zamknięte"], start=2):
        slownik[f"D{i}"] = status

    ws["K1"] = "Region (z innego arkusza)"
    ws["L1"] = "Status (nazwa zdefiniowana)"
    ws.column_dimensions["K"].width = 24
    ws.column_dimensions["L"].width = 24

    # PETLA 1 (K): odwolanie BEZPOSREDNIE do arkusza - openpyxl to zapisze,
    # ale jesli plik potem zapisze Excel, regula trafi do bloku x14 (2.12).
    dv_k = DataValidation(
        type="list",
        formula1="{0}!$A$2:$A$5".format(quote_sheetname(slownik.title)),
        allow_blank=True,
    )
    ws.add_data_validation(dv_k)
    dv_k.add("K2:K50")
    print(f"\nK2:K50 formula1 = {dv_k.formula1!r}")

    # PETLA 2 (L): przez NAZWE ZDEFINIOWANA - forma najodporniejsza.
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names.add(
        DefinedName("Statusy", attr_text=f"'{slownik.title}'!$D$2:$D$4")
    )
    dv_l = DataValidation(type="list", formula1="Statusy", allow_blank=True)
    ws.add_data_validation(dv_l)
    dv_l.add("L2:L50")
    print(f"L2:L50 formula1 = {dv_l.formula1!r}")

    wb.save(cel)
    wb.close()
    print(f"\nZapisano: {cel}")


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Walidacje"]
        print()
        print("=" * 100)
        print(f"Reguł w arkuszu 'Walidacje': {len(ws.data_validations)}")
        print("=" * 100)
        print(f"{'#':<4}{'type':<12}{'operator':<20}{'f1':<34}{'sqref'}")
        print("-" * 100)
        for i, dv in enumerate(ws.data_validations.dataValidation, 1):
            f1 = str(dv.formula1)[:32]
            print(f"{i:<4}{str(dv.type):<12}{str(dv.operator):<20}"
                  f"{f1:<34}{dv.sqref}")

        # nazwy zdefiniowane (modul 26 wprowadzi je pelniej)
        print()
        print("Nazwy zdefiniowane w skoroszycie:")
        for nazwa, dn in wb.defined_names.items():
            print(f"  {nazwa!r} -> {dn.attr_text!r}")
    finally:
        wb.close()

    print()
    print("=" * 100)
    print("EKSPERYMENT 'X14' - NAJWAZNIEJSZY W TYM MODULE")
    print("=" * 100)
    print("  1. Otwórz output/17_typy_walidacji.xlsx W EXCELU.")
    print("  2. Kliknij na komorke L2 (walidacja przez nazwe) i sprawdz, ze")
    print("     lista rozwijana DZIALA.")
    print("  3. Zapisz plik z Excela pod nazwa 17_po_excelu.xlsx i zamknij.")
    print("  4. Uruchom:  python examples/17_inwentarz.py output/17_po_excelu.xlsx")
    print("  5. Uruchom:  python examples/17_inwentarz.py output/17_typy_walidacji.xlsx")
    print("     Porownaj liczbe regul PRZED i PO przejsciu przez Excela.")
    print()
    print("  Dlaczego: Excel zapisuje reguly odwolujace sie do innego arkusza")
    print("  w bloku rozszerzen x14. openpyxl tego bloku nie rozumie ->")
    print("  ostrzezenie 'Data Validation extension is not supported and")
    print("  will be removed'. Regula oparta na NAZWIE ZDEFINIOWANEJ jest")
    print("  zapisywana w standardowym miejscu i przez to przetrwa (2.12).")


def main() -> None:
    zbuduj(CEL)
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Osiem reguł budowanych w pętli — **każda to osobny obiekt `DataValidation`**. To ważne: openpyxl **nie scala** automatycznie reguł o identycznych parametrach. Jeśli masz osiem pól z identyczną listą, dostaniesz **osiem reguł**, a nie jedną z ośmioma zakresami. Dlatego — jak pokazuje to pętla — **budowanie reguł w pętli jest tanie**, ale **jedna reguła na wiele zakresów jest jeszcze tańsza**: wystarczy `dv.add("B2:B50")`, `dv.add("D2:D50")` na tym samym obiekcie.

**Co trafi do pliku.** Blok `<dataValidations count="10">` z dziesięcioma `<dataValidation>`, każda z własnym `<sqref>`. Plus — w `xl/workbook.xml` — wpis w `<definedNames>`:

```xml
<definedNames>
  <definedName name="Statusy">'Słownik'!$D$2:$D$4</definedName>
</definedNames>
```

**Trzy rzeczy do przemyślenia:**

1. **Walidacje `date` i `time` operują na serialu Excela, nie na `datetime`.** `=DATE(2026,1,1)` to formuła, którą ocenia Excel. Ale gdybyś chciał podać wartość liczbowo, musiałbyś użyć konwersji z modułu 05 (`to_excel`). **Formuła jest tu bezpieczniejsza i czytelniejsza** — nie musisz znać serialu. To jedna z niewielu sytuacji, gdzie formuła w arkuszu jest **lepsza** niż wartość z Pythona (bo tu nie potrzebujesz *wyniku*, tylko *wzorca porównania*).
2. **`operator` działa tylko dla typów liczbowych i tekstowych.** Dla `list` i `custom` jest ignorowany (w XML-u po prostu się nie pojawi). Ustawienie go tam nie szkodzi, ale nie rób tego „na wszelki wypadek" — czytelnik kodu będzie się zastanawiał, co miałeś na myśli.
3. **Ukryty arkusz `Słownik` nadal działa jako źródło listy rozwijanej.** Ukrycie (`sheet_state = "hidden"`) nie wpływa na działanie odwołań — to tylko kwestia widoczności dla użytkownika. Zapamiętaj to, bo to najczęstszy sposób budowy „list słownikowych": arkusz jest schowany, ale jego wartości są widoczne w listach rozwijanych.

### Przykład 3 — formularz wejściowy z ochroną (🔴)

To najważniejszy przykład modułu: łączy walidację, wzorzec „białych pól" i ochronę.

```python
"""Modul 17 - formularz wejściowy: walidacja + odblokowane pola + ochrona.

Kolejnosc myslenia (sekcja 2.13):
  1. struktura  2. slownik + walidacje  3. odblokuj pola
  4. zablokuj/ukryj formuly  5. wlacz ochrone  6. sprawdz

Uruchom:  python examples/17_formularz.py
"""

from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "17_formularz.xlsx"

HASŁO = "Kurs2026!"

# Rejestr wizualny - jedno miejsce decyzji (modul 09)
STYL = {
    "naglowek":   Font(bold=True, size=12),
    "etykieta":   Font(bold=True),
    "pole":       PatternFill("solid", fgColor="FFF7E6"),   # biale kratki
    "wynik_bg":   PatternFill("solid", fgColor="EAF3EA"),
    "wynik_font": Font(italic=True, color="3BA55C"),
}

# Slownik: lista wartosci per pole (wzorzec B z 2.13)
SLOWNIK = {
    "Region":    ["Północ", "Południe", "Wschód", "Zachód"],
    "Kategoria": ["Sprzęt", "Usługi", "Licencje", "Szkolenia"],
    "Status":    ["Nowe", "W toku", "Zamknięte"],
}

# Pola formularza: (wiersz, etykieta, typ, operator, f1, f2, komunikat)
POLA = [
    (4,  "Data",             "date",   "greaterThanOrEqual", "=DATE(2026,1,1)", None,
         "Data nie może być wcześniejsza niż 2026-01-01."),
    (5,  "Kwota netto",      "decimal", "between", 0, 10_000_000,
         "Kwota musi być liczbą z zakresu 0 – 10 000 000."),
    (6,  "Ilość",            "whole",   "greaterThan", 0, None,
         "Ilość musi być liczbą całkowitą większą od zera."),
    (7,  "Uwagi (≤200 zn.)", "textLength", "lessThanOrEqual", 200, None,
         "Uwagi mogą mieć maksymalnie 200 znaków."),
]

# Wiersze 11-13 to formuly - zablokowane i UKRYTE
FORMUŁY = [
    (11, "Suma netto",  "=SUM(B5:B5)"),              # przyklad roboczy
    (12, "Liczba pól",  "=COUNTA(B4:B7)"),
    (13, "Wersja pliku", "=\"1.0\""),
]


def _ukryj_formule(cell) -> None:
    """Ustawia hidden=True BEZ gubienia pozostalych pol Protection.

    Modul 09: ochrony tez nie mozna mutowac - trzeba zrobic kopie i przypisac.
    """
    nowa = copy(cell.protection)
    nowa.hidden = True
    nowa.locked = True          # formula ma byc i zablokowana, i ukryta
    cell.protection = nowa


def _odblokuj(cell) -> None:
    """Odblokowuje komorke do wpisywania + daje jej tlo 'wpisz tu'."""
    nowa = copy(cell.protection)
    nowa.locked = False
    cell.protection = nowa
    cell.fill = STYL["pole"]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Formularz"

    # ---------- 1. STRUKTURA --------------------------------------
    ws["A1"] = "Formularz zapotrzebowania"
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 44

    ws["A2"] = "Wypełnij białe pola. Pozostałe komórki są zablokowane."
    ws["A2"].font = Font(italic=True, size=9)

    # ---------- 2. SLOWNIK + WALIDACJE -----------------------------
    slownik = wb.create_sheet("Słownik")
    zakresy = {}
    for i, (nazwa, wartosci) in enumerate(SLOWNIK.items(), start=1):
        litera = get_column_letter(i)
        slownik.cell(row=1, column=i, value=nazwa).font = STYL["etykieta"]
        for r, w in enumerate(wartosci, start=2):
            slownik.cell(row=r, column=i, value=w)
        zakresy[nazwa] = (f"{quote_sheetname('Słownik')}"
                          f"!${litera}$2:${litera}${len(wartosci) + 1}")

    # (a) walidacje list - TRZY osobne pola: Region, Kategoria, Status
    pola_listowe = [
        ("C4:C200", "Region",    "Wybierz region z listy."),
        ("C5:C200", "Kategoria", "Wybierz kategorię z listy."),
        ("C6:C200", "Status",    "Wybierz status z listy."),
    ]
    for zakres, nazwa_slownika, komunikat in pola_listowe:
        dv = DataValidation(
            type="list",
            formula1=zakresy[nazwa_slownika],
            allow_blank=True,
            showErrorMessage=True,
        )
        dv.promptTitle = "Wybór z listy"
        dv.prompt = komunikat
        dv.errorTitle = "Wartość spoza listy"
        dv.error = f"Dozwolone wartości: {', '.join(SLOWNIK[nazwa_slownika])}."
        dv.errorStyle = "stop"
        ws.add_data_validation(dv)              # KROK 2 - bez tego nie ma DV!
        dv.add(zakres)
        print(f"DV list  {zakres:<12} -> {nazwa_slownika}  "
              f"({len(SLOWNIK[nazwa_slownika])} pozycji)")

    # (b) walidacje liczbowe/datowe/tekstowe - z listy POLA
    for wiersz, etykieta, typ, operator, f1, f2, komunikat in POLA:
        dv = DataValidation(type=typ, operator=operator,
                            formula1=f1, formula2=f2,
                            allow_blank=True, showErrorMessage=True)
        dv.errorTitle = f"Błędna wartość: {etykieta}"
        dv.error = komunikat
        dv.errorStyle = "stop"
        ws.add_data_validation(dv)
        dv.add(f"B{wiersz}:B200")
        print(f"DV {typ:<10} B{wiersz}:B200  op={operator}")

    # ---------- 3. ODBLOKOWANIE POL -------------------------------
    WSZYSTKIE_POLA = ["B4:B200", "B5:B200", "B6:B200", "B7:B200",
                      "C4:C200", "C5:C200", "C6:C200"]

    # etykiety w kolumnie A
    for wiersz, etykieta, *_ in POLA:
        ws.cell(row=wiersz, column=1, value=etykieta).font = STYL["etykieta"]

    for zakres in WSZYSTKIE_POLA:
        for row in ws[zakres]:
            for c in row:
                _odblokuj(c)

    # ---------- 4. ZABLOKOWANIE + UKRYCIE FORMUL ------------------
    ws["A11"].font = STYL["etykieta"]
    for wiersz, etykieta, formula in FORMUŁY:
        c_etykieta = ws.cell(row=wiersz, column=1, value=etykieta)
        c_etykieta.font = STYL["etykieta"]
        c_wynik = ws.cell(row=wiersz, column=2, value=formula)
        c_wynik.fill = STYL["wynik_bg"]
        c_wynik.font = STYL["wynik_font"]
        _ukryj_formule(c_wynik)          # locked=True + hidden=True
        print(f"Formula {c_wynik.coordinate}: hidden=True, locked=True")

    # ---------- 5. OCHRONA ARKUSZA --------------------------------
    print()
    print("Stan flag przed wlaczeniem ochrony:")
    print(f"  ws.protection.sheet = {ws.protection.sheet}")
    print(f"  formatCells={ws.protection.formatCells}  "
          f"sort={ws.protection.sort}  autoFilter={ws.protection.autoFilter}")
    print("  (True = ZABLOKOWANE, False = dozwolone - sekcja 2.6)")

    # Zostawmy uzytkownikowi mozliwosc korzystania z filtra i sortowania
    ws.protection.sort = False            # sortowanie DOZWOLONE
    ws.protection.autoFilter = False      # strzalki filtra DOZWOLONE
    ws.protection.formatCells = False     # formatowanie komorek DOZWOLONE
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

    # PW: ustawienie hasla AUTOMATYCZNIE wlacza ochrone (sekcja 2.7)
    ws.protection.password = HASŁO
    print(f"\nPo ustawieniu hasla: ws.protection.sheet = {ws.protection.sheet}")

    # slownik jest konfiguracja - schowaj go
    slownik.sheet_state = "hidden"

    wb.active = wb.sheetnames.index("Formularz")

    wb.save(cel)
    wb.close()
    print(f"Zapisano: {cel}")


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Formularz"]
        print()
        print("=" * 84)
        print("INWENTARZ OCHRONY I WALIDACJI")
        print("=" * 84)

        p = ws.protection
        print(f"sheet={p.sheet}  (bool() = {bool(p)!r})")
        print(f"haslo zapisane w pliku: {p.password!r}   <- 4 znaki hex!")
        print()
        print("Flagi (True = ZABLOKOWANE):")
        for flaga in ("objects", "scenarios", "formatCells", "formatColumns",
                      "formatRows", "insertColumns", "insertRows",
                      "insertHyperlinks", "deleteColumns", "deleteRows",
                      "selectLockedCells", "selectUnlockedCells",
                      "sort", "autoFilter", "pivotTables"):
            print(f"  {flaga:<22} = {getattr(p, flaga)}")

        print()
        print("Komorki z hidden=True (ukryte formuly):")
        licznik = 0
        for row in ws.iter_rows():
            for c in row:
                if c.protection and c.protection.hidden:
                    licznik += 1
                    print(f"  {c.coordinate}: locked={c.protection.locked} "
                          f"hidden={c.protection.hidden}  wartość={c.value!r}")
        print(f"  razem: {licznik}")

        print()
        print(f"Reguł walidacji: {len(ws.data_validations)}")
    finally:
        wb.close()

    print()
    print("=" * 84)
    print("CO SPRAWDZIC W EXCELU (to jest sedno tego przykladu)")
    print("=" * 84)
    print("  1. Kliknij B6 i SPRÓBUJ wpisać cokolwiek -> pojawi się komunikat")
    print("     'Błędna wartość: Ilość' z Twoim tekstem błędu.")
    print("  2. Kliknij B4 -> strzałka listy Region. Wybierz wartość spoza")
    print("     listy z palca -> komunikat 'Wartość spoza listy'.")
    print("  3. Kliknij komórkę B11 (Suma netto). SPRÓBUJ coś wpisać ->")
    print("     Excel odmówi: 'Ta komórka jest na chronionym arkuszu'.")
    print("  4. Kliknij B11 i popatrz na PASEK FORMUŁY -> zamiast")
    print("     '=SUM(B5:B5)' zobaczysz PUSTO/ukryte. To zasługa hidden=True")
    print("     + włączonej ochrony (sekcja 2.8).")
    print("  5. Spróbuj dodać wiersz albo kolumnę (prawy klik na nagłówku) ->")
    print("     zablokowane. Spróbuj posortować -> działa (sort=False).")
    print("  6. Zakładka 'Słownik' jest UKRYTA, ale lista Region nadal działa.")
    print("     Ukrycie arkusza nie psuje odwołań.")
    print()
    print("  EKSPERYMENT 'ZLAMANE HASLO PRZEZ PYTHON':")
    print("    openpyxl wczyta ten plik i BEZ HASŁA doda arkusz:")
    print("      wb = load_workbook('output/17_formularz.xlsx')")
    print("      wb.create_sheet('Hack'); wb.save('output/hack.xlsx')")
    print("    ...i to się uda. Ochrona to kłódka, nie sejf (sekcja 2.9).")


def main() -> None:
    zbuduj(CEL)
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Kluczowe momenty:

- **`copy(cell.protection)`** — dlatego, że w module 09 poznałeś **niemutowalność stylów**. `cell.protection.hidden = True` **nie zadziała** (albo zadziała lokalnie na obiekcie tymczasowym, który i tak nie trafi do komórki). Trzeba: skopiować → zmienić → przypisać.
- **`ws.protection.sort = False`** — to **pojedyncza flaga**, żadnego kopiowania nie wymaga. `SheetProtection` to zwykły obiekt, nie niemutowalny (nie jest stylem).
- **`ws.protection.password = HASŁO`** — ustawia skrót hasła **i włącza ochronę**. Dwukrotnie to sprawdzaj w kodzie, bo to nieoczywiste.

**Co trafi do pliku.** Trzy miejsca:

```xml
<!-- xl/worksheets/sheet1.xml -->
<sheetProtection sheet="1" objects="0" scenarios="0"
                 formatCells="0" formatRows="0" formatColumns="0"
                 insertColumns="1" insertRows="1" insertHyperlinks="1"
                 deleteColumns="1" deleteRows="1"
                 selectLockedCells="0" selectUnlockedCells="0"
                 sort="0" autoFilter="0" pivotTables="1" password="7C4A"/>
```

```xml
<!-- xl/styles.xml -> cellXfs (stan blokady MIESZKA W STYLU, sekcja 2.8) -->
<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1">
  <protection locked="0"/>              <!-- pole do wpisywania -->
</xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1">
  <protection locked="1" hidden="1"/>   <!-- formula: zablokowana i ukryta -->
</xf>
```

```xml
<!-- xl/workbook.xml -->
<definedNames>...</definedNames>       <!-- jesli uzywasz nazw -->
```

**Trzy rzeczy do przemyślenia:**

1. **`password="7C4A"` to skrót, nie hasło.** Cztery cyfry heksadecymalne. Twoje `Kurs2026!` nie występuje w pliku w żadnej czytelnej postaci — i to jest dobra wiadomość. Zła wiadomość: **65 536 możliwych wartości** tego skrótu, więc „odzyskanie" hasła sprowadza się do policzenia 65 536 skrótów i znalezienia pasującego. To kwestia milisekund. **Wiele różnych haseł odblokuje ten sam plik.**
2. **Pola listowe mają trzy **osobne** reguły, nie jedną.** Bo każda lista jest inna. Gdyby listy były identyczne, wystarczyłaby jedna reguła z wieloma zakresami `dv.add(...)` — mniej rekordu, szybszy plik. To ten sam rachunek co „jedna Tabela zamiast trzech zakresów" (moduł 12).
3. **Ukryty `Słownik` nadal służy jako źródło listy.** Ale uwaga: jeśli ten plik trafi z powrotem do Excela i zostanie **przez niego zapisany**, reguły w `C4`, `C5`, `C6` **prawdopodobnie wylądują w bloku `x14`** — i przy następnym przepuszczeniu przez openpyxl zostaną zgubione z ostrzeżeniem. To jest realistyczny scenariusz cyklu życia szablonu i dlatego warto rozważyć **nazwy zdefiniowane** zamiast odwołań bezpośrednich.

### Przykład 4 — inwentarz walidacji i ochrony z nasłuchem ostrzeżeń (🟡)

To narzędzie, które uruchamiasz **przed** każdym zapisem cudzego pliku. Bezpośrednio przygotowuje moduł 18.

```python
"""Modul 17 - inwentarz: walidacje, ochrona, ostrzezenia o utracie danych.

Uruchom:  python examples/17_inwentarz.py [sciezka.xlsx]
"""

from __future__ import annotations

import sys
import warnings
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
DOMYSLNY = ROOT / "output" / "17_formularz.xlsx"


def flaga(wartosc) -> str:
    """Czytelny zapis flagi ochrony (True = zablokowane)."""
    if wartosc is True:
        return "ZABLOKOWANE"
    if wartosc is False:
        return "dozwolone  "
    return f"{wartosc!r}"


def inwentarz(sciezka: Path) -> None:
    print("=" * 96)
    print(f"INWENTARZ WALIDACJI I OCHRONY: {sciezka.name}")
    print("=" * 96)

    # NASLUCH OSTRZEZEN - tu wykryjemy blok x14 (sekcja 2.12)!
    with warnings.catch_warnings(record=True) as zlapane:
        warnings.simplefilter("always")        # bez tego tylko PIERWSZE ostrzezenie

        # KRYTYCZNE: NIE read_only! Tam walidacji nie ma (sekcja 2.11)
        wb = load_workbook(sciezka)
        try:
            # ---------- OCHRONA SKOROSZYTU --------------------------
            sec = wb.security
            print("\nOCHRONA SKOROSZYTU (xl/workbook.xml -> workbookProtection)")
            print(f"  lockStructure     = {flaga(sec.lockStructure)}")
            print(f"  lockWindows       = {flaga(sec.lockWindows)}")
            print(f"  lockRevision      = {flaga(sec.lockRevision)}")
            print(f"  workbookPassword  = {sec.workbookPassword!r}")
            print(f"  revisionsPassword = {sec.revisionsPassword!r}")
            if sec.lockStructure and not sec.workbookPassword:
                print("  ⚠️  lockStructure=True BEZ hasla = brak egzekwowania (2.10)!")

            # ---------- OCHRONA ARKUSZY ----------------------------
            for ws in wb.worksheets:
                print(f"\nARKUSZ '{ws.title}'  (stan={ws.sheet_state})")
                p = ws.protection
                print(f"  ochrona arkusza: {flaga(p.sheet)}  "
                      f"hasz hasla={p.password!r}")
                if p.sheet:
                    aktywne = [
                        nazwa for nazwa in (
                            "objects", "scenarios", "formatCells",
                            "formatColumns", "formatRows", "insertColumns",
                            "insertRows", "insertHyperlinks", "deleteColumns",
                            "deleteRows", "selectLockedCells",
                            "selectUnlockedCells", "sort", "autoFilter",
                            "pivotTables",
                        ) if getattr(p, nazwa) is True
                    ]
                    print(f"  zablokowane akcje ({len(aktywne)}): "
                          f"{', '.join(aktywne) if aktywne else '-'}")

                # ---------- KOMORKI UKRYTE / ODBLOKOWANE -----------
                odblokowane, ukryte = [], []
                for row in ws.iter_rows():
                    for c in row:
                        prot = c.protection
                        if prot is None:
                            continue
                        if prot.locked is False:
                            odblokowane.append(c.coordinate)
                        if prot.hidden:
                            ukryte.append(c.coordinate)
                print(f"  komórki ODBLOKOWANE (do wpisywania): "
                      f"{len(odblokowane)}"
                      + (f"  -> {odblokowane[:6]}{'...' if len(odblokowane) > 6 else ''}"
                         if odblokowane else ""))
                print(f"  komórki z UKRYTĄ FORMUŁĄ: {len(ukryte)}"
                      + (f"  -> {ukryte}" if ukryte else ""))

                # ---------- WALIDACJE ------------------------------
                dv_list = ws.data_validations
                print(f"  reguł walidacji: {len(dv_list)}")
                for i, dv in enumerate(dv_list.dataValidation, 1):
                    f1 = str(dv.formula1)
                    ostrzezenie = ""
                    if dv.type == "list" and not f1.startswith('"'):
                        # lista z zakresu/nazwy - potencjalne ryzyko x14
                        ostrzezenie = "  <- lista spoza komórki: RYZYKO x14!"
                    if dv.type == "list" and len(f1) > 255:
                        ostrzezenie += "  <- PRZEKROCZONY LIMIT 255 ZNAKOW!"
                    print(f"    [{i}] type={dv.type}  "
                          f"sqref={dv.sqref}  f1={f1[:44]}{ostrzezenie}")

                if not ws.data_validations.dataValidation and not p.sheet:
                    continue
        finally:
            wb.close()

        # ---------- OSTRZEZENIA ------------------------------------
        print()
        print("=" * 96)
        print("OSTRZEZENIA openpyxl PRZY WCZYTANIU")
        print("=" * 96)
        if zlapane:
            for w in zlapane:
                print(f"  ⚠️  {w.category.__name__}: {w.message}")
            print()
            print("  ⚠️  KAŻDE ostrzeżenie = potencjalna utrata danych przy zapisie.")
            print("  Ostrzeżenie o 'Data Validation extension' oznacza, że plik ma")
            print("  walidacje w bloku x14 (typowo: listy z innego arkusza),")
            print("  których openpyxl NIE przeniesie do nowego pliku (sekcja 2.12).")
            print("  Obejścia: (1) nazwy zdefiniowane, (2) odbudowa z XML-a.")
        else:
            print("  Brak ostrzeżeń. To dobra wiadomość, ale nie gwarancja.")
            print("  Uwaga: walidacja, która w ogóle się nie zapisała (bo")
            print("  pominąłeś dv.add), nie generuje ostrzeżenia - jest cicho.")

    print()
    print("=" * 96)
    print("CHECKLISTA PRZED ZAPISEM (skrocona wersja modułu 18)")
    print("=" * 96)
    print("  [ ] Czy plik ma walidacje x14? -> zobacz ostrzeżenia wyżej.")
    print("  [ ] Czy ochrona arkusza ma hasło? Bez hasła użytkownik zdejmie")
    print("      ochronę jednym kliknięciem (2.9).")
    print("  [ ] Czy wiesz, że ochrona NIE jest bezpieczeństwem?")
    print("  [ ] Czy walidacje są kompletne (każda ma niepusty sqref)?")
    print("  [ ] Czy któraś lista wprost nie przekracza 255 znaków?")
    print("  [ ] ZAPISZ pod nową nazwą i uruchom ten inwentarz PONOWNIE;")
    print("      porównanie wyników to najtańszy test utraty danych.")


def main() -> None:
    sciezka = Path(sys.argv[1]) if len(sys.argv) > 1 else DOMYSLNY
    if not sciezka.exists():
        print(f"Brak pliku: {sciezka}")
        print("Uruchom najpierw: python examples/17_formularz.py")
        return
    inwentarz(sciezka)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `load_workbook` bez `read_only` parsuje **cały** arkusz — a więc także walidacje, ochronę i **style każdej komórki** (bo `cell.protection` to część stylu, sekcja 2.8). Na pliku 100 000 komórek ten inwentarz będzie **wolny** — to cena kompletności. W module 19 zobaczysz, jak to obejść dla ogromnych plików (podpowiedź: parsując XML bezpośrednio, tak jak robi to `extension_validations` z 2.12).

**Trzy rzeczy do przemyślenia:**

1. **`warnings.simplefilter("always")` to znowu najważniejsza linia.** Python domyślnie pokazuje każde **unikalne** ostrzeżenie tylko raz na sesję. Bez `always` przy pliku z dwoma takimi samymi problemami zobaczysz jedno ostrzeżenie i pomyślisz, że problem jest jeden. Ostrzeżenie o `x14` może wystąpić **raz na arkusz** — a bez `always` zobaczysz tylko pierwszy.
2. **Test „lista spoza komórki" to heurystyka, nie pewnik.** Sprawdzenie `not f1.startswith('"')` wychwyci większość przypadków, ale prawdziwa diagnoza polega na **zajrzeniu do XML-a** (`zipfile` + szukanie `x14:dataValidation`) — dokładnie tak, jak w `extension_validations` z 2.12.
3. **Inwentarz przed i po zapisie to Twój test regresji.** To ten sam wzorzec co w module 16: **uruchom narzędzie na oryginale, zapisz, uruchom na kopii, porównaj liczby.** Różnica to Twoja utrata danych. W module 18 zrobisz z tego pełny, automatyczny raport ryzyka.

## 4. Anatomia API

| Klasa / metoda / atrybut | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `DataValidation(type=, formula1=, formula2=)` | Obiekt reguły | `type`: `list`/`whole`/`decimal`/`date`/`time`/`textLength`/`custom` | To **nie** jest jeszcze reguła w arkuszu! |
| `dv.type` / `dv.validation_type` | Typ walidacji | `NoneSet` — 7 wartości | `validation_type` to alias (stare API) |
| `dv.operator` | Operator porównania | `between`, `notBetween`, `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan`, `greaterThanOrEqual` | Ignorowany dla `list` i `custom` |
| `dv.formula1`, `dv.formula2` | Warunek / lista | `NestedText` — napis; liczby też działają | Lista wprost: **cudzysłowy w napisie** |
| `dv.allowBlank` / `dv.allow_blank` | Puste pole dozwolone | `bool` | `allow_blank` w `__init__` **wygrywa** z `allowBlank` |
| `dv.showDropDown` | ⚠️ **ODWROTNIE:** `True` = **ukryj strzałkę** | `bool` (domyślnie `False`) | Alias: `hide_drop_down` — nazwa uczciwa |
| `dv.showErrorMessage` | Pokazuj błąd przy złej wartości | `bool` (domyślnie `False`) | Ustaw `True` **jawnie** dla formularzy |
| `dv.showInputMessage` | Pokazuj podpowiedź przy wejściu | `bool` (domyślnie `False`) | Wymagane, by `prompt` był widoczny |
| `dv.errorStyle` | Surowość błędu | `stop` (blokuje), `warning`, `information` | `stop` = nie da się zatwierdzić |
| `dv.errorTitle`, `dv.error` | Komunikat błędu | `str` | `stop` + komunikat = najlepszy UX |
| `dv.promptTitle`, `dv.prompt` | Komunikat wejścia | `str` | Wymaga `showInputMessage=True` |
| `dv.sqref` / `dv.cells` / `dv.ranges` | Zakresy reguły | `MultiCellRange` | **Aliasy tego samego** w 3.1.x |
| `dv.add(cell_or_ref)` | Dodaje zakres do `sqref` | napis (`"B2:B100"`) albo `Cell` | **Nie zastępuje** — dokłada |
| `"B4" in dv` | Czy komórka jest w regule | — | `__contains__` działa |
| `ws.add_data_validation(dv)` | **Rejestruje regułę na arkuszu** | `DataValidation` | **Pominięcie = reguła nie trafi do pliku** |
| `ws.data_validations` | Kolekcja reguł arkusza | `DataValidationList` | `len(ws.data_validations)` działa |
| `ws.data_validations.dataValidation` | Lista obiektów reguł | — | Do iteracji po odczytanym pliku |
| `ws.data_validations.append(dv)` | To samo co `add_data_validation` | — | API niższego poziomu |
| `quote_sheetname("Dane 2026")` | Cytuje nazwę arkusza | — | → `'Dane 2026'`; **wymagane dla list z innego arkusza** |
| `wb.defined_names.add(DefinedName(nazwa, attr_text=...))` | Nazwa zdefiniowana | `attr_text`: `'Arkusz'!$A$2:$A$10` | **Najodporniejsza forma listy** (2.12) |
| `ws.protection.sheet` / `ws.protection.enabled` | Główny wyłącznik ochrony | `bool` | `enabled` to alias |
| `ws.protection.enable()` / `.disable()` | Włącza / wyłącza ochronę | — | Zwykłe `sheet = True/False` |
| `ws.protection.password` | Hasło | `str` | ⚠️ **Setter WŁĄCZA ochronę** (`set_password` → `enable()`) |
| `ws.protection.set_password(value, already_hashed=False)` | Hasło z gotowym skrótem | — | Też **włącza** ochronę |
| `ws.protection.<flaga>` | 15 flag dostępu | `bool` — **`True` = ZABLOKOWANE** | `formatCells`, `sort`, `autoFilter`, `insertRows`, … |
| `bool(ws.protection)` | Czy chroniony | — | Zwraca `ws.protection.sheet` |
| `Protection(locked=, hidden=)` | Stan blokady **komórki** (styl) | `bool` | **Część stylu** → trafia do `styles.xml` |
| `cell.protection` | Obiekt `Protection` komórki | — | **Niemutowalny** — użyj `copy()` i przypisz |
| `wb.security` | `WorkbookProtection` | — | Domyślnie istnieje na `Workbook` |
| `wb.security.lockStructure` | Blokada struktury arkuszy | `bool` | **Egzekwowane tylko z hasłem** (2.10) |
| `wb.security.workbookPassword` | Hasło struktury | `str` | Ustawiany jako skrót |
| `wb.security.revisionsPassword` | Hasło zmian śledzonych | `str` | Legacy, rzadko potrzebne |
| `wb.security.lockWindows` | Blokada okien | `bool` | **Nowszy Excel nie oferuje** — zostaw `None` |
| `wb.security.lockRevision` | Blokada zmian śledzonych | `bool` | Legacy — zwykle pomijaj |
| `wb.security.set_workbook_password(v, already_hashed=True)` | Hasło z gotowym skrótem | — | Gdy masz skrót z innego źródła |
| `openpyxl.utils.protection.hash_password(plaintext)` | Skrót hasła (algorytm legacy) | — | 16 bitów → 4 znaki hex |
| `load_workbook(..., read_only=True)` | Tryb leniwy | — | **Walidacji i ochrony NIE MA** (2.11) |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — lista rozwijana „Tak / Nie / Nie wiem".**

Napisz `examples/17_cw1.py`, który tworzy `output/17_cw1.xlsx` z arkuszem `Ankieta`:

1. Nagłówki `Pytanie` / `Odpowiedź` w wierszu 1 (pogrubione).
2. Pięć pytań w kolumnie A (wiersze 3–7).
3. **Jedna** reguła `DataValidation` typu `list` z opcjami `Tak`, `Nie`, `Nie wiem`, zastosowana do `B3:B7` — jedną regułą, nie pięć!
4. `allow_blank=True`.
5. Komunikaty **po polsku**: `promptTitle="Wybór"`, `prompt="Wybierz jedną z trzech odpowiedzi."`, `errorTitle="Nieprawidłowa odpowiedź"`, `error="Dozwolone są tylko: Tak, Nie, Nie wiem."`, `errorStyle="stop"`.
6. `showErrorMessage=True` ustawione **jawnie**.
7. Pola `B3:B7` mają tło `FFF7E6` (żeby było widać, gdzie wpisywać).

Funkcja `weryfikuj()` ma:

- wypisać **liczbę reguł** w arkuszu,
- wypisać `formula1`, `allowBlank`, `showDropDown`, `errorStyle`, `sqref`,
- **policzyć i wypisać** `len(formula1)` oraz ostrzec, jeśli przekracza 255.

W pliku `output/17_cw1_wnioski.md` odpowiedz:

1. **Ile reguł jest w arkuszu — jedna czy pięć?** Zajrzyj do `ws.data_validations` i wyjaśnij, dlaczego jedna reguła może objąć pięć komórek (2.1 — `sqref` to zbiór).
2. **Ile znaków ma Twój `formula1`** i ile zapasu zostało do limitu Excela?
3. **Co by się stało, gdybyś wywołał `dv.add("B3:B7")`, ale zapomniał `ws.add_data_validation(dv)`?** Sprawdź to eksperymentalnie (utwórz plik, wczytaj go i policz reguły). Czy Python zgłosił błąd?
4. **Co by się stało przy odwrotnym błędzie:** `ws.add_data_validation(dv)` bez `dv.add(...)`? Sprawdź i wyjaśnij, odwołując się do `DataValidationList.to_tree`.
5. **Co się zmieni, gdy ustawisz `showDropDown=True`?** Sprawdź w Excelu i wyjaśnij zaskakujący efekt (2.4).
6. **Czy Excel rozróżnia wielkość liter w liście?** Wpisz `tak` zamiast `Tak` i sprawdź. Opisz wynik.
7. **Czy `openpyxl` sprawdził Twoją listę przy zapisie?** Wpisz przez openpyxl do `B3` wartość `"Kompletnie coś innego"`, zapisz i otwórz w Excelu. Opisz, co zrobił Excel (1.1).

### 🟡 Warsztat

**Zadanie 2 — lista z innego arkusza i komunikaty po polsku.**

Napisz `examples/17_cw2.py`, tworzący `output/17_cw2.xlsx`:

1. Arkusz `Zamówienia` z kolumnami: `Klient`, `Region`, `Kategoria`, `Status`, `Ilość`, `Uwagi`.
2. Arkusz `Słownik` (ukryty, `sheet_state="hidden"`) z trzema listami: `Region` (6 wartości), `Kategoria` (5 wartości), `Status` (4 wartości).
3. **Dla kolumny `Region` użyj odwołania bezpośredniego do arkusza** (`quote_sheetname`).
4. **Dla kolumny `Status` użyj nazwy zdefiniowanej** (`DefinedName` + `formula1="Statusy"`).
5. **Dla kolumny `Kategoria`** — użyj listy z zakresu, ale przez nazwę zdefiniowaną `="Kategorie"`.
6. Walidacja `Ilość`: `whole`, `greaterThan`, `0`.
7. Walidacja `Uwagi`: `textLength`, `lessThanOrEqual`, `200`.
8. **Wszystkie komunikaty po polsku**, `errorStyle="stop"`, `showErrorMessage=True`.
9. Pola do wpisywania mają tło `FFF7E6`, nagłówki są pogrubione i mają tło `DDDDDD`.

Funkcja `weryfikuj()` ma wypisać tabelę wszystkich reguł z kolumnami: `type`, `operator`, `formula1` (skrócone), `sqref`, oraz **flagę ryzyka**: `RYZYKO x14` dla list opartych na odwołaniu do arkusza.

W pliku `output/17_cw2_wnioski.md`:

1. **Czym różni się `formula1` dla trzech kolumn list?** Wklej trzy wartości i wyjaśnij (2.3).
2. **Dlaczego `quote_sheetname` był konieczny?** Co by się stało bez niego, skoro nazwa arkusza `Słownik` nie ma spacji? Sprawdź, czy da się bez i opisz różnicę.
3. **Jak wygląda nazwa zdefiniowana w `xl/workbook.xml`?** Rozpakuj plik i wklej odpowiedni fragment XML.
4. **Wykonaj test `x14`:** otwórz plik w Excelu, zapisz, zamknij, wczytaj przez `examples/17_inwentarz.py`. Ile reguł zostało? Które przetrwały? Uzasadnij wynik, odwołując się do 2.12 — i wyjaśnij, **dlaczego walidacja oparta na nazwie zdefiniowanej ma większą szansę przetrwać**.
5. **Jaki jest realny limit liczby reguł walidacji na arkusz?** Zbuduj test: jedna reguła na setki komórek vs setki reguł po jednej komórce. Zmierz rozmiar pliku i czas zapisu. Który wariant wybierzesz i dlaczego?
6. **Co by się stało, gdyby lista `Region` miała 40 pozycji o długich nazwach?** Policz długość napisu i powiedz, czy limit 255 zostałby przekroczony (2.5).

### 🔴 Wyzwanie

**Zadanie 3 — formularz chroniony: 5 pól odblokowanych, formuły ukryte, reszta zablokowana.**

Napisz `examples/17_cw3.py`, który buduje `output/17_cw3.xlsx` — działający formularz wejściowy.

**Wymagania:**

1. **Arkusz `Formularz`** z sekcjami: *Dane podstawowe* (5 pól), *Podsumowanie* (3 formuły), *Metodyka* (tekst, tylko do czytania).

2. **Pięć pól wejściowych, każde z walidacją:**

   | Pole | Typ walidacji | Warunek | Komunikat błędu (po polsku) |
   |---|---|---|---|
   | `Data zgłoszenia` | `date` | `greaterThanOrEqual` `=DATE(2026,1,1)` | „Data musi być z 2026 roku lub późniejsza." |
   | `Klient` | `textLength` | `lessThanOrEqual` `60` | „Nazwa klienta może mieć maks. 60 znaków." |
   | `Kwota netto` | `decimal` | `between` `0` – `10_000_000` | „Kwota musi być liczbą z zakresu 0 – 10 000 000." |
   | `Ilość` | `whole` | `greaterThan` `0` | „Ilość musi być liczbą całkowitą większą od zera." |
   | `Status` | `list` | z `Słownik!$D$2:$D$5` (nazwa zdefiniowana `Statusy`) | „Wybierz status z listy." |

3. **Wszystkie pięć pól** ma tło `FFF7E6`, **odblokowane** (`Protection(locked=False)` przez `copy()` — pamiętaj o niemutowalności z modułu 09!).

4. **Trzy formuły** w sekcji *Podsumowanie* — **zablokowane i ukryte** (`Protection(locked=True, hidden=True)` przez `copy()`).

5. **Wszystko pozostałe zablokowane** (domyślne `locked=True`).

6. **Ochrona arkusza włączona z hasłem** — ale **przez jawne `ws.protection.password = HASŁO`** (i skomentuj w kodzie, że to automatycznie włącza ochronę).

7. **Uprawnienia dozwolone:** sortowanie (`sort=False`), filtr (`autoFilter=False`), formatowanie (`formatCells=False`). Wszystko inne zablokowane.

8. **Arkusz `Słownik`** ukryty (`hidden`), z listą statusów w kolumnie `D` i nazwą zdefiniowaną `Statusy`.

9. **Funkcja `inwentarz()`** ma wypisać:
   - stan ochrony arkusza i **wszystkie 15 flag** w formie tabeli z opisem „ZABLOKOWANE / dozwolone",
   - liczbę komórek odblokowanych i ich adresy,
   - liczbę komórek z ukrytą formułą i ich adresy,
   - wszystkie reguły walidacji z pełnymi komunikatami,
   - **ostrzeżenie**, jeśli `sort=True` mimo że chcesz sortowania, albo `hidden=True` na komórce, która nie jest formułą.

10. **Funkcja `test_przenikania()`** — pokazuje, że ochrona **nie jest** bezpieczeństwem:
    - wczytuje plik **bez hasła**,
    - zmienia wartości w zablokowanych komórkach,
    - dodaje arkusz,
    - zapisuje jako `output/17_cw3_obejscie.xlsx`.

W pliku `output/17_cw3_wnioski.md`:

1. **Co dokładnie robi `ws.protection.password = HASŁO`** — odnieś się do źródła `_Protected.set_password` i `SheetProtection.set_password` (2.7). Czy da się ustawić hasło **bez** włączania ochrony?
2. **Dlaczego odblokowanie pól wymagało `copy(cell.protection)`?** Odnieś się do modułu 09 i wyjaśnij, co się dzieje, gdy przypiszesz `Protection(locked=False)` „na goło" do komórki, która miała `hidden=True`.
3. **W którym pliku w archiwum `.xlsx` mieszka stan blokady?** Rozpakuj plik i wklej odpowiedni fragment `styles.xml` oraz `sheet1.xml`. Uzasadnij, dlaczego tak jest (2.8).
4. **Jak wygląda `password` w zapisanym pliku?** Wklej wartość i wyjaśnij, dlaczego nie jest to Twoje hasło.
5. **Uruchom `test_przenikania()`.** Czy udało się zmienić zablokowane komórki? Czy udało się dodać arkusz? Co się stało z `lockStructure`, jeśli je ustawisz? Sformułuj w **jednym zdaniu** prawdę o ochronie w OOXML.
6. **Przekonaj się o sile hasła.** Napisz skrypt, który **bez znajomości hasła** wyliczy, że przestrzeń skrótu ma $2^{16}$ wartości. Nie musisz łamać hasła — wystarczy policzyć i uzasadnić, dlaczego to jest realne w milisekundach. Odnieś się do *„16-bit hash, four hexadecimal digits"* (2.9).
7. **Zaprojektuj audyt.** Co jeszcze musiałbyś sprawdzić, żeby móc powiedzieć klientowi „ten raport jest chroniony przed przypadkową zmianą"? Wypisz **pięć** punktów i wskaż, które z nich openpyxl **nie** potrafi sprawdzić programowo (podpowiedź: eksperyment „przed i po zapisie" z modułu 16).

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Formularz chroniony - kluczowe fragmenty zadania 3."""

from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HASŁO = "Kurs2026!"
AUTOR_BLADU = "Formularz zapotrzebowań"

# ======================================================================
# 1. REJESTRY - jedno miejsce decyzji (modul 09, 10, 15, 16)
# ======================================================================
STYL = {
    "sekcja":  Font(bold=True, size=12, color="1F4E79"),
    "etykieta": Font(bold=True),
    "pole_bg": PatternFill("solid", fgColor="FFF7E6"),
    "wynik_bg": PatternFill("solid", fgColor="EAF3EA"),
    "wazne":   Font(italic=True, size=9, color="7F7F7F"),
}

# pola: (wiersz, etykieta, opis, type, operator, f1, f2, komunikat)
POLA = [
    (4,  "Data zgłoszenia", "RRRR-MM-DD",
     "date", "greaterThanOrEqual", "=DATE(2026,1,1)", None,
     "Data musi być z 2026 roku lub późniejsza."),
    (5,  "Klient", "nazwa firmy",
     "textLength", "lessThanOrEqual", 60, None,
     "Nazwa klienta może mieć maks. 60 znaków."),
    (6,  "Kwota netto", "w zł, bez separatora",
     "decimal", "between", 0, 10_000_000,
     "Kwota musi być liczbą z zakresu 0 – 10 000 000."),
    (7,  "Ilość", "sztuki",
     "whole", "greaterThan", 0, None,
     "Ilość musi być liczbą całkowitą większą od zera."),
    (8,  "Status", "lista",
     "list", None, "Statusy", None,
     "Wybierz status z listy."),
]

STATUSY = ["Nowe", "W toku", "Zamknięte", "Anulowane"]

FORMUŁY = [
    (13, "Wartość brutto",   "=ROUND(B6*1.23,2)"),
    (14, "Zgłoszeń ogółem",  "=COUNTA(B5:B220)"),
    (15, "Kompletność (%)",  "=ROUND(COUNTA(B4:B8)/5*100,0)"),
]


# ======================================================================
# 2. CZYNNOSCI ELEMENTARNE - wzorzec "jedna funkcja = jedna czynność"
# ======================================================================
def odblokuj(cell) -> None:
    """Odblokowuje pole do wpisywania. Kopia, bo Protection jest NIEMUTOWALNE."""
    nowa = copy(cell.protection)
    nowa.locked = False
    cell.protection = nowa
    cell.fill = STYL["pole_bg"]


def zablokuj_i_ukryj(cell) -> None:
    """Formuła: zablokowana (locked=True) I ukryta (hidden=True)."""
    nowa = copy(cell.protection)
    nowa.locked = True
    nowa.hidden = True
    cell.protection = nowa
    cell.fill = STYL["wynik_bg"]


def dodaj_walidacje(ws, pole: str, typ: str, operator, f1, f2,
                    komunikat: str, naglowek_prompt: str) -> None:
    """Buduje regule i - DWA KROKI, KTORYCH NIE WOLNO POMINAC - rejestruje + zakresy."""
    dv = DataValidation(
        type=typ,
        operator=operator,
        formula1=f1,
        formula2=f2,
        allow_blank=True,
        showErrorMessage=True,      # jawnie! (sekcja 2.4 i 2.1)
        # showDropDown NIE ustawiamy - domyslnie False = strzalka widoczna
    )
    dv.errorTitle = f"Błędna wartość: {naglowek_prompt}"
    dv.error = komunikat
    dv.errorStyle = "stop"          # stop = nie da sie zatwierdzic
    dv.promptTitle = naglowek_prompt
    dv.prompt = komunikat

    ws.add_data_validation(dv)      # KROK 2 - bez tego cicho nic nie ma!
    dv.add(pole)                    # KROK 3 - bez tego cicho nic nie ma!


# ======================================================================
# 3. STRUKTURA ARKUSZA
# ======================================================================
def zbuduj_formularz(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Formularz"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 30

    ws["A1"] = "Formularz zapotrzebowania"
    ws["A1"].font = Font(bold=True, size=15)
    ws["A2"] = ("Wypełnij pola na żółtym tle. Pozostałe komórki są chronione "
                "i zawierają ukryte formuły.")
    ws["A2"].font = STYL["wazne"]

    # ---- sekcja: Dane podstawowe ----
    ws["A3"] = "Dane podstawowe"
    ws["A3"].font = STYL["sekcja"]

    for wiersz, etykieta, opis, *_ in POLA:
        ws.cell(row=wiersz, column=1, value=etykieta).font = STYL["etykieta"]
        ws.cell(row=wiersz, column=3, value=opis).font = STYL["wazne"]

    # ---- sekcja: Podsumowanie ----
    ws["A12"] = "Podsumowanie (formuły ukryte)"
    ws["A12"].font = STYL["sekcja"]
    for wiersz, etykieta, formula in FORMUŁY:
        ws.cell(row=wiersz, column=1, value=etykieta).font = STYL["etykieta"]
        ws.cell(row=wiersz, column=2, value=formula)

    # ---- sekcja: Metodyka ----
    ws["A17"] = "Metodyka"
    ws["A17"].font = STYL["sekcja"]
    ws["A18"] = ("Walidacje sprawdzają poprawność wpisywanych danych "
                 "(bramkarz). Ochrona blokuje edycję wszystkich komórek poza "
                 "pięcioma polami (kłódka na szufladzie, nie sejf).")
    ws["A18"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[18].height = 44


def zbuduj_slownik(wb: Workbook) -> str:
    """Slownik + nazwa zdefiniowana. Zwraca zakres do diagnostyki."""
    ws = wb.create_sheet("Słownik")
    ws["D1"] = "Status"
    ws["D1"].font = STYL["etykieta"]
    for i, status in enumerate(STATUSY, start=2):
        ws.cell(row=i, column=4, value=status)

    zakres = f"{quote_sheetname('Słownik')}!$D$2:$D${len(STATUSY) + 1}"
    wb.defined_names.add(DefinedName("Statusy", attr_text=zakres))
    wb.active = wb.sheetnames.index("Formularz")
    ws.sheet_state = "hidden"
    return zakres


# ======================================================================
# 4. ZLOZENIE CALOSCI - KOLEJNOSC MA ZNACZENIE
# ======================================================================
def zbuduj(sciezka: Path) -> None:
    wb = Workbook()
    zakres_statusow = zbuduj_slownik(wb)
    zbuduj_formularz(wb)
    ws = wb["Formularz"]

    # KROK A: walidacje (zanim cokolwiek zablokujemy)
    for wiersz, etykieta, _opis, typ, operator, f1, f2, komunikat in POLA:
        suma_wartosc = f1 if typ == "list" else f1
        dodaj_walidacje(ws, f"B{wiersz}:B220", typ, operator,
                        suma_wartosc, f2, komunikat, etykieta)
    print(f"Walidacja dla 'Status' używa nazwy zdefiniowanej -> {zakres_statusow}")

    # KROK B: odblokuj POLA (biale kratki formularza)
    for wiersz, *_ in POLA:
        odblokuj(ws.cell(row=wiersz, column=2))

    # KROK C: zablokuj i ukryj FORMULY
    for wiersz, _etykieta, _formula in FORMUŁY:
        zablokuj_i_ukryj(ws.cell(row=wiersz, column=2))

    # KROK D: uprawnienia - WSZYSTKO zablokowane poza tym, co potrzebne
    p = ws.protection
    p.sort = False              # sortowanie DOZWOLONE (False = nie blokuj)
    p.autoFilter = False        # strzalki filtra DOZWOLONE
    p.formatCells = False       # formatowanie komorek DOZWOLONE
    p.selectLockedCells = False # mozna zaznaczac zablokowane (da sie kopiowac)
    p.selectUnlockedCells = False
    # reszta flag zostaje na domyslnych True = ZABLOKOWANE

    # KROK E: PW - ustawienie hasla AUTOMATYCZNIE wlacza ochrone (2.7)
    p.password = HASŁO
    assert p.sheet is True, "Ustawienie hasla powinno włączyć ochronę!"

    wb.save(sciezka)
    wb.close()


# ======================================================================
# 5. INWENTARZ I TEST PRZENIKANIA
# ======================================================================
FLAGI_OPIS = {
    "objects": "edycja obiektów (wykresy, obrazy)",
    "scenarios": "edycja scenariuszy",
    "formatCells": "formatowanie komórek",
    "formatColumns": "szerokość/ukrywanie kolumn",
    "formatRows": "wysokość/ukrywanie wierszy",
    "insertColumns": "wstawianie kolumn",
    "insertRows": "wstawianie wierszy",
    "insertHyperlinks": "wstawianie hiperlinków",
    "deleteColumns": "usuwanie kolumn",
    "deleteRows": "usuwanie wierszy",
    "selectLockedCells": "zaznaczanie zablokowanych komórek",
    "selectUnlockedCells": "zaznaczanie odblokowanych komórek",
    "sort": "sortowanie",
    "autoFilter": "używanie strzałek filtra",
    "pivotTables": "edycja tabel przestawnych",
}


def inwentarz(sciezka: Path) -> None:
    wb = load_workbook(sciezka)          # NIE read_only! (2.11)
    try:
        ws = wb["Formularz"]
        p = ws.protection
        print("=" * 92)
        print(f"OCHRONA arkusza '{ws.title}': sheet={p.sheet}   "
              f"hasz hasła={p.password!r}")
        print("=" * 92)
        print(f"{'flaga':<22}{'stan':<14}co blokuje")
        print("-" * 92)
        for nazwa, opis in FLAGI_OPIS.items():
            wartosc = getattr(p, nazwa)
            stan = "ZABLOKOWANE" if wartosc is True else "dozwolone"
            print(f"{nazwa:<22}{stan:<14}{opis}")

        # --- komorki: odblokowane i ukryte formuly ---
        odblokowane, ukryte = [], []
        for row in ws.iter_rows():
            for c in row:
                if c.protection is None:
                    continue
                if c.protection.locked is False:
                    odblokowane.append(c.coordinate)
                if c.protection.hidden:
                    # kontrola spojnosci: hidden bez formuly = prawdopodobnie blad
                    if not (isinstance(c.value, str) and c.value.startswith("=")):
                        print(f"  ⚠️  {c.coordinate}: hidden=True, ale to NIE "
                              f"formuła (wartość={c.value!r})")
                    ukryte.append(c.coordinate)
        print()
        print(f"Pola ODBLOKOWANE ({len(odblokowane)}): {odblokowane}")
        print(f"Formuły UKRYTE   ({len(ukryte)}): {ukryte}")

        # --- walidacje z pelnymi komunikatami ---
        print()
        print(f"Walidacje ({len(ws.data_validations)}):")
        for i, dv in enumerate(ws.data_validations.dataValidation, 1):
            print(f"  [{i}] {dv.type:<10} {str(dv.operator):<18} "
                  f"f1={str(dv.formula1)[:26]:<28} sqref={dv.sqref}")
            print(f"       error='{dv.error}'  style={dv.errorStyle}")
    finally:
        wb.close()


def test_przenikania(sciezka: Path, cel: Path) -> None:
    """Dowod, ze ochrona nie jest bezpieczenstwem (sekcja 2.9)."""
    print()
    print("=" * 92)
    print("TEST PRZENIKANIA: openpyxl czyta i modyfikuje BEZ HASLA")
    print("=" * 92)

    wb = load_workbook(sciezka)
    try:
        ws = wb["Formularz"]
        print(f"  hasło w pliku: {ws.protection.password!r}  "
              f"(to NIE jest 'Kurs2026!')")
        ws["B13"] = "=1/0"                    # podmiana ukrytej formuly
        ws["A1"] = "ZNISZCZONE"
        wb.create_sheet("Hack")               # nowy arkusz
        wb.save(cel)
        print(f"  ✓ Zmieniono ukrytą formułę (B13), nagłówek (A1)")
        print(f"  ✓ Dodano arkusz 'Hack'")
        print(f"  ✓ Zapisano jako {cel.name}")
    finally:
        wb.close()

    wb2 = load_workbook(sciezka)
    wb3 = load_workbook(cel)
    try:
        print(f"  ochrona w oryginale: {wb2['Formularz'].protection.sheet}")
        print(f"  ochrona po obejściu:  {wb3['Formularz'].protection.sheet}")
        print(f"  arkusze oryginału: {wb2.sheetnames}")
        print(f"  arkusze po obejściu: {wb3.sheetnames}")
    finally:
        wb2.close()
        wb3.close()

    print()
    print("  WNIOSEK: ochrona to instrukcja dla INTERFEJSU Excela, nie")
    print("  właściwość danych. openpyxl nie egzekwuje jej ani nie pyta")
    print("  o hasło. OOXML: 'cannot protect it from malicious modification'.")
```

**Kluczowe decyzje i uzasadnienia:**

- **`copy()` w obu funkcjach pomocniczych.** To bezpośrednie zastosowanie modułu 09. `Protection` to **część stylu**, więc podlega tej samej niemutowalności co `Font`. Bez `copy()` przypisanie „na goło" `Protection(locked=False)` **wygubiłoby** wcześniejszy `hidden=True` — bo `Protection(locked=False)` ma `hidden=False` domyślnie. To subtelna cisza i drugi najczęstszy błąd w tym module.

- **Trzy kroki walidacji zamknięte w jednej funkcji.** `dodaj_walidacje` przyjmuje gotowe parametry i **robi** rejestrację oraz przypisanie zakresu. W ten sposób **nie da się** pominąć `add_data_validation` ani `dv.add`. To ta sama filozofia co `dodaj_link_z_napisem` z modułu 16 — **interfejs funkcji wymusza poprawną kolejność**.

- **`showDropDown` nie pojawia się ani raz.** To celowa decyzja: domyślne `False` daje widoczną strzałkę, a jawny parametr tylko wprowadzałby nieporozumienie. W komentarzu jest to wyjaśnione (2.4).

- **`p.password = HASŁO` jako ostatnia czynność.** Kolejność: walidacje → odblokowanie pól → ukrycie formuł → uprawnienia → hasło/ochrona. Ta kolejność odwzorowuje sekwencję myślową z 2.13 i sprawia, że kod czyta się jak instrukcja obsługi. **Mechanicznie w openpyxl kolejność nie jest wymuszona** (stan komórek i flaga ochrony serializują się niezależnie przy zapisie), ale **jest wymuszona dydaktycznie** — i przy każdej późniejszej modyfikacji kodu jest to porządek, który pozwala szybko znaleźć problem.

- **`assert p.sheet is True`.** Jedna linia, która **dokumentuje** nieoczywiste zachowanie `set_password` → `enable()` (2.7). Asercje jako dokumentacja wykonawcza to wzorzec, który w module 20 zamienisz na pełne testy.

- **`test_przenikania` jako część modułu, nie przypis.** To jest **najważniejsza funkcja tego zadania**. Bez niej student wyniósłby z modułu umiejętność „ustawiania ochrony", a nie **zrozumienie, czym ochrona jest**. Test pokazuje wprost: podmiana ukrytej formuły, zniszczony nagłówek, dodatkowy arkusz — wszystko **bez hasła**, wszystko zapisane, a flaga `sheet=True` nadal `True` w obu plikach.

- **Blokada struktury.** Jeśli chcesz dodać `wb.security.lockStructure = True`, pamiętaj o warunku z 2.10: **bez `workbookPassword` to nie jest egzekwowane**. W kodzie rozwiązania pominięto to celowo — pole do refleksji w punkcie 5 wniosków.

- **Co tu jest antywzorcem?** W rozwiązaniu widać dwie rzeczy, które w module 24 poprawisz: (a) funkcje przyjmują `ws` (typ openpyxl) w sygnaturach — to przeciek abstrakcji; poprawisz to na własny obiekt raportu; (b) reguły walidacji są zapisane jako krotki w `POLA`, a nie jako obiekty `Specyfikacja` — poprawisz to w module 25, gdy zobaczysz, jak z tej samej struktury wygenerować walidacje, komunikaty i dokumentację.

</details>

## 6. Typowe błędy i pułapki

**1. „Lista rozwijana nie działa — nie ma strzałki" (objaw) → **`showDropDown=True` UKRYWA strzałkę.** To alias `hide_drop_down`; OOXML ma odwrócone znaczenie tego atrybutu (2.4) (przyczyna) → **Nie ustawiaj `showDropDown` wcale** (domyślne `False` = widowczna strzałka). Jeśli już musisz — `showDropDown=False`. Zajrzyj też do XML-a: `showDropDown="1"` to Twój problem (naprawa).**

**2. „Utworzyłem `DataValidation`, ustawiłem komunikaty — a w pliku nie ma walidacji" (objaw) → **pominięty `ws.add_data_validation(dv)`.** `dv.add("B2:B100")` zadziała bez błędu, ale reguła nigdy nie trafi do kolekcji arkusza (2.2) (przyczyna) → **Zawsze trzy kroki: zbuduj → `ws.add_data_validation(dv)` → `dv.add(zakres)`.** Zamknij to w funkcji pomocniczej, żeby nie dało się pominąć (Przykład 3, szkic zadania 3) (naprawa).**

**3. „Walifacja dodana, komunikaty ustawione, `add_data_validation` wywołane — a w pliku nadal nic" (objaw) → **pominięty `dv.add(...)`.** Dokumentacja mówi wprost: *„Validations without any cell ranges will be ignored when saving a workbook"*, a `DataValidationList.to_tree` filtruje `if bool(r.sqref)` (2.2) (przyczyna) → **Wywołaj `dv.add(...)` choć raz.** Sprawdź `if not dv.sqref: raise ValueError(...)` przed zapisem (naprawa).**

**4. „Excel pokazuje dziwną/uciętą listę" (objaw) → **lista wpisana wprost przekracza 255 znaków** (limit Excela, nie openpyxl). openpyxl **nie ostrzega** — XlsxWriter ostrzega, ale tutaj zapis przejdzie cicho (2.5) (przyczyna) → **Dla długich list zawsze zakres komórek albo nazwa zdefiniowana.** Policz `len(",".join(opcje))` i porównaj z 255 **przed** zapisem (naprawa).**

**5. „Lista dzieli wartości tam, gdzie nie powinna / nie dzieli, gdzie powinna" (objaw) → **przecinek jako separator wpisany w wartości albo brak cudzysłowów wokół całości.** Trzy postacie `formula1` mylą się nagminnie (2.3) (przyczyna) → **Zapamiętaj: lista wprost = `'"a,b,c"'`, zakres = `"'Arkusz'!$A$1:$A$10"`, nazwa = `"NazwaZdefiniowana"`.** Sprawdź swój napis w `print(repr(dv.formula1))` (naprawa).**

**6. „Po zapisaniu pliku przez openpyxl zniknęły wszystkie listy rozwijane" (objaw) → **plik miał walidację w bloku `x14`** (typowo lista z innego arkusza; tak zapisuje Excel 2010+) **i openpyxl ją zignorował.** Towarzyszy temu ostrzeżenie `Data Validation extension is not supported and will be removed`, które przeoczyłeś (2.12) (przyczyna) → **Wykryj ostrzeżenie** przez `warnings.catch_warnings(record=True)` + `simplefilter("always")`. **Odbuduj reguły** z XML-a **albo** przejdź w szablonie na **nazwy zdefiniowane** (naprawa).**

**7. „Wpisałem zablokowanie hasłem, ale użytkownik zdjął ochronę jednym kliknięciem" (objaw) → **ochrona arkusza bez hasła.** Dokumentacja: *„If no password is specified, users can disable configured sheet protection without specifying a password"* (2.6) (przyczyna) → **Zawsze ustaw `ws.protection.password`.** Pamiętaj, że ustawienie hasła **automatycznie włącza ochronę** (2.7) (naprawa).**

**8. „Włączyłem ochronę i nikt nie może nic wpisać — nawet w polach formularza" (objaw) → **wszystkie komórki są domyślnie zablokowane**, a Ty nie odblokowałeś pól. Wzorzec nie brzmi „zablokuj formuły", a „odblokuj wejścia" (1.3) (przyczyna) → **`Protection(locked=False)` na polach do wpisywania, najlepiej przez `copy(cell.protection)`** — a **potem** włącz ochronę. Sprawdź, że `odblokowane` w inwentarzu nie są puste (naprawa).**

**9. „Odblokowałem pole, ale straciło ono ukrycie formuły / własne ustawienie" (objaw) → **`cell.protection = Protection(locked=False)` nadpisuje CAŁY składnik ochrony**, a `Protection()` ma `hidden=False` domyślnie (2.8) (przyczyna) → **`copy(cell.protection)` → zmień jedno pole → przypisz z powrotem.** Dotyczy **każdego** składnika stylu (moduł 09) (naprawa).**

**10. „Ustawiłem hasło, żeby je potem skonfigurować — i ochrona włączyła się, zanim byłem gotowy" (objaw) → **`ws.protection.password = x` wywołuje `set_password` → `enable()`.** Ustawienie hasła i włączenie ochrony to **jedna operacja** (2.7) (przyczyna) → **Ustawiaj hasło jako OSTATNIĄ czynność** względem struktury arkusza. Jeśli chcesz mieć hasło bez ochrony — to nie jest możliwe przez `password`; użyj `set_password(..., already_hashed=True)` i **pamiętaj, że też włączy** (naprawa).**

**11. „Ustawiłem `sort=False` i `autoFilter=False`, a użytkownik i tak nie może filtrować/sortować" (objaw) → **Excel nakłada dodatkowe warunki, których flagi nie obejmują:** *„Users can't sort ranges that contain locked cells on a protected worksheet, regardless of this setting"* i *„Users cannot apply or remove AutoFilters on a protected worksheet"* (2.6) (przyczyna) → **Zostaw zakres danych odblokowany** i **założ filtr przed** włączeniem ochrony (najlepiej jako Tabelę — moduł 12). Flaga daje **uprawnienie**, nie gwarancję (naprawa).**

**12. „Ukryłem formułę, ale każdy ją widzi" (objaw) → **`hidden=True` działa TYLKO na chronionym arkuszu**, a poza tym formuła i tak jest czytelna w `sheet1.xml` (2.8) (przyczyna) → **`hidden=True` + `sheet=True` w parze.** I nie traktuj tego jako ochrony danych — to wyłącznie kosmetyka paska formuły. W każdej niechronionej kopii formuła będzie jawna (naprawa).**

**13. „Chcę WYMUSIĆ, żeby nikt nie zmienił danych w pliku" (objaw) → **nie istnieje mechanizm, który to robi w OOXML dla danych.** Ochrona to *„instruction to Excel's UI, not a property of the data"*; openpyxl **nie egzekwuje** jej przy czytaniu i zapisie, a `pandas`/import do bazy widzą wszystko (2.9) (przyczyna) → **Szyfruj plik (hasło do otwarcia), kontroluj uprawnienia w systemie plików, albo nie umieszczaj danych w pliku.** Uświadom to sobie **przed** obiecywaniem klientowi „zabezpieczonego raportu" (naprawa).**

**14. „Zbudowałem narzędzie diagnostyczne i wyszło, że plik nie ma żadnych walidacji" (objaw) → **użyłeś `load_workbook(..., read_only=True)`.** Czytnik robi wtedy `continue` przed parsowaniem arkusza, więc walidacji, ochrony, obrazów i reguł warunkowych **nie ma w modelu** (2.11) (przyczyna) → **Do diagnostyki używaj trybu normalnego.** Zaakceptuj koszt pamięci. W module 19 poznasz trzecią drogę: parsowanie XML-a bezpośrednio (naprawa).**

**15. „Utworzyłem 60 000 drobnych walidacji (po jednej na komórkę) i plik przestał działać / waży gigabajty" (objaw) → **jedna reguła na komórkę mnoży blok XML i strukturę modelu** w pamięci (2.1) (przyczyna) → **Jedna reguła na kolumnę** — `dv.add("B2:B100000")`. Jeśli reguła ma zależeć od innego pola w tym samym wierszu, użyj `custom` z `ROW()`/`COLUMN()`/`INDIRECT()` w formule zamiast tworzyć regułę na wiersz (naprawa).**

**16. „`wb.security.lockStructure = True` ustawione, a struktura nadal jest edytowalna" (objaw) → **bez `workbookPassword` dokumentacja mówi jawnie: ograniczenia *„will only be enforced if the appropriate password is set"* (2.10) (przyczyna) → **Ustaw oba:** `wb.security.workbookPassword = "..."` **oraz** `wb.security.lockStructure = True`. I pamiętaj, że i tak nie jest to wymuszane poza Excelem (naprawa).**

**17. „Ostrzeżenia znikają — widzę tylko pierwsze" (objaw) → **domyślny filtr Pythona pokazuje każde unikalne ostrzeżenie tylko raz na sesję**. Przy pięciu arkuszach z `x14` zobaczysz jedno ostrzeżenie (Przykład 4) (przyczyna) → **`warnings.catch_warnings(record=True)` + `warnings.simplefilter("always")`.** Zbieraj ostrzeżenia i wypisuj jako raport ryzyka (naprawa).**

**18. „Ukryłem arkusz słownika i lista rozwijana pokazuje mi puste wartości" (objaw) → **lista odwołuje się do zakresu na arkuszu `veryHidden`** albo odwołanie obejmuje niewłaściwy zakres; Excel ma historyczne ograniczenia dla arkuszy o najgłębszym ukryciu (przyczyna) → **Użyj `sheet_state = "hidden"`, nie `"veryHidden"`, dla arkusza ze słownikiem.** Sprawdź odwołanie w `print(repr(dv.formula1))` i obejrzyj plik w Excelu. W razie problemu — nazwa zdefiniowana zamiast odwołania bezpośredniego (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Walidacja to bramkarz na wejściu głównym, nie zamek w drzwiach.** Excel sprawdza wartość **tylko wtedy**, gdy wpisuje ją człowiek przez swój interfejs. openpyxl, pandas, import do bazy — wchodzą od kuchni. **Konsekwencja:** walidacja to **UX**, nie kontrola integralności. Logika biznesowa musi być walidowana w Pythonie, a walidacja w Excelu jest tylko podpowiedzią dla użytkownika. Trzy kroki zawsze razem: **zbuduj → `add_data_validation` → `dv.add(...)`**; brak któregokolwiek jest cichy.

2. **Trzy postacie `formula1` i jeden odwrócony parametr.** Lista wprost wymaga **cudzysłowów w napisie** i ma limit **255 znaków** (limit Excela, nie openpyxl — i openpyxl o nim **nie ostrzega**). Zakres wymaga **`quote_sheetname`** i idzie bez `=`. Nazwa zdefiniowana jest **najodporniejsza**. A `showDropDown=True` **ukrywa** strzałkę — to najczęstsza przyczyna „lista nie działa".

3. **Ochrona to 15 przełączników z odwróconą logiką, a stan blokady komórki MIESZKA W STYLU.** `True` znaczy **„zablokowane"**, nie „dozwolone". Wszystkie komórki są domyślnie zablokowane, więc wzorzec brzmi **„odblokuj pola, potem włącz ochronę"** — nigdy „zablokuj formuły". A `Protection` to **szósty składnik stylu**, więc podlega niemutowalności z modułu 09: `copy()` → zmień → przypisz. `hidden=True` działa **tylko** na chronionym arkuszu i **nie chroni danych** — jest zamalowanym wzorem, nie sejfem.

4. **Dwa haczyki, które trzeba znać na pamięć: hasło = ochrona włączona, i `lockStructure` bez hasła = brak egzekwowania.** `ws.protection.password = x` wywołuje `set_password` → `enable()` — nie da się ustawić hasła „na później". `wb.security.lockStructure = True` **bez `workbookPassword`** nie robi nic, bo nie ma czego wymagać od użytkownika. Ustawiaj oba.

5. **To nie jest bezpieczeństwo — i musisz to powiedzieć głośno.** Skrót hasła to **16 bitów / 4 znaki hex / 65 536 możliwości**; wiele haseł odblokuje ten sam plik. **openpyxl nie egzekwuje ochrony** przy czytaniu: doda arkusz do skoroszytu z `lockStructure=True`, podmieni ukrytą formułę i zapisze — **bez hasła**. Specyfikacja: *„cannot protect it from malicious modification"*. Kłódka na szufladzie biurka, nie sejf. Prawdziwa ochrona to szyfrowanie pliku albo uprawnienia w systemie.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ==================================================================
# 2. WALIDACJA - TRZY KROKI, KTORYCH NIE WOLNO POMINAC
# ==================================================================
dv = DataValidation(
    type="list",                    # list|whole|decimal|date|time|textLength|custom
    operator=None,                  # between|notBetween|equal|notEqual|lessThan|
                                    # lessThanOrEqual|greaterThan|greaterThanOrEqual
    formula1='"Tak,Nie,Nie wiem"',  # patrz sekcja 3 - TRZY POSTACIE
    formula2=None,                  # druga wartosc dla 'between'
    allow_blank=True,               # allow_blank WYGRYWA z allowBlank
    showErrorMessage=True,          # ustaw JAWNIE (domyslne False!)
    showInputMessage=True,          # bez tego 'prompt' nie zadziala
    # showDropDown NIE USTAWIAJ! True = UKRYWA strzalke (2.4)
)
dv.errorTitle = "Nieprawidłowa wartość"
dv.error = "Dozwolone są tylko: Tak, Nie, Nie wiem."
dv.errorStyle = "stop"              # stop (blokuje) | warning | information
dv.promptTitle = "Wybór"
dv.prompt = "Wybierz jedną z trzech wartości."

ws.add_data_validation(dv)          # KROK 2: REJESTRACJA (bez tego: CICHO NIC)
dv.add("B2:B100")                   # KROK 3: ZAKRESY  (bez tego: CICHO NIC)
dv.add("D2:D100")                   # dokladanie kolejnych zakresow

"B50" in dv                         # True - __contains__ dziala
len(ws.data_validations)            # liczba regul
for r in dv.sqref: print(r)         # iteracja po zakresach (MultiCellRange)

# ==================================================================
# 3. TRZY POSTACIE formula1
# ==================================================================
# (a) LISTA WPROST - cudzysloby SA CZESCIA napisu; limit 255 znakow!
opcje = ["Tak", "Nie", "Nie wiem"]
f1_a = '"' + ",".join(opcje) + '"'          # -> "Tak,Nie,Nie wiem"
assert len(f1_a) <= 255, f"Lista za dluga: {len(f1_a)} znakow!"

# (b) ZAKRES - bez cudzyslowow, bez '='; CYTOWANIE nazwy arkusza!
f1_b = "{0}!$B$1:$B$10".format(quote_sheetname("Dane 2026"))   # 'Dane 2026'!$B$1:$B$10

# (c) NAZWA ZDEFINIOWANA - najbardziej odporna (2.12)
wb.defined_names.add(DefinedName("Statusy", attr_text="'Słownik'!$D$2:$D$5"))
f1_c = "Statusy"

# (d) typ 'custom' - formula z '=' (dokumentacja openpyxl)
#     formula1="=ISNUMBER(J2)"

# ==================================================================
# 4. OCHRONA ARKUSZA - 15 FLAG (True = ZABLOKOWANE!)
# ==================================================================
ws.protection.sheet = True          # glowny wylacznik
ws.protection.enable()              # to samo
ws.protection.disable()             # sheet = False

# uprawnienia: False = DOZWOLONE
ws.protection.sort = False          # ...ale locked cells i tak blokuja sortowanie!
ws.protection.autoFilter = False    # ...filtr musi byc zalozony PRZED ochrona
ws.protection.formatCells = False
ws.protection.selectLockedCells = False
ws.protection.selectUnlockedCells = False
# domyslnie ZABLOKOWANE: objects, scenarios, formatColumns, formatRows,
# insertColumns, insertRows, insertHyperlinks, deleteColumns, deleteRows,
# pivotTables

# ⚠️ USTAWIENIE HASLA AUTOMATYCZNIE WŁĄCZA OCHRONĘ (set_password -> enable())!
ws.protection.password = "haslo"
assert ws.protection.sheet is True
bool(ws.protection)                 # True jesli arkusz chroniony

# ==================================================================
# 5. STAN BLOKADY KOMORKI - CZESC STYLU (niemutowalne! modul 09)
# ==================================================================
# ZLE: cell.protection.locked = False      <- nie zadziala
# ZLE: cell.protection = Protection(locked=False)  <- GUBI 'hidden'!

# DOBRZE:
p = copy(cell.protection)
p.locked = False            # odblokuj do wpisywania
cell.protection = p

p = copy(cell.protection)
p.locked = True
p.hidden = True             # ukryj formule w pasku (dziala TYLKO z sheet=True)
cell.protection = p

# styl pola: zeby uzytkownik WIDZIAL, gdzie wpisywac
cell.fill = PatternFill("solid", fgColor="FFF7E6")

# ==================================================================
# 6. OCHRONA SKOROSZYTU (2.10)
# ==================================================================
wb.security.workbookPassword = "haslo"      # BEZ TEGO lockStructure nic nie robi!
wb.security.lockStructure = True            # bez dodawania/usuwania/przenoszenia
wb.security.revisionsPassword = "haslo2"    # legacy - zwykle pomijaj
# wb.security.lockWindows  -> zostaw None (nowszy Excel juz tego nie ma)
# wb.security.lockRevision -> legacy - pomijaj

# ==================================================================
# 7. ODCZYT WALIDACJI I OCHRONY (diagnostyka!)
# ==================================================================
# ⚠️ NIE read_only! Tam walidacji i ochrony NIE MA (2.11)
import warnings

with warnings.catch_warnings(record=True) as zlapane:
    warnings.simplefilter("always")          # bez tego tylko PIERWSZE ostrzezenie
    wb = load_workbook(path)
    try:
        for ws in wb.worksheets:
            print(ws.title, "ochrona:", ws.protection.sheet,
                  "hasz:", ws.protection.password)

            # komorki odblokowane i ukryte formuly
            for row in ws.iter_rows():
                for c in row:
                    if c.protection and c.protection.locked is False:
                        print("  odblokowana:", c.coordinate)
                    if c.protection and c.protection.hidden:
                        print("  ukryta formula:", c.coordinate)

            for dv in ws.data_validations.dataValidation:
                print(dv.type, dv.operator, repr(dv.formula1), str(dv.sqref))
    finally:
        wb.close()

for w in zlapane:
    print("⚠️", w.message)
# "Data Validation extension is not supported and will be removed"
#   -> plik ma walidacje w bloku x14; openpyxl ich NIE przeniesie (2.12)

# obejscie 1: nazwy zdefiniowane zamiast odwolan do arkusza
# obejscie 2: odczyt x14 z XML-a i odbudowa regul
#   zipfile.ZipFile(path).read("xl/worksheets/sheet1.xml")
#   szukaj: <x14:dataValidation ... type="..." ...><xm:f>...</xm:f>
#                                          <xm:sqref>...</xm:sqref>

# ==================================================================
# 8. HAMOWANIE OSTRZEZENIA (tylko gdy wiesz, ze czytasz bez zapisu)
# ==================================================================
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported",
    category=UserWarning,
    module="openpyxl",
)

# ==================================================================
# 9. TYPY WALIDACJI - SCIAGA
# ==================================================================
# list        formula1='"' + ','.join(v) + '"'  LUB  "'Ark'!$A$1:$A$10"  LUB  "Nazwa"
# whole       operator=..., formula1=int, formula2=int (dla between)
# decimal     operator=..., formula1=float, formula2=float
# date        operator=..., formula1="=DATE(2026,1,1)"
# time        operator=..., formula1="=TIME(18,0,0)"
# textLength  operator="lessThanOrEqual", formula1=200
# custom      formula1="=ISNUMBER(A1)"
# operator:   between | notBetween | equal | notEqual | lessThan |
#             lessThanOrEqual | greaterThan | greaterThanOrEqual

# ==================================================================
# 10. TRZY ZDANIA DO ZAPAMIETANIA
# ==================================================================
# 1) Walidacja to UX, nie bezpieczeństwo - kazdy skrypt ja ominie.
# 2) Ochrona to kłódka na szufladzie, nie sejf - 16-bit hash, 65 536 wartosci,
#    openpyxl nie egzekwuje jej ani przy czytaniu, ani przy zapisie.
# 3) hidden=True dziala TYLKO z sheet=True i NIE chroni danych - formuly
#    sa jawne w sheet1.xml.
```

## 9. Co dalej

Za nami dwa moduły o „rzeczach, które nie są danymi" — najpierw nakładki (16), teraz reguły i zamki (17). I w obu powtórzyła się ta sama lekcja w innym przebraniu:

> **Bliżej pliku (XML, `styles.xml`, `sqref`, bloki `x14`) rozumiesz więcej niż przez samo API.**

W module 16 odkryłeś, że obraz to kopia bajtów, a nie adres — bo zajrzałeś do `xl/media/`. W module 17 odkryłeś, że stan blokady komórki to **styl**, a lista z innego arkusza to **rozszerzenie** — bo zajrzałeś do `styles.xml` i `extLst`. **Za każdym razem odpowiedź była w archiwum ZIP, nie w dokumentacji funkcji.** To nie przypadek; to jest właściwy model mentalny tej biblioteki i wracaj do niego, gdy coś „nie działa".

Trzy rzeczy, które zabierasz ze sobą:

- **Dyscyplina trzech kroków.** Walidacja ma **trzy** kroki i **żaden** nie jest opcjonalny. Dwa z nich kończą się ciszą, gdy je pominiesz. Dlatego **zamykaj je w funkcji** — interfejs, którego nie da się użyć źle, jest lepszy niż komentarz „pamiętaj o `add_data_validation`".
- **Ochrona jako UX, nie jako zabezpieczenie.** Umiesz teraz ustawić 15 flag, włączyć ochronę struktury i ukryć formuły. Ale umiesz też **powiedzieć, czego to nie robi** — a to jest różnica między programistą a inżynierem. Specyfikacja OOXML wprost: *„cannot protect it from malicious modification"*. Zapamiętaj to zdanie; w module 21 stanie się ono osią całego modułu bezpieczeństwa.
- **Nasłuch ostrzeżeń jako narzędzie inżynierskie.** `warnings.simplefilter("always")` to nie ciekawostka. To jedyny kanał, przez który openpyxl mówi Ci „tracę dane". W module 16 był to komentarz w scalonej komórce; tutaj — cały blok `x14`. **W module 18 zamienimy ten nasłuch w pełny, automatyczny raport ryzyka pliku wejściowego.**

**Zadanie do zrobienia teraz, zanim ruszysz dalej.** Wykonaj pełny **cykl życia szablonu** — to doświadczenie, które przygotuje Cię na moduł 18 lepiej niż jakikolwiek opis:

1. Uruchom `examples/17_formularz.py`. Otwórz plik w Excelu i sprawdź, że listy rozwijane działają, a pola są odblokowane.
2. **Zapisz plik z Excela** pod nową nazwą i zamknij.
3. Uruchom `examples/17_inwentarz.py` na oryginale i na wersji „po Excelu". **Porównaj liczbę reguł walidacji.**
4. Uruchom `test_przenikania()` z Ćwiczenia 3 — otwórz plik bez hasła, zmień ukrytą formułę, dodaj arkusz, zapisz.
5. Wczytaj plik po „przebudowie przez openpyxl" i uruchom inwentarz **po raz trzeci**. Policz, ile reguł zostało w każdym z trzech plików.

Zobaczysz na własne oczy trzy etapy tej samej utraty danych: **Excel → `x14` → openpyxl → nic.** A jednocześnie zobaczysz, że flaga `sheet=True` przetrwała wszystkie trzy pliki i nadal wygląda na „chronioną" — mimo że skrypt przeszedł przez nią jak przez otwarte drzwi.

W **module 18** wchodzimy w **najważniejszy temat całego kursu: modyfikację istniejących plików bez utraty zawartości.** Zbierzemy wszystko, co do tej pory ostrzegało o utracie danych — kształty, sparkline'y, slicery, makra, `x14`, komentarze w scaleniach, uproszczone wykresy — w **jeden inwentarz**. Zobaczysz wzorzec **„scan-before-save"**, **atomowy zapis** (`os.replace`), **kopie z timestampem**, **idempotencję** i **test regresji formatowania**. A nad tym wszystkim stanie jedna reguła kciuka, którą powtórzymy trzy razy: **jeśli plik był tworzony ręcznie przez człowieka i jest bogaty wizualnie — nie przepuszczaj go przez openpyxl.**