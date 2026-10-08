# Moduł 04 — Pierwszy skoroszyt

> **Część:** I — Podstawy pracy ze skoroszytem · **Poziom:** ⭐ · **Wymaga:** modułów 00–03

## 0. W tym module nauczysz się

- Piszesz **pierwszy kompletny program**: utworzenie skoroszytu w pamięci, wpisanie danych, zapis na dysk, ponowne wczytanie i odczyt. Cały cykl w jednym pliku, który uruchomisz bez żadnej wiedzy dodatkowej.
- Rozumiesz, dlaczego `Workbook()` **zawsze** daje Ci jeden domyślny arkusz o nazwie `Sheet` — i dlaczego to nie jest uprzejmość biblioteki, a wymóg formatu pliku.
- Rozumiesz dokładnie, co się dzieje z istniejącym plikiem, gdy wywołasz `wb.save("ten_sam_plik.xlsx")` — i dlaczego to zachowanie nazywa się „Zapisz", a nie „Zapisz jako".
- Umiesz czytać istniejące skoroszyty: `load_workbook()`, `wb.sheetnames`, `wb["Nazwa"]`, `wb.active`, `ws.title` — oraz odróżniasz sytuację, w której dostajesz **tekst** (nazwę), od sytuacji, w której dostajesz **obiekt** (arkusz).
- Zarządzasz strukturą skoroszytu w podstawowym zakresie: `wb.create_sheet()`, `del wb["Nazwa"]`, `wb.remove(ws)`, ustawianie aktywnego arkusza. Pełne zasady nazewnictwa, kopiowania i widoczności przyjdą w module 08.
- Budujesz **jawne ścieżki** przez `pathlib.Path` i rozumiesz, dlaczego ścieżka względna to bomba zegarowa. Tworzysz katalog `output/` samodzielnie, zanim cokolwiek zapiszesz.
- Zapisujesz skoroszyt do **pamięci** (`BytesIO`), nie na dysk — fundament pod testy (moduł 20) i odpowiedzi HTTP (moduł 26).
- Ustawiasz **właściwości dokumentu** (`wb.properties`) i rozumiesz, po co w firmie wpisuje się tam autora, tytuł i datę — i co z tego czyta audytor.

## 1. Intuicja i analogia

Masz już za sobą cztery moduły teorii: wiesz, że `.xlsx` to archiwum ZIP z XML-em, wiesz, że openpyxl buduje graf obiektów w pamięci i przepisuje cały pakiet od nowa. Czas na pierwszy **działający program**. Ten moduł jest krótki koncepcyjnie i długi praktycznie — bo zawiera dokładnie te elementy, które później, w każdym kolejnym module, będziesz powtarzać w każdej wersji swojego kodu.

### Analogia pierwsza: „Zapisz" znaczy „Zapisz", nie „Zapisz jako"

Otwórz w głowie edytor tekstu. Piszesz notatkę. Naciskasz `Ctrl+S`. Jeśli dokument był już zapisany na dysku pod nazwą `notatka.txt`, edytor **nie pyta Cię o nic** — po prostu zastępuje plik nową treścią. To jest „Zapisz". „Zapisz jako" to inny klawisz: ten pyta o nazwę.

`wb.save("sciezka")` działa jak „Zapisz". Jeśli plik pod tą ścieżką istnieje:

- nie dostaniesz ostrzeżenia,
- nie dostaniesz pytania o potwierdzenie,
- nie powstanie żadna kopia zapasowa,
- stara zawartość **przestaje istnieć** w momencie, gdy zaczyna się nowy zapis.

I ważne: to nie jest „zła wola" openpyxl ani cecha biblioteki. Tak działa zapis plików w każdym języku programowania, jeśli otworzysz plik w trybie zapisu. Python robi dokładnie to samo: `open("notatka.txt", "w")` **obcina plik do zera bajtów** jeszcze przed wpisaniem pierwszej litery.

To ma jeszcze jedną konsekwencję, o której łatwo zapomnieć: **jeśli zapis się nie uda w połowie** (zabraknie miejsca na dysku, proces zostanie zabity, przerwiesz `Ctrl+C`), plik nie wróci do poprzedniego stanu. Będzie **obcięty albo niekompletny**. Dlatego w module 00 przyjęliśmy zasadę kursu — nigdy nie nadpisujemy pliku źródłowego — a w tym module zbudujemy pierwszą wersję funkcji, która tej zasady pilnuje za Ciebie.

### Analogia druga: ścieżka względna to „gdzie akurat stoję"

To najczęstsza przyczyna sytuacji „u mnie działa, a na serwerze nie".

Wyobraź sobie, że mówisz komuś: **„podaj mi to z szuflady"**. Jeśli ta osoba stoi w Twoim biurku — wie, o co chodzi. Jeśli stoi w innym pokoju, pójdzie szukać „szuflady" gdzie indziej i albo nic nie znajdzie, albo znajdzie cudzą szufladę.

Ścieżka względna (`"output/raport.xlsx"`) znaczy dokładnie to: **„szuflada w pokoju, w którym teraz stoisz"**. A „pokój, w którym stoisz" to w Pythonie **katalog roboczy procesu** (ang. *current working directory*, CWD) — miejsce, z którego **uruchomiłeś program**, a nie miejsce, w którym **leży plik ze skryptem**.

To dwa różne miejsca i bardzo łatwo je pomylić:

```text
C:\projekty\kurs\                           <- tu leży skrypt: examples\04_pierwszy.py
├── examples\
│   └── 04_pierwszy.py                      <- zapisuje do "output\dane.xlsx"
└── output\                                 <- tu CHCESZ zapisać

Uruchamiasz z C:\projekty\kurs\     ->  "output\dane.xlsx" = C:\projekty\kurs\output\dane.xlsx  ✔
Uruchamiasz z C:\projekty\kurs\examples\  ->  "output\dane.xlsx" = C:\projekty\kurs\examples\output\dane.xlsx  ✘
Uruchamiasz z harmonogramu zadań (C:\Windows\System32\)  ->  ...  ✘✘✘
```

Ten sam plik, ten sam kod, trzy różne wyniki. Rozwiązanie jest jedno i w kursie stosujemy je zawsze: **ścieżki budujemy jawnie, od położenia pliku ze skryptem.**

### Analogia trzecia: `try/finally` to zamykanie drzwi

Są drzwi, przez które wchodzisz i wychodzisz. Jeśli w środku potkniesz się i przewrócisz — nadal musisz je za sobą zamknąć, żeby nie wywiało Ci dokumentów.

Blok `finally` w Pythonie to właśnie te drzwi. Kod w nim wykonuje się **zawsze**: gdy funkcja zakończy się normalnie, gdy zwróci wartość, gdy rzuci wyjątek, gdy wyjątek przeleci wyżej. Nie ma sytuacji, w której `finally` się „nie uda wykonać", o ile proces w ogóle żyje.

Po co to w openpyxl? Bo skoroszyt może trzymać uchwyt do pliku. W trybie `read_only` trzyma go **na pewno**, aż do wywołania `wb.close()`. Jeśli w połowie czytania wielkiego pliku poleci wyjątek i nie zamkniesz skoroszytu, uchwyt zostaje otwarty. Na Windows oznacza to, że plik jest zablokowany — nie usuniesz go, nie nadpiszesz, a w długo działającym procesie po kilkuset iteracjach skończy Ci się limit otwartych plików.

To nie jest teoria — w historii zmian openpyxl są na to konkretne wpisy: „File handlers not always released in read-only mode" oraz „Workbook files not properly closed on Python ≥ 3.11.8 and Windows" (naprawione w 3.1.3). Biblioteka poprawia się, ale **reguła po Twojej stronie jest jedna**: otworzyłeś skoroszyt — zamknij go, i to w `finally`.

## 2. Teoria

### 2.1. Dwa importy, z których korzysta 90% kursu

```python
from openpyxl import Workbook, load_workbook
```

To są dwie różne operacje i warto je odróżnić już na starcie:

- `Workbook()` — **tworzy** nowy, pusty skoroszyt w pamięci. Nie dotyka żadnego pliku. Nie potrzebuje żadnego pliku.
- `load_workbook(...)` — **wczytuje** istniejący plik do pamięci i zwraca gotowy obiekt `Workbook`.

Warto zapamiętać, że `Workbook` jest **klasą**, a nie funkcją — napisane z dużej litery, bo tworzysz instancję. W starszych tutorialach znajdziesz `load_workbook()` pisane małą literą i to nie jest niekonsekwencja: `load_workbook` to funkcja (procedura wczytująca), a `Workbook` to typ obiektu, który powstaje w wyniku.

Dodatkowo, na potrzeby diagnostyki, warto znać wersję biblioteki:

```python
import openpyxl
print(openpyxl.__version__)      # np. '3.1.5'
```

Zrobisz to w ćwiczeniu 🟢 i będziesz powtarzać w każdym razie, gdy coś „nie działa jak w tutorialu". Wersja to pierwsza informacja, o którą warto zapytać przed szukaniem błędu.

### 2.2. `Workbook()` — dlaczego od razu jest w nim arkusz `Sheet`

Zacznijmy od zaskoczenia, które łapie prawie każdego:

```python
from openpyxl import Workbook

wb = Workbook()
print(wb.sheetnames)      # ['Sheet']
```

Utworzyłeś „pusty" skoroszyt, a w środku już coś jest. Dlaczego?

**Nie jest to uprzejmość biblioteki. To wymóg formatu pliku.**

W module 02 rozbieraliśmy `.xlsx` na części. Jedną z nich był `xl/workbook.xml` — spis treści skoroszytu. Ten plik musi zawierać listę arkuszy. I nie może być na niej **nic**. Skoroszyt bez ani jednego arkusza nie jest poprawnym dokumentem arkusza kalkulacyjnego — Excel nie ma czego wyświetlić, nie ma gdzie postawić kursora, nie ma czego zapisać.

openpyxl rozwiązuje to tak, że **od pierwszej chwili** tworzy model, który jest poprawny: jeden arkusz o nazwie `Sheet`. Dzięki temu możesz natychmiast pisać:

```python
ws = wb.active          # to jest ten domyślny arkusz
ws["A1"] = "cokolwiek"
```

…bez żadnego `create_sheet`. Gdyby domyślnego arkusza nie było, każde rozpoczęcie pracy wymagałoby trzech dodatkowych linijek.

Trzy rzeczy warte zapamiętania z tego miejsca:

**`wb.active` to nie „pierwszy arkusz" — to „arkusz wskazany jako aktywny".** W świeżo utworzonym skoroszycie wskazany jest arkusz o indeksie 0, czyli właśnie `Sheet`. Ale to ustawienie można zmienić i w module 08 zrobimy to świadomie (na przykład po to, żeby użytkownik otwierał raport na zakładce „Podsumowanie", a nie na surowych danych).

**`wb.active` to właściwość, nie metoda.** Piszesz `wb.active`, nie `wb.active()`. Dodanie nawiasów da Ci `TypeError: 'Worksheet' object is not callable`, bo próbujesz wywołać obiekt arkusza jak funkcję.

**Nazwa `Sheet` jest tylko domyślną wartością.** Zmienisz ją jednym przypisaniem:

```python
ws = wb.active
ws.title = "Dane"
print(wb.sheetnames)      # ['Dane']
```

I jeszcze rzecz praktyczna: **nie zostawiaj arkusza o nazwie `Sheet`** w raporcie, który ktoś dostanie. Nazwa jest techniczną pozostałością i wygląda jak niedokończona robota. Nadawaj nazwy od razu.

### 2.3. `wb.save()` — co dokładnie się dzieje i co grozi

W module 03 opisaliśmy cztery fazy życia modelu. `save()` to faza czwarta. Rozłóżmy ją na kroki techniczne, bo to wyjaśnia zachowanie w sytuacjach awaryjnych:

```text
wb.save("output/raport.xlsx")
        │
        ├─ 1. Python otwiera plik w trybie binarnym zapisu
        │     -> plik, jeśli istniał, zostaje OBCIĘTY DO ZERA BAJTÓW
        │
        ├─ 2. openpyxl przechodzi przez CAŁY model i generuje nowe części XML
        │     ([Content_Types].xml, xl/workbook.xml, xl/worksheets/*.xml,
        │      xl/styles.xml, docProps/*.xml i resztę, którą model zna)
        │
        ├─ 3. wygenerowane części są pakowane do archiwum ZIP
        │     i strumieniowane na dysk
        │
        └─ 4. archiwum jest zamykane, uchwyt zwalniany
```

Zwróć uwagę na krok 1. **Plik jest niszczony zanim powstanie nowa treść.** Jeżeli krok 2 albo 3 przerwie się w połowie, na dysku zostanie obcięty albo częściowo zapisany plik. W module 18 nauczymy się temu zapobiegać techniką zapisu atomowego (piszemy do pliku tymczasowego, a dopiero potem podmieniamy nazwę). Na razie zapamiętaj samą zasadę: **`save()` to operacja destrukcyjna, która najpierw usuwa, potem tworzy.**

Trzy praktyczne konsekwencje:

**Konsekwencja 1 — nadpisanie jest ciche.** Jeśli chcesz zapisać wynik pod nową nazwą, po prostu podaj inną nazwę. Nic więcej nie musisz robić. Ale jeśli przez pomyłkę podasz nazwę pliku źródłowego, nie dostaniesz żadnego sygnału ostrzegawczego.

**Konsekwencja 2 — katalog musi istnieć.** Zapis do ścieżki, której katalog nie istnieje, kończy się `FileNotFoundError`:

```python
wb.save("output/2026/październik/raport.xlsx")
# FileNotFoundError: [Errno 2] No such file or directory:
#   'output/2026/październik/raport.xlsx'
```

Zwróć uwagę: **openpyxl nie tworzy katalogów samodzielnie.** Ani `save()`, ani `load_workbook()`. To Twoja odpowiedzialność i zrobimy to jedną linijką w sekcji 2.6.

**Konsekwencja 3 — plik otwarty w Excelu.** Na Windows, jeśli użytkownik ma ten plik otwarty w Excelu, zapis zakończy się:

```text
PermissionError: [Errno 13] Permission denied: 'C:\\...\\raport.xlsx'
```

Excel trzyma plik zablokowany. Na Linuksie i macOS zapis **zwykle się powiedzie** — plik zostanie nadpisany pod spodem, a Excel tego nie zauważy i przy swoim zamknięciu nadpisze Twoją pracę swoją. Ten sam mechanizm utraconej aktualizacji, który omawialiśmy w module 03, tylko widziany z drugiej strony.

Wniosek operacyjny jest jeden i będzie wracać: **do pliku, który może być otwarty w Excelu, nie pisz nigdy.** Pisz do nowego pliku i przekaż użytkownikowi ścieżkę.

#### Rozszerzenie pliku ma znaczenie

Dwie rzeczy, które łatwo pomylić:

**`.xlsx` to zwykły skoroszyt. `.xlsm` to skoroszyt z makrami.** Jeśli wczytasz plik `.xlsm` bez flagi `keep_vba=True`, a potem go zapiszesz, dostaniesz plik `.xlsm` **bez makr**. Excel otworzy go, ale:

- ostrzeże, że „format pliku i rozszerzenie nie są zgodne" albo
- przy otwarciu zgłosi problem z zawartością,
- w każdym razie **Twoje makra zniknęły** i nic Cię o tym nie poinformowało.

Pełne omówienie makr i flagi `keep_vba` znajduje się w module 18. Tutaj zapamiętaj jedno zdanie: **jeśli plik ma makra, wczytuj go z `keep_vba=True` albo wcale go nie zapisuj przez openpyxl.**

**`.xls` to zupełnie inny format.** To stary format binarny (BIFF) sprzed 2007 roku. openpyxl go **nie obsługuje** — ani do czytania, ani do pisania. Próba wczytania kończy się `InvalidFileException`. Próba zapisania modelu pod nazwą `.xls` jest jeszcze gorsza: openpyxl zapisze **format `.xlsx`** pod rozszerzeniem `.xls`, a Excel będzie zdezorientowany z komunikatem o niezgodności formatu i rozszerzenia. Jeśli dostaniesz `.xls` od kogoś — najpierw go przekonwertuj.

```python
from openpyxl.utils.exceptions import InvalidFileException

try:
    wb = load_workbook("stary_plik.xls")
except InvalidFileException as blad:
    print("To nie jest .xlsx:", blad)
    # openpyxl does not support the old .xls file format, please use
    # xlrd to read this file, or convert it to the more recent .xlsx file format
```

### 2.4. Odczyt: `load_workbook()` i cztery sposoby dostania się do arkusza

```python
from openpyxl import load_workbook

wb = load_workbook("output/raport.xlsx")
```

Po tym wywołaniu masz w pamięci pełny model (tryb normalny — przypomnienie z modułu 03) i cztery sposoby na dojście do arkusza:

```python
wb.sheetnames            # ['Dane', 'Podsumowanie']  - lista TEKSTÓW
wb["Dane"]               # obiekt Worksheet, po nazwie
wb[0]                    # obiekt Worksheet, po indeksie od 0
wb.active                # obiekt Worksheet, wskazany jako aktywny
```

Dwie pary łatwo pomylić, więc zróbmy z tego tabelę:

| Wywołanie | Co zwraca | Co z tym robić |
|---|---|---|
| `wb.sheetnames` | `list[str]` — same nazwy | przejrzeć, wyszukać, porównać, wydrukować |
| `wb.worksheets` | `list[Worksheet]` — same obiekty | iterować, gdy potrzebujesz treści |
| `wb["Dane"]` | `Worksheet` | dalej pracować z arkuszem |
| `wb.active` | `Worksheet` | dalej pracować z arkuszem |

**`wb.sheetnames` to zwykła lista Pythona** — możesz na niej robić wszystko, co na liście: `"Dane" in wb.sheetnames`, `wb.sheetnames.index("Dane")`, `sorted(wb.sheetnames)`. I, co ważne, **kolejność tej listy to kolejność zakładek** w Excelu. To najprostszy sposób sprawdzenia, czy raport ma właściwą strukturę.

**Wielkość liter ma znaczenie.** `wb["dane"]` na skoroszycie z arkuszem `"Dane"` rzuci `KeyError`. To najczęstsza literówka w historii automatyzacji Excela:

```python
wb["Dane "]      # KeyError - spacja na końcu nazwy
wb["dane"]       # KeyError - mała litera
```

Komunikat jest dość konkretny i warto go czytać zamiast zgadywać:

```text
KeyError: 'Worksheet dane does not exist.'
```

**Cztery nawyki, które warto wdrożyć od razu:**

1. **Nie zakładaj istnienia arkusza.** Sprawdź: `if "Dane" not in wb.sheetnames: raise ValueError(...)`. Kod, który wybucha z czytelnym komunikatem, jest lepszy niż kod, który wybucha z `KeyError` w środku pętli po 40 sekundach pracy.
2. **Nie odwołuj się do arkuszy po indeksie, jeśli możesz po nazwie.** `wb[0]` przestanie oznaczać to samo, gdy ktoś przestawi zakładki. `wb["Dane"]` nie przestanie.
3. **`wb.active` w pliku wczytanym z dysku to arkusz, który był aktywny w chwili ostatniego zapisu** — niekoniecznie pierwszy. Po modyfikacji skoroszytu możesz go zmienić.
4. **`load_workbook` nie odświeży pliku.** Jeśli ktoś zmienił plik po Twoim wczytaniu, Twój model jest nieaktualny (moduł 03, sekcja „obiekt żywy").

### 2.5. Struktura skoroszytu: tworzenie i usuwanie arkuszy

Operacje na arkuszu mają trzy warianty, które łatwo pomylić, bo robią prawie to samo:

```python
ws = wb.create_sheet(title="Dane")        # tworzy arkusz na KOŃCU
ws = wb.create_sheet(title="Dane", index=0)  # tworzy arkusz na POCZĄTKU
ws = wb.create_sheet()                    # bez nazwy -> automatyczna
```

**`index` to pozycja w kolejności zakładek**, liczona od 0 (bo to lista Pythona — przypomnienie o dwóch systemach indeksowania z modułu 03). `index=0` wstawia arkusz przed wszystkimi innymi. `index=None` (domyślnie) dopisuje go na końcu.

**Bez `title` openpyxl nada nazwę automatyczną.** W świeżo utworzonym skoroszycie, który ma już `Sheet`, kolejne wywołania `wb.create_sheet()` dadzą `Sheet1`, `Sheet2` i tak dalej. To wygodne w kodzie jednorazowym, ale **nie buduj na tym logiki** — nadawaj nazwy jawnie. Nazwa `Sheet3` w raporcie dla zarządu wygląda jak błąd.

**Duplikaty nazw są rozwiązywane automatycznie.** Jeśli spróbujesz utworzyć arkusz o nazwie, która już istnieje, nie dostaniesz wyjątku — openpyxl sam nada nazwę zmodyfikowaną:

```python
wb = Workbook()
wb["Sheet"].title = "Dane"

dubel = wb.create_sheet(title="Dane")
print(dubel.title)       # 'Dane1'  <- NIE 'Dane'
```

To zachowanie (funkcja `avoid_duplicate_name` w kodzie openpyxl) jest wygodne, ale **bardzo łatwo je przeoczyć**. Jeśli Twój kod zakłada, że po `create_sheet(title="Dane")` arkusz nazywa się dokładnie `"Dane"`, a w skoroszycie już był taki arkusz — dostaniesz cichą niespójność. Waliduj to samodzielnie:

```python
def unikalna_nazwa(wb, proponowana: str) -> str:
    """Zwraca proponowaną nazwę, dodając numer, jeśli jest zajęta."""
    if proponowana not in wb.sheetnames:
        return proponowana
    licznik = 1
    while f"{proponowana}_{licznik}" in wb.sheetnames:
        licznik += 1
    return f"{proponowana}_{licznik}"
```

**Usuwanie arkusza** — dwa równoważne sposoby o różnej wygodzie:

```python
del wb["Dane"]           # po nazwie (czytelne)
wb.remove(ws)            # po obiekcie (przydatne, gdy masz zmienną, a nie nazwę)
```

**Nazwy arkuszy podlegają regułom.** Excel narzuca dwa ograniczenia: maksymalnie 31 znaków oraz zakaz używania znaków `: \ / ? * [ ]`. openpyxl waliduje przynajmniej część z nich i zgłasza `ValueError`:

```python
wb.create_sheet(title="Dane/2026")
# ValueError: Invalid character / found in sheet title
```

Uwaga uczciwościowa: **w historii zmian openpyxl pojawiały się zmiany w zakresie walidacji długości nazwy** (był choćby wpis „Allow sheet titles longer than 31 characters"). Nie polegaj więc na tym, że biblioteka zatrzyma każdą złą nazwę — napisz własny walidator. Pełny zestaw reguł i funkcję `bezpieczna_nazwa_arkusza()` zbudujemy w module 08.

**Ostatnia rzecz, o której trzeba pamiętać:** openpyxl **nie zabroni Ci usunąć ostatniego arkusza**. Możesz zrobić `del wb["Sheet"]` i zostanie Ci skoroszyt bez ani jednego arkusza. Taki plik nie jest poprawnym dokumentem — Excel będzie miał z nim problem. Zasada jest prosta i bezwzględna: **zawsze zostaw co najmniej jeden arkusz.**

**Zmiana aktywnego arkusza** — tu jest pułapka warta zapamiętania. Setter `wb.active` przyjmuje **liczbę** (indeks) albo **obiekt arkusza**, ale **nie przyjmuje tekstu**:

```python
wb.active = 1                   # OK - indeks w wb.sheetnames
wb.active = wb["Podsumowanie"]  # OK - obiekt Worksheet (openpyxl 3.1.x)
wb.active = "Podsumowanie"      # TypeError!
# TypeError: Value must be either a worksheet, chartsheet or numerical index
```

Jeśli masz tylko nazwę i chcesz ustawić aktywny arkusz, użyj pośrednictwa:

```python
wb.active = wb.sheetnames.index("Podsumowanie")
```

Dodatkowo: **nie można uaktywnić arkusza, który jest ukryty.** Próba kończy się `ValueError: Only visible sheets can be made active`. To ma sens — użytkownik nie może zobaczyć schowanej zakładki — i wróci przy okazji arkuszy konfiguracyjnych w module 08.

### 2.6. Ścieżki — „półka nr 3, segregator B", nie „szuflada"

Znamy problem z analogii. Teraz rozwiązanie.

**Reguła:** każdy skrypt buduje ścieżkę od **położenia własnego pliku**, nigdy od katalogu roboczego.

```python
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
```

Rozbierzmy to na fragmenty, bo każdy coś robi:

**`Path(__file__)`** — `__file__` to zmienna, którą Python automatycznie wstawia do każdego uruchamianego skryptu. Zawiera ścieżkę do tego pliku, ale **taką, jaką zna interpreter** — może być względna, może zawierać `..`.

**`.resolve()`** — zamienia ścieżkę na **absolutną i kanoniczną**: rozwiązuje `..`, dokleja katalog główny dysku, zamienia separatory na właściwe dla systemu. Po `.resolve()` masz pewność, że operujesz na jednej, konkretnej lokalizacji.

**`.parent`** — katalog nadrzędny. Jeden `.parent` to katalog, w którym leży plik (`examples/`), drugi `.parent` to katalog projektu. Liczba `.parent`-ów zależy od tego, jak głęboko w strukturze leży skrypt — i **musisz ją policzyć świadomie**. Jeśli przeniesiesz skrypt o poziom wyżej, będzie trzeba poprawić.

**`ROOT / "output"`** — operator `/` na `Path` składa ścieżki. Nie `os.path.join`, nie sklejanie tekstu z `+`, nie wpisywanie `\\` ręcznie. `Path` sam dobierze separator odpowiedni dla systemu (na Windows `\`, na Linuksie `/`).

**`OUTPUT.mkdir(parents=True, exist_ok=True)`** — i to jest ta jedna linijka, o której mówiłem:
- `parents=True` — utwórz **wszystkie brakujące katalogi po drodze**, nie tylko ostatni. Bez tego `mkdir` na ścieżce `output/2026/10` wywali się, jeśli nie ma `output/2026`.
- `exist_ok=True` — **nie zgłaszaj błędu, jeśli katalog już istnieje**. Bez tego drugie uruchomienie skryptu kończy się `FileExistsError`.

Bez obu tych flag funkcja jest bezużyteczna w praktyce — skrypt albo działa raz, albo wywala się przy pierwszym uruchomieniu. To najczęściej kopiowany fragment kodu w projektach z openpyxl.

#### Kiedy `__file__` nie istnieje

`__file__` to mechanizm działający dla uruchamianych plików `.py`. **W sesji interaktywnej i w Jupyterze ta zmienna nie istnieje** i odwołanie do niej kończy się `NameError`. Dla notebooków warto mieć wariant awaryjny:

```python
from pathlib import Path


def katalog_projektu() -> Path:
    """Katalog projektu: dla skryptu - od położenia pliku, dla notebooka - od CWD."""
    try:
        return Path(__file__).resolve().parent.parent
    except NameError:
        # __file__ nie istnieje w sesji interaktywnej ani w notebooku
        return Path.cwd()
```

Ten wzorzec warto mieć w każdym projekcie — jest krótki i oszczędza frustracji.

#### Względne vs bezwzględne — kiedy wolno na skróty

Uczciwie: **w kodzie jednorazowym, który uruchamiasz ręcznie z jednego miejsca, ścieżka względna jest dopuszczalna.** Jeśli piszesz skrypt „na dziś", który zapisuje `output/x.xlsx`, a Ty zawsze uruchamiasz go z katalogu projektu — nic złego się nie stanie.

Ale granica jest ostra: **ścieżka względna przestaje być dopuszczalna w momencie, gdy skrypt uruchomi ktokolwiek inny albo cokolwiek innego.** Zmiana na:

- inny komputer,
- harmonogram zadań (Windows), `cron` (Linux), `launchd` (macOS),
- serwer aplikacyjny (Django, FastAPI, Flask),
- CI (GitHub Actions),
- wsad do Dockera,

…natychmiast zamienia ścieżkę względną w błąd. A wtedy debugowanie „coś się zapisało w złym miejscu" trwa dłużej, niż napisanie od razu trzech linijek z `Path`. Dlatego w tym kursie stosujemy wzorzec jawny **zawsze**, od pierwszego modułu — żeby nie trzeba było o tym pamiętać w module 26.

#### Narzędzia pomocnicze dla ścieżek

```python
from pathlib import Path

p = Path("output/04_raport.xlsx")

p.name           # '04_raport.xlsx'          - sama nazwa pliku
p.stem           # '04_raport'               - nazwa bez rozszerzenia
p.suffix         # '.xlsx'                   - rozszerzenie
p.parent         # Path('output')            - katalog
p.exists()       # czy plik/katalog istnieje
p.stat().st_size # rozmiar w bajtach (jeśli istnieje)
p.with_name("04_raport_backup.xlsx")   # ta sama lokalizacja, inna nazwa
p.with_suffix(".xlsm")                 # to samo, inne rozszerzenie
p.expanduser()   # rozwija '~' na katalog domowy
```

**Word o bezpieczeństwie:** nigdy nie buduj ścieżki ze **stringa pochodzącego od użytkownika** przez proste sklejenie. Nazwa pliku `"../../etc/passwd"` albo `"..\\..\\Windows\\System32\\x.dll"` to próba wyjścia poza katalog, do którego miałeś prawo pisać. Bezpieczeństwo plików to temat modułu 21, ale nawyk warto mieć już teraz: **nazwę pliku bierzemy od użytkownika, ale katalog — zawsze swój.**

### 2.7. `try/finally` i pliki tymczasowe

#### Zamykanie skoroszytu

Zaskoczenie: **`Workbook` w openpyxl nie obsługuje protokołu menedżera kontekstu.** To znaczy, że ten kod **nie zadziała**:

```python
with load_workbook("plik.xlsx") as wb:      # ✘
    ...
# TypeError: 'Workbook' object does not support the context manager protocol
```

Warto to sobie zapisać, bo w internecie krąży dużo przykładów z `with load_workbook(...)`. Nie pochodzą one z działającego kodu — forma `with` na obiekcie `Workbook` **nigdy nie była** częścią API openpyxl i nie jest nią w 3.1.x. Jeśli chcesz użyć menedżera kontekstu, użyj go na poziomie *swojej* funkcji:

```python
from contextlib import closing
from openpyxl import load_workbook

with closing(load_workbook("plik.xlsx")) as wb:
    ...
```

`contextlib.closing` to standardowe narzędzie Pythona: „weź dowolny obiekt mający metodę `close()` i zamknij go na wyjściu z bloku, niezależnie od tego, co się stało w środku". Skoro `Workbook` ma `close()`, ten wzorzec działa.

Ale **wzorzec kanoniczny, którego będziemy używać w całym kursie**, to jawne `try/finally`:

```python
from openpyxl import load_workbook

wb = None
try:
    wb = load_workbook("plik.xlsx", read_only=True)
    ws = wb["Dane"]
    for wiersz in ws.iter_rows(values_only=True):
        przetworz(wiersz)
finally:
    if wb is not None:
        wb.close()
```

Dlaczego `try/finally`, a nie `try/except`? Bo `finally` nie ma nic wspólnego z obsługą błędu — **wykona się także wtedy, gdy błędu nie ma**. To gwarancja sprzątania, nie łapanie wyjątków:

| Sytuacja | `except` zadziała? | `finally` zadziała? |
|---|---|---|
| brak błędu | nie | **tak** |
| błąd obsłużony | tak | **tak** |
| błąd nieobsłużony (leci wyżej) | nie | **tak** |
| `return` w środku `try` | nie | **tak** |
| `sys.exit()` w środku `try` | nie | **tak** |

Dlaczego `if wb is not None`? Bo jeśli `load_workbook` sam rzucił wyjątek (`FileNotFoundError`, `InvalidFileException`), zmienna `wb` nadal jest `None` i nie ma czego zamykać.

**Kiedy `close()` naprawdę ma znaczenie?** W trybie normalnym skoroszyt najczęściej nie trzyma otwartego pliku po wczytaniu (moduł 03 — model jest samodzielny, plik zostaje zamknięty). Ale:

- w trybie `read_only` trzyma uchwyt **na pewno** i `close()` jest obowiązkowe,
- przy wielokrotnym otwieraniu plików w pętli (na przykład 1000 raportów) brakujące `close()` doprowadzi do wyczerpania uchwytów,
- na Windows niezamknięty uchwyt potrafi zablokować plik tak, że nie da się go usunąć ani nadpisać,
- wersje openpyxl starsze niż 3.1.3 miały znane przypadki nieuwalniania uchwytów w trybie read-only.

Nawyk: **wywołuj `wb.close()` zawsze, gdy kończysz pracę ze skoroszytem.** Nawet jeśli teoretycznie nie jest potrzebne — jest tanie, a chroni przed klasą błędów, których objawy są bardzo mylące („nie mogę usunąć pliku").

#### Pliki tymczasowe

Czasem potrzebujesz pliku roboczego: na wynik pośredni, na test, na porównanie dwóch wersji. **Nie wynajduj nazw ręcznie** (`tmp1.xlsx`, `test_final_v2.xlsx`) — używaj `tempfile`. Zaleta praktyczna: system sam dba o unikalność, a w większości systemów katalog tymczasowy jest regularnie czyszczony.

```python
import tempfile
from pathlib import Path

from openpyxl import Workbook

with tempfile.TemporaryDirectory() as katalog_tmp:
    plik = Path(katalog_tmp) / "roboczy.xlsx"

    wb = Workbook()
    wb.active["A1"] = "wynik pośredni"
    wb.save(plik)

    print("istnieje w środku bloku:", plik.exists())        # True
    # ... tu robisz, co chcesz: wczytujesz, porównujesz, konwertujesz ...

print("istnieje po wyjściu z bloku:", plik.exists())        # False
```

`TemporaryDirectory` użyty jako menedżer kontekstu **usuwa cały katalog wraz z zawartością** przy wyjściu z bloku `with` — niezależnie od tego, czy blok zakończył się sukcesem, czy wyjątkiem. Idealne do plików roboczych.

**Pułapka na Windows:** `tempfile.NamedTemporaryFile()` domyślnie tworzy plik **otwarty i zablokowany**, więc próba otwarcia go po raz drugi pod tą samą nazwą kończy się `PermissionError`. Jeśli potrzebujesz nazwy pliku tymczasowego, którą podasz dalej (na przykład do openpyxl), użyj jednego z dwóch podejść:

```python
import os
import tempfile
from pathlib import Path

# Wariant A: nazwa + jawne zamknięcie deskryptora (przenośnie)
deskryptor, nazwa = tempfile.mkstemp(suffix=".xlsx")
os.close(deskryptor)                # zamykamy PRZED podaniem nazwy openpyxl
plik_tmp = Path(nazwa)

# Wariant B: katalog tymczasowy + własna nazwa (najprostsze i najbezpieczniejsze)
# with tempfile.TemporaryDirectory() as katalog:
#     plik_tmp = Path(katalog) / "roboczy.xlsx"
```

Oba warianty znajdziesz w praktyce. Wariant B jest czytelniejszy i w kursie używamy jego.

### 2.8. Zapis do pamięci — `BytesIO`

**Plik nie musi być plikiem na dysku.** W Pythonie istnieje pojęcie **obiektu plikopodobnego** (ang. *file-like object*): czegoś, co zachowuje się jak plik — można do tego pisać, z tego czytać — ale niekoniecznie jest czymś na dysku.

**Analogia: rura kontra szuflada.** Zwykły plik to szuflada w biurku — zawartość zostaje, inni mogą do niej zajrzeć, czasem blokujesz ją kluczem. `BytesIO` to rura, do której wrzucasz bajty: one tam są, ale dopóki ich nie wyjmiesz, nie istnieją nigdzie indziej. Gdy rura zniknie, znikają i bajty.

```python
from io import BytesIO

from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws.title = "Dane"
ws["A1"] = "W pamięci, nie na dysku"

# Zapis do "pliku", który istnieje tylko w RAM
bufor = BytesIO()
wb.save(bufor)
dane = bufor.getvalue()

print(type(dane))       # <class 'bytes'>
print(len(dane))        # np. 4838  - tyle bajtów ma cały skoroszyt
print(dane[:2])         # b'PK'    - bo .xlsx to archiwum ZIP (moduł 02)
```

Trzy rzeczy, które warto zauważyć:

**`dane[:2] == b'PK'`** — to sygnatura archiwum ZIP. To ta sama obserwacja, którą zrobiłeś w module 02 na pliku z dysku. Skoroszyt w pamięci i skoroszyt na dysku to **te same bajty**.

**openpyxl nie zamyka Twojego `BytesIO`.** Po `save(bufor)` możesz swobodnie wywołać `bufor.getvalue()`. To jest zachowanie celowe i różni się od sytuacji, gdy `ZipFile` sam otwiera plik na dysku (wtedy go zamknie).

**Jeśli używasz tego samego bufora ponownie** (na przykład do wczytania tego, co właśnie zapisałeś), ustaw pozycję na początek — `bufor.seek(0)`. ZipFile zwykle sobie z tym radzi samodzielnie, ale jawne `seek(0)` jest czytelniejsze i eliminuje klasę problemów:

```python
bufor.seek(0)
wb2 = load_workbook(bufor)
print(wb2["Dane"]["A1"].value)      # 'W pamięci, nie na dysku'
```

…albo prościej i czyściej — stwórz nowy bufor z przechwyconych bajtów:

```python
wb3 = load_workbook(BytesIO(dane))
```

**Po co to wszystko?** Trzy konkretne zastosowania, które wrócą w kursie:

1. **Testy (moduł 20).** Zamiast zapisywać pliki w katalogu tymczasowym, test możesz przeprowadzić w całości w pamięci — szybko, bez sprzątania, bez ryzyka, że równolegle uruchomiony test nadpisze ten sam plik.
2. **Odpowiedź HTTP (moduł 26).** Zamiast zapisywać raport na dysku serwera, wysyłasz bajty bezpośrednio do przeglądarki użytkownika.
3. **Załącznik e-mail.** Biblioteki pocztowe przyjmują bajty; nie musisz niczego dotykać dysku.

Sygnatura jest ta sama w obu kierunkach: `wb.save(...)` i `load_workbook(...)` przyjmują **ścieżkę albo obiekt plikopodobny**. To znaczy, że możesz napisać raz funkcję przyjmującą „cokolwiek, gdzie da się zapisać", i używać jej w skrypcie, w teście i w serwerze:

```python
from typing import Union
from io import BytesIO
from pathlib import Path


def zapisz(wb, cel: Union[str, Path, BytesIO]) -> None:
    wb.save(cel)          # działa dla wszystkich trzech typów
```

### 2.9. Właściwości dokumentu — ślad audytowy

Każdy skoroszyt ma metadane: kto go utworzył, kiedy, jakim programem, jaki ma tytuł. W otwartym pliku zobaczysz je w `Plik → Informacje`. W pakiecie trafiają do `docProps/core.xml` i `docProps/app.xml` — tych samych części, które oglądałeś w module 02.

**Analogia: metryka na odwrocie zdjęcia.** Zdjęcie wygląda tak samo, ale na odwrocie ktoś zapisał: „Warszawa, lipiec 2026, aparat Nikon, fot. A. Kowalski". Dla zwykłego oglądania to nieistotne. Dla archiwum — bezcenne.

```python
import openpyxl
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Dane"
ws["A1"] = "treść raportu"

print("Wartości domyślne:")
for nazwa in ("creator", "title", "subject", "description", "keywords",
              "category", "lastModifiedBy", "created", "modified"):
    print(f"  {nazwa:16} = {getattr(wb.properties, nazwa)!r}")
```

Co zobaczysz? W 3.1.x pole `creator` ma domyślną wartość `'openpyxl'`, a `title`, `subject`, `description` i pozostałe są zwykle puste (`None`). **Sprawdź to u siebie** — dokładne wartości domyślne bywały zmieniane między wydaniami. Ważniejsza od wartości jest sama nauka: **pole istnieje i możesz je ustawić.**

```python
from datetime import datetime

wb.properties.title = "Raport sprzedaży — październik 2026"
wb.properties.creator = f"generator-raportow/1.4.0 (openpyxl {openpyxl.__version__})"
wb.properties.lastModifiedBy = "automat: generator-raportow/1.4.0"
wb.properties.description = (
    "Dokument wygenerowany automatycznie. Zmiany wprowadzone ręcznie "
    "zostaną nadpisane przy następnym uruchomieniu."
)
wb.properties.subject = "Sprzedaż, październik 2026"
wb.properties.keywords = "raport; sprzedaż; automatyzacja"
wb.properties.category = "Raport wewnętrzny"
wb.properties.created = datetime(2026, 10, 7, 9, 30, 0)
```

Trzy uwagi techniczne, które oszczędzą Ci niespodzianek:

**Ustawiaj `created` i `modified` jawnie, jeśli wartość ma znaczenie.** W 3.1.1 naprawiano błąd „DocumentProperties times set by module import only" — w starszej wersji czas ustawiał się raz, przy pierwszym użyciu modułu, więc **wiele skoroszytów w jednym przebiegu miało identyczny znacznik**. Jeśli zobaczysz u siebie taki objaw, wiesz już, od czego zacząć.

**Podawaj daty bez strefy czasowej.** openpyxl nie ma pełnej obsługi stref. Do `properties` wstawiaj datę **naiwną** — `datetime(2026, 10, 7, 9, 30)` zamiast `datetime.now(timezone.utc)`. Temat stref wraca w module 05 przy wartościach w komórkach; tutaj zasada jest prostsza: nie mieszaj stref, jeśli nie musisz.

**Ustawienie `creator` to nie zabezpieczenie, to higiena.** Każdy może te pola zmienić — w Excelu, w openpyxl, w edytorze XML w archiwum. Nie traktuj ich jako dowodu niepodważalnego. Traktuj jako **konwencję operacyjną**, która w firmie bardzo się opłaca.

#### Po co to w firmie: cztery konkretne scenariusze

**1. Ślad audytowy.** Audytor pyta: „kto wygenerował ten plik z danymi klientów i kiedy?". Jeśli raport pochodzi z Twojego generatora, `creator` mówi wprost: `generator-raportow/1.4.0 (openpyxl 3.1.5)`, a `created` podaje godzinę. Bez tego masz plik „od kogoś, kiedyś" — a plik z danymi osobowymi bez proweniencji to problem zgodnościowy.

**2. Wersjonowanie generatora.** Wpisanie wersji swojego skryptu w `creator` pozwala odpowiedzieć na pytanie „które raporty w archiwum powstały przed poprawką błędu w podatku?" — bez zgadywania. To tanie i bardzo praktyczne.

**3. Wykrywanie plików „z excela".** Jeśli proces zakłada, że raporty przychodzą **wyłącznie** z generatora, to plik z `creator = 'Python'`, `'openpyxl'` lub pustym `creator`... właściwie: jeśli zakładasz, że przyszły plik ma być ręcznie przygotowany, a widzisz `creator = 'openpyxl'` — wiesz, że przeszedł przez Twoją automatyzację. I odwrotnie: jeśli cały proces jest automatyczny, a pojawia się plik z `creator = 'Anna Kowalska'` albo z datą modyfikacji w weekend — wiesz, że ktoś edytował go ręcznie. **Proweniencja to najprostszy test „czy plik przyszedł tam, gdzie powinien".**

**4. Komunikacja z użytkownikiem.** `description` z tekstem „Dokument wygenerowany automatycznie. Zmiany ręczne zostaną nadpisane." rozwiąże Ci dziesiątki pytań mailowych typu „zgubiła mi się moja kolumna". Napisanie tego w metadanych kosztuje jedną linijkę kodu.

**Uczciwa uwaga o tym, co się dzieje z cudzymi plikami.** Jeśli wczytasz skoroszyt, który ktoś utworzył w Excelu, i zapiszesz go pod nową nazwą, **właściwości przejdą razem z modelem** — polski `creator` pozostanie nazwiskiem autora oryginału. Twój ślad audytowy będzie wskazywał na kogoś innego. Dlatego w każdym generowanym raporcie **ustawiaj `creator` i `lastModifiedBy` jawnie, samodzielnie**. To trzy linijki i jedna decyzja projektowa.

### 2.10. Jak wygląda pierwszy dobry program

Zbierzmy wszystko w jeden obraz. Dobry skrypt generujący skoroszyt ma zawsze tę samą anatomię:

```text
1. IMPORTY                    - biblioteka standardowa, potem openpyxl
2. STAŁE ŚCIEŻEK              - ROOT / OUTPUT, obliczone od __file__
3. USTAWIENIE ŚRODOWISKA      - OUTPUT.mkdir(parents=True, exist_ok=True)
4. FUNKCJE                    - jedna odpowiedzialność na funkcję
   ├── zbuduj_skoroszyt()     - tworzy model i zwraca wb
   └── zapisz_bezpiecznie()   - dba o katalog, kopię, zapis
5. main()                     - złożenie całości
6. if __name__ == "__main__"  - uruchomienie
```

To jest dokładnie struktura, którą rozwiną moduły 20 (testy), 22–26 (wzorce i architektura). Ale jej zalążek piszesz już teraz — nie dlatego, że tak wypada, ale dlatego, że **kolejność „zbuduj model, potem zapisz" jest jedyną, która pozwala przetestować to, co zbudowałeś**.

## 3. Przykłady krok po kroku

Wszystkie przykłady zapisują do `output/`, zgodnie z konwencją modułu 00. Każdy jest samodzielnym, uruchamialnym skryptem.

### Przykład 1 — pierwszy program w 20 linijkach

```python
"""Najkrótszy pełny cykl: utwórz -> zapisz -> wczytaj -> odczytaj.

Uruchom:  python examples/04_pierwszy.py
"""

from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook

# --- ŚCIEŻKI: budujemy je OD POŁOŻENIA TEGO PLIKU, nie od katalogu roboczego ---
ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)      # 1) katalog musi istnieć 2) brak błędu, gdy istnieje

sciezka = OUTPUT / "04_pierwszy.xlsx"


def zbuduj() -> Workbook:
    """FAZA 1 + 2: model powstaje i jest wypełniany. Dysk pozostaje nietknięty."""
    wb = Workbook()                 # -> skoroszyt z jednym arkuszem 'Sheet'
    ws = wb.active                  # -> ten arkusz (właściwość, nie metoda!)
    ws.title = "Dane"               # <- nie zostawiamy technicznej nazwy 'Sheet'

    ws["A1"] = "Produkt"
    ws["B1"] = "Kwota"
    ws["A2"] = "Widzet"
    ws["B2"] = 1234.5
    ws["A3"] = "Wkręt"
    ws["B3"] = 89.99

    return wb


def zapisz(wb: Workbook, cel: Path) -> None:
    """FAZA 4: przepisanie całego modelu do nowego pakietu na dysku."""
    wb.save(cel)


def odczytaj(cel: Path) -> None:
    """Ponowny odczyt - jedyny sposób sprawdzenia, co FAKTYCZNIE jest w pliku."""
    wb = load_workbook(cel)
    try:
        print(f"  arkusze:        {wb.sheetnames}")
        print(f"  aktywny:        {wb.active.title!r}")
        print(f"  A1 / B1:        {wb['Dane']['A1'].value!r} / {wb['Dane']['B1'].value!r}")
        print(f"  A2 / B2:        {wb['Dane']['A2'].value!r} / {wb['Dane']['B2'].value!r}")
        print(f"  typ wartości B2: {type(wb['Dane']['B2'].value).__name__}")
    finally:
        wb.close()                  # nawyk: zamknij, co otworzyłeś


def main() -> None:
    print(f"openpyxl {openpyxl.__version__}")
    print(f"katalog wyjściowy: {OUTPUT}")
    print()

    wb = zbuduj()
    print("W MODELU (RAM):")
    print(f"  arkusze: {wb.sheetnames}")
    print(f"  B2 = {wb.active['B2'].value}")

    print(f"\nPlik na dysku przed save(): {sciezka.exists()}")

    zapisz(wb, sciezka)

    print(f"Plik na dysku po save():    {sciezka.exists()} "
          f"({sciezka.stat().st_size} B)")
    print()

    print("Z PLIKU (po ponownym wczytaniu):")
    odczytaj(sciezka)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `Workbook()` tworzy graf: jeden `Workbook`, jeden `Worksheet` o nazwie `Dane`, sześć komórek w słowniku `_cells`. Zmiana `ws.title` podmienia nazwę w modelu — i w liście `wb.sheetnames`, bo to ta sama informacja.

**Co trafi do pliku.** Nic, dopóki nie wykona się `save()`. Po `save()` w katalogu `output/` leży `04_pierwszy.xlsx` (kilka tysięcy bajtów) — archiwum ZIP, które otworzysz w Excelu.

**Oczekiwany wynik:**

```text
openpyxl 3.1.5
katalog wyjściowy: /home/ty/kurs-openpyxl/output

W MODELU (RAM):
  arkusze: ['Dane']
  B2 = 1234.5

Plik na dysku przed save(): False
Plik na dysku po save():    True (4789 B)

Z PLIKU (po ponownym wczytaniu):
  arkusze:        ['Dane']
  aktywny:        'Dane'
  A1 / B1:        'Produkt' / 'Kwota'
  A2 / B2:        'Widzet' / 1234.5
  typ wartości B2: float
```

Zwróć uwagę na jedną rzecz, którą zrobiliśmy celowo: **sprawdziliśmy zawartość pliku przez ponowne wczytanie, a nie przez zaufanie modelowi.** Przyzwyczaj się do tego. Model może zawierać coś innego niż plik — na przykład wtedy, gdy zapisu nie było, albo gdy openpyxl czegoś nie umiał odtworzyć. Weryfikacja przez odczyt to najtańszy test, jaki możesz sobie zafundować.

### Przykład 2 — dlaczego ścieżki względne są zdradliwe

```python
"""Demonstracja różnicy między katalogiem roboczym a położeniem skryptu.

Uruchom DWA RAZY, z różnych katalogów:

    cd kurs-openpyxl          && python examples/04_sciezki.py
    cd kurs-openpyxl/examples && python 04_sciezki.py

...i porównaj sekcję ROZWIĄZANIE.
"""

import os
from pathlib import Path


def katalog_projektu() -> Path:
    """Katalog projektu - niezależny od tego, skąd uruchomiono skrypt."""
    try:
        return Path(__file__).resolve().parent.parent
    except NameError:
        # np. w Jupyterze __file__ nie istnieje
        return Path.cwd()


def main() -> None:
    print("=" * 72)
    print("DIAGNOZA")
    print("=" * 72)

    print(f"  katalog roboczy (CWD):     {Path.cwd()}")
    print(f"  zmienna __file__:          {__file__!r}")
    print(f"  __file__ po resolve():     {Path(__file__).resolve()}")
    print(f"  katalog tego skryptu:      {Path(__file__).resolve().parent}")
    print(f"  separator systemowy:       {os.sep!r}")

    print()
    print("=" * 72)
    print("PROBLEM: ŚCIEŻKA WZGLĘDNA")
    print("=" * 72)

    wzgledna = Path("output/04_sciezki.xlsx")
    bezwzgledna_wzglednej = wzgledna.resolve()
    print(f"  Path('output/04_sciezki.xlsx')")
    print(f"    -> zostanie rozwiązane do: {bezwzgledna_wzglednej}")
    print()
    print("  ^ Ta ścieżka ZALEŻY od katalogu roboczego.")
    print("    Uruchom ten skrypt z innego katalogu - zmieni się.")
    print("    W harmonogramie zadań / cronie / Dockerze - prawie zawsze zły.")

    print()
    print("=" * 72)
    print("ROZWIĄZANIE: ŚCIEŻKA JAWNA, OD POŁOŻENIA PLIKU")
    print("=" * 72)

    root = katalog_projektu()
    output = root / "output"
    output.mkdir(parents=True, exist_ok=True)
    cel = output / "04_sciezki.xlsx"

    print(f"  ROOT   = {root}")
    print(f"  OUTPUT = {output}")
    print(f"  cel    = {cel}")
    print()
    print("  ^ Ta ścieżka jest TA SAMA niezależnie od katalogu roboczego.")

    # Mały dowód: zapiszmy coś i sprawdźmy, gdzie naprawdę wylądowało.
    from openpyxl import Workbook

    wb = Workbook()
    wb.active["A1"] = f"zapisane z CWD={Path.cwd()}"
    wb.save(cel)

    print(f"\n  Zapisano. Rozmiar: {cel.stat().st_size} B")
    print(f"  Katalog pliku:     {cel.parent}")

    print()
    print("=" * 72)
    print("OPERACJE NA ŚCIEŻKACH - ŚCIĄGAWKA")
    print("=" * 72)
    p = Path("dane/2026/raport_sprzedazy.xlsx")
    print(f"  p           = {p}")
    print(f"  p.name      = {p.name}")
    print(f"  p.stem      = {p.stem}")
    print(f"  p.suffix    = {p.suffix}")
    print(f"  p.parent    = {p.parent}")
    print(f"  p.parts     = {p.parts}")
    print(f"  p.with_suffix('.xlsm') = {p.with_suffix('.xlsm')}")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `Path` to obiekt, nie tekst. Trzyma listę segmentów (`p.parts`) i informację o systemie. Operacje `/`, `.parent`, `.with_suffix()` zwracają **nowe** obiekty `Path` — nie modyfikują oryginału (tak jak style w module 03: nic nie mutujemy, tworzymy nowe).

**Co trafi do pliku.** Plik `output/04_sciezki.xlsx` — zawsze w tym samym miejscu, niezależnie od tego, skąd uruchomiłeś skrypt. W komórce `A1` **znajdziesz swój katalog roboczy** — to celowy dowód, że działałeś z konkretnego CWD, a mimo to plik trafił tam, gdzie miał trafić.

**Oczekiwany wniosek po dwóch uruchomieniach:** sekcja DIAGNOZA pokaże **różne** `Path.cwd()`, ale sekcja ROZWIĄZANIE pokaże **identyczny** `cel`. To jest cała różnica.

### Przykład 3 — arkusze: tworzenie, kolejność, usuwanie, aktywny

```python
"""Zarządzanie strukturą skoroszytu. Pełne zasady - moduł 08.

Uruchom:  python examples/04_arkusze.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def pokaz(wb, opis: str) -> None:
    aktywny = getattr(wb.active, "title", None)
    print(f"  {opis:<24} {wb.sheetnames}   aktywny={aktywny!r}")


def main() -> None:
    print("=" * 72)
    print("TWORZENIE I KOLEJNOŚĆ")
    print("=" * 72)

    wb = Workbook()
    pokaz(wb, "po Workbook():")                  # ['Sheet']

    # Zmieniamy nazwę domyślnego arkusza - nie chcemy 'Sheet' w raporcie.
    wb.active.title = "Podsumowanie"
    pokaz(wb, "po zmianie nazwy:")

    # Nowy arkusz trafia NA KOŃCU (index=None to domyślne zachowanie).
    dane = wb.create_sheet(title="Dane")
    pokaz(wb, "po create_sheet('Dane'):")

    # index=0 wstawia arkusz NA POCZĄTKU.
    konf = wb.create_sheet(title="Konfiguracja", index=0)
    pokaz(wb, "po create_sheet(index=0):")

    print()
    print("=" * 72)
    print("NAZWY AUTOMATYCZNE I DUPLIKATY")
    print("=" * 72)

    auto1 = wb.create_sheet()
    auto2 = wb.create_sheet()
    print(f"  bez title:            {auto1.title!r}, {auto2.title!r}")
    print("  ^ openpyxl numeruje kolejne arkusze (Sheet, Sheet1, Sheet2, ...)")
    print("    NIE buduj logiki na tych nazwach - nadawaj nazwy jawnie.")

    # Duplikat: openpyxl NIE rzuca wyjątku, tylko modyfikuje nazwę.
    dubel = wb.create_sheet(title="Dane")
    print(f"\n  create_sheet(title='Dane'): {dubel.title!r}")
    print("  ^ openpyxl sam rozwiązał konflikt nazw (funkcja avoid_duplicate_name).")
    print("    Cicha zmiana nazwy bywa zaskakująca - waliduj samodzielnie.")

    # Nasza własna walidacja:
    def unikalna_nazwa(wb, proponowana: str) -> str:
        if proponowana not in wb.sheetnames:
            return proponowana
        licznik = 1
        while f"{proponowana}_{licznik}" in wb.sheetnames:
            licznik += 1
        return f"{proponowana}_{licznik}"

    print(f"  nasza propozycja:              {unikalna_nazwa(wb, 'Dane')!r}")

    print()
    print("=" * 72)
    print("NAZWY NIEDOZWOLONE")
    print("=" * 72)

    for zla in ("Dane/2026", "Raport: Q3"):
        try:
            wb.create_sheet(title=zla)
            print(f"  {zla!r:<16} -> utworzono (!)  sprawdź zachowanie w SWOJEJ wersji")
        except ValueError as blad:
            print(f"  {zla!r:<16} -> ValueError: {blad}")

    print()
    print("  Reguła Excela: nazwa <= 31 znaków, bez znaków  : \\ / ? * [ ]  ")
    print("  openpyxl waliduje część z tego. Zawsze waliduj samodzielnie,")
    print("  bo zachowanie bywało zmieniane między wersjami. Pełny walidator - moduł 08.")

    print()
    print("=" * 72)
    print("USUWANIE ARKUSZY")
    print("=" * 72)
    pokaz(wb, "przed usuwaniem:")

    del wb["Konfiguracja"]        # po nazwie
    pokaz(wb, "del wb['Konfiguracja']:")

    wb.remove(auto1)              # po obiekcie
    pokaz(wb, "wb.remove(auto1):")

    print("\n  ZASADA: nigdy nie usuwaj ostatniego arkusza.")
    print("  openpyxl Ci na to pozwoli, ale skoroszyt bez arkuszy")
    print("  nie jest poprawnym plikiem i Excel będzie miał z nim problem.")

    print()
    print("=" * 72)
    print("AKTYWNY ARKUSZ")
    print("=" * 72)
    pokaz(wb, "stan wyjściowy:")

    # Sposób 1: indeks (liczba)
    wb.active = 1
    pokaz(wb, "wb.active = 1:")

    # Sposób 2: obiekt arkusza (działa w openpyxl 3.1.x)
    wb.active = wb["Podsumowanie"]
    pokaz(wb, "wb.active = wb['Podsumow...']:")

    # Sposób 3: tekst - TAK NIE DZIAŁA
    try:
        wb.active = "Dane"
        print("  wb.active = 'Dane'      zadziałało (niespodzianka w tej wersji)")
    except TypeError as blad:
        print(f"  wb.active = 'Dane'      TypeError: {blad}")

    print("\n  Z nazwy na indeks: wb.active = wb.sheetnames.index('Dane')")

    print()
    print("=" * 72)
    print("ZAPIS I KONTROLA Z DYSKU")
    print("=" * 72)

    cel = OUTPUT / "04_arkusze.xlsx"
    wb.active = wb.sheetnames.index("Podsumowanie")   # raport ma się otwierać tu
    wb.save(cel)

    kontrola = load_workbook(cel)
    try:
        pokaz(kontrola, "z pliku:")
        print(f"\n  który arkusz otworzy się w Excelu? {kontrola.active.title!r}")
        print(f"  rozmiar pliku: {cel.stat().st_size} B")
    finally:
        kontrola.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `wb._sheets` to lista obiektów arkuszy. `create_sheet(index=0)` **wstawia** element na początek listy (przesuwając resztę), a nie dopisuje. `del wb["Konfiguracja"]` szuka arkusza po nazwie i usuwa go z listy — ale **nie czyści odwołań do niego wewnątrz innych części modelu**. Jeśli jakiś arkusz miał nazwę zdefiniowaną albo regułę wskazującą na usunięty arkusz, to odwołanie zostanie martwe. To dokładnie analogia z modułu 03: usunąłeś stronę z indeksem, a indeks nadal na nią wskazuje. Szczegóły w module 08.

**Co trafi do pliku.** `output/04_arkusze.xlsx` z arkuszem `Podsumowanie` ustawionym jako aktywny. Po otwarciu w Excelu kursor będzie na zakładce „Podsumowanie", mimo że nie jest ona pierwsza na liście — i to jest zachowanie, które chcesz, gdy raport ma się zaczynać od wyników, a nie od danych źródłowych.

### Przykład 4 — zapis do pamięci i pliki tymczasowe

```python
"""Skoroszyt, który nigdy nie dotyka dysku - oraz katalog tymczasowy.

Uruchom:  python examples/04_strumien.py
"""

import tempfile
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def zbuduj() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    ws["A1"] = "W pamięci, nie na dysku"
    ws["A2"] = 42
    return wb


def main() -> None:
    print("=" * 72)
    print("ZAPIS DO PAMIĘCI (BytesIO)")
    print("=" * 72)

    wb = zbuduj()

    bufor = BytesIO()          # "rura" na bajty zamiast szuflady na dysku
    wb.save(bufor)             # ta sama metoda, inny cel
    dane = bufor.getvalue()    # wyjmujemy bajty

    print(f"  typ obiektu:      {type(dane).__name__}")
    print(f"  rozmiar:          {len(dane)} B")
    print(f"  pierwsze 2 bajty: {dane[:2]!r}   <- 'PK' = archiwum ZIP (moduł 02)")
    print(f"  bufor zamknięty?  {bufor.closed}   <- openpyxl NIE zamyka Twojego bufora")
    print(f"  plik na dysku?    {list(OUTPUT.glob('04_strumien.xlsx')) or 'nie zapisano'}")

    print()
    print("=" * 72)
    print("ODCZYT Z PAMIĘCI (round-trip)")
    print("=" * 72)

    wb2 = load_workbook(BytesIO(dane))
    try:
        print(f"  arkusze:          {wb2.sheetnames}")
        print(f"  A1:               {wb2['Dane']['A1'].value!r}")
        print(f"  A2:               {wb2['Dane']['A2'].value!r}")
        print("  ^ pełny skoroszyt wrócił z bajtów. Dysk nie był potrzebny.")
    finally:
        wb2.close()

    # Ten sam bufor użyty ponownie - dla porządku przewijamy na początek.
    bufor.seek(0)
    wb3 = load_workbook(bufor)
    try:
        print(f"  z tego samego bufora:  {wb3['Dane']['A1'].value!r}")
    finally:
        wb3.close()

    print()
    print("=" * 72)
    print("PORÓWNANIE: PAMIĘĆ vs DYSK")
    print("=" * 72)

    cel = OUTPUT / "04_strumien.xlsx"
    wb.save(cel)
    print(f"  w pamięci: {len(dane):>7} B")
    print(f"  na dysku:  {cel.stat().st_size:>7} B")
    print("  ^ różnice rzędu kilkudziesięciu bajtów to metadane pakietu")
    print("    (kolejność części, znaczniki czasu w nagłówkach ZIP).")
    print("    Rozmiar bywa INNY przy każdym zapisie - nie porównuj plików")
    print("    bajt w bajt (wróci to w module 20).")

    print()
    print("=" * 72)
    print("KATALOG TYMCZASOWY")
    print("=" * 72)

    with tempfile.TemporaryDirectory() as tmp:
        roboczy = Path(tmp) / "roboczy.xlsx"
        wb.save(roboczy)
        print(f"  katalog tymczasowy:  {tmp}")
        print(f"  plik w środku bloku: {roboczy.exists()}  ({roboczy.stat().st_size} B)")

        # Tu możesz spokojnie wczytywać, porównywać, konwertować...
        kontrola = load_workbook(roboczy)
        try:
            print(f"  odczyt:              {kontrola['Dane']['A1'].value!r}")
        finally:
            kontrola.close()

    # Wychodzimy z bloku 'with' -> cały katalog zostaje usunięty.
    print(f"  po wyjściu z with:   {roboczy.exists()}")
    print("  ^ TemporaryDirectory sprząta samo, także przy wyjątku.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Przy `wb.save(bufor)` openpyxl generuje **te same części XML** co przy zapisie na dysk i pakuje je w archiwum ZIP, ale strumień wyjściowy prowadzi do bufora w RAM. `bufor.getvalue()` zwraca `bytes`, czyli niemutowalną sekwencję bajtów — możesz ją przesłać e-mailem, wysłać jako odpowiedź HTTP, wrzucić do testu.

**Co trafi do pliku.** `output/04_strumien.xlsx` powstaje tylko dlatego, że **dodatkowo** wywołaliśmy `wb.save(cel)` z jawną ścieżką. Gdyby tego zabrakło, katalog `output/` nie zawierałby nic nowego — mimo że wykonaliśmy pełny, poprawny zapis skoroszytu.

### Przykład 5 — właściwości dokumentu i podglądanie `docProps`

```python
"""Metadane skoroszytu: ustawianie, odczyt po wczytaniu, podgląd części pakietu.

Uruchom:  python examples/04_wlasciwosci.py
"""

import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

POLA = ("creator", "title", "subject", "description", "keywords",
        "category", "lastModifiedBy", "created", "modified")


def wypisz(properties, naglowek: str) -> None:
    print(f"  --- {naglowek} ---")
    for nazwa in POLA:
        wartosc = getattr(properties, nazwa, "<brak pola w tej wersji>")
        print(f"    {nazwa:16} = {wartosc!r}")


def lokalna_nazwa(tag: str) -> str:
    """'<{ns}creator>' -> 'creator' (bez przestrzeni nazw XML)."""
    return tag.rsplit("}", 1)[-1]


def podglad_czesci(cel: Path) -> None:
    """Pokazuje, co openpyxl naprawdę zapisał w docProps/."""
    print("  Zawartość części pakietu:")
    with zipfile.ZipFile(cel) as archiwum:
        nazwy = archiwum.namelist()

        for kandydat in ("docProps/core.xml", "docProps/app.xml"):
            if kandydat not in nazwy:
                print(f"    {kandydat:<24} <- NIE MA w tym pliku (sprawdź swoją wersję)")
                continue

            print(f"\n    === {kandydat} ===")
            tresc = archiwum.read(kandydat)
            root = ET.fromstring(tresc)

            for element in root:
                nazwa = lokalna_nazwa(element.tag)
                tekst = (element.text or "").strip()
                if tekst:
                    print(f"      {nazwa:<20} {tekst[:90]}")


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    ws["A1"] = "Treść raportu"

    print("=" * 72)
    print("WARTOŚCI DOMYŚLNE (sprawdź u siebie - mogą się różnić między wersjami)")
    print("=" * 72)
    wypisz(wb.properties, "świeży Workbook()")

    print()
    print("=" * 72)
    print("USTAWIAMY METADANE")
    print("=" * 72)

    wb.properties.title = "Raport sprzedaży — październik 2026"
    wb.properties.creator = (
        f"generator-raportow/1.4.0 (openpyxl {openpyxl.__version__}; Python {__import__('sys').version.split()[0]})"
    )
    wb.properties.lastModifiedBy = "automat: generator-raportow/1.4.0"
    wb.properties.subject = "Sprzedaż — październik 2026"
    wb.properties.description = (
        "Dokument wygenerowany automatycznie. Zmiany ręczne zostaną nadpisane "
        "przy następnym uruchomieniu generatora."
    )
    wb.properties.keywords = "raport; sprzedaz; automatyzacja"
    wb.properties.category = "Raport wewnętrzny"

    # UWAGA: data NAIWNA, bez strefy czasowej. openpyxl nie obsługuje stref
    # w pełni - szczegóły w module 05.
    wb.properties.created = datetime(2026, 10, 7, 9, 30, 0)

    wypisz(wb.properties, "po ustawieniu")

    cel = OUTPUT / "04_wlasciwosci.xlsx"
    wb.save(cel)
    print(f"\n  Zapisano: {cel}  ({cel.stat().st_size} B)")

    print()
    print("=" * 72)
    print("ODCZYT Z PLIKU")
    print("=" * 72)
    kontrola = load_workbook(cel)
    try:
        wypisz(kontrola.properties, "wczytane z pliku")
    finally:
        kontrola.close()

    print()
    print("=" * 72)
    print("ZAGLĄDAMY DO WNĘTRZA PAKIETU (moduł 02 w praktyce)")
    print("=" * 72)
    podglad_czesci(cel)

    print()
    print("=" * 72)
    print("PRZYPADEK: CUDZY PLIK (metadane zostają z oryginału)")
    print("=" * 72)

    obcy = OUTPUT / "04_wlasciwosci_obcy.xlsx"
    wb_obcy = Workbook()
    wb_obcy.active["A1"] = "plik, który przyszedł z zewnątrz"
    wb_obcy.properties.creator = "Anna Kowalska"
    wb_obcy.properties.lastModifiedBy = "Anna Kowalska"
    wb_obcy.properties.created = datetime(2026, 8, 12, 14, 5, 0)
    wb_obcy.save(obcy)

    # Ktoś wczytuje plik, coś dopisuje i zapisuje pod nową nazwą:
    przepuszczony = load_workbook(obcy)
    try:
        przepuszczony.active["B1"] = "dopisane przez automat"
        przepuszczony.save(OUTPUT / "04_wlasciwosci_przepuszczony.xlsx")
        print(f"  creator oryginału:          {przepuszczony.properties.creator!r}")
    finally:
        przepuszczony.close()

    final = load_workbook(OUTPUT / "04_wlasciwosci_przepuszczony.xlsx")
    try:
        print(f"  creator po przepuszczeniu:  {final.properties.creator!r}")
        print("  ^ NIE ZMIENIŁ SIĘ. Metadane podróżują z modelem.")
        print("    Jeśli generujesz raport - ustaw je JAWNIE, żeby")
        print("    ślad audytowy wskazywał na Twój generator, nie na autora oryginału.")
    finally:
        final.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `wb.properties` to obiekt `DocumentProperties` z zestawem atrybutów tekstowych i dwóch dat (`created`, `modified`). Ustawiasz je na poziomie skoroszytu — nie arkusza i nie komórki. Po `load_workbook` te same wartości przyjdą z pliku.

**Co trafi do pliku.** `docProps/core.xml` (metadane podstawowe: autor, tytuł, daty — `dc:creator`, `dcterms:created`) i `docProps/app.xml` (metadane aplikacji: nazwa programu, który zapisał plik). Zobaczysz je w sekcji podglądu. W Excelu znajdziesz je w `Plik → Informacje`.

### Przykład 6 — „zapisz bez pomyłki" (zapowiedź modułu 18)

```python
"""Trzy zabezpieczenia, które ratują Cię przed najczęstszym błędem openpyxl.

To wersja MINIMALNA. Pełny zapis bezpieczny (kopia zapasowa + zapis atomowy
+ dziennik audytowy) zbudujesz w ćwiczeniu 🔴.

Uruchom:  python examples/04_zapis_bez_pomylek.py
"""

import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def zapisz_bez_pomylek(wb, sciezka, *, nadpisz: bool = False) -> Path:
    """Zapisuje skoroszyt po sprawdzeniu trzech rzeczy.

    1. Katalog docelowy istnieje (tworzy go, jeśli trzeba).
    2. Plik docelowy nie istnieje - chyba że wprost pozwolono nadpisać.
    3. Zapis idzie przez plik tymczasowy, więc nieudany zapis NIE niszczy
       pliku docelowego.

    Parametry:
        wb       - obiekt Workbook gotowy do zapisu
        sciezka  - ścieżka docelowa (str lub Path)
        nadpisz  - czy pozwolić na zastąpienie istniejącego pliku

    Zwraca:
        Ścieżkę absolutną zapisanego pliku.
    """
    cel = Path(sciezka).expanduser().resolve()

    # (1) KATALOG - openpyxl sam go nie utworzy.
    cel.parent.mkdir(parents=True, exist_ok=True)

    # (2) NIE NADPISUJEMY PO CICHU. Domyślnie wolimy wybuchnąć niż zniszczyć.
    if cel.exists() and not nadpisz:
        raise FileExistsError(
            f"Plik istnieje: {cel}\n"
            f"  - podaj inną nazwę, albo\n"
            f"  - wywołaj z nadpisz=True, jeśli naprawdę chcesz go zastąpić."
        )

    # (3) ZAPIS ATOMOWY: piszemy obok (w tym samym katalogu - wymóg dla
    #     os.replace), a potem podmieniamy nazwę. Jeśli coś pójdzie nie tak,
    #     plik docelowy pozostaje nietknięty.
    deskryptor, tmp_nazwa = tempfile.mkstemp(
        prefix=f".{cel.stem}_", suffix=".tmp", dir=str(cel.parent)
    )
    import os
    os.close(deskryptor)              # mkstemp daje otwarty deskryptor - zamykamy
    plik_tmp = Path(tmp_nazwa)

    try:
        wb.save(plik_tmp)
        os.replace(plik_tmp, cel)     # atomowa podmiana nazwy
    except BaseException:
        plik_tmp.unlink(missing_ok=True)   # sprzątamy po nieudanym zapisie
        raise

    return cel


def main() -> None:
    print("=" * 72)
    print("1. ZWYKŁY ZAPIS (wszystko w porządku)")
    print("=" * 72)

    def nowy_wb(tekst: str) -> Workbook:
        wb = Workbook()
        wb.active.title = "Dane"
        wb.active["A1"] = tekst
        return wb

    cel = zapisz_bez_pomylek(nowy_wb("pierwsza wersja"), OUTPUT / "04_bezpiecznie.xlsx")
    print(f"  zapisano: {cel}  ({cel.stat().st_size} B)")

    print()
    print("=" * 72)
    print("2. PRÓBA NADPISANIA BEZ ZGODY")
    print("=" * 72)
    try:
        zapisz_bez_pomylek(nowy_wb("druga wersja"), OUTPUT / "04_bezpiecznie.xlsx")
    except FileExistsError as blad:
        print(f"  FileExistsError:")
        for linia in str(blad).splitlines():
            print(f"    {linia}")
    print("  ^ Plik z pierwszą wersją NIE został tknięty. Sprawdźmy:")

    sprawdzenie = load_workbook(OUTPUT / "04_bezpiecznie.xlsx")
    try:
        print(f"    A1 w pliku: {sprawdzenie['Dane']['A1'].value!r}")
    finally:
        sprawdzenie.close()

    print()
    print("=" * 72)
    print("3. ŚWIADOME NADPISANIE")
    print("=" * 72)
    cel = zapisz_bez_pomylek(
        nowy_wb("druga wersja - nadpisana świadomie"),
        OUTPUT / "04_bezpiecznie.xlsx",
        nadpisz=True,
    )
    sprawdzenie = load_workbook(cel)
    try:
        print(f"  A1 w pliku: {sprawdzenie['Dane']['A1'].value!r}")
    finally:
        sprawdzenie.close()

    print()
    print("=" * 72)
    print("4. KATALOG, KTÓREGO NIE BYŁO")
    print("=" * 72)
    gleboko = OUTPUT / "2026" / "10" / "podkatalog" / "raport.xlsx"
    print(f"  katalog istnieje przed zapisem? {gleboko.parent.exists()}")
    cel = zapisz_bez_pomylek(nowy_wb("zapis w nowym katalogu"), gleboko)
    print(f"  katalog istnieje po zapisie?    {gleboko.parent.exists()}")
    print(f"  zapisano: {cel}")
    print("  ^ mkdir(parents=True, exist_ok=True) utworzył całą ścieżkę.")

    print()
    print("=" * 72)
    print("PO CO BYŁ PLIK TYMCZASOWY?")
    print("=" * 72)
    print("""
  Bez pliku tymczasowego:
      wb.save(cel)  ->  Python OBCINA plik docelowy do zera bajtów,
                        dopiero potem wpisuje nową treść.
                        Błąd w połowie = plik docelowy ZNISZCZONY.

  Z plikiem tymczasowym:
      wb.save(tmp)  ->  zapisujemy obok, plik docelowy nadal istnieje
      os.replace()  ->  podmiana nazwy w jednej operacji systemowej

  Efekt: albo mamy STARY plik, albo NOWY plik. Nigdy połowiczny.
  To właśnie zapis atomowy. Pełna procedura - moduł 18.
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Funkcja `zapisz_bez_pomylek` nie tworzy żadnych trwałych struktur — sprawdza warunki, tworzy plik tymczasowy, zapisuje model, podmienia nazwę. `os.replace()` to operacja systemowa: w jednym kroku znika nazwa starego pliku i pojawia się nowa, wskazująca na nowe dane. Warunek techniczny: **plik tymczasowy musi leżeć w tym samym katalogu co cel** (i najlepiej na tym samym wolumenie), bo tylko wtedy podmiana jest naprawdę atomowa. Dlatego `dir=str(cel.parent)`, a nie domyślny katalog systemowy.

**Co trafi do pliku.** Po nieudanej próbie nadpisania w pliku `04_bezpiecznie.xlsx` **nadal jest pierwsza wersja** — i to jest cała wartość tego wzorca. Po świadomym nadpisaniu jest druga wersja. Plików tymczasowych nie zostaje: albo zostają podmienione, albo usunięte w `except`.

**Przykładowy wynik:**

```text
  1. ZWYKŁY ZAPIS  -> zapisano: /home/ty/kurs-openpyxl/output/04_bezpiecznie.xlsx (4752 B)
  2. PRÓBA NADPISANIA -> FileExistsError: Plik istnieje: .../04_bezpiecznie.xlsx
     ^ Plik z pierwszą wersją NIE został tknięty.
       A1 w pliku: 'pierwsza wersja'
  3. ŚWIADOME NADPISANIE -> A1 w pliku: 'druga wersja - nadpisana świadomie'
  4. KATALOG, KTÓREGO NIE BYŁO -> katalog istnieje po zapisie? True
```

> Uwaga o rozszerzeniu pliku tymczasowego: openpyxl nie wybiera formatu na podstawie rozszerzenia — zawartość zawsze jest OOXML. Nazwa `.tmp` jest bezpieczna i pomaga w sortowaniu katalogu. Jeśli w Twojej wersji pojawi się ostrzeżenie, zignoruj je lub użyj przyrostka `.xlsx` w nazwie tymczasowej.

## 4. Anatomia API

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook(write_only=False, iso_dates=False)` | Tworzy nowy, pusty skoroszyt w RAM | `write_only` — tryb prasy drukarskiej; `iso_dates` — ISO 8601 dla dat | Zawsze zawiera jeden arkusz `Sheet`. Nie dotyka dysku |
| `load_workbook(filename, ...)` | Wczytuje plik do modelu | `read_only=False`, `keep_vba=False`, `data_only=False`, `keep_links=True`, `rich_text=False` | `filename` może być ścieżką **lub** obiektem plikopodobnym (np. `BytesIO`) |
| `wb.save(filename)` | Serializuje cały model → nowy pakiet | ścieżka **lub** obiekt plikopodobny | Nadpisuje bez ostrzeżenia. Wymaga istniejącego katalogu. Otwiera plik w trybie zapisu (obcina!) |
| `wb.close()` | Zwalnia uchwyt do archiwum | — | Obowiązkowe w `read_only`. **`Workbook` NIE jest menedżerem kontekstu** — `with` nie zadziała |
| `wb.sheetnames` | Lista nazw arkuszy | — | Kolejność = kolejność zakładek. Zwykła lista Pythona |
| `wb.worksheets` | Lista obiektów arkuszy | — | Indeksowanie od 0 |
| `wb.active` | Aktywny arkusz | getter → `Worksheet` lub `None`; setter przyjmuje `Worksheet` **albo** `int` | **Nie przyjmuje tekstu.** Ukrytego arkusza nie da się uaktywnić (`ValueError`) |
| `wb["Nazwa"]` | Arkusz po nazwie | `str` lub `int` (indeks) | `KeyError`, jeśli nie ma; wielkość liter ma znaczenie |
| `wb.create_sheet(title=None, index=None)` | Tworzy arkusz | `index=None` → na końcu; `0` → na początku | Bez `title` nadaje nazwę automatycznie. Duplikaty rozwiązuje samodzielnie |
| `wb.remove(worksheet)` | Usuwa arkusz po obiekcie | obiekt `Worksheet` | Odpowiednik `del wb[nazwa]` |
| `del wb["Nazwa"]` | Usuwa arkusz | `str` lub `int` | Pozwala usunąć także ostatni arkusz — nie rób tego |
| `wb.copy_worksheet(ws)` | Kopiuje arkusz w obrębie skoroszytu | obiekt `Worksheet` | Obrazy i wykresy **nie** są kopiowane. Moduł 08 |
| `wb.move_sheet(ws, offset=0)` | Przesuwa zakładkę | `offset` (int) | Bezpieczniejszy niż manipulowanie listą prywatną. Moduł 08 |
| `wb.index(ws)` | Indeks arkusza na liście | obiekt `Worksheet` | Wygodne przy ustawianiu `wb.active` |
| `wb.properties` | Metadane dokumentu | `title`, `creator`, `subject`, `description`, `keywords`, `category`, `lastModifiedBy`, `created`, `modified` | Trafiają do `docProps/core.xml`. Daty podawaj **bez** strefy |
| `wb.calculation` | Ustawienia przeliczania | `fullCalcOnLoad` | Domyślnie Excel przelicza formuły po otwarciu |
| `wb.security` | Ochrona struktury skoroszytu | `lockStructure`, `workbookPassword` | Moduł 17 |
| `wb.epoch` | Epoka dat skoroszytu | `CALENDAR_WINDOWS_1900` / `CALENDAR_MAC_1904` | Moduł 05 |
| `wb.defined_names` | Nazwy zdefiniowane | — | API przebudowane w 3.1. Moduł 06 |
| `ws.title` | Nazwa arkusza | setter przyjmuje tekst | Niedozwolone znaki → `ValueError`. Max 31 znaków (waliduj sam) |
| `ws.parent` | Skoroszyt arkusza | — | Ten sam obiekt, którym go utworzyłeś |
| `ws["A1"]`, `ws.cell(row, column)` | Komórka | — | Moduł 05 |
| `Path(__file__).resolve().parent` | Katalog skryptu | — | Fundament jawnych ścieżek |
| `Path.mkdir(parents=True, exist_ok=True)` | Tworzy katalog i całą ścieżkę | `parents`, `exist_ok`, `mode` | Bez obu flag funkcja jest praktycznie bezużyteczna w skryptach |
| `BytesIO()` | Bufor bajtów w RAM | `initial_bytes` | `getvalue()` zwraca `bytes`. openpyxl go nie zamyka |
| `tempfile.TemporaryDirectory()` | Katalog tymczasowy usuwany przy wyjściu z `with` | `suffix`, `prefix`, `dir` | Najbezpieczniejszy sposób pracy z plikami roboczymi |
| `tempfile.mkstemp(suffix, prefix, dir)` | Tworzy plik tymczasowy i zwraca `(fd, ścieżka)` | — | Trzeba samemu wywołać `os.close(fd)` |
| `os.replace(src, dst)` | Atomowa podmiana nazwy | — | Działa atomowo na tym samym wolumenie. Podstawa zapisu bezpiecznego (moduł 18) |
| `openpyxl.utils.exceptions.InvalidFileException` | Wyjątek dla nieobsługiwanego formatu | — | Rzucany m.in. dla `.xls` |
| `contextlib.closing(...)` | Menedżer kontekstu dla obiektu z `close()` | — | Wariant, gdy chcesz `with` przy skoroszycie |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — Pierwszy skoroszyt z datą.** Napisz skrypt `examples/04_cw1.py`, który:

1. Wypisuje wersję openpyxl (`openpyxl.__version__`).
2. Tworzy skoroszyt, zmienia nazwę domyślnego arkusza na `"Raport"`.
3. Wprowadza do `A1` tytuł raportu (tekst), do `A2` datę wykonania jako **obiekt `datetime`** (nie tekst!), a do `A3` liczbę `3.14159`.
4. Ustawia `wb.properties.title` i `wb.properties.creator`.
5. Tworzy katalog `output/` **w sposób jawny** (`parent=True, exist_ok=True`).
6. Zapisuje plik jako `output/04_cw1.xlsx` i wypisuje jego rozmiar w bajtach.
7. **Wczytuje plik ponownie** i wypisuje dla każdej z trzech komórek: wartość, typ Pythona (`type(...).__name__`) oraz współrzędną (`cell.coordinate`).
8. Zamyka oba skoroszyty.

Na koniec odpowiedz w pliku `output/04_cw1_wnioski.md`:

- Czy typ wartości odczytanej z `A2` jest taki sam jak typ wartości, którą wpisałeś? Dlaczego?
- Co się stanie, jeśli usuniesz `mkdir`? Sprawdź eksperymentalnie (usuń katalog `output` na chwilę).
- Co się stanie, jeśli nie wywołasz `save()`? Sprawdź.

**Zadanie 2 — Katalog roboczy.** Uruchom **ten sam** skrypt (np. `examples/04_pierwszy.py`) dwa razy: raz z katalogu głównego projektu (`python examples/04_pierwszy.py`), a raz z katalogu `examples` (`cd examples && python 04_pierwszy.py`). W pliku `output/04_cw2_wnioski.md` odpowiedz:

- Czy plik wynikowy powstał w obu przypadkach w tym samym miejscu? Dlaczego?
- Co by się stało, gdyby skrypt zapisywał do `"output/plik.xlsx"` (ścieżka względna)? Sprawdź to eksperymentalnie — zmień ścieżkę w kopii skryptu i uruchom z obu katalogów.
- Wypisz, co zwraca `Path.cwd()` w obu przypadkach.

### 🟡 Warsztat

**Zadanie 3 — Trzy arkusze z konwencją nazw.** Napisz skrypt `examples/04_cw3.py`, który tworzy skoroszyt o strukturze:

| Nazwa arkusza | Zawartość | Uwagi |
|---|---|---|
| `Konfiguracja` | `A1="wersja"`, `B1="1.0"`, `A2="autor"`, `B2=twój nick` | **pierwsza zakładka** (index=0) |
| `Dane` | nagłówki w wierszu 1: `Data`, `Produkt`, `Kwota` + 5 wierszy danych | — |
| `Podsumowanie` | `A1="Raport testowy"`, `A3="Liczba pozycji"`, `B3=5` | **aktywna** przy otwarciu |

Dodatkowo:

1. Usuń domyślny arkusz `Sheet` (nie zmieniaj go — **usuń**), ale zrób to dopiero po utworzeniu arkusza `Konfiguracja`, żeby nie zostać bez arkuszy.
2. Ustaw `Konfiguracja` jako pierwszą zakładkę, a `Podsumowanie` jako aktywną.
3. Zapisz jako `output/04_cw3.xlsx`.
4. **Wczytaj plik i zweryfikuj programowo**: czy lista `sheetnames` zgadza się z oczekiwaną? Czy `wb.active.title == "Podsumowanie"`? Czy `Konfiguracja` jest pierwsza?
5. Wypisz ostrzeżenie (`print`), jeśli weryfikacja się nie powiedzie.

W pliku `output/04_cw3_wnioski.md` odpowiedz:

- Dlaczego usunięcie arkusza `Sheet` **po** utworzeniu nowego jest bezpieczniejsze niż **przed**?
- Co by się stało, gdybyś ustawił `wb.active = "Podsumowanie"` (tekst)? Sprawdź i wklej komunikat błędu.
- Jak zmieniłbyś kod, gdyby lista arkuszy była nieznana z góry (np. przychodzi z konfiguracji)?

**Zadanie 4 — Skoroszyt w pamięci i w pliku.** Napisz skrypt `examples/04_cw4.py`, który:

1. Buduje skoroszyt **w funkcji** `zbuduj(zawartosc: str) -> Workbook` (funkcja nie zapisuje niczego na dysk).
2. Zapisuje go do `BytesIO` i wypisuje `len(danych)`.
3. Wypisuje pierwsze dwa bajty i wyjaśnia w komentarzu, dlaczego są właśnie takie.
4. Wczytuje skoroszyt z powrotem z `BytesIO` i sprawdza, czy zawartość się zgadza.
5. Zapisuje ten sam skoroszyt do katalogu tymczasowego (`TemporaryDirectory`) i porównuje rozmiar z rozmiarem w pamięci.
6. Zapisuje ten sam skoroszyt do `output/04_cw4.xlsx`.
7. Wychodzi z bloku `with` i **sprawdza**, czy plik tymczasowy zniknął.

W pliku `output/04_cw4_wnioski.md` odpowiedz:

- Czy rozmiar `BytesIO` i pliku na dysku jest identyczny? Jeśli nie — podaj różnicę i wyjaśnij, czym może być spowodowana.
- Dlaczego `zbuduj()` nie powinna zawierać `wb.save()`? Podaj dwa powody — jeden związany z testami, jeden z projektem (zapowiedź modułów 20 i 22).
- Zapisz ten sam skoroszyt dwa razy do dwóch buforów i porównaj długości. Czy są równe? Uruchom to trzy razy — czy wynik jest zawsze ten sam?

### 🔴 Wyzwanie

**Zadanie 5 — `zapisz_bezpiecznie()` w pełnej wersji.** To zadanie łączy wszystko, co poznałeś, i jest zalążkiem procedury z modułu 18. Napisz funkcję `zapisz_bezpiecznie(wb, sciezka, **opcje)`, która spełnia **wszystkie** poniższe wymagania:

**Wymagania funkcjonalne:**

1. **Tworzy katalog** docelowy, jeśli nie istnieje (`parents=True, exist_ok=True`) i **raportuje w wyniku**, czy katalog powstał.
2. **Wykonuje kopię zapasową** istniejącego pliku docelowego, jeśli ten istnieje. Nazwa kopii zawiera znacznik czasu w formacie `RRRRMMDD_GGMMSS`, plik trafia do katalogu `_backup/` obok celu (albo do katalogu wskazanego argumentem). Kopię zapisuj przez `shutil.copy2()`, żeby zachować metadane.
3. **Zapisuje atomowo**: do pliku tymczasowego utworzonego **w tym samym katalogu** (`tempfile.mkstemp` z `dir=`), przez `wb.save(tmp)` i `os.replace(tmp, cel)`. W razie dowolnego wyjątku plik tymczasowy jest usuwany, a wyjątek leci dalej.
4. **Odmawia nadpisania**, jeśli plik istnieje, a nie przekazano `nadpisz=True` — z komunikatem podającym ścieżkę i sugestię (podaj inną nazwę / użyj `nadpisz=True`).
5. **Sprawdza, czy plik da się odczytać po zapisie** — wykonuje `load_workbook(cel, read_only=True)` i natychmiast `close()`. Jeśli weryfikacja się nie powiedzie, zgłasza wyjątek **i przywraca plik z kopii zapasowej** (jeśli kopia istnieje).
6. **Zwraca obiekt wyniku** (`dataclass`), zawierający co najmniej: ścieżkę docelową, rozmiar w bajtach, ścieżkę kopii zapasowej (albo `None`), informację czy utworzono katalog, czas trwania zapisu w sekundach.
7. **Mierzy czas** przez `time.perf_counter()`.
8. **Nic nie wypisuje na ekran** — funkcja biblioteczna nie drukuje. Cały raport wraca w obiekcie wyniku, a decyzję o wypisaniu podejmuje wywołujący.

**Wymagania dodatkowe (na ocenę bardzo dobrą):**

9. Funkcja przyjmuje też obiekt plikopodobny (`BytesIO`) — w tym przypadku krok 1, 2, 4 i 5 są pomijane, a wynik ma `sciezka=None`.
10. Ogranicz liczbę kopii zapasowych: argument `max_kopii` (domyślnie 5) — po zapisie usuń najstarsze kopie, zostawiając `max_kopii` najnowszych.

**Test, który musi przejść:** uruchom funkcję cztery razy pod rząd na tej samej ścieżce, z `nadpisz=True` trzeci i czwarty raz. Sprawdź, czy: (a) przy drugim wywołaniu dostajesz `FileExistsError` i plik pozostaje niezmieniony, (b) po czwartym wywołaniu istnieje dokładnie `1 + min(ile, max_kopii)` plików (cel + kopie), (c) żaden plik `.tmp` nie został w katalogu.

<details>
<summary><strong>Szkic rozwiązania zadania 5 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""Pełny zapis bezpieczny - rozwiązanie zadania 5."""

from __future__ import annotations

import os
import shutil
import tempfile
import time
from dataclasses import dataclass, field
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Union

from openpyxl import Workbook, load_workbook

# Typ przyjmowany jako cel zapisu: ścieżka albo obiekt plikopodobny.
Cel = Union[str, os.PathLike, BytesIO]


@dataclass
class WynikZapisu:
    """Raport z operacji zapisu. Funkcja nic nie drukuje - wszystko tutaj."""

    sciezka: Path | None
    rozmiar_bajtow: int
    kopia_zapasowa: Path | None = None
    utworzono_katalog: bool = False
    zweryfikowano: bool = False
    usuniete_kopie: list[Path] = field(default_factory=list)
    czas_s: float = 0.0

    def __str__(self) -> str:
        cel = self.sciezka if self.sciezka else "<bufor pamięci>"
        czesci = [f"zapisano: {cel}", f"{self.rozmiar_bajtow} B", f"{self.czas_s:.3f} s"]
        if self.utworzono_katalog:
            czesci.append("utworzono katalog")
        if self.kopia_zapasowa:
            czesci.append(f"kopia: {self.kopia_zapasowa.name}")
        if self.usuniete_kopie:
            czesci.append(f"usunięto starych kopii: {len(self.usuniete_kopie)}")
        if self.sciezka is not None:
            czesci.append("zweryfikowano" if self.zweryfikowano else "BEZ WERYFIKACJI")
        return " | ".join(czesci)


def _jest_buforem(cel: object) -> bool:
    """Czy cel to obiekt plikopodobny (ma metodę write), a nie ścieżka?"""
    return hasattr(cel, "write")


def _posprzataj_stare_kopie(katalog: Path, wzorzec: str, zostaw: int) -> list[Path]:
    """Zostawia `zostaw` najnowszych kopii pasujących do wzorca, resztę usuwa."""
    kandydaci = sorted(
        katalog.glob(wzorzec),
        key=lambda p: p.stat().st_mtime,
        reverse=True,                    # najnowsze na początku
    )
    usuniete = []
    for plik in kandydaci[zostaw:]:
        plik.unlink(missing_ok=True)
        usuniete.append(plik)
    return usuniete


def zapisz_bezpiecznie(
    wb: Workbook,
    cel: Cel,
    *,
    nadpisz: bool = False,
    kopia_zapasowa: bool = True,
    katalog_backupow: str | os.PathLike | None = None,
    max_kopii: int = 5,
    weryfikuj: bool = True,
) -> WynikZapisu:
    """Zapisuje skoroszyt w sposób odporny na awarie i pomyłki.

    Gwarancje:
      * katalog docelowy istnieje po powrocie funkcji,
      * przy `nadpisz=False` istniejący plik NIE zostanie tknięty,
      * nieudany zapis nie zostawia uszkodzonego pliku docelowego,
      * opcjonalnie: plik da się wczytać po zapisie (weryfikacja),
      * przy nieudanej weryfikacji plik wraca z kopii zapasowej.
    """
    start = time.perf_counter()

    # --- ŚCIEŻKA W PAMIĘCI: kroki plikowe pomijamy ----------------------
    if _jest_buforem(cel):
        wb.save(cel)
        rozmiar = len(cel.getvalue())          # type: ignore[union-attr]
        return WynikZapisu(
            sciezka=None,
            rozmiar_bajtow=rozmiar,
            czas_s=time.perf_counter() - start,
        )

    cel = Path(cel).expanduser().resolve()

    # --- (1) KATALOG ----------------------------------------------------
    utworzono_katalog = False
    if not cel.parent.exists():
        cel.parent.mkdir(parents=True, exist_ok=True)
        utworzono_katalog = True

    # --- (4) OCHRONA PRZED CICHYM NADPISANIEM ---------------------------
    if cel.exists() and not nadpisz:
        raise FileExistsError(
            f"Plik istnieje: {cel}\n"
            f"  - podaj inną nazwę,\n"
            f"  - albo wywołaj funkcję z nadpisz=True."
        )

    # --- (2) KOPIA ZAPASOWA ---------------------------------------------
    kopia: Path | None = None
    usuniete: list[Path] = []

    if cel.exists() and kopia_zapasowa:
        gdzie = Path(katalog_backupow) if katalog_backupow else cel.parent / "_backup"
        gdzie.mkdir(parents=True, exist_ok=True)

        znacznik = datetime.now().strftime("%Y%m%d_%H%M%S")
        kopia = gdzie / f"{cel.stem}_{znacznik}{cel.suffix}"

        # Przy zapisach w tej samej sekundzie nazwa może się powtórzyć.
        licznik = 1
        while kopia.exists():
            kopia = gdzie / f"{cel.stem}_{znacznik}_{licznik}{cel.suffix}"
            licznik += 1

        shutil.copy2(cel, kopia)           # copy2 zachowuje metadane pliku

        if max_kopii > 0:
            usuniete = _posprzataj_stare_kopie(
                gdzie, f"{cel.stem}_*{cel.suffix}", max_kopii
            )

    # --- (3) ZAPIS ATOMOWY ----------------------------------------------
    deskryptor, tmp_nazwa = tempfile.mkstemp(
        prefix=f".{cel.stem}_", suffix=".tmp", dir=str(cel.parent)
    )
    os.close(deskryptor)                   # mkstemp zostawia OTWARTY deskryptor
    plik_tmp = Path(tmp_nazwa)

    try:
        wb.save(plik_tmp)
        os.replace(plik_tmp, cel)          # atomowo: stary albo nowy, nigdy wpół
    except BaseException:
        plik_tmp.unlink(missing_ok=True)   # nie zostawiamy śmieci
        raise

    # --- (5) WERYFIKACJA -------------------------------------------------
    zweryfikowano = False
    if weryfikuj:
        try:
            kontrola = load_workbook(cel, read_only=True)
            try:
                _ = kontrola.sheetnames      # wymusza odczyt nagłówka pakietu
            finally:
                kontrola.close()
            zweryfikowano = True
        except Exception as blad_weryfikacji:
            # Ratujemy poprzednią wersję, jeśli ją mamy.
            if kopia is not None and kopia.exists():
                shutil.copy2(kopia, cel)
                raise RuntimeError(
                    f"Weryfikacja nie powiodła się ({blad_weryfikacji!r}). "
                    f"Przywrócono poprzednią wersję z {kopia.name}."
                ) from blad_weryfikacji
            raise RuntimeError(
                f"Weryfikacja nie powiodła się ({blad_weryfikacji!r}), "
                f"a kopii zapasowej nie było."
            ) from blad_weryfikacji

    # --- (6) WYNIK -------------------------------------------------------
    return WynikZapisu(
        sciezka=cel,
        rozmiar_bajtow=cel.stat().st_size,
        kopia_zapasowa=kopia,
        utworzono_katalog=utworzono_katalog,
        zweryfikowano=zweryfikowano,
        usuniete_kopie=usuniete,
        czas_s=time.perf_counter() - start,
    )
```

**Test czterech wywołań:**

```python
if __name__ == "__main__":
    import tempfile
    from io import BytesIO

    cel = Path("output/04_zapis_bezpieczny.xlsx")

    def wb_z(tekst: str) -> Workbook:
        wb = Workbook()
        wb.active.title = "Dane"
        wb.active["A1"] = tekst
        return wb

    print(zapisz_bezpiecznie(wb_z("wer. 1"), cel))
    # -> zapisano: .../04_zapis_bezpieczny.xlsx | 4760 B | ... | zweryfikowano

    try:
        zapisz_bezpiecznie(wb_z("wer. 2"), cel)
    except FileExistsError as blad:
        print("Odmowa nadpisania OK:", str(blad).splitlines()[0])

    print(zapisz_bezpiecznie(wb_z("wer. 3"), cel, nadpisz=True))
    print(zapisz_bezpiecznie(wb_z("wer. 4"), cel, nadpisz=True))

    # Kontrole:
    kontrola = load_workbook(cel)
    try:
        print("W pliku:", kontrola["Dane"]["A1"].value)      # 'wer. 4'
    finally:
        kontrola.close()

    kopie = list((cel.parent / "_backup").glob(f"{cel.stem}_*{cel.suffix}"))
    print("Liczba kopii:", len(kopie))                        # <= max_kopii
    smieci = list(cel.parent.glob(f".{cel.stem}_*.tmp"))
    print("Plików tymczasowych:", len(smieci))                # 0

    # Bufor pamięci - te same gwarancje poza plikiem
    bufor = BytesIO()
    wynik = zapisz_bezpiecznie(wb_z("w pamięci"), bufor)
    print(wynik)
    print("Bajtów w buforze:", len(bufor.getvalue()))
```

**Kluczowe decyzje projektowe i dlaczego właśnie tak:**

- **Plik tymczasowy w katalogu docelowym, nie w katalogu systemowym.** `os.replace()` jest atomowe tylko w obrębie jednego systemu plików. Gdyby plik tymczasowy leżał na innej partycji, `replace` musiałby kopiować zawartość — i przestałoby być atomowe. Dlatego `dir=str(cel.parent)`.
- **`mkstemp` + `os.close()`.** `mkstemp` tworzy plik i zwraca otwarty deskryptor. Gdybyśmy go nie zamknęli, mielibyśmy plik zablokowany na Windows — dokładnie ta sama pułapka, co przy `NamedTemporaryFile`.
- **`except BaseException`, a nie `except Exception`.** `BaseException` obejmuje `KeyboardInterrupt` (przerwanie przez `Ctrl+C`). Po przerwaniu plik tymczasowy też ma zostać posprzątany, a wyjątek — polecieć dalej. Używanie `BaseException` w `except` jest w ogólności ryzykowne, ale tutaj robimy w tym bloku wyłącznie sprzątanie i natychmiast podnosimy wyjątek dalej — to bezpieczny, świadomy wyjątek od reguły.
- **Kopia przed zapisem, nie po.** Kopia ma sens tylko wtedy, gdy chroni **stan poprzedni**. Kopia po zapisie chroniłaby już nową wersję, czyli była bezwartościowa.
- **Weryfikacja przez `read_only=True`.** Sprawdzamy, czy pakiet da się otworzyć — to tani test integralności archiwum i podstawowej struktury. Wczytanie w trybie `read_only` nie buduje całego modelu, więc nie płacimy za to pełnym czasem wczytywania.
- **Funkcja nic nie wypisuje.** Biblioteka nie wie, czy wynik ma trafić do konsoli, do logu, czy do interfejsu graficznego. Zwracanie obiektu wyniku pozwala wywołującemu zdecydować — a testom pozwala sprawdzać dane, a nie przechwytywać `stdout`.
- **Deduplikacja nazw kopii.** Bez pętli `while kopia.exists()` dwa zapisy w tej samej sekundzie nadpisałyby sobie nawzajem kopie. To bardzo realny scenariusz przy szybkich testach.
- **`dataclass` z `field(default_factory=list)`.** Mutowalna wartość domyślna (`= []`) to klasyczna pułapka Pythona: zostałaby współdzielona między **wszystkie** instancje. `default_factory` tworzy nową listę przy każdym wywołaniu.

**Co jeszcze warto by dodać w wersji produkcyjnej** (moduł 18): dziennik audytowy do pliku, sumę kontrolną (hash) wejścia i wyjścia, obsługę wyjątku `PermissionError` z komunikatem „plik otwarty w Excelu", blokadę na czas zapisu (plik `.lock`), oraz – przy plikach `.xlsm` – obowiązkowe `keep_vba=True` przy wczytaniu i sprawdzenie, czy `wb.vba_archive` nie jest `None`.

</details>

## 6. Typowe błędy i pułapki

**1. „Program się wykonał bez błędu, a w katalogu nie ma pliku" (objaw) → brak wywołania `wb.save()`; model żył tylko w RAM i zniknął razem z procesem (przyczyna) → dodaj `wb.save(sciezka)` przed końcem programu; w skryptach produkcyjnych umieść zapis tak, żeby wykonał się także przy wyjątku (na przykład w `finally` albo wołając funkcję zapisującą z bloku `try/except`) (naprawa).**

Szczególny wariant tego błędu: `save()` wywołany **we wnętrzu pętli**, po którym następują dalsze modyfikacje — zapisujesz stan pośredni, a końcowy nigdy nie trafia na dysk.

**2. „Nadpisałem sobie plik źródłowy i nie mam kopii" (objaw) → `wb.save()` zachowuje się jak „Zapisz", nie „Zapisz jako": istniejący plik jest obcinany do zera bajtów i zastępowany nową treścią, bez ostrzeżenia i bez kopii zapasowej (przyczyna) → zapisuj **pod inną nazwą** niż plik wejściowy; jeśli chcesz pisać pod tę samą, użyj wzorca `zapisz_bez_pomylek` z przykładu 6 albo pełnego `zapisz_bezpiecznie` z ćwiczenia 🔴 (naprawa).**

To najdroższy błąd w całym kursie — kosztuje ludzi dane, a nie czas. Reguła z modułu 00 („nigdy nie nadpisujemy pliku źródłowego") istnieje dokładnie z tego powodu.

**3. „`FileNotFoundError: [Errno 2] No such file or directory`" (objaw) → katalog docelowy nie istnieje, a openpyxl go **nie tworzy**; zapis do zagnieżdżonej ścieżki (`output/2026/10/raport.xlsx`) nie utworzy brakujących poziomów (przyczyna) → wywołaj `Path(parent).mkdir(parents=True, exist_ok=True)` **przed** `wb.save()`; obie flagi są potrzebne: `parents` tworzy całą ścieżkę, `exist_ok` zapobiega `FileExistsError` przy drugim uruchomieniu (naprawa).**

Uwaga: ten sam błąd zdarza się, gdy podasz ścieżkę do pliku zamiast do katalogu jako `dir=` w funkcjach `tempfile` — warto czytać komunikat, bo podaje dokładnie tę ścieżkę, która nie istnieje.

**4. „`KeyError: 'Worksheet dane does not exist.'`" (objaw) → nazwa arkusza w kodzie nie zgadza się z nazwą w pliku: zła wielkość liter, spacja na końcu, polski znak vs jego brak, albo arkusz został zmieniony/usunięty przez kogoś innego (przyczyna) → wypisz `wb.sheetnames` **przed** odwołaniem i skopiuj nazwę dokładnie; w kodzie produkcyjnym sprawdzaj jawnie: `if "Dane" not in wb.sheetnames: raise ValueError(f"Brak arkusza 'Dane'. Dostępne: {wb.sheetnames}")` (naprawa).**

Ten komunikat jest jednym z niewielu naprawdę przyjaznych w openpyxl — zawiera pytaną nazwę. Warto go czytać, a nie zakładać „to na pewno coś innego".

**5. „Zapisałem model, w którym zmieniłem `ws2`, a zmiany nie ma w pliku" (objaw) → dwa skoroszyty, dwie zmienne; zapisano `wb1`, zmieniono `wb2` (najczęściej przez pomyłkę przy `load_workbook` w drugiej zmiennej) (przyczyna) → przy pracy z więcej niż jednym skoroszytem nadawaj zmiennym nazwy opisowe (`wb_zrodlo`, `wb_wynik`), zamiast `wb1`, `wb2`; sprawdź, na którym obiekcie faktycznie pracujesz (naprawa).**

Powiązany wariant: zapisujesz skoroszyt, który został **tylko wczytany**, żeby coś odczytać — i przy okazji przepisujesz cały plik od nowa, gubiąc elementy nieobjęte modelem (patrz moduł 18).

**6. „`PermissionError: [Errno 13] Permission denied`" przy zapisie (objaw) → plik docelowy jest otwarty w Excelu na Windows i zablokowany przez system; openpyxl nie blokuje plików i nie potrafi „poczekać", a Excel trzyma uchwyt (przyczyna) → zapisuj **do nowego pliku** i przekazuj użytkownikowi ścieżkę; jeśli to niemożliwe, wykryj sytuację i podaj czytelny komunikat: `except PermissionError: raise RuntimeError(f"Zamknij plik {cel} w Excelu i spróbuj ponownie.")` (naprawa).**

Na Linuksie/macOS ten błąd zwykle **nie wystąpi** — zapis się powiedzie, a Excel nadpisze Twoją pracę przy swoim zamknięciu. To gorsza sytuacja od błędu, bo cicha. Nigdy nie pisz do pliku, który może być otwarty.

**7. „Plik `.xlsm` otworzył się, ale makra zniknęły, a Excel ostrzega o formacie" (objaw) → wczytano plik `.xlsm` bez `keep_vba=True`, więc archiwum VBA nie zostało dołączone do modelu; zapis odtworzył pakiet bez `vbaProject.bin` (przyczyna) → wczytuj pliki z makrami jawnie: `load_workbook(path, keep_vba=True)`; dodatkowo zapisz wynik pod nazwą `.xlsm`, a nie `.xlsx`, i sprawdź `wb.vba_archive` po wczytaniu (naprawa).**

Pełne omówienie utraty części pakietu — moduł 18. Tutaj istotne jest to, że **objaw jest dwuczłonowy**: nie tylko tracisz makra, ale też Excel pokazuje ostrzeżenie o niezgodności formatu z rozszerzeniem.

**8. „Zapisałem plik z rozszerzeniem `.xls` i Excel mówi, że format nie zgadza się z rozszerzeniem" (objaw) → openpyxl nie obsługuje starego formatu BIFF; zapisze zawartość OOXML (`.xlsx`) pod dowolnym rozszerzeniem, także `.xls` (przyczyna) → używaj `.xlsx` dla zwykłych skoroszytów, `.xlsm` dla tych z makrami; jeśli potrzebujesz prawdziwego `.xls`, użyj innego narzędzia albo najpierw przekonwertuj plik (naprawa).**

Powiązany objaw: `load_workbook("plik.xls")` → `InvalidFileException` z komunikatem o nieobsługiwanym formacie. Rozróżnienie „nie umiem tego wczytać" od „nie umiem tego zapisać" jest tu istotne — w pierwszym przypadku dostajesz wyjątek, w drugim cichy problem.

**9. „`TypeError: 'Worksheet' object is not callable`" (objaw) → `wb.active` to **właściwość**, nie metoda; dodanie nawiasów próbuje **wywołać** obiekt arkusza (przyczyna) → pisz `ws = wb.active`, bez nawiasów; ta sama zasada dotyczy `wb.sheetnames`, `ws.max_row`, `ws.dimensions` (naprawa).**

Błąd pojawia się najczęściej po przejściu z bibliotek, które mają `wb.get_active_sheet()`, albo po przeczytaniu starego tutoriala z `wb.active()`.

**10. „`TypeError: Value must be either a worksheet, chartsheet or numerical index`" przy `wb.active = "Podsumowanie"` (objaw) → setter `wb.active` przyjmuje obiekt arkusza albo indeks liczbowy; **tekstem nie da się ustawić aktywnego arkusza** (przyczyna) → użyj `wb.active = wb["Podsumowanie"]` albo `wb.active = wb.sheetnames.index("Podsumowanie")` (naprawa).**

I pamiętaj o drugim warunku: `wb.active = ws` gdzie `ws` jest **ukryty** → `ValueError: Only visible sheets can be made active`.

**11. „Zmieniłem nazwę arkusza na `"Dane/2026"` i dostałem `ValueError: Invalid character / found in sheet title`" albo — gorzej — nazwa „jakoś się zmieniła" (objaw) → Excel zabrania znaków `: \ / ? * [ ]` w nazwach arkuszy i ogranicza długość do 31 znaków; openpyxl waliduje część przypadków, ale w historii zmian zmieniał zakres walidacji, a **duplikaty rozwiązuje samodzielnie** (na przykład `Dane` → `Dane1`) (przyczyna) → napisz własną funkcję walidującą nazwy (znaki, długość) i sprawdzaj kolizje przez `if nazwa in wb.sheetnames`; pełny walidator zbudujemy w module 08 (naprawa).**

Najgorszy wariant tego błędu: kod zakłada, że arkusz nazywa się `"Dane"`, a openpyxl po cichu nazwał go `"Dane1"` — i odwołanie `wb["Dane"]` działa (arkusz istnieje), ale wskazuje na **inny** arkusz niż zamierzony.

**12. „`with load_workbook("plik.xlsx") as wb:` → `TypeError: 'Workbook' object does not support the context manager protocol`" (objaw) → `Workbook` **nie implementuje** protokołu menedżera kontekstu w openpyxl 3.1.x; forma `with` na skoroszycie nie jest i nigdy nie była częścią API, mimo że pojawia się w wielu przykładach w internecie (przyczyna) → użyj jawnego `try/finally` z `wb.close()` w bloku `finally`, albo `with closing(load_workbook(...)) as wb:` z modułu `contextlib` (naprawa).**

Ten błąd ma ciekawą własność: pojawia się **natychmiast** i z jednoznacznym komunikatem, więc jest łatwy do naprawienia. Gorzej, gdy zamienisz `try/finally` na samo `wb.close()` na końcu funkcji — wtedy przy wyjątku w środku skoroszyt nie zostanie zamknięty i dostaniesz problem, którego objaw („nie mogę usunąć pliku na Windows") nie wskazuje na przyczynę.

**13. „Rozmiar pliku różni się przy każdym zapisie tego samego modelu" (objaw) → nagłówki ZIP zawierają znaczniki czasu, a kolejność części i kompresja mogą się nieznacznie różnić między zapisami (przyczyna) → nie porównuj plików `.xlsx` bajt w bajt ani nie używaj hasha pliku jako identyfikatora „czy coś się zmieniło"; porównuj **wnętrze**: nazwy części, liczbę komórek, wartości i style (moduł 20 pokaże, jak to zrobić porządnie) (naprawa).**

To istotna pułapka już teraz, bo intuicja podpowiada „ten sam model = ten sam plik". Nie jest tak i w module 20 zapłacilibyśmy za to testami, które nigdy nie przechodzą.

**14. „W wielu plikach wygenerowanych w jednym przebiegu pole `created` ma identyczną wartość, choć pliki powstawały w różnym czasie" (objaw) → w openpyxl przed 3.1.1 czas właściwości dokumentu ustawiał się raz, przy pierwszym użyciu modułu, a nie dla każdego skoroszytu (błąd naprawiony w 3.1.1) (przyczyna) → zaktualizuj openpyxl do co najmniej 3.1.3 i **ustawiaj `created` jawnie** w każdym skoroszycie, jeśli data ma znaczenie audytowe (naprawa).**

To dobry przykład szerszej zasady: **jeśli jakaś wartość ma dla Ciebie znaczenie, ustaw ją sam.** Poleganie na wartościach domyślnych biblioteki jest ryzykowne, bo mogą się zmienić między wersjami.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **`Workbook()` tworzy model w RAM, `load_workbook()` wczytuje plik do modelu, `save()` przepisuje model z powrotem do pliku.** Trzy operacje, trzy różne konsekwencje. Żadna z nich nie jest „edycją w miejscu" — to konsekwencja tego, że openpyxl implementuje **format** pliku, a nie program Excel. Bez `save()` nie istnieje żaden plik; z `save()` istnieje **kompletny nowy plik** — także wtedy, gdy część informacji nie zmieściła się w modelu.

2. **Nowy skoroszyt zawsze ma arkusz `Sheet` — bo format pliku wymaga co najmniej jednego arkusza.** `wb.active` to **właściwość** (nie metoda) wskazująca aktualnie aktywny arkusz; można ją ustawić na obiekt arkusza albo jego numer, ale **nie na tekst**. `del wb["Nazwa"]` i `wb.remove(ws)` usuwają arkusze, ale **nigdy nie należy usuwać ostatniego** — openpyxl na to pozwoli, Excel już nie.

3. **`save()` zachowuje się jak „Zapisz": nadpisuje istniejący plik bez ostrzeżenia, bez pytań i bez kopii — a plik docelowy zostaje obcięty do zera jeszcze przed zapisaniem nowej treści.** Do tego dochodzą dwa warunki brzegowe: openpyxl **nie tworzy katalogów** (musisz je utworzyć sam przez `mkdir(parents=True, exist_ok=True)`) i **nie obsługuje plików `.xls`** (ani ich nie wczyta, ani nie zapisze we właściwym formacie). Dlatego od tego modułu w kursie obowiązują dwa wzorce: **jawne ścieżki od `__file__`** oraz **zapis pod nową nazwą / z kopią zapasową**.

4. **Sprzątanie jest Twoją odpowiedzialnością, i to bezwarunkową.** `Workbook` **nie jest** menedżerem kontekstu (`with load_workbook(...)` nie zadziała), więc skoroszyt zamykasz przez `wb.close()` **w bloku `finally`** — tak, żeby zamknięcie nastąpiło także przy wyjątku, przy `return` i przy `Ctrl+C`. Ma to znaczenie krytyczne w trybie `read_only` i przy pracy w pętli po wielu plikach. Pliki robocze trzymaj w `tempfile.TemporaryDirectory()`, a pliki tymczasowe twórz w tym samym katalogu, w którym docelowo ma powstać wynik.

5. **Ten sam skoroszyt może trafić do pliku albo do pamięci — sygnatury są identyczne.** `wb.save(sciezka)` i `wb.save(BytesIO())` przyjmują ten sam typ argumentu „gdzie zapisać", a `load_workbook()` przyjmuje ścieżkę albo obiekt plikopodobny. To daje Ci trzy rzeczy naraz: **testowalność** (testy w pamięci, moduł 20), **integrację** (odpowiedź HTTP i załącznik e-mail, moduł 26) i **bezpieczeństwo** (nic nie ląduje na dysku, jeśli nie musi). A **`wb.properties`** to warstwa, którą w firmie traktuje się jak metrykę na odwrocie zdjęcia: ustawiaj w niej `creator`, `title` i `created` jawnie, bo domyślne wartości bywają mylące, a przy przepuszczaniu cudzych plików metadane przechodzą razem z treścią.

## 8. Ściągawka modułu

```python
# ==================================================================
# 0. IMPORTY I WERSJA
# ==================================================================
from pathlib import Path
from io import BytesIO
from datetime import datetime
import tempfile
import os
import shutil
import time

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

print(openpyxl.__version__)        # pytanie numer jeden przy każdym "nie działa"

# ==================================================================
# 1. ŚCIEŻKI - ZAWSZE JAWNE, OD POŁOŻENIA PLIKU
# ==================================================================
ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)     # parents=True + exist_ok=True

# Wariant dla notebooka / sesji interaktywnej (brak __file__):
def katalog_projektu() -> Path:
    try:
        return Path(__file__).resolve().parent.parent
    except NameError:
        return Path.cwd()

# Operacje na ścieżkach:
#   p.name  p.stem  p.suffix  p.parent  p.exists()  p.stat().st_size
#   p.with_name("nowa.xlsx")   p.with_suffix(".xlsm")   p / "podkatalog"

# ==================================================================
# 2. UTWORZENIE SKOROSZYTU I ZAPIS
# ==================================================================
wb = Workbook()                    # model w RAM, jeden arkusz 'Sheet'
ws = wb.active                     # WŁAŚCIWOŚĆ, nie metoda!
ws.title = "Dane"                  # nie zostawiaj technicznej nazwy 'Sheet'

ws["A1"] = "Produkt"
ws["B1"] = "Kwota"
ws["A2"] = "Widzet"
ws["B2"] = 1234.5
ws["A3"] = datetime(2026, 10, 7, 9, 30)

wb.save(OUTPUT / "raport.xlsx")    # NADPISUJE bez ostrzeżenia; katalog musi istnieć

# ==================================================================
# 3. ODCZYT
# ==================================================================
wb = load_workbook(OUTPUT / "raport.xlsx")
try:
    wb.sheetnames                  # ['Dane']        - lista TEKSTÓW
    wb.worksheets                  # [Worksheet]     - lista OBIEKTÓW (od 0)
    wb.active                      # Worksheet (aktywna zakładka przy zapisie)
    wb.active.title                # 'Dane'

    ws = wb["Dane"]                # po nazwie; KeyError gdy brak; LITERY MAJĄ ZNACZENIE
    ws = wb[0]                     # po indeksie (od 0) - kruche, gdy ktoś przestawi

    wartosc = ws["B2"].value
    typ = type(wartosc).__name__   # 'float'
    adres = ws["B2"].coordinate    # 'B2'
finally:
    wb.close()                     # NAWYK: zamykaj w finally
```

```python
# ==================================================================
# 4. STRUKTURA SKOROSZYTU
# ==================================================================
wb = Workbook()
wb.active.title = "Podsumowanie"

dane = wb.create_sheet(title="Dane")                 # NA KOŃCU (domyślnie)
konf = wb.create_sheet(title="Konfiguracja", index=0)  # NA POCZĄTKU

auto = wb.create_sheet()                             # 'Sheet1', 'Sheet2', ...
dubel = wb.create_sheet(title="Dane")                # -> 'Dane1' (cicha zmiana!)

del wb["Konfiguracja"]                               # po nazwie
wb.remove(konf)                                      # po obiekcie

# AKTYWNY ARKUSZ - trzy warianty, tylko dwa działają:
wb.active = 1                                        # indeks (OK)
wb.active = wb["Podsumowanie"]                       # obiekt Worksheet (OK w 3.1.x)
wb.active = wb.sheetnames.index("Podsumowanie")      # z nazwy na indeks (najbezpieczniej)
# wb.active = "Podsumowanie"                         # TypeError!

# WALIDACJA NAZWY NA WŁASNĄ RĘKĘ:
def unikalna_nazwa(wb, proponowana: str) -> str:
    if proponowana not in wb.sheetnames:
        return proponowana
    licznik = 1
    while f"{proponowana}_{licznik}" in wb.sheetnames:
        licznik += 1
    return f"{proponowana}_{licznik}"

# ZASADA: nigdy nie usuwaj ostatniego arkusza.

# ==================================================================
# 5. ZAPIS BEZ NADPISANIA - WZORZEC MINIMALNY
# ==================================================================
def zapisz_bez_pomylek(wb, sciezka, *, nadpisz: bool = False) -> Path:
    cel = Path(sciezka).expanduser().resolve()
    cel.parent.mkdir(parents=True, exist_ok=True)

    if cel.exists() and not nadpisz:
        raise FileExistsError(f"Plik istnieje: {cel}. Podaj inną nazwę lub nadpisz=True.")

    deskryptor, tmp_nazwa = tempfile.mkstemp(
        prefix=f".{cel.stem}_", suffix=".tmp", dir=str(cel.parent)
    )
    os.close(deskryptor)
    plik_tmp = Path(tmp_nazwa)
    try:
        wb.save(plik_tmp)
        os.replace(plik_tmp, cel)          # atomowa podmiana nazwy
    except BaseException:
        plik_tmp.unlink(missing_ok=True)
        raise
    return cel

# ==================================================================
# 6. ZAPIS I ODCZYT W PAMIĘCI (BytesIO)
# ==================================================================
bufor = BytesIO()
wb.save(bufor)                             # ten sam typ celu co ścieżka
dane = bufor.getvalue()                    # bytes; openpyxl NIE zamyka bufora
print(dane[:2])                            # b'PK' - to archiwum ZIP (moduł 02)

wb2 = load_workbook(BytesIO(dane))         # round-trip bez dysku
try:
    print(wb2.sheetnames)
finally:
    wb2.close()

# Ponowne użycie tego samego bufora:
bufor.seek(0)
wb3 = load_workbook(bufor)
wb3.close()

# ==================================================================
# 7. PLIKI TYMCZASOWE
# ==================================================================
with tempfile.TemporaryDirectory() as tmp:
    roboczy = Path(tmp) / "roboczy.xlsx"
    wb.save(roboczy)
    # ... praca ...
# <- cały katalog zostaje usunięty, także przy wyjątku

# Uwaga na Windows: NamedTemporaryFile trzyma plik ZABLOKOWANY.
# Jeśli potrzebujesz nazwy, użyj mkstemp + os.close(fd) albo TemporaryDirectory.

# ==================================================================
# 8. WŁAŚCIWOŚCI DOKUMENTU (ŚLAD AUDYTOWY)
# ==================================================================
wb.properties.title = "Raport sprzedaży — październik 2026"
wb.properties.creator = f"generator-raportow/1.4.0 (openpyxl {openpyxl.__version__})"
wb.properties.lastModifiedBy = "automat: generator-raportow/1.4.0"
wb.properties.subject = "Sprzedaż — październik 2026"
wb.properties.description = "Dokument wygenerowany automatycznie. Zmiany ręczne zostaną nadpisane."
wb.properties.keywords = "raport; sprzedaz; automatyzacja"
wb.properties.category = "Raport wewnętrzny"
wb.properties.created = datetime(2026, 10, 7, 9, 30, 0)   # data NAIWNA, bez strefy

# UWAGA: przy przepuszczaniu cudzego pliku metadane przechodzą z oryginału.
# Ustawiaj creator/lastModifiedBy JAWNIE w każdym generowanym raporcie.

# ==================================================================
# 9. OBSŁUGA BŁĘDÓW - CO ŁAPAĆ
# ==================================================================
try:
    wb = load_workbook("plik.xlsx")
except FileNotFoundError:
    ...                                     # nie ma takiego pliku / katalogu
except InvalidFileException as blad:
    ...                                     # np. stary format .xls
except KeyError as blad:
    ...                                     # brak arkusza w wb["
    # "]
finally:
    ...

try:
    wb.save(cel)
except PermissionError:
    ...                                     # Windows: plik otwarty w Excelu
except FileNotFoundError:
    ...                                     # nie istnieje katalog docelowy
except OSError as blad:
    ...                                     # brak miejsca, uprawnienia, inny błąd systemu

# ==================================================================
# 10. SZKIELET DOBREGO SKRYPTU
# ==================================================================
# 1. importy
# 2. ROOT / OUTPUT od __file__ + mkdir(parents=True, exist_ok=True)
# 3. funkcja budująca model (NIE zapisuje nigdzie)
# 4. funkcja zapisująca (tworzy katalog, nie nadpisuje po cichu)
# 5. main() - złożenie
# 6. if __name__ == "__main__": main()
```

## 9. Co dalej

Masz już pełny, działający cykl: model → plik → model. Wiesz, jak budować ścieżki, jak nie zniszczyć cudzej pracy, jak zamknąć to, co otworzyłeś, i jak zapisać skoroszyt tam, gdzie nigdy nie dotknie dysku. To wszystko jest fundamentem — w każdym kolejnym module będziesz używać dokładnie tych pięciu wzorców:

```text
ROOT/OUTPUT od __file__           -> moduły 05-27
mkdir(parents=True, exist_ok=..)  -> moduły 05-27
try/finally + wb.close()          -> moduły 05-27
zapis pod nową nazwą               -> moduł 18 (w pełnej wersji)
wb.properties                     -> moduły 18, 26, 27
```

W **module 05** zejdziemy o poziom niżej: przestaniemy mówić „komórka ma wartość" i zaczniemy mówić **jaką wartość i jakiego typu**. To moduł, który zaskakuje najbardziej doświadczonych, bo dotyczy rzeczy, których się nie widzi: różnicy między pustą komórką a `None`, liczbą całkowitą, która przychodzi jako `float`, tekstem, który wygląda jak liczba, oraz datą, która w Excelu jest **liczbą**, a w Pythonie obiektem `datetime`. Tam też wyjaśnimy, dlaczego `#N/A` wczytany z pliku jest **tekstem**, a nie wyjątkiem — i dlaczego to jest ważne, jeśli walidujesz dane klienta.

Przygotuj do modułu 05:

- **pliki z `output/`**, które powstały w tym module — będziesz na nich testować konwersje typów,
- **jeden plik utworzony ręcznie w Excelu**, w którym w kolumnie są liczby sformatowane jako waluta, w drugiej daty, a w trzeciej kod pocztowy z zerem na początku. To ostatnie pole będzie Twoim poligonem doświadczalnym — i zobaczysz, co się z nim stanie po przejściu przez openpyxl,
- **odpowiedź na pytanie**: w których miejscach kodu z tego modułu zmieniałeś **model**, a w których dotykałeś **pliku**? Powinieneś umieć wskazać palcem każdą linię.