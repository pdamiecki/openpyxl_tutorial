# Moduł 05 — Komórki i wartości

> **Część:** I — Podstawy pracy ze skoroszytem · **Poziom:** ⭐ · **Wymaga:** modułów 00–04

## 0. W tym module nauczysz się

- Rozumiesz, że **komórka to obiekt, a nie pudełko na tekst** — i że `ws["A1"]` zwraca obiekt `Cell`, którego treścią jest `cell.value`. To rozróżnienie wyjaśnia połowę błędów z tego modułu.
- Znasz **pełną listę typów**, jakie możesz wpisać do komórki (`str`, `int`, `float`, `bool`, `None`, `datetime`, `date`, `time`, `timedelta`, `Decimal`) i wiesz, dlaczego `list`, `dict`, `set` ani `numpy.int64` się nie zmieszczą.
- Odróżniasz **pustą komórkę (`None`)** od **pustego tekstu (`""`)** od **zera (`0`)** — i rozumiesz konsekwencje tego rozróżnienia dla `ISBLANK`, `COUNTA`, `COUNTBLANK` i `AVERAGE` oraz dla Twoich warunków w Pythonie.
- Wiesz, **skąd bierze się `"1"` tam, gdzie spodziewałeś się `1`** (i odwrotnie), dlaczego długie identyfikatory tracą cyfry, a kody pocztowe tracą zera — i umiesz to naprawić.
- Rozumiesz, że **data w pliku Excela jest liczbą** (serialem dni), co znaczy `wb.epoch`, czym różni się `CALENDAR_WINDOWS_1900` od `CALENDAR_MAC_1904` i dlaczego `datetime.date` wraca z pliku jako `datetime.datetime`.
- Wiesz, że **wartości błędów (`#N/A`, `#DIV/0!`) wczytują się jako tekst**, a nie wyjątki — i jak je wykrywać niezawodnie (`cell.data_type`).
- Umiesz czytać dane **wierszami** (`ws.values`, `iter_rows(values_only=True)`) i przepuszczać je przez warstwę mapowania (`bezpieczna_wartosc()`), zamiast rozsypywać `is None` po całym kodzie.

## 1. Intuicja i analogia

W module 04 nauczyłeś się tworzyć plik. Ale plik, w którym nie ma **wiarygodnych danych**, jest tylko ładnie opakowanym problemem. Ten moduł jest o danych — a konkretnie o jednej rzeczy, która zaskakuje nawet ludzi z kilkuletnim stażem: **Excel i Python mają różne modele typów, a openpyxl jest tłumaczem poruszającym się między nimi**. Tłumacz bywa wierny, bywa twórczy, a czasem coś gubi. Ten moduł uczy Cię rozpoznawać, w którym z tych trzech trybów właśnie pracujesz.

### Analogia pierwsza: magazynier i pudełka z etykietami

Wyobraź sobie bardzo uporządkowanego magazyniera. Ma regał z komórkami (ha — komórkami!). Do każdej komórki może włożyć **dokładnie jedną rzecz**, ale tylko z zamkniętej listy asortymentu: tekst, liczbę całkowitą, liczbę zmiennoprzecinkową, prawda/fałsz, datę, godzinę, czas trwania. Każda z tych rzeczy dostaje **inną etykietę** („to jest tekst", „to jest liczba"). Magazynier nie przyjmie na magazyn słownika, listy ani własnego obiektu Twojej klasy. Jeśli podasz mu coś spoza listy, wywali Cię z magazynu z komunikatem.

Ten magazynier to **openpyxl**. Regał to **arkusz**. Komórka to **jedna przegródka**. A etykieta to **`cell.data_type`** — jednoliterowy kod, który mówi, co naprawdę leży w środku.

Dlaczego to takie ważne? Bo **Python ma na to samo inne etykiety**. `datetime.date` w Pythonie jest osobnym typem od `datetime.datetime`. Excel nie zna takiego rozróżnienia — dla niego oba są tą samą liczbą dni od umownej daty. `Decimal` w Pythonie potrafi trzymać nieskończenie wiele cyfr po przecinku. Excel trzyma **15 cyfr znaczących** i ani jednej więcej. Różnice na wejściu i wyjściu z magazynu są więc nieuniknione — i trzeba je znać, zamiast się na nie natykać.

### Analogia druga: puste krzesło na sali

Wyobraź sobie salę, na której mają usiąść ludzie. Są trzy sytuacje, które z daleka wyglądają podobnie:

1. **Nie ma krzesła.** — Tak naprawdę to „brak miejsca". W Excelu: komórka pusta, w której **nic nigdy nie było**.
2. **Jest krzesło, ale nikt na nim nie siedzi.** — Miejsce istnieje, jest przygotowane, ale puste. W Excelu: komórka z **pustym tekstem `""`** — ma wartość, tylko ta wartość jest zerowej długości.
3. **Na krześle siedzi ktoś, kto się nie odzywa — i jego obecność liczy się jako liczba obecnych.** W Excelu: **`0`** albo **`FALSE`**. To nie jest brak danych. To pełnoprawna wartość.

Zastanów się, co zrobić z pytaniem „ile osób jest na sali?". Zależy, jakiej metody użyjesz:

- **`COUNTA`** odpowiada na pytanie „ile miejsc jest zajętych czymkolwiek?" — liczy krzesła z ludźmi, ale też liczy **puste krzesło, jeśli jest to krzesło przygotowane** (bo jest „niepuste" — ma wartość, choć pustą).
- **`COUNTBLANK`** odpowiada na pytanie „ile miejsc wygląda na wolne?" — liczy zarówno brak krzesła, jak i krzesło przygotowane, ale nieobsadzone.
- **`ISBLANK`** odpowiada precyzyjnie na pytanie „czy tu jest krzesło?" — `TRUE` tylko wtedy, gdy nie ma nic.
- **`AVERAGE`** odpowiada na pytanie „jaka jest średnia liczba osób na krzesło?" — pomija te puste miejsca całkowicie, ale **osobę milczącą (zero) policzy jako zero**, obniżając średnią.

To jest ta różnica, która kosztuje ludzi najwięcej czasu w Excelu. Komórka z zerem i komórka pusta wyglądają niemal identycznie na wydruku, jeśli zero jest sformatowane tak, by się nie pokazywało — a dają zupełnie inne wyniki w każdej formule.

**Wniosek dla Pythona:** skoro te trzy rzeczy są różne w Excelu, to są też różne w Twoim kodzie. `if not cell.value:` potraktuje `None`, `""` i `0` **identycznie** — a Ty prawie nigdy nie chcesz tego zrobić. Pisz `if cell.value is None:` i bądź świadomy, co odrzucasz.

### Analogia trzecia: numer telefonu w notatniku i w kalkulatorze

Zapisz numer telefonu `0048 512 345 678`. W **notatniku** to ciąg znaków — liczy się każdy znak, każde zero na początku, każda spacja. W **kalkulatorze** ten sam ciąg da liczbę `48512345678`, bo kalkulator nie ma pojęcia, że początkowe zera były ważne.

Excel to kalkulator z ambicjami arkusza kalkulacyjnego. Gdy wpiszesz do komórki coś, co wygląda jak liczba, **Excel traktuje to jako liczbę** — z zerami wiodącymi odciętymi i zaokrągleniem do 15 cyfr znaczących. I robi to niezależnie od tego, czy Ty tego chciałeś. Kod pocztowy `00-950` podzieli się na „datę" i liczbę. Numer ewidencyjny o 18 cyfrach zostanie zaokrąglony.

Rozwiązanie brzmi: **dane, które tylko wyglądają jak liczby, ale którymi nie będziesz liczyć, zapisuj jako tekst** i to jawnie, w kodzie. `ws["A1"] = "007"` — z cudzysłowem. Ten moduł pokaże Ci, jak wykrywać, że coś poszło nie tak, i jak to naprawić w już istniejącym pliku.

### Analogia czwarta: data jako licznik dni

Zapamiętaj jedno zdanie, a zrozumiesz cały podrozdział o datach: **Excel nie przechowuje daty. Excel przechowuje liczbę dni, które upłynęły od umownego momentu, i tylko wyświetla ją w formie daty.**

To jak **licznik dni w kalendarzu biurowym**: nie ma tam napisu „7 października 2026", jest numer strony — 46 302. A to, że wyświetla się „7 października 2026", to zasługa **formatu wyświetlania**, czyli tej samej warstwy, która sprawia, że liczba `0,25` pokazuje się jako `25%`.

Z tego wynikają trzy rzeczy, które wrócą w tym module:

1. **Format liczby i wartość liczby to dwie rozłączne rzeczy.** Zaokrąglasz wyświetlanie, nie wartość.
2. **openpyxl musi zgadnąć, czy liczba to data.** Robi to, patrząc na format komórki. Liczba z formatem `0,00` to liczba. Ta sama liczba z formatem `yyyy-mm-dd` to data.
3. **Data bez formatu to dla człowieka katastrofa.** Jeśli zapiszesz datę i nie zadbasz o format, użytkownik zobaczy `46302,39583` i zadzwoni do Ciebie. To jedna z tych pomyłek, które zdarzają się nawet doświadczonym — i której poświęcimy w tym module spory fragment.

## 2. Teoria

### 2.1. Komórka to obiekt — `ws["A1"]` a `ws["A1"].value`

Zacznijmy od rozróżnienia, które trzeba sobie naprawdę przyswoić:

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

uchwyt = ws["A1"]          # to jest OBIEKT Cell
print(type(uchwyt))        # <class 'openpyxl.cell.cell.Cell'>
print(uchwyt.value)        # None   <- dopiero to jest TREŚĆ

ws["A1"] = "Widzet"        # to jest przypisanie do .value, w skrócie
print(ws["A1"].value)      # 'Widzet'
```

**Analogia: szafka i jej zawartość.** `ws["A1"]` to **uchwyt do szafki**. `ws["A1"].value` to **zawartość szafki**. `ws["A1"] = x` to skrót od `ws["A1"].value = x` — Python pozwala tak zrobić dzięki mechanizmowi zwanemu **setterem właściwości** (metoda, którą interpreter wywołuje automatycznie przy przypisaniu; spotkałeś się z nim w module 04 przy `wb.active`).

Komórka to pełnoprawny obiekt z własnymi właściwościami. Najważniejsze:

```python
komorka = ws["A1"]
komorka.value              # treść (str, int, float, ... albo None)
komorka.data_type          # 's' (string), 'n' (number), 'd' (date), 'b' (bool), 'f' (formuła), 'e' (błąd)
komorka.coordinate         # 'A1'   - adres w notacji A1
komorka.row                # 1      - numer wiersza (od 1!)
komorka.column             # 1      - numer kolumny (od 1!)
komorka.number_format      # 'General'
komorka.is_date            # True, jeśli format jest formatem daty
komorka.parent             # obiekt Worksheet, do którego należy
```

Zapamiętaj kierunek zależności: **komórka wie, do jakiego arkusza należy** (`cell.parent` → `Worksheet`, a `cell.parent.parent` → `Workbook`). Nie odwrotnie. Nie ma czegoś takiego jak „arkusz przypisany do komórki" — komórka zawsze należy do dokładnie jednego arkusza i nie da się jej przenieść.

**Konsekwencja praktyczna nr 1: każda komórka, której dotkniesz, zaczyna istnieć w modelu.** To jedna z tych rzeczy, które nie robią nic złego w małym pliku, a w dużym robią różnicę: gdy wywołasz `ws["ZZ1000"]` na pustym arkuszu, openpyxl **tworzy** obiekt komórki w słowniku `_cells` i zapamiętuje go. Od tego momentu może wpływać na to, co openpyxl uważa za „rozmiar danych" w arkuszu. Wrócimy do tego w module 07 (sekcja o `max_row`), ale już teraz zapamiętaj: **czytanie komórki to też modyfikacja modelu** (niewidzialna w pliku, ale istniejąca w pamięci).

**Konsekwencja praktyczna nr 2: `cena = ws["B2"]` to prawdopodobnie błąd.** Masz w zmiennej obiekt komórki, nie liczbę. Jeśli chcesz liczbę, napisz `cena = ws["B2"].value`. Ten błąd jest o tyle podstępny, że obiekt komórki często zachowuje się „prawie jak" wartość — na przykład `print()` pokaże coś sensownego — ale przy dodawaniu dostaniesz `TypeError: unsupported operand type(s)`.

### 2.2. Dwa sposoby adresowania: notacja A1 i para `(row, column)`

Do tej samej komórki możesz dotrzeć na dwa sposoby:

```python
ws["B3"] = 100                            # notacja A1 - litera kolumny + numer wiersza
ws.cell(row=3, column=2, value=100)       # notacja liczbowa - (wiersz, kolumna), obie od 1
```

**Analogia: adres pocztowy i współrzędne GPS.** Adres („Kwiatowa 3, mieszkanie 5") jest zrozumiały dla człowieka — od razu wiesz, gdzie to jest. Współrzędne (`52.2297, 21.0122`) są zrozumiałe dla urządzenia — można je policzyć, dodać, przeliczyć. Oba opisują to samo miejsce, ale inne zadania wykonuje się wygodniej w jednym z nich.

**Kiedy A1?** Gdy adres jest stały i znany z góry — nagłówki, komórki konfiguracyjne, konkretne pola formularza, tytuł raportu. `ws["A1"] = "Raport"` czyta się jak zdanie.

**Kiedy `(row, column)`?** Gdy adres jest **wynikiem obliczeń**: pętla po wierszach, wyszukiwanie kolumny, przeliczanie indeksu. W pętli `ws.cell(row=i, column=j)` jest naturalne, a składanie tekstu `ws[f"{litera}{i}"]` — żmudne i wolniejsze.

```python
# Źle: składanie adresu z tekstu w gorącej pętli
for i in range(2, 1000):
    ws[f"B{i}"] = dane[i]

# Dobrze: adres liczbowy, bez konwersji
for i in range(2, 1000):
    ws.cell(row=i, column=2, value=dane[i])

# Znakomicie: append, gdy dane tworzą regularną tabelę (szczegóły w module 07)
for wiersz in dane:
    ws.append(wiersz)
```

Trzy rzeczy, o których warto wiedzieć przy `ws.cell()`:

**`ws.cell()` przyjmuje `value=`.** To nie jest zwykły getter — możesz ustawić wartość w tym samym wywołaniu. Działa to trochę jak `dict.get(key, default)`, tylko dla komórek.

**`ws.cell()` nie tworzy nowego arkusza ani nie sprawdza granic sensownie.** Wywołanie `ws.cell(row=0, column=1)` lub `ws.cell(row=1, column=0)` rzuci `ValueError`, bo wiersze i kolumny numeruje się **od 1**. Granice arkusza to 1–1 048 576 wierszy i 1–16 384 kolumny (A do XFD).

**`ws["B3"]` to nie „string w nawiasach" a specjalny mechanizm.** Obiekt `Worksheet` implementuje metodę `__getitem__`/`__setitem__`, która rozpoznaje zarówno pojedyncze adresy (`"B3"`), jak i **zakresy** (`"A1:C10"`) oraz **całe kolumny i wiersze** (`"B"`, `"3"`):

```python
ws["A1:C3"]     # krotka krotek komórek - 3 wiersze po 3 komórki
ws["B"]         # krotka komórek z kolumny B (od wiersza 1 do max_row)
ws["3"]         # krotka komórek z wiersza 3 (od kolumny 1 do max_column)
```

To wygodnE, ale ma pułapkę: `ws["B"]` na arkuszu, w którym dane sięgają wiersza 100, zwróci też puste komórki od 101 w górę — jeśli wcześniej dotknąłeś większych adresów. Zakresy zależą od aktualnego rozmiaru modelu, nie od tego, co „naprawdę" ma dane.

### 2.3. Co dokładnie może leżeć w komórce — macierz typów

Oto pełna, praktyczna lista. Wszystko, co nie jest na tej liście, zostanie odrzucone (patrz 2.4).

| Typ Pythona | Przykład | `data_type` w pliku | Co zobaczysz po odczycie | Uwagi |
|---|---|---|---|---|
| `str` | `"Widzet"` | `'s'` (shared string) | `'Widzet'` — ten sam `str` | Długość ≤ 32 767 znaków; `\n` wymaga `wrap_text` |
| `int` | `42` | `'n'` | `42` — `int` | Powyżej 15 cyfr znaczących Excel zaokrągla |
| `float` | `3.14` | `'n'` | `3.14` — `float` | Standard IEEE 754; `0.1 + 0.2 ≠ 0.3` |
| `bool` | `True` | `'b'` | `True` — `bool` | W XML to `1`/`0`, ale wraca jako `bool` |
| `None` | `None` | (brak wpisu) | `None` — komórka pusta | Komórka bez wartości i bez stylu nie trafia do XML |
| `datetime.datetime` | `datetime(2026,10,7,9,30)` | `'d'` → serial | `datetime` | Wymaga formatu daty; bez strefy czasowej |
| `datetime.date` | `date(2026,10,7)` | `'d'` → serial | **`datetime`**, nie `date`! | Excel nie odróżnia daty od północy |
| `datetime.time` | `time(9,30)` | `'d'` → ułamek | `time` | Wartość z przedziału 0–1 dnia |
| `datetime.timedelta` | `timedelta(hours=27)` | `'d'` → dni | `timedelta` | Format `[hh]:mm:ss`; przekracza 24 h |
| `Decimal` | `Decimal("19.99")` | `'n'` | `19.99` — **`float`** | Zapisywany jako liczba; wraca jako `float` |

Trzy wnioski z tej tabeli, które warto od razu wyciągnąć:

**`Decimal` nie wraca jako `Decimal`.** Otrzymałeś akceptację typu, ale w pliku jest zwykła liczba zmiennoprzecinkowa. Jeśli więc liczysz pieniądze i chcesz precyzji, **rób całą arytmetykę w `Decimal` przed zapisem**, a do komórki wstawiaj już policzoną wartość. Nigdy nie odczytuj kwot z pliku z powrotem na `Decimal` z nadzieją na zachowanie precyzji — dostaniesz `float`.

**`date` wraca jako `datetime`.** To nie błąd biblioteki, to konsekwencja modelu Excela z sekcji 1. Jeśli Twój kod dalej porównuje `if komorka.value == date(2026, 10, 7):` — to porównanie **nie zadziała**, bo `datetime(2026,10,7,0,0) != date(2026,10,7)`. Zapisuj i odczytuj konsekwentnie ten sam typ, albo normalizuj na wejściu do logiki biznesowej.

**`None` to jedyna „prawdziwa pustka".** Zaraz się tym zajmiemy szczegółowo, bo to najważniejsza rzecz w module.

### 2.4. Typy spoza listy — `ValueError` i jak go naprawić

Jeśli spróbujesz wpisać coś, czego magazynier nie zna, dostaniesz wyjątek:

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ws["A1"] = {"klient": "ACME", "kwota": 100}
```

```text
ValueError: Cannot convert {'klient': 'ACME', 'kwota': 100} to Excel
```

Komunikat jest dosłowny i podaje wartość — w dużych plikach bywa długi, ale od razu mówi, co go zabolało. Ten sam błąd dostaniesz dla `list`, `set`, `tuple`, obiektów własnych klas, a także — co jest najczęstszym realnym przypadkiem — dla **typów z numpy**:

```python
import numpy as np

ws["A1"] = np.int64(42)
# ValueError: Cannot convert 42 to Excel
```

To boli podwójnie, bo komunikat pokazuje `42`, patrzysz na `42` i myślisz „przecież to liczba!". I masz rację — ale to liczba z **innej rodziny**. `numpy.int64` nie dziedziczy po wbudowanym `int` w Pythonie, a openpyxl sprawdza typ przez `isinstance`. Ciekawostka: `np.float64` **przechodzi**, bo dziedziczy po `float`. To dokładnie ten rodzaj niespójności, który sprawia, że błędy typów w openpyxl są tak frustrujące dla ludzi pracujących z pandas.

**Cztery wzorce naprawy** — od najprostszego do najbardziej „architektonicznego":

```python
# 1. Konwersja w miejscu - gdy wartość naprawdę jest liczbą
ws["A1"] = int(np.int64(42))            # int() sprowadza do wbudowanego typu

# 2. Serializacja - gdy wartość jest strukturą, którą chcesz zachować
import json
ws["A2"] = json.dumps({"klient": "ACME", "kwota": 100})   # czytelny tekst JSON

# 3. Reprezentacja tekstowa - gdy struktura ma być tylko widoczna
ws["A3"] = ", ".join(str(x) for x in [1, 2, 3])           # '1, 2, 3'

# 4. Rozbicie na kolumny - gdy struktura ma sens biznesowy
#    Zamiast wciskać słownik w jedną komórkę, zrób z jego kluczy nagłówki.
```

**Wzorzec 4 jest właściwym rozwiązaniem w 90% przypadków.** Jeśli masz listę słowników, w których każdy słownik to jeden rekord — nie serializuj ich do JSON-a „bo tak jest szybciej". Zapisz je jako **tabelę**, w której klucze są nagłówkami, a wartości kolumnami. Excel to arkusz kalkulacyjny; struktura danych powinna przypominać tabelę, nie drzewo JSON. JSON w komórce to znak, że dane są w złym miejscu — a nie w złym formacie.

Dodatkowa uwaga: `bytes` też nie wejdzie wprost. Zdekoduj go jawnie: `b"tekst".decode("utf-8")`. Jeśli masz dane binarne (obraz, PDF), nie należą do komórki — mają swoje miejsce w `xl/media/` (moduł 16).

### 2.5. `None`, `""` i `0` — trzy różne rzeczy, jedno słowo „puste"

To najważniejszy podrozdział w całym module. Wracamy do analogii z krzesłem i przekładamy ją na kod.

```python
from openpyxl import Workbook, load_workbook
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "Puste"

# Trzy komórki, które "wyglądają" tak samo na pierwszy rzut oka:
ws["B1"] = None         # pustka - komórka bez wartości
ws["B2"] = ""           # pusty TEKST - wartość zerowej długości
ws["B3"] = 0            # ZERO - pełnoprawna liczba

cel = OUTPUT / "05_puste.xlsx"
wb.save(cel)

kontrola = load_workbook(cel)
try:
    for adres in ("B1", "B2", "B3"):
        k = kontrola["Puste"][adres]
        print(f"{adres}: value={k.value!r:>6}  typ={type(k.value).__name__:<5}  data_type={k.data_type!r}")
finally:
    kontrola.close()
```

```text
B1: value=  None  typ=NoneType  data_type='n'
B2: value=    ''  typ=str       data_type='s'
B3: value=     0  typ=int       data_type='n'
```

Trzy różne stany, trzy różne etykiety. `None` to brak wartości. `""` to wartość istniejąca, ale zerowej długości. `0` to wartość liczbowa.

**Co się dzieje w pamięci.** `ws["B1"] = None` ustawia `value` na `None` i `data_type` na `'n'` (domyślna wartość, bo openpyxl nie ma osobnego kodu dla pustki). Komórka istnieje w `_cells`, ale nie ma treści. `ws["B2"] = ""` ma `data_type = 's'` i wartość `""`.

**Co trafi do pliku.** To jest clou: **komórka bez wartości i bez stylu nie jest w ogóle zapisywana do XML**. Nie ma jej w `sheet1.xml`. Z punktu widzenia pliku komórka `B1` „nie istnieje" — dokładnie tak, jak puste miejsce na sali, na którym nie ma nawet krzesła. Natomiast `B2` zostanie zapisana jako komórka tekstowa ze wskaźnikiem do pustego łańcucha w `sharedStrings.xml`. To komórka, na której stoi krzesło — ale nikt nie siedzi.

Efekt jest taki:

| Adres | `ISBLANK` | `COUNTA` | `COUNTBLANK` | `AVERAGE(B1:B3)` |
|---|---|---|---|---|
| `B1` = brak wartości | `TRUE` | nie liczy | liczy | pomija |
| `B2` = `""` | `FALSE` | **liczy** | **liczy** | pomija |
| `B3` = `0` | `FALSE` | liczy | nie liczy | **liczy jako 0** |
| Razem | — | **2** | **2** | **0** |

Zwróć uwagę na dwa wyniki, które wyglądają na sprzeczne:

- `COUNTA` zwraca `2`, bo **pusty łańcuch nie jest pustką** — komórka jest „niepusta".
- `COUNTBLANK` też zwraca `2`, bo jego definicja obejmuje **zarówno komórki naprawdę puste, jak i te zawierające pusty tekst**. To jedyna formuła o takiej „szerokiej" definicji pustki i źródło niezliczonych nieporozumień w raportach.

Otwórz wygenerowany plik w Excelu i sprawdź te wartości na własne oczy. To jedna z tych rzeczy, które warto zobaczyć raz, żeby nigdy nie pomylić ich ponownie.

#### Konsekwencje dla kodu w Pythonie

```python
# ŹLE - wszystkie trzy przypadki wpadają w ten sam warunek
if not komorka.value:
    pomin()                      # pomija None, "" ORAZ 0

# DOBRZE - jawnie mówisz, co uznajesz za brak danych
if komorka.value is None:
    brak_danych()

# DOBRZE - gdy pusty tekst też jest "brakiem"
if komorka.value is None or komorka.value == "":
    brak_danych()

# DOBRZE - gdy chcesz tylko liczby
if isinstance(komorka.value, (int, float)) and not isinstance(komorka.value, bool):
    licz(komorka.value)
```

Trzeci przykład zasługuje na wyjaśnienie: `bool` jest w Pythonie podklasą `int`, więc `isinstance(True, int)` zwraca `True`. Chcesz więc **jawnie wykluczyć** boole, gdy sprawdzasz „czy to liczba". Inaczej prawda/fałsz wliczy się do sumy jako 1/0 — co zresztą w Excelu jest poprawne, ale w Twojej logice biznesowej prawdopodobnie nie.

#### Czyszczenie komórki

```python
komorka.value = None      # usuwa wartość, zostawia styl
komorka.value = ""        # wstawia pusty tekst - komórka NIE jest pusta w rozumieniu Excela
```

To rozróżnienie jest istotne, gdy „czyścisz" szablon przed wypełnieniem. Jeśli w szablonie było zero, a Ty ustawisz `""`, użytkownik dostanie komórkę, którą Excel będzie liczył w `COUNTA`, choć wygląda na pustą. Ustawiaj `None`, nie `""`.

### 2.6. Konwersje przy odczycie — `int` kontra `float`, `"1"` kontra `1`

#### Dlaczego czasem dostajesz `float`, a czasem `int`

openpyxl czyta z pliku **tekst** i musi zdecydować, czy zamienić go na liczbę całkowitą, czy zmiennoprzecinkową. Robi to po prostu patrząc na zapis: jeśli w ciągu jest kropka albo wykładnik, używa `float()`, w przeciwnym razie `int()`.

Praktyczny wniosek: **`23` zostanie odczytane jako `23`, a `23,0` jako `23.0`**. Jeśli więc w kolumnie masz kwoty, do których ktoś ręcznie wpisał „23,0" (a Excel zapisał to jako 23,0), dostaniesz `float`. Twoja walidacja „typ ma być `int`" polegnie na pozornie identycznych danych. Waliduj **wartość**, nie typ: `if float(v).is_integer():`.

Druga pułapka jest subtelniejsza: kwoty pochodzące z **formuł** i kwoty z **cache formuł** mogą przyjść jako `float`, nawet jeśli wizualnie są całkowite. I trzecia, praktyczna: jeśli komórka ma format z odejmowaniem, zaokrągleniem albo przełączaniem jednostek, **wartość w komórce nie ma nic wspólnego z tym, co widzisz w Excelu**. Format `#,##0,"tys."` sprawia, że widzisz `1 235`, a w komórce jest `1234567`. Zawsze operuj na wartości, nigdy na wyglądzie.

#### Dlaczego długie liczby tracą cyfry

Excel przechowuje liczby jako liczby zmiennoprzecinkowe podwójnej precyzji i **wyświetla maksymalnie 15 cyfr znaczących**. Numer faktury `123456789012345678` zostanie:

1. zapisany przez openpyxl do XML **w pełnej postaci** (bo `str()` na `int` nie traci cyfr),
2. **odczytany przez openpyxl z powrotem jako pełna liczba** — jeśli nikt w międzyczasie nie otwierał pliku w Excelu,
3. **zmieniony przez Excel** w momencie otwarcia i ponownego zapisu — na `1,23456789012346E+17`.

To wyjątkowo zdradliwy przypadek: **ten sam plik może pokazać różną wartość zależnie od tego, kto go czyta**. Openpyxl zwraca to, co jest w XML; Excel zwraca to, co jego model liczbowy potrafi zmieścić. Nie ma tu dobrego rozwiązania poza jednym: **identyfikatory zapisuj jako tekst**. Numer faktury, NIP, EAN, IBAN, numer przesyłki — wszystko to są **etykiety**, a nie liczby. Nie będziesz ich dodawać ani uśredniać.

#### Kiedy `1` staje się `"1"` i odwrotnie

Dwa różne scenariusze, dwie różne przyczyny:

**Scenariusz A: liczba stała się tekstem.** Dzieje się tak, gdy ktoś „dla bezpieczeństwa" zamienił wartość na tekst — albo gdy dane przyszły z CSV i każda kolumna była tekstem. Objaw: w Excelu liczby są wyrównane do lewej, w rogu komórki jest zielony trójkąt („Liczba zapisana jako tekst"), a `SUMA` zwraca `0`.

**Scenariusz B: tekst stał się liczbą.** Dzieje się tak, gdy ktoś wpisał lub wkleił kod pocztowy `00-950`, a Excel rozpoznał w tym datę (i wyświetla `00-950` jako... cokolwiek, zależy od interpretera) albo gdy wkleił `007` i dostał `7`. Objaw: zniknęły zera wiodące, wartości się „przesunęły" o jeden dzień, jeśli to data.

**Detektor, którym posłużysz się w praktyce — wyrównanie w komórce.** Excel domyślnie wyrównuje:

- **tekst → do lewej**,
- **liczby i daty → do prawej**,
- **wartości logiczne → do środka**.

Patrząc na kolumnę, w której wszystkiego spodziewałeś się po prawej, a jest po lewej — wiesz, że to tekst. To najszybsza diagnostyka na świecie i nie wymaga ani linijki kodu.

**Detektor programistyczny — `data_type`:**

```python
if komorka.data_type == "s" and komorka.value.strip().replace(",", ".").lstrip("-").replace(".", "", 1).isdigit():
    print(f"{komorka.coordinate}: liczba ukryta w tekście -> {komorka.value!r}")
```

(`data_type == 's'` znaczy „tekst", `'n'` — „liczba", `'d'` — „data", `'b'` — „boolean", `'e'` — „błąd", `'f'` — „formuła".)

### 2.7. `bool` — prawda, fałsz i cicha arytmetyka

Wartości logiczne w Excelu to `PRAWDA`/`FAŁSZ` (albo `TRUE`/`FALSE` w wersji angielskiej) i są osobnym typem danych — nie tekstem, nie liczbą. openpyxl zapisuje je poprawnie:

```python
ws["A1"] = True       # data_type = 'b'; w XML <c t="b"><v>1</v></c>
ws["A2"] = False      # data_type = 'b'; w XML <c t="b"><v>0</v></c>
```

Po odczycie dostajesz `True` i `False` — nie liczby. To dobra wiadomość.

**Trzy pułapki, o których warto wiedzieć:**

**Pułapka 1: `ws["A1"] = "PRAWDA"` to tekst.** Cudzysłów zmienia wszystko. W Excelu zobaczysz słowo `PRAWDA` wyrównane do lewej (bo to tekst), a nie wartość logiczną wyrównaną do środka. Formuła `=JEŻELI(A1;...)` potraktuje tekst `"PRAWDA"` jako prawdę tylko przez przypadek konwersji — i w wielu kontekstach wcale nie. Zapisuj bez cudzysłowu, gdy chcesz wartość logiczną.

**Pułapka 2: porównania w Pythonie.** `True == 1` i `False == 0` zwracają `True`. Jeśli mieszasz w kolumnie boole i liczby (np. „tak/nie" obok wartości), to `suma = sum(kolumna)` policzy Ci `PRAWDA` jako 1. Czasem to jest dokładnie to, czego chcesz; czasem jest to cichy błąd. Filtruj jawnie:

```python
same_liczby = [v for v in kolumna if isinstance(v, (int, float)) and not isinstance(v, bool)]
same_boole = [v for v in kolumna if isinstance(v, bool)]
```

**Pułapka 3: puste „prawda" mylone z pustką.** Nie ma czegoś takiego jak „trójstanowy boolean" w Excelu. Jeśli potrzebujesz „tak / nie / brak odpowiedzi", użyj listy rozwijanej z tekstem (`"Tak"`, `"Nie"`, `""`) albo trzech stanów w osobnej kolumnie. `TRUE`/`FALSE`/`None` da się tak zbudować, ale natkniesz się na wszystkie problemy z sekcji 2.5.

### 2.8. Limity tekstu — 32 767 znaków

Excel ma twarde limity na jedno pole:

| Element | Limit |
|---|---|
| Znaki w komórce | **32 767** |
| Znaki w formule | 8 192 |
| Wiersze w arkuszu | 1 048 576 |
| Kolumny w arkuszu | 16 384 (do `XFD`) |
| Nazwa arkusza | 31 znaków |

Co się dzieje, gdy wpiszesz dłuższy tekst? openpyxl **emituje ostrzeżenie (`UserWarning`)** o przekroczeniu maksymalnej długości łańcucha i obcina wartość do 32 767 znaków. Warto to potwierdzić eksperymentalnie w swojej wersji — to jeden z tych szczegółów implementacyjnych, które bywały zmieniane:

```python
import warnings
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

dlugi = "x" * 40_000

with warnings.catch_warnings(record=True) as zlapane:
    warnings.simplefilter("always")
    ws["A1"] = dlugi
    for ostrzezenie in zlapane:
        print(f"OSTRZEŻENIE: {ostrzezenie.message}")

print(f"wpisano:  {len(dlugi)} znaków")
print(f"w komórce: {len(ws['A1'].value)} znaków")
```

**Nigdy nie polegaj na automatycznym obcinaniu.** Nawet gdyby działało idealnie, dostajesz dane po cichej modyfikacji — najgorszy możliwy scenariusz w raportowaniu. Jeśli tekst naprawdę jest długi (np. pełny opis, notatka z CRM, treść maila), **podziel go świadomie**:

```python
LIMIT = 32_000      # zostawiamy margines - nie dobijamy do 32 767


def podziel_tekst(tekst: str, limit: int = LIMIT) -> list[str]:
    """Dzieli długi tekst na kawałki mieszczące się w limicie komórki.

    Dzielimy po granicach słów, żeby kawałki były czytelne. Jeśli pojedyncze
    słowo jest dłuższe niż limit (np. długi URL), tniemy twardo.
    """
    if len(tekst) <= limit:
        return [tekst]

    kawalki = []
    pozostalo = tekst

    while len(pozostalo) > limit:
        ciecie = pozostalo.rfind(" ", 0, limit)
        if ciecie <= 0:                       # nie ma spacji w limicie - tnij twardo
            ciecie = limit
        kawalki.append(pozostalo[:ciecie].rstrip())
        pozostalo = pozostalo[ciecie:].lstrip()

    if pozostalo:
        kawalki.append(pozostalo)

    return kawalki
```

I wtedy zapisujesz kolejne kawałki do kolejnych komórek (albo do jednej kolumny w dół), informując użytkownika, że tekst jest dzielony. Alternatywa — i często lepsza — to **zapisanie tekstu jako pliku obok skoroszytu** i wstawienie hiperlinku (moduł 16). Komórka to nie miejsce na eseje.

### 2.9. Daty i czas — licznik dni, epoka i strefy

Wracamy do licznika dni. Konkrety:

**Serial to liczba dni od umownej północy.** W pliku, który powstał w większości systemów, obowiązuje epoka **`CALENDAR_WINDOWS_1900`**. Ustalmy jeden punkt odniesienia, żebyś miał do czego wracać: **data `2026-10-07` to serial `46302`**, a `2026-10-07 09:30` to `46302,395833…` (bo 9,5 godziny z 24 to 0,3958333…). Sprawdzisz to sam w ćwiczeniu 🟢.

**Część ułamkowa to pora dnia.** `0,5` to południe, `0,25` to 6:00, `0,75` to 18:00. Czas bez daty jest liczbą z zakresu od 0 do 1. Dlatego `time(9, 30)` zapisany samodzielnie to `0,395833…`, a nie „09:30".

**Excel ma też drugą epokę.** Skoroszyty tworzone na starych wersjach Excela dla Macintosha używają **`CALENDAR_MAC_1904`**, gdzie dni liczy się od `1904-01-01`. Różnica między epokami to **1 462 dni** (4 lata plus fałszywy 29 lutego 1900). Praktyczne znaczenie: jeśli wczytasz taki plik i pomyślisz, że serial `1` to `1900-01-01`, dostaniesz datę starszą o 1 462 dni i wszystkie daty w raporcie przesuną się o cztery lata. Nie rób arytmetyki na serialach ręcznie — openpyxl obsługuje epokę, jeśli tylko mu na to pozwolisz.

```python
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904

wb = Workbook()
print(wb.epoch)                        # 2415018.5  (CALENDAR_WINDOWS_1900)
print(wb.epoch == CALENDAR_WINDOWS_1900)   # True
print(CALENDAR_MAC_1904 - CALENDAR_WINDOWS_1900)   # 1462.0
```

> **Uwaga techniczna:** `wb.epoch` to atrybut skoroszytu; przy `load_workbook()` openpyxl odczytuje go z pliku (z `xl/workbook.xml`). Jeśli w Twojej wersji `wb.epoch` zachowuje się inaczej niż tutaj, sprawdź w dokumentacji swojej wersji — to mało uczęszczany zakątek API.

#### Konwersja w obie strony

openpyxl udostępnia dwie funkcje do ręcznej konwersji:

```python
from datetime import datetime
from openpyxl.utils.datetime import to_excel, from_excel

serial = to_excel(datetime(2026, 10, 7, 9, 30))
print(serial)                            # 46302.395833333336

data = from_excel(46302)
print(data)                              # 2026-10-07 00:00:00
print(type(data).__name__)               # datetime  <- uwaga: NIE date
```

Do czego Ci się to przyda w praktyce? Do trzech rzeczy:

1. **Diagnostyka.** „W pliku widzę `46302`, czy to na pewno 7 października?" — tak, to dokładnie ta funkcja.
2. **Praca z danymi, które przyszły jako liczby** (np. z eksportu, w którym daty nie mają formatu daty).
3. **Debugowanie własnego kodu** — gdy Excel pokazuje `#####` albo `46302`, wiesz, że format komórki nie jest formatem daty.

Jeśli używasz `to_excel()` ręcznie, podawaj `datetime.datetime`. Dla zwykłego `date` najpewniejsza jest konwersja:

```python
from datetime import date, datetime, time

d = date(2026, 10, 7)
serial = to_excel(datetime.combine(d, time.min))     # 46302.0
```

#### Jak openpyxl rozpoznaje datę

To kluczowy mechanizm, który trzeba zrozumieć, bo wyjaśnia najczęstszą pomyłkę z datami w tym module:

**openpyxl pyta o format komórki.** Jeśli format liczby wygląda na format daty, liczba zostanie zamieniona na `datetime`. Jeśli nie — zostanie liczbą. Do sprawdzenia służy między innymi właściwość `cell.is_date`:

```python
if komorka.is_date:
    print(f"{komorka.coordinate}: to data -> {komorka.value}")
else:
    print(f"{komorka.coordinate}: to liczba -> {komorka.value}")
```

Praktyczny wniosek: **data bez formatu daty jest dla openpyxl liczbą.** Jeśli wczytasz plik, w którym ktoś zapisał datę jako zwykłą liczbę (albo wyczyścił format), dostaniesz `46302`, a nie `datetime`. Twoja walidacja musi to przewidywać:

```python
from datetime import datetime
from openpyxl.utils.datetime import from_excel

def na_date(value):
    """Zamienia wczytaną wartość na datetime, niezależnie od tego, jak przyszła."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return from_excel(value)            # traktujemy liczbę jako serial
    return None                             # tekst, pustka, błąd - nie nasza sprawa
```

#### Daty bez strefy — ograniczenie openpyxl

**openpyxl nie obsługuje stref czasowych.** To nie jest luka w bibliotece, którą kiedyś domkną — to konsekwencja modelu Excela. Excel nie przechowuje informacji o strefie; data to liczba dni, pełna kropka. Nie ma gdzie zapisać „+02:00".

Efekt jest taki: jeśli wpiszesz do komórki `datetime` ze strefą, dostaniesz błąd z komunikatem w rodzaju `Excel does not support datetimes with timezones.` (dokładna treść i klasa wyjątku mogą się różnić między wersjami). A więc:

```python
from datetime import datetime, timezone

wb = Workbook()
ws = wb.active

# ŹLE - data świadoma strefy trafi na opór
# ws["A1"] = datetime.now(timezone.utc)

# DOBRZE - konwertuj do lokalnego czasu i usuń strefę ŚWIADOMIE
teraz_utc = datetime.now(timezone.utc)
teraz_lokalnie = teraz_utc.astimezone()          # przesuwa wskazanie zegara
ws["A1"] = teraz_lokalnie.replace(tzinfo=None)   # usuwa metadane strefy

# ALBO - zapisz osobno czas i osobno strefę, jako tekst
ws["A2"] = teraz_utc.replace(tzinfo=None)
ws["B2"] = "UTC"
```

**Reguła na cały kurs:** do komórek trafiają **wyłącznie daty naiwne** (bez `tzinfo`). Konwersję stref robisz w warstwie aplikacji, przed zapisem, i dokumentujesz w metadanych skoroszytu (`wb.properties.description`) lub w arkuszu konfiguracyjnym, w jakiej strefie są dane. Bardzo często wystarczy dopisek „wszystkie czasy w CET" — ale musi gdzieś być.

#### Czas trwania — `timedelta` i przekroczenie 24 godzin

`timedelta` to przypadek szczególny i bardzo praktyczny (nadgodziny, czasy obsługi, czasy realizacji). Excel przechowuje czas trwania jako liczbę dni, ale z odpowiednim formatem, który pozwala przekroczyć 24 godziny: `[hh]:mm:ss`.

```python
from datetime import timedelta

ws["A10"] = timedelta(hours=27, minutes=15)
ws["A10"].number_format = "[hh]:mm:ss"      # bez nawiasów Excel pokaże 03:15!
```

W nawiasach kwadratowych tkwi cała magia: `hh` przewija się po 24 godzinach, `[hh]` — nie. Bez nawiasów dwudniowy czas pracy pokaże się jako jeden dzień i trzy godziny, i nikt się nie zorientuje, że w raporcie jest błąd. Jeśli masz w komórce `timedelta`, **zawsze** ustawiaj format z nawiasami.

### 2.10. Wartości błędów — `#N/A` jest tekstem

Gdy Excel nie potrafi policzyć formuły, zapisuje w komórce **kod błędu**. W pakiecie to osobny typ komórki (`t="e"`) i jedna z siedmiu wartości: `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` (nowsze wersje Excela dodają też m.in. `#SPILL!`, `#CALC!`).

**Przy odczycie openpyxl nie rzuca wyjątku.** Dostajesz `Cell`, którego `value` jest **łańcuchem znaków** (np. `'#DIV/0!'`), a `data_type` ma wartość `'e'`:

```python
k = ws["D2"]
print(k.value)         # '#DIV/0!'
print(k.data_type)     # 'e'
print(type(k.value))   # <class 'str'>
```

To ma trzy poważne konsekwencje:

**Konsekwencja 1: błąd nie zatrzymuje Twojego programu, więc możesz go nie zauważyć.** Skrypt przejdzie bez mrugnięcia okiem, a w raporcie będzie kolumna `#REF!`, której nikt nie przeczyta — bo przecież „policzyło się bez błędu". Walidacja musi być jawna:

```python
KODY_BLEDOW = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}

def czy_blad(komorka) -> bool:
    """Rozpoznaje prawdziwą wartość błędu (t='e') albo tekst udający błąd."""
    if komorka.data_type == "e":
        return True
    return isinstance(komorka.value, str) and komorka.value in KODY_BLEDOW
```

**Test `data_type == 'e'` jest jedynym niezawodnym.** Samo sprawdzanie, czy tekst zaczyna się od `#`, da fałszywe alarmy, bo użytkownik może wpisać `#N/A` jako zwykły tekst (na przykład w kolumnie „komentarz").

**Konsekwencja 2: błąd wygląda jak tekst, więc formuły w dół się rozsypują.** Jeśli z pliku wczytasz `#N/A` i policzysz na tym `float()`, dostaniesz `ValueError`. A jeszcze częściej — policzysz `sum()` po kolumnie i dostaniesz... błąd, bo Python nie doda `str` do `float`. Zawsze filtruj:

```python
liczby = [
    k.value for k in kolumna
    if isinstance(k.value, (int, float)) and not isinstance(k.value, bool)
]
suma = sum(liczby)
print(f"pominięto {len(kolumna) - len(liczby)} komórek bez danych liczbowych")
```

**Konsekwencja 3: nie ma łatwej drogi w drugą stronę.** openpyxl nie oferuje wygodnego API do zapisania „prawdziwego" błędu z Twojego kodu. Wpisanie `ws["A1"] = "#N/A"` da **tekst**, nie błąd. To znaczy, że w raporcie mogą współistnieć dwa różne `#N/A` — prawdziwy błąd Excela i tekst udający błąd — wyglądające identycznie. I że Excel może próbować interpretować Twoją kolumnę z tekstem jako kolumnę z błędami.

Osobna sprawa to **`NaN`**. `float('nan')` nie ma odpowiednika w Excelu — nie da się zapisać „nie-liczby" jako liczby. Jeśli taka wartość dotrze do komórki (na przykład przez `dataframe_to_rows` albo ręcznie), to niezależnie od tego, co zrobi z nią openpyxl, **Ty po odczycie nie dostaniesz tego, co wpisałeś** — możesz zobaczyć kod błędu (`#NUM!`) albo pustkę, zależnie od drogi, którą wartość weszła. Równie źle kończy się konwersja `str(np.nan)` — dostajesz w komórce zwykły tekst `'nan'`, i kolumna z liczbami miesza się z tekstem, więc żadna formuła jej nie policzy.

**Wniosek praktyczny, jednolinijkowy: `NaN` zamieniaj na `None` sam, przed zapisem.**

```python
from math import isnan

def bez_nan(value):
    if isinstance(value, float) and isnan(value):
        return None                      # pusta komórka - jedyna sensowna reprezentacja
    return value
```

### 2.11. `data_type` — jak sprawdzić, co naprawdę jest w komórce

Skoro tyle mówimy o typach, zbierzmy kody w jednym miejscu:

| Kod | Znaczenie | Typ Pythona po odczycie |
|---|---|---|
| `'n'` | liczba (ang. *number*) | `int`, `float` lub `None` dla pustki |
| `'s'` | tekst (ang. *shared string*) | `str` |
| `'str'` | tekst wyniku formuły | `str` |
| `'inlineStr'` | tekst zapisany w komórce, nie w `sharedStrings` | `str` |
| `'b'` | wartość logiczna | `bool` |
| `'d'` | data lub czas | `datetime`, `date`, `time`, `timedelta` |
| `'f'` | formuła (nie wynik!) | `str` zaczynający od `=` |
| `'e'` | błąd | `str` z kodem błędu |

To jest Twoja diagnostyka pierwszego kontaktu. Gdy nie rozumiesz, co się dzieje z danymi, **wypisz `data_type` obok `repr(value)`** i wszystko się wyjaśni:

```python
def opisz(komorka) -> str:
    """Jednolinijkowy opis komórki - wklejaj do debugowania bez żalu."""
    return (f"{komorka.coordinate}: data_type={komorka.data_type!r} "
            f"value={komorka.value!r} typ={type(komorka.value).__name__} "
            f"format={komorka.number_format!r}")
```

### 2.12. Unikanie „magicznych komórek" — nazwy zdefiniowane

Zanim przejdziemy dalej, jedna decyzja projektowa, którą warto podjąć już teraz. Wyobraź sobie arkusz konfiguracyjny, w którym komórka `B2` trzyma stawkę VAT:

```python
vat = ws_konf["B2"].value        # co to za komórka? dlaczego B2?
```

Taki kod ma dwa problemy. Po pierwsze, **czytelnik nie wie, co jest w `B2`** — musi otworzyć plik. Po drugie, **gdy ktoś wstawi wiersz w arkuszu konfiguracji, `B2` przestanie być stawką VAT**, a Twój kod będzie dalej czytał `B2` i dostanie coś innego. To klasyczny „magic number" — wartość wpisana na sztywno w kod, bez nazwy wyjaśniającej jej sens.

**Analogia: etykiety na szufladach.** Zamiast mówić „weź to z trzeciej szuflady od góry", mówisz „weź to z szuflady VAT". Pierwsze jest prawdziwe do pierwszego przestawienia mebli.

openpyxl 3.1 pozwala tworzyć **nazwy zdefiniowane** — czyli nazwy, które wskazują na komórki lub zakresy:

```python
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import quote_sheetname, absolute_coordinate

wb = Workbook()
konf = wb.active
konf.title = "Konfiguracja"
konf["A1"] = "Stawka VAT"
konf["B1"] = 0.23
konf["A2"] = "Waluta"
konf["B2"] = "PLN"

# Tworzymy nazwę STAWKA_VAT wskazującą na 'Konfiguracja'!$B$1
odniesienie = f"{quote_sheetname(konf.title)}!{absolute_coordinate('B1')}"
wb.defined_names.add(DefinedName(name="STAWKA_VAT", attr_text=odniesienie))

# Od tego momentu w Excelu działa =STAWKA_VAT
# a w Pythonie masz jedno miejsce, w którym definiujesz lokalizację danych:
nazwa = wb.defined_names["STAWKA_VAT"]
for arkusz, adres in nazwa.destinations:
    print(f"STAWKA_VAT -> {arkusz}!{adres}")
```

Pełne omówienie nazw, ich zasięgów i pułapek przyjdzie w module 06 (w kontekście formuł). Tutaj zapamiętaj samą zasadę, bo ona jest architektoniczna: **lokalizacje danych konfiguracyjnych definiuj raz, w jednym miejscu i pod nazwą.** To zalążek tego, co w module 26 nazwiemy konfiguracją deklaratywną.

### 2.13. Czytanie danych — trzy sposoby i kiedy który

#### Sposób 1: pojedyncze komórki

```python
wartosc = ws["B2"].value                  # najkrótszy zapis
wartosc = ws.cell(row=2, column=2).value  # gdy adres jest obliczony
```

Dobre do odczytu kilku konkretnych pól: nagłówek, parametr, komórka sumy. Złe do czytania tabel — w pętli po tysiącach komórek tworzysz tysiące obiektów i tracisz wydajność (moduł 19).

#### Sposób 2: `ws.values` — generator wierszy

```python
for wiersz in ws.values:
    print(wiersz)             # ('Produkt', 'Kwota', 'Data') - KROTKA wartości
```

`ws.values` daje **generator** — leniwą sekwencję, która produkuje kolejny wiersz dopiero wtedy, gdy o niego poprosisz. **Analogia: taśma produkcyjna, nie magazyn.** Nie musisz trzymać w pamięci całego arkusza, żeby przejść po nim wiersz po wierszu.

Trzy rzeczy, które trzeba o generatorach wiedzieć:

**Generator to nie lista.** Jeśli zapiszesz go w zmiennej i przejdziesz po nim dwa razy, drugie przejście zwróci nic:

```python
dane = ws.values
pierwsze = list(dane)      # ('Produkt', 'Kwota', ...), (...), ...
drugie = list(dane)        # [] <- pusto! taśma się skończyła

# Rozwiązanie: albo opakuj w list() od razu,
dane = list(ws.values)
# albo odwołaj się do ws.values ponownie - każde odwołanie tworzy świerzy generator
```

**Każdy wiersz ma tyle elementów, ile kolumn w zakresie.** Jeśli w arkuszu dane kończą się w kolumnie C, a `max_column` wynosi 5 (bo ktoś sformatował kolumnę E), każdy wiersz będzie miał 5 elementów — dwa ostatnie to `None`. To da się odfiltrować, ale trzeba o tym wiedzieć.

**`ws.values` nie widzi pustych wierszy na końcu.** Zwraca wiersze do aktualnego rozmiaru modelu. Jeśli „brakujące" dane są wynikiem sformatowania, zobaczysz je jako wiersze z `None`.

#### Sposób 3: `iter_rows()` — pełna kontrola

```python
# Wartości (szybko, dla przetwarzania danych)
for a, b, c in ws.iter_rows(min_row=2, values_only=True):
    ...

# Komórki (gdy potrzebujesz adresu, formatu albo data_type)
for wiersz in ws.iter_rows(min_row=2):
    for komorka in wiersz:
        if komorka.data_type == "e":
            print(f"błąd w {komorka.coordinate}: {komorka.value}")
```

Różnica między `ws.values` a `iter_rows()` jest prosta: `ws.values` to skrót, który daje **tylko wartości** i zawsze cały arkusz. `iter_rows()` pozwala **ograniczyć zakres** (`min_row`, `max_row`, `min_col`, `max_col`), a przez `values_only=False` (domyślnie) daje **obiekty komórek** — czyli dostęp do `data_type`, `number_format` i `coordinate`. Szczegóły i wydajność — moduł 07.

**Najlepszy nawyk, którego warto się nauczyć od razu:** gdy dane przychodzą z pliku, w którym nie panujesz nad typami, przepuść każdą wartość przez **jedną, centralną funkcję normalizującą**. Nie rozsypuj `if value is None` po całym kodzie.

```python
from datetime import date, datetime
from decimal import Decimal


def bezpieczna_wartosc(value):
    """Sprowadza wartość z komórki do postaci bezpiecznej dla logiki biznesowej.

    Reguły:
      * None            -> ""      (bo pusta komórka w raporcie ma być pustym tekstem)
      * str             -> bez spacji na brzegach
      * Decimal         -> float    (Excel i tak trzyma liczby jako float)
      * datetime        -> ISO 8601 (tekst - nadaje się do porównań i JSON-a)
      * date            -> ISO 8601
      * bool            -> zostaje bool (nie chcemy przypadkowej zamiany na 1/0)
      * wszystko inne   -> str(value)  (ostatnia linia obrony)
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    return value
```

To jeszcze nie jest wzorzec projektowy — to zdrowy nawyk. Ale zobacz: funkcja ma **jedną odpowiedzialność** (określić, jak wartość z pliku trafia do świata biznesowego) i **jedno miejsce**, w którym zmienisz reguły. W module 24 nazwiemy to **Adapterem**.

## 3. Przykłady krok po kroku

Wszystkie przykłady zapisują do `output/` i są samodzielnymi, uruchamialnymi skryptami.

### Przykład 1 — tabela typów: co weszło, co wyszło

```python
"""Wpisujemy po jednej wartości każdego obsługiwanego typu i sprawdzamy,
co openpyxl zapisał i co odczytał z powrotem.

Uruchom:  python examples/05_typy.py
"""

from datetime import date, datetime, time, timedelta
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# (etykieta, wartość do wpisania)
PRZYKLADY = [
    ("tekst",                "Widzet"),
    ("int",                  42),
    ("int duży (18 cyfr)",   123456789012345678),
    ("float",                3.14),
    ("float 0.1+0.2",        0.1 + 0.2),
    ("bool True",            True),
    ("bool False",           False),
    ("None",                 None),
    ("pusty string",         ""),
    ("datetime",             datetime(2026, 10, 7, 9, 30, 0)),
    ("date",                 date(2026, 10, 7)),
    ("time",                 time(9, 30, 0)),
    ("timedelta 27h15m",     timedelta(hours=27, minutes=15)),
    ("Decimal",              Decimal("19.99")),
    ("tekst liczbowy '007'", "007"),
    ("tekst liczbowy '1'",   "1"),
]


def zbuduj() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Typy"

    ws["A1"] = "Co wpisano"
    ws["C1"] = "Co wpisano (repr)"

    for indeks, (etykieta, wartosc) in enumerate(PRZYKLADY, start=2):
        ws.cell(row=indeks, column=1, value=etykieta)     # adres liczbowy: w pętli
        ws.cell(row=indeks, column=2, value=wartosc)      # <- tu dzieje się magia typów
        ws.cell(row=indeks, column=3, value=repr(wartosc))

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 16
    return wb


def main() -> None:
    wb = zbuduj()

    # Co jest w MODELU - przed zapisem
    print("=" * 96)
    print("W MODELU (przed save) - typ obiektu trzymany w komórce")
    print("=" * 96)
    for indeks, (etykieta, wartosc) in enumerate(PRZYKLADY, start=2):
        k = wb["Typy"].cell(row=indeks, column=2)
        print(f"  {etykieta:<22} {type(k.value).__name__:<10} {k.value!r}")

    cel = OUTPUT / "05_typy.xlsx"
    wb.save(cel)
    print(f"\nZapisano: {cel} ({cel.stat().st_size} B)")

    # Co jest w PLIKU - po ponownym wczytaniu
    print()
    print("=" * 96)
    print("Z PLIKU (po odczycie) - co openpyxl odtworzył")
    print("=" * 96)
    print(f"  {'etykieta':<22} {'typ':<10} {'data_type':<10} {'is_date':<8} "
          f"{'format':<18} {'repr':<26}")
    print("  " + "-" * 92)

    kontrola = load_workbook(cel)
    try:
        ws = kontrola["Typy"]
        for indeks, (etykieta, _) in enumerate(PRZYKLADY, start=2):
            k = ws.cell(row=indeks, column=2)
            print(f"  {etykieta:<22} {type(k.value).__name__:<10} {k.data_type:<10} "
                  f"{str(k.is_date):<8} {k.number_format:<18} {k.value!r:<26}")
    finally:
        kontrola.close()

    # Wnioski - programowo, bez zgadywania
    print()
    print("=" * 96)
    print("WNIOSKI (wykryte automatycznie)")
    print("=" * 96)

    kontrola = load_workbook(cel)
    try:
        ws = kontrola["Typy"]
        for indeks, (etykieta, wpisane) in enumerate(PRZYKLADY, start=2):
            odczytane = ws.cell(row=indeks, column=2).value
            if type(wpisane) is not type(odczytane):
                print(f"  ZMIANA TYPU:  {etykieta:<22} "
                      f"{type(wpisane).__name__} -> {type(odczytane).__name__}")
        print("\n  (typ NoneType dla pustej komórki jest oczekiwany)")
    finally:
        kontrola.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `ws.cell(row=i, column=2, value=...)` wywołuje `cell.value = ...`, a to z kolei `_bind_value()`, które bada typ Pythona i ustawia `data_type` oraz — tam, gdzie trzeba — konwertuje wartość na postać zrozumiałą dla Excela. `Decimal("19.99")` zostaje zapamiętany jako obiekt `Decimal`, ale przy zapisie stanie się liczbą w XML. Data zostaje zamieniona na **serial** (`46302.395833…`) już w pamięci — w komórce nie ma już obiektu `datetime` w postaci „oryginalnej", tylko liczba z etykietą `'d'`.

**Co trafi do pliku.** W `xl/worksheets/sheet1.xml` zobaczysz liczby, indeksy do tekstów i kilka komórek z `t="b"`. Pustka (`None`) **nie pojawi się w pliku wcale**. Data pojawi się jako liczba z formatem daty w `styles.xml`.

**Czego się z tego nauczysz, patrząc na wydruk:** wiersz `date` pokaże zmianę typu na `datetime` — to najważniejsza obserwacja tego przykładu. Wiersz `Decimal` pokaże zmianę na `float`. A wiersz `int duży` — ciekawostkę: openpyxl odtworzy pełną liczbę, ale **Excel pokaże `1,23456789012346E+17`**. Otwórz plik i zobacz sam.

### Przykład 2 — typy, które się nie mieszczą

```python
"""Co openpyxl odrzuca i jak to naprawić.

Uruchom:  python examples/05_odrzucone.py
"""

import json
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# Typy, które openpyxl odrzuca
ODRZUCANE = [
    ("list",   ["Widzet", "Wkręt"]),
    ("dict",   {"klient": "ACME", "kwota": 100}),
    ("set",    {"a", "b"}),
    ("tuple",  (1, 2, 3)),
    ("object", object()),
]

# Sprawdźmy też numpy, jeśli jest dostępne
try:
    import numpy as np
    ODRZUCANE.append(("numpy.int64", np.int64(42)))
    ODRZUCANE.append(("numpy.float64", np.float64(3.14)))
except ImportError:
    print("numpy nie jest zainstalowane - pomijam przypadki numpy\n")


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Typy"

    print("=" * 88)
    print("PRÓBA WPISANIA TYPÓW SPOZA LISTY")
    print("=" * 88)

    dobre = []
    for indeks, (etykieta, wartosc) in enumerate(ODRZUCANE, start=1):
        try:
            ws.cell(row=indeks, column=1, value=wartosc)
            wynik = f"OK (zapisano jako {type(ws.cell(row=indeks, column=1).value).__name__})"
            dobre.append(etykieta)
        except (ValueError, TypeError) as blad:
            komunikat = str(blad)
            if len(komunikat) > 60:
                komunikat = komunikat[:57] + "..."
            wynik = f"{type(blad).__name__}: {komunikat}"
        print(f"  {etykieta:<16} -> {wynik}")

    print()
    print("=" * 88)
    print("CZTERY SPOSOBY NAPRAWY")
    print("=" * 88)

    ws2 = wb.create_sheet(title="Naprawione")

    # 1. int() - gdy wartość NAPRAWDĘ jest liczbą, tylko z obcej rodziny
    try:
        import numpy as np
        ws2["A1"] = int(np.int64(42))               # <- doprowadzamy do wbudowanego int
        ws2["B1"] = "int(): konwersja typu"
    except ImportError:
        pass

    # 2. json.dumps() - gdy struktura ma zostać w jednej komórce (rzadko dobre!)
    struktura = {"klient": "ACME", "kwota": 100}
    ws2["A2"] = json.dumps(struktura, ensure_ascii=False)
    ws2["B2"] = "json.dumps(): struktura jako tekst"

    # 3. str.join() - gdy chodzi o czytelną reprezentację
    ws2["A3"] = ", ".join(["Widzet", "Wkręt"])
    ws2["B3"] = "join(): lista jako tekst"

    # 4. NAJLEPSZE: rozbicie na kolumny - struktura staje się tabelą
    rekordy = [
        {"klient": "ACME", "kwota": 100},
        {"klient": "Beta", "kwota": 250},
    ]
    naglowki = list(rekordy[0].keys())
    ws2.append(naglowki)                            # <- wiersz nagłówków
    for rekord in rekordy:
        ws2.append([rekord[k] for k in naglowki])   # <- wartości w kolumnach

    ws2["B5"] = "rozłożenie na kolumny: najczęstsze właściwe rozwiązanie"

    cel = OUTPUT / "05_odrzucone.xlsx"
    wb.save(cel)
    print(f"  Zapisano: {cel}")

    print()
    print("=" * 88)
    print("STAN ARKUSZA 'Naprawione' PO ODCZYCIE")
    print("=" * 88)
    kontrola = load_workbook(cel)
    try:
        for wiersz in kontrola["Naprawione"].iter_rows(values_only=True):
            if any(v is not None for v in wiersz):
                print("  " + " | ".join("" if v is None else repr(v) for v in wiersz))
    finally:
        kontrola.close()

    print()
    print(f"  Odrzucone typy: {[t for t, _ in ODRZUCANE if t not in dobre] or 'brak'}")
    print(f"  Przyjęte:       {dobre}")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Przy każdym nieudanym przypisaniu openpyxl nie tworzy komórki z wartością — wyjątek leci, zanim wartość trafi do modelu. Jeśli wywołasz `ws.cell(row, column)` **bez** wartości, komórka powstanie jako pusta. W kodzie powyżej po nieudanej próbie komórka i tak istnieje (bo `ws.cell` został już wywołany) — i to jest dobry moment, by zobaczyć w praktyce uwagę z sekcji 2.1: **dotknięcie komórki tworzy ją w modelu.**

**Co trafi do pliku.** Do arkusza `Typy` trafią wyłącznie komórki, którym udało się przypisać wartość. `np.float64` przejdzie (bo dziedziczy po `float`), `np.int64` nie przejdzie. W arkuszu `Naprawione` zobaczysz cztery rozwiązania obok siebie — w tym prawdziwą miniaturową tabelę z nagłówkami.

### Przykład 3 — `None`, `""` i `0` w jednym arkuszu

```python
"""Trzy rodzaje "pustki" i ich wpływ na formuły Excela.

openpyxl formuł nie liczy - zrobi to Excel przy otwarciu. Ten skrypt
przygotowuje poligon i wypisuje, co dokładnie znajdziesz w komórkach.

Uruchom:  python examples/05_pustka.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pustka"

    ws["A1"] = "Opis"
    ws["B1"] = "Wartość"
    ws["C1"] = "ISBLANK"

    ws["A2"] = "brak wartości (None)"
    ws["B2"] = None                                  # puste miejsce - nie ma krzesła
    ws["C2"] = "=ISBLANK(B2)"

    ws["A3"] = "pusty tekst ('')"
    ws["B3"] = ""                                    # krzesło stoi, nikt nie siedzi
    ws["C3"] = "=ISBLANK(B3)"

    ws["A4"] = "zero (0)"
    ws["B4"] = 0                                     # ktoś siedzi i się nie odzywa
    ws["C4"] = "=ISBLANK(B4)"

    # Formuły podsumowujące - policzy je Excel po otwarciu pliku
    ws["A6"] = "COUNTA(B2:B4)"
    ws["B6"] = "=COUNTA(B2:B4)"
    ws["A7"] = "COUNTBLANK(B2:B4)"
    ws["B7"] = "=COUNTBLANK(B2:B4)"
    ws["A8"] = "AVERAGE(B2:B4)"
    ws["B8"] = "=AVERAGE(B2:B4)"
    ws["A9"] = "COUNT(B2:B4)"
    ws["B9"] = "=COUNT(B2:B4)"

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12

    cel = OUTPUT / "05_pustka.xlsx"
    wb.save(cel)
    print(f"Zapisano: {cel}")
    print("Otwórz plik w Excelu i sprawdź, czy widzisz:")
    print("  ISBLANK:  PRAWDA / FAŁSZ / FAŁSZ")
    print("  COUNTA:   2     (drugie miejsce 'zajęte' to pusty tekst!)")
    print("  COUNTBLANK: 2   (brak + pusty tekst)")
    print("  AVERAGE:  0     (jedyna liczba to zero)")
    print("  COUNT:    1     (tylko zero jest liczbą)")

    print()
    print("=" * 76)
    print("CO WIDZI OPENPYXL (bez liczenia formuł - tylko wartości)")
    print("=" * 76)

    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Pustka"]
        for adres in ("B2", "B3", "B4"):
            k = ws2[adres]
            print(f"  {adres}: value={k.value!r:<6} data_type={k.data_type!r} "
                  f"typ={type(k.value).__name__}")

        print()
        print("  Formuły w C2:C4 są w komórkach jako TEKST:")
        for adres in ("C2", "C3", "C4"):
            k = ws2[adres]
            print(f"  {adres}: value={k.value!r} data_type={k.data_type!r}")
    finally:
        kontrola.close()

    print()
    print("=" * 76)
    print("PUŁAPKA W PYTHONIE")
    print("=" * 76)
    print("""
  for komorka in ws["B2":"B4"]:   # (uproszczenie - zakres daje krotkę krotek)
      ...

  if not komorka.value:
      print("pusto")     # <- TO ZŁAPIE RÓWNIEŻ ZERO!

  Aby uniknąć pułapki:
      if komorka.value is None:            # tylko brak wartości
      if komorka.value in (None, ""):      # brak wartości albo pusty tekst
      if isinstance(komorka.value, (int, float)) \\
              and not isinstance(komorka.value, bool):   # tylko liczby
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Trzy różne stany w trzech komórkach. Zauważ, że `ws["B2"] = None` **nie usuwa** komórki z modelu — tylko jej wartość. Gdyby ta komórka miała wcześniej ustawiony styl, jej wpis zostałby w pliku (bo komórki ze stylem są zapisywane), a wynik `ISBLANK` byłby i tak `PRAWDA`. Styl i wartość to dwie niezależne warstwy.

**Co trafi do pliku.** `B2` **nie trafi** (brak wartości, brak stylu). `B3` trafi jako komórka tekstowa ze wskaźnikiem do pustego łańcucha. `B4` trafi jako liczba `0`. To dokładnie ta różnica, którą potem „widzi" Excel w `COUNTA`, `COUNTBLANK` i `ISBLANK`.

**Dowód lepszy niż wszystko:** otwórz plik w Excelu i kliknij `COUNTA`. Zobaczysz `2` w polu, w którym kusiłoby Cię `1`. A potem sprawdź `COUNTBLANK` — też `2`. Ta pozorna sprzeczność to nie błąd Excela, to dwie różne definicje pustki w jednym programie.

### Przykład 4 — daty od podszewki: serial, format, epoka

```python
"""Data jako liczba. Serial, format, epoka i powrót do datetime.

Uruchom:  python examples/05_daty.py
"""

from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import (
    CALENDAR_MAC_1904,
    CALENDAR_WINDOWS_1900,
    from_excel,
    to_excel,
)

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def main() -> None:
    print("=" * 84)
    print("1. EPOKA")
    print("=" * 84)

    wb = Workbook()
    print(f"  wb.epoch                      = {wb.epoch}")
    print(f"  CALENDAR_WINDOWS_1900         = {CALENDAR_WINDOWS_1900}")
    print(f"  CALENDAR_MAC_1904             = {CALENDAR_MAC_1904}")
    print(f"  różnica epok (dni)            = {CALENDAR_MAC_1904 - CALENDAR_WINDOWS_1900}")
    print("  ^ 1462 dni = 4 lata + fałszywy 29 lutego 1900 r.")
    print("    Sprawdź w dokumentacji swojej wersji, jeśli wb.epoch zachowa się inaczej.")

    print()
    print("=" * 84)
    print("2. SERIAL - RĘCZNIE")
    print("=" * 84)

    for dt in (
        datetime(1900, 1, 1),
        datetime(2026, 1, 1),
        datetime(2026, 10, 7),
        datetime(2026, 10, 7, 9, 30),
        datetime(2026, 10, 7, 18, 0),
    ):
        serial = to_excel(dt)
        z_powrotem = from_excel(serial)
        print(f"  {dt:%Y-%m-%d %H:%M}  ->  serial {serial:<20.10f}  ->  {z_powrotem}")

    print()
    print("  Pory dnia jako ułamki:")
    for t in (time(0, 0), time(6, 0), time(12, 0), time(18, 0), time(23, 59, 59)):
        print(f"    {t}  ->  {to_excel(t):.8f}")

    print()
    print("=" * 84)
    print("3. ZAPIS I ODCZYT - CO Z CZYM WRACA")
    print("=" * 84)

    wb = Workbook()
    ws = wb.active
    ws.title = "Daty"

    wartosci = [
        ("datetime", datetime(2026, 10, 7, 9, 30, 0)),
        ("date", date(2026, 10, 7)),
        ("time", time(9, 30, 0)),
        ("timedelta", timedelta(hours=27, minutes=15)),
    ]

    for indeks, (etykieta, wartosc) in enumerate(wartosci, start=2):
        ws.cell(row=indeks, column=1, value=etykieta)
        komorka = ws.cell(row=indeks, column=2, value=wartosc)

        # ŚWIADOMIE ustawiamy format zamiast polegać na domyślnym
        if etykieta == "datetime":
            komorka.number_format = "yyyy-mm-dd hh:mm:ss"
        elif etykieta == "date":
            komorka.number_format = "yyyy-mm-dd"
        elif etykieta == "time":
            komorka.number_format = "hh:mm:ss"
        elif etykieta == "timedelta":
            komorka.number_format = "[hh]:mm:ss"      # NAWIASASY! inaczej 03:15 zamiast 27:15

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22

    cel = OUTPUT / "05_daty.xlsx"
    wb.save(cel)

    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Daty"]
        print(f"  {'etykieta':<12} {'typ wejściowy':<14} {'typ wyjściowy':<14} "
              f"{'wartość':<24} {'format'}")
        print("  " + "-" * 78)
        for indeks, (etykieta, wartosc) in enumerate(wartosci, start=2):
            k = ws2.cell(row=indeks, column=2)
            print(f"  {etykieta:<12} {type(wartosc).__name__:<14} "
                  f"{type(k.value).__name__:<14} {str(k.value):<24} {k.number_format}")
    finally:
        kontrola.close()

    print()
    print("  ^ ZWRÓĆ UWAGĘ: 'date' wraca jako 'datetime'! Excel nie odróżnia")
    print("    daty od północy. Jeśli porównujesz z datetime.date(), to porównanie")
    print("    NIE zadziała. Normalizuj typ przed porównaniem.")

    print()
    print("=" * 84)
    print("4. DATA BEZ FORMATU = LICZBA DLA OPENPYXL")
    print("=" * 84)

    wb = Workbook()
    ws = wb.active
    ws.title = "BezFormatu"

    ws["A1"] = "data z formatem daty"
    ws["B1"] = datetime(2026, 10, 7)
    ws["B1"].number_format = "yyyy-mm-dd"          # <- to decyduje o rozpoznaniu

    ws["A2"] = "TA SAMA wartość, format 'General'"
    ws["B2"] = 46302                               # <- to jest ten sam dzień, jako liczba
    ws["B2"].number_format = "General"             # <- openpyxl tego NIE rozpozna jako daty

    cel2 = OUTPUT / "05_daty_bez_formatu.xlsx"
    wb.save(cel2)

    kontrola = load_workbook(cel2)
    try:
        ws2 = kontrola["BezFormatu"]
        for adres in ("B1", "B2"):
            k = ws2[adres]
            print(f"  {adres}: value={k.value!r:<24} typ={type(k.value).__name__:<10} "
                  f"is_date={k.is_date}")
        print("  ^ Ta sama data. Rozpoznana tylko tam, gdzie format jest formatem daty.")
    finally:
        kontrola.close()

    print()
    print("=" * 84)
    print("5. STREFA CZASOWA - CZEGO NIE WOLNO")
    print("=" * 84)

    wb = Workbook()
    ws = wb.active
    teraz_utc = datetime.now(timezone.utc)
    print(f"  datetime ze strefą: {teraz_utc}")
    try:
        ws["A1"] = teraz_utc
        print("  -> zapisano (!)  sprawdź zachowanie w SWOJEJ wersji openpyxl")
    except (ValueError, TypeError) as blad:
        print(f"  -> {type(blad).__name__}: {blad}")

    # Poprawna droga:
    ws["A2"] = teraz_utc.astimezone().replace(tzinfo=None)
    print(f"  poprawnie (naiwna): {ws['A2'].value}")
    print("  ^ Konwertuj strefę ŚWIADOMIE, zanim wartość dotknie komórki.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Przy `ws["B1"] = datetime(...)` openpyxl **natychmiast** zamienia obiekt na serial (`46302.0`) i ustawia `data_type = 'd'`. W komórce nie ma już `datetime`. Przy odczycie mechanizm działa w drugą stronę — ale **tylko wtedy, gdy format komórki jest formatem daty**. Wiersz 4 tego przykładu pokazuje to czarno na białym: ta sama liczba `46302` z formatem `General` pozostaje liczbą i openpyxl zwraca `int`.

**Co trafi do pliku.** Wartości liczbowe (seriale), a w `xl/styles.xml` — wpisy formatów, do których komórki się odwołują. Data nie jest w pliku „napisana" — jest **zaszyfrowana jako liczba**, a format jest kluczem do jej odczytania.

### Przykład 5 — tekst, który udaje liczbę, i liczba, która udaje identyfikator

```python
"""Diagnoza i naprawa: zera wiodące, długie identyfikatory, "liczby" jako tekst.

Uruchom:  python examples/05_tekst_vs_liczba.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def czy_liczba_w_tekscie(value) -> bool:
    """Czy wartość jest tekstem, który wygląda na liczbę zapisaną po polsku?"""
    if not isinstance(value, str):
        return False
    proba = value.strip().replace("\xa0", "").replace(" ", "")
    # Obsługa zapisu polskiego (przecinek) i angielskiego (kropka)
    proba = proba.replace(",", ".")
    if proba.count(".") > 1:                 # "1.234.567" - separatory tysięcy
        proba = proba.replace(".", "", proba.count(".") - 1)
    return proba.replace(".", "", 1).replace("-", "", 1).isdigit()


def main() -> None:
    # --- Budujemy "brudny" plik, jaki dostalibyśmy od kogoś z firmy ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"

    ws["A1"] = "Kod pocztowy"
    ws["B1"] = "NIP"
    ws["C1"] = "Kwota"
    ws["D1"] = "Ilość"

    ws["A2"] = "00-950"                       # tekst - poprawnie (ale tylko dlatego, że ma myślnik)
    ws["B2"] = "1234567890"                   # tekst - poprawnie, ale to liczba udająca tekst
    ws["C2"] = "1 234,56"                     # LICZBA ZAPISANA JAKO TEKST - problem!
    ws["D2"] = 12                             # liczba - poprawnie

    ws["A3"] = "00950"                        # tekst z zerem wiodącym
    ws["B3"] = 1234567890                     # liczba - straciła potencjalne zero wiodące
    ws["C3"] = "89,99"                        # tekst zamiast liczby
    ws["D3"] = 7

    ws["A4"] = 1234.56                        # liczba na pozycji tekstu
    ws["B4"] = "PL1234567890"                 # identyfikator - poprawnie jako tekst
    ws["C4"] = 0                              # ZERO - pełnoprawna liczba, nie pustka
    ws["D4"] = None                           # brak danych

    cel = OUTPUT / "05_brudny.xlsx"
    wb.save(cel)
    print(f"Utworzono plik testowy: {cel}")

    print()
    print("=" * 100)
    print("DIAGNOZA: co naprawdę jest w komórkach")
    print("=" * 100)

    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Dane"]
        naglowki = [c.value for c in ws2[1]]
        print(f"  {'':<8}" + "".join(f"{n:<18}" for n in naglowki))
        print("  " + "-" * 92)

        for wiersz in ws2.iter_rows(min_row=2):
            linia = f"  {wiersz[0].row:<8}"
            for komorka in wiersz:
                opis = f"{komorka.data_type}:{komorka.value!r}"
                linia += f"{opis[:17]:<18}"
            print(linia)

        print()
        print("  Legenda data_type:  s = tekst   n = liczba   d = data   b = bool   e = błąd")
    finally:
        kontrola.close()

    print()
    print("=" * 100)
    print("DETEKCJA PROBLEMÓW")
    print("=" * 100)

    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Dane"]
        naglowki = [c.value for c in ws2[1]]

        for wiersz in ws2.iter_rows(min_row=2):
            for komorka in wiersz:
                nazwa = naglowki[komorka.column - 1]
                v = komorka.value

                if komorka.data_type == "s" and czy_liczba_w_tekscie(v):
                    print(f"  {komorka.coordinate} ({nazwa}): LICZBA JAKO TEKST -> {v!r}")

                if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) >= 10**11:
                    print(f"  {komorka.coordinate} ({nazwa}): RYZYKO UTRATY CYFR    -> {v!r}")
    finally:
        kontrola.close()

    print()
    print("=" * 100)
    print("NAPRAWA: normalizacja z zachowaniem tego, co ważne")
    print("=" * 100)

    # Kolumny, w których liczby MOGĄ być konwertowane na typ liczbowy
    LICZBOWE = {"Kwota", "Ilość"}
    # Kolumny, które ZAWSZE zostają tekstem (identyfikatory, kody)
    TEKSTOWE = {"Kod pocztowy", "NIP"}


def normalizuj(value, nazwa_kolumny: str):
    """Zwraca wartość w postaci właściwej dla danej kolumny."""
    if value is None:
        return None

    if nazwa_kolumny in TEKSTOWE:
        return str(value).strip()               # RÓB TEKST ZAWSZE - nie trać zer

    if nazwa_kolumny in LICZBOWE:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return value                        # już liczba
        if isinstance(value, str):
            tekst = value.strip().replace("\xa0", "").replace(" ", "")
            tekst = tekst.replace(",", ".")
            try:
                liczba = float(tekst)
            except ValueError:
                return value                    # nie da się - zostawiamy jak jest
            # Całkowite zapisujemy jako int, żeby "12.0" nie wyglądało dziwnie
            return int(liczba) if liczba.is_integer() else liczba

    return value


def main2() -> None:
    zrodlo = OUTPUT / "05_brudny.xlsx"
    cel = OUTPUT / "05_brudny_naprawiony.xlsx"

    wb = load_workbook(zrodlo)
    try:
        ws = wb["Dane"]
        naglowki = [c.value for c in ws[1]]
        zmiany = []

        for wiersz in ws.iter_rows(min_row=2):
            for komorka in wiersz:
                nazwa = naglowki[komorka.column - 1]
                przed = komorka.value
                po = normalizuj(przed, nazwa)
                if przed != po:
                    zmiany.append((komorka.coordinate, nazwa, przed, po))
                    komorka.value = po
                    # Dla kolumn tekstowych wymuszamy format tekstowy w Excelu
                    if nazwa in TEKSTOWE:
                        komorka.number_format = "@"      # '@' = format tekstowy

        wb.properties.description = (
            f"Plik znormalizowany automatycznie. Zmieniono {len(zmiany)} komórek."
        )
        wb.save(cel)
    finally:
        wb.close()

    print(f"  Zapisano: {cel}")
    print()
    for adres, nazwa, przed, po in zmiany:
        print(f"  {adres:<6} {nazwa:<16} {przed!r:<18} ->  {po!r}")

    print()
    print("=" * 100)
    print("KONTROLA PO NAPRAWIE")
    print("=" * 100)
    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Dane"]
        for wiersz in ws2.iter_rows(min_row=2):
            opis = "  "
            for komorka in wiersz:
                opis += f"{komorka.value!r:<20}"
            print(opis)
        print()
        print("  ^ Kod '00950' i NIP zachowały pełną postać (tekst).")
        print("    Kwoty są już LICZBAMI - Excel policzy je w SUMIE.")
    finally:
        kontrola.close()


if __name__ == "__main__":
    main()
    print()
    main2()
```

**Co się dzieje w pamięci.** `normalizuj()` to **funkcja czysta** — nie dotyka arkusza, tylko mapuje wartość na wartość. Dopiero pętla `main2()` decyduje, kiedy wynik zapisać do komórki. Ten podział jest ważny: **logika przekształcenia jest oddzielona od operacji na arkuszu**. Właśnie dzięki temu tę funkcję da się przetestować bez tworzenia jakiegokolwiek pliku (moduł 20) i wymienić bez dotykania kodu zapisu (moduł 23).

**Co trafi do pliku.** Nowy plik `05_brudny_naprawiony.xlsx`, w którym `"1 234,56"` stało się liczbą `1234.56`, `"89,99"` — liczbą `89.99`, a kolumny identyfikatorów mają format `@` (tekstowy), więc Excel nie będzie „poprawiał" zer wiodących przy edycji. Format `@` to ważny szczegół: bez niego użytkownik wpisze `00950`, a Excel natychmiast zamieni to na `950`.

### Przykład 6 — czytanie tabeli do słowników z pełnym raportem typów

```python
"""Wzorzec czytania: nagłówki + wiersze -> lista słowników + raport jakości.

Ten skrypt to zalążek tego, co w module 24 nazwiemy Adapterem.

Uruchom:  python examples/05_czytanie.py
"""

from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def czytaj_tabele(ws, wiersz_naglowkow: int = 1) -> list[dict]:
    """Czyta arkusz tabelaryczny do listy słowników.

    Zakłada układ: jeden wiersz nagłówków, dane poniżej, bez pustych wierszy
    w środku. Zwraca listę rekordów z kluczami z nagłówków.
    """
    naglowki = [
        k.value if k.value is not None else f"kolumna_{k.column}"
        for k in ws[wiersz_naglowkow]
    ]

    rekordy = []
    for wiersz in ws.iter_rows(min_row=wiersz_naglowkow + 1, values_only=True):
        if all(v is None for v in wiersz):       # całkowicie pusty wiersz - pomijamy
            continue
        rekordy.append(dict(zip(naglowki, wiersz)))
    return rekordy


def raport_typow(ws, wiersz_naglowkow: int = 1) -> dict[str, Counter]:
    """Zlicza typy wartości w każdej kolumnie - szybki audyt jakości danych."""
    naglowki = [k.value for k in ws[wiersz_naglowkow]]
    raport: dict[str, Counter] = {n: Counter() for n in naglowki}

    for wiersz in ws.iter_rows(min_row=wiersz_naglowkow + 1):
        for komorka in wiersz:
            nazwa = naglowki[komorka.column - 1]
            raport[nazwa][komorka.data_type] += 1

    return raport


def main() -> None:
    # Budujemy realistyczny arkusz z różnorodnością typów
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaz"

    ws.append(["Data", "Klient", "Kwota", "Sztuk", "Rabat", "Uwagi"])
    ws.append(["2026-10-01", "ACME", 1234.5, 12, True, None])
    ws.append(["2026-10-02", "Beta sp. z o.o.", "89,99", 3, False, "zwrot"])
    ws.append(["2026-10-03", "Gamma", 0, 0, True, None])
    ws.append(["2026-10-04", "Delta", "#N/A", 5, False, None])     # błąd!
    ws.append(["2026-10-05", "Epsilon", 450.0, None, True, "pilne"])
    ws.append([None, None, None, None, None, None])                # pusty wiersz

    cel = OUTPUT / "05_sprzedaz.xlsx"
    wb.save(cel)
    print(f"Utworzono: {cel}")

    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Sprzedaz"]

        print()
        print("=" * 96)
        print("RAPORT TYPÓW (data_type) - pierwszy krok każdego audytu")
        print("=" * 96)
        for nazwa, licznik in raport_typow(ws2).items():
            rozklad = ", ".join(f"{typ}={ile}" for typ, ile in sorted(licznik.items()))
            print(f"  {nazwa:<12} {rozklad}")

        print()
        print("  ^ Kolumna 'Kwota' ma 's' i 'n' - czyli MIESZANKĘ tekstu i liczb.")
        print("    To natychmiast widoczny sygnał, że SUMY w Excelu nie zadziałają.")
        print("    Kolumna 'Data' ma 's' - to tekstdaty, nie daty.")

        print()
        print("=" * 96)
        print("RZECZY DO NAPRAWY")
        print("=" * 96)
        for wiersz in ws2.iter_rows(min_row=2):
            for komorka in wiersz:
                if komorka.data_type == "e":
                    print(f"  {komorka.coordinate}: BŁĄD w komórce -> {komorka.value!r}")
                elif komorka.data_type == "n" and komorka.is_date:
                    print(f"  {komorka.coordinate}: data w formacie daty -> {komorka.value}")

        print()
        print("=" * 96)
        print("ODCZYT DO SŁOWNIKÓW")
        print("=" * 96)
        rekordy = czytaj_tabele(ws2)
        for rekord in rekordy:
            print(f"  {rekord}")

        print()
        print(f"  wierszy danych (bez pustego): {len(rekordy)}")
        print("  ^ Pusty wiersz został pominięty programowo, a nie przypadkiem.")
    finally:
        kontrola.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `iter_rows(values_only=True)` tworzy po stronie Pythona krotki wartości — nie obiekty komórek. `raport_typow()` używa wersji **bez** `values_only`, bo potrzebuje `data_type` i `column`. To dobry przykład, że wybór wariantu iteracji zależy od tego, czy potrzebujesz samej wartości, czy też metadanych komórki.

**Co trafi do pliku.** Nic nowego — to skrypt czytający. Ale plik `05_sprzedaz.xlsx` jest celowo „brudny" i posłuży Ci za materiał do ćwiczenia 🔴.

## 4. Anatomia API

| Metoda / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `ws["A1"]` | Zwraca obiekt `Cell` (tworzy go, jeśli nie istnieje) | adres A1, zakres `"A1:C3"`, kolumna `"B"`, wiersz `"3"` | To **nie** jest wartość. Wartość to `ws["A1"].value` |
| `ws["A1"] = x` | Ustawia wartość komórki | dowolna wartość z listy typów | Skrót od `ws["A1"].value = x` |
| `ws.cell(row, column, value=None)` | Zwraca komórkę po adresie liczbowym | obie współrzędne **od 1** | Wygodne w pętlach. Przyjmuje `value=` |
| `cell.value` | Treść komórki | — | `None` dla pustej; `str` dla kodu błędu |
| `cell.data_type` | Kod typu: `'n'`, `'s'`, `'str'`, `'inlineStr'`, `'b'`, `'d'`, `'f'`, `'e'` | — | Najpewniejszy sposób rozpoznania zawartości |
| `cell.is_date` | `True`, jeśli format liczby jest formatem daty | — | Nie mówi „czy wartość jest datą", tylko „jak sformatowana" |
| `cell.number_format` | Kod formatu wyświetlania (`'General'`, `'0.00'`, `'yyyy-mm-dd'`, `'@'`) | — | Nie zmienia wartości, zmienia wyświetlanie |
| `cell.coordinate` | Adres w notacji A1 | — | `'B7'` |
| `cell.row`, `cell.column` | Współrzędne liczbowe (od 1) | — | `cell.row == 1` dla wiersza pierwszego |
| `cell.column_letter` | Litera kolumny | — | W nowszym kodzie preferuj `get_column_letter(cell.column)` |
| `cell.parent` | Arkusz, do którego należy komórka | — | `cell.parent.parent` → `Workbook` |
| `cell.offset(row=0, column=0)` | Komórka przesunięta o wektor | liczby całkowite, mogą być ujemne | Wygodne przy poruszaniu się po układzie |
| `cell.get_comment()` / `cell.comment` | Komentarz przypisany do komórki | — | Moduł 16 |
| `ws.values` | Generator krotek wartości (cały arkusz) | — | Każde odwołanie tworzy nowy generator; zapisany generator zużywa się raz |
| `ws.iter_rows(min_row, max_row, min_col, max_col, values_only)` | Generator wierszy | `values_only=True` → krotki wartości; domyślnie krotki `Cell` | Pełna kontrola zakresu |
| `ws.iter_cols(...)` | Jak wyżej, ale kolumnami | — | Uwaga na wydajność (moduł 19) |
| `ws.append(iterable)` | Dopisuje wiersz na końcu | **iterowalne** — uwaga na stringi! | Szczegóły w module 07 |
| `ws.max_row`, `ws.max_column` | Rozmiar modelu | — | Rośnie też od formatowania i od „dotkniętych" komórek |
| `ws.calculate_dimension(force=False)` | Zakres w notacji A1 | `force=True` przelicza od nowa | Moduł 07 |
| `get_column_letter(28)` | `28` → `'AB'` | liczba od 1 | `openpyxl.utils` |
| `column_index_from_string("AB")` | `'AB'` → `28` | tekst | `openpyxl.utils` |
| `coordinate_from_string("AB12")` | `'AB12'` → `('AB', 12)` | tekst | `openpyxl.utils` |
| `coordinate_to_tuple("AB12")` | `'AB12'` → `(12, 28)` | tekst | `openpyxl.utils`; kolejność: (wiersz, kolumna) |
| `range_boundaries("A1:C3")` | `(1, 1, 3, 3)` = `(min_col, min_row, max_col, max_row)` | tekst | Uwaga na kolejność zwracanych wartości! |
| `quote_sheetname("Dane 2026")` | `"'Dane 2026'"` | tekst | Wymagane dla nazw ze spacjami |
| `absolute_coordinate("B1")` | `'$B$1'` | tekst | Przydatne w nazwach i regułach |
| `to_excel(dt, epoch=...)` | `datetime` → serial | `openpyxl.utils.datetime` | Podawaj `datetime`, nie `date` |
| `from_excel(serial, epoch=..., timedelta=False)` | serial → `datetime` | `openpyxl.utils.datetime` | Zwraca `datetime`, nigdy `date` |
| `CALENDAR_WINDOWS_1900`, `CALENDAR_MAC_1904` | Stałe epok | `openpyxl.utils.datetime` | Różnica: 1 462 dni |
| `wb.epoch` | Epoka skoroszytu | — | Ustawiana przy wczytaniu; sprawdź w swojej wersji |
| `wb.defined_names.add(DefinedName(...))` | Tworzy nazwę zdefiniowaną | `DefinedName(name=..., attr_text=...)` | Nowe API w 3.1.x |
| `wb.defined_names["NAZWA"]` | Odczyt nazwy | — | `.destinations` → generator `(arkusz, adres)` |
| `DefinedName(name, attr_text)` | Obiekt nazwy | `openpyxl.workbook.defined_name` | `attr_text` wymaga `quote_sheetname` przy spacjach |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — Tabela typów.** Napisz skrypt `examples/05_cw1.py`, który:

1. Tworzy skoroszyt z arkuszem `Typy`.
2. Wpisuje do kolumny B (od wiersza 2) po jednej wartości następujących typów: `str`, `int`, `float`, `bool` (`True`), `None`, `datetime`, `date`, `time`, `timedelta`, `Decimal`. W kolumnie A umieszcza etykietę typu.
3. Zapisuje plik jako `output/05_cw1.xlsx`.
4. Wczytuje plik i wypisuje dla każdej komórki: etykietę, `repr(value)`, `type(value).__name__`, `cell.data_type`, `cell.is_date`, `cell.number_format`.
5. Wypisuje listę typów, które **zmieniły się** przy przejściu przez plik.

W pliku `output/05_cw1_wnioski.md` uzupełnij tabelę i odpowiedz:

| Wpisany typ | Typ po odczycie | `data_type` | `is_date` | Zmiana? |
|---|---|---|---|---|
| `str` | | | | |
| `int` | | | | |
| ... | | | | |

- Które typy przeszły **bez zmiany**, a które się zmieniły? Dlaczego?
- Co się stało z `Decimal("19.99")`? Czym się różni `float` od `Decimal` w arytmetyce pieniężnej? (Podpowiedź: spróbuj w Pythonie `Decimal("0.1") + Decimal("0.2")` i `0.1 + 0.2`.)
- Zamień `date(2026, 10, 7)` na `datetime(2026, 10, 7, 0, 0)`. Czy teraz typ po odczycie się zgadza? Otwórz plik w Excelu — czy widzisz jakąkolwiek różnicę w wyglądzie tych dwóch komórek?

**Zadanie 2 — Trzy rodzaje pustki.** Zmodyfikuj przykład 3 tak, aby:

1. Do komórki `B2` (pusta) **dopisać wartość, a następnie ją usunąć** przez `ws["B2"].value = None`. Czy po zapisie i odczycie komórka nadal jest pusta?
2. Do `B5` wpisać `False` i dodać formułę `=ISBLANK(B5)` oraz `=COUNT(B2:B5)`.
3. Wypisać w konsoli **własne** przewidywania wyników formuł przed otwarciem pliku, a potem otworzyć plik w Excelu i porównać.

W pliku `output/05_cw2_wnioski.md`:

- Zapisz różnice między przewidywaniem a rzeczywistością (jeśli były).
- Odpowiedz: która z komórek `None`/`""`/`0`/`False` jest — Twoim zdaniem — najbardziej myląca i w jakim kontekście biznesowym prowadzi do błędu?
- Wypisz trzy realne przypadki z pracy z danymi, gdzie `if not cell.value:` dałoby zły wynik.

### 🟡 Warsztat

**Zadanie 3 — `bezpieczna_wartosc()` z testami.** Zaimplementuj funkcję mapującą wartości z pliku na postać bezpieczną dla logiki biznesowej. Wymagania:

1. Sygnatura: `bezpieczna_wartosc(value, *, puste_na_none: bool = False) -> object`.
2. Reguły domyślne: `None` → `""`; `Decimal` → `float`; `datetime`/`date` → tekst w formacie ISO 8601; `str` → po `strip()`; `bool` → bez zmian; liczby → bez zmian; typy nieznane → `str(value)`.
3. Przy `puste_na_none=True` tekst zerowej długości zamieniany jest na `None`.
4. Napisz **drugą** funkcję, odwrotną: `wartosc_do_komorki(value)`, która przygotowuje wartość do zapisu (konwersja `NaN` → `None`, konwersja `numpy.int64` → `int` przez `int()`, konwersja `list`/`dict` → `json.dumps`).
5. Napisz **trzeci** element: `wykryj_problemy(komorka) -> list[str]`, która zwraca listę nazw problemów wykrytych w komórce, np. `["liczba zapisana jako tekst"]`, `["błąd w komórce: #N/A"]`, `["ryzyko utraty cyfr: 18 cyfr"]`, `["podejrzany kod pocztowy: brak zera wiodącego"]`.
6. Napisz testy w `pytest`, które używają obu funkcji na skoroszycie zbudowanym w pamięci (`BytesIO`) — bez zapisywania pliku. Co najmniej 15 przypadków, w tym co najmniej 3 przypadki brzegowe.

W pliku `output/05_cw3_wnioski.md`:

- Wklej uruchomienie `pytest` z wynikiem wszystkich testów.
- Wyjaśnij, dlaczego `bezpieczna_wartosc()` jest funkcją **czystą** i dlaczego to jest zaleta.
- Wyjaśnij, dlaczego rozdzielenie `bezpieczna_wartosc()` (odczyt) od `wartosc_do_komorki()` (zapis) jest lepsze niż jedna funkcja robiąca oba.

### 🔴 Wyzwanie

**Zadanie 4 — Diagnostyka i naprawa pliku z kwotami jako tekst.** To zadanie łączy wszystko z tego modułu i jest zalążkiem pracy z plikami od klientów (moduł 18) oraz walidacji (moduł 21).

Napisz **jedno narzędzie** `examples/05_cw4.py`, które wykonuje pełny cykl: **wykryj → opisz → napraw → zweryfikuj → zraportuj**.

**Krok 1: Generator brudnego pliku.** Funkcja `zbuduj_brudny(cel)`, która tworzy skoroszyt z arkuszem `Faktury` i taką zawartością (nagłówki w wierszu 1):

| Faktura | NIP | Kod pocztowy | Netto | VAT | Brutto |
|---|---|---|---|---|---|
| `"FV/2026/09/001"` | `"1234567890"` | `"00-950"` | `"1 234,56"` | 0.23 | `"1518,51"` |
| `"FV/2026/09/002"` | `1234567890` | 950 | `"89,99"` | 0.23 | `"110,69"` |
| `"FV/2026/09/003"` | `"PL1234567890"` | `"31-501"` | 0 | 0.23 | 0 |
| `"FV/2026/09/004"` | None | `"00950"` | `"#N/A"` | 0.23 | `"0,00"` |
| `"FV/2026/09/005"` | `"99999999999999999999"` | `"00-001"` | `"12 345,67"` | 0.08 | `"13333,32"` |

**Krok 2: Audyt.** Funkcja `audyt(ws) -> dict`, która dla każdej kolumny zwraca:

- liczbę komórek według `data_type`,
- listę adresów komórek z błędami (`data_type == 'e'`),
- listę adresów „liczb zapisanych jako tekst",
- listę adresów komórek, w których istnieje ryzyko utraty cyfr (liczba ≥ 11 cyfr),
- listę adresów komórek, które **wyglądają na kod pocztowy** (tekst pasujący do wzorca `\d{2}-?\d{3}` lub liczba w zakresie 0–99999), ale nie mają poprawnego formatu,
- łączną liczbę problemów.

**Krok 3: Naprawa.** Funkcja `napraw(ws) -> list[str]` (lista opisów zmian), która:

- kolumny `NIP`, `Kod pocztowy` i `Faktura` **wymusza jako tekst** (`str()` + `number_format = "@"`) — bez wyjątków, bo to identyfikatory;
- kody pocztowe **normalizuje** do postaci `XX-XXX` (jeśli wartość ma 5 cyfr, wstawia myślnik; jeśli ma 4 cyfry, dokleja zero wiodące);
- kolumny `Netto`, `VAT`, `Brutto` konwertuje na liczby, obsługując **polski format** (`1 234,56` → `1234.56`) — a jeśli konwersja jest niemożliwa, wpisuje `None` i zapisuje to w liście zmian;
- formatuje `Netto` i `Brutto` jako walutę (`#,##0.00 "zł"`), a `VAT` jako procent (`0%`);
- **zapisuje listę zmian** w arkuszu `Audyt` (nowy arkusz) z kolumnami: `Adres`, `Kolumna`, `Przed`, `Po`, `Powód`;
- zapisuje wynik jako **nowy plik** (`05_faktury_naprawione.xlsx`), nigdy nadpisując źródła.

**Krok 4: Weryfikacja.** Funkcja `zweryfikuj(cel) -> bool`, która wczytuje wynikowy plik i sprawdza:

- brak komórek z `data_type == 'e'`;
- wszystkie kwoty w kolumnach liczbowych są typu liczbowego;
- wszystkie kody pocztowe pasują do wzorca `\d{2}-\d{3}`;
- liczba wierszy danych w pliku wynikowym zgadza się z liczbą wierszy w pliku źródłowym.

**Krok 5: Raport.** Skrypt wypisuje podsumowanie: liczba przebadanych komórek, liczba problemów wg kategorii, liczba naprawionych, wynik weryfikacji (z pełną ścieżką do pliku wynikowego).

**Wymagania dodatkowe (na piątkę):**

- Całe narzędzie uruchamiane z linii poleceń: `python 05_cw4.py sciezka_wejsciowa.xlsx`.
- Jeśli nie podano argumentu, narzędzie samo tworzy brudny plik testowy i przetwarza go.
- Kod ma pracować z **dowolnym** plikiem, w którym istnieją kolumny o tych nazwach — nie zakładaj, że są w konkretnych kolumnach. Znajdź kolumny **po nagłówkach**.
- Jeśli kolumna nie istnieje, narzędzie to raportuje i kontynuuje, zamiast się wywalić.

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""Diagnostyka i naprawa pliku z kwotami jako tekst - rozwiązanie zadania 4.

Uruchom:
    python examples/05_cw4.py                      # tworzy i przetwarza plik testowy
    python examples/05_cw4.py sciezka/do/pliku.xlsx
"""

from __future__ import annotations

import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# --- KONFIGURACJA: wszystko, co dotyczy konkretnego raportu, w jednym miejscu ---
KOLUMNY_TEKSTOWE = {"Faktura", "NIP", "Kod pocztowy"}
KOLUMNY_LICZBOWE = {"Netto", "VAT", "Brutto"}
FORMATY = {
    "Netto": '#,##0.00 "zł"',
    "Brutto": '#,##0.00 "zł"',
    "VAT": "0%",
}
WZORZEC_KODU_POCZTOWEGO = re.compile(r"^\d{2}-\d{3}$")
KODY_BLEDOW = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


# ----------------------------------------------------------------------
# POMOCNIKI
# ----------------------------------------------------------------------
def naglowki(ws) -> dict[str, int]:
    """Mapa: nazwa kolumny -> indeks kolumny (1-based). Po nagłówkach z wiersza 1."""
    return {
        komorka.value: komorka.column
        for komorka in ws[1]
        if komorka.value is not None
    }


def czy_liczba_w_tekscie(value) -> bool:
    """Czy tekst da się zinterpretować jako liczba w formacie PL albo EN?"""
    if not isinstance(value, str):
        return False
    tekst = value.strip().replace("\xa0", "").replace(" ", "")
    tekst = tekst.replace(",", ".")
    if tekst.count(".") > 1:                      # "1.234.567" - separatory tysięcy
        tekst = tekst.replace(".", "", tekst.count(".") - 1)
    return tekst.replace(".", "", 1).replace("-", "", 1).isdigit()


def tekst_na_liczbe(value, *, vat: bool = False):
    """Konwersja tekstu w formacie PL/EN na float. None, gdy się nie da."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    if not isinstance(value, str):
        return None

    tekst = value.strip().replace("\xa0", "").replace(" ", "")
    if tekst.endswith("%"):
        tekst = tekst[:-1].replace(",", ".")
        try:
            return float(tekst) / 100.0
        except ValueError:
            return None

    tekst = tekst.replace(",", ".")
    if tekst.count(".") > 1:
        tekst = tekst.replace(".", "", tekst.count(".") - 1)

    try:
        liczba = float(tekst)
    except ValueError:
        return None

    # Kolumna VAT w formacie "23" oznacza 23%, nie 2300%
    if vat and liczba > 1:
        liczba = liczba / 100.0
    return liczba


def normalizuj_kod_pocztowy(value) -> str | None:
    """'950' -> '00950' -> '00-950'. Zwraca None, gdy nie da się zrobić kodu."""
    if value is None:
        return None
    tekst = str(value).strip().split(".")[0]       # 950.0 -> "950"
    cyfry = re.sub(r"\D", "", tekst)
    if not cyfry:
        return None
    cyfry = cyfry.zfill(5)                         # dokleja zera wiodące
    if len(cyfry) != 5:
        return None
    return f"{cyfry[:2]}-{cyfry[2:]}"


def czy_ryzyko_utraty_cyfr(value) -> bool:
    """Liczba, która w Excelu przekroczy 15 cyfr znaczących."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    return len(str(abs(int(value)))) >= 16 if float(value).is_integer() else False


# ----------------------------------------------------------------------
# KROK 1: GENERATOR BRUDNEGO PLIKU
# ----------------------------------------------------------------------
def zbuduj_brudny(cel: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Faktury"

    ws.append(["Faktura", "NIP", "Kod pocztowy", "Netto", "VAT", "Brutto"])
    ws.append(['FV/2026/09/001', '1234567890', '00-950', '1 234,56', 0.23, '1518,51'])
    ws.append(['FV/2026/09/002', 1234567890, 950, '89,99', 0.23, '110,69'])
    ws.append(['FV/2026/09/003', 'PL1234567890', '31-501', 0, 0.23, 0])
    ws.append(['FV/2026/09/004', None, '00950', '#N/A', 0.23, '0,00'])
    ws.append(['FV/2026/09/005', '99999999999999999999', '00-001',
               '12 345,67', 0.08, '13333,32'])

    wb.properties.description = "Plik testowy celowo zbudowany z błędami typów."
    wb.save(cel)
    return cel


# ----------------------------------------------------------------------
# KROK 2: AUDYT
# ----------------------------------------------------------------------
@dataclass
class Problem:
    adres: str
    kolumna: str
    kategoria: str
    opis: str


def audyt(ws) -> tuple[dict[str, Counter], list[Problem]]:
    """Zwraca (raport typów po kolumnach, lista problemów)."""
    mapa = naglowki(ws)
    typy: dict[str, Counter] = defaultdict(Counter)
    problemy: list[Problem] = []

    for wiersz in ws.iter_rows(min_row=2):
        for komorka in wiersz:
            nazwa = next((n for n, c in mapa.items() if c == komorka.column), None)
            if nazwa is None:
                continue

            typy[nazwa][komorka.data_type] += 1
            v = komorka.value

            if komorka.data_type == "e" or (isinstance(v, str) and v in KODY_BLEDOW):
                problemy.append(Problem(
                    komorka.coordinate, nazwa, "błąd",
                    f"wartość błędu w komórce: {v!r}"))

            if czy_liczba_w_tekscie(v) and nazwa in KOLUMNY_LICZBOWE:
                problemy.append(Problem(
                    komorka.coordinate, nazwa, "liczba jako tekst",
                    f"kwota zapisana jako tekst: {v!r}"))

            if czy_ryzyko_utraty_cyfr(v):
                problemy.append(Problem(
                    komorka.coordinate, nazwa, "ryzyko utraty cyfr",
                    f"liczba o długości >= 16 cyfr: {v!r}"))

            if nazwa == "Kod pocztowy":
                if not isinstance(v, str) or not WZORZEC_KODU_POCZTOWEGO.match(v.strip()):
                    problemy.append(Problem(
                        komorka.coordinate, nazwa, "kod pocztowy",
                        f"niepoprawny format kodu: {v!r}"))

            if nazwa in KOLUMNY_TEKSTOWE and not isinstance(v, str) and v is not None:
                problemy.append(Problem(
                    komorka.coordinate, nazwa, "identyfikator jako liczba",
                    f"identyfikator nie jest tekstem: {v!r}"))

    return dict(typy), problemy


# ----------------------------------------------------------------------
# KROK 3: NAPRAWA
# ----------------------------------------------------------------------
def napraw(ws) -> list[tuple[str, str, object, object, str]]:
    """Naprawia arkusz W MIEJSCU. Zwraca listę krotek (adres, kolumna, przed, po, powód)."""
    mapa = naglowki(ws)
    zmiany: list[tuple[str, str, object, object, str]] = []

    def zapisz(komorka, nazwa: str, nowa, powod: str) -> None:
        if komorka.value != nowa:
            zmiany.append((komorka.coordinate, nazwa, komorka.value, nowa, powod))
            komorka.value = nowa

    for wiersz in ws.iter_rows(min_row=2):
        for komorka in wiersz:
            nazwa = next((n for n, c in mapa.items() if c == komorka.column), None)
            if nazwa is None or komorka.value is None:
                continue

            v = komorka.value

            # --- identyfikatory: zawsze tekst ---
            if nazwa in KOLUMNY_TEKSTOWE:
                nowa = str(v).strip()
                if nazwa == "NIP":
                    # Zostawiamy litery (np. prefiks PL) i cyfry, resztę odrzucamy
                    nowa = re.sub(r"[^A-Za-z0-9]", "", nowa)
                zapisz(komorka, nazwa, nowa, "wymuszony typ tekstowy")
                komorka.number_format = "@"

                if nazwa == "Kod pocztowy":
                    kod = normalizuj_kod_pocztowy(nowa)
                    if kod and kod != nowa:
                        zapisz(komorka, nazwa, kod, "normalizacja kodu pocztowego")
                continue

            # --- kwoty: tekst / błąd -> liczba ---
            if nazwa in KOLUMNY_LICZBOWE:
                if komorka.data_type == "e" or (isinstance(v, str) and v in KODY_BLEDOW):
                    zapisz(komorka, nazwa, None, f"usunięto wartość błędu {v!r}")
                    continue

                liczba = tekst_na_liczbe(v, vat=(nazwa == "VAT"))
                if liczba is None:
                    zapisz(komorka, nazwa, None, f"nie da się skonwertować: {v!r}")
                else:
                    zapisz(komorka, nazwa, liczba, "konwersja tekstu na liczbę")

                komorka.number_format = FORMATY.get(nazwa, "General")

    return zmiany


# ----------------------------------------------------------------------
# KROK 3b: ARKUSZ AUDYTU
# ----------------------------------------------------------------------
def dodaj_arkusz_audytu(wb, zmiany, problemy) -> None:
    ws = wb.create_sheet(title="Audyt")

    ws.append(["Adres", "Kolumna", "Przed", "Po", "Powód"])
    for adres, kolumna, przed, po, powod in zmiany:
        ws.append([adres, kolumna, repr(przed), repr(po), powod])

    ws.append([])
    ws.append(["PODSUMOWANIE PROBLEMÓW WYKRYTYCH"])
    ws.append(["Kategoria", "Liczba"])
    for kategoria, ile in sorted(Counter(p.kategoria for p in problemy).items()):
        ws.append([kategoria, ile])

    naglowek = Font(bold=True, color="FFFFFF")
    tlo = PatternFill(fill_type="solid", fgColor="4472C4")
    for komorka in ws[1]:
        komorka.font = naglowek
        komorka.fill = tlo
        komorka.alignment = Alignment(horizontal="center")

    for kolumna, szerokosc in (("A", 10), ("B", 18), ("C", 24), ("D", 24), ("E", 42)):
        ws.column_dimensions[kolumna].width = szerokosc


# ----------------------------------------------------------------------
# KROK 4: WERYFIKACJA
# ----------------------------------------------------------------------
def zweryfikuj(cel: Path, oczekiwana_liczba_wierszy: int) -> tuple[bool, list[str]]:
    bledy: list[str] = []
    wb = load_workbook(cel)
    try:
        ws = wb["Faktury"]
        mapa = naglowki(ws)

        wierszy = 0
        for wiersz in ws.iter_rows(min_row=2):
            if all(k.value is None for k in wiersz):
                continue
            wierszy += 1
            for komorka in wiersz:
                nazwa = next((n for n, c in mapa.items() if c == komorka.column), None)
                if nazwa is None:
                    continue
                v = komorka.value

                if komorka.data_type == "e":
                    bledy.append(f"{komorka.coordinate}: nadal błąd -> {v!r}")

                if nazwa in KOLUMNY_LICZBOWE and v is not None:
                    if not isinstance(v, (int, float)) or isinstance(v, bool):
                        bledy.append(f"{komorka.coordinate}: {nazwa} nie jest liczbą -> {v!r}")

                if nazwa == "Kod pocztowy" and v is not None:
                    if not WZORZEC_KODU_POCZTOWEGO.match(str(v)):
                        bledy.append(f"{komorka.coordinate}: zły kod pocztowy -> {v!r}")

        if wierszy != oczekiwana_liczba_wierszy:
            bledy.append(
                f"liczba wierszy danych: {wierszy} != oczekiwana {oczekiwana_liczba_wierszy}")
    finally:
        wb.close()

    return (len(bledy) == 0), bledy


# ----------------------------------------------------------------------
# KROK 5: ORKIESTRACJA (to już zalążek warstwy aplikacji - moduł 26)
# ----------------------------------------------------------------------
def main(argv: list[str]) -> int:
    if len(argv) > 1:
        zrodlo = Path(argv[1]).expanduser().resolve()
        if not zrodlo.exists():
            print(f"BŁĄD: plik nie istnieje: {zrodlo}")
            return 1
    else:
        zrodlo = OUTPUT / "05_faktury_brudne.xlsx"
        zbuduj_brudny(zrodlo)
        print(f"utworzono plik testowy: {zrodlo}")

    cel = OUTPUT / "05_faktury_naprawione.xlsx"

    wb = load_workbook(zrodlo)
    try:
        if "Faktury" not in wb.sheetnames:
            print(f"BŁĄD: brak arkusza 'Faktury'. Dostępne: {wb.sheetnames}")
            return 1

        ws = wb["Faktury"]

        # liczymy wiersze danych PRZED zmianami - posłużą do weryfikacji
        liczba_wierszy = sum(
            1 for w in ws.iter_rows(min_row=2)
            if any(k.value is not None for k in w)
        )

        typy, problemy = audyt(ws)
        zmiany = napraw(ws)
        dodaj_arkusz_audytu(wb, zmiany, problemy)

        wb.properties.description = (
            f"Plik znormalizowany automatycznie. Wykryto {len(problemy)} problemów, "
            f"zmieniono {len(zmiany)} komórek. Źródło: {zrodlo.name}"
        )
        wb.save(cel)
    finally:
        wb.close()

    ok, bledy = zweryfikuj(cel, liczba_wierszy)

    print()
    print("=" * 78)
    print("PODSUMOWANIE")
    print("=" * 78)
    print(f"  źródło:            {zrodlo}")
    print(f"  wynik:             {cel}")
    print(f"  wiersze danych:    {liczba_wierszy}")
    print(f"  problemy:          {len(problemy)}")
    for kategoria, ile in sorted(Counter(p.kategoria for p in problemy).items()):
        print(f"    - {kategoria:<24} {ile}")
    print(f"  zmienione komórki: {len(zmiany)}")

    print()
    print("  Typy po kolumnach (przed naprawą):")
    for nazwa, licznik in typy.items():
        rozklad = ", ".join(f"{t}={n}" for t, n in sorted(licznik.items()))
        print(f"    {nazwa:<16} {rozklad}")

    print()
    if ok:
        print("  WERYFIKACJA: OK - plik wynikowy jest spójny.")
    else:
        print("  WERYFIKACJA: PROBLEMY:")
        for blad in bledy:
            print(f"    - {blad}")

    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
```

**Kluczowe decyzje projektowe i dlaczego właśnie tak:**

- **Konfiguracja w jednym miejscu.** Wszystkie reguły (`KOLUMNY_TEKSTOWE`, `KOLUMNY_LICZBOWE`, `FORMATY`, wzorce) są zadeklarowane w nagłówku pliku. Dzięki temu „zmiana reguły biznesowej" to nie wyszukiwanie po kodzie, a zmiana jednej linijki. To zalążek **konfiguracji deklaratywnej** z modułu 26 — tyle że zapisanej w Pythonie, a nie w YAML-u.
- **Kolumny rozpoznawane po nagłówkach, nie po literach.** `naglowki(ws)` buduje mapę nazwa → indeks. Dzięki temu narzędzie działa na pliku, w którym kolumny są w innej kolejności albo przesunięte. To praktyczna wersja zasady „nie buduj na `B2`".
- **Trzy osobne fazy: `audyt` → `napraw` → `zweryfikuj`.** Każda robi jedną rzecz i nic nie wie o pozostałych. `audyt` jest funkcją czytającą — nie modyfikuje arkusza. Dzięki temu możesz uruchomić sam audyt i tylko **opisać** plik klientowi, bez zmieniania go. To jest **rozdział odpowiedzialności**, do którego wrócimy w module 22.
- **`Problem` jako `dataclass`, nie `dict`.** Struktura jest opisana typem, ma nazwane pola, jest czytelna w debugerze i może być rozszerzona bez psucia kodu, który ją czyta. W module 23 nazwiemy to **DTO**.
- **`napraw()` zwraca listę zmian i nie wypisuje niczego.** Cała komunikacja z użytkownikiem jest w `main()`. Funkcja, która drukuje, jest funkcją, której nie da się ponownie użyć — a tutaj chcesz ją wywołać także z testu i z serwera.
- **`None` jako wynik nieudanej konwersji, a nie zero.** Wpisanie `0` w miejsce, gdzie nie dało się policzyć kwoty, jest **wpisaniem fałszywej informacji** — i to takiej, której nikt nie zauważy. `None` jest uczciwe: „tu nie ma poprawnej wartości", i widać to w `COUNT`.
- **Zapis do nowego pliku.** Źródło nigdy nie jest nadpisywane — to reguła z modułu 00 i 04, tutaj zrealizowana w praktyce.
- **Narzędzie weryfikuje się samo.** `zweryfikuj()` to osobny przebieg po **pliku z dysku**, nie po modelu. To jedyny sposób, by mieć pewność, że to, co zapisaliśmy, naprawdę się zapisało.
- **Kody pocztowe przez `zfill(5)`.** Wartość `950` (zinterpretowana przez kogoś jako liczba) zamienia się w `00950`, a potem w `00-950`. Bez tego kroku `950` zostałoby „naprawione" na `95-0` — czyli gorzej niż przed naprawą. **Zawsze sprawdzaj, czy normalizacja nie tworzy śmieci.**
- **`vat=True` przy konwersji VAT.** Kolumna `VAT` może przychodzić raz jako `0.23`, raz jako `23`. Heurystyka „powyżej 1 to procent" jest niedoskonała (stawka 100% też jest możliwa), więc **trzeba ją zweryfikować** — w prawdziwym narzędziu dodałbyś do arkusza audytu ostrzeżenie o takich komórkach. To dobry przykład, że konwersja danych nigdy nie jest w pełni automatyczna, i że uczciwe narzędzie pokazuje swoje założenia.

**Czego ten szkic jeszcze nie robi** (a w wersji produkcyjnej powinien): zapisu atomowego z kopią zapasową (moduł 18), notowania wartości źródłowych w pliku wynikowym dla audytu, limitu rozmiaru pliku wejściowego i sanitacji danych użytkownika (moduł 21), testów (moduł 20). Ten szkic jest **poprawną logiką w niepełnej oprawie** — i dokładnie tak powinna wyglądać nauka: najpierw logika, potem oprawa.

</details>

## 6. Typowe błędy i pułapki

**1. „Kolumna z kwotami jest pusta w formule `SUMA` — zwraca zero" (objaw) → kwoty przyszły jako **tekst** (`data_type == 's'`), najczęściej z eksportu albo przez wklejenie z zachowaniem formatowania; Excel i `SUMA` ignorują tekst całkowicie (objaw wtórny: liczby wyrównane do lewej, zielone trójkąty w rogach) (przyczyna) → wykryj przez `data_type == 's'` na komórkach, które wyglądają na liczby, i przekonwertuj na `float` w Pythonie (przykład 5), a w Excelu — funkcją „Tekst jako kolumny" albo przez `=WARTOŚĆ()` (naprawa).**

Diagnostyka „na oko", która działa zawsze: **liczby i daty Excel wyrównuje do prawej, tekst do lewej.** Jedno spojrzenie w kolumnę często wystarcza, żeby rozpoznać problem, zanim napiszesz linijkę kodu.

**2. „Tabela ma 12,5 tysiąca wierszy, a się nie otwiera"/„`np.int64` daje `ValueError: Cannot convert 42 to Excel`" (objaw) → typy spoza modelu openpyxl: `list`, `dict`, `set`, `tuple`, obiekty własnych klas, a z ekosystemu danych — `numpy.int64` i inne typy bez dziedziczenia po wbudowanych `int`/`float`/`str` (przyczyna) → doprowadź typ do wbudowanego (`int()`, `float()`, `str()`), dla struktur użyj `json.dumps()`, a najlepiej **rozbij strukturę na kolumny**; w pipeline'ach z pandas konwertuj typy wcześniej (naprawa).**

Zwróć uwagę, że `np.float64` **przechodzi** (dziedziczy po `float`), a `np.int64` nie. Ta niespójność jest źródłem wielu godzin debugowania — jeśli pracujesz z pandas, zawsze rób `int(...)` albo `df.astype(object)` przed pętlą zapisu.

**3. „Zero w kolumnie zniknęło — wcześniej były zera wiodące, teraz nie ma" (objaw) → Excel zinterpretował wartość jako liczbę i odciął zera wiodące; dotyczy kodów pocztowych (`00-950` zapisane jako liczba `950`), numerów seryjnych, silników PESEL, NIP-ów, numerów kont (przyczyna) → zapisuj identyfikatory **jako tekst** (`str()` + cudzysłów przy wpisywaniu) i **ustaw format komórki na tekstowy** (`cell.number_format = "@"`), żeby Excel nie „poprawiał" ich przy edycji; napraw już zepsute dane funkcją `normalizuj_kod_pocztowy()` z ćwiczenia 🔴 (naprawa).**

Sama zamiana na tekst nie wystarczy, jeśli format pozostanie `General` — użytkownik kliknie komórkę, poprawi jedną cyfrę i Excel natychmiast odetnie zera. Format `@` jest tu częścią rozwiązania, nie kosmetyką.

**4. „Napisałem `if not cell.value:` i przeszło, gdy wartość to zero, a nie powinno" (objaw) → w Pythonie `not` traktuje `None`, `""` i `0` identycznie, a w danych finansowych zero jest jedną z najczęstszych prawdziwych wartości (zwrot, korekta, zerowy VAT) (przyczyna) → używaj jawnych porównań: `if cell.value is None:` dla braku danych albo `if cell.value in (None, "")` dla „pustego"; nigdy `if not cell.value:` (naprawa).**

Ten błąd jest zdradziecki, bo wprowadza **ciche przekłamanie wyniku**, a nie wyjątek. Test, który go złapie, to test z zerem w danych wejściowych — i dlatego zawsze trzeba pisać przypadki brzegowe z `0`, `None` i `""`.

**5. „`datetime.date` zapisałem, a po odczycie dostałem `datetime.datetime`" (objaw) → Excel nie ma osobnego typu dla daty i dla daty z godziną; istnieje jeden serial dni, a format decyduje, co widać. `date(2026,10,7)` staje się północą `2026-10-07 00:00:00` (przyczyna) → normalizuj typ po odczycie przed porównaniem: `value.date()` jeśli potrzebujesz tylko daty, albo zapisuj `datetime` od początku i nie mieszaj `date` z `datetime` w porównaniach (naprawa).**

Ten błąd ma bardzo konkretny objaw: **kod przechodzi, ale porównanie `if value == date(2026, 10, 7)` zawsze zwraca `False`** — i nikt nie widzi, że dane są w porządku, a zepsuty jest operator porównania.

**6. „`cell.font.bold = True` nie działa" — w wersji typowej (objaw) → to nie jest błąd typów, ale warto go tu wymienić, bo ma dokładnie tę samą przyczynę co większość problemów tego modułu: **obiekt w komórce jest niemutowalny** i przypisanie czegoś do jego atrybutu nie zmienia modelu; powstaje nowy efemeryczny obiekt, który nigdzie nie trafia. Dotyczy to także stylu: `komorka.number_format = "..."` działa (bo to właściwość komórki), ale `komorka.font.bold = True` nie działa (bo to właściwość obiektu stylu) (przyczyna) → `komorka.font = Font(bold=True, ...)` albo `nowy = copy(komorka.font); nowy.bold = True; komorka.font = nowy` (naprawa).**

Warto zwrócić uwagę: `cell.number_format = "@"` **działa**, a `cell.font.bold = True` **nie działa**. To nie jest niespójność — `number_format` jest atrybutem komórki, a `bold` atrybutem obiektu stylu. Moduł 09 wyjaśni to systematycznie.

**7. „Data pokazuje się w komórce jako `46302` zamiast daty" (objaw) → liczba jest zapisana z formatem `General` albo ktoś wyczyścił format; openpyxl (i Excel) nie mają innej informacji o tym, że to data, niż format komórki (przyczyna) → ustaw format jawnie: `komorka.number_format = "yyyy-mm-dd"`; przy wczytywaniu cudzych plików sprawdzaj `komorka.is_date` i konwertuj `from_excel()` w razie potrzeby; „naprawa" wyłącznie przez nadanie formatu, bo wartość jest poprawna (naprawa).**

Priorytetowa wersja tego błędu: wczytujesz plik, w którym **część** komórek ma format daty, a część nie. Dostaniesz kolumnę, w której mieszają się `datetime` i `int`. Wtedy `max()` albo sortowanie dadzą bezsensowne wyniki — i trzeba to znormalizować jawnie (`na_date()` z sekcji 2.9).

**8. „`TypeError` z komunikatem o strefach czasowych przy zapisie daty" (objaw) → openpyxl nie obsługuje `datetime` ze strefą (`tzinfo`), bo Excel nie ma gdzie zapisać przesunięcia; komunikat w rodzaju „Excel does not support datetimes with timezones" (przyczyna) → usuń `tzinfo` **świadomie**, po konwersji do właściwej strefy: `dt.astimezone().replace(tzinfo=None)`; strefę zapisz jako osobny tekst/konfigurację (naprawa).**

Zwróć uwagę na kolejność: najpierw `astimezone()` (konwersja wskazania zegara), potem `replace(tzinfo=None)` (usunięcie metadanych). Odwrotna kolejność **zmieni wartość godziny** — bo usunięcie strefy bez konwersji zostawia czas UTC udający czas lokalny. To klasyczny błąd o skutkach wykrywalnych dopiero po miesiącach.

**9. „`Decimal` wpisałem, a wyliczenia dają `1234.5600000000001`" (objaw) → `Decimal` jest akceptowany przy zapisie jako liczba, ale w pliku i po odczycie jest to zwykła liczba zmiennoprzecinkowa; arytmetyka na `float` ma błędy reprezentacji (przyczyna) → prowadź **całą arytmetykę pieniężną w `Decimal`**, a do komórki wstawiaj już policzony wynik; przy odczycie pamiętaj, że dostaniesz `float`, więc albo konwertuj `Decimal(str(value))`, albo — najlepiej — licz w **groszach jako `int`** (naprawa).**

`Decimal(str(value))` zamiast `Decimal(value)` to ważny szczegół: `Decimal(0.1)` zapamiętuje wszystkie błędy reprezentacji binarnej, a `Decimal("0.1")` — nie. Konwersja przez `str()` jest jedyną bezpieczną drogą z `float` do `Decimal`.

**10. „Kolumna z klasyfikacją ma `PRAWDA`/`FAŁSZ` jako tekst, a nie jako wartości logiczne" (objaw) → ktoś zapisał je w cudzysłowie (`"PRAWDA"`) albo dane przyszły z CSV; Excel traktuje je jak zwykły tekst, więc porównania i filtry działają inaczej (przyczyna) → rozpoznaj po wyrównaniu (tekst do lewej, wartość logiczna do środka) i `data_type == 's'`; konwertuj świadomie: `value.strip().upper() in ("PRAWDA", "TRUE")` → `True` (naprawa).**

Dodatkowa pułapka po stronie Pythona: `True == 1` i `False == 0` zwracają `True`. Jeśli w jednej kolumnie mieszają się boole i liczby, `sum()` policzy `PRAWDA` jako 1 — czasem to dobrze, a czasem to cichy błąd. Filtruj przez `isinstance(v, bool)`.

**11. „W komórce jest `#N/A`, ale mój skrypt się nie wywalił" (objaw) → wartość błędu wczytuje się jako **tekst** (`data_type == 'e'`, `value == '#N/A'`), a nie jako wyjątek; Python traktuje ją jak zwykły `str`, więc `sum()` może polec dopiero na etapie dodawania, a w innych kontekstach przejść cicho (przyczyna) → sprawdzaj `komorka.data_type == "e"` **oraz** czy tekst nie jest jednym z kodów błędów; filtruj komórki tak, by do logiki biznesowej wchodziły wyłącznie wartości sprawdzone (naprawa).**

Druga strona tego samego problemu: **nie da się łatwo zapisać prawdziwego błędu z Pythona.** Zapisanie `"#N/A"` utworzy tekst, który wygląda jak błąd, ale nim nie jest. W raporcie mogą więc współistnieć dwa różne `#N/A` — i to jest dokładnie ten rodzaj niejednoznaczności, którego warto unikać: lepiej wpisać `None` i opis w kolumnie „uwagi".

**12. „Długi opis w komórce został ucięty / Excel zgłasza problem z plikiem" (objaw) → przekroczenie limitu 32 767 znaków na komórkę; openpyxl ostrzega o maksymalnej długości łańcucha i może obciąć wartość, a Excel może zgłosić potrzebę naprawy pliku — sprawdź zachowanie w swojej wersji (przyczyna) → **nie polegaj na obcinaniu**: dziel tekst samodzielnie (`podziel_tekst()` z sekcji 2.8) i zapisuj kawałki do kolejnych komórek, albo zapisz tekst jako plik obok i wstaw hiperlink (moduł 16) (naprawa).**

Powiązany limit: **32 767 to także granica dla formuł** (8 192 znaki) i dla nazw arkuszy (31 znaków). Jeśli budujesz tekst dynamicznie, waliduj jego długość **przed** przypisaniem, a nie po fakcie.

**13. „Wielki identyfikator, który zapisałem, ma inną wartość po otwarciu w Excelu" (objaw) → Excel trzyma liczby jako liczby zmiennoprzecinkowe podwójnej precyzji i wyświetla maksymalnie 15 cyfr znaczących; 18-cyfrowy numer zostaje zaokrąglony przy pierwszym otwarciu i zapisaniu pliku, nawet jeśli openpyxl zapisał pełne cyfry (przyczyna) → wszystkie identyfikatory zapisuj jako **tekst** z formatem `"@"`; jeśli dopiero diagnozujesz, porównaj `repr(cell.value)` z tym, co widać w Excelu — to najszybszy sposób wykrycia (naprawa).**

Ten błąd jest wyjątkowo zdradliwy, bo **zależy od tego, kto czyta plik**: openpyxl zwróci to, co jest w XML, a Excel — to, co jego model potrafi zmieścić. Ten sam plik może więc pokazać dwie różne wartości. Dlatego identyfikatory to dane tekstowe; nie ma tu wyjątków.

**14. „`NaN` z pandas wylądowało w komórce jako tekst `'nan'` i kolumna się rozsypała" (objaw) → `float('nan')` nie ma reprezentacji w Excelu; jeśli w ścieżce zapisu pojawi się `str(nan)`, do pliku trafia zwykły tekst, a kolumna miesza liczby z tekstem — albo, przy innym przebiegu, wartość błędu/pustka, więc wynik po odczycie nie jest tym, co wpisałeś (przyczyna) → konwertuj `NaN` na `None` **sam, przed zapisem** (`bez_nan()` z sekcji 2.10), i nigdy nie przepuszczaj `NaN` przez `str()`; po odczycie sprawdź `data_type` kolumny (naprawa).**

Szersza zasada: **puste komórki to jedyna sensowna reprezentacja braku danych w Excelu.** Nie tekst `"nan"`, nie `"N/A"`, nie `"-"`, nie `0`. Każda z tych atrap wcześniej czy później zostanie policzona albo porównana jak prawdziwa wartość.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Komórka to obiekt, a jej treść to `cell.value`.** `ws["A1"]` zwraca `Cell`, nie wartość — musisz napisać `.value`. Komórka wie o wiele więcej niż samą treść: `data_type` (jaki to typ: `'n'`, `'s'`, `'d'`, `'b'`, `'f'`, `'e'`), `number_format` (jak to pokazać), `is_date`, `coordinate`. **Dotknięcie komórki tworzy ją w modelu** — czytanie też jest modyfikacją. Adresuj po A1, gdy adres jest stały; po `(row, column)`, gdy jest obliczony — oba indeksy **od 1**.

2. **Excel przyjmuje zamkniętą listę typów: `str`, `int`, `float`, `bool`, `None`, `datetime`, `date`, `time`, `timedelta`, `Decimal`.** Wszystko inne — `list`, `dict`, `set`, `numpy.int64`, obiekty własnych klas — kończy się `ValueError: Cannot convert ... to Excel`. Naprawa jest zawsze jedna z czterech: konwersja do typu wbudowanego, serializacja (`json.dumps`), reprezentacja tekstowa, **albo — najlepiej — rozbicie struktury na kolumny**. A na wyjściu pamiętaj: `Decimal` wraca jako `float`, a `date` jako `datetime`.

3. **Pusta komórka nie jest tym samym co pusty tekst ani co zero.** `None` to brak krzesła, `""` to krzesło bez osoby, `0` to osoba milcząca — i Excel traktuje je trzy różne sposoby w `ISBLANK`, `COUNTA`, `COUNTBLANK` i `AVERAGE`. W Pythonie pisz `if value is None:` zamiast `if not value:`, bo drugie **skasuje Ci zera** — a zero w danych finansowych jest jedną z najczęstszych prawdziwych wartości. Komórkę „czyścisz" przez `None`, nigdy przez `""`.

4. **Dwie najczęstsze katastrofy typów to identyfikator i data.** Identyfikatory (NIP, kod pocztowy, numer faktury, IBAN) **zawsze** zapisuj jako tekst z formatem `"@"` — inaczej tracą zera wiodące, a długie liczby tracą cyfry, bo Excel trzyma tylko 15 cyfr znaczących i zaokrągla przy pierwszym otwarciu. Daty są w pliku **liczbami dni od epoki**; format komórki decyduje, czy openpyxl rozpozna je jako `datetime`, więc **ustawiaj format jawnie** (`yyyy-mm-dd`, `[hh]:mm:ss` dla czasów trwania) i nigdy nie przepuszczaj dat ze strefą czasową — openpyxl ich nie obsługuje, bo Excel nie ma gdzie zapisać przesunięcia.

5. **Wartości wracają z pliku w postaci, której nie kontrolujesz — więc normalizuj je w jednym miejscu.** Wartości błędów (`#N/A`, `#DIV/0!`) wczytują się jako **tekst**, nie jako wyjątki, i przechodzą przez Twój kod cicho; liczby wracają raz jako `int`, raz jako `float`; kolumna z kwotami może mieszać liczby i tekst. Rozwiązanie jest jedno i architektoniczne: **jedna funkcja normalizująca** (`bezpieczna_wartosc()`) i **jedna funkcja przygotowująca do zapisu** (`wartosc_do_komorki()`), do których prowadzi każda wartość przekraczająca granicę plik↔kod. Ten podział — czysta logika mapowania oddzielona od operacji na arkuszu — to zalążek Adaptera, który rozbudujemy w module 24.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal
from math import isnan
from pathlib import Path
from io import BytesIO
import re

from openpyxl import Workbook, load_workbook
from openpyxl.utils import (
    get_column_letter,
    column_index_from_string,
    coordinate_from_string,      # 'AB12' -> ('AB', 12)
    coordinate_to_tuple,         # 'AB12' -> (12, 28)
    range_boundaries,            # 'A1:C3' -> (1, 1, 3, 3)  (min_col, min_row, max_col, max_row)
    quote_sheetname,             # 'Dane 2026' -> "'Dane 2026'"
    absolute_coordinate,         # 'B1' -> '$B$1'
)
from openpyxl.utils.datetime import (
    to_excel, from_excel,
    CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904,
)
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# ==================================================================
# 2. ZAPIS - ADRESOWANIE
# ==================================================================
wb = Workbook()
ws = wb.active

ws["A1"] = "Produkt"                    # notacja A1 - czytelna
ws.cell(row=2, column=1, value="Widzet")  # adres liczbowy - do pętli; OBA od 1!

ws["B2"] = 1234.5
ws["C2"] = datetime(2026, 10, 7, 9, 30)
ws["C2"].number_format = "yyyy-mm-dd hh:mm:ss"   # USTAWIAJ FORMAT JAWNIE

ws["D2"] = timedelta(hours=27, minutes=15)
ws["D2"].number_format = "[hh]:mm:ss"            # NAWIASY! bez nich 03:15

ws["E2"] = "007"                                  # identyfikator = tekst
ws["E2"].number_format = "@"                      # format tekstowy

# ==================================================================
# 3. ODCZYT - KOMÓRKA I JEJ METADANE
# ==================================================================
k = ws["B2"]
k.value            # 1234.5
k.data_type        # 'n' | 's' | 'str' | 'inlineStr' | 'b' | 'd' | 'f' | 'e'
k.number_format    # 'General'
k.is_date          # True, gdy format liczby jest formatem daty
k.coordinate       # 'B2'
k.row, k.column    # 2, 2   (od 1)
k.column_letter    # 'B'
k.parent           # <Worksheet>  (k.parent.parent -> <Workbook>)
k.offset(row=1, column=0)         # komórka o wiersz niżej

# DIAGNOSTYKA PIERWSZEGO KONTAKTU:
def opisz(komorka):
    return (f"{komorka.coordinate}: data_type={komorka.data_type!r} "
            f"value={komorka.value!r} typ={type(komorka.value).__name__}")

# ==================================================================
# 4. TYPY - CO PRZEJDZIE, A CO NIE
# ==================================================================
# PRZEJDZIE:
#   str, int, float, bool, None,
#   datetime.datetime / .date / .time / .timedelta, Decimal
#
# NIE PRZEJDZIE (ValueError: Cannot convert ... to Excel):
#   list, dict, set, tuple, bytes, object(), numpy.int64, własne klasy

# NAPRAWY:
ws["A1"] = int(np_value)                      # obcy typ liczbowy -> wbudowany int
ws["A2"] = json.dumps(slownik, ensure_ascii=False)   # struktura -> tekst JSON
ws["A3"] = ", ".join(map(str, lista))         # lista -> tekst
# NAJLEPSZE: rozbij strukturę na kolumny (nagłówki + wiersze)

# ==================================================================
# 5. PUstka: None vs "" vs 0
# ==================================================================
ws["A1"] = None        # brak wartości - komórka NIE trafia do XML
ws["A2"] = ""          # pusty TEKST - trafia; ISBLANK = FALSE; COUNTA liczy
ws["A3"] = 0           # zero - pełnoprawna liczba

komorka.value = None   # "czyszczenie" - ustawiaj None, nigdy ""

# WARUNKI W PYTHONIE:
if komorka.value is None:                          # tylko brak danych
if komorka.value in (None, ""):                    # brak albo pusty tekst
if isinstance(komorka.value, (int, float)) and not isinstance(komorka.value, bool):
    ...                                            # TYLKO liczby (bez bool!)

# ==================================================================
# 6. DATY - SERIAL, FORMAT, KONWERSJA
# ==================================================================
from openpyxl.utils.datetime import to_excel, from_excel

to_excel(datetime(2026, 10, 7))          # 46302.0
to_excel(datetime(2026, 10, 7, 9, 30))   # 46302.395833333336
from_excel(46302)                        # datetime(2026, 10, 7, 0, 0)  <- NIE date!

# Zwykły date -> datetime przed konwersją:
to_excel(datetime.combine(date(2026, 10, 7), time.min))

# Epoka:
wb.epoch                                 # 2415018.5 (Windows 1900) - sprawdź w swojej wersji
CALENDAR_MAC_1904 - CALENDAR_WINDOWS_1900   # 1462.0 dni różnicy

# ROZPOZNANIE DATY ZALEŻY OD FORMATU:
komorka.is_date                          # True, gdy format jest formatem daty

# NORMALIZACJA WCZYTYWANEJ DATY (liczby i datetime razem):
def na_date(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return from_excel(value)
    return None

# FORMATY, KTÓRE WARTO USTAWIAĆ JAWNIE:
#   datetime    -> 'yyyy-mm-dd hh:mm:ss'
#   date        -> 'yyyy-mm-dd'
#   time        -> 'hh:mm:ss'
#   timedelta   -> '[hh]:mm:ss'   (NAWIASASY dla czasów > 24 h)
#   tekst       -> '@'
#   waluta      -> '#,##0.00 "zł"'
#   procent     -> '0%'
#   tysiące     -> '#,##0'

# STREFY CZASOWE - openpyxl NIE OBSŁUGUJE:
# zle:   ws["A1"] = datetime.now(timezone.utc)
# dobrze:
ws["A1"] = datetime.now(timezone.utc).astimezone().replace(tzinfo=None)
#          ^ konwersja wskazania zegara        ^ dopiero potem usunięcie strefy

# ==================================================================
# 7. BŁĘDY I NaN
# ==================================================================
KODY_BLEDOW = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
# Wartość błędu wczytuje się jako TEKST (data_type == 'e'), BEZ wyjątku!

def czy_blad(komorka):
    return komorka.data_type == "e" or (
        isinstance(komorka.value, str) and komorka.value in KODY_BLEDOW
    )

def bez_nan(value):
    return None if isinstance(value, float) and isnan(value) else value
# NaN -> None PRZED zapisem. Nigdy str(nan) - to da tekst 'nan'.

# ==================================================================
# 8. CZYTANIE WIERSZAMI
# ==================================================================
for wiersz in ws.values:                  # krotki WARTOŚCI, cały arkusz
    print(wiersz)

for wiersz in ws.iter_rows(values_only=True):      # tylko wartości, zakres kontrolowany
    print(wiersz)

for wiersz in ws.iter_rows(min_row=2, max_col=6):  # krotki OBIEKTÓW Cell
    for komorka in wiersz:
        if komorka.data_type == "e":
            print(f"błąd: {komorka.coordinate}")

# UWAGA: generator zapisany w zmiennej zużywa się przy pierwszym przejściu!
dane = ws.values
list(dane)      # [...] - pełna lista
list(dane)      # [] - taśma się skończyła
# Rozwiązanie: dane = list(ws.values)  ALBO odwołaj się do ws.values ponownie

# ==================================================================
# 9. WARSTWA NORMALIZACJI - WZORZEC DO KOPIOWANIA
# ==================================================================
def bezpieczna_wartosc(value):
    """Wartość z komórki -> postać bezpieczna dla logiki biznesowej."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return value                       # bool PRZED liczbami (bool dziedziczy po int!)
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str):
        return value.strip()
    return value


def wartosc_do_komorki(value):
    """Wartość z logiki biznesowej -> postać akceptowana przez openpyxl."""
    if isinstance(value, float) and isnan(value):
        return None
    if isinstance(value, str) and not value.strip():
        return None                        # puste komórki, nie puste teksty
    if isinstance(value, (list, dict, set, tuple)):
        raise TypeError(
            f"Typ {type(value).__name__} nie należy do komórki. "
            f"Rozbij go na kolumny albo użyj json.dumps()."
        )
    return value

# ==================================================================
# 10. NAZWY ZDEFINIOWANE (ZALĄŻEK KONFIGURACJI)
# ==================================================================
konf = wb.create_sheet(title="Konfiguracja")
konf["A1"], konf["B1"] = "Stawka VAT", 0.23

odniesienie = f"{quote_sheetname(konf.title)}!{absolute_coordinate('B1')}"
wb.defined_names.add(DefinedName(name="STAWKA_VAT", attr_text=odniesienie))

for arkusz, adres in wb.defined_names["STAWKA_VAT"].destinations:
    print(f"STAWKA_VAT -> {arkusz}!{adres}")

# ==================================================================
# 11. SZKIELET SKRYPTU TYPUJĄCEGO (wzorzec do kopiowania)
# ==================================================================
# 1. KONFIGURACJA na górze (nazwy kolumn, formaty, wzorce)
# 2. ROOT / OUTPUT od __file__ + mkdir(parents=True, exist_ok=True)
# 3. FUNKCJE CZYSTE: parsowanie i walidacja (bez arkusza, łatwo testować)
# 4. FUNKCJA CZYTAJĄCA: audyt - nic nie modyfikuje
# 5. FUNKCJA PISZĄCA: napraw - zwraca listę zmian, nic nie drukuje
# 6. FUNKCJA WERYFIKUJĄCA: czyta plik z dysku i sprawdza warunki
# 7. main(): spina wszystko, drukuje raport, ZAPISUJE DO NOWEGO PLIKU
# 8. if __name__ == "__main__": raise SystemExit(main())
```

## 9. Co dalej

Umiesz już wpisać do komórki dokładnie to, co chcesz, i odczytać z niej dokładnie to, co jest — łącznie z rozpoznaniem, gdy coś po drodze się zmieniło. To ważniejsza umiejętność, niż się wydaje: w praktyce **większość czasu w automatyzacji Excela nie idzie na pisanie raportów, a na zmaganie się z danymi, które nie są tym, czym wyglądają**. Kod, który umie to wykryć i nazwać, jest wart dziesięć razy więcej niż kod, który „chyba działa".

W **module 06** zajmiemy się tym, co w komórce wygląda jak liczba, ale nią nie jest: **formułami**. Zobaczysz, że openpyxl traktuje formułę jako **tekst zaczynający się od `=`**, że nie policzy jej za Ciebie (i nigdy nie będzie), że `data_only=True` czyta **wartości z pamięci podręcznej**, a nie wyniki, oraz — najważniejsze — że **zapisanie skoroszytu wczytanego przez `data_only=True` zamienia wszystkie formuły na wartości**. To jeden z tych błędów, których naprawienie graniczy z niemożliwością, jeśli nie masz kopii zapasowej. Zapowiedź: przyda się `zapisz_bez_pomylek()` z modułu 04.

Przygotuj do modułu 06:

- **plik z tego modułu** (`output/05_sprzedaz.xlsx`) oraz jego kopię — będziesz na nim testować odczyt formuł i wartości z cache,
- **jeden plik utworzony w Excelu**, w którym w kolumnie B są liczby, w kolumnie C formuły (`=B2*1.23`), a w kolumnie D wyniki tych formuł — i **otwórz go i zamknij w Excelu**, żeby plik miał zapisane wartości w pamięci podręcznej. Bez tego kroku `data_only=True` zwróci Ci `None` i nie zrozumiesz dlaczego,
- **odpowiedź na pytanie**: w których miejscach tego modułu zmieniałeś **typ** wartości w komórce, a w których tylko **format jej wyświetlania**? Powinieneś umieć wskazać palcem oba rodzaje linii i wyjaśnić, czy wpłynęły na wartość, czy tylko na jej wygląd.