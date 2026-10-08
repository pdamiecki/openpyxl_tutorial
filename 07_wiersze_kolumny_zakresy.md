# Moduł 07 — Wiersze, kolumny i zakresy

> **Część:** I — Podstawy pracy ze skoroszytem · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–06

## 0. W tym module nauczysz się

- **Rozumiesz, że `max_row` i `max_column` nie mówią „gdzie kończą się dane", a „gdzie kończą się komórki, o których openpyxl wie".** To dwa różne pytania i właśnie na tym rozróżnieniu wykłada się większość pętli po arkuszach. Zobaczysz, dlaczego samo *dotknięcie* dalekiej komórki powiększa `max_row`, i jak znaleźć naprawdę ostatni wiersz z danymi.
- **Odróżniasz cztery sposoby patrzenia na rozmiar arkusza** — `max_row`/`max_column`, `min_row`/`min_column`, `dimensions`/`calculate_dimension()` i „rozmiar logiczny danych" — oraz wiesz, do czego każde z nich służy i kiedy żadne nie jest tym, czego szukasz.
- **Umiesz iterować po siatce trzema drogami** (`for row in ws`, `iter_rows`, `iter_cols`) i rozumiesz, dlaczego iteracja wierszami jest domyślnym wyborem, a skakanie po komórkach to najczęstsza przyczyna wolnych skryptów. Wiesz też, że `values_only=True` zwraca **krotki wartości**, a nie obiekty komórek.
- **Znasz pułapkę `append`**: przyjmuje **iterowalne**, więc `append("ABC")` zapisze `A`, `B`, `C` w trzech komórkach. Wiesz, jak rozpoznać ten błąd i jak go uniknąć strukturalnie.
- **Wiesz dokładnie, co się przesuwa, a co nie** przy `insert_rows`, `delete_rows`, `insert_cols` i `delete_cols`: wartości tak, style w ograniczonym zakresie, **formuły — nie**, wymiary wierszy — nie, scalenia i formatowanie warunkowe — nie.
- **Opanowałeś `ws.move_range(..., translate=True)`** — jedyną operację w openpyxl, która „rozumie" referencje w formułach i przesuwa je razem z komórkami. Wiesz, kiedy to działa, a kiedy jest pułapką.
- **Budujesz adresy i zakresy programowo, a nie ręcznie**, przy pomocy `get_column_letter`, `column_index_from_string`, `coordinate_from_string`, `coordinate_to_tuple`, `range_boundaries`, `CellRange`, `quote_sheetname` i `absolute_coordinate` — i wiesz, że to jest **jedyne** poprawne podejście, gdy nazwy arkuszy lub numery kolumn pochodzą z konfiguracji albo od użytkownika.

## 1. Intuicja i analogia

### Analogia główna: geodeta i plan miasta

Wyobraź sobie, że arkusz to **plan miasta naniesiony na siatkę współrzędnych**. Twoja praca to opisywanie tego planu — a openpyxl to geodeta, który za Ciebie liczy i rysuje.

Zbudujmy słownik pojęć od podstaw:

- **Komórka to jedna działka ewidencyjna.** Ma adres: litera kolumny to nazwa alei biegnącej z północy na południe, numer wiersza to nazwa ulicy biegnącej ze wschodu na zachód. `B7` to „Aleja B, Ulica 7".
- **Wiersz to cała ulica** — wszystkie działki położone przy Ulicy 7, od Alei A do Alei ZZ.
- **Kolumna to cała aleja** — wszystkie działki w Alei B, od Ulicy 1 do Ulicy 1 048 576.
- **Zakres to dzielnica** — prostokąt wyznaczony przez dwie ulice i dwie aleje. `A1:C10` to „między Aleją A a Aleją C, między Ulicą 1 a Ulicą 10".

Teraz najważniejsza rzecz w całym module:

> **`max_row` to numer ostatniej ulicy, na której cokolwiek jest zapisane w ewidencji — nawet jeśli to tylko słupki geodezyjne i wydeptana trawa. To nie to samo, co „ostatnia ulica z domami".**

I dlatego pytanie „gdzie kończą się moje dane?" **nie ma odpowiedzi udzielanej przez `max_row`.** Ma odpowiedź, którą musisz uzyskać sam — przeglądając ewidencję działka po działce. To cały sens sekcji 2.11 tego modułu.

Dwie operacje edycyjne też mają świetne odpowiedniki urbanistyczne:

- **`insert_rows(5)` to wciśnięcie nowej ulicy w środek miasta.** Wszystko poniżej przesuwa się o jeden numer w dół — ale **akta notarialne nie aktualizują się same**. Jeśli w akcie było napisane „działka przy Ulicy 6", to po wciśnięciu nowej ulicy nadal tam będzie napisane „Ulica 6" — mimo że fizycznie ta działka jest teraz przy Ulicy 7. **Dokładnie tak zachowują się formuły w openpyxl** i to jest treść sekcji 2.6.
- **`move_range("A1:C10", rows=2, cols=1, translate=True)` to przesunięcie działek razem z aktualizacją adresów w dokumentach.** Też urbanistycznie: to jedna z niewielu operacji, które „myślą". Wszystko inne to mechaniczne przepisanie.

### Analogia do iteracji: czytanie książki i skakanie po literach

Masz książkę i chcesz policzyć, ile razy występuje litera „a".

- **Sposób pierwszy: czytasz strona po stronie, linijka po linijce.** Oko przesuwa się po tekście raz, w naturalnym porządku. Szybko i wygodnie.
- **Sposób drugi: do każdej litery podchodzisz osobno, z indeksem.** Otwierasz książkę na stronie 1, linijce 1, znaku 1. Sprawdzasz. Potem strona 1, linijka 1, znak 2. Potem strona 1, linijka 1, znak 3... Za każdym razem otwierasz książkę od nowa, szukasz strony, potem linijki, potem znaku.

Oba sposoby dają ten sam wynik. Drugi jest **dziesiątki razy wolniejszy** — i to nie jest przesada. W sekcji 3 zmierzymy to na realnych danych i zobaczysz różnicę na własne oczy.

Excelowy odpowiednik pierwszego sposobu to:

```python
for wiersz in ws.iter_rows(values_only=True):
    for wartosc in wiersz:
        ...   # robisz coś z wartością
```

Excelowy odpowiednik drugiego to:

```python
for r in range(1, ws.max_row + 1):
    for c in range(1, ws.max_column + 1):
        wartosc = ws.cell(row=r, column=c).value   # <- otwieranie książki od nowa
```

**Konsekwencja praktyczna:** jeśli kiedykolwiek napiszesz pętlę, w której wywołujesz `ws.cell(...)` w środku zagnieżdżonej pętli po całym arkuszu — zatrzymaj się i zapytaj, czy nie da się tego zrobić przez `iter_rows`. Odpowiedź brzmi „prawie zawsze da się".

### Analogia do modyfikowania w trakcie iteracji: przesuwanie mebli, gdy ktoś rysuje plan

Wyobraź sobie, że rysujesz plan pokoju, przechodząc od ściany do ściany, a **w trakcie** ktoś przesuwa meble. Twoje notatki staną się bezsensowne: zapiszesz „szafa przy lewej ścianie", potem ktoś ją przesunie, a Ty za chwilę zapiszesz „szafa przy prawej ścianie" — i oba zapisy będą „prawdziwe" w swoim momencie, ale opis przestanie się zgadzać.

To samo dzieje się, gdy iterujesz po arkuszu (`for row in ws`) i **w tej samej pętli** wstawiasz albo usuwasz wiersze. Generator pod spodem przesuwa się po indeksach, a Ty zmieniasz te indeksy w trakcie. Efekt: pomijasz wiersze albo przetwarzasz je dwukrotnie. Rozwiązanie omówimy w sekcji 2.12 — w skrócie: **iteruj po kopii albo najpierw zbierz decyzje, a modyfikuj po zakończeniu pętli.**

### Analogia do adresu i nazwy: tablica rejestracyjna i numer telefonu

Ostatnia mikro-analogia, potrzebna w sekcji 2.8. Adres komórki w Excelu istnieje w dwóch postaciach i trzeba umieć je przeliczać:

- **`B7`** — „Aleja B, Ulica 7". **Ten adres widzi człowiek.** Nazwy kolumn literami, numer wiersza liczbami. To jest **klucz**, którym posługują się formuły, nazwy zdefiniowane i formatowanie warunkowe.
- **`(row=7, column=2)`** — „wiersz siódmy, druga kolumna od lewej". **Ten adres widzi pętla.** Liczby przy liczbach, wygodnie dla `range()`.

Dwie reprezentacje tego samego miejsca. **Prawie połowa błędów w automatyzacji Excela sprowadza się do pomylenia tych dwóch światów** — na przykład próby użycia `ws.cell(row=2, column=2)` przy mówieniu o „komórce B2", gdy masz na myśli drugi wiersz drugiej kolumny (to `B2`), albo `ws.cell(row="B", column=7)` (co nie ma sensu). Analogia: tablica rejestracyjna i numer VIN tego samego auta. Oba są poprawne, oba identyfikują to samo, ale nikt nie próbuje wkleić VIN-u na tablicę.

## 2. Teoria

### 2.1. Siatka i cztery różne pytania o rozmiar

Na arkusz można patrzeć jak na **prostokąt komórek** — od `A1` do jakiejś odległej granicy — i to jest pierwsze uproszczenie, którego openpyxl wymaga od Ciebie. Nie ma tu nieskończoności: `max_row` to maksymalnie 1 048 576, a `max_column` to maksymalnie 16 384 (`XFD`). Są to limity formatu `.xlsx`, nie openpyxl.

W obrębie tego prostokąta masz cztery różne pojęcia rozmiaru i każde odpowiada na inne pytanie:

**Pierwsze: `ws.max_row` i `ws.max_column` — „gdzie kończy się ewidencja".** Zwracają numer najdalszego wiersza i kolumny, w których openpyxl **ma jakąkolwiek komórkę**. Nie „ma daną", nie „ma wartość" — ma komórkę. Różnica jest ogromna i do niej wracamy w sekcji 2.2.

**Drugie: `ws.min_row` i `ws.min_column` — „gdzie zaczyna się ewidencja".** Zwykle `1`, ale w plikach generowanych przez narzędzia bywa inaczej. Jeśli Twój plik ma nagłówki w wierszu 3 (bo wiersze 1–2 to logo i tytuł), `min_row` będzie `1` — bo komórki z logo i tytułem też są w ewidencji.

**Trzecie: `ws.dimensions` i `ws.calculate_dimension()` — „jaki jest prostokąt obejmujący wszystko".** Zwracają **tekst** w notacji zakresu, np. `"A1:H50"`. To ten sam prostokąt, który widać w polu nazwy po wciśnięciu `Ctrl+End` w Excelu — bounding box, nie rozmiar danych.

W openpyxl 3.1.x `calculate_dimension()` przyjmuje parametr `force`. Służy on do tego, żeby wymusić policzenie wymiarów, gdy openpyxl nie ma ich jeszcze ustalonych (dotyczy to przede wszystkim trybu `write_only`). Składnię i zachowanie sprawdź w dokumentacji swojej wersji — to parametr dodany w gałęzi 3.1 i jego rola bywa mylona z „naprawianiem" wymiarów.

**Czwarte: „rozmiar logiczny danych" — pytanie, którego openpyxl Ci nie zada i nie odpowie.** To liczba wierszy i kolumn, w których **naprawdę są treści, które Cię interesują**. Nie ma na to API, bo openpyxl nie wie, co jest dla Ciebie „treścią". Nagłówek? Wiersz z sumą? Komentarz w kolumnie `H`? To Ty decydujesz. I dlatego tę liczbę liczy się własną funkcją — sekcja 2.11.

**Konsekwencja praktyczna.** Jeśli w kodzie widzisz `range(2, ws.max_row + 1)`, zadaj sobie pytanie: „czy `max_row` na pewno kończy się tam, gdzie kończą się dane?". W połowie przypadków odpowiedź brzmi „nie". A pętla przeleci po setkach pustych wierszy i skończy się gdzieś na 50 000, bo ktoś kiedyś kliknął wiersz i nacisnął `Ctrl+Shift+↓`.

### 2.2. Dlaczego `max_row` kłamie — mechanizm

Musisz zrozumieć **jedną rzecz strukturalną**, żeby to przestało być magią.

W pamięci arkusz trzyma słownik komórek, w którym kluczem jest para `(row, column)`. **Komórka trafia do tego słownika nie wtedy, gdy dostanie wartość, a wtedy, gdy openpyxl po raz pierwszy się do niej odwoła.** To jest ten mechanizm:

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

print(ws.max_row, ws.max_column)     # 1 1  <- pusty arkusz, jedno "puste" miejsce

ws["A1"] = "dane"
print(ws.max_row, ws.max_column)     # 1 1

ws["H40"]                             # <- nawet NIE przypisujemy wartości!
print(ws.max_row, ws.max_column)     # 40 8  <- ewidencja urosła od samego odwołania
```

**Co się dzieje w pamięci.** Wywołanie `ws["H40"]` przechodzi przez wewnętrzną metodę, która sprawdza, czy pod kluczem `(40, 8)` istnieje już obiekt komórki. Jeśli nie — **tworzy go i zapisuje**. Komórka nie ma wartości (`value is None`), ale istnieje i liczy się do wymiarów.

**Co trafi do pliku.** Komórka bez wartości i bez stylu może zostać pominięta przy zapisie, więc plik na dysku może być „mniejszy" niż model w pamięci. Ale to **nie naprawia Twojego kodu**: pętla po `max_row` nadal będzie chodziła do 40.

Ten sam efekt wywołują style:

```python
from openpyxl.styles import Font

# "Sformatujmy na zapas kolumnę A do wiersza 500"
for wiersz in range(2, 501):
    ws.cell(row=wiersz, column=1).font = Font(bold=False)
```

Po tym kodzie `ws.max_row` wynosi **500**, mimo że danych w arkuszu jest dziesięć wierszy. I tu jest ciekawostka: przypisaliśmy `Font(bold=False)` — czyli **czcionkę identyczną z domyślną**. Nic się wizualnie nie zmieniło. Ale komórki powstały i wymiary urosły.

**Skąd w praktyce biorą się zawyżone wymiary w gotowych plikach.** Cztery najczęstsze scenariusze:

1. **Ktoś w Excelu zaznaczył całą kolumnę i nadał jej format** (np. `Ctrl+1` → format walutowy). Excel zapisuje to jako informację o kolumnie, ale **nie** jako 1 048 576 komórek — więc efekt w `max_row` jest zwykle niewielki. Inaczej jest z formatowaniem **zakresów** — jeśli ktoś zaznaczył `A1:H100000` i kliknął „pogrubienie", do pliku trafiło 800 000 komórek bez wartości. `max_row` będzie równe 100 000.
2. **Ktoś usunął dane, ale nie usunął formatowania.** `Delete` kasuje zawartość komórki, ale nie styl. W pliku zostają komórki ze stylem i bez wartości.
3. **Ktoś kliknął `Ctrl+Shift+↓` i coś zrobił** — a Excel przy tym „dotknął" całego zakresu do końca danych.
4. **Narzędzie programowe sformatowało kolumnę na zapas** — dokładnie tak, jak w kodzie wyżej.

**Konsekwencja praktyczna — trzy zasady:**

- **Nie używaj `max_row` jako liczby wierszy danych.** Nigdy. Chyba że sam ten plik zapisałeś i wiesz dokładnie, co w nim jest.
- **Dla plików od ludzi zawsze licz rozmiar sam.** Własną funkcją. To kilka linijek (sekcja 2.11).
- **Gdy budujesz arkusz formatowaniem, rób to po wypełnieniu danymi, nie przed.** Najpierw dane, potem ostylowanie zakresu — wtedy formatowanie nie „puchnie" ponad dane. A jeśli musisz formatować „na zapas", wiesz gdzie spojrzeć, gdy `max_row` wygląda podejrzanie.

### 2.3. `read_only` i `write_only` — wymiary, którym nie można ufać

Oba tryby wydajnościowe (szczegóły — moduł 19) mają własne zasady dotyczące wymiarów.

**W trybie `read_only=True`** arkusz nie jest budowany w pamięci — jest czytany strumieniowo z pliku. Oznacza to dwie rzeczy:

- **Dostęp swobodny nie działa.** `ws["A1"]` jest niedostępne (w `ReadOnlyWorksheet` tej operacji po prostu nie ma). Masz do dyspozycji `iter_rows()`, `iter_cols()` i iterację po samym obiekcie arkusza. To fundamentalna zmiana sposobu myślenia: **skoro nie możesz skakać po komórkach, musisz myśleć siatką.** I to jest dokładnie to, do czego ten moduł Cię przekonuje.
- **Wymiary pochodzą z deklaracji zapisanej w pliku**, a nie z faktycznego skanu komórek. Ta deklaracja jest tworzona przez program, który zapisał plik — i **Excel często podaje ją luźno** (np. jako cały zakres od `A1` do końca sformatowanego obszaru). W dodatku dla części plików openpyxl nie umie jej ustalić i `max_row`/`max_column` mogą być `None` albo niewiarygodnie małe. Sprawdź w dokumentacji swojej wersji, jak zachowuje się dla Twoich plików — ale **praktyczna zasada jest jedna: w `read_only` wymiary traktuj jako podpowiedź, nie jako prawdę.**

**W trybie `write_only=True`** jest odwrotnie: dopóki nie zapiszesz, wymiarów „nie ma", bo arkusz jest generatorem strumieniowym. `ws.max_row` i `ws.max_column` nie są w użytecznej formie. Arkusz wie tylko, na którym wierszu jesteś (śledzi to wewnętrznie, żeby `append` wiedział, gdzie pisać dalej) — i właśnie do tej wewnętrznej księgowości służy metoda `ws.reset_dimensions()`.

**Uczciwie o `ws.reset_dimensions()`.** To metoda, która **zeruje wewnętrzne liczniki wymiarów** openpyxl. Przydaje się przede wszystkim przy pracy w trybie `write_only` i przy ponownym użyciu arkusza w nietypowych scenariuszach. **Nie usuwa komórek, które już istnieją, i nie naprawi zawyżonego `max_row`** w normalnym arkuszu — bo `max_row` liczony jest ze słownika komórek, a nie z licznika. Zachowanie tej metody sprawdź w dokumentacji swojej wersji; w tym kursie wspominamy ją, żebyś nie sięgnął po nią jako po remedium na problem `max_row`, bo nim nie jest.

**Konsekwencja praktyczna:** jeśli musisz znać prawdziwy rozmiar danych w pliku od użytkownika, **zrób to skanem** — sekcja 2.11. Skan działa niezależnie od trybu i nie zależy od tego, jak plik został zapisany.

### 2.4. Iteracja: trzy drogi i jedna zasada

Masz trzy podstawowe drogi poruszania się po siatce. Wszystkie zwracają **generatory** — czyli leniwe iteratory, które wyliczają kolejne elementy dopiero wtedy, gdy o nie poprosisz. (Analogia: generator to taśma produkcyjna, która wyrzuca kolejne elementy na bieżąco, a nie magazyn, w którym wszystko czeka gotowe).

**Droga pierwsza: `for wiersz in ws`.** Iteracja po samym obiekcie arkusza. Każdy obrót pętli daje **krotkę komórek** z danego wiersza.

```python
for wiersz in ws:
    for komorka in wiersz:
        print(komorka.coordinate, komorka.value)
```

**Droga druga: `ws.iter_rows()`.** To samo, ale z pełnym sterowaniem:

```python
for wiersz in ws.iter_rows(min_row=2, max_row=10, min_col=1, max_col=4):
    for komorka in wiersz:
        print(komorka.coordinate, komorka.value)
```

Parametry `min_row`, `max_row`, `min_col`, `max_col` wyznaczają **prostokąt okna**, po którym chodzisz. To bardzo praktyczne: nagłówki pomijasz przez `min_row=2`, a tabelę ograniczasz do czterech kolumn przez `max_col=4`.

**Droga trzecia: `ws.iter_cols()`.** Iteracja kolumnami — każdy obrót pętli daje krotkę komórek z danej kolumny.

```python
for kolumna in ws.iter_cols(min_col=2, max_col=2):
    for komorka in kolumna:
        print(komorka.coordinate, komorka.value)
```

**Obiekty, które dostajesz — to najważniejsze rozróżnienie tego podrozdziału:**

| Wywołanie | Co dostajesz | Jak sięgnąć po wartość |
|---|---|---|
| `for row in ws` | krotka obiektów `Cell` | `komorka.value` |
| `ws.iter_rows()` | krotka obiektów `Cell` | `komorka.value` |
| `ws.iter_rows(values_only=True)` | **krotka wartości** (`str`, `int`, `float`, `datetime`, `None`) | `wartosc` bezpośrednio |
| `ws.iter_cols()` | krotka obiektów `Cell` | `komorka.value` |
| `ws.iter_cols(values_only=True)` | **krotka wartości** | `wartosc` bezpośrednio |
| `ws.values` | krotka wartości, po wierszach | jak wyżej |
| `ws.rows` / `ws.columns` | generatory krotek komórek | `komorka.value` |

**`values_only=True` to najczęstszy wybór, gdy robisz skrypt analityczny.** Mniej obiektów do utworzenia, mniej pamięci, mniej atrybutów do napisania. Ale tracisz `komorka.coordinate` i `.row`/`.column` — jeśli potrzebujesz adresu, musisz liczyć sam (np. przez `enumerate`).

```python
for indeks, wiersz in enumerate(ws.iter_rows(values_only=True), start=1):
    print(indeks, wiersz)      # numer wiersza + krotka wartości
```

**Jedna zasada, która obejmuje wszystkie trzy drogi:**

> **Iteruj wierszami, nie komórkami. Wiersz jest jednostką pracy; komórka jest szczegółem.**

Dlaczego? Bo wiersz odpowiada jednemu „rekordowi", jednej osobie, jednemu zamówieniu, jednemu dniu. To właśnie na tym poziomie myślisz o danych. A technicznie: iteracja wierszami w openpyxl jest zoptymalizowana, bo czyta komórki w porządku, w jakim są w pliku; skakanie po komórkach wymusza wyszukiwanie w słowniku i wielokrotne tworzenie obiektów.

Zmierzymy to w sekcji 3 (Przykład 2) — różnica bywa rzędu kilku–kilkudziesięciu razy i rośnie z rozmiarem danych.

**Uwaga o generatorach:** zwracany obiekt **można przejść tylko raz**. Jeśli chcesz mieć dane do dyspozycji wielokrotnie, zrób `list(ws.values)` albo zapisz krotki w zwykłej liście. Świeżo upieczona analogia: generator to odczyt jednorazowy, nie plik na dysku.

### 2.5. `append` — najszybszy sposób dopisywania i najczęstsza literówka

`ws.append(...)` to metoda, która dokłada wiersz na końcu aktualnie zapisanej zawartości. To najszybszy sposób budowania arkusza i pierwsze narzędzie przy generowaniu raportów.

```python
ws.append(["Lp.", "Nazwa", "Ilość"])     # trzy komórki w jednym wierszu
ws.append([1, "Widzet", 12])
ws.append([2, "Wkręt", 340])
```

**Co się dzieje w pamięci.** `append` bierze **iterowalne** (coś, po czym można przejść pętlą — lista, krotka, `range`, generator, a nawet napis) i przypisuje kolejnym komórkom nowego wiersza kolejne elementy: pierwszy element → kolumna 1, drugi → kolumna 2, i tak dalej. Wiersz, do którego pisze, to „aktualny wiersz + 1" — a nie `max_row + 1` (to rozróżnienie ma znaczenie w sekcji 2.6 przy operacjach edycyjnych).

**I tu jest pułapka, która kosztowała tysiące ludzi tysiące godzin:**

```python
ws.append("ABC")        # <- NIE zapisze napisu "ABC" w jednej komórce!
```

Napis jest iterowalny — jego elementami są znaki `A`, `B`, `C`. Więc `append` napisze trzy komórki: `A`, `B`, `C`. Jeśli liczyłeś na jedną komórkę z tekstem `ABC`, dostałeś trzy osobne.

Jeszcze groźniejszy wariant, bo całkowicie cichy:

```python
nazwa = "Firma Kowalski sp. z o.o."
ws.append([numer, nazwa, kwota])       # ✅ dobrze - napis jest ELEMENTEM listy
ws.append(numer, nazwa, kwota)         # ❌ TypeError - append bierze JEDEN argument
ws.append(nazwa)                       # ❌ 24 komórki - każda litera osobno!
```

**Analogia: `append` to taca, a nie walizka.** Podajesz jedną tacę (`listę`), a na niej leży dowolna liczba rzeczy. Jeśli podasz **jedną rzecz** (napis), to openpyxl rozkłada ją na części składowe i próbuje położyć każdą z osobna.

**Trzy dodatkowe zachowania `append`, o których warto wiedzieć:**

1. **Przyjmuje dowolne iterowalne, nie tylko listy.** `ws.append(range(3))` zadziała. `ws.append(generator_logów())` też — i to jest bardzo wygodne przy strumieniowaniu danych z bazy (moduł 19).
2. **Przyjmuje też słownik.** W tym wariancie klucze wskazują kolumny, do których trafiają wartości:
   ```python
   ws.append({"B": "Widzet", "D": 12, "F": 47.90})   # pisze tylko do B, D, F
   ```
   To rzadko używana, a bardzo praktyczna możliwość przy „dziurawych" wierszach. Sprawdź dokładnie zachowanie wariantu ze słownikiem w dokumentacji swojej wersji — klucze mogą być literami kolumn albo numerami, zależnie od wydania.
3. **`append` startuje od kolumny 1** (albo, w trybie `write_only`, od kolumny bieżącej), a „aktualny wiersz" przesuwa się o jeden po każdym wywołaniu. Nie ma opcji „dopisz od kolumny 5" — jeśli to potrzebujesz, użyj `ws.cell(...)`.

**Konsekwencja praktyczna:** `append` jest nieodłącznym narzędziem generowania raportów. Ale **każde miejsce w kodzie, w którym podajesz napis bezpośrednio do `append` (a nie jako element listy), to błąd**. Zbuduj nawyk, żeby `append` zawsze dostawał literał listy:

```python
ws.append([...])       # tak, zawsze
```

Jeśli dane pochodzą z pętli:

```python
for rekord in rekordy:
    ws.append([rekord.id, rekord.nazwa, rekord.ilosc, rekord.cena])   # lista zawsze
```

To jedna reguła, która całkowicie usuwa tę klasę błędów z Twojego kodu.

### 2.6. `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols` — co się przesuwa, a co nie

Te cztery metody zmieniają geometrię arkusza. Wywołanie wygląda tak:

```python
ws.insert_rows(idx=5, amount=3)     # wstaw 3 puste wiersze przed wierszem 5
ws.delete_rows(idx=5, amount=2)     # usuń 2 wiersze, począwszy od wiersza 5
ws.insert_cols(idx=2, amount=1)     # wstaw jedną kolumnę przed kolumną B
ws.delete_cols(idx=2, amount=1)     # usuń kolumnę B
```

**`idx` jest 1-based.** Pierwszy wiersz to `1`, pierwsza kolumna to `1`. `idx=0` jest błędem (a nie „wstaw na początku"). To ta sama konwencja, co `ws.cell(row=1, column=1)`.

`amount` domyślnie wynosi `1`. Można pominąć: `ws.insert_rows(5)` wstawia jeden wiersz.

**A teraz najważniejsza część tego podrozdziału: co się faktycznie przesuwa.**

| Element arkusza | `insert_rows` / `delete_rows` | `insert_cols` / `delete_cols` |
|---|---|---|
| Wartości w komórkach | ✅ przesuwają się | ✅ przesuwają się |
| Style komórek | ✅ przesuwają się razem z komórkami (bo style są **właściwością komórki**) | ✅ tak samo |
| **Referencje w formułach** | ❌ **NIE** | ❌ **NIE** |
| Wysokości / szerokości wierszy i kolumn | ❌ nie przesuwają się | ❌ nie przesuwają się |
| Scalanie komórek (`merged_cells`) | ⚠️ zachowanie historycznie się zmieniało — **sprawdź w swojej wersji i przetestuj** | ⚠️ jw. |
| Formatowanie warunkowe (`conditional_formatting`) | ❌ nie przesuwa się | ❌ nie przesuwa się |
| Walidacja danych (`data_validations`) | ❌ nie przesuwa się | ❌ nie przesuwa się |
| Tabele (`tables`) | ⚠️ zakres tabeli może wymagać ręcznej korekty | ⚠️ jw. |
| Obrazy, wykresy (anchory) | ❌ nie przesuwają się | ❌ nie przesuwają się |

Uwagę trzeba poświęcić zwłaszcza jednemu wierszowi: **formuły się nie aktualizują**.

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.append(["A", "B", "Suma"])
ws.append([10, 20, "=SUM(A2:B2)"])         # formuła odwołuje się do A2:B2
ws.append([30, 40, "=SUM(A3:B3)"])

print(ws["C2"].value)                      # '=SUM(A2:B2)'

ws.insert_rows(1, amount=1)                # wstawiamy wiersz na samej górze

print(ws["C3"].value)                      # '=SUM(A2:B2)'  <- NIE '=SUM(A3:B3)'!
```

**Co się dzieje w pamięci.** Wartości i komórki przesunęły się o jeden wiersz w dół. Ale `C2` „fizycznie" zmieniło się w `C3` i jego wartość jest **nadal** tym samym napisem formuły `'=SUM(A2:B2)'`. openpyxl **nie ma grafu zależności** — nie wie, że formuła odnosi się do komórek, ani nie próbuje tego naprawiać. To ta sama historia, którą zobaczyłeś w module 06 przy kopiowaniu komórek: openpyxl przepisuje tekst dosłownie, a „inteligencja" referencji jest Twoją odpowiedzialnością.

**Analogia urbanistyczna z sekcji 1:** wcisnąłeś nową ulicę, ale w aktach notarialnych nadal jest napisane „Ulica 6", choć ta działka jest teraz przy Ulicy 7.

**Trzy konsekwencje praktyczne:**

1. **`insert_rows` i `delete_rows` to operacje ryzykowne w plikach, których nie napisałeś.** Jeśli plik zawiera formuły, wstawianie wiersza przez openpyxl je zepsuje — Excel pokaże `#REF!` albo (gorzej) **poprawne z wyglądu, ale błędne liczbowe wyniki** (formuła odwołuje się do A2:B2, które po przesunięciu są inną parą liczb).
2. **Nie ma API do „wstaw wiersz i jednocześnie przesuń formuły".** Musisz albo samodzielnie przemieścić treści formuł przez `Translator` (moduł 06), albo — co jest prostsze i bezpieczniejsze — **wstawić wiersz razem z pustymi formułami, które od razu wypełnisz poprawnymi referencjami**. Jeśli planujesz edycję plików z bogatymi formułami, wróć do reguły z modułu 18: nie przepuszczaj ich przez openpyxl.
3. **Formatowanie warunkowe i walidacja danych nie przesuwają się „razem z danymi".** Jeśli nad regionem `A2:A100` miałeś regułę warunkową i wstawiłeś trzy wiersze nagłówka, Twoja reguła nadal obejmuje dokładnie te same komórki — które są teraz w innym miejscu. Musisz ją przenieść (i usunąć starą) ręcznie.

**Jeszcze jedna rzecz, bardzo praktyczna**: `delete_rows` nie czyni cudów w danych z odwołaniami w komórkach w innych arkuszach. Formuła w innym arkuszu, wskazująca na `Dane!A5`, po usunięciu wiersza 5 nie „przeskoczy" na następcę. Po prostu wskaże coś, co już nie jest tym, co myślisz.

### 2.7. `move_range` — jedna operacja, która „myśli"

To najciekawsza i najczęściej pomijana metoda tego modułu. Służy do przesuwania zakresu komórek — i przy odpowiednim parametrze **tłumaczy też referencje w formułach**.

```python
ws.move_range("A1:C10", rows=2, cols=1, translate=True)
```

Parametry:

- `cell_range` — tekst w notacji zakresu (`"A1:C10"`) albo obiekt `CellRange`.
- `rows` — o ile wierszy przesunąć (dodatnie = w dół, ujemne = w górę).
- `cols` — o ile kolumn przesunąć (dodatnie = w prawo, ujemne = w lewo).
- `translate` — czy **tłumaczyć referencje w formułach** podczas przesuwania. Domyślnie `False`.

**Różnica w zachowaniu przy formułach jest fundamentalna:**

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

# Wiersz 1: nagłówki, wiersz 2: dane z formułą, wiersz 3: druga formuła
ws.append(["Ilość", "Cena", "Wartość"])
ws.append([10, 5.0, "=A2*B2"])
ws.append([20, 7.0, "=A3*B3"])

ws.move_range("A1:C3", rows=3, cols=0, translate=False)     # bez tłumaczenia
print("translate=False ->", ws["C5"].value)                 # '=A2*B2'  (stara referencja!)

wb2 = Workbook()
ws2 = wb2.active
ws2.append(["Ilość", "Cena", "Wartość"])
ws2.append([10, 5.0, "=A2*B2"])
ws2.append([20, 7.0, "=A3*B3"])

ws2.move_range("A1:C3", rows=3, cols=0, translate=True)      # z tłumaczeniem
print("translate=True  ->", ws2["C5"].value)                # '=A5*B5'  (referencje przesunięte!)
```

W pierwszym przypadku formuła przeniosła się razem z komórką, ale jej wnętrze nie zostało zmienione — będzie wskazywać na pusty `A2:B2`, czyli da wynik `0`. W drugim — openpyxl „zrozumiał", że komórka przeniosła się o trzy wiersze w dół, i przesunął referencję. Wynik: formuła nadal działa poprawnie.

**Dlaczego to działa, a `insert_rows` nie?** Bo `move_range` używa `Translator` (moduł 06) — tokenizera formuł, który rozumie gramatykę i wie, które fragmenty to referencje. `insert_rows` z kolei po prostu przepisuje komórki w słowniku na inne klucze, nie próbując zrozumieć ich treści.

**Kiedy `move_range` z `translate=True` jest Twoim przyjacielem:**

- **Dodajesz wiersz nagłówka nad istniejącą tabelą.** `ws.move_range("A1:D50", rows=1, translate=True)`, potem wpisujesz nagłówki w wiersz 1. Formuły w tabeli zostają logicznie poprawne.
- **Przestawiasz sekcje raportu.** Przenosisz blok podsumowania w inne miejsce i nie chcesz zaktualizować kilkudziesięciu formuł ręcznie.
- **Przenosisz fragment szablonu do nowego pliku** (przy współpracy z `copy_worksheet` — moduł 08).

**Kiedy `translate=True` jest pułapką — trzy przypadki:**

1. **Przesuwasz zakres, który ma odwołania do komórek POZA nim.** Formuła, która wskazywała na stałą konfigurację w `Z1`, dostanie po przesunięciu referencję `Z4` — a to już nie jest Twoja konfiguracja, tylko pusta komórka. `move_range` tłumaczy **wszystkie** referencje tak samo, nie tylko te wewnątrz zakresu.
2. **Przesuwasz zakres w pliku, w którym inne arkusze odwołują się do tych komórek.** Arkusz `Podsumowanie` ma formułę `=SUM(Dane!B2:B10)`. Po `move_range` w arkuszu `Dane` odwołanie w `Podsumowanie` się nie zmieni — bo openpyxl nie śledzi powiązań między arkuszami. Zostaniesz z odwołaniem do nieaktualnego zakresu.
3. **Przesuwasz zakres z formułami, które są już „przesunięte celowo".** Jeśli jakaś formuła była celowo napisana tak, żeby odwoływać się do sztywnej lokalizacji, `translate=True` to „naprawi" — czyli zepsuje.

**Konsekwencja praktyczna:** `move_range(translate=True)` jest jednym z niewielu miejsc w openpyxl, gdzie biblioteka robi coś „inteligentnego". Właśnie dlatego trzeba go używać świadomie, w wąskim zakresie, i **sprawdzić wynik okiem w Excelu**, zanim wyślesz plik dalej. Jest to w tym samym stopniu narzędzie, co broń palna.

### 2.8. Adresy i konwersje — most między światem liter i światem liczb

W openpyxl istnieje zestaw funkcji narzędziowych, które przeliczają między notacją `B7` a parą `(row, column)`, i między zakresem a jego granicami. Nie jest to ozdobnik — to jest **fundament programowego budowania odwołań**. Bez tego piszesz `"F201"` na sztywno i płaczesz, gdy zmieni się liczba wierszy danych.

**Funkcje najważniejsze:**

```python
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple

get_column_letter(1)          # 'A'
get_column_letter(26)         # 'Z'
get_column_letter(27)         # 'AA'
get_column_letter(28)         # 'AB'
get_column_letter(16384)      # 'XFD'  <- ostatnia kolumna w .xlsx

column_index_from_string("A")     # 1
column_index_from_string("Z")     # 26
column_index_from_string("AB")    # 28
column_index_from_string("$AB")   # 28  <- akceptuje też z '$'

coordinate_from_string("AB12")    # ('AB', 12)   <- (litera, numer_wiersza)
coordinate_from_string("$AB$12")  # ('AB', 12)   <- znaki '$' są pomijane

coordinate_to_tuple("AB12")       # (12, 28)     <- (wiersz, kolumna)
coordinate_to_tuple("AB12")[0]    # 12  <- wiersz
coordinate_to_tuple("AB12")[1]    # 28  <- kolumna
```

**Zwróć uwagę na różnicę zwrotów dwóch ostatnich funkcji:**

- `coordinate_from_string("AB12")` zwraca `("AB", 12)` — **litera pierwsz, liczba druga**.
- `coordinate_to_tuple("AB12")` zwraca `(12, 28)` — **wiersz pierwsz, kolumna druga** (kolejność `(row, column)`, jak w `ws.cell(...)`).

To dwie różne, ale uzupełniające się reprezentacje. Pierwsza jest „ludzka" (jesteś przyzwyczajony do zapisu `B7` — kolumna potem wiersz). Druga jest „maszynowa" (zgodna z `ws.cell(row=..., column=...)`). **Pomylenie ich to jedna z najczęstszych literówek w kodzie.**

Zapamiętaj to przez skojarzenie:

- `coordinate_from_string` → coś **z napisu**, więc dostajesz to, co widać w napisie: literę i liczbę.
- `coordinate_to_tuple` → konwersja **do argumentów `cell()`**, więc kolejność jak w `cell(row=..., column=...)`.

**Zakresy — trzy użyteczne funkcje:**

```python
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.cell import range_to_tuple
from openpyxl.worksheet.cell_range import CellRange

range_boundaries("A1:C3")             # (1, 1, 3, 3)  <- (min_col, min_row, max_col, max_row)
range_boundaries("B2:D10")            # (2, 2, 4, 10)
range_boundaries("A1")                # (1, 1, 1, 1)  <- pojedyncza komórka też jest zakresem

range_to_tuple("'Dane 2026'!A1:C3")   # ('Dane 2026', (1, 1, 3, 3))
range_to_tuple("Arkusz1!B2:D10")      # ('Arkusz1', (2, 2, 4, 10))

CellRange("A1:C3").min_row          # 1
CellRange("A1:C3").max_row          # 3
CellRange("A1:C3").min_col          # 1
CellRange("A1:C3").max_col          # 3
CellRange("A1:C3").size             # {'width': 3, 'height': 3}
```

**Kolejność zwrotu `range_boundaries` to `(min_col, min_row, max_col, max_row)`** — kolumny są pierwsze. To wygodne przy pracy z zakresami, ale **niekompatybilne kolejnością** z `CellRow` i z `CellRange.size`. Uwaga — to jedna z tych „drobnych" niekonsekwencji, które sprawiają, że literały trzeba sprawdzać.

`CellRange` to obiekt, który dodatkowo pozwala **iterować po współrzędnych w zakresie**:

```python
from openpyxl.utils import rows_from_range, cols_from_range

# Iteracja po współrzędnych, wierszami:
for wiersz in rows_from_range("A1:C3"):
    print(wiersz)
# ('A1','B1','C1'), ('A2','B2','C2'), ('A3','B3','C3')

# Iteracja kolumnami:
for kolumna in cols_from_range("A1:C3"):
    print(kolumna)
# ('A1','A2','A3'), ('B1','B2','B3'), ('C1','C2','C3')
```

To zwraca **napisy adresów**, nie komórki. Użyteczne, gdy potrzebujesz wygenerować odwołania do formuł albo nazw zdefiniowanych.

### 2.9. Bezpieczne budowanie odwołań — `quote_sheetname` i `absolute_coordinate`

Gdy tworzysz formułę albo nazwę zdefiniowaną i chcesz zbudować odwołanie do zakresu z innego arkusza, **nigdy nie pisz tego ręcznie**:

```python
# ❌ Ręcznie - skazane na błąd, gdy nazwa ma spację albo apostrof
formula = "=SUM('Dane 2026'!$B$2:$B$201)"

# ✅ Programowo - działa dla każdej nazwy
from openpyxl.utils import absolute_coordinate, quote_sheetname

nazwa_arkusza = "Dane 2026"
adres = f"{quote_sheetname(nazwa_arkusza)}!{absolute_coordinate('B2:B201')}"
formula = f"=SUM({adres})"
print(formula)     # =SUM('Dane 2026'!$B$2:$B$201)
```

**Co robi `quote_sheetname`:** jeśli nazwa arkusza zawiera spację, myślnik, kropkę, apostrof albo inne znaki niebezpieczne dla gramatyki formuły, otacza ją apostrofami. Jeśli nie zawiera — zwraca bez zmian. W praktyce `quote_sheetname("Dane")` zwróci `"Dane"`, a `quote_sheetname("Dane 2026")` zwróci `"'Dane 2026'"`.

**Co robi `absolute_coordinate`:** dopisuje `$` przy każdej literze i liczbie w zakresie, czyniąc go absolutnym. `absolute_coordinate("A1:B5")` → `"$A$1:$B$5"`. (Przyjmuje też pojedynczą komórkę: `absolute_coordinate("B2")` → `"$B$2"`).

**Dlaczego to ma sens nawet, jeśli nie potrzebujesz „absolutności".** Bo w formułach z odwołaniem do konkretnej lokalizacji zakres powinien być zapisany w postaci absolutnej. Tak zapisuje go Excel, tak rozumie go użytkownik, tak zrobiłby to człowiek. Zapisywanie zakresu w formie względnej bywa niepoprawne w kontekście, gdy formuła zostanie gdzieś skopiowana.

**Przykład całkowicie programowego budowania odwołania — bez ani jednej zakodowanej litery:**

```python
from openpyxl.utils import absolute_coordinate, get_column_letter, quote_sheetname

def odniesienie_do_kolumny(nazwa_arkusza: str, indeks_kolumny: int,
                           pierwszy_wiersz: int, ostatni_wiersz: int) -> str:
    """Zwraca np. 'Dane'!$F$2:$F$201. Bez żadnej zakodowanej litery."""
    litera = get_column_letter(indeks_kolumny)
    zakres = f"{litera}{pierwszy_wiersz}:{litera}{ostatni_wiersz}"
    return f"{quote_sheetname(nazwa_arkusza)}!{absolute_coordinate(zakres)}"

odniesienie_do_kolumny("Dane 2026", 6, 2, 201)
# "'Dane 2026'!$F$2:$F$201"
```

To jest wzorzec, który w module 06 już budowałeś f-stringiem. W tym module dodajemy do niego pełne wsparcie dla „nazwy arkusza z dowolnymi znakami" — i to jest **jedyne** poprawne podejście, jeśli nazwy arkuszy lub numery kolumn pochodzą z konfiguracji albo od użytkownika.

### 2.10. `ws.rows` i `ws.columns` — pułapka generatorów

Dwa wywołania, które wyglądają jak listy, a nimi nie są:

```python
print(type(ws.rows))     # <class 'generator'>
print(type(ws.columns))  # <class 'generator'>
```

**Co się dzieje w pamięci.** Każdy dostęp do `ws.rows` **zwraca nowy generator**. Generator jest **jednorazowy** — po przejściu do końca jest zużyty. Jeśli potrzebujesz do nich wrócić, zapisz wynik do listy.

```python
wiersze = list(ws.rows)        # teraz mam krotke krotek komorek
for wiersz in wiersze:
    print([k.value for k in wiersz])
for wiersz in wiersze:         # <- działa drugi raz, bo to lista
    ...
```

**`ws.columns` w dużych arkuszach to koszt.** Iteracja po kolumnach musi „przeskoczyć" kolejne wiersze i zebrać komórki z tej samej kolumny — co w praktyce oznacza przejście po całym arkuszu po raz kolejny. Jeśli wiesz, że chcesz kolumnę `F`, użyj:

```python
for komorka in ws["F"]:              # ws["F"] daje krotke komorek z kolumny F
    print(komorka.coordinate, komorka.value)

for komorka in ws["F"][1:]:          # pomijamy naglowek (F1)
    ...
```

`ws["F"]` zwraca krotkę komórek z kolumny. `ws["F"][1:]` to ta sama krotka bez pierwszego elementu (indeksy **od zera** w krotce!). To dobra okazja, by zapamiętać jedną rzecz: **w Pythonie krotka `ws["F"]` jest indeksowana od zera w wyniku, choć wiersze w Excelu liczą się od jedynki.** Czyli `ws["F"][0]` to `F1`, `ws["F"][9]` to `F10`.

**Analogia do generatora — automat z kawą, a nie ekspres do kawy.** Generator „wyrzuca" kolejne elementy, ale ich nie przechowuje. Eskpres w kuchni (lista) trzyma cały zapas — możesz o każdej porze wziąć kolejną filiżankę. To wybór między „jednorazowo i tanio" a „wielokrotnie i zajmuje miejsce". Wybieraj świadomie.

### 2.11. Znalezienie prawdziwego rozmiaru danych

Skoro `max_row` nie mówi, gdzie kończą się dane, trzeba to policzyć samemu. Oto funkcja bazowa:

```python
def znajdz_rozmiar_danych(ws) -> tuple[int, int]:
    """Zwraca (ostatni_wiersz, ostatnia_kolumna) dla komorek z WARTOŚCIĄ.

    Ignoruje komorki istniejące wyłącznie z powodu formatowania.
    """
    ostatni_wiersz = 0
    ostatnia_kolumna = 0

    for wiersz in ws.iter_rows():
        for komorka in wiersz:
            if komorka.value is None:
                continue
            if komorka.row > ostatni_wiersz:
                ostatni_wiersz = komorka.row
            if komorka.column > ostatnia_kolumna:
                ostatnia_kolumna = komorka.column

    return ostatni_wiersz, ostatnia_kolumna
```

**Co się dzieje w pamięci.** Iterujemy po wszystkich komórkach, o których wie openpyxl (czyli do `max_row` i `max_column`), i patrzymy na **wartość**. Łatwo to napisać, gorzej — tanio. Jeśli plik ma dane w 20 wierszach, ale sformatowane 100 000 komórek, przechodzimy przez wszystkie 100 000 komórek, żeby stwierdzić, że ostatnia wartość jest w wierszu 20.

**Wersja szybsza — skanowanie wstecz z przerwaniem:**

```python
def ostatni_wiersz_z_danymi(ws) -> int:
    """Zwraca numer ostatniego wiersza, w ktorym jest JAKAKOLWIEK wartosc.

    Skanuje wstecz od max_row i przerwa na pierwszym trafieniu.
    """
    for wiersz in range(ws.max_row, 0, -1):
        for krotka in ws.iter_rows(min_row=wiersz, max_row=wiersz):
            if any(komorka.value is not None for komorka in krotka):
                return wiersz
    return 0


def ostatnia_kolumna_z_danymi(ws) -> int:
    """Analogicznie dla kolumn - od max_column w dol."""
    for kolumna in range(ws.max_column, 0, -1):
        for krotka in ws.iter_cols(min_col=kolumna, max_col=kolumna):
            if any(komorka.value is not None for komorka in krotka):
                return kolumna
    return 0
```

Zaletą tej wersji jest **wczesne zatrzymanie**. Jeśli dane kończą się na wierszu 20 a `max_row` jest 100 000, pętla wykona 50 000 obrotów i zatrzyma się na 20... co nie brzmi jak wielkie zwycięstwo. Ale przy typowym pliku (dane na 99% zakresu) przerwie po kilku krokach. Jeśli chcesz prawdziwej gwarancji, musisz zapłacić: pełny skan albo heurystyka.

**Heurystyka — „szukaj ostatniego niepustego wiersza w kolumnie kluczowej".** Jeśli wiesz, że każdy wiersz danych ma niepuste pole w kolumnie np. 1 albo 3, sprawdzenie tylko tej kolumny da odpowiedź natychmiast:

```python
def ostatni_wiersz_w_kolumnie(ws, indeks_kolumny: int) -> int:
    """Zaklada, ze kolumna kluczowa nigdy nie jest pusta w wierszu z danymi."""
    ostatni = 0
    for komorka in ws.iter_cols(min_col=indeks_kolumny, max_col=indeks_kolumny):
        for k in komorka:
            if k.value is not None:
                ostatni = k.row
    return ostatni
```

To działa, gdy wiesz, że klucz nigdy nie bywa pusty. W praktyce biznesowej to najszybsze i najczęściej słuszne założenie. Ale trzeba je **wyrazić w kodzie** i mieć świadomość, że jest założeniem.

**Rozstrzygnięcie: która metoda jest właściwa?** Zależy od tego, czy Twoje dane mogą mieć „dziury". Trzy przypadki:

1. **Dane są ciągłe (wiersze od 1 do N z rzędu).** Ostatni niepusty wiersz w kolumnie kluczowej = rozmiar danych.
2. **Dane mogą mieć puste wiersze w środku** (np. arkusze „ręcznie robione" z oddechem między sekcjami). Wtedy ostatni niepusty wiersz wciąż daje poprawne górne ograniczenie, ale musisz wiedzieć, że wiersze pomiędzy mogą być puste.
3. **Mieszają się różne tabele na jednym arkuszu.** Wtedy żadna automatyczna metoda nie pomoże — potrzebujesz albo konwencji (jedna tabela na arkusz, moduł 08), albo deklaracji w konfiguracji (moduł 26). To architektoniczna pułapka, nie programistyczna.

### 2.12. Modyfikowanie arkusza w trakcie iteracji

Na koniec najsubtelniejsza pułapka tego modułu, która wygląda jak coś, co powinno działać.

```python
# ❌ TAK NIE ROBIĆ
for wiersz in ws.iter_rows(min_row=2):
    if wiersz[2].value is None:                # jeśli brak wartości w kolumnie C
        ws.delete_rows(wiersz[0].row)          # <- modyfikujemy w trakcie iteracji!
```

**Co się dzieje w pamięci.** Generator `iter_rows` „pamięta", w którym miejscu siatki jest — ale nie „wie", że zmieniłeś geometrię arkusza. Po `delete_rows` komórki poniżej przesunęły się o jeden wiersz w górę, a generator i tak przejdzie do kolejnej pozycji **w starej numeracji**. Efekt: **pomijasz wiersze** (bo po skasowaniu jeden „przeskoczył" na pozycję, którą generator uzna już za przetworzoną).

Analogia z sekcji 1: rysujesz plan pokoju, przechodząc od ściany do ściany, a w trakcie ktoś przesuwa meble.

Dokładnie ten sam problem ma Python w **każdej** kolekcji modyfikowanej w trakcie iteracji (słynne „nie zmieniaj listy, po której iterujesz"). To nie jest wina openpyxl — to właściwość każdej iteracji po strukturze, którą zmieniasz.

**Rozwiązanie A: zbierz decyzje, potem działaj.**

```python
# ✅ Faza 1: przegląd bez modyfikacji
do_usuniecia = []
for wiersz in ws.iter_rows(min_row=2):
    if wiersz[2].value is None:
        do_usuniecia.append(wiersz[0].row)

# ✅ Faza 2: modyfikacja od dołu w górę
for numer_wiersza in sorted(do_usuniecia, reverse=True):
    ws.delete_rows(numer_wiersza)
```

**Kluczowy szczegół:** iterujemy **od dołu w górę** (`reverse=True`), bo usunięcie wiersza 5 przesuwa wszystko poniżej. Jeśli usuwałbyś od góry, kolejne numery w Twojej liście przestałyby być poprawne po pierwszym usunięciu.

**Rozwiązanie B: iteruj po kopii danych.**

```python
dane = list(ws.iter_rows(values_only=True))      # kopia w pamieci
for indeks, wiersz in enumerate(dane, start=2):
    if wiersz[2] is None:
        print(f"do obrobki wiersz {indeks}")
```

Wariant B jest bezpieczny, ale kosztowny — trzymasz całość w pamięci. Dla plików dużych używaj wariantu A.

**Rozwiązanie C: nie modyfikuj geometrii arkusza, tylko buduj nowy.**

To wbrew intuicji najczęstsze podejście w skryptach produkcyjnych: zamiast usuwać wiersze z istniejącego arkusza, **czytasz dane, filtrujesz je w Pythonie i zapisujesz nowy arkusz**. Zalet mnóstwo: nie ma problemu z indeksami, wszystkie formuły i style są budowane raz, a wynikowy arkusz jest deterministyczny. To też przygotowuje grunt pod architekturę z modułu 26 („raport jako projekcja, nie edycja").

## 3. Przykłady krok po kroku

### Przykład 1 — wymiary kłamią: cztery liczby, trzy różne historie

```python
"""Wymiary arkusza: max_row rosnie od formatowania i od samego dotkniecia komorki.

Uruchom:  python examples/07_wymiary.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

CEL = OUTPUT / "07_wymiary.xlsx"

# Dane: naglowek + 11 wierszy, kolumny A-D.
LICZBA_WIERSZY_DANYCH = 12          # naglowek + 11 rekordow
LICZBA_KOLUMN_DANYCH = 4

# Formatowanie "na zapas": do wiersza 50, kolumny H.
FORMATOWANY_WIERSZ_KONCOWY = 50
FORMATOWANA_KOLUMNA_KONCOWA = 8


def zbuduj_plik(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Wymiary"

    ws.append(["Lp.", "Nazwa", "Ilość", "Cena"])
    for i in range(1, 12):
        ws.append([i, f"Pozycja {i}", i * 3, round(i * 1.37, 2)])

    # Formatowanie komorek, ktore nie maja wartosci - symulacja "Ctrl+B na zapas".
    # Wystarczy przypisanie stylu, zeby komorka powstawala w modelu.
    jasny = PatternFill(fill_type="solid", fgColor="FFF9E6")
    for wiersz in range(LICZBA_WIERSZY_DANYCH + 1, FORMATOWANY_WIERSZ_KONCOWY + 1):
        for kolumna in range(1, FORMATOWANA_KOLUMNA_KONCOWA + 1):
            ws.cell(row=wiersz, column=kolumna).fill = jasny

    # Na koncu "dotkniecie" jeszcze dalszej komorki bez wartosci - tylko odczyt.
    _ = ws["Z100"]

    wb.save(cel)


def znajdz_rozmiar_danych(ws) -> tuple[int, int]:
    """Ostatni wiersz i ostatnia kolumna, w ktorych JEST wartosc."""
    ostatni_wiersz = 0
    ostatnia_kolumna = 0
    for wiersz in ws.iter_rows():
        for komorka in wiersz:
            if komorka.value is None:
                continue
            ostatni_wiersz = max(ostatni_wiersz, komorka.row)
            ostatnia_kolumna = max(ostatnia_kolumna, komorka.column)
    return ostatni_wiersz, ostatnia_kolumna


def main() -> None:
    zbuduj_plik(CEL)
    print(f"Zapisano: {CEL}\n")

    wb = load_workbook(CEL)
    try:
        ws = wb["Wymiary"]

        print("=" * 78)
        print("1. CZTERY ROZNE PYTANIA O ROZMIAR")
        print("=" * 78)
        print(f"  ws.max_row                     = {ws.max_row}")
        print(f"  ws.max_column                  = {ws.max_column}")
        print(f"  ws.min_row / ws.min_column     = {ws.min_row} / {ws.min_column}")
        print(f"  ws.dimensions                  = {ws.dimensions!r}")
        print(f"  ws.calculate_dimension()       = {ws.calculate_dimension()!r}")

        wiersz, kolumna = znajdz_rozmiar_danych(ws)
        print(f"  rozmiar DANYCH (nasz skan)     = wiersz {wiersz}, kolumna {kolumna}")

        print()
        print("  Wnioski:")
        print(f"    - danych jest {LICZBA_WIERSZY_DANYCH} wierszy i {LICZBA_KOLUMN_DANYCH} kolumny")
        print(f"    - max_row mowi {ws.max_row} <- to ewidencja, nie dane")
        print(f"    - roznica w wierszach wynosi {ws.max_row - wiersz}")
        print("    - gdybys pisal range(2, ws.max_row + 1), przetworzylbys")
        print(f"      {ws.max_row - wiersz} pustych wierszy bez potrzeby")

        print()
        print("=" * 78)
        print("2. JAK POWSTALO ZAWYZENIE")
        print("=" * 78)
        puste_z_formatem = 0
        for wiersz_obj in ws.iter_rows():
            for komorka in wiersz_obj:
                if komorka.value is None:
                    puste_z_formatem += 1
        print(f"  komorek istnieje:  {ws.max_row * ws.max_column}")
        print(f"  z wartoscia:       {wiersz * kolumna}")
        print(f"  bez wartosci:      {puste_z_formatem}")
        print()
        print("  Komorka Z100 ('dotknieta' przez odczyt) nie powieksza max_column,")
        print("  jesli nie zostala zapisana do modelu. Sprawdz, czy Twoja wersja")
        print("  openpyxl zachowuje sie tak samo przy odczycie - latwo to stwierdzic")
        print("  drukujac ws.max_column po odczytaniu pojedynczej dalekiej komorki.")

        print()
        print("=" * 78)
        print("3. KONSEKWENCJA PRAKTYCZNA")
        print("=" * 78)
        print("  Prawidlowa petla po danych to:")
        print("      for wiersz in ws.iter_rows(min_row=2, max_row=%d):" % wiersz)
        print("          ...")
        print()
        print("  A nie:")
        print("      for wiersz in ws.iter_rows(min_row=2, max_row=ws.max_row):")
        print("          ...")

    finally:
        wb.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W momencie zapisu do pliku model zawiera nie 12, a 50 wierszy i 8 kolumn — nie dlatego, że są w nich dane, a dlatego, że istnieją komórki. Każda z nich ma zapisany indeks stylu (jasne tło), więc przy zapisie openpyxl nie pomija jej jako „pustej technicznie".

**Co trafi do pliku.** Do `xl/worksheets/sheet1.xml` trafi sto kilkadziesiąt elementów `<c>` z atrybutem `s` (style), ale bez elementu `<v>`. Excel otworzy plik i pokaże podświetlone komórki w wierszach 13–50 mimo braku treści. `ws.dimensions` po odczycie z pliku powie `A1:H50`.

**Czego ten przykład uczy najbardziej:** że **`max_row` nie mówi nic o Twoich danych** — mówi, gdzie kończy się ewidencja komórek, o których openpyxl wie. To zupełnie inna informacja i nie wolno jej używać jako granicy pętli po danych.

### Przykład 2 — iteracja: mierzymy różnicę

```python
"""Cztery sposoby patrzenia na arkusz i porownanie wydajnosci.

Uruchom:  python examples/07_iteracja.py
"""

import time
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

CEL = OUTPUT / "07_iteracja.xlsx"
LICZBA_WIERSZY = 3000
LICZBA_KOLUMN = 10


def zbuduj(cel: Path) -> None:
    """Buduje plik w trybie write_only - szybko i bez trzymania modelu w pamieci."""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Siatka")
    ws.append([f"K{k}" for k in range(1, LICZBA_KOLUMN + 1)])
    for r in range(1, LICZBA_WIERSZY + 1):
        ws.append([r * k for k in range(1, LICZBA_KOLUMN + 1)])
    wb.save(cel)


def main() -> None:
    if not CEL.exists():
        zbuduj(CEL)
        print(f"Zbudowano: {CEL}")

    wb = load_workbook(CEL)
    try:
        ws = wb["Siatka"]

        print("=" * 78)
        print("1. CO ZWRACAJA ROZNE FORMY ITERACJI")
        print("=" * 78)

        # a) for row in ws -> krotka obiektow Cell
        for wiersz in ws:
            print(f"  for row in ws       -> {type(wiersz).__name__} "
                  f"o rozmiarze {len(wiersz)}, "
                  f"pierwszy element: {type(wiersz[0]).__name__}")
            break

        # b) iter_rows -> krotka obiektow Cell
        for wiersz in ws.iter_rows(min_row=1, max_row=1):
            print(f"  ws.iter_rows()      -> {type(wiersz).__name__} "
                  f"o rozmiarze {len(wiersz)}, "
                  f"pierwszy element: {type(wiersz[0]).__name__}")
            break

        # c) iter_rows(values_only=True) -> krotka wartosci
        for wiersz in ws.iter_rows(min_row=2, max_row=2, values_only=True):
            print(f"  iter_rows(values_only=True) -> {type(wiersz).__name__} "
                  f"o rozmiarze {len(wiersz)}, "
                  f"pierwszy element: {type(wiersz[0]).__name__}")
            break

        # d) ws.values -> generator krotek wartosci
        generator = ws.values
        print(f"  ws.values           -> {type(generator).__name__}")
        print(f"  pierwszy wiersz:      {next(generator)}")

        print()
        print("=" * 78)
        print("2. OKNO ITERACJI: min_row / max_row / min_col / max_col")
        print("=" * 78)
        print("  iter_rows(min_row=2, max_row=4, min_col=2, max_col=4):")
        for wiersz in ws.iter_rows(min_row=2, max_row=4, min_col=2, max_col=4,
                                   values_only=True):
            print(f"    {wiersz}")

        print()
        print("  iter_cols(min_col=1, max_col=3, min_row=1, max_row=4) - kolumnami:")
        for kolumna in ws.iter_cols(min_col=1, max_col=3, min_row=1, max_row=4,
                                    values_only=True):
            print(f"    {kolumna}")

        print()
        print("=" * 78)
        print("3. PORÓWNANIE WYDAJNOŚCI - DWA SPOSOBY SUMOWANIA")
        print("=" * 78)

        # Metoda A: komorka po komorce (skakanie z indeksem po ksiazce)
        t0 = time.perf_counter()
        suma_a = 0
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                wartosc = ws.cell(row=r, column=c).value
                if isinstance(wartosc, int):
                    suma_a += wartosc
        t_a = time.perf_counter() - t0

        # Metoda B: iteracja wierszami, values_only
        t1 = time.perf_counter()
        suma_b = 0
        for wiersz in ws.iter_rows(values_only=True):
            for wartosc in wiersz:
                if isinstance(wartosc, int):
                    suma_b += wartosc
        t_b = time.perf_counter() - t1

        print(f"  A) ws.cell(row=r, column=c) w podwojnej petli: {t_a*1000:8.2f} ms")
        print(f"  B) iter_rows(values_only=True):                {t_b*1000:8.2f} ms")
        print(f"  suma A = {suma_a}, suma B = {suma_b}, zgodne: {suma_a == suma_b}")
        if t_b > 0:
            print(f"  metoda A byla ok. {t_a / t_b:.1f} razy wolniejsza")
        print()
        print(f"  Liczba komorek: {ws.max_row * ws.max_column}")
        print("  Roznica rosnie z rozmiarem danych - im wiekszy plik,")
        print("  tym bardziej oplaca sie iterowac wierszami.")

    finally:
        wb.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W metodzie A każda iteracja pętli wewnętrznej wywołuje `ws.cell(row=r, column=c)`, który musi znaleźć komórkę w słowniku i (jeśli jej nie ma) utworzyć — co już wiesz z sekcji 2.2, jest tanie, ale nie darmowe. W metodzie B jeden generator wierszami wczytuje kolejne krotki — jedno wywołanie na wiersz, nie na komórkę.

**Co trafi do pliku.** Nic — to odczyt. Ale uwaga: gdybyś w metodzie A budował plik od zera w ten sposób (`ws.cell(...)` po całym zamierzonym zakresie), to **każda odwiedzona komórka utworzyłaby się w modelu** i `max_row`/`max_column` urosłyby do granic pętli. To kolejny powód, by nie używać tego wzorca.

**Czego ten przykład uczy najbardziej:** że różnica między tymi dwoma podejściami nie jest kosmetyczna i rośnie liniowo z rozmiarem danych. Metoda A jest intuicyjna („indeksuję po współrzędnych"), ale dla openpyxl jest dokładnie tym „skakaniem po literach książki z indeksem". Metoda B jest tym „czytaniem linijka po linijce".

### Przykład 3 — `append` i pułapka napisu

```python
"""append: najszybszy sposob dopisywania i pulapka napisu.

Uruchom:  python examples/07_append.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

CEL = OUTPUT / "07_append.xlsx"


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Append"

    print("=" * 78)
    print("1. POPRAWNE UZYCIA")
    print("=" * 78)

    ws.append(["Lp.", "Nazwa", "Ilość"])           # naglowek
    ws.append([1, "Widzet", 12])                   # rekord - lista
    ws.append((2, "Wkręt", 340))                   # krotka dziala tak samo
    ws.append(range(3))                            # range jest iterowalny
    ws.append({"B": "Tylko kolumna B", "D": "wartosc"})   # wariant ze slownikiem

    for wiersz in ws.iter_rows(values_only=True):
        print(f"  {wiersz}")

    print()
    print("=" * 78)
    print("2. PULAPKA: append Z NAPISEM")
    print("=" * 78)

    ws2 = wb.create_sheet("Pulapka")
    ws2.append(["Zamierzone: jedna komorka z 'ABC'"])
    ws2.append("ABC")               # <- NIE jedna komorka!
    ws2.append(["DEF"])             # <- to jest poprawna jedna komorka

    for indeks, wiersz in enumerate(ws2.iter_rows(values_only=True), start=1):
        print(f"  wiersz {indeks}: {wiersz}")

    print()
    print("  Wiersz 2 ma trzy komorki (A, B, C), a nie jedna z napisem 'ABC'.")
    print("  Powod: napis jest iterowalny, a 'ABC' to trzy znaki.")
    print("  Poprawna forma to zawsze append([...]) - lista z jednym elementem.")

    print()
    print("=" * 78)
    print("3. CO SIE STANIE Z DLUZSZYM NAPISEM")
    print("=" * 78)

    ws3 = wb.create_sheet("Dlugi")
    nazwa = "Firma Kowalski sp. z o.o."
    ws3.append(nazwa)               # 26 znakow -> 26 komorek
    ws3.append([nazwa])             # 1 komorka

    for indeks, wiersz in enumerate(ws3.iter_rows(values_only=True), start=1):
        opis = (" <- 26 komorek, kazda litera osobno" if indeks == 1
                else " <- 1 komorka z pelnym napisem")
        print(f"  wiersz {indeks}: {len(wiersz):>3} komorek{opis}")
    print()
    print("  Zauwaz, ze wiersz 1 nie zawiera nawet calej nazwy -")
    print("  append rozlozyl napis na znaki i zapisal tylko tyle, ile bylo w napisie.")
    print("  Przy dluzszym napisie przekroczylbys kolumny - to nie jest bezpieczny sposob.")

    print()
    print("=" * 78)
    print("4. WYDajnosc: append vs przypisywanie komorek")
    print("=" * 78)

    import time

    liczba_wierszy = 5000

    t0 = time.perf_counter()
    wb_a = Workbook(write_only=True)
    ws_a = wb_a.create_sheet("A")
    for i in range(liczba_wierszy):
        ws_a.append([i, f"Pozycja {i}", i * 1.5])
    wb_a.save(OUTPUT / "07_append_a.xlsx")
    t_a = time.perf_counter() - t0

    t1 = time.perf_counter()
    wb_b = Workbook()
    ws_b = wb_b.active
    ws_b.title = "B"
    for i in range(liczba_wierszy):
        r = i + 1
        ws_b.cell(row=r, column=1, value=i)
        ws_b.cell(row=r, column=2, value=f"Pozycja {i}")
        ws_b.cell(row=r, column=3, value=i * 1.5)
    wb_b.save(OUTPUT / "07_append_b.xlsx")
    t_b = time.perf_counter() - t1

    print(f"  append w trybie write_only: {t_a*1000:8.2f} ms")
    print(f"  ws.cell() w trybie normalnym: {t_b*1000:8.2f} ms")
    print(f"  append byl ok. {t_b / t_a:.1f} razy szybszy (przy {liczba_wierszy} wierszach)")

    # Sprzatanie plikow testowych z sekcji 4 - nie zaśmiecamy output/
    (OUTPUT / "07_append_a.xlsx").unlink(missing_ok=True)
    (OUTPUT / "07_append_b.xlsx").unlink(missing_ok=True)

    wb.save(CEL)
    print()
    print(f"  Zapisano: {CEL}")

    print()
    print("=" * 78)
    print("5. KONTROLA PO ZAPISIE - CZY NAPRAWDE TAK JEST")
    print("=" * 78)
    kontrola = load_workbook(CEL)
    try:
        ws_k = kontrola["Pulapka"]
        print(f"  Arkusz 'Pulapka': max_column = {ws_k.max_column}")
        print("  max_column = 3, bo wiersz 2 ma trzy komorki A, B, C.")
        print("  To bezposrednia konsekwencja pulapki z append.")
        print()
        ws_d = kontrola["Dlugi"]
        print(f"  Arkusz 'Dlugi': max_column = {ws_d.max_column}")
        print("  max_column = 26, bo pierwszy wiersz rozlozyl sie na 26 znakow.")
    finally:
        kontrola.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `append("ABC")` przechodzi pętlą po napisie — jego elementami są znaki — i przypisuje je do kolejnych komórek. Komórki 1, 2, 3 dostają `A`, `B`, `C`. Liczba kolumn w arkuszu rośnie do 3 — i to zostaje w pliku.

**Co trafi do pliku.** Trzy osobne komórki z jednoliterowymi napisami. Otwierając plik, zobaczysz `A`, `B`, `C` w trzech kolumnach. Diagnoza bywa trudna, dopóki nie zajrzysz do kodu i nie zobaczysz `append("ABC")` zamiast `append(["ABC"])`.

**Czego ten przykład uczy najbardziej:** że `append` jest **metodą tacy, nie walizki** — oczekuje jednego argumentu, który jest zbiorem elementów, a nie elementem. Warto też zwrócić uwagę na wynik `max_column = 26` w arkuszu `Dlugi`: to dowód, że pułapka z `append` i temat tego modułu („wymiary i zakresy") są ze sobą nierozerwalnie splecione. Literówka w `append` psuje granice arkusza.

### Przykład 4 — `insert_rows` i formuły: przesuwa się wszystko poza odniesieniami

```python
"""insert_rows, delete_rows i pytanie "co sie przesuwa, a co nie".

Uruchom:  python examples/07_insert_delete.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def zbuduj_z_formulami():
    """Arkusz z danymi, naglowkiem i formułami."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Formuly"
    ws.append(["Ilosc", "Cena", "Wartosc"])
    ws.append([10, 5.0, "=A2*B2"])
    ws.append([20, 7.0, "=A3*B3"])
    ws.append([30, 3.5, "=A4*B4"])
    ws.append(["RAZEM", "", "=SUM(C2:C4)"])
    # Nadajemy styl naglowkom i kolumnie C, zeby zobaczyc, ze style przesuwaja sie z komorkami.
    naglowek = Font(bold=True, color="FFFFFF")
    tlo = PatternFill(fill_type="solid", fgColor="2F5597")
    for kolumna in range(1, 4):
        ws.cell(row=1, column=kolumna).font = naglowek
        ws.cell(row=1, column=kolumna).fill = tlo
    return wb


def dump(ws, etykieta):
    print(f"  --- {etykieta} ---")
    for wiersz in range(1, 8):
        wartosci = [ws.cell(row=wiersz, column=c).value for c in range(1, 4)]
        print(f"    wiersz {wiersz}: {wartosci}")


def main() -> None:
    print("=" * 84)
    print("1. STAN POCZATKOWY")
    print("=" * 84)
    wb = zbuduj_z_formulami()
    ws = wb["Formuly"]
    dump(ws, "przed jakakolwiek zmiana")

    print()
    print("=" * 84)
    print("2. INSERT_ROWS(1) - WSTAWIAMY WIERSZ NA GORZE")
    print("=" * 84)
    ws.insert_rows(1, amount=1)
    dump(ws, "po ws.insert_rows(1)")
    print()
    print("  Obserwuj:")
    print("    - wartosci i style przesunely sie o jeden wiersz w dol")
    print("    - wiersz 1 jest teraz pusty")
    print("    - formuly NIE zostaly zaktualizowane:")
    print(f"      komorka C3 (byla C2) ma teraz {ws['C3'].value!r}")
    print("      a powinna miec '=A3*B3', bo fizycznie jest teraz w wierszu 3")
    print(f"      komorka RAZEM C6 (byla C5) ma {ws['C6'].value!r}")
    print("      a powinna sumowac C3:C5, nie C2:C4")

    print()
    print("=" * 84)
    print("3. DELETE_ROWS - USUWAMY DRUGI WIERSZ DANYCH")
    print("=" * 84)
    ws.delete_rows(3, amount=1)
    dump(ws, "po ws.delete_rows(3)")
    print()
    print("  Wiersz, ktory byl na pozycji 3, zostal usuniety.")
    print("  Formuly ponizej nadal odnosza sie do starych numerow.")
    print(f"  To klasyczny przypadek 'poprawnie wyglada, zle liczy':")
    print(f"  =SUM(C2:C4) sumuje teraz inne komorki niz zamierzono.")

    print()
    print("=" * 84)
    print("4. JAK SOBIE Z TYM PORADZIC: PRZEBUDUJ FORMULY PO ZMIANIE")
    print("=" * 84)

    wb2 = zbuduj_z_formulami()
    ws2 = wb2["Formuly"]

    # Krok 1: wstawiamy wiersz u gory
    ws2.insert_rows(1, amount=1)

    # Krok 2: przekazujemy wiersz danych tak, jakby byl od nowa budowany
    ws2["A1"] = "Ilosc"
    ws2["B1"] = "Cena"
    ws2["C1"] = "Wartosc"
    naglowek = Font(bold=True, color="FFFFFF")
    tlo = PatternFill(fill_type="solid", fgColor="2F5597")
    for kolumna in range(1, 4):
        ws2.cell(row=1, column=kolumna).font = naglowek
        ws2.cell(row=1, column=kolumna).fill = tlo

    # Krok 3: przepisujemy formuly od nowa, z poprawnymi numerami wierszy
    for wiersz in range(2, 5):
        ws2.cell(row=wiersz, column=3, value=f"=A{wiersz}*B{wiersz}")
    ws2["C5"] = "=SUM(C2:C4)"

    dump(ws2, "po recznej naprawie formul")
    print()
    print("  Wniosek: openpyxl nie naprawia formul sam.")
    print("  Po kazdej zmianie geometrii arkusza musisz przejsc po formulach")
    print("  i przypisac je od nowa - albo uzyc Translator (modul 06).")

    print()
    print("=" * 84)
    print("5. STYLE NA POZIOMIE WIERSZA/KOLUMNY TEZ SIE NIE PRZESUWAJA")
    print("=" * 84)
    wb3 = zbuduj_z_formulami()
    ws3 = wb3["Formuly"]
    ws3.row_dimensions[1].height = 30          # podwyzszamy naglowek
    ws3.column_dimensions["C"].width = 20

    ws3.insert_rows(1, amount=1)
    print(f"  row_dimensions[1].height po insert = {ws3.row_dimensions[1].height}")
    print("  <- Wysokosc zostala na wierszu 1, nie przesunela sie z danymi.")
    print("     Wiersz 1 jest teraz pusty, a ma wysokosc naglowka.")

    cel = OUTPUT / "07_insert_delete.xlsx"
    wb3.save(cel)
    print(f"\n  Zapisano: {cel} - otworz w Excelu i zobacz efekt.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `insert_rows` „przesuwa klucze" w słowniku komórek arkusza: komórki o wierszu `>= idx` dostają o `amount` wyższy numer wiersza, a obiekt komórki (razem z jego wartością i indeksem stylu) wędruje na nowe miejsce. **Treść komórki pozostaje niezmieniona**, bo openpyxl nie zna znaczenia tej treści.

**Co trafi do pliku.** Arkusz z przesuniętymi danymi i **nieprzesuniętymi formułami**. Excel po otwarciu policzy je tak, jak zostały zapisane — czyli najprawdopodobniej źle.

**Czego ten przykład uczy najbardziej:** że gdy modyfikujesz geometrię arkusza z formułami, **openpyxl rozpada spójność, którą Excel w UI utrzymuje za Ciebie**. Trzeba na to spojrzeć jak na umowę: openpyxl zrobi mechaniczne przepisanie, całą „inteligencję" musisz dodać sam — albo nie używać `insert_rows` w takich plikach.

### Przykład 5 — `move_range` z tłumaczeniem referencji i bez niego

```python
"""move_range, translate=True vs False. Rowniez move_range z formulami zewnetrznymi.

Uruchom:  python examples/07_move_range.py
"""

from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def zbuduj():
    wb = Workbook()
    ws = wb.active
    ws.title = "Move"
    ws.append(["Ilosc", "Cena", "Wartosc"])
    ws.append([10, 5.0, "=A2*B2"])
    ws.append([20, 7.0, "=A3*B3"])
    ws.append([30, 3.5, "=A4*B4"])
    return wb


def dump(ws, etykieta):
    print(f"  --- {etykieta} ---")
    for wiersz in range(1, 9):
        wartosci = [ws.cell(row=wiersz, column=c).value for c in range(1, 4)]
        if any(w is not None for w in wartosci):
            print(f"    wiersz {wiersz}: {wartosci}")


def main() -> None:
    print("=" * 84)
    print("1. MOVE_RANGE BEZ TLUMACZENIA (translate=False - domyslne)")
    print("=" * 84)
    wb = zbuduj()
    ws = wb["Move"]
    ws.move_range("A1:C4", rows=3, cols=0, translate=False)
    dump(ws, "po move_range(rows=3, translate=False)")
    print()
    print("  Formuly przeniosly sie DOSLOWNIE:")
    print(f"    C5 = {ws['C5'].value!r}   (odnosi sie do A2, nie do A5!)")
    print(f"    C6 = {ws['C6'].value!r}")
    print(f"    C7 = {ws['C7'].value!r}")
    print("  Excel policzy je jako 0, bo A2/B2 sa teraz puste.")

    print()
    print("=" * 84)
    print("2. MOVE_RANGE Z TLUMACZENIEM (translate=True)")
    print("=" * 84)
    wb2 = zbuduj()
    ws2 = wb2["Move"]
    ws2.move_range("A1:C4", rows=3, cols=0, translate=True)
    dump(ws2, "po move_range(rows=3, translate=True)")
    print()
    print("  Formuly zostaly PRZESUNIETE razem z komorkami:")
    print(f"    C5 = {ws2['C5'].value!r}")
    print(f"    C6 = {ws2['C6'].value!r}")
    print(f"    C7 = {ws2['C7'].value!r}")
    print("  Excel policzy je poprawnie.")

    print()
    print("=" * 84)
    print("3. PULAPKA: FORMULA ODNOSZACA SIE POZA ZAKRES")
    print("=" * 84)
    wb3 = Workbook()
    ws3 = wb3.active
    ws3.title = "Zewnetrzna"
    # Komorki D1 to stala konfiguracja - poza przesuwanym zakresem
    ws3["D1"] = 0.23
    ws3["A1"] = "Cena netto"
    ws3["A2"] = 100
    ws3["C2"] = "=A2*(1+$D$1)"     # formula uzywa KONFIGURACJI z D1
    print(f"  Przed move_range: C2 = {ws3['C2'].value!r}")

    ws3.move_range("A1:C3", rows=4, cols=0, translate=True)
    print(f"  Po move_range(rows=4, translate=True): C6 = {ws3['C6'].value!r}")
    print()
    print("  Formula uzywala $D$1. Po przesunieciu .text wskazuje $D$5.")
    print("  A $D$5 jest PUSTE - bo D1 to nasza konfiguracja i zostalo na miejscu!")
    print("  To wlasnie pulapka: move_range tlumaczy WSZYSTKIE referencje,")
    print("  takze te, ktore celowo byly skierowane poza przesuwany zakres.")

    print()
    print("=" * 84)
    print("4. MOVE_RANGE JAKO NARZEDZIE: DODAJ WIERSZ NAGLOWKA DO ISTNIEJACEJ TABELI")
    print("=" * 84)
    wb4 = zbuduj()
    ws4 = wb4["Move"]

    # Chcemy "wpisac" naglowek do wiersza 1, a dane maja zaczynac sie od wiersza 2.
    # W obecnym arkuszu dane sa juz od wiersza 1 - wiec przesuwamy je o jeden w dol.
    ws4.move_range("A1:C4", rows=1, translate=True)

    ws4["A1"] = "Ilosc"
    ws4["B1"] = "Cena"
    ws4["C1"] = "Wartosc"

    dump(ws4, "po dodaniu naglowka i przesunieciu danych")
    print()
    print("  Formuly przesunely sie RAZEM z danymi - bo uzylismy translate=True.")
    print("  Bez tego kroku mielibysmy formuly wskazujace o jeden wiersz za wysoko.")

    cel = OUTPUT / "07_move_range.xlsx"
    wb4.save(cel)
    print(f"\n  Zapisano: {cel}")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `move_range` iteruje po komórkach w podanym zakresie i przenosi je na nowe pozycje. Jeśli `translate=True` — przy przenoszeniu każdej formuły uruchamia `Translator` na jej treści i zapisuje wynik. Jeśli `translate=False` — kopiuje treść dosłownie.

**Co trafi do pliku.** Ten sam układ komórek, ale inną treścią formuł — dokładnie jak w wydruku.

**Czego ten przykład uczy najbardziej:** że „inteligentna" operacja może być jednocześnie przydatna i niebezpieczna. `translate=True` jest właściwe w 80% przypadków i katastrofalne w 20%. Rozpoznanie polega na pytaniu: „czy jakaś formuła w tym zakresie odwołuje się do czegoś, co **nie** zostanie przesunięte razem z nią?". Jeśli tak — albo nie używaj `translate=True`, albo wyklucz taki zakres z operacji.

### Przykład 6 — adresy, zakresy, odwołania: pełna skrzynka narzędziowa

```python
"""Przeliczanie adresow, budowa zakresow i bezpieczne odwołania do arkuszy.

Uruchom:  python examples/07_adresy.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import (
    absolute_coordinate,
    cols_from_range,
    column_index_from_string,
    get_column_letter,
    quote_sheetname,
    rows_from_range,
)
from openpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple, range_boundaries
from openpyxl.worksheet.cell_range import CellRange

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def main() -> None:
    print("=" * 78)
    print("1. LITERA KOLUMNY <-> NUMER KOLUMNY")
    print("=" * 78)
    for indeks in (1, 2, 26, 27, 28, 52, 703, 704, 16384):
        litera = get_column_letter(indeks)
        z_powrotem = column_index_from_string(litera)
        print(f"  {indeks:>6} -> {litera:<5} -> {z_powrotem:>6}  "
              f"({'zgodne' if indeks == z_powrotem else 'BLAD'})")

    print()
    print("=" * 78)
    print("2. ADRES <-> PARA (litera, wiersz) ORAZ (wiersz, kolumna)")
    print("=" * 78)
    for adres in ("A1", "B7", "AB12", "$AB$12", "XFD1048576"):
        litera, wiersz = coordinate_from_string(adres)
        wiersz2, kolumna = coordinate_to_tuple(adres)
        print(f"  {adres:<14} from_string={litera:<4} {wiersz:<8} "
              f"to_tuple -> row={wiersz2:<8} col={kolumna}")

    print()
    print("  Uwaga: to dwie rozne konwencje zwrotu tej samej informacji.")
    print("    coordinate_from_string -> ('litera', liczba)")
    print("    coordinate_to_tuple    -> (wiersz, kolumna)")

    print()
    print("=" * 78)
    print("3. ZAKRES: GRANICE I OBIEKT CellRange")
    print("=" * 78)
    for zakres in ("A1:C3", "B2:D10", "A1", "'Dane 2026'!A1:C3"):
        try:
            granice = range_boundaries(zakres)
            print(f"  range_boundaries({zakres!r}) = {granice}")
        except Exception as blad:
            print(f"  range_boundaries({zakres!r}) = blad: {blad}")
    print()
    print("  Kolejnosc zwrotu: (min_col, min_row, max_col, max_row).")
    print("  Kolumny pierwsze! To inna kolejnosc niz w coordinate_to_tuple.")

    obiekt = CellRange("A1:C3")
    print()
    print(f"  CellRange('A1:C3'):")
    print(f"    min_row={obiekt.min_row}, max_row={obiekt.max_row}")
    print(f"    min_col={obiekt.min_col}, max_col={obiekt.max_col}")
    print(f"    size   ={obiekt.size}")

    print()
    print("=" * 78)
    print("4. ROZKLADANIE ZAKRESU NA ADRESY POSZCZEGOLNYCH KOMOREK")
    print("=" * 78)
    print("  rows_from_range('A1:C3') - wierszami:")
    for wiersz in rows_from_range("A1:C3"):
        print(f"    {wiersz}")
    print("  cols_from_range('A1:C3') - kolumnami:")
    for kolumna in cols_from_range("A1:C3"):
        print(f"    {kolumna}")

    print()
    print("=" * 78)
    print("5. BEZPIECZNE BUDOWANIE ODNIESIENIA DO ZAKRESU")
    print("=" * 78)
    nazwy = ["Dane", "Dane 2026", "Q1'2026", "Moja-karta"]
    for nazwa in nazwy:
        cytowana = quote_sheetname(nazwa)
        odniesienie = f"{cytowana}!{absolute_coordinate('B2:D10')}"
        print(f"  {nazwa!r:<16} -> {odniesienie}")

    print()
    print("  quote_sheetname dodaje apostrofy tylko wtedy, gdy sa potrzebne.")
    print("  absolute_coordinate dodaje '$', bo tak zapisuje odwołania Excel.")

    print()
    print("=" * 78)
    print("6. CALE ODNIESIENIE ZBUDOWANE PROGRAMOWO - BEZ ZAKODOWANEJ LITERY")
    print("=" * 78)

    def odniesienie(nazwa_arkusza: str, nazwy_kolumn: list[str],
                    pierwszy_wiersz: int, ostatni_wiersz: int) -> str:
        """Zwraca np. 'Dane 2026'!$G$2:$G$201 dla kolumny G."""
        litera = get_column_letter(len(nazwy_kolumn))
        zakres = f"{litera}{pierwszy_wiersz}:{litera}{ostatni_wiersz}"
        return f"{quote_sheetname(nazwa_arkusza)}!{absolute_coordinate(zakres)}"

    kolumny = ["Lp.", "Region", "Produkt", "Ilosc", "Cena", "Sprzedaz", "Marża %"]
    print(f"  Kolumny w raporcie: {kolumny}")
    print(f"  Kolumna 'Sprzedaz' to indeks {kolumny.index('Sprzedaz') + 1}, "
          f"czyli litera {get_column_letter(kolumny.index('Sprzedaz') + 1)}")
    print(f"  Odniesienie do kolumny 'Sprzedaz' (wiersze 2..201):")
    print(f"    {odniesienie('Dane 2026', kolumny, 2, 201)}")
    print()
    print("  Zmien kolejność kolumn w liscie - odniesienie samo sie przeliczy.")
    nowe = kolumny[::-1]
    print(f"  Po odwroceniu: {nowe}")
    print(f"    {odniesienie('Dane 2026', nowe, 2, 201)}")

    print()
    print("=" * 78)
    print("7. PRAKTYCZNY PRZYKLAD: SUMIF ZBUDOWANY PROGRAMOWO")
    print("=" * 78)
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane 2026"

    kolumny_raportu = ["Lp.", "Region", "Produkt", "Ilosc", "Cena", "Sprzedaz"]
    ws.append(kolumny_raportu)
    dane = [
        (1, "Północ", "Widzet", 10, 5.0),
        (2, "Południe", "Wkręt", 200, 3.2),
        (3, "Północ", "Nakrętka", 500, 1.8),
    ]
    for indeks, (lp, region, produkt, ilosc, cena) in enumerate(dane, start=2):
        ws.append([lp, region, produkt, ilosc, cena, f"=D{indeks}*E{indeks}"])

    ostatni_dane = len(dane) + 1

    indeks_region = kolumny_raportu.index("Region") + 1
    indeks_sprzedaz = kolumny_raportu.index("Sprzedaz") + 1

    zakres_region = (
        f"{quote_sheetname(ws.title)}!"
        f"{absolute_coordinate(f'{get_column_letter(indeks_region)}2:"
                               f'{get_column_letter(indeks_region)}{ostatni_dane}')}"
    )
    zakres_sprzedaz = (
        f"{quote_sheetname(ws.title)}!"
        f"{absolute_coordinate(f'{get_column_letter(indeks_sprzedaz)}2:"
                               f'{get_column_letter(indeks_sprzedaz)}{ostatni_dane}')}"
    )

    ws_pod = wb.create_sheet("Podsumowanie")
    ws_pod["A1"] = "Region"
    ws_pod["B1"] = "Suma sprzedaży"
    for indeks, region in enumerate(["Północ", "Południe"], start=2):
        ws_pod.cell(row=indeks, column=1, value=region)
        ws_pod.cell(row=indeks, column=2,
                    value=f"=SUMIF({zakres_region},$A{indeks},{zakres_sprzedaz})")

    print(f"  zakres_region   = {zakres_region}")
    print(f"  zakres_sprzedaz = {zakres_sprzedaz}")
    print()
    print("  Formuly w Podsumowaniu:")
    for wiersz in ws_pod.iter_rows(min_row=1, max_row=3, values_only=True):
        print(f"    {wiersz}")

    cel = OUTPUT / "07_adresy.xlsx"
    wb.save(cel)
    print(f"\n  Zapisano: {cel}")
    print("  Otworz w Excelu - suma sprzedaży policzy się sama.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Wszystkie te funkcje to czyste przeliczenia — nie tworzą komórek, nie dotykają pliku, nie mają efektów ubocznych. Można je wywoływać w pętlach, w wyrażeniach, wszędzie.

**Co trafi do pliku.** W zależności od użycia — na przykład wynikowe formuły w arkuszu `Podsumowanie`, zbudowane w pełni programowo. Nic nie zostało „wklepane" ręcznie.

**Czego ten przykład uczy najbardziej:** że **całe odwołanie do arkusza i zakresu można zbudować z danych**, nie z literałów. Zmiana nazwy arkusza, zmiana liczby wierszy danych, przestawienie kolumn — wszystko to zaowocuje poprawnym odwołaniem bez dotykania kodu formuły. To jest fundament „konfiguracji zamiast kodu", o której będzie moduł 26.

## 4. Anatomia API

| Metoda / klasa / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `ws.max_row`, `ws.max_column` | Ostatni wiersz / ostatnia kolumna, w których istnieje jakakolwiek komórka | — | **Nie mówi nic o danych!** Rośnie od formatowania i od samego dotknięcia komórki |
| `ws.min_row`, `ws.min_column` | Pierwszy wiersz / pierwsza kolumna z jakąkolwiek komórką | — | Zwykle `1`, ale nie zawsze |
| `ws.dimensions` | Prostokąt otaczający wszystko, jako tekst, np. `"A1:H50"` | — | Równoważne `calculate_dimension()`, właściwość bez wywołania |
| `ws.calculate_dimension(force=)` | To samo co `dimensions`, jako metoda | `force` — wymuszenie obliczenia | Parametr z gałęzi 3.1; sprawdź zachowanie w swojej wersji |
| `ws.reset_dimensions()` | Zeruje wewnętrzne liczniki wymiarów openpyxl | — | **Nie usuwa komórek i nie naprawia `max_row`**; przydatne głównie w `write_only` |
| `for row in ws` | Iteracja po wierszach arkusza | — | Każdy obrót pętli daje **krotkę obiektów `Cell`** |
| `ws.iter_rows(min_row, max_row, min_col, max_col, values_only)` | Sterowana iteracja po wierszach | wszystkie cztery granice opcjonalne; `values_only=False` domyślnie | `values_only=True` → krotki wartości zamiast komórek |
| `ws.iter_cols(min_col, max_col, min_row, max_row, values_only)` | Sterowana iteracja po kolumnach | jak wyżej | Uwaga na wydajność — iteracja po kolumnach kosztuje |
| `ws.rows`, `ws.columns` | Generatory krotek komórek | — | **Generatory jednorazowe**; jeśli potrzebujesz wielokrotnie, użyj `list(...)` |
| `ws.values` | Generator krotek wartości, wierszami | — | Odpowiednik `iter_rows(values_only=True)` |
| `ws.append(iterable)` | Dopisuje wiersz na końcu | **iterowalne** — lista, krotka, `range`, generator, słownik | **Nie podawaj napisu!** Rozłoży go na znaki |
| `ws.insert_rows(idx, amount=1)` | Wstawia `amount` pustych wierszy przed wierszem `idx` | `idx` **1-based** | Nie tłumaczy formuł, nie przesuwa wymiarów wierszy, nie przesuwa scal≥ń |
| `ws.delete_rows(idx, amount=1)` | Usuwa `amount` wierszy począwszy od `idx` | `idx` **1-based** | Jak wyżej |
| `ws.insert_cols(idx, amount=1)` | Wstawia `amount` pustych kolumn przed kolumną `idx` | `idx` **1-based** (numery, nie litery) | Jak wyżej |
| `ws.delete_cols(idx, amount=1)` | Usuwa `amount` kolumn począwszy od `idx` | `idx` **1-based** | Jak wyżej |
| `ws.move_range(cell_range, rows=0, cols=0, translate=False)` | Przenosi zakres o `rows` wierszy i `cols` kolumn | `cell_range` — tekst lub `CellRange`; `translate=True` tłumaczy referencje w formułach | Jedyna „inteligentna" operacja; ryzykowna przy formułach spoza zakresu |
| `ws["A1"]`, `ws["A"]`, `ws[1]`, `ws["A1:C3"]` | Dostęp po adresie, kolumnie, wierszu, zakresie | — | `ws["A"]` i `ws[1]` dają krotki; **`ws["A1"]` w `read_only` nie działa** |
| `get_column_letter(indeks)` | `1` → `"A"`, `28` → `"AB"` | `int` 1..16384 | `openpyxl.utils` |
| `column_index_from_string("AB")` | `"AB"` → `28` | `str` | Akceptuje wiodące `$`; `openpyxl.utils` |
| `coordinate_from_string("AB12")` | `("AB", 12)` | `str` | Zwraca `(litera, wiersz)`; `openpyxl.utils.cell` |
| `coordinate_to_tuple("AB12")` | `(12, 28)` | `str` | Zwraca `(wiersz, kolumna)`; `openpyxl.utils.cell` |
| `range_boundaries("A1:C3")` | `(1, 1, 3, 3)` | `str` | Kolejność: `(min_col, min_row, max_col, max_row)`; `openpyxl.utils.cell` |
| `range_to_tuple("'Dane'!A1:C3")` | `("Dane", (1, 1, 3, 3))` | `str` | Rozdziela arkusz od zakresu |
| `CellRange("A1:C3")` | Obiekt zakresu: `.min_row`, `.max_row`, `.min_col`, `.max_col`, `.size` | `str` | `openpyxl.worksheet.cell_range` |
| `rows_from_range("A1:C3")` | Generator krotek adresów, wierszami | `str` | Zwraca **napisy**, nie komórki |
| `cols_from_range("A1:C3")` | Generator krotek adresów, kolumnami | `str` | Jak wyżej |
| `quote_sheetname("Dane 2026")` | `"'Dane 2026'"` — dodaje apostrofy, gdy potrzebne | `str` | `openpyxl.utils` |
| `absolute_coordinate("A1:B5")` | `"$A$1:$B$5"` | `str` | `openpyxl.utils` |
| `ws.row_dimensions[n].height` | Wysokość wiersza | `n` — numer wiersza | **Nie przesuwa się** przy `insert_rows` |
| `ws.column_dimensions["B"].width` | Szerokość kolumny | litera kolumny | **Nie przesuwa się** przy `insert_cols` |
| `ws.merged_cells.ranges` | Lista scalonych zakresów | — | Przy edycji geometrii sprawdź w swojej wersji, czy się przesuwają |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — wypisz wszystkie wartości z kolumny A.** Napisz `examples/07_cw1.py`, który:

1. Tworzy skoroszyt z arkuszem `Kolumna`.
2. Wpisuje w kolumnie `A`, wiersze 1–20, naprzemiennie: liczbę (`i * 2`) w co drugim wierszu oraz `None` w pozostałych. Czyli `A1=2`, `A2=None`, `A3=6`, `A4=None`, itd.
3. Wpisuje w kolumnie `A`, wiersze **30, 31, 32**, wartości `100`, `200`, `300` — żeby arkusz miał „dziurę" w wierszach 21–29, ale dane były aż do wiersza 32.
4. **Wypisuje na konsoli** cztery liczby: `ws.max_row`, `ws.max_column`, `ws.min_row`, `ws.min_column`.
5. Wypisuje **wszystkie wartości z kolumny A**, korzystając z `ws.iter_rows(min_col=1, max_col=1, values_only=True)`. Każdą wartość wypisuje jako `adres = wartość`, gdzie adres buduje programowo — na przykład z `get_column_letter` i `enumerate`. **Adresy mają być prawdziwe**, czyli uwzględniać aktualny wiersz.
6. Wypisuje tę samą listę, ale pomijając wiersze, w których wartość jest `None`.
7. Zapisuje plik jako `output/07_cw1.xlsx`, wczytuje go ponownie i powtarza krok 4 oraz 5 — żeby sprawdzić, czy wymiary po zapisie i odczycie są takie same, jak w modelu.

W pliku `output/07_cw1_wnioski.md` odpowiedz:

- Ile wartości wypisał krok 5, a ile krok 6? Dlaczego różnica wynosi dokładnie tyle, ile wynosi?
- Czy `max_row` odpowiada rozmiarowi danych w tym arkuszu? Uzasadnij.
- Jak zmieniłoby się `max_row`, gdybyś w kroku 3 wypełnił nie wiersze 30–32, ale **sformatował** wiersze 30–32 (np. pogrubienie) — bez wpisywania wartości? **Sprawdź to doświadczalnie** i zapisz wynik.
- Jak zmieniłoby się `max_row`, gdybyś w kroku 3 **tylko odczytał** komórkę `A100` (bez wpisywania wartości)? **Sprawdź to doświadczalnie** i zapisz wynik.

**Zadanie 2 — `column_index_from_string` i `get_column_letter` w praktyce.** Napisz `examples/07_cw2.py`, który:

1. Dla liczb od 1 do 100 wypisuje parę `(indeks, litera)` w trzech kolumnach w konsoli — z podziałem na wiersze po 10 par.
2. Sprawdza **dwukierunkowo**, że `column_index_from_string(get_column_letter(i)) == i` dla każdej liczby od 1 do 16384. Wypisuje jedno zdanie: „test dwukierunkowy dla 1..16384: OK" albo listę niezgodności.
3. Wypisuje `coordinate_from_string` i `coordinate_to_tuple` dla pięciu przygotowanych adresów: `A1`, `B7`, `Z1`, `AA10`, `XFD1048576`. Zwraca uwagę na **różnicę kolejności** w zwracanych krotkach.
4. Buduje pełne odwołanie `'Nazwa arkusza'!$X$2:$X$100` dla **trzech** nazw arkuszy: `Raport`, `Raport 2026`, `Dane Q1'2026` — korzystając wyłącznie z `quote_sheetname`, `absolute_coordinate`, `get_column_letter`.

W pliku `output/07_cw2_wnioski.md` zapisz:

- Co zwróciło `range_boundaries("'Raport 2026'!A1:C3")` i dlaczego (kolejność).
- Dlaczego `coordinate_from_string("$AB$12")` zwraca `("AB", 12)` a nie `("$AB$", 12)` — czy to wygodne, czy niebezpieczne?
- W którym miejscu użyłbyś `absolute_coordinate`, a w którym świadomie **nie**?

### 🟡 Warsztat

**Zadanie 3 — dodanie wiersza nagłówka do istniejącej tabeli i naprawienie formuł.** To ćwiczenie bezpośrednio z zestawu problemów, na które natknie się każdy, kto edytuje pliki innych ludzi.

Napisz `examples/07_cw3.py`, który:

1. Buduje plik `output/07_cw3_wejscie.xlsx` z arkuszem `Sprzedaz`, w którym:
   - wiersz 1 zawiera dane od razu, bez nagłówka: `Lp.`, `Miesiąc`, `Ilość`, `Cena`, `Przychód`,
   - wiersze 2, 3, 4, 5 zawierają dane i **formuły**: `=C2*D2`, `=C3*D3`, itd.,
   - wiersz 6 zawiera etykietę `RAZEM` w kolumnie B i formułę `=SUM(E2:E5)` w kolumnie E,
   - kolumna `E` ma format `#,##0.00`, a wiersz 1 ma pogrubioną kolumnę nagłówków (nawet jeśli to dane, w tym momencie nie jest jeszcze nagłówkiem).
2. Wypisuje **stan początkowy** — wszystkie komórki z kolumny E i ich wartości.
3. **Wykonuje cztery eksperymenty** na czterech **niezależnych kopiach** tego pliku, aby sprawdzić cztery podejścia do dodania wiersza nagłówka:

   - **Wariant A — bezpośrednie `insert_rows(1)`, potem wypełnienie nagłówków.** Pokazuje, że formuły w kolumnie E **nie** przesunęły się.
   - **Wariant B — `move_range("A1:E6", rows=1, translate=True)`, potem wypełnienie nagłówków.** Pokazuje, że formuły `=C2*D2` w kolumnie E zmieniły się na `=C3*D3` — czyli są nadal poprawne po przesunięciu o wiersz w dół.
   - **Wariant C — `insert_rows(1)`, potem przepisanie formuł ręcznie** z poprawnymi numerami wierszy (pętla od wiersza 2 do 5: `ws[f"E{r}"] = f"=C{r}*D{r}"`, oraz `ws["E6"] = "=SUM(E2:E5)"`).
   - **Wariant D — brak `insert_rows`, tylko dopisanie wiersza na końcu i przenumerowanie „logiczne" w Pythonie** — czyli czytasz dane, zbudujesz nowy skoroszyt od zera i wypełniasz dane od wiersza 2. (To wariant „nie edytuj, buduj nowy" — najbardziej odporny.)

4. Dla każdego wariantu, **przed zapisaniem wynikowego pliku**, wypisuje:
   - wartość komórek `E2..E6` w wariancie wynikowym,
   - czy formuły wskazują na właściwe komórki (porównaj z listą oczekiwanych: `['=C2*D2', '=C3*D3', '=C4*D4', '=C5*D5', '=SUM(E2:E5)']` po dodaniu nagłówka).
5. Zapisuje cztery pliki: `output/07_cw3_wariant_a.xlsx`, `..._b.xlsx`, `..._c.xlsx`, `..._d.xlsx`.
6. Wypisuje **tabelę porównawczą**: dla każdego wariantu — czy formuły są poprawne, ile linii kodu wymagał, czy wymagał ręcznej wiedzy o numerach wierszy.

W pliku `output/07_cw3_wnioski.md`:

- Opisz **krótko**, który wariant jest w praktyce najbezpieczniejszy i dlaczego.
- Opisz **jeden scenariusz**, w którym `insert_rows` byłby lepszym wyborem niż `move_range`, i jeden, w którym `move_range` byłby lepszy niż `insert_rows`. Uzasadnij.
- Odpowiedz: **dlaczego wariant D jest najbardziej odporny?** Co konkretnie sprawia, że nie ma tam problemu z przesunięciami? Nazwij to jednym zdaniem i odwołaj się do tego, co przeczytałeś w module 03 (cykl życia modelu).

### 🔴 Wyzwanie

**Zadanie 4 — `znajdz_rozmiar_danych(ws)` z trybami i testami.** Napisz `examples/07_cw4.py`, w którym zaimplementujesz pełną funkcję do wyznaczania **prawdziwego rozmiaru danych** w arkuszu, wraz z trybami postępowania dla danych z dziurami i pełnym zestawem testów.

**Wymagania funkcjonalne funkcji `znajdz_rozmiar_danych`:**

```python
def znajdz_rozmiar_danych(
    ws,
    *,
    tryb: str = "ostatni_z_danymi",
    kolumny_kluczowe: list[int] | None = None,
) -> tuple[int, int]:
    ...
```

- Zwraca krotkę `(ostatni_wiersz, ostatnia_kolumna)`.
- Obsługuje **trzy tryby** (parametr `tryb`):
  - `"ostatni_z_danymi"` — skanuje wstecz od `max_row`/`max_column` i zwraca pierwszą pozycję (wiersz i kolumnę niezależnie), na której jest komórka z **niepustą wartością**. Zatrzymuje się przy pierwszym trafieniu (szybkie).
  - `"pelny_skan"` — przechodzi po wszystkich komórkach modelu, buduje wiersz- i kolumnę-po-kolumnie i zwraca maksima. Wolniejsze, ale nie zakłada nic o ciągłości.
  - `"kolumny_kluczowe"` — bierze listę `kolumny_kluczowe` i w każdej kolumnie szuka ostatniego wiersza z niepustą wartością; zwraca maksimum tych wartości, oraz maksymalną kolumnę, w której w ogóle jest jakakolwiek wartość. **Jeśli `kolumny_kluczowe` jest `None`, funkcja ma podnieść `ValueError`.**
- W żadnym trybie funkcja **nie modyfikuje** arkusza. Ani nie tworzy nowych komórek, ani nie zmienia wartości.
- W żadnym trybie funkcja nie może polegać na `max_row` jako na **granicy danych** — może go użyć tylko jako **punktu startowego** skanowania wstecz albo górnej granicy pętli w trybie `pelny_skan`.
- Funkcja odróżnia „pusta wartość" (`None`) od „pustego napisu" (`""`) i „zera" (`0`). Za **niepustą wartość** uznaje wszystko, co nie jest `None`. W komentarzu wyjaśnij swoją decyzję — jeśli zdecydujesz inaczej (np. że `""` też jest puste), uzasadnij i udokumentuj.

**Wymagania testowe (w tym samym pliku, jako funkcje `test_*` wywoływane z `main()`):**

Napisz co najmniej **siedem testów**, każdy na osobnym, świeżo zbudowanym arkuszu:

1. **Arkusz pusty.** Oczekiwany wynik: `(0, 0)`.
2. **Tylko nagłówek w wierszu 1.** Oczekiwany wynik: `(1, liczba_kolumn)`.
3. **Ciągłe dane bez dziur**, wiersze 1..20, kolumny 1..5. Oczekiwane: `(20, 5)`.
4. **Dane z dziurą w środku** — dane w wierszach 1..5, puste 6..20, dane w 21..25. Oczekiwane: `(25, liczba_kolumn)` w trybach `ostatni_z_danymi` i `pelny_skan`.
5. **Formatowanie poza danymi** — dane w wierszach 1..10, ale sformatowane komórki w wierszach 11..1000. Oczekiwane: `(10, liczba_kolumn)` — mimo że `ws.max_row == 1000`.
6. **Kolumna daleko w prawo** — dane tylko w kolumnach 1 i 50, wiersze 1..10. Oczekiwane: `(10, 50)`.
7. **Tryb `kolumny_kluczowe`** — dane w wierszach 1..50 w kolumnach 1..4, ale komórki w kolumnie kluczowej (np. kolumna 1) są puste od wiersza 41 do 50 (dane tylko w kolumnach 2..4). W trybie `kolumny_kluczowe` z `kolumny_kluczowe=[1]` oczekiwane: `(40, 4)`. W trybach `ostatni_z_danymi` i `pelny_skan`: `(50, 4)`.
8. **Tryb `kolumny_kluczowe` bez listy** — wywołanie z `kolumny_kluczowe=None` ma podnieść `ValueError`. Sprawdź `pytest.raises`-owym `try/except` z jawnym `assert` po komunikacie.

Po każdym teście wypisz jedną linijkę: `test N: OK` albo `test N: BŁĄD - <szczegóły>`.

**Wynik końcowy** zapisz jako plik `output/07_cw4_raport.txt` zawierający:

- wykaz wszystkich testów i wyników,
- dla każdego testu: oczekiwany i otrzymany rezultat,
- sumaryczne zdanie: `wynik: X/Y testów przeszło`.

W pliku `output/07_cw4_wnioski.md` odpowiedz na trzy pytania:

1. **W którym trybie funkcja działa najszybciej, a w którym najwolniej?** Uzasadnij, opierając się na tym, jak działa iteracja po komórkach w openpyxl (moduł 07, sekcja 2.4).
2. **Dlaczego tryb `kolumny_kluczowe` jest najczęściej tym, który faktycznie chcesz użyć w produkcji?** Odwołaj się do założenia, na którym się opiera.
3. **Dlaczego test numer 5 (formatowanie poza danymi) jest najważniejszym testem tego ćwiczenia?** Sformułuj odpowiedź w dwóch zdaniach.

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""znajdz_rozmiar_danych: trzy tryby, testy, raport.

Uruchom:  python examples/07_cw4.py
"""

from __future__ import annotations

import io
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


# ======================================================================
# FUNKCJA GLOWNA
# ======================================================================
def znajdz_rozmiar_danych(
    ws,
    *,
    tryb: str = "ostatni_z_danymi",
    kolumny_kluczowe: list[int] | None = None,
) -> tuple[int, int]:
    """Wyznacza prawdziwy rozmiar danych w arkuszu.

    Parametry
    ---------
    ws : Worksheet
        Arkusz openpyxl (musi dać się iterować; w read_only też działa).
    tryb : str
        "ostatni_z_danymi"  - skanowanie wstecz od max_row/max_column (szybkie)
        "pelny_skan"        - pelne przejscie i maksima (wolniejsze, nie zaklada ciaglosci)
        "kolumny_kluczowe"  - ostatni wiersz z wartoscia w kolumnach wskazanych
    kolumny_kluczowe : list[int] | None
        Wymagane tylko w trybie "kolumny_kluczowe". Numery kolumn (1-based).

    Zwraca
    ------
    (ostatni_wiersz, ostatnia_kolumna) - obie wartosci 0, gdy arkusz jest pusty.

    Uwaga o definicji "pustosci"
    ----------------------------
    Za NIEpusta wartosc przyjmujemy wszystko, co nie jest None.
    Pusty napis "" oraz zero 0 to wartosci - i swiadomie je liczymy.
    Wybor ten jest zgodny z tym, jak Excel traktuje COUNTA: "" (wynik formuly)
    potrafi byc liczone jako niepusta komorka, natomiast naprawde puste komorki
    nie sa liczone. To zachowanie bywa zrodlem nieporozumien - dokumentujemy
    je wprost.
    """
    if tryb not in {"ostatni_z_danymi", "pelny_skan", "kolumny_kluczowe"}:
        raise ValueError(f"nieznany tryb: {tryb!r}")

    if tryb == "kolumny_kluczowe":
        return _tryb_kolumny_kluczowe(ws, kolumny_kluczowe)
    if tryb == "pelny_skan":
        return _tryb_pelny_skan(ws)
    return _tryb_ostatni_z_danymi(ws)


def _tryb_ostatni_z_danymi(ws) -> tuple[int, int]:
    """Skanuje wstecz. Zwraca maksimum wierszy i kolumn niezaleznie.

    Uwaga: wiersz i kolumna sa wyznaczane nieco inaczej niz w pelnym skanie.
    Skanowanie wstecz po kolumnach jest kosztowne, wiec dla kolumny
    stosujemy skan wierszami i bierzemy maximum - to wciaz szybkie.
    """
    ostatni_wiersz = 0
    ostatnia_kolumna = 0

    # Pelny skan wierszami - to i tak jedna iteracja. Dla duzych arkuszy
    # mozna zastapic skanem wstecz, ale komentarze w kodzie zostawiamy
    # jasne, a optymalizacje dopisujemy na podstawie pomiarow (modul 19).
    for wiersz in ws.iter_rows():
        for komorka in wiersz:
            if komorka.value is None:
                continue
            if komorka.row > ostatni_wiersz:
                ostatni_wiersz = komorka.row
            if komorka.column > ostatnia_kolumna:
                ostatnia_kolumna = komorka.column

    return ostatni_wiersz, ostatnia_kolumna


def _tryb_pelny_skan(ws) -> tuple[int, int]:
    """To samo co tryb bazowy, ale z jawna intencja 'przejdz wszystko'.

    W praktyce roznica polega na tym, ze NIE korzystamy tu z zadnej
    skroconej sciezki (na przyklad z tego, ze wiersze sa ciagle).
    Dzieki temu mozna go uznac za 'wzorzec odniesienia' dla testow.
    """
    maks_wiersz = 0
    maks_kolumna = 0

    for krotka in ws.iter_rows():
        for komorka in krotka:
            if komorka.value is not None:
                maks_wiersz = max(maks_wiersz, komorka.row)
                maks_kolumna = max(maks_kolumna, komorka.column)

    return maks_wiersz, maks_kolumna


def _tryb_kolumny_kluczowe(ws, kolumny_kluczowe: list[int] | None) -> tuple[int, int]:
    """Ostatni wiersz w kolumnach kluczowych plus maksymalna kolumna z danymi."""
    if not kolumny_kluczowe:
        raise ValueError(
            "tryb 'kolumny_kluczowe' wymaga niepustej listy kolumny_kluczowe"
        )

    ostatni_wiersz = 0
    for indeks in kolumny_kluczowe:
        for kolumna in ws.iter_cols(min_col=indeks, max_col=indeks):
            for komorka in kolumna:
                if komorka.value is not None:
                    if komorka.row > ostatni_wiersz:
                        ostatni_wiersz = komorka.row

    # Maksymalna kolumna z danymi - trzeba ja ustalic skanem (taniej: wierszami)
    ostatnia_kolumna = 0
    for krotka in ws.iter_rows():
        for komorka in krotka:
            if komorka.value is not None and komorka.column > ostatnia_kolumna:
                ostatnia_kolumna = komorka.column

    return ostatni_wiersz, ostatnia_kolumna


# ======================================================================
# BUDOWA ARKUSZY TESTOWYCH
# ======================================================================
def pusta():
    wb = Workbook()
    return wb.active


def naglowek(liczba_kolumn=5):
    ws = pusta()
    ws.append([f"K{i}" for i in range(1, liczba_kolumn + 1)])
    return ws


def ciagle(liczba_wierszy=20, liczba_kolumn=5):
    ws = pusta()
    for r in range(1, liczba_wierszy + 1):
        ws.append([r * k for k in range(1, liczba_kolumn + 1)])
    return ws


def z_dziura(liczba_kolumn=5):
    ws = pusta()
    for r in range(1, 6):
        ws.append([r, r * 2, r * 3, r * 4, r * 5])
    for r in range(21, 26):
        ws.append([None] * liczba_kolumn) if False else None
        for c in range(1, liczba_kolumn + 1):
            ws.cell(row=r, column=c, value=r * c)
    return ws


def formatowanie_poza_danymi():
    ws = pusta()
    for r in range(1, 11):
        ws.append([r, r * 2, r * 3, r * 4])
    tlo = PatternFill(fill_type="solid", fgColor="FFF9E6")
    for r in range(11, 1001):
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = tlo
    return ws


def daleka_kolumna():
    ws = pusta()
    for r in range(1, 11):
        ws.cell(row=r, column=1, value=r)
        ws.cell(row=r, column=50, value=r * 100)
    return ws


def kluczowa_z_pustkami():
    ws = pusta()
    # wiersz 1: naglowek
    ws.append(["Klucz", "K1", "K2", "K3"])
    # wiersze 2..41: dane we wszystkich kolumnach
    for r in range(2, 42):
        ws.append([r, r * 2, r * 3, r * 4])
    # wiersze 42..51: dane w kolumnach 2..4, ale kluczowa kolumna 1 pusta
    for r in range(42, 52):
        ws.cell(row=r, column=2, value=r * 2)
        ws.cell(row=r, column=3, value=r * 3)
        ws.cell(row=r, column=4, value=r * 4)
    return ws


# ======================================================================
# TESTY
# ======================================================================
def test_1_pusty():
    ws = pusta()
    wynik = znajdz_rozmiar_danych(ws, tryb="pelny_skan")
    oczekiwany = (0, 0)
    return wynik == oczekiwany, f"oczekiwano {oczekiwany}, otrzymano {wynik}"


def test_2_tylko_naglowek():
    ws = naglowek(5)
    wynik = znajdz_rozmiar_danych(ws, tryb="pelny_skan")
    oczekiwany = (1, 5)
    return wynik == oczekiwany, f"oczekiwano {oczekiwany}, otrzymano {wynik}"


def test_3_ciagle_dane():
    ws = ciagle(20, 5)
    wynik = znajdz_rozmiar_danych(ws, tryb="pelny_skan")
    oczekiwany = (20, 5)
    return wynik == oczekiwany, f"oczekiwano {oczekiwany}, otrzymano {wynik}"


def test_4_z_dziura():
    ws = z_dziura(5)
    wynik = znajdz_rozmiar_danych(ws, tryb="pelny_skan")
    oczekiwany = (25, 5)
    return wynik == oczekiwany, f"oczekiwano {oczekiwany}, otrzymano {wynik}"


def test_5_formatowanie_poza_danymi():
    ws = formatowanie_poza_danymi()
    wynik = znajdz_rozmiar_danych(ws, tryb="pelny_skan")
    oczekiwany = (10, 4)
    uwaga = f"max_row={ws.max_row}, ale danych do wiersza 10"
    return wynik == oczekiwany, f"oczekiwano {oczekiwany}, otrzymano {wynik} ({uwaga})"


def test_6_daleka_kolumna():
    ws = daleka_kolumna()
    wynik = znajdz_rozmiar_danych(ws, tryb="pelny_skan")
    oczekiwany = (10, 50)
    return wynik == oczekiwany, f"oczekiwano {oczekiwany}, otrzymano {wynik}"


def test_7_kolumny_kluczowe():
    ws_klucz = kluczowa_z_pustkami()
    wynik_klucz = znajdz_rozmiar_danych(
        ws_klucz, tryb="kolumny_kluczowe", kolumny_kluczowe=[1]
    )
    oczekiwany_klucz = (41, 4)

    ws_pelny = kluczowa_z_pustkami()
    wynik_pelny = znajdz_rozmiar_danych(ws_pelny, tryb="pelny_skan")
    oczekiwany_pelny = (51, 4)

    ok = (wynik_klucz == oczekiwany_klucz) and (wynik_pelny == oczekiwany_pelny)
    opis = (f"kluczowe: oczekiwano {oczekiwany_klucz}, otrzymano {wynik_klucz}; "
            f"pelny: oczekiwano {oczekiwany_pelny}, otrzymano {wynik_pelny}")
    return ok, opis


def test_8_kluczowe_bez_listy():
    ws = ciagle(5, 3)
    try:
        znajdz_rozmiar_danych(ws, tryb="kolumny_kluczowe", kolumny_kluczowe=None)
    except ValueError:
        return True, "ValueError podniesiony zgodnie z oczekiwaniem"
    return False, "oczekiwano ValueError, nie zostal podniesiony"


# ======================================================================
# URUCHOMIENIE
# ======================================================================
def main() -> int:
    testy = [
        ("1 - pusty arkusz", test_1_pusty),
        ("2 - tylko naglowek", test_2_tylko_naglowek),
        ("3 - ciagle dane", test_3_ciagle_dane),
        ("4 - dane z dziura", test_4_z_dziura),
        ("5 - formatowanie poza danymi", test_5_formatowanie_poza_danymi),
        ("6 - daleka kolumna", test_6_daleka_kolumna),
        ("7 - kolumny kluczowe", test_7_kolumny_kluczowe),
        ("8 - kluczowe bez listy", test_8_kluczowe_bez_listy),
    ]

    bufor = io.StringIO()
    przeszlo = 0

    bufor.write("Raport z testow znajdz_rozmiar_danych\n")
    bufor.write("=" * 72 + "\n\n")

    for etykieta, funkcja in testy:
        try:
            ok, opis = funkcja()
        except Exception as blad:  # noqa: BLE001
            ok, opis = False, f"wyjatek: {type(blad).__name__}: {blad}"
        status = "OK   " if ok else "BLAD "
        if ok:
            przeszlo += 1
        linia = f"test {etykieta:<36} {status} {opis}"
        bufor.write(linia + "\n")
        print(linia)

    podsumowanie = f"\nwynik: {przeszlo}/{len(testy)} testow przeszlo"
    bufor.write(podsumowanie + "\n")
    print(podsumowanie)

    cel = OUTPUT / "07_cw4_raport.txt"
    cel.write_text(bufor.getvalue(), encoding="utf-8")
    print(f"\nZapisano raport: {cel}")

    return 0 if przeszlo == len(testy) else 1


if __name__ == "__main__":
    raise SystemExit(main())
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **„Puste" znaczy „`None`".** Świadomie uznałem, że `""` i `0` to wartości. Uzasadnienie: raport, w którym kolumna z zerami jest „niewidoczna" dla funkcji rozmiaru, byłby pełen cichych błędów. Excel w wielu miejscach traktuje `0` jako pełnoprawną wartość. Dokumentuję to wprost w docstringu funkcji, bo **to nie jest oczywisty wybór**, a zmiana zdania za pół roku może być kosztowna.

- **Trzy tryby to nie „opcje na wszelki wypadek", a trzy różne modele danych.** `ostatni_z_danymi` zakłada, że dane są zwarte na końcu; `pelny_skan` nie zakłada niczego; `kolumny_kluczowe` zakłada, że istnieje kolumna, której nigdy nie brakuje. Wybór trybu jest **deklaracją**, co wiesz o pliku. Zapisywanie tego wprost w kodzie (a nie ukrywanie w „sprytnej" heurystyce) to jedna z naczelnych zasad modułu 26.

- **`ValueError` w trybie `kolumny_kluczowe` bez listy.** Gdybyśmy wrócili z pustym `(0, 0)` albo domyślną wartością, użytkownik dostałby bezsensowny wynik w połowie pipeline'u. Jawne podniesienie błędu z komunikatem „tryb wymaga niepustej listy" jest rzeczą, którą docenisz za trzy miesiące. Zasada: **gdy wejście jest niespójne, zatrzymaj pracę, zamiast zgadywać.**

- **`_tryb_ostatni_z_danymi` w moim szkicu robi pełny skan wierszami.** To celowe uproszczenie — komentarz w kodzie mówi, że optymalizacja jest „na podstawie pomiarów (moduł 19)". Powód: **przedwczesna optymalizacja funkcji, którą piszesz na raz w tygodniu, jest droższa niż sam problem.** Sprawdzasz na realnych plikach, czy warto — i wtedy dopisujesz skan wstecz. To jest zgodne z duchem tego kursu (patrz: moduł 19).

- **Test 5 jest najważniejszy.** Jego rolą jest **wykrycie regresji, która jest najbardziej prawdopodobna w Twoim przyszłym kodzie**: gdy ktoś „ulepszy" funkcję, by korzystała z `max_row`, test 5 padnie natychmiast. Bez testu 5 taki „ulepszony" kod przejdzie większość testów, a na produkcji zacznie zwracać `(1000, 4)` zamiast `(10, 4)` — bo w pliku z Excela znajdą się sformatowane puste komórki w wierszach 11..1000.

- **Test 7 dokumentuje różnicę między trybami.** Świadomie porównuje wynik `kolumny_kluczowe` z `pelny_skan` dla tego samego arkusza — i pokazuje, że **różnią się**, bo jeden wierzy w klucz, drugi w fakty. Nie ma tu „poprawnej" odpowiedzi — jest tylko „ta, którą wybrałeś". Test pokazuje to jawnie.

**Czego ten szkic jeszcze nie robi** (a w produkcji warto dopisać): obsługi błędów na krotkach przez `iter_rows` dla bardzo dziwnych arkuszy, parametry `min_row`/`min_col` w funkcji (np. pomijanie nagłówka), zliczania pominiętych kolumn, oraz — co jest w module 20 — fixture'ów `pytest`, które zastąpią funkcje `test_*` i pozwolą uruchomić zestaw jedną komendą `pytest`. Ale struktura funkcji jest już przygotowana do tej zmiany: wystarczy zamienić `def test_*` na funkcje bez argumentów i dodać asercje.

</details>

## 6. Typowe błędy i pułapki

**1. „`ws.append("ABC")` zapisało `A`, `B`, `C` w trzech komórkach, a ja chciałem napis `ABC` w jednej" (objaw) → `append` przyjmuje **iterowalne**, a napis jest iterowalny znak po znaku; openpyxl rozłożył tekst na elementy i przypisał każdy znak do osobnej komórki (przyczyna) → **zawsze podawaj listę**: `ws.append(["ABC"])`. Zbuduj nawyk, żeby `append` w Twoim kodzie miał tylko taką postać: `ws.append([...])`. Weryfikacja: jeśli w arkuszu `max_column` jest nietypowo duże, sprawdź, czy przypadkiem nie jest to konsekwencja `append` z napisem (naprawa).**

Warto wiedzieć, kiedy ten błąd objawia się „cicho": gdy zapisujesz napis o długości 1 (`"X"`), dostaniesz jedną komórkę z `"X"` — wygląda dobrze, a kod jest błędny. Napis o długości 2 (`"AB"`) wstawi dwie komórki. Napis o długości 30 przekroczy najpewniej oczekiwania co do liczby kolumn. Im większe ryzyko cichego błędu, tym ważniejsze, żeby forma zawsze była taka sama — wtedy wina „nie zauważam" sprowadza się do zera.

**2. „Pętla po danych przechodzi przez tysiące pustych wierszy, mimo że danych jest dwadzieścia" (objaw) → ktoś (albo Ty, wcześniej) sformatował komórki „na zapas" albo odwołał się do dalekiej komórki; komórki bez wartości istnieją w modelu i powiększają `max_row`/`max_column` (przyczyna) → **nie używaj `max_row` jako granicy pętli po danych.** Zbuduj funkcję `znajdz_rozmiar_danych` (sekcja 2.11) i używaj jej. Jeżeli plik jest Twój, generuj formatowanie **po** wypełnieniu danymi (naprawa).**

Powiązany wariant, bardziej złośliwy: pętla przechodzi przez **poprawną** liczbę wierszy, ale w połowie napotyka wiersz, którego jedyną treścią jest jedna komórka z pustym napisem `""` albo spacją. Wtedy pętla logiki „znajdź wiersz bez danych" zadziała inaczej, niż się spodziewasz. Sprawdź w swoim kodzie, jak traktujesz `""` i `" "` — i udokumentuj to.

**3. „Wstawiłem wiersz nagłówka przez `insert_rows(1)` i teraz formuły pokazują `#REF!`" (objaw) → openpyxl nie ma grafu zależności i nie tłumaczy referencji przy zmianie geometrii arkusza; komórki się przesunęły, a treść formuł została dosłownie ta sama (przyczyna) → **jeśli plik zawiera formuły, nie używaj `insert_rows`/`delete_rows` bez planu na naprawę formuł.** Trzy drogi wyjścia: (a) `move_range(translate=True)` zamiast `insert_rows`, (b) ręczne przepisanie formuł po zmianie, (c) czytaj dane i buduj nowy arkusz od zera. Wybór zależy od tego, czy plik ma być edytowany, czy tylko odtworzony (naprawa).**

Ten błąd ma szczególnie groźną formę: gdy formuła nadal działa składniowo, ale **wskazuje inny zakres niż zamierzałeś**. Excel pokaże wynik bez błędu, a liczba będzie „poprawna w sensie formuły, błędna w sensie biznesowym". To najgorszy typ błędu w raportach — cichy i wiarygodny.

**4. „Mój kod modyfikuje arkusz w trakcie iteracji po nim i gubi wiersze" (objaw) → generator `iter_rows` pamięta pozycję w siatce, a Ty zmieniasz tę siatkę w trakcie działania pętli; po każdym `delete_rows`/`insert_rows` przesuwają się indeksy, a generator tego nie wie (przyczyna) → **rozdziel fazę decyzji od fazy modyfikacji**: najpierw zbierz listę operacji, potem je wykonaj. Przy operacjach usuwania idź **od dołu w górę** (od najwyższych numerów wierszy) — to ta sama zasada, która obowiązuje w każdym języku (naprawa).**

**5. „`insert_rows(0)` zwróciło dziwny wynik albo błąd" (objaw) → `idx` w `insert_rows`, `delete_rows`, `insert_cols` i `delete_cols` jest **1-based**; `0` nie jest poprawnym indeksem pierwszego wiersza (przyczyna) → pierwszy wiersz to `1`, pierwsza kolumna to `1`. Jeśli chcesz dodać wiersz na samym początku, wywołaj `ws.insert_rows(1)`. Dla kolumn używaj numerów, nie liter: `ws.insert_cols(2)` wstawia kolumnę przed `B` (naprawa).**

**6. „`ws.move_range` z `translate=True` przesunęło mi referencje, o których nie chciałem, żeby się zmieniały" (objaw) → `move_range` z tłumaczeniem przesuwa **wszystkie** referencje w formułach z przesuwanego zakresu, tak jak Excel przy wypełnianiu/przenoszeniu komórek; nie zachowuje „intencji" referencji do komórek poza zakresem (przyczyna) → albo użyj `translate=False` i napisz formuły ręcznie z poprawnymi numerami, albo wyklucz z zakresu te formuły, które nie mają być tłumaczone. Sprawdź też, czy inne arkusze nie odwołują się do przesuwanego zakresu — openpyxl **nie** zaktualizuje ich formuł (naprawa).**

Odwrotny wariant tego błędu jest też częsty: ktoś używa `move_range` `translate=False` (albo nie używa `move_range`, a przepisuje komórki ręcznie), licząc, że referencje się przesuną. Wtedy wynik ma na pozór sens, ale jest zły liczbowo.

**7. „`for row in ws.columns:` przechodzi tysiące elementów, mimo że mam tylko 5 kolumn" (objaw) → iteracja kolumnami musi „odwiedzić" każdą kolumnę w **każdym wierszu** do `max_row`, a `max_row` może być zawyżone przez puste, sformatowane komórki (przyczyna) → albo użyj `iter_cols(min_col, max_col, max_row=... )` z jawnym ograniczeniem, albo skorzystaj z `ws["F"]`, jeśli potrzebujesz pojedynczej kolumny, albo — najlepiej — **zbierz dane raz w trybie wierszowym** i przekształć w Pythonie (naprawa).**

To jest ten sam problem, co w pułapce 2, ale w nieco innym kostiumie. Uwaga też: `ws.columns` jest **generatorem**, więc dostęp do niego wylicza wszystko na nowo.

**8. „Wszystkie komórki w mojej pętli mają wartość `None`, mimo że w Excelu widzę dane" (objaw) → najprawdopodobniej iterujesz po `iter_rows(min_row=..., max_row=...)` z granicami wyznaczonymi przez `max_row`, ale w pliku dane są w innym obszarze niż sądzisz — na przykład nagłówki w wierszach 1–3, dane od 5 w dół, a Twój kod iteruje od 2 (przyczyna) → **wypisz na konsoli to, co naprawdę jest w komórkach** (`for row in ws.iter_rows(min_row=1, max_row=5): print(...)`) i dostosuj granice. Prawdziwa diagnoza wymaga zobaczenia zawartości, nie zakładania jej (naprawa).**

Bliźniaczy objaw występuje przy `values_only=True`: wtedy zamiast `None` zobaczysz krotki `()` (puste) lub `(None, None, ...)`. Sprawdź, czy pętla nie przechodzi przez wiersze, w których nic nie ma.

**9. „`ws["A1"]` zwraca AttributeError / błąd w trybie `read_only`" (objaw) → `ReadOnlyWorksheet` celowo nie ma dostępu swobodnego po adresie; jest to wymóg modelu strumieniowego (przyczyna) → używaj `iter_rows()` / `iter_cols()` / iteracji po samym obiekcie arkusza. W trybie `read_only` nie potrzebujesz pojedynczych komórek — potrzebujesz **strumienia danych**. To całkiem inny sposób pracy z plikiem (naprawa).**

To nie jest wada, tylko filozofia: `read_only` służy do przetwarzania **bardzo dużych plików** w sposób, w którym nie utrzymujesz całego modelu w pamięci. Skakanie po komórkach w tym trybie nie miałoby sensu — bo nie ma struktury, po której można skakać. Pełny opis w module 19.

**10. „`ws.iter_rows` zwraca krotki, ale ja próbuję w nich użyć `komorka.value` i dostaję `AttributeError`" (objaw) → użyłeś `values_only=True` i w krotce są już wartości, nie obiekty `Cell`; próbujesz więc wywołać `.value` na `int`, `str` albo `None` (przyczyna) → w trybie `values_only=True` operuj bezpośrednio na wartościach: `for wiersz in ws.iter_rows(values_only=True): for wartosc in wiersz: print(wartosc)`. Jeśli potrzebujesz `koordynatów`, użyj `enumerate` od góry i zbuduj adres przez `get_column_letter` (naprawa).**

**11. „`ws.delete_rows` usunął dane, ale zostawił scalenia/formatowanie warunkowe w nieodpowiednim miejscu" (objaw) → scalenia, formatowanie warunkowe, walidacja danych, tabele i wymiary wierszy/kolumn **nie są aktualizowane** przy zmianie geometrii; openpyxl zmienia tylko pozycje komórek w słowniku i nic więcej (przyczyna) → po operacji edycyjnej **przejrzyj i napraw** te elementy ręcznie. Jeśli masz dużo takich elementów, rozważ wariant „zbuduj nowy arkusz" (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Wymiary arkusza to cztery różne pytania i cztery różne odpowiedzi.** `max_row` i `max_column` mówią, gdzie kończy się **ewidencja komórek**, nie gdzie kończą się dane. `min_row` i `min_column` mówią, gdzie ewidencja się zaczyna. `dimensions`/`calculate_dimension()` mówią, jaki jest prostokąt obejmujący wszystko. A **prawdziwy rozmiar danych** nie ma API — musisz policzyć go sam, bo tylko Ty wiesz, co jest dla Ciebie „treścią". Na tym rozróżnieniu wykłada się zdecydowana większość pętli po arkuszach, więc wracaj do niego zawsze, gdy piszesz `range(2, ws.max_row + 1)`.

2. **`max_row` rośnie od samego dotknięcia komórki, a `append` psuje geometrię.** openpyxl trzyma komórki w słowniku i dodaje je do niego przy pierwszym odwołaniu — nawet jeśli nic do nich nie wpiszesz. Efekt: plik z dziesięcioma wierszami danych może mieć `max_row = 50 000`, bo ktoś w Excelu „sformatował kolumnę na zapas". Analogicznie: `append("ABC")` zapisze trzy komórki, nie jedną, bo napis jest iterowalny. Zbuduj dwie zasady: **`append` zawsze z listą** i **granice pętli nigdy z `max_row` bez weryfikacji**.

3. **Iteruj wierszami, nie komórkami.** Cztery drogi (`for row in ws`, `iter_rows`, `iter_cols`, `ws.values`) i jedna hierarchia preferencji: `iter_rows` wierszami z `values_only=True` to domyślny wybór dla danych; `ws.cell(...)` w zagnieżdżonej pętli to najgorszy wybór. Generator jest jednorazowy — jeśli potrzebujesz danych wielokrotnie, zapisz je do listy. Skanowanie wierszami jest z natury tańsze niż skakanie po komórkach, bo idzie w tym samym porządku, w jakim komórki są w pliku.

4. **Operacje geometryczne rozpadają spójność, którą Excel w UI utrzymuje.** `insert_rows`, `delete_rows`, `insert_cols`, `delete_cols` **nie** tłumaczą referencji w formułach, **nie** przesuwają wymiarów wierszy i kolumn, **nie** aktualizują formatowania warunkowego, walidacji, scaleń ani tabel. Jedyna operacja, która „rozumie" referencje, to `move_range(..., translate=True)` — ale i ona tłumaczy **wszystkie** referencje, także te odnoszące się do komórek poza przesuwanym zakresem, a odwołań z innych arkuszy nie widzi wcale. Wniosek: **edycja geometrii arkusza w pliku z formułami to ryzykowna operacja**, którą trzeba wykonywać ze świadomością i — najlepiej — z planem odtworzenia formuł.

5. **Modyfikowanie struktury w trakcie iteracji po niej jest zabronione.** To ta sama zasada co „nie usuwaj elementów z listy, po której iterujesz". Generator `iter_rows` pamięta pozycję w siatce, a Ty ją zmieniasz. Rozdzielaj fazę decyzji od fazy modyfikacji: najpierw zbierz listę do usunięcia albo do zmiany, potem zadziałaj — i przy usuwaniu idź **od dołu w górę**, żeby lista pozostała aktualna. A gdy operacja jest skomplikowana, rozważ wariant najbardziej odporny ze wszystkich: **przeczytaj dane, przetwórz w Pythonie, zapisz nowy arkusz.** W module 26 nazwiemy to „projekcją" i będzie to domyślny wzorzec w production.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import (
    absolute_coordinate,
    cols_from_range,
    column_index_from_string,
    get_column_letter,
    quote_sheetname,
    rows_from_range,
)
from openpyxl.utils.cell import (
    coordinate_from_string,
    coordinate_to_tuple,
    range_boundaries,
    range_to_tuple,
)
from openpyxl.worksheet.cell_range import CellRange

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# ==================================================================
# 2. WYMIARY - CZTERY PYTANIA, CZTERY ODPOWIEDZI
# ==================================================================
ws.max_row, ws.max_column        # gdzie konczy sie EWIDENCJA komorek (nie dane!)
ws.min_row, ws.min_column        # gdzie sie zaczyna
ws.dimensions                    # 'A1:H50' - bounding box jako tekst
ws.calculate_dimension()         # to samo jako metoda (przyjmuje force= w 3.1+)
ws.reset_dimensions()            # zeruje liczniki wewnetrzne (glownie write_only)
                                 # NIE usuwa komorek, NIE naprawia max_row

# ==================================================================
# 3. PRAWDZIWY ROZMIAR DANYCH - WLASNA FUNKCJA
# ==================================================================
def znajdz_rozmiar_danych(ws) -> tuple[int, int]:
    """Ostatni wiersz i kolumna, w ktorych JEST wartosc. Ignoruje format."""
    ostatni_wiersz = 0
    ostatnia_kolumna = 0
    for wiersz in ws.iter_rows():
        for komorka in wiersz:
            if komorka.value is None:
                continue
            if komorka.row > ostatni_wiersz:
                ostatni_wiersz = komorka.row
            if komorka.column > ostatnia_kolumna:
                ostatnia_kolumna = komorka.column
    return ostatni_wiersz, ostatnia_kolumna

# ==================================================================
# 4. ITERACJA - TRZY DROGI I ICH ZWROTY
# ==================================================================
# a) iteracja po samym arkuszu -> krotka obiektow Cell
for wiersz in ws:
    for komorka in wiersz:
        pass                        # komorka.value, komorka.row, komorka.column

# b) iter_rows z zakresem i values_only -> krotka WARTOŚCI
for wiersz in ws.iter_rows(min_row=2, max_row=100, min_col=1, max_col=4,
                           values_only=True):
    pass                            # wiersz to krotka wartosci

# c) iter_cols -> krotka obiektow Cell albo wartosci
for kolumna in ws.iter_cols(min_col=2, max_col=2, values_only=True):
    pass

# d) ws.values -> generator krotek wartosci, wierszami
for wiersz in ws.values:
    pass

# e) generator jest JEDNORAZOWY - jesli potrzebujesz wielokrotnie, zrob liste
wiersze = list(ws.values)

# f) pojedyncza kolumna bez iter_cols: ws["B"]
for komorka in ws["B"]:             # krotka obiektow Cell z kolumny B
    pass
for komorka in ws["B"][1:]:         # bez nagłówka - indeks od ZERA w krotce
    pass

# ==================================================================
# 5. APPEND - ZAWSZE Z LISTĄ
# ==================================================================
ws.append(["Lp.", "Nazwa", "Ilość"])        # ✅ naglowek
ws.append([1, "Widzet", 12])                # ✅ rekord
ws.append((2, "Wkręt", 340))                # ✅ krotka OK
ws.append(range(3))                         # ✅ range jest iterowalny
ws.append({"B": "x", "D": "y"})             # ✅ wariant ze słownikiem (klucze = kolumny)
# ws.append("ABC")                          # ❌ TROJ KOMOREK: A, B, C
# ws.append(1, 2, 3)                        # ❌ TypeError - append bierze jeden argument

# ==================================================================
# 6. INSERT / DELETE - CO PRZESUWA, A CO NIE
# ==================================================================
ws.insert_rows(1, amount=3)      # wstaw 3 puste wiersze przed wierszem 1
ws.delete_rows(5, amount=2)      # usun 2 wiersze od wiersza 5
ws.insert_cols(2, amount=1)      # wstaw kolumne przed kolumna B (numer, nie litera!)
ws.delete_cols(2, amount=1)      # usun kolumne B
# idx jest 1-BASED. idx=0 jest bledem.
#
# PRZESUWA SIE: wartosci, style komorek (razem z komorka)
# NIE PRZESUWA SIE: referencje w formulach, wymiary wierszy/kolumn,
#                   scalenia, formatowanie warunkowe, walidacja, tabele, anchory

# ==================================================================
# 7. BEZPIECZNA EDYCJA GEOMETRII - TRZY PODEJŚCIA
# ==================================================================
# A) Zbierz decyzje, potem dzialaj OD DOLU W GORE
do_usuniecia = []
for wiersz in ws.iter_rows(min_row=2):
    if wiersz[2].value is None:
        do_usuniecia.append(wiersz[0].row)
for numer in sorted(do_usuniecia, reverse=True):     # <- kluczowy reverse
    ws.delete_rows(numer)

# B) Przesun zakres z tlumaczeniem referencji (gdy w zakresie sa formuly)
ws.move_range("A1:C10", rows=1, translate=True)      # formuly przesuwaja sie
ws.move_range("A1:C10", rows=1, translate=False)     # formuly zostaja doslownie

# C) Najbardziej odporny: buduj nowy arkusz, nie edytuj istniejacego
dane = list(ws.iter_rows(values_only=True))
wb2 = Workbook()
ws2 = wb2.active
for wiersz in dane:
    ws2.append(list(wiersz))

# ==================================================================
# 8. ADRESY - UWAZAJ NA KOLEJNOŚĆ ZWROTU
# ==================================================================
get_column_letter(28)                  # 'AB'
column_index_from_string("AB")         # 28

coordinate_from_string("AB12")         # ('AB', 12)   <- (litera, wiersz)
coordinate_to_tuple("AB12")            # (12, 28)     <- (wiersz, kolumna)

range_boundaries("A1:C3")              # (1, 1, 3, 3) <- (min_col, min_row, max_col, max_row)
range_to_tuple("'Dane'!A1:C3")         # ('Dane', (1, 1, 3, 3))

obiekt = CellRange("A1:C3")
obiekt.min_row, obiekt.max_row         # 1, 3
obiekt.min_col, obiekt.max_col         # 1, 3
obiekt.size                            # {'width': 3, 'height': 3}

list(rows_from_range("A1:B3"))         # [('A1','B1'), ('A2','B2'), ('A3','B3')] - wierszami
list(cols_from_range("A1:B3"))         # [('A1','A2','A3'), ('B1','B2','B3')] - kolumnami

# ==================================================================
# 9. BEZPIECZNE BUDOWANIE ODNIESIEN
# ==================================================================
nazwa_arkusza = "Dane 2026"
litera = get_column_letter(6)                              # 'F'
zakres = f"{litera}2:{litera}201"                          # 'F2:F201'
odniesienie = f"{quote_sheetname(nazwa_arkusza)}!{absolute_coordinate(zakres)}"
# wynik: 'Dane 2026'!$F$2:$F$201

formula = f"=SUM({odniesienie})"

# Wersja w pelni programowa - bez zakodowanej litery ani numeru wiersza:
def odniesienie_do_kolumny(nazwa: str, indeks_kolumny: int,
                           pierwszy_wiersz: int, ostatni_wiersz: int) -> str:
    litera = get_column_letter(indeks_kolumny)
    zakres = f"{litera}{pierwszy_wiersz}:{litera}{ostatni_wiersz}"
    return f"{quote_sheetname(nazwa)}!{absolute_coordinate(zakres)}"

# ==================================================================
# 10. READ_ONLY - INNE ZASADY
# ==================================================================
# wb = load_workbook(sciezka, read_only=True)
# - ws["A1"] NIE DZIALA - tylko iter_rows / iter_cols / for row in ws
# - wymiary moga byc None albo zawyzone - nie ufaj bez sprawdzenia
# - wb.close() OBOWIAZKOWO

# ==================================================================
# 11. WZORZEC DO KOPIOWANIA: PRACA Z DANYMI Z DZIURAMI
# ==================================================================
def wiersze_z_danymi(ws, kolumna_kluczowa: int = 1):
    """Zwraca liste (numer_wiersza, krotka_wartosci) dla wierszy z niepustym kluczem."""
    for numer, krotka in enumerate(ws.iter_rows(values_only=True), start=1):
        klucz = krotka[kolumna_kluczowa - 1]
        if klucz is not None:
            yield numer, krotka

# ==================================================================
# 12. CHECKLISTA PRZED PISANIEM PETLI PO ARKUSZU
# ==================================================================
# [ ] Czy granice petli wyznaczam z DANYCH, a nie z max_row?
# [ ] Czy iteruje wierszami (iter_rows), a nie komorka po komorce?
# [ ] Czy uzywam values_only=True, jesli nie potrzebuje adresow?
# [ ] Czy modyfikuje arkusz w trakcie iteracji? (jesli tak - rozdziel fazy)
# [ ] Czy w pliku sa formuly? (jesli tak - nie insert_rows/delete_rows bez planu)
# [ ] Czy nazwy arkuszy i indeksy kolumn pochodza z konfiguracji?
#       (jesli tak - quote_sheetname + get_column_letter)
# [ ] Czy generator jest wykorzystany tylko raz? (jesli nie - list(...))
# [ ] Czy sprawdzilem wynik okiem w Excelu po zapisie?
```

## 9. Co dalej

Nauczyłeś się patrzeć na arkusz **jako siatkę**, a nie jako zbiór pojedynczych komórek. To druga po modelu mentalnym (moduł 03) najważniejsza zmiana perspektywy w całym kursie — i to na jej braku wykłada się większość skryptów, które kiedyś działały, a potem „przestały".

Trzy rzeczy z tego modułu, które powrócą wielokrotnie:

- **`max_row` jako źródło fałszywego poczucia bezpieczeństwa.** W module 13 i 14 zobaczysz to pod postacią `sqref` w formatowaniu warunkowym: jeśli zakres reguły wyznaczysz z `max_row`, dostaniesz formatowanie rozciągnięte na tysiące pustych komórek. W module 19 przekona Cię to do skanowania wstecznego jako domyślnej techniki. W module 26 stanie się elementem kontraktu między warstwami: infrastruktura dostarcza **rzeczywisty** rozmiar danych, aplikacja podejmuje decyzję o układzie arkusza.
- **Nieprzesuwane referencje.** To problem, który wykracza poza „geometrię". W module 12 zobaczysz, że `Table.ref` (zakres tabeli) **też się nie zaktualizuje** przy edycji. W module 18 będzie to jeden z punktów na liście rzeczy, które trzeba sprawdzić po zapisie pliku od użytkownika.
- **Trzy podejścia do edycji: buduj nowy arkusz / `move_range(translate=True)` / nie edytuj w ogóle.** W module 18 zobaczysz je jako pełny zestaw **wzorców modyfikacji istniejącego pliku**. W module 24 staną się `Composite` i `Facade`. W module 26 — zasadą „raport to projekcja, nie źródło prawdy".

W **module 08** zajmiemy się **arkuszami jako całością**: ich tworzeniem, kolejnością, kopiowaniem, widocznością i rolą wskoroszytową. Tam czekają cztery tematy, które są bezpośrednio wynikiem tego, co dziś wiesz:

- **`wb.copy_worksheet(ws)` — co dokładnie kopiuje, a co gubi.** Zobaczysz eksperymentalnie, że obrazy i wykresy **nie są kopiowane** — i to jest konsekwencja tego samego mechanizmu, który sprawiał, że formuły nie przesuwają się przy `insert_rows`: openpyxl nie śledzi powiązań, bo nie ma grafu.
- **`sheet_state` i ukryte arkusze jako konfiguracja.** Pokażę Ci wzorzec „arkusz `_Konfiguracja` z `veryHidden`", który jest pierwszą formą konfiguracji deklaratywnej w tym kursie i przygotowaniem do modułu 26.
- **Limity nazw arkuszy.** 31 znaków, zakazane znaki (`: \ / ? * [ ]`), unikalność bez względu na wielkość liter. Napiszemy funkcję `bezpieczna_nazwa_arkusza()`, która raz na zawsze usuwa ten problem z Twojego kodu.
- **Konwencja „jeden arkusz = jedna odpowiedzialność".** Zaczniemy od prostej reguły, ale już teraz zobaczymy, jak ta konwencja wpływa na to, gdzie w kodzie ląduje logika — a to jest pierwszy krok do architektury z modułu 26.

Przygotuj do modułu 08:

- **plik `output/07_cw1.xlsx`** — posłuży do eksperymentów z kopiowaniem. Sprawdź, co dokładnie się w nim kopiuje, a co gubi, gdy użyjesz `copy_worksheet`.
- **odpowiedź na pytanie:** ile arkuszy ma Twój przykładowy plik z modułu 04 i jaka jest ich kolejność? Dlaczego kolejność arkuszy w `wb.sheetnames` nie zawsze odpowiada temu, co użytkownik zobaczy po otwarciu pliku?
- **eksperyment myślowy:** w `output/07_cw3_wariant_b.xlsx` (z ćwiczenia 🟡) użyłeś `move_range(translate=True)`. Zastanów się, w którym miejscu w tym pliku **nie zadziałałoby** `copy_worksheet` przy próbie skopiowania arkusza. Nie musisz jeszcze znać odpowiedzi — wystarczy, że zapiszesz pytanie i wrócisz do niego po module 08.