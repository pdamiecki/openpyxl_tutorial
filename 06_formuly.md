# Moduł 06 — Formuły

> **Część:** I — Podstawy pracy ze skoroszytem · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–05

## 0. W tym module nauczysz się

- Rozumiesz, że **formuła to dla openpyxl zwykły tekst zaczynający się od znaku `=`** — i że openpyxl zapisuje go do pliku jako tekst, nic więcej.
- Wiesz, że **openpyxl nie ma silnika obliczeniowego i nigdy go nie będzie miał** — potrafi zapisać wzór, ale nie policzy go za Ciebie. Policzysz dopiero Ty (w Pythonie) albo Excel (przy otwarciu pliku).
- Rozumiesz, czym jest **pamięć podręczna wartości** (ang. *cache*): skąd się bierze, kiedy jej nie ma i dlaczego `data_only=True` zwraca Ci czasem `None` zamiast liczby.
- Znasz **wszystkie trzy pułapki `data_only=True`** — a przede wszystkim wiesz, że **zapisanie skoroszytu wczytanego w tym trybie bezpowrotnie zamienia formuły na wartości** i jak się przed tym zabezpieczyć.
- Odróżniasz **adresowanie względne (`B2`) od absolutnego (`$B$2`)** i wiesz, że openpyxl **nie tłumaczy referencji przy kopiowaniu formuł** — musisz albo powiedzieć wprost, co chcesz, albo użyć `Translator`.
- Umiesz zapisać formułę sięgającą **innego arkusza** (`='Dane 2026'!$B$2`) i **nazwy zdefiniowanej** (`=SUMA_ROCZNA`), oraz wiesz, dlaczego w pliku formuły muszą być po **angielsku i z przecinkami**.
- Wiesz, jak działają **formuły tablicowe** (`ArrayFormula`) i **dynamiczne** (`=UNIQUE()`, `=XLOOKUP()`, `=FILTER()`), gdzie leży pułapka prefiksu `_xlfn.` i czego openpyxl w tym obszarze **nie modeluje**.
- Znasz **kryteria decyzyjne**: kiedy pisać formuły, a kiedy wartości — i rozumiesz zdanie „Excel to warstwa prezentacji, nie silnik biznesowy", do którego wrócimy w module 26.

## 1. Intuicja i analogia

### Analogia główna: instrukcja montażu i zdjęcie złożonego mebla

Wyobraź sobie, że kupujesz mebel, do którego dołączona jest **instrukcja montażu**: „przykręć element A do elementu B, dokręć cztery śruby". Instrukcja opisuje **jak coś zrobić** — ale samego mebla nie składa.

Do tej samej teczki ktoś mógł dołączyć **zdjęcie już złożonego mebla**, zrobione przy ostatnim montażu. To zdjęcie pokazuje **efekt** — ale tylko taki, jaki był wtedy, gdy ktoś naprawdę składał. Jeśli po zdjęciu zmieniono instrukcję, zdjęcie jest nieaktualne. Jeśli nikt nigdy nie składał mebla, zdjęcia nie ma wcale.

Ta teczka to **plik `.xlsx`**. I teraz najważniejsze zdanie tego modułu:

- **Instrukcja montażu to formuła.** Leży w komórce, zaczyna się od `=`, opisuje **jak** coś policzyć.
- **Zdjęcie złożonego mebla to wartość z pamięci podręcznej.** Leży obok, jako druga informacja w tej samej komórce. Pokazuje **wynik**, ale tylko z ostatniego policzenia.
- **openpyxl to notariusz, który spisuje teczkę.** Spisze instrukcję słowo w słowo (bo to tekst). Zajrzy do zdjęcia i odczyta, co na nim jest (bo to dana). Ale **sam mebla nie złoży** — nie ma do tego narzędzi.
- **Excel to stolarz, który składa mebel.** Po otwarciu teczki czyta instrukcję i naprawdę liczy. I przy okazji robi nowe zdjęcie.

Z tego wynikają **wszystkie** konsekwencje, które omówimy:

| Teczka | Odpowiednik w openpyxl | Co dostaniesz |
|---|---|---|
| Tylko instrukcja, bez zdjęcia | Komórka z formułą, zapisana przez openpyxl | `data_only=False` → tekst formuły; `data_only=True` → `None` |
| Tylko zdjęcie, bez instrukcji | Komórka z wartością (liczbą, tekstem) | Ta sama wartość w obu trybach |
| Instrukcja **i** zdjęcie | Plik zapisany przez Excel, w którym formuła ma zapisany wynik | `data_only=False` → tekst formuły; `data_only=True` → liczba |
| Instrukcja ze starym zdjęciem | Plik, w którym ktoś zmienił dane wejściowe, ale nikt nie otworzył w Excelu | `data_only=True` → **nieaktualna** liczba, bez ostrzeżenia |

Ostatni wiersz jest najgroźniejszy: **cache potrafi kłamać, a openpyxl Cię o tym nie ostrzeże.**

### Analogia uzupełniająca: przepisy kucharskie i gotowe dania

Wyobraź sobie zeszyt z przepisami w restauracji. Każda strona ma przepis („podsmaż cebulę, dodaj pomidory, duś 20 minut"), a na marginesie ktoś dopisał ołówkiem, ile porcji wyszło ostatnim razem.

- Jeśli chcesz **wiedzieć, jak zrobić danie** — czytasz przepis. To jest `data_only=False`.
- Jeśli chcesz **wiedzieć, ile wyszło** — patrzysz na dopisek. To jest `data_only=True`.
- Jeśli chcesz **przepisać zeszyt na czysto** i wywalić przepisy, zostawiając tylko dopiski — to jest `load_workbook(data_only=True)` + `save()`. **Zniszczyłeś książkę kucharską i nie masz jak zrobić dania ponownie.** Ten błąd ma dokładnie taki charakter: plik się zapisuje bez błędu, otwiera bez ostrzeżenia, a jedyne, co się stało, to utrata całej logiki.
- Jeśli nikt nigdy nie gotował dania, dopisku nie ma — tylko pusty margines. To jest `None` w trybie `data_only=True`.

### Analogia do adresowania: palec i adres pocztowy

Pomyśl, jak wskazujesz miejsce w wielkiej hali:

- **Palcem: „to pudełko obok tamtego".** Jeśli Ty albo rozmówca się przesuniecie, wskazanie zrozumie się **relatywnie do waszej nowej pozycji**. To adres **względny** — w Excelu `B2`.
- **Adresem: „Magazyn Główny, regał 4, półka 2, pudełko 7".** Zawsze to samo miejsce, niezależnie od tego, kto gdzie stoi. To adres **absolutny** — w Excelu `$B$2`.

W Excelu różnica między `B2` i `$B$2` nie ma żadnego znaczenia dla **jednej** komórki. Ma znaczenie dopiero wtedy, gdy **kopiujesz** formułę w inne miejsce — bo Excel automatycznie przesuwa wskazania palca, a adresy zostawia w spokoju.

I tu jest rzecz, którą trzeba zapamiętać twardo: **openpyxl nie jest rozmówcą, który rozumie Twoje wskazanie palcem. To notariusz, który kopiuje akt dosłownie.** Jeśli napiszesz `ws["D5"] = ws["C5"].value` i w `C5` leży `=SUM(A5:B5)`, to do `D5` trafi **dokładnie** `=SUM(A5:B5)` — a nie `=SUM(A6:B6)`. Excel zachowałby się inaczej. openpyxl nie. Do przesuwania referencji służy osobne narzędzie: `Translator`.

### Analogia do nowych funkcji: `_xlfn.`

Format pliku `.xlsx` powstał w 2006 roku i ma **zamknięty słownik nazw funkcji**. Funkcje dodane później — `UNIQUE`, `XLOOKUP`, `FILTER`, `TEXTSPLIT` — nie są w oryginalnym słowniku. Excel radzi sobie z tym tak, że w pliku zapisuje je z przedrostkiem `_xlfn.` (albo `_xlws.` dla części funkcji arkusza), a **na ekranie pokazuje ładną nazwę bez przedrostka**.

To jak wewnętrzny kod magazynowy: na fakturze dla klienta widnieje „Krzesło tapicerowane", ale w systemie magazynowym to `KRZ-TAP-0417`. Jeśli wpiszesz na fakturze kod magazynowy, klient go nie zrozumie. Jeśli zapiszesz do pliku `UNIQUE` bez przedrostka, Excel może pokazać `#NAME?` — bo szuka w słowniku i nie znajduje.

Wrócimy do tego w sekcji 2.12, ale już teraz zapamiętaj: **przed zapisaniem nowej funkcji do pliku sprawdź, jak zapisuje ją sam Excel** — to test, który zajmuje dwie minuty i oszczędza godziny.

### Analogia do kryterium wyboru: paragon i księga główna

Zadaj sobie jedno pytanie: **czy ten plik ma być dokumentem, czy narzędziem?**

- **Dokument (raport, faktura, zestawienie).** Ma pokazać liczby. Nikt nie kliknie w komórkę, żeby zobaczyć formułę. Wartości wystarczą — a nawet są lepsze, bo są stabilne i widoczne dla systemów czytających plik.
- **Narzędzie (model finansowy, kalkulator, formularz).** Ma się przeliczać, gdy użytkownik zmieni dane wejściowe. Tu formuły są **konieczne**, bo inaczej użytkownik dostanie zamrożony obraz, a nie działające narzędzie.

Analogia: **paragon** pokazuje kwotę (wartość), a **księga główna** zawiera zapisy, które się sumują (formuły). Nie wydaje się paragonu z formułami, ani nie prowadzi księgi głównej jako listy gotowych sum. Większość generatorów raportów w Pythonie powinna produkować **paragony**. Moduł 27 pokaże, jak to rozstrzygać w praktyce.

## 2. Teoria

### 2.1. Formuła to tekst rozpoczynający się od `=`

To dosłownie wszystko, co openpyxl wie o formułach:

```python
from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active
ws.title = "Formuly"

ws["A1"] = 10
ws["B1"] = 20
ws["C1"] = "=SUM(A1:B1)"          # <- to jest formuła

k = ws["C1"]
print(repr(k.value))              # '=SUM(A1:B1)'
print(k.data_type)                # 'f'   <- f jak formula
```

**Co się dzieje w pamięci.** Metoda `Cell._bind_value()` (wywoływana automatycznie przy przypisaniu do `cell.value`) sprawdza, co jej podajesz. Widzi `str`, którego pierwszy znak to `=`. **Nie parsuje go. Nie sprawdza składni. Nie sprawdza, czy `SUM` istnieje.** Ustawia tylko `data_type = 'f'` i zapamiętuje tekst. Koniec.

**Co trafi do pliku.** W `xl/worksheets/sheet1.xml` pojawi się:

```xml
<c r="C1" s="1"><f>SUM(A1:B1)</f></c>
```

Zwróć uwagę na dwie rzeczy. Po pierwsze, **znak `=` nie trafia do pliku** — jest tylko znacznikiem dla openpyxl („to jest formuła"), a Excel w pliku zapisuje samą treść bez `=`. Po drugie, **nie ma elementu `<v>`** — czyli nie ma zapisanego wyniku. To właśnie ta brakująca część odpowiada za brak zdjęcia złożonego mebla i za `None` przy `data_only=True`.

**Konsekwencja praktyczna nr 1: openpyxl nie sprawdzi, czy Twoja formuła ma sens.**

```python
ws["A1"] = "=SUM(A1:B1"        # brak nawiasu - openpyxl zapisze bez mrugnięcia
ws["A2"] = "=NIE_MA_TAKIEJ_FUNKCJI(1;2)"   # nazwa po polsku i średnik - też przejdzie
ws["A3"] = "="                  # sam znak równości - też
ws["A4"] = "=1+"                # niekompletne wyrażenie - też
```

Wszystkie te komórki zostaną zapisane jako formuły, wszystkie z `data_type = 'f'`. Excel po otwarciu pliku pokaże w nich `#NAME?` albo zaproponuje naprawę pliku. **openpyxl jest notariuszem, nie korektorem.**

To jest dokładnie ten moment, w którym warto zapamiętać jedną z **trzech kategorii** kursu:

> **openpyxl potrafi** zapisać w komórce dowolny tekst rozpoczynający się od `=`.
> **openpyxl potrafi tylko częściowo** obsłużyć formuły — konkretnie: odczytać je, zapisać i przetłumaczyć referencje prostym mechanizmem.
> **openpyxl nie potrafi** ich policzyć, zweryfikować składni ani rozstrzygnąć, czy funkcja istnieje w danej wersji Excela.

**Konsekwencja praktyczna nr 2: bez `=` nie ma formuły.**

```python
ws["A1"] = "SUM(A1:B1)"     # <- to jest ZWYKŁY TEKST
```

W Excelu zobaczysz napis `SUM(A1:B1)` wyrównany do lewej. To jeden z najczęstszych i najbardziej irytujących błędów początkujących — kod wygląda dobrze, a plik nie liczy. Zapamiętaj: **`=` jest częścią treści, którą wpisujesz, i jest obowiązkowe.**

**Konsekwencja praktyczna nr 3: formuła to tekst, więc działa interpolacja.**

```python
ostatni = 100
ws["C1"] = f"=SUM(A2:A{ostatni})"      # '=SUM(A2:A100)'
```

Bardzo wygodne przy generowaniu raportów o zmiennej długości — ale też miejsce, w którym łatwo o błąd, gdy `ostatni` wynosi `0` (formuła `=SUM(A2:A0)` jest niepoprawna).

### 2.2. openpyxl nie ma silnika obliczeniowego

To zdanie trzeba powiedzieć wprost, bo jest źródłem największych frustracji:

> **openpyxl nie zawiera implementacji funkcji Excela.** Nie ma tam `SUM`, `VLOOKUP`, `SUMIF`, `TEXT`, `IFERROR`. Nie ma parsera wyrażeń, nie ma arkusza obliczeniowego, nie ma zależności między komórkami. **Nie ma czegoś takiego, co liczy formułę.**

Dlaczego? Bo openpyxl implementuje **format pliku** (OOXML, patrz moduł 02), a nie **program Excel**. Format to sposób zapisania danych; program to maszyna licząca. To dwa zupełnie różne zadania, a napisanie kompatybilnego silnika formuł to przedsięwzięcie na skalę osobnej biblioteki (i takie biblioteki istnieją, np. `formulas` czy `pycel` — ale są niepełne, wolne i nieprzewidywalne przy złożonych formułach; nie są częścią kursu i nie polecam ich w produkcji).

**Trzy praktyczne konsekwencje, które trzeba mieć w głowie zawsze:**

**Konsekwencja A: nie odczytasz wyniku formuły, którego nikt nie policzył.** Jeśli w pliku jest formuła bez cache (bo plik powstał z openpyxl, albo z narzędzia, które nie liczy), to `data_only=True` zwróci `None`. Nie ma tam żadnej liczby do odczytania. Fizyka.

**Konsekwencja B: Twój kod nie może „poczekać, aż się policzy".** Nie ma czego czekać. Formuła nie liczy się „w tle" po zapisaniu pliku. Liczy się dopiero w momencie, gdy program obsługujący arkusz (Excel, LibreOffice, Google Sheets, Numbers, albo biblioteka czytająca cache) otworzy plik i wykona obliczenia.

**Konsekwencja C: logika biznesowa NIE MOŻE mieszkać w formułach.** Jeśli Twój system musi znać sumę, żeby podjąć decyzję (np. „jeśli sprzedaż > 1 mln, wyślij alert"), to **ta suma musi być policzona w Pythonie**. Formuła w Excelu jest dla człowieka, który otworzy plik — nie jest dla Twojego programu. Ten wątek rozwinie się w module 26 jako zasada „Excel to warstwa prezentacji".

### 2.3. Cache — skąd się bierze i kiedy go nie ma

Pamięć podręczna wartości (cache) to po prostu element `<v>` zapisany **obok** elementu `<f>` w tej samej komórce:

```xml
<!-- Komórka z formułą I wynikiem (plik zapisany przez Excel) -->
<c r="C1" s="1"><f>SUM(A1:B1)</f><v>30</v></c>

<!-- Komórka z formułą BEZ wyniku (plik zapisany przez openpyxl) -->
<c r="C1" s="1"><f>SUM(A1:B1)</f></c>
```

**Cache istnieje wtedy, gdy ostatni program, który zapisał plik, umiał liczyć.** Konkretnie:

| Sytuacja | Cache jest? |
|---|---|
| Plik utworzony i zapisany w Excelu | ✅ tak |
| Plik utworzony i zapisany w LibreOffice | ✅ tak (o ile wczytanie przeliczyło formuły — patrz 2.15) |
| Plik utworzony przez openpyxl (zapis formuł) | ❌ **nie** — openpyxl nie liczy, więc nie ma czego zapisać |
| Plik z Excela, ale otwarty i zapisany przez openpyxl | ❌ **nie** — openpyxl przepisuje `<f>`, ale `<v>` po drodze gubi |
| Plik z Excela, w którym ktoś zmienił dane wejściowe, ale nie otworzył go ponownie | ⚠️ cache jest, ale **nieaktualny** |

Trzeci wiersz jest szczególnie ważny i zaskakujący: **przepuszczenie pliku przez openpyxl kasuje cache wartości.** Nawet jeśli tylko wczytasz i zapiszesz plik bez żadnych zmian. W module 18 zobaczymy to jako jeden z elementów „co ginie przy zapisie".

Czwarty wiersz to pułapka, której nie da się wykryć: **nie ma w pliku znacznika „cache jest nieaktualny"**. Jest tylko liczba, która może być stara.

### 2.4. `data_only=True` — trzy pułapki w jednym parametrze

```python
from openpyxl import load_workbook

# Tryb domyślny: dostajesz LITERAŁY FORMUL
wb_f = load_workbook("raport.xlsx")                       # data_only=False

# Tryb "tylko wartości": dostajesz CACHE
wb_w = load_workbook("raport.xlsx", data_only=True)
```

**Co robi `data_only=True`:** instruuje czytnik, żeby przy każdej komórce z formułą **zignorował element `<f>` i wziął element `<v>`** (jeśli istnieje).

**Czego NIE robi:** nie liczy niczego. Nie „odświeża". Nie „pobiera wyników". Czyta to, co już jest w pliku.

**Co dzieje się z modelem w pamięci w trybie `data_only=True`:** komórki z formułami stają się **zwykłymi komórkami z wartościami**. Nie ma śladu, że tam była formuła. `cell.data_type` będzie `'n'` albo `'s'`, a `cell.value` — liczbą albo tekstem. Formuła **nie istnieje w modelu**.

To jest właśnie źródło katastrofy opisanej w kolejnej pułapce.

#### Pułapka 1: brak cache → `None`

```python
# Plik zapisany przez openpyxl - formuła bez cache
wb = load_workbook("output/06_pierwsza_formula.xlsx", data_only=True)
print(wb.active["C1"].value)      # None
```

Widzisz `None` i myślisz: „formuła nie działa". Nie — **formuła jest w pliku zapisana poprawnie, ale nikt jej nie policzył**. Otwórz plik w Excelu, zamknij (Excel zapyta o zapisanie — zapisz), a potem wczytaj ponownie z `data_only=True`. Zobaczysz `30`.

**Analogia:** patrzysz na zdjęcie, które jeszcze nie zostało zrobione. Brakuje nie instrukcji, a fotografii.

Szczególny przypadek: **jeśli pliku nigdy nie otworzy Excel, `data_only=True` zwraca `None` dla wszystkich formuł.** Dlatego program, który ma czytać plik generowany przez inny program w Pythonie, w ogóle nie powinien opierać się na formułach — powinien czytać wartości zapisane wprost.

#### Pułapka 2: zapis po `data_only=True` niszczy formuły — BEZPOWROTNIE

To najważniejsza pułapka tego modułu i jedna z najważniejszych w całym kursie:

```python
from openpyxl import load_workbook

wb = load_workbook("raport_z_excela.xlsx", data_only=True)   # <- UWAGA
wb.save("raport_z_excela_zniszczony.xlsx")
```

**Co się stało.** Przy wczytaniu openpyxl przeczytał tylko elementy `<v>` i zbudował model z **samych wartości**. Elementy `<f>` zostały zignorowane i **w ogóle nie trafiły do modelu**. Przy `save()` openpyxl serializuje cały model od nowa — a w modelu nie ma ani jednej formuły, tylko liczby.

**Efekt:** plik wyjściowy zawiera liczby tam, gdzie były sumy. Wszystkie formuły zniknęły. Nie ma `#REF!`, nie ma ostrzeżenia, nie ma komunikatu. Plik wygląda na poprawny — do momentu, gdy ktoś zmieni dane wejściowe i zauważy, że sumy się nie przeliczają.

**Analogia z sekcji 1:** przepisałeś zeszyt kucharski, wpisując w miejsce przepisów tylko to, ile porcji wyszło ostatnim razem. Zeszyt wygląda na pełny — dopóki nie spróbujesz ugotować dania ponownie.

**Dlaczego tego nie widać?** Bo:
- model w pamięci nie odróżnia „komórki z wartością" od „komórki, która była formułą" — nie ma czego sprawdzić,
- zapis kończy się sukcesem,
- Excel otwiera plik bez ostrzeżeń,
- a liczby są **poprawne** — tylko martwe.

Więc jeśli nie masz kopii zapasowej (a nie masz, jeśli nadpisałeś oryginał), **dane są bezpowrotnie utracone**.

**Jak się zabezpieczyć — trzy poziomy obrony:**

1. **Zasada bezwzględna z modułu 00: nigdy nie zapisuj do pliku wejściowego.** Zawsze nowa nazwa. To ratuje sytuację, ale nie ratuje sensu raportu.
2. **Nigdy nie łącz `data_only=True` z `save()` w tym samym obiegu.** Jeśli potrzebujesz wartości — wczytaj je w osobnej sesji, wyciągnij do Pythona i zamknij. Do zapisu otwórz plik **drugi raz**, w trybie domyślnym.
3. **Napisz strażnika (`guard`), który sprawdza, czy plik w modelu w ogóle ma formuły.** Ale uwaga: w trybie `data_only=True` to niemożliwe (bo w modelu nie ma formuł). Strażnik musi działać **przed** decyzją o wczytaniu — na podstawie osobnego, taniego rozpoznania pliku albo na podstawie tego, jak plik powstał.

Praktyczny strażnik, który sprawdza plik **z dysku** w trybie domyślnym:

```python
from openpyxl import load_workbook


def policz_formuly(sciezka) -> int:
    """Ile komórek z formułami ma plik? Odczyt w trybie domyślnym, tylko do odczytu."""
    wb = load_workbook(sciezka, read_only=True)     # read_only: szybko i tanio
    try:
        ile = 0
        for ws in wb.worksheets:
            for wiersz in ws.iter_rows():
                for komorka in wiersz:
                    if komorka.data_type == "f":
                        ile += 1
        return ile
    finally:
        wb.close()


def bezpieczny_zapis(wb, cel, *, zrodlo_z_formulami: bool) -> None:
    """Zapisuje, ale najpierw sprawdza, czy nie niszczymy pracy czyjejś pracy."""
    if zrodlo_z_formulami:
        raise RuntimeError(
            "Skoroszyt pochodzi z pliku zawierającego formuły, a został wczytany "
            "w trybie data_only=True. Zapis zniszczyłby formuły. "
            "Wczytaj ponownie bez data_only i dopiero wtedy zapisuj."
        )
    wb.save(cel)
```

Zwróć uwagę na kształt tego rozwiązania: **funkcja nie próbuje naprawić problemu — odmawia wykonania.** To dobra praktyka przy operacjach nieodwracalnych (moduł 24 nazwie to Proxy ochronnym).

#### Pułapka 3: cache może być nieaktualny, a Ty się o tym nie dowiesz

Scenariusz: ktoś wygenerował plik w Excelu. Potem — innym narzędziem, np. skryptem VBA, programem ERP albo Twoim poprzednim skryptem w Pythonie — **zmienił dane wejściowe w komórkach A1:A10**. Formuła `=SUM(A1:A10)` w `C1` ma nadal stary wynik w `<v>`, bo nikt jej nie przeliczył.

Wczytujesz plik z `data_only=True` i dostajesz **liczbę, która nie odpowiada danym w tym samym pliku**. Nie ma żadnego sposobu, żeby to wykryć z poziomu openpyxl. Jedyna obrona to **nie ufać cache bez sprawdzenia** — albo policzyć samemu z danych wejściowych i porównać.

To jest dobry moment, żeby wprowadzić nawyk, który wróci w module 20: **każda dana, którą odczytujesz z pliku wygenerowanego przez kogoś innego, jest tylko hipotezą, dopóki jej nie sprawdzisz.**

### 2.5. `data_only=False` — tryb domyślny i co znaczy „domyślny"

```python
wb = load_workbook("raport.xlsx")                # data_only=False
wb = load_workbook("raport.xlsx", data_only=False)  # to samo, jawnie
```

W tym trybie:

- komórka z formułą ma `cell.value == '=SUM(A1:B1)'` i `cell.data_type == 'f'`,
- komórka z wartością ma swoją wartość,
- element `<v>` (cache) jest **ignorowany**.

Ważna konsekwencja: **nie da się w jednej sesji mieć jednocześnie formuły i wartości.** Model trzyma albo jedno, albo drugie. Jeśli potrzebujesz obu, musisz otworzyć plik **dwa razy** i porównać komórka po komórce. To jest poprawny, kanoniczny sposób pracy z plikami zawierającymi formuły:

```python
def porownaj_dwa_swiaty(sciezka):
    """Zwraca listę krotek (arkusz, adres, formula, wartosc_z_cache)."""
    wb_f = load_workbook(sciezka)                   # formuły
    wb_w = load_workbook(sciezka, data_only=True)   # wartości z cache
    wynik = []

    try:
        for nazwa in wb_f.sheetnames:
            ws_f = wb_f[nazwa]
            ws_w = wb_w[nazwa]
            for wiersz in ws_f.iter_rows():
                for komorka in wiersz:
                    if komorka.data_type != "f":
                        continue
                    wartosc = ws_w[komorka.coordinate].value
                    wynik.append(
                        (nazwa, komorka.coordinate, komorka.value, wartosc)
                    )
        return wynik
    finally:
        wb_f.close()
        wb_w.close()
```

**Co się dzieje w pamięci.** Trzymasz dwa niezależne grafy obiektów — dwa razy tyle pamięci. Przy dużych plikach to odczuwalne (moduł 19), więc dla plików rzędu setek tysięcy komórek rozważ `read_only=True` przynajmniej dla jednej z sesji.

**Co trafi do pliku.** Nic — to odczyt. Ale ten wzorzec jest podstawą każdego audytu pliku z formułami i wróci w ćwiczeniu 🔴.

### 2.6. Dwa światy w jednej komórce — dlaczego to jest projekt, a nie błąd

Łatwo pomyśleć: „dlaczego openpyxl nie czyta po prostu obu rzeczy naraz?". Odpowiedź jest architektoniczna: **formuła i wartość to dwie różne informacje o tej samej rzeczy, i tylko jedna z nich jest w danym momencie prawdziwa.**

- Formuła jest **prawdą o tym, jak liczyć** — stabilna, dopóki ktoś jej nie zmieni.
- Wartość w cache jest **prawdą o tym, co wyszło ostatnim razem** — może być nieaktualna.

Gdyby model trzymał obie, każdy kod czytający `cell.value` musiałby zdecydować, którą bierze — i w połowie przypadków podjąłby złą decyzję po cichu. Rozdzielenie na dwa tryby jest więc **jawnym kontraktem**: Ty mówisz, której prawdy chcesz.

**Analogia: dokument i jego skan.** Masz oryginał umowy i skan z pieczątką notarialną. Skan jest wygodniejszy, ale nie wiadomo, czy nie jest sprzed zmiany. Wybierasz, na co patrzysz — i to jest Twoja decyzja, nie biblioteki.

### 2.7. Adresowanie względne i absolutne

Trzy rodzaje odwołań w Excelu:

| Zapis | Nazwa | Co się dzieje przy kopiowaniu w prawo/w dół |
|---|---|---|
| `B2` | względne | kolumna i wiersz się przesuwają |
| `$B$2` | absolutne | **nic się nie przesuwa** |
| `B$2` | mieszane (absolutny wiersz) | kolumna się przesuwa, wiersz nie |
| `$B2` | mieszane (absolutna kolumna) | wiersz się przesuwa, kolumna nie |

**Analogia: palec i adres.** `B2` to „pudełko obok" — przesuwa się razem z Tobą. `$B$2` to „regał 4, półka 2, pudełko 7" — zawsze to samo.

Klasyczny wzorzec, w którym to widać: tabela mnożników. Kolumna nagłówków lat u góry (`B$1` — wiersz się nie przesuwa), wiersz produktów z lewej (`$A2` — kolumna się nie przesuwa):

```python
ws["B2"] = "=B$1*$A2"
```

Wpisujesz to raz w `B2`, kopiujesz na cały prostokąt `B2:D10` i działa. To **wzorzec krzyżowy** (ang. *cross-tabulation*) i pojawi się w ćwiczeniu 🔴.

### 2.8. openpyxl nie tłumaczy referencji przy kopiowaniu — i to jest istotne

Najważniejsza różnica w zachowaniu między Excelem a openpyxl:

```python
# Excel: skopiowanie C1 i wklejenie w C2 daje formułę z przesuniętymi referencjami
# openpyxl: to samo "kopiowanie" daje formułę DOSŁOWNIE taką samą

ws["C1"] = "=SUM(A1:B1)"
ws["C2"] = ws["C1"].value          # '=SUM(A1:B1)' - NIE '=SUM(A2:B2)'!
```

**Dlaczego tak jest.** openpyxl nie wie, że „kopiujesz". Widzi tylko przypisanie tekstu do komórki. A przesuwanie referencji wymaga zrozumienia gramatyki formuły — czego openpyxl nie robi automatycznie, bo zrobiłby to źle w połowie przypadków (bo nie wie, czy chciałeś przesunąć, czy nie).

To jest **zachowanie poprawne i bezpieczne**: openpyxl robi dokładnie to, o co go poprosiłeś, bez „inteligentnego" zgadywania. Ale za tę przewidywalność płaci się obowiązkiem: **jeśli chcesz przesunąć referencje, powiedz to wprost.**

**Jak powiedzieć wprost — trzy podejścia:**

**Podejście 1: generuj formułę z właściwymi numerami (najprostsze i najczęstsze).**

```python
for i in range(2, 101):
    ws.cell(row=i, column=3, value=f"=SUM(A{i}:B{i})")
```

Zaleta: pełna kontrola, zero magii. Wada: formuły są „wypalone" na sztywno — zmiana układu wymaga zmiany kodu. Dla większości generatorów raportów to jest **właściwy wybór**.

**Podejście 2: `Translator` — tłumacz referencji.**

```python
from openpyxl.formula.translate import Translator

oryginal = "=SUM($A2:B2)*$D$1"
tlumacz = Translator(oryginal, origin="C2")

print(tlumacz.translate_formula("C3"))    # '=SUM($A3:B3)*$D$1'
print(tlumacz.translate_formula("D4"))    # '=SUM($A4:C4)*$D$1'
```

`Translator` przyjmuje formułę i **komórkę, w której się znajduje** (`origin`), a potem `translate_formula(dest)` zwraca tę samą formułę ze świadomie przesuniętymi referencjami — dokładnie tak, jak zrobiłby to Excel. Referencje absolutne zostają w miejscu.

To jest **jedyne narzędzie w openpyxl, które „rozumie" gramatykę formuły** (ma tokenizer). Używaj go, gdy naprawdę potrzebujesz kopiowania wzorca — na przykład przy generowaniu siatek formuł.

> **Uwaga:** klasa `Translator` ma też metody pomocnicze do przesuwania formuł przy wstawianiu/usuwaniu wierszy i kolumn. Sprawdź nazwy i dostępność w dokumentacji swojej wersji openpyxl — to obszar, który zmieniał się między wydaniami.

**Podejście 3: `ws.move_range(..., translate=True)`.**

```python
ws.move_range("C2:C10", rows=1, cols=0, translate=True)
```

Przesuwa zakres **i tłumaczy zawarte w nim formuły**. Bez `translate=True` przenosi komórki „jak są". Szczegóły w module 07.

### 2.9. Formuły międzyarkuszowe — apostrofy, spacje, `$`

```python
ws["A1"] = "='Dane 2026'!$B$2"          # arkusz ze spacją - apostrofy OBOWIĄZKOWE
ws["A2"] = "=Dane2026!$B$2"             # arkusz bez spacji - apostrofów nie trzeba
ws["A3"] = "=SUM('Dane 2026'!$C$2:$C$100)"
```

Zasada: **jeśli nazwa arkusza zawiera spację, myślnik, kropkę albo inny znak specjalny, w formule trzeba ją ująć w apostrofy.** Bez nich Excel pokaże błąd składni albo — gorzej — zinterpretuje nazwę inaczej, niż myślisz.

Zamiast pamiętać o tym ręcznie, użyj funkcji z `openpyxl.utils`, którą poznałeś w module 05:

```python
from openpyxl.utils import quote_sheetname, absolute_coordinate

nazwa = "Dane 2026"
odniesienie = f"{quote_sheetname(nazwa)}!{absolute_coordinate('B2')}"
ws["A1"] = f"={odniesienie}"

print(ws["A1"].value)     # ='Dane 2026'!$B$2
```

`quote_sheetname` wstawi apostrofy tylko wtedy, gdy są potrzebne, a `absolute_coordinate` dopisze `$`. To **jedyne poprawne podejście**, jeśli nazwy arkuszy pochodzą z konfiguracji albo od użytkownika.

**Trzy pułapki międzyarkuszowe:**

**Zmiana nazwy arkusza łamie formuły.** Jeśli w kodzie masz `ws.title = "Dane Q4"`, a gdzie indziej formuła `='Dane Q3'!A1`, to openpyxl **nie zaktualizuje formuły** — piszesz tekst, nie masz grafu zależności. Excel zaktualizowałby. openpyxl nie.

**Usunięcie arkusza nie usuwa odwołań.** `del wb["Dane 2026"]` pozostawi w innych arkuszach formuły wskazujące na nieistniejący arkusz — Excel pokaże `#REF!`. Analogia: wyrwałeś stronę z indeksem; odwołania w tekście zostały.

**Referencje do nieistniejącego arkusza w chwili zapisu są dozwolone.** openpyxl nie sprawdza niczego. Więc literówka w nazwie arkusza przejdzie przez cały Twój pipeline i objawi się dopiero w Excelu.

### 2.10. Nazwy zdefiniowane w formułach

Nazwy zdefiniowane (moduł 05, sekcja 2.12) to nie tylko wygoda w kodzie — to **formuła, która nie mówi, gdzie leży wartość**:

```python
from openpyxl import Workbook
from openpyxl.utils import absolute_coordinate, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName

wb = Workbook()
konf = wb.active
konf.title = "Konfiguracja"
konf["A1"] = "Stawka VAT"
konf["B1"] = 0.23

odniesienie = f"{quote_sheetname(konf.title)}!{absolute_coordinate('B1')}"
wb.defined_names.add(DefinedName(name="STAWKA_VAT", attr_text=odniesienie))

ws = wb.create_sheet("Kalkulacje")
ws["A1"] = 100.0
ws["B1"] = "=A1*STAWKA_VAT"          # <- nazwa, nie adres!
```

W Excelu komórka `B1` pokaże `23`. A jeśli kiedyś ktoś przeniesie stawkę VAT do innego arkusza, **zmiana jest w jednym miejscu** (w definicji nazwy) i wszystkie formuły nadal działają.

**Zasięg nazwy — rzecz, która bywa źródłem cichych błędów.** Nazwa może być:

- **globalna** (widoczna w całym skoroszycie) — domyślnie,
- **lokalna dla arkusza** (`localSheetId=0`, `1`, ...) — widoczna tylko z tego arkusza i **przesłaniająca** nazwę globalną o tej samej nazwie.

Ta druga możliwość sprawia, że ten sam `=STAWKA_VAT` może w dwóch arkuszach oznaczać dwie różne komórki. Nie rób tego celowo, chyba że masz naprawdę dobry powód, a jeśli już musisz — nazywaj tak, żeby się nie myliło.

**Nazwa może też wskazywać na formułę, nie na komórkę.** `attr_text` przyjmuje dowolny tekst, więc `DefinedName(name="SUMA_ROCZNA", attr_text="SUM(Suma!$B$2:$B$13)")` utworzy nazwę liczącą, a nie wskazującą. To działa, ale utrudnia debugowanie — w formule nie widać, co tam naprawdę jest. Używaj oszczędnie.

**Nazwy a `data_only`.** Nazwy nie mają cache — cache ma komórka z formułą. Więc cała dyskusja z sekcji 2.4 stosuje się do komórek używających nazw bez żadnych zmian.

### 2.11. Formuły tablicowe — `ArrayFormula`

Klasyczna formuła tablicowa to formuła, która w jednym działaniu operuje na całym zakresie — na przykład zwraca sumę iloczynów. W Excelu wprowadza się ją przez `Ctrl+Shift+Enter` i taka formuła „zajmuje" cały zakres.

**Analogia: kalka kreślarska.** Zwykła formuła to pisanie na jednym kwadracie. Formuła tablicowa to kalka nałożona na prostokąt — piszesz raz, a odbija się na całym obszarze.

W openpyxl reprezentuje ją klasa `ArrayFormula`:

```python
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula

wb = Workbook()
ws = wb.active
ws.title = "Tablicowa"

for i in range(1, 11):                       # kolumny A (ilość) i B (cena)
    ws.cell(row=i, column=1, value=i)
    ws.cell(row=i, column=2, value=i * 1.5)

# Formuła tablicowa obejmująca zakres A1:A10 - przypisujemy ją RAZ,
# w komórce najwyższej, z ref wskazującym na CAŁY zakres.
ws["D1"] = ArrayFormula(ref="D1:D10", text="=A1:A10*B1:B10")

wb.save("output/06_tablicowa.xlsx")
```

**Co się dzieje w pamięci.** Wartością komórki `D1` nie jest `str`, a obiekt `ArrayFormula` z dwoma atrybutami: `text` (sama formuła) i `ref` (zakres, który obejmuje). openpyxl wie, że to formuła tablicowa, i zapisze ją z odpowiednim znacznikiem.

**Co trafi do pliku.** Excel zapisze to jako `<f t="array" ref="D1:D10">` — czyli z informacją „ta formuła obejmuje cały zakres". Dzięki temu Excel wie, że ma ją rozprowadzić na dziesięć komórek.

**Trzy pułapki `ArrayFormula`:**

**Pułapka 1: przypisz raz, nie dziesięć razy.** Jeśli w pętli przypiszesz ten sam `ArrayFormula` do dziesięciu komórek, dostaniesz dziesięć niezależnych formuł tablicowych nałożonych na siebie — a Excel zgłosi uszkodzenie pliku.

**Pułapka 2: `ref` musi być zgodne z rzeczywistym umieszczeniem.** Jeśli przypisujesz formułę do `D1` z `ref="D1:D10"`, ale zakres nie istnieje albo pokrywa się z czymś innym, Excel może zgłosić problem.

**Pułapka 3: odczyt daje obiekt, nie tekst.**

```python
from openpyxl import load_workbook

wb = load_workbook("output/06_tablicowa.xlsx")
k = wb["Tablicowa"]["D1"]
print(type(k.value))       # <class 'openpyxl.worksheet.formula.ArrayFormula'>
print(k.value.text)        # 'A1:A10*B1:B10'  (uwaga: bez znaku '=')
print(k.value.ref)         # 'D1:D10'
```

Uwaga na dwie rzeczy: to nie jest `str`, więc proste `k.value.startswith("=")` wywali się z `AttributeError`. Oraz — w tym obiekcie tekst formuły **nie ma znaku `=` na początku** (to wewnętrzna reprezentacja pliku). Te szczegóły sprawdzaj w swojej wersji, bo reprezentacja bywała zmieniana.

**Czy `ArrayFormula` jest dziś potrzebne?** W większości nowoczesnych scenariuszy **nie**, bo Excel 365 ma formuły dynamiczne (sekcja 2.12), które automatycznie rozlewają się na zakres. `ArrayFormula` zostaje dla:

- zgodności ze starszymi wersjami Excela (2016, 2019),
- plików, w których starsze formuły CSE już istnieją i wczytujesz je z powrotem,
- sytuacji, gdy chcesz mieć pewność, że formuła zachowa się identycznie niezależnie od wersji.

> **Uwaga:** w `openpyxl.worksheet.formula` znajdują się też klasy dla innych wariantów formuł w pliku (m.in. dla formuł udostępnionych i tabel danych). Nie są potrzebne w codziennej pracy, ale jeśli wczytasz plik z Excela i zobaczysz, że `cell.value` nie jest ani `str`, ani liczbą — sprawdź w dokumentacji swojej wersji, co to za klasa.

### 2.12. Formuły dynamiczne i prefiks `_xlfn.`

Formuły dynamiczne to jedna z największych zmian w Excelu od dekad: `UNIQUE`, `SORT`, `FILTER`, `XLOOKUP`, `TEXTSPLIT`, `SEQUENCE`, `LET`, `LAMBDA` i inne. Ich wspólna cecha: **wynik rozlewa się na wiele komórek sam z siebie** (ang. *spill*), bez `Ctrl+Shift+Enter`.

```python
ws["F1"] = "=UNIQUE(A2:A100)"       # lista unikalnych wartości
ws["H1"] = "=XLOOKUP(A2,Klienci!A:A,Klienci!B:B,\"nie znaleziono\")"
ws["J1"] = "=FILTER(A2:B100, B2:B100>1000, \"brak danych\")"
```

**czego się spodziewać — trzy rzeczy do zapamiętania:**

**1. Te funkcje wymagają Excela 365 / 2021+.** W Excelu 2016 zobaczysz `#NAME?` — i to jest zachowanie zgodne z oczekiwaniem, nie błąd Twojego kodu. Jeśli plik ma trafiać do osób ze starszym Excelem, dynamiczne funkcje są złą decyzją.

**2. W pliku te funkcje są zapisane z prefiksem.** Jak w analogii z kodem magazynowym: Excel zapisuje je jako `_xlfn.UNIQUE`, `_xlfn.XLOOKUP`, `_xlfn.FILTER` (część funkcji arkusza dostaje przedrostek `_xlws.`), a wyświetla bez przedrostka. **Jeśli zapiszesz `=UNIQUE(...)` bez przedrostka, Excel może pokazać `#NAME?`, mimo że używasz najnowszej wersji.**

Nie zgaduj, które funkcje i w jakim wariantie wymagają przedrostka — **sprawdź empirycznie**:

1. W Excelu wpisz formułę z nową funkcją (np. `=XLOOKUP(...)`).
2. Zapisz plik jako `.xlsx`.
3. Zmień rozszerzenie na `.zip`, rozpakuj, otwórz `xl/worksheets/sheet1.xml`.
4. Zobacz, jaką dokładnie postać formuły zapisał Excel.
5. Wpisz w openpyxl **dokładnie to samo**.

Ten test zajmuje dwie minuty i jest **jedyną** metodą, która daje pewność przy Twojej wersji Excela. Listy przedrostków w internecie się różnią i szybko się starzeją.

**3. openpyxl nie tworzy metadanych rozlania (spill).** Gdy Excel 365 zapisuje formułę dynamiczną, dopisuje w pliku informacje o metadanych komórki (atrybuty `cm`/`vm`), które opisują, że komórka jest częścią rozlanego wyniku. **openpyxl tego nie generuje.** W praktyce Excel przy otwarciu takiego pliku zwykle sam rozlewa formułę od nowa — ale to zachowanie zależy od wersji i nie jest gwarantowane. Przetestuj na swojej wersji i nie zakładaj, że plik z formułą dynamiczną będzie wyglądał identycznie po wygenerowaniu przez Pythona i ręcznie przez Excela.

> **Krótko i uczciwie:** openpyxl **potrafi** zapisać tekst formuły dynamicznej. **Potrafi tylko częściowo** zagwarantować, że Excel ją poprawnie rozleje i że zachowanie będzie identyczne z ręcznie utworzonym plikiem. **Nie modeluje** metadanych rozlania ani tego, czy funkcja istnieje w Twojej wersji Excela.

**Alternatywa dla przypadków, w których to musi „po prostu działać":** policz `UNIQUE` i `FILTER` w Pythonie. To zwykle kilka linii (zbiór, słownik, list comprehension), a wynik zapisany jako wartości działa w **każdej** wersji Excela, jest stabilny i natychmiast widoczny dla innych programów. Zapamiętaj to — wrócimy do tego w tabeli decyzyjnej w sekcji 2.17.

### 2.13. Locale: formuły w pliku są zawsze po angielsku i zawsze z przecinkami

Bardzo ważna i bardzo często mylona rzecz. W polskim Excelu widzisz `=SUMA(A1;B1)`. Ale w **pliku** Excel zapisuje to jako `SUM(A1,B1)`.

Dlaczego? Bo format OOXML (ECMA-376) definiuje gramatykę formuły **w wariancie en-US** — ze względu na przenośność plików. Konwersję na lokalny język robi program wyświetlający, w momencie otwarcia. Plik jest bezjęzykowy.

**Konsekwencja: w openpyxl pisz formuły po angielsku i z przecinkiem.**

```python
ws["C1"] = "=SUM(A1:B1)"            # ✅ poprawnie w pliku
ws["C2"] = "=SUMA(A1;B1)"           # ❌ Excel pokaże #NAME?
ws["C3"] = "=IF(A1>0,\"tak\",\"nie\")"   # ✅ przecinki, angielskie IF
ws["C4"] = "=JEŻELI(A1>0;\"tak\";\"nie\")"   # ❌ to nie zadziała
```

Uwaga na przecinek dziesiętny: w pliku jest **kropka**.

```python
ws["D1"] = "=ROUND(A1*1.23,2)"      # ✅ kropka dziesiętna, przecinek jako separator
```

**Wyjątek: stałe tablicowe.** W pliku stałe tablicowe zapisuje się z **przecinkiem jako separatorem kolumn i średnikiem jako separatorem wierszy**:

```python
ws["E1"] = "=SUM({1,2,3;4,5,6})"    # tablica 2x3: {kol,kol,kol;kol,kol,kol}
```

**Skąd się biorą błędy w tym obszarze.** Dwie przyczyny praktyczne:

1. **Kopiowanie formuły z UI Excela.** Użytkownik kopiuje `=SUMA(A1;B1)`, Ty wklejasz to do kodu i dziwisz się, że nie działa. To najczęstszy przypadek.
2. **Generowanie formuły z tekstu zależnego od locale.** Jeśli budujesz formułę w f-stringu z liczbami, pamiętaj o **punkcie dziesiętnym**, nie przecinku. `f"=ROUND({wartosc},2)"` da przy polskim locale coś takiego: `=ROUND(1234,56,2)` — i Excel zobaczy **trzy** argumenty. To cichy i bardzo frustrujący błąd.

Bezpieczna konwersja liczby do fragmentu formuły:

```python
def liczba_do_formuly(value: float) -> str:
    """Konwertuje liczbę do zapisu akceptowanego w pliku .xlsx (kropka dziesiętna)."""
    return repr(float(value)).replace("inf", "#").replace("nan", "#")   # NaN/inf nie mają odpowiednika
```

(Zamiana `nan`/`inf` na `#` to zabezpieczenie przed wstawieniem do formuły tekstu, który Excel odrzuci — lepiej, by formuła była jawnie błędna i widoczna, niż by cały plik wyglądał podejrzanie.)

### 2.14. Wymuszanie przeliczenia: `fullCalcOnLoad`

Praktyczny problem: generujesz plik z formułami w openpyxl. Wiesz, że nie mają cache. Chcesz, żeby użytkownik po otwarciu **od razu** zobaczył liczby, a nie puste komórki albo komunikat o przeliczaniu.

Rozwiązanie: ustaw w skoroszycie flagę „przelicz wszystko przy wczytywaniu". W pliku odpowiada za to element `calcPr` z atrybutem `fullCalcOnLoad`.

```python
from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties


def wymus_przeliczenie(wb) -> None:
    """Ustawia flagę pełnego przeliczenia przy otwarciu skoroszytu.

    Sprawdź nazwy atrybutów w dokumentacji swojej wersji openpyxl -
    to rzadko używany obszar API.
    """
    if wb.calculation is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True


wb = Workbook()
ws = wb.active
ws["A1"], ws["B1"] = 10, 20
ws["C1"] = "=SUM(A1:B1)"

wymus_przeliczenie(wb)
wb.save("output/06_wymuszone.xlsx")
```

**Co trafi do pliku.** W `xl/workbook.xml` pojawi się `<calcPr calcId="..." fullCalcOnLoad="1"/>`. Excel po otwarciu wykona pełne przeliczenie całego skoroszytu, zanim cokolwiek pokaże.

**Dwie uczciwe uwagi:**

- To **nie sprawia, że openpyxl liczy**. To prośba skierowana do programu, który otworzy plik. Jeśli plik otworzy Twój własny skrypt w Pythonie z `data_only=True`, **nadaj cache nie będzie** — bo nikt go nie policzył i nie zapisał.
- To rzadko używany fragment API. Sprawdź, czy `wb.calculation` i `CalcProperties` zachowują się w Twojej wersji tak, jak tutaj. Kod powyżej jest defensywny (`if wb.calculation is None`) właśnie dlatego.

### 2.15. Liczenie „bez Excela" — LibreOffice w trybie headless

Jeśli Twój pipeline działa na serwerze bez Excela, a Ty **naprawdę** potrzebujesz wartości z formuł, masz jedno realistyczne wyjście: **LibreOffice w trybie wsadowym**.

```bash
soffice --headless --convert-to xlsx --outdir output/przeliczone output/06_formuly.xlsx
```

LibreOffice otwiera plik, wykonuje obliczenia i zapisuje nowy plik — tym razem z cache. Potem możesz go wczytać w openpyxl z `data_only=True` i dostać liczby.

**Trzy zastrzeżenia, o których trzeba wiedzieć:**

1. **Ustawienie przeliczania przy wczytywaniu.** LibreOffice domyślnie bywa ostrożny wobec formuł z obcych formatów. W opcjach (Narzędzia → Opcje → Calc → Formuła / „Recalculation on File Load") można to zmienić. Sprawdź zachowanie w swojej wersji — różni się między wydaniami i konfiguracjami.
2. **Wierność obliczeń.** LibreOffice to inny silnik niż Excel. Dla 95% formuł wyniki będą identyczne; przy funkcjach statystycznych, formatowaniu tekstu, funkcjach finansowych i nowych funkcjach dynamicznych **mogą się różnić**. Nie używaj tego do rozliczeń.
3. **Koszt.** Konwersja dużego pliku to sekundy (albo dziesiątki sekund). Jeśli twój pipeline generuje plik na żądanie, to może być nieakceptowalne. To architektoniczna decyzja — wróć do niej w module 26.

**Alternatywa, która jest zwykle lepsza:** policz to w Pythonie. Jeżeli potrzebujesz sumy, to `sum()` w Pythonie liczy w mikrosekundach i daje **dokładnie ten sam wynik, co Excel** (poza drobnymi różnicami w zaokrągleniach zmiennoprzecinkowych). Jeśli potrzebujesz `SUMIF` — to `defaultdict` i pętla. Jeśli `VLOOKUP` — słownik. Cała tabela decyzyjna w sekcji 2.17 jest o tym, kiedy warto, a kiedy nie warto.

### 2.16. Formuły w trybach `read_only` i `write_only`

**`read_only=True`:** czytasz formuły albo wartości bez zmian. `data_only=True` działa normalnie. Bardzo przydatne do szybkiego audytu dużego pliku (jak `policz_formuly()` z sekcji 2.4). Pamiętaj o `wb.close()`.

**`write_only=True`:** formuły są **trywialne** — `append` z tekstem zaczynającym się od `=` po prostu działa:

```python
from openpyxl import Workbook

wb = Workbook(write_only=True)
ws = wb.create_sheet("Dane")
ws.append(["A", "B", "Suma"])
ws.append([10, 20, "=SUM(A2:B2)"])     # <- działa
wb.save("output/06_write_only.xlsx")
```

Ale w `write_only`:

- nie ma dostępu swobodnego — nie odczytasz komórki ani nie sprawdzisz, co się zapisało,
- nie zaplanujesz formuły, która odwołuje się do komórki „wyżej" w tym samym arkuszu... właściwie możesz, bo Excel rozwiązuje odwołania po wczytaniu — ale Ty nie masz jak tego zweryfikować,
- nie ma sensu mówić o cache — openpyxl i tak nic nie liczy.

To dobry tryb dla generowania dużych raportów z formułami. Szczegóły wydajnościowe — moduł 19.

### 2.17. `=` w danych użytkownika — punkt ostrzeżenia

To najniebezpieczniejsza rzecz w tym module, więc mówię wprost: **jeśli wstawiasz do komórki tekst pochodzący od użytkownika i ten tekst zaczyna się od `=`, openpyxl zapisze go jako formułę.**

```python
komentarz_od_klienta = "=HYPERLINK(\"http://zly.example\",\"kliknij\")"
ws["A1"] = komentarz_od_klienta      # <- TO NIE JEST TEKST. TO JEST FORMULA.
```

Excel otworzy plik, wykona formułę i pokaże klikalny link. Jeśli komentarz brzmiał `=cmd|'/c calc'!A1`, sytuacja jest znacznie gorsza — to klasyczny atak **formula injection / DDE injection**. Pełne omówienie i funkcja sanitizująca — moduł 21. Tutaj zapamiętaj samo ostrzeżenie:

> **Każda wartość pochodząca spoza Twojego kodu musi być jawnie „uprzejmie poproszona", żeby była tekstem, zanim trafi do komórki.**

Jeśli **naprawdę** musisz zapisać tekst zaczynający się od `=`, istnieje niskopoziomowy sposób: najpierw ustaw wartość, a potem nadpisz `data_type`:

```python
komorka = ws.cell(row=1, column=1)
komorka.value = "=SUM(A1:B1)"      # openpyxl ustawi tu data_type='f'
komorka.data_type = "s"            # <- nadpisujemy na "string"

wb.save("output/06_tekst_udajacy_formule.xlsx")
```

Excel pokaże w tej komórce **tekst** `=SUM(A1:B1)`, wyrównany do lewej, bez uruchamiania formuły. To działa, bo `data_type` decyduje o zapisie do XML. Sprawdź w swojej wersji — to niskopoziomowy trik, ale szeroko stosowany.

**Lepsza droga — i ta, którą polecam:** nie wstawiaj do komórek niczego, co zaczyna się od znaku wyzwalającego. Usuń lub zamień ten znak na wejściu do systemu:

```python
ZNAKI_WYZWALAJACE = ("=", "+", "-", "@")


def jako_tekst(value) -> str:
    """Zamienia wartość w tekst, który Excel NIE zinterpretuje jako formułę."""
    tekst = "" if value is None else str(value)
    if tekst.startswith(ZNAKI_WYZWALAJACE):
        return "'" + tekst          # apostrof jest widoczny i jednoznaczny
    return tekst
```

Apostrof na początku **jest widoczny w komórce** (w odróżnieniu od apostrofu wpisanego ręcznie w Excelu, który jest tylko znacznikiem UI) — i to jest zaleta: użytkownik widzi, że wartość została zmieniona świadomie. W module 21 zbudujemy z tego pełną, przetestowaną funkcję sanitizującą.

### 2.18. Wzorzec decyzyjny: wartości czy formuły?

To najważniejsza tabela w tym module. Wracaj do niej, gdy nie wiesz, co wybrać.

| Kryterium | Wybierz **WARTOŚCI** | Wybierz **FORMUŁY** |
|---|---|---|
| **Kto odbiera plik** | System, inny program, archiwum | Człowiek w Excelu, który ma dalej pracować |
| **Kto czyta plik automatycznie** | ✅ zawsze wartości — formuły bez cache to `None` | ❌ tylko jeśli plik zawsze przejdzie przez Excel |
| **Czy dane wejściowe będą się zmieniać w Excelu** | Nie | Tak — wtedy formuły są niezbędne |
| **Audytowalność** | Liczba bez pochodzenia | Widać, jak powstała — czasem wymóg formalny |
| **Deterministyczność** | ✅ ten sam wsad = ten sam plik | Zależy od wersji i ustawień Excela |
| **Wydajność** | Szybko: piszesz `int`/`float` | Wolniej: tysiące łańcuchów + przeliczanie w Excelu |
| **Rozmiar pliku** | Mniejszy | Większy (formuła + format) |
| **Ryzyko błędu** | Błąd widać od razu, bo liczysz sam | Błąd ukryty w formule, objawia się dopiero w Excelu |
| **Kiedy Excel nie jest dostępny** | ✅ działa wszędzie | ❌ bez Excela wartości nie ma |
| **Zgodność ze starszym Excelem** | ✅ | ⚠️ nowe funkcje → `#NAME?` |
| **Bezpieczeństwo** | ✅ brak injection | ⚠️ dane użytkownika w formule = ryzyko |

**Reguła kciuka, którą proponuję na cały kurs:**

> **Domyślnie pisz wartości. Formuły dodawaj świadomie, jako element funkcji pliku, a nie jako sposób liczenia.**

Uzasadnienie jest praktyczne: **formuła to obietnica, że ktoś ją policzy.** Jeśli nie kontrolujesz, co stanie się z plikiem po jego wygenerowaniu — nie możesz tej obietnicy złożyć.

**Wzorzec hybrydowy, który jest często najlepszy** (i który zbudujemy w ćwiczeniu 🔴): **podsumowanie formułą + kontrola w wartościach.**

- Arkusz `Podsumowanie` zawiera formuły (`=SUMIF(...)`, `=COUNTIF(...)`) — użytkownik widzi żywe liczby i może zmieniać dane wejściowe.
- Ten sam kod w Pythonie liczy te same agregaty i wpisuje je do arkusza `Kontrola` jako **wartości**, złotą czcionką, z komentarzem „policzone w Pythonie, {data}".
- Rozbieżność między tymi dwiema kolumnami (**jeśli wystąpi**) oznacza, że dane w pliku zmieniły się po wygenerowaniu — albo że Excel liczy inaczej, niż myślisz.

To nie jest paranoja — to dokładnie ten rodzaj zabezpieczenia, który w audycie wyłapuje ciche przekłamania. I jest przy okazji świetnym przykładem na to, że **formuły i wartości nie wykluczają się — pełnią różne funkcje.**

## 3. Przykłady krok po kroku

### Przykład 1 — pierwsza formuła i co naprawdę leży w pliku

```python
"""Formuła: co jest w modelu, co jest w pliku, co widzi Excel.

Uruchom:  python examples/06_pierwsza.py
"""

from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Formuly"

    ws["A1"] = 10
    ws["B1"] = 20
    ws["C1"] = "=SUM(A1:B1)"
    ws["C1"].number_format = "0.00"

    print("=" * 78)
    print("1. CO JEST W MODELU (przed zapisem)")
    print("=" * 78)
    k = ws["C1"]
    print(f"  value     = {k.value!r}")
    print(f"  typ       = {type(k.value).__name__}")
    print(f"  data_type = {k.data_type!r}")
    print("  ^ openpyxl nie liczy. To dosłownie tekst zaczynający się od '='.")

    cel = OUTPUT / "06_pierwsza.xlsx"
    wb.save(cel)
    print(f"\n  Zapisano: {cel} ({cel.stat().st_size} B)")

    print()
    print("=" * 78)
    print("2. CO TRAFIŁO DO PLIKU (zaglądamy do XML)")
    print("=" * 78)
    with ZipFile(cel) as archiwum:
        xml = archiwum.read("xl/worksheets/sheet1.xml").decode("utf-8")

    # Wypisujemy tylko fragment z komórką C1
    poczatek = xml.find('r="C1"')
    if poczatek > 0:
        fragment = xml[max(0, poczatek - 20):poczatek + 80]
        print(f"  ...{fragment}...")

    print()
    print("  ^ Widzisz <f>SUM(A1:B1)</f>, ale NIE MA <v>30</v>.")
    print("    Brak <v> = brak wartości w pamięci podręcznej = brak 'zdjęcia'.")
    print("    Zwróć też uwagę: w pliku NIE MA znaku '=' - to znacznik openpyxl.")

    print()
    print("=" * 78)
    print("3. ODCZYT - DWA TRYBY, DWIE ODPOWIEDZI")
    print("=" * 78)

    wb_f = load_workbook(cel)
    wb_w = load_workbook(cel, data_only=True)
    try:
        kf = wb_f["Formuly"]["C1"]
        kw = wb_w["Formuly"]["C1"]
        print(f"  data_only=False -> value={kf.value!r:<16} "
              f"typ={type(kf.value).__name__:<5} data_type={kf.data_type!r}")
        print(f"  data_only=True  -> value={kw.value!r:<16} "
              f"typ={type(kw.value).__name__:<5} data_type={kw.data_type!r}")
        print()
        print("  ^ W trybie data_only=True dostajesz None, bo Excel nigdy tego pliku")
        print("    nie policzył. Formuła JEST poprawna - brakuje tylko wyniku.")
    finally:
        wb_f.close()
        wb_w.close()

    print()
    print("=" * 78)
    print("4. CO TERAZ ZROBIĆ")
    print("=" * 78)
    print("  Otwórz plik w Excelu, zapisz go i uruchom ten skrypt ponownie -")
    print("  zobaczysz, że data_only=True zwraca 30.")
    print("  Tego kroku openpyxl nie wykona za Ciebie.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Przy `ws["C1"] = "=SUM(A1:B1)"` openpyxl widzi tylko, że tekst zaczyna się od `=`. Nie sprawdza, czy `SUM` istnieje, czy referencje są poprawne, czy nawiasy się domykają. Ustawia `data_type='f'` i idzie dalej.

**Co trafi do pliku.** Element `<f>` z treścią bez `=`, bez elementu `<v>`. Zapamiętaj ten obrazek — będzie się powtarzał przez cały moduł.

**Czego się nauczysz po uruchomieniu.** Że `None` w trybie `data_only=True` **nie jest błędem** — to brak zdjęcia w teczce. I że naprawa jest poza Twoim skryptem: potrzebny jest program, który policzy.

### Przykład 2 — dwa światy: formuły i wartości w jednym pliku

```python
"""Ten sam plik wczytany dwa razy: formuły i wartości z pamięci podręcznej.

Do działania potrzebny jest plik, który Excel już policzył i zapisał.
Skrypt sam wykryje, czy tak jest, i podpowie, co zrobić.

Uruchom:  python examples/06_dwa_swiaty.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

CEL = OUTPUT / "06_dwa_swiaty.xlsx"


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaz"

    ws.append(["Miesiąc", "Sprzedaż", "Marża %", "Zysk", "Kumulacja"])
    dane = [("styczeń", 10_000, 0.12), ("luty", 12_500, 0.15),
            ("marzec", 9_800, 0.09), ("kwiecień", 14_200, 0.18)]

    for wiersz, (miesiac, sprzedaz, marza) in enumerate(dane, start=2):
        ws.cell(row=wiersz, column=1, value=miesiac)
        ws.cell(row=wiersz, column=2, value=sprzedaz)
        ws.cell(row=wiersz, column=3, value=marza)
        # Zysk = sprzedaż * marża - formuła z odwołaniem w tym samym wierszu
        ws.cell(row=wiersz, column=4, value=f"=B{wiersz}*C{wiersz}")
        # Kumulacja - odwołanie do wiersza wyżej, z pierwszym wierszem bez sumy
        if wiersz == 2:
            ws.cell(row=wiersz, column=5, value=f"=D{wiersz}")
        else:
            ws.cell(row=wiersz, column=5, value=f"=E{wiersz - 1}+D{wiersz}")

    ostatni = 5
    ws.cell(row=6, column=1, value="RAZEM")
    ws.cell(row=6, column=2, value=f"=SUM(B2:B{ostatni})")
    ws.cell(row=6, column=3, value=f"=AVERAGE(C2:C{ostatni})")
    ws.cell(row=6, column=4, value=f"=SUM(D2:D{ostatni})")

    for adres in ("C2", "C3", "C4", "C5", "C6"):
        ws[adres].number_format = "0.0%"
    for adres in ("B2", "B6", "D2", "D3", "D4", "D5", "D6", "E2", "E3", "E4", "E5"):
        ws[adres].number_format = "#,##0.00"

    ws.column_dimensions["A"].width = 12
    for kolumna in "BCDE":
        ws.column_dimensions[kolumna].width = 14

    wb.save(cel)


def main() -> None:
    if not CEL.exists():
        zbuduj(CEL)
        print(f"Utworzono plik testowy: {CEL}")
        print("!! Otwórz go teraz w Excelu, ZAPISZ i zamknij, żeby powstał cache. !!\n")

    wb_f = load_workbook(CEL)                   # formuły
    wb_w = load_workbook(CEL, data_only=True)   # wartości z cache

    try:
        ws_f = wb_f["Sprzedaz"]
        ws_w = wb_w["Sprzedaz"]

        print("=" * 92)
        print("FORMUŁA  vs  WARTOŚĆ Z CACHE")
        print("=" * 92)
        print(f"  {'Adres':<7} {'Formuła (data_only=False)':<32} "
              f"{'Wartość (data_only=True)':<24} {'data_type'}")
        print("  " + "-" * 88)

        ile_formul = 0
        ile_z_cache = 0

        for wiersz in ws_f.iter_rows():
            for komorka in wiersz:
                if komorka.data_type != "f":
                    continue
                ile_formul += 1
                wartosc = ws_w[komorka.coordinate].value
                if wartosc is not None:
                    ile_z_cache += 1
                print(f"  {komorka.coordinate:<7} {komorka.value:<32} "
                      f"{str(wartosc):<24} {komorka.data_type!r}")

        print()
        print(f"  formuł: {ile_formul}, z cache: {ile_z_cache}")

        if ile_z_cache == 0:
            print()
            print("  DIAGNOZA: plik nie ma pamięci podręcznej.")
            print("  Powód: nikt go nie policzył (zapisał go openpyxl albo narzędzie,")
            print("         które nie liczy), ALBO openpyxl przepisał plik po drodze.")
            print("  Rozwiązanie: otwórz plik w Excelu i zapisz, albo użyj")
            print("               LibreOffice --headless (sekcja 2.15 kursu).")
        elif ile_z_cache < ile_formul:
            print()
            print("  DIAGNOZA: cache częściowy - część formuł ma wynik, część nie.")
            print("  To typowe dla plików składanych z kilku źródeł.")
        else:
            print()
            print("  DIAGNOZA: cache kompletny. Możesz ufać wartościom")
            print("  — pod warunkiem, że nikt nie zmienił danych po ostatnim przeliczeniu.")

        print()
        print("=" * 92)
        print("KONTROLA: CZY CACHE ZGADZA SIĘ Z DANYMI WEJŚCIOWYMI?")
        print("=" * 92)
        print("  Liczymy to samo w Pythonie i porównujemy z cache.")

        sprzedaz = [ws_w.cell(row=r, column=2).value for r in range(2, 6)]
        marze = [ws_w.cell(row=r, column=3).value for r in range(2, 6)]
        zyski_python = [s * m for s, m in zip(sprzedaz, marze)]

        print(f"  {'Wiersz':<8} {'Formuła (cache)':<20} {'Python':<20} {'Zgodne?'}")
        print("  " + "-" * 62)
        for indeks, wiersz in enumerate(range(2, 6)):
            z_excela = ws_w.cell(row=wiersz, column=4).value
            z_pythona = zyski_python[indeks]
            zgodne = (
                z_excela is not None
                and abs(z_excela - z_pythona) < 1e-9
            )
            print(f"  {wiersz:<8} {str(z_excela):<20} {z_pythona:<20.4f} "
                  f"{'OK' if zgodne else 'ROZBIEŻNOŚĆ'}")
    finally:
        wb_f.close()
        wb_w.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Dwa całkowicie niezależne grafy obiektów. W `wb_f` komórki mają `data_type='f'`, w `wb_w` te same komórki mają `data_type='n'` i liczby — bo openpyxl w tym trybie **nigdy nie widział formuł**. Pamięci zajmujesz dwa razy tyle.

**Co trafi do pliku.** Nic — to odczyt. Ale ostatnia sekcja skryptu pokazuje wzorzec, który warto zapamiętać na całe życie: **policz to samo niezależnie i porównaj.** To jedyny sposób, żeby wykryć nieaktualny cache.

### Przykład 3 — katastrofa: zapis po `data_only=True`

```python
"""DEMONSTRACJA ZNISZCZENIA: co się dzieje, gdy zapiszesz skoroszyt
wczytany z data_only=True. Uruchamiaj świadomie - pliki są w output/.

Uruchom:  python examples/06_katastrofa.py
"""

from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

ZRODLO = OUTPUT / "06_do_katastrofy.xlsx"
ZNISZCZONY = OUTPUT / "06_ZNISZCZONY.xlsx"


def przygotuj(cel: Path) -> None:
    """Tworzy plik z formułami. Uruchom, potem OTWÓRZ I ZAPISZ w Excelu."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Model"
    ws["A1"] = 100
    ws["A2"] = 200
    ws["A3"] = "=SUM(A1:A2)"
    ws["B1"] = "=A3*2"
    wb.save(cel)


def policz_formuly_z_dysku(sciezka: Path) -> int:
    """Ile formuł jest w pliku? Audyt pliku bez wczytywania do modelu na stałe."""
    wb = load_workbook(sciezka, read_only=True)
    try:
        return sum(
            1
            for ws in wb.worksheets
            for wiersz in ws.iter_rows()
            for komorka in wiersz
            if komorka.data_type == "f"
        )
    finally:
        wb.close()


def ile_elementow_f(sciezka: Path) -> int:
    """Liczy elementy <f> bezpośrednio w XML - dowód 'twardy', niezależny od openpyxl."""
    with ZipFile(sciezka) as archiwum:
        xml = archiwum.read("xl/worksheets/sheet1.xml").decode("utf-8")
    return xml.count("<f>")


def main() -> None:
    if not ZRODLO.exists():
        przygotuj(ZRODLO)
        print(f"Utworzono: {ZRODLO}")
        print("!! Otwórz w Excelu i ZAPISZ, żeby powstał cache. Potem uruchom ponownie. !!")
        return

    print("=" * 84)
    print("STAN PRZED KATASTROFĄ")
    print("=" * 84)
    print(f"  Formuły wg openpyxl (read_only):  {policz_formuly_z_dysku(ZRODLO)}")
    print(f"  Elementy <f> w XML:               {ile_elementow_f(ZRODLO)}")

    # --- TU DZIEJE SIĘ KATASTROFA ---
    wb = load_workbook(ZRODLO, data_only=True)
    try:
        ws = wb["Model"]
        print()
        print("  Model w pamięci (data_only=True):")
        for adres in ("A3", "B1"):
            k = ws[adres]
            print(f"    {adres}: value={k.value!r} data_type={k.data_type!r}")
        print("    ^ W modelu NIE MA ani jednej formuły. Są tylko liczby.")

        wb.save(ZNISZCZONY)
    finally:
        wb.close()

    print()
    print("=" * 84)
    print("STAN PO ZAPISIE")
    print("=" * 84)
    print(f"  Zapisano: {ZNISZCZONY}")
    print(f"  Formuły wg openpyxl (read_only):  {policz_formuly_z_dysku(ZNISZCZONY)}")
    print(f"  Elementy <f> w XML:               {ile_elementow_f(ZNISZCZONY)}")
    print()
    print("  ^ ZERO formuł. To nie jest błąd openpyxl - to konsekwencja modelu.")
    print("    Wczytałeś TYLKO wartości, więc zapisałeś TYLKO wartości.")
    print("    Żadnego ostrzeżenia nie było. Plik wygląda poprawnie.")

    print()
    print("=" * 84)
    print("JAK TEGO NIE ROBIĆ - POPRAWNY WZORZEC")
    print("=" * 84)
    print("""
  # 1. Rozpoznanie PRZED decyzją o wczytaniu
  liczba_formul = policz_formuly_z_dysku(ZRODLO)

  # 2. Potrzebujesz wartości -> OSOBNA sesja, tylko do odczytu
  wb_w = load_workbook(ZRODLO, data_only=True)
  try:
      wartosci = {k.coordinate: k.value for w in wb_w for k in w["Model"]}
  finally:
      wb_w.close()

  # 3. Potrzebujesz zmodyfikować i zapisać -> DRUGA sesja, tryb domyślny
  wb = load_workbook(ZRODLO)          # formuły zachowane
  try:
      wb["Model"]["C1"] = "dopisana kolumna"
      wb.save(OUTPUT / "nowa_nazwa.xlsx")     # ZAWSZE nowa nazwa
  finally:
      wb.close()
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** To najważniejszy moment w module. W trybie `data_only=True` openpyxl **nie widzi elementu `<f>` w ogóle** — pomija go przy budowaniu modelu. Nie ma więc czego „przywrócić" przy zapisie. Model jest prosty: „`A3` to liczba 300, `B1` to liczba 600".

**Co trafi do pliku.** Formuły zniknęły. W XML nie ma ani jednego `<f>`, są tylko `<v>`. Liczby pozostają **poprawne** — więc nikt nie zauważy problemu, dopóki nie spróbuje zmienić danych wejściowych.

**Dlaczego demonstracja liczy elementy `<f>` bezpośrednio w XML.** Bo to **twardy dowód**, niezależny od openpyxl. Gdybyś liczył tylko przez `data_type == 'f'` w trybie domyślnym, mógłbyś podejrzewać, że problem jest w bibliotece. Zajrzenie do pliku ZIP pokazuje, że problem jest w danych — dokładnie tak, jak w module 02 zapowiadaliśmy.

### Przykład 4 — `Translator`: kopiowanie formuły bez zepsucia referencji

```python
"""Kopiowanie formuł: dlaczego openpyxl nie tłumaczy referencji
i jak zrobić to świadomie.

Uruchom:  python examples/06_translator.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.formula.translate import Translator

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# Formuła z trzema rodzajami odwołań:
#   $A2  - stała kolumna (mnożnik z kolumny A), wiersz się przesuwa
#   B2   - w pełni względne
#   $D$1 - w pełni absolutne (kurs waluty)
WZORZEC = "=B2*$A2*$D$1"
ORIGIN = "C2"


def main() -> None:
    print("=" * 84)
    print("1. CO ROBI openpyxl, GDY PRZYPISUJESZ TEKST FORMULY")
    print("=" * 84)

    wb = Workbook()
    ws = wb.active
    ws.title = "Kopiowanie"

    ws["C2"] = WZORZEC
    ws["C3"] = ws["C2"].value          # "kopiowanie" przez openpyxl
    ws["D4"] = ws["C2"].value

    for adres in ("C2", "C3", "D4"):
        print(f"  {adres}: {ws[adres].value!r}")
    print()
    print("  ^ WSZYSTKIE TRZY SĄ IDENTYCZNE. To NIE jest błąd:")
    print("    openpyxl to notariusz, który przepisuje tekst dosłownie.")
    print("    Nie zna pojęcia 'kopiowanie komórki' i nie przesuwa referencji.")

    print()
    print("=" * 84)
    print("2. CO ZROBIŁBY EXCEL (i czego brakuje w podejściu powyżej)")
    print("=" * 84)
    print("  Excel przy wklejeniu C2 do C3 dałby: '=B3*$A3*$D$1'")
    print("  Zauważ: B i $A przesunęły się na wiersz 3, $D$1 zostało.")

    print()
    print("=" * 84)
    print("3. TRANSLATOR - ŚWIADOME PRZESUWANIE REFERENCJI")
    print("=" * 84)

    tlumacz = Translator(WZORZEC, origin=ORIGIN)
    print(f"  origin={ORIGIN}, wzorzec={WZORZEC!r}")
    print()
    for cel in ("C2", "C3", "C4", "C5", "D2", "F7"):
        przetlumaczona = tlumacz.translate_formula(cel)
        print(f"  -> {cel:<5} {przetlumaczona}")
    print()
    print("  ^ Referencje względne (B) przesuwają się.")
    print("    Referencje absolutne ($A, $D$1) zostają w miejscu.")
    print("    Dokładnie tak, jak zachowałby się Excel.")

    print()
    print("=" * 84)
    print("4. PRAKTYCZNE ZASTOSOWANIE: TABELA KRZYŻOWA")
    print("=" * 84)

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "Krzyzowa"

    ws2["A1"] = "Produkt \\ Rok"
    lata = [2024, 2025, 2026]
    ceny = [1000, 1200, 1350]

    for indeks, rok in enumerate(lata, start=2):
        kolumna = indeks                       # B, C, D
        ws2.cell(row=1, column=kolumna, value=rok)
        ws2.cell(row=2, column=kolumna, value=ceny[indeks - 2])

    produkty = ["Widzet", "Wkręt", "Nakrętka"]
    wspolczynniki = [1.0, 1.5, 0.8]

    for indeks, produkt in enumerate(produkty, start=3):
        ws2.cell(row=indeks, column=1, value=produkt)
        ws2.cell(row=indeks, column=6, value=wspolczynniki[indeks - 3])

    # Wzorzec: cena z wiersza 2 (absolutny wiersz) * współczynnik z kolumny F (absolutna kolumna)
    wzorzec = "=B$2*$F3"
    tlumacz_krzyzowy = Translator(wzorzec, origin="B3")

    for wiersz in range(3, 6):
        for kolumna in range(2, 5):
            from openpyxl.utils import get_column_letter
            adres = f"{get_column_letter(kolumna)}{wiersz}"
            ws2[adres] = tlumacz_krzyzowy.translate_formula(adres)
            ws2[adres].number_format = "#,##0.00"

    cel = OUTPUT / "06_translator.xlsx"
    wb2.save(cel)

    print(f"  Zapisano: {cel}")
    print()
    for wiersz in ws2.iter_rows(min_row=1, max_row=5, max_col=6, values_only=True):
        print("  " + " | ".join("" if v is None else str(v) for v in wiersz))
    print()
    print("  ^ Każda komórka ma WŁASNĄ, poprawnie przesuniętą formułę.")
    print("    Referencje 'B$2' (stały wiersz) i '$F3' (stała kolumna)")
    print("    zadziałały jak w Excelu.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `Translator` przy tworzeniu **tokenizuje** formułę — dzieli ją na kawałki (operandy, operatory, funkcje, separatory) i rozpoznaje, które kawałki są referencjami. Przy `translate_formula(dest)` przelicza różnicę wierszy i kolumn między `origin` a `dest` i przesuwa tylko te referencje, które są **względne w danym wymiarze**.

**Co trafi do pliku.** Komórki z gotowymi formułami — już przesuniętymi. Excel nie musi nic tłumaczyć.

**Kiedy używać `Translator`, a kiedy f-stringa.** Przykład 4 jest dydaktyczny: pokazuje, że to narzędzie istnieje i działa. Ale w praktyce **f-string jest zwykle lepszy**, bo czytelniejszy:

```python
# Translator - trzeba pamiętać o origin i o tym, że referencje są w notacji A1
wzorzec = Translator("=B$2*$F3", origin="B3")
ws[adres] = wzorzec.translate_formula(adres)

# f-string - widać dokładnie, co powstaje
ws[adres] = f"={get_column_letter(kolumna)}$2*$F{wiersz}"
```

`Translator` wygrywa w dwóch sytuacjach: gdy formuła jest **długa i skomplikowana** (przepisywanie jej f-stringiem to koszmar i źródło błędów) oraz gdy formuła **pochodzi z zewnątrz** — na przykład z pliku szablonu, z konfiguracji, od użytkownika. Wtedy `Translator` pozwala ją skopiować bez ręcznego rozbierania na części.

### Przykład 5 — formuły międzyarkuszowe i nazwy zdefiniowane

```python
"""Formuły sięgające innych arkuszy i nazw zdefiniowanych.
Trzy pułapki, które je łamią.

Uruchom:  python examples/06_miedzy_arkuszami.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import absolute_coordinate, quote_sheetname
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def main() -> None:
    wb = Workbook()

    # --- arkusz z danymi (nazwa ZE SPACJĄ - celowo) ---
    dane = wb.active
    dane.title = "Dane 2026"
    dane.append(["Miesiąc", "Sprzedaż", "Koszt"])
    for miesiac, sprzedaz, koszt in [
        ("styczeń", 10_000, 8_200),
        ("luty", 12_500, 10_100),
        ("marzec", 9_800, 9_100),
    ]:
        dane.append([miesiac, sprzedaz, koszt])

    # --- arkusz podsumowania ---
    podsumowanie = wb.create_sheet("Podsumowanie")

    # 1. Formuła z apostrofami (nazwa ze spacją) i odwołaniem absolutnym
    podsumowanie["A1"] = "Suma sprzedaży (formuła międzyarkuszowa)"
    podsumowanie["B1"] = (
        f"=SUM({quote_sheetname('Dane 2026')}!$B$2:$B$4)"
    )

    # 2. To samo, ale zbudowane ręcznie - dla porównania
    podsumowanie["A2"] = "Suma kosztów (zapis ręczny)"
    podsumowanie["B2"] = "='Dane 2026'!$C$2+$C$3+$C$4"

    # 3. Formuła z nazwy zdefiniowanej
    podsumowanie["A3"] = "Średnia marża (nazwa zdefiniowana)"
    podsumowanie["B3"] = "=SREDNIA_MARZA"

    # --- tworzymy nazwy zdefiniowane ---
    odniesienie_sprzedaz = (
        f"{quote_sheetname(dane.title)}!{absolute_coordinate('B2:B4')}"
    )
    wb.defined_names.add(DefinedName(name="SPRZEDAZ_MIESIECZNA",
                                     attr_text=odniesienie_sprzedaz))

    # Nazwa wskazująca na FORMULĘ, nie na komórkę
    marza_formula = f"AVERAGE({quote_sheetname(dane.title)}!$B$2:$B$4)"
    wb.defined_names.add(DefinedName(name="SREDNIA_MARZA", attr_text=marza_formula))

    # --- użycie nazwy w formule ---
    podsumowanie["A4"] = "Suma przez nazwę"
    podsumowanie["B4"] = "=SUM(SPRZEDAZ_MIESIECZNA)"

    podsumowanie["B1"].number_format = "#,##0"
    podsumowanie["B2"].number_format = "#,##0"
    podsumowanie["B3"].number_format = "#,##0"
    podsumowanie["B4"].number_format = "#,##0"

    cel = OUTPUT / "06_miedzy_arkuszami.xlsx"
    wb.save(cel)
    print(f"Zapisano: {cel}")

    print()
    print("=" * 84)
    print("ZAWARTOŚĆ FORMUL")
    print("=" * 84)
    kontrola = load_workbook(cel)
    try:
        ws = kontrola["Podsumowanie"]
        for wiersz in ws.iter_rows(min_row=1, max_row=4):
            etykieta = wiersz[0].value
            komorka = wiersz[1]
            print(f"  {etykieta:<40} {komorka.value}")

        print()
        print("  Nazwy zdefiniowane w pliku:")
        for nazwa in kontrola.defined_names:
            definicja = kontrola.defined_names[nazwa]
            print(f"    {nazwa:<24} -> {definicja.attr_text}")
    finally:
        kontrola.close()

    print()
    print("=" * 84)
    print("TRZY PUŁAPKI, KTÓRE ŁAMIĄ TAKIE FORMULY")
    print("=" * 84)
    print("""
  1. ZMIANA NAZWY ARKUSZA
     Jeśli zmienisz ws.title z "Dane 2026" na "Dane Q4", formuła
     ='Dane 2026'!$B$2 NIE zostanie zaktualizowana. Excel pokaże #REF!.
     openpyxl nie ma grafu zależności - pisze tekst, nie śledzi powiązań.
     Wniosek: nazwę arkusza ustal RAZ i traktuj ją jak stałą publiczną.

  2. USUNIĘCIE ARKUSZA
     del wb["Dane 2026"] pozostawi formuły wskazujące donikąd.
     Analogia: wyrwałeś stronę z indeksem, odwołania w tekście zostały.

  3. LITERÓWKA W NAZWIE
     openpyxl NIE sprawdza, czy arkusz o takiej nazwie istnieje w chwili
     zapisu. Formuła ='Dane 2027'!$B$2 przejdzie przez cały pipeline
     i objawi się dopiero w Excelu.
     Wniosek: waliduj nazwy arkuszy przed budowaniem formuł.
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `DefinedName` to zwykły obiekt z dwoma polami: `name` i `attr_text`. `wb.defined_names.add()` wstawia go do słownika skoroszytu. **Nic nie jest sprawdzane** — ani czy arkusz istnieje, ani czy komórka istnieje, ani czy nazwa nie ma spacji.

**Co trafi do pliku.** W `xl/workbook.xml` pojawi się sekcja `<definedNames>` z wpisami `<definedName name="...">treść</definedName>`. A formuły używające tych nazw — dokładnie tak, jak je napisałeś.

**Zwróć uwagę na `B3`.** To nie jest odwołanie do komórki, a **nazwa wskazująca na formułę** `AVERAGE('Dane 2026'!$B$2:$B$4)`. W Excelu komórka pokaże wynik. Ale w pliku nie zobaczysz, skąd ten wynik pochodzi — i to jest cena za wygodę. **Dla wartości biznesowych używaj nazw wskazujących na komórki.** Nazwy-formuły mają sens tylko dla stałych konfiguracyjnych (np. `PODATEK = 0.23*1.05`).

### Przykład 6 — `ArrayFormula` i formuły dynamiczne obok siebie

```python
"""Formuła tablicowa (ArrayFormula) i formuła dynamiczna (Excel 365).
Dwie drogi do tego samego wyniku - i to, co openpyxl modeluje, a czego nie.

Uruchom:  python examples/06_tablicowe.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Tablica"

    ws["A1"] = "Ilość"
    ws["B1"] = "Cena"
    for i in range(1, 11):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=round(i * 1.37, 2))

    # --- 1. FORMULA TABLICOWA (klasyczna CSE) ---
    ws["D1"] = "Iloczyn (ArrayFormula)"
    ws["D2"] = ArrayFormula(ref="D2:D11", text="=A2:A11*B2:B11")
    print("ArrayFormula zapisana w D2, obejmuje D2:D11")

    # --- 2. FORMULA ZWYKŁA SUMUJĄCA ten sam wynik ---
    ws["E1"] = "Suma iloczynów (zwykła formuła)"
    ws["E2"] = "=SUMPRODUCT(A2:A11,B2:B11)"

    # --- 3. FORMULA DYNAMICZNA (wymaga Excel 365) ---
    ws["G1"] = "Unikalne (formuła dynamiczna)"
    ws["G2"] = "=_xlfn.UNIQUE(A2:A11)"
    print("Formuła dynamiczna w G2 z prefiksem _xlfn. - sprawdź go w swoim Excelu!")

    # --- 4. TO SAMO, ALE POLICZONE W PYTHONIE (zawsze działa) ---
    ws["I1"] = "Unikalne (wartości z Pythona)"
    wartosci = [ws.cell(row=r, column=1).value for r in range(2, 12)]
    unikalne = sorted(set(wartosci))
    for i, wartosc in enumerate(unikalne, start=2):
        ws.cell(row=i, column=9, value=wartosc)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 26
    ws.column_dimensions["G"].width = 26
    ws.column_dimensions["I"].width = 26

    cel = OUTPUT / "06_tablicowe.xlsx"
    wb.save(cel)
    print(f"\nZapisano: {cel}")

    print()
    print("=" * 84)
    print("CO WIDAĆ PO PONOWNYM WCZYTANIU")
    print("=" * 84)
    kontrola = load_workbook(cel)
    try:
        ws2 = kontrola["Tablica"]
        for adres in ("D2", "E2", "G2", "I2"):
            k = ws2[adres]
            typ = type(k.value).__name__
            if isinstance(k.value, ArrayFormula):
                opis = f"ArrayFormula(text={k.value.text!r}, ref={k.value.ref!r})"
            else:
                opis = repr(k.value)
            print(f"  {adres}: {typ:<14} {opis}")
        print()
        print("  ^ D2 to NIE jest str - to obiekt ArrayFormula.")
        print("    Proste k.value.startswith('=') wywali się na tym z AttributeError.")
        print("    Sprawdź reprezentację w dokumentacji swojej wersji openpyxl.")

        print()
        print("  Odczyt w trybie data_only=True (cache):")
        kontrola.close()
        kontrola = load_workbook(cel, data_only=True)
        ws3 = kontrola["Tablica"]
        for adres in ("D2", "E2", "G2", "I2"):
            k = ws3[adres]
            print(f"    {adres}: {k.value!r}")
        print()
        print("    ^ Wszystkie None - nikt tego pliku nie policzył.")
        print("      Ale I2..I6 (wartości z Pythona) mają dane, bo to NIE formuły.")
    finally:
        kontrola.close()

    print()
    print("=" * 84)
    print("CZTERY PODEJŚCIA DO TEGO SAMEGO PROBLEMU")
    print("=" * 84)
    print("""
  D/E:  formuła tablicowa + SUMPRODUCT
        -> działa w każdym Excelu od 2007, ale użytkownik nie może edytować
           komórek objętych ArrayFormula (Excel tego zabrania).

  G:    formuła dynamiczna UNIQUE z prefiksem _xlfn.
        -> elegancka, ale WYMAGA Excel 365/2021+ i metadanych rozlania,
           których openpyxl NIE generuje. Sprawdź na swojej wersji.

  I:    wartości policzone w Pythonie.
        -> działa ZAWSZE, w każdej wersji Excela, w każdym odbiorcy
           (w tym w programach, które nie liczą). To nie jest "gorsze".
           To jest właściwy wybór, jeśli plik nie ma być narzędziem.

  Wniosek: wybierz świadomie. "Najnowsze" nie znaczy "najlepsze" -
  znaczy "najmniej przenośne".
""")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `ArrayFormula` to obiekt, nie tekst. To pierwszy moment w kursie, w którym wartość komórki **nie jest** jednym z typów z modułu 05 — i dlatego trzeba o tym wiedzieć: `cell.value` może zwrócić coś, czego się nie spodziewasz.

**Co trafi do pliku.** Excel zapisze formułę tablicową jako `<f t="array" ref="D2:D11">`. Formuła dynamiczna zostanie zapisana jako `<f>_xlfn.UNIQUE(A2:A11)</f>` — z prefiksem, o ile go podasz.

**Czego ten przykład uczy najbardziej:** kolumna `I` z wartościami policzonymi w Pythonie **nie wymaga niczyjej łaski**. Nie potrzebuje Excela 365, nie potrzebuje przeliczenia, nie potrzebuje metadanych. Działa. W module 26 nazwiemy to „obciążeniem poznawczym przeniesionym tam, gdzie jest kontrolowane".

### Przykład 7 — wymuszanie przeliczenia i liczenie bez Excela

```python
"""Dwie drogi do wartości z formuł bez otwierania Excela.

Uruchom:  python examples/06_przeliczanie.py
Uruchom:  bash:  soffice --headless --convert-to xlsx --outdir output/przel output/06_formuly_przeliczone.xlsx
"""

import shutil
import subprocess
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def wymus_przeliczenie(wb) -> None:
    """Ustawia flagę pełnego przeliczenia przy otwarciu.

    Uwaga: to nie sprawia, że openpyxl liczy - to prośba do programu,
    który otworzy plik. Sprawdź nazwy atrybutów w swojej wersji.
    """
    if wb.calculation is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Model"

    ws["A1"] = "Cena netto"
    ws["B1"] = 100.0
    ws["A2"] = "VAT"
    ws["B2"] = 0.23
    ws["A3"] = "Cena brutto"
    ws["B3"] = "=B1*(1+B2)"
    ws["A4"] = "Podatek"
    ws["B4"] = "=B1*B2"

    ws["B3"].number_format = '#,##0.00 "zł"'
    ws["B4"].number_format = '#,##0.00 "zł"'
    ws["B2"].number_format = "0%"

    wymus_przeliczenie(wb)
    wb.properties.description = (
        "Wygenerowane przez openpyxl. Ustawiono fullCalcOnLoad=True, "
        "więc Excel przeliczy formuły przy otwarciu."
    )
    wb.save(cel)


def main() -> None:
    cel = OUTPUT / "06_formuly_przeliczone.xlsx"
    zbuduj(cel)
    print(f"Zapisano: {cel}")
    print("  - w xl/workbook.xml jest teraz <calcPr ... fullCalcOnLoad=\"1\"/>")
    print("  - Excel przeliczy formuły przy otwarciu")

    print()
    print("=" * 84)
    print("CZYM TO NIE JEST")
    print("=" * 84)
    wb = load_workbook(cel, data_only=True)
    try:
        print(f"  data_only=True na tym pliku: B3 = {wb['Model']['B3'].value!r}")
        print("  ^ Nadal None! fullCalcOnLoad mówi EXCELOWI, żeby policzył.")
        print("    Python czyta plik tak, jak został zapisany - bez obliczeń.")
    finally:
        wb.close()

    print()
    print("=" * 84)
    print("LIBREOFFICE W TRYBIE HEADLESS (opcjonalnie)")
    print("=" * 84)

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        print("  Nie znaleziono LibreOffice. Instalacja (przykłady):")
        print("    Debian/Ubuntu:  sudo apt install libreoffice-calc")
        print("    macOS:          brew install --cask libreoffice")
        print("    Windows:        instalator z strony LibreOffice")
        print()
        print("  Po instalacji uruchom:")
        print(f"    soffice --headless --convert-to xlsx "
              f"--outdir {OUTPUT / 'przel'} {cel}")
        return

    katalog = OUTPUT / "przel"
    katalog.mkdir(exist_ok=True)
    print(f"  Znaleziono: {soffice}")
    print(f"  Konwertuję do: {katalog}")

    wynik = subprocess.run(
        [soffice, "--headless", "--convert-to", "xlsx",
         "--outdir", str(katalog), str(cel)],
        capture_output=True, text=True, timeout=180,
    )
    print(f"  zwrot: {wynik.returncode}")
    if wynik.stdout.strip():
        print(f"  stdout: {wynik.stdout.strip()[:400]}")
    if wynik.stderr.strip():
        print(f"  stderr: {wynik.stderr.strip()[:400]}")

    przekonwertowany = katalog / cel.name
    if not przekonwertowany.exists():
        print("  Nie udało się utworzyć pliku wynikowego.")
        return

    wb = load_workbook(przekonwertowany, data_only=True)
    try:
        ws = wb["Model"]
        print()
        print("  Wartości po przeliczeniu przez LibreOffice:")
        for adres in ("B1", "B2", "B3", "B4"):
            print(f"    {adres} = {ws[adres].value!r}")
        if ws["B3"].value is None:
            print()
            print("  Nadal None -> LibreOffice nie przeliczył przy wczytywaniu.")
            print("  To ustawienie (Narzędzia -> Opcje -> Calc -> Formuła)")
            print("  różni się między wersjami i konfiguracjami.")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `wymus_przeliczenie()` zapisuje jeden atrybut w obiekcie `CalcProperties` — nic więcej. Cała „magia" dzieje się dopiero po stronie programu, który otworzy plik.

**Co trafi do pliku.** Element `<calcPr calcId="..." fullCalcOnLoad="1"/>` w `xl/workbook.xml`. Zajrzyj tam sam po uruchomieniu — to dobra okazja, żeby zobaczyć, że „ustawienie w Excelu" to konkretny kawałek XML.

**Czego ten przykład uczy najbardziej:** że **`fullCalcOnLoad` i `data_only=True` to dwa różne światy**. Pierwsze mówi Excelowi, żeby policzył. Drugie mówi openpyxl, żeby przeczytał to, co już policzone. Jeśli nie zrozumiesz tego rozróżnienia, będziesz w kółko próbować „zmusić openpyxl do liczenia" — czego nie da się zrobić, bo tam nie ma czego zmuszać.

## 4. Anatomia API

| Metoda / klasa / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `ws["C1"] = "=..."` | Zapisuje formułę | tekst zaczynający się od `=` | openpyxl **nie waliduje** składni ani nazw funkcji |
| `cell.value` (przy formule) | Zwraca literał formuły albo obiekt formuły | — | `str` (`'=SUM(...)'`) albo `ArrayFormula` — sprawdzaj typ! |
| `cell.data_type` | `'f'` dla formuły | — | W trybie `data_only=True` **nigdy** nie będzie `'f'` |
| `load_workbook(data_only=True)` | Czyta **cache** (`<v>`), ignoruje `<f>` | `True`/`False` (domyślnie `False`) | **Nigdy nie łącz z `save()`** na tym samym obiekcie |
| `load_workbook(data_only=False)` | Czyta **literały formuł** | domyślne | Cache jest ignorowany |
| `load_workbook(read_only=True)` | Szybki odczyt bez zapisu do modelu na stałe | — | Świetne do audytu formuł; `wb.close()` obowiązkowo |
| `wb.calculation` | Obiekt `CalcProperties` (element `calcPr`) | — | Może być `None`; sprawdź wersję |
| `CalcProperties(fullCalcOnLoad=True)` | Każę Excelowi przeliczyć wszystko przy otwarciu | — | `openpyxl.workbook.properties`; nie działa na openpyxl |
| `Translator(formula, origin)` | Tokenizuje formułę i przygotowuje do tłumaczenia | `origin` = adres, w którym formuła leży | `openpyxl.formula.translate` |
| `Translator.translate_formula(dest)` | Zwraca formułę z przesuniętymi referencjami | adres docelowy | Referencje z `$` zostają w miejscu |
| `Translator` — metody przesuwania | Przesuwanie formuł przy wstawianiu/usuwaniu wierszy i kolumn | — | Sprawdź nazwy i dostępność w swojej wersji |
| `ArrayFormula(ref, text)` | Formuła tablicowa (klasyczna CSE) | `ref` = zakres, `text` = formuła | `openpyxl.worksheet.formula`; przypisz **raz**, do lewej-górnej |
| `ArrayFormula.text`, `.ref` | Odczyt przy wczytaniu pliku | — | `text` często **bez** znaku `=` |
| `openpyxl.worksheet.formula` — inne klasy | Reprezentacje innych wariantów formuł w pliku | — | Sprawdź w dokumentacji swojej wersji |
| `Tokenizer(formula)` | Rozbija formułę na tokeny | — | `openpyxl.formula`; `Tokenizer("=SUM(A1,B1)").items` |
| `token.type`, `token.subtype`, `token.value` | Opis pojedynczego tokenu | — | `type`: `OPERAND`, `FUNC`, `OP_IN`, `OP_PRE`, `OP_POST`, `SEP`, `WSPACE`, `ARRAY` |
| `Tokenizer.render()` | Składa tokeny z powrotem w tekst | — | Przydatne przy modyfikacji formuł |
| `quote_sheetname(nazwa)` | `'Dane 2026'` — dodaje apostrofy | — | `openpyxl.utils`; wymagane przy spacji w nazwie |
| `absolute_coordinate("B2")` | `'$B$2'` | — | `openpyxl.utils` |
| `wb.defined_names.add(DefinedName(...))` | Dodaje nazwę zdefiniowaną | `name`, `attr_text` | API zmienione w 3.1.x |
| `wb.defined_names["NAZWA"]` | Odczyt nazwy | — | `.attr_text` → treść, `.destinations` → miejsca |
| `DefinedName(..., localSheetId=0)` | Nazwa o zasięgu lokalnym dla arkusza | indeks arkusza od 0 | Przesłania nazwę globalną o tej samej nazwie |
| `wb._external_links` | Lista odwołań do zewnętrznych skoroszytów | — | Atrybut wewnętrzny; sprawdź w swojej wersji |
| `load_workbook(..., keep_links=True)` | Zachowanie odwołań zewnętrznych | `True` domyślnie | `False` może spowodować `#REF!` przy odwołaniach zewnętrznych |
| `ws.move_range(zakres, rows, cols, translate=)` | Przenosi zakres i (opcjonalnie) tłumaczy formuły | `translate=True/False` | Bez `translate` formuły **nie** są przesuwane; moduł 07 |
| `ws.insert_rows()`, `ws.delete_rows()` | Wstawia / usuwa wiersze | — | **NIE aktualizują formuł**; moduł 07 |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — suma i średnia, czyli pierwsza formuła.** Napisz `examples/06_cw1.py`, który:

1. Tworzy skoroszyt z arkuszem `Liczby`.
2. Wpisuje do `A1:A10` dziesięć liczb (dowolnych, np. `1..10`).
3. Wpisuje w `C1` formułę `=SUM(A1:A10)`, a w `C2` — `=AVERAGE(A1:A10)`.
4. Ustawia w `C1` i `C2` format `#,##0.00`.
5. Zapisuje plik jako `output/06_cw1.xlsx`.
6. Wypisuje w konsoli swoje **przewidywane** wartości dla `C1` i `C2`.
7. Wczytuje plik **trzykrotnie**: raz domyślnie, raz z `data_only=True`, raz z `read_only=True` — i wypisuje, co dostał w `C1` i `C2` w każdym trybie.
8. Zagląda do `xl/worksheets/sheet1.xml` w pliku ZIP i wypisuje fragmenty z komórkami `C1` i `C2`.

Po uruchomieniu **otwórz plik w Excelu**, sprawdź, czy przewidywania się zgadzają, zapisz plik i uruchom skrypt ponownie.

W pliku `output/06_cw1_wnioski.md` uzupełnij tabelę i odpowiedz:

| Tryb odczytu | `C1` (`=SUM`) | `C2` (`=AVERAGE`) |
|---|---|---|
| domyślny (`data_only=False`) | | |
| `data_only=True` **przed** otwarciem w Excelu | | |
| `data_only=True` **po** otwarciu i zapisaniu w Excelu | | |
| `read_only=True` | | |

- Co dokładnie znajduje się w XML dla komórki `C1` przed otwarciem w Excelu, a co po? Wypisz oba fragmenty.
- Co się zmieniło w pliku po tym, jak otworzyłeś go w Excelu i zapisałeś? (Podpowiedź: użyj `zipfile` i porównaj rozmiary oraz zawartość `sheet1.xml` przed i po.)
- Dlaczego `read_only=True` zwraca to samo co tryb domyślny, a nie to samo co `data_only=True`?

**Zadanie 2 — formuła jako tekst.** Zmodyfikuj `06_cw1.py` tak, aby:

1. W komórce `E1` zapisać **tekst** `"=SUM(A1:A10)"` — tak, żeby Excel pokazał go jako napis, a nie policzył (użyj triku z `data_type = "s"`).
2. W komórce `E2` zapisać ten sam **tekst** bez żadnego triku (czyli `data_type` pozostanie `'f'`).
3. W komórce `E3` zapisać tekst **bez** znaku `=` — `"SUM(A1:A10)"`.
4. Zapisać, wczytać i wypisać dla `E1`, `E2`, `E3`: `value` i `data_type`.
5. Otworzyć plik w Excelu i **opisać słowami**, czym różnią się wyglądy tych trzech komórek (wyrównanie, zawartość, ewentualny wynik).

W pliku `output/06_cw2_wnioski.md` zapisz odpowiedzi, w tym:

- Który z trzech przypadków **wygląda w kodzie identycznie**, a zachowuje się zupełnie inaczej?
- Dlaczego `E3` (bez `=`) jest najczęstszym błędem początkujących?
- W której sytuacji potrzebujesz `E1` (tekst udający formułę) i jak to się wiąże z bezpieczeństwem (moduł 21)?

### 🟡 Warsztat

**Zadanie 3 — audyt formuł w cudzym pliku.** Napisz `examples/06_cw3.py`, który dla dowolnego pliku `.xlsx` podanego w argumencie (a jeśli nie podano — dla `output/06_dwa_swiaty.xlsx`) wykonuje pełny audyt:

1. **Znajdź wszystkie formuły.** Wypisz, ile ich jest, w których arkuszach i w których kolumnach (pogrupuj po kolumnie — to zwykle najbardziej diagnostyczny widok).
2. **Sklasyfikuj formuły** według funkcji: wyodrębnij nazwy funkcji z każdej formuły (np. `SUM`, `IF`, `VLOOKUP`) i wypisz ranking: która funkcja występuje najczęściej. Wykorzystaj `openpyxl.formula.Tokenizer` — tokeny o `type == "FUNC"` to nazwy funkcji.
3. **Sprawdź cache.** Wczytaj plik drugi raz z `data_only=True` i wypisz:
   - liczbę formuł, które **mają** wartość w cache,
   - liczbę formuł, które **nie mają** (czyli `None`),
   - procent pokrycia cache.
4. **Wykryj formuły „podejrzane"** i wypisz je z adresem i uzasadnieniem:
   - puste (`"="` albo `"=+"`),
   - niemające zamykającego nawiasu (proste sprawdzenie: liczba `(` ≠ liczba `)`),
   - zawierające `#REF!` w treści,
   - odwołujące się do arkusza, którego nie ma w pliku (znajdź w formule fragment `'Nazwa'!` albo `Nazwa!` i sprawdź, czy `Nazwa` istnieje w `wb.sheetnames`),
   - dłuższe niż 1000 znaków (limitowo 8192, ale już 1000 to sygnał ostrzegawczy).
5. **Sprawdź, czy formuły odwołują się do komórek poza zakresem danych** — na przykład w arkuszu, gdzie `max_row` to 50, formuła sięga wiersza 1000. To klasyczny znak formuł „na zapas", które ktoś wstawił i zapomniał.
6. **Wypisz raport** jako tekstową tabelkę w konsoli oraz — bonus — zapisz ten sam raport do **nowego arkusza** `Audyt` w kopii pliku.

Wymagania techniczne:

- Plik **nie może być modyfikowany** w miejscu. Audyt ma być bezpieczny: czytaj `read_only=True` tam, gdzie to możliwe.
- Kod ma być odporny na duże pliki: nie buduj listy wszystkich komórek, jeśli plik ma 500 000 wierszy — dla `read_only` używaj `iter_rows()` z ograniczeniem do arkusza, w którym są formuły.
- Narzędzie ma działać na dowolnym pliku, nie tylko na tych z kursu.

W pliku `output/06_cw3_wnioski.md`:

- Wklej raport z co najmniej trzech różnych plików (w tym jednego wygenerowanego przez openpyxl i jednego utworzonego w Excelu).
- Odpowiedz: **jakie dwa sygnały w raporcie najszybciej wskazują, że plik był generowany automatycznie, a nie tworzony przez człowieka?**
- Odpowiedz: dlaczego `Tokenizer` jest lepszym sposobem wyodrębniania nazw funkcji niż wyrażenie regularne? Podaj jeden konkretny przypadek, w którym wyrażenie regularne dałoby zły wynik.

### 🔴 Wyzwanie

**Zadanie 4 — podsumowanie formułą, walidacja w Pythonie.** To zadanie jest zalążkiem modułu 26: rozdzielenia roli „Excel jako narzędzie" i „Python jako silnik".

Napisz `examples/06_cw4.py`, który generuje skoroszyt `output/06_cw4_raport.xlsx` opisany poniżej.

**Arkusz `Dane` (surowe dane wejściowe):** 200 wierszy wygenerowanych programowo, kolumny:

| Lp. | Region | Produkt | Ilość | Cena | Sprzedaż |
|---|---|---|---|---|---|
| 1..200 | jedno z: `Północ`, `Południe`, `Wschód`, `Zachód` | jedno z 5 produktów | 1..50 | jedna z 5 cen | **formuła** `=D2*E2` |

Kolumny `Sprzedaż` **muszą być formułami** (`=D2*E2`, `=D3*E3`, ...) — tu celowo używamy formuł, bo użytkownik może chcieć poprawić ilość albo cenę i zobaczyć nową wartość.

**Arkusz `Podsumowanie` (żywych formuł):** tabela przestawna zbudowana formułami:

- wiersze: cztery regiony plus `RAZEM`,
- kolumny: `Suma sprzedaży`, `Średnia sprzedaż`, `Liczba transakcji`, `Średnia cena`,
- każda liczba to formuła:
  - `=SUMIF(Dane!$B$2:$B$201,$A2,Dane!$F$2:$F$201)` dla sumy sprzedaży regionu,
  - `=AVERAGEIF(Dane!$B$2:$B$201,$A2,Dane!$F$2:$F$201)` dla średniej,
  - `=COUNTIF(Dane!$B$2:$B$201,$A2)` dla liczby transakcji,
  - `=AVERAGEIF(Dane!$B$2:$B$201,$A2,Dane!$E$2:$E$201)` dla średniej ceny,
  - wiersz `RAZEM` używa `SUM`, `AVERAGE`, `COUNTA` po całym zakresie.
- formatowanie: kwoty jako `#,##0.00 "zł"`, ceny jako `#,##0.00`, liczba transakcji jako `0`, średnie z jednym miejscem po przecinku.

**Arkusz `Kontrola` (wartości z Pythona):** **ten sam** zestaw agregatów policzony w Pythonie i wpisany jako **wartości**, nie formuły:

- te same nagłówki i ten sam układ co `Podsumowanie`,
- liczby w kolorze ciemnoszarym z komentarzem w `A1`: „Wartości policzone w Pythonie. Porównaj z arkuszem Podsumowanie, jeśli podejrzewasz, że dane zmieniły się po wygenerowaniu."
- dodatkowa kolumna `ZGODNE` z formułą `=IF(ROUND(B2,6)=ROUND(Podsumowanie!B2,6),"OK","SPRAWDŹ")` — tak, tu **celowo** używamy formuły, bo jej zadaniem jest porównanie dwóch światów.

**Arkusz `Metodyka`:** tekstowy opis (kilka komórek), który zawiera odpowiedzi na pytania:

1. Które elementy raportu są formułami i **dlaczego**?
2. Które elementy są wartościami i **dlaczego**?
3. Co się stanie z arkuszem `Podsumowanie`, jeśli osoba otwierająca plik nie ma Excela 365?
4. Co się stanie z arkuszem `Podsumowanie`, jeśli inny program odczyta plik bez uprzedniego otwarcia go w Excelu?
5. Gdzie w tym raporcie leży logika biznesowa — w formule czy w Pythonie? Uzasadnij w 3–5 zdaniach.

**Arkusz `Ustawienia`:** konfiguracja raportu zapisana jako dane (a nie jako kod):

- wersja raportu, data wygenerowania, autor, nazwy arkuszy, progi walidacji (np. minimalna akceptowalna sprzedaż).
- **Uwaga:** to zalążek deklaratywnej konfiguracji z modułu 26 — ale jeszcze w prostszej wersji: po prostu dane w komórkach z nagłówkami obok.

**Walidacja po stronie Pythona (ma być widoczna w pliku):**

1. Wykryj wiersze z błędami danych wejściowych: ujemna ilość, cena poza dozwolonym zakresem, brak regionalu, tekst w kolumnie liczbowej.
2. W arkuszu `Dane` wypełnij **całe błędne wiersze** kolorem jasnoczerwonym (`PatternFill`), a w dodatkowej kolumnie `Problem` (obok `Sprzedaż`) wpisz **tekstowy** opis problemu — nie formułę.
3. W arkuszu `Kontrola` dopisz blok `Walidacja` z liczbą wykrytych problemów według kategorii — jako wartości.
4. Ustaw `fullCalcOnLoad=True` na skoroszycie i wyjaśnij w `Metodyka`, dlaczego.

**Wymagania dodatkowe:**

- Nazwy kolumn i zakresy formuł **muszą być budowane programowo** (f-stringiem albo przy pomocy `get_column_letter`), a nie wpisane na sztywno jako `"F201"`.
- Liczba wierszy danych jest parametrem funkcji (`liczba_wierszy=200`), ale liczba **kolumn i ich kolejność** są stałe.
- Skrypt musi być uruchamialny wielokrotnie — za każdym razem generuje plik **od nowa** i jest **deterministyczny**: te same dane wejściowe dają bajtowo tę samą treść (pomijając datę w `Ustawienia`). Użyj stałego ziarna (`random.seed(42)`).
- Skrypt wypisuje na końcu **raport z porównania**: dla każdego regionu — wartość z Pythona, oraz informację, czy formuła `ZGODNE` będzie mogła potwierdzić zgodność (bez otwierania Excela nie możesz tego sprawdzić, ale możesz policzyć to samo w Pythonie i wypisać oczekiwany wynik).

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""Podsumowanie formułą + walidacja w Pythonie - rozwiązanie zadania 4.

Uruchom:  python examples/06_cw4.py
"""

from __future__ import annotations

import random
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# ----------------------------------------------------------------------
# KONFIGURACJA RAPORTU - jedno miejsce, z którego wszystko wynika
# ----------------------------------------------------------------------
REGIONY = ["Północ", "Południe", "Wschód", "Zachód"]
PRODUKTY = ["Widzet", "Wkręt", "Nakrętka", "Podkładka", "Śruba"]
CENY = {"Widzet": 12.50, "Wkręt": 3.20, "Nakrętka": 1.80,
        "Podkładka": 0.45, "Śruba": 7.90}

KOLUMNY_DANE = ["Lp.", "Region", "Produkt", "Ilość", "Cena", "Sprzedaż", "Problem"]
KOLUMNY_PODSUMOWANIA = ["Region", "Suma sprzedaży", "Średnia sprzedaż",
                        "Liczba transakcji", "Średnia cena", "ZGODNE"]

FORMAT_KWOTA = '#,##0.00 "zł"'
FORMAT_CENA = "#,##0.00"
FORMAT_LICZBA = "0"
FORMAT_SREDNIA = "#,##0.0"

WERSJA_RAPORTU = "1.0"
AUTOR = "Kurs openpyxl, moduł 06"


# ----------------------------------------------------------------------
# DANE WEJŚCIOWE (deterministyczne)
# ----------------------------------------------------------------------
def generuj_dane(n: int, ziarno: int = 42) -> list[dict]:
    """Zwraca n rekordów. Celowo wstrzykuje kilka wadliwych przypadków."""
    generator = random.Random(ziarno)
    rekordy = []

    for i in range(1, n + 1):
        region = generator.choice(REGIONY)
        produkt = generator.choice(PRODUKTY)
        ilosc = generator.randint(1, 50)
        cena = CENY[produkt]

        # Co 40. rekord ma wadę - żeby walidacja miała co robić
        if i % 40 == 0:
            rodzaj = i // 40 % 3
            if rodzaj == 0:
                ilosc = -abs(ilosc)                  # ujemna ilość
            elif rodzaj == 1:
                cena = None                          # brak ceny
            else:
                region = ""                          # brak regionu

        rekordy.append({
            "Lp.": i, "Region": region, "Produkt": produkt,
            "Ilość": ilosc, "Cena": cena,
        })
    return rekordy


# ----------------------------------------------------------------------
# WALIDACJA - CZYSTA LOGIKA, BEZ ARKUSZA
# ----------------------------------------------------------------------
def waliduj(rekord: dict) -> list[str]:
    """Zwraca listę problemów w rekordzie. Pusta lista = rekord poprawny."""
    problemy = []

    if not rekord.get("Region"):
        problemy.append("brak regionu")

    if rekord.get("Produkt") not in CENY:
        problemy.append(f"nieznany produkt: {rekord.get('Produkt')!r}")

    ilosc = rekord.get("Ilość")
    if not isinstance(ilosc, int):
        problemy.append(f"ilość nie jest liczbą całkowitą: {ilosc!r}")
    elif ilosc <= 0:
        problemy.append(f"ilość musi być dodatnia: {ilosc}")

    cena = rekord.get("Cena")
    if cena is None:
        problemy.append("brak ceny")
    elif not isinstance(cena, (int, float)):
        problemy.append(f"cena nie jest liczbą: {cena!r}")
    elif cena <= 0:
        problemy.append(f"cena musi być dodatnia: {cena}")

    return problemy


def policz_sprzedaz(rekord: dict) -> float | None:
    """Sprzedaż = ilość * cena. None, gdy danych brak lub są niepoprawne."""
    ilosc, cena = rekord.get("Ilość"), rekord.get("Cena")
    if not isinstance(ilosc, int) or not isinstance(cena, (int, float)):
        return None
    if ilosc <= 0 or cena <= 0:
        return None
    return round(ilosc * cena, 2)


# ----------------------------------------------------------------------
# AGREGACJE W PYTHONIE - to samo, co robią formuły w arkuszu
# ----------------------------------------------------------------------
def agreguj(rekordy: list[dict]) -> dict[str, dict]:
    """Agregaty per region. Odpowiednik SUMIF/COUNTIF/AVERAGEIF."""
    suma: dict[str, float] = defaultdict(float)
    liczba: dict[str, int] = defaultdict(int)

    for rekord in rekordy:
        region = rekord.get("Region")
        sprzedaz = policz_sprzedaz(rekord)
        if not region or sprzedaz is None:
            continue
        if region not in REGIONY:                 # nieznany region nie wchodzi
            continue
        suma[region] += sprzedaz
        liczba[region] += 1

    wynik = {}
    for region in REGIONY:
        n = liczba[region]
        wynik[region] = {
            "suma": round(suma[region], 2),
            "srednia": round(suma[region] / n, 2) if n else 0.0,
            "liczba": n,
        }

    wszystkie = [policz_sprzedaz(r) for r in rekordy]
    wszystkie = [w for w in wszystkie if w is not None]
    wynik["RAZEM"] = {
        "suma": round(sum(wszystkie), 2),
        "srednia": round(sum(wszystkie) / len(wszystkie), 2) if wszystkie else 0.0,
        "liczba": len(wszystkie),
    }
    return wynik


def srednia_cena(rekordy: list[dict]) -> dict[str, float]:
    """Średnia cena per region (odpowiednik AVERAGEIF po kolumnie Cena)."""
    sumy: dict[str, float] = defaultdict(float)
    liczba: dict[str, int] = defaultdict(int)

    for rekord in rekordy:
        region, cena = rekord.get("Region"), rekord.get("Cena")
        if not region or region not in REGIONY:
            continue
        if not isinstance(cena, (int, float)) or cena <= 0:
            continue
        sumy[region] += cena
        liczba[region] += 1

    wszystkie = [r["Cena"] for r in rekordy
                 if isinstance(r.get("Cena"), (int, float)) and r["Cena"] > 0]

    wynik = {r: round(sumy[r] / liczba[r], 2) if liczba[r] else 0.0 for r in REGIONY}
    wynik["RAZEM"] = round(sum(wszystkie) / len(wszystkie), 2) if wszystkie else 0.0
    return wynik


# ----------------------------------------------------------------------
# BUDOWA ARKUSZY
# ----------------------------------------------------------------------
NAGLOWEK_FONT = Font(bold=True, color="FFFFFF", size=11)
NAGLOWEK_FILL = PatternFill(fill_type="solid", fgColor="2F5597")
NAGLOWEK_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
BLAD_FILL = PatternFill(fill_type="solid", fgColor="FCE4E4")
RAZEM_FILL = PatternFill(fill_type="solid", fgColor="E8EEF7")
KONTROLA_FONT = Font(color="595959", italic=True)


def styl_naglowka(ws, ostatnia_kolumna: int) -> None:
    for kolumna in range(1, ostatnia_kolumna + 1):
        komorka = ws.cell(row=1, column=kolumna)
        komorka.font = NAGLOWEK_FONT
        komorka.fill = NAGLOWEK_FILL
        komorka.alignment = NAGLOWEK_ALIGN
    ws.freeze_panes = "A2"


def zbuduj_dane(ws, rekordy: list[dict]) -> Counter:
    """Arkusz z danymi. Kolumna Sprzedaż jest FORMULA, kolumna Problem - tekstem."""
    ws.append(KOLUMNY_DANE)

    liczba_kolumn = len(KOLUMNY_DANE)
    kol_sprzedaz = KOLUMNY_DANE.index("Sprzedaż") + 1
    kol_problem = KOLUMNY_DANE.index("Problem") + 1
    litera_sprzedaz = get_column_letter(kol_sprzedaz)

    kategorie = Counter()

    for indeks, rekord in enumerate(rekordy, start=2):
        problemy = waliduj(rekord)

        for kolumna, nazwa in enumerate(KOLUMNY_DANE, start=1):
            if nazwa == "Sprzedaż":
                # FORMULA: użytkownik może poprawić ilość/cenę i zobaczyć wynik
                if problemy:
                    ws.cell(row=indeks, column=kolumna, value=None)
                else:
                    ilosc_lit = get_column_letter(KOLUMNY_DANE.index("Ilość") + 1)
                    cena_lit = get_column_letter(KOLUMNY_DANE.index("Cena") + 1)
                    ws.cell(row=indeks, column=kolumna,
                            value=f"={ilosc_lit}{indeks}*{cena_lit}{indeks}")
            elif nazwa == "Problem":
                ws.cell(row=indeks, column=kolumna,
                        value="; ".join(problemy) if problemy else None)
            else:
                ws.cell(row=indeks, column=kolumna, value=rekord.get(nazwa))

        if problemy:
            for kolumna in range(1, liczba_kolumn + 1):
                ws.cell(row=indeks, column=kolumna).fill = BLAD_FILL
            for problem in problemy:
                kategoria = problem.split(":")[0].split(" ")[0]
                kategorie[kategoria] += 1

    # formaty
    kol_ilosc = KOLUMNY_DANE.index("Ilość") + 1
    kol_cena = KOLUMNY_DANE.index("Cena") + 1
    for indeks in range(2, len(rekordy) + 2):
        ws.cell(row=indeks, column=kol_ilosc).number_format = FORMAT_LICZBA
        ws.cell(row=indeks, column=kol_cena).number_format = FORMAT_CENA
        ws.cell(row=indeks, column=kol_sprzedaz).number_format = FORMAT_KWOTA

    # szerokości
    szerokosci = {"Lp.": 7, "Region": 14, "Produkt": 14, "Ilość": 9,
                  "Cena": 11, "Sprzedaż": 14, "Problem": 46}
    for nazwa, szerokosc in szerokosci.items():
        ws.column_dimensions[get_column_letter(KOLUMNY_DANE.index(nazwa) + 1)].width = szerokosc

    styl_naglowka(ws, liczba_kolumn)
    return kategorie


def zbuduj_podsumowanie(ws, liczba_wierszy: int, agregaty: dict,
                        ceny_srednie: dict) -> None:
    """Arkusz żywych formuł. Wszystkie liczby to SUMIF/COUNTIF/AVERAGEIF."""
    ws.append(KOLUMNY_PODSUMOWANIA)

    ostatni_dane = liczba_wierszy + 1          # ostatni wiersz danych w arkuszu Dane
    litera_region = "B"
    litera_ilosc = "D"
    litera_cena = "E"
    litera_sprzedaz = "F"

    zakres_region = f"Dane!${litera_region}$2:${litera_region}${ostatni_dane}"
    zakres_sprzedaz = f"Dane!${litera_sprzedaz}$2:${litera_sprzedaz}${ostatni_dane}"
    zakres_cena = f"Dane!${litera_cena}$2:${litera_cena}${ostatni_dane}"

    for indeks, region in enumerate(REGIONY + ["RAZEM"], start=2):
        ws.cell(row=indeks, column=1, value=region)

        if region == "RAZEM":
            # Wiersz RAZEM liczy po całym zakresie, nie po warunku
            ws.cell(row=indeks, column=2, value=f"=SUM({zakres_sprzedaz})")
            ws.cell(row=indeks, column=3, value=f"=AVERAGE({zakres_sprzedaz})")
            ws.cell(row=indeks, column=4, value=f"=COUNT({zakres_sprzedaz})")
            ws.cell(row=indeks, column=5, value=f"=AVERAGE({zakres_cena})")
            for kolumna in range(1, len(KOLUMNY_PODSUMOWANIA) + 1):
                ws.cell(row=indeks, column=kolumna).fill = RAZEM_FILL
        else:
            kryterium = f"$A{indeks}"
            ws.cell(row=indeks, column=2,
                    value=f"=SUMIF({zakres_region},{kryterium},{zakres_sprzedaz})")
            ws.cell(row=indeks, column=3,
                    value=f"=AVERAGEIF({zakres_region},{kryterium},{zakres_sprzedaz})")
            ws.cell(row=indeks, column=4,
                    value=f"=COUNTIF({zakres_region},{kryterium})")
            ws.cell(row=indeks, column=5,
                    value=f"=AVERAGEIF({zakres_region},{kryterium},{zakres_cena})")

        # ZGODNE: porównuje FORMULE z WARTOŚCIĄ z arkusza Kontrola.
        # To jedyne miejsce, w którym formuła ma sens - porównuje dwa światy.
        ws.cell(row=indeks, column=6,
                value=f'=IF(ROUND(B{indeks},6)=ROUND(Kontrola!B{indeks},6),"OK","SPRAWDŹ")')

    for indeks in range(2, len(REGIONY) + 3):
        ws.cell(row=indeks, column=2).number_format = FORMAT_KWOTA
        ws.cell(row=indeks, column=3).number_format = FORMAT_KWOTA
        ws.cell(row=indeks, column=4).number_format = FORMAT_LICZBA
        ws.cell(row=indeks, column=5).number_format = FORMAT_CENA

    for kolumna, szerokosc in zip("ABCDEF", (14, 18, 20, 20, 16, 12)):
        ws.column_dimensions[kolumna].width = szerokosc

    styl_naglowka(ws, len(KOLUMNY_PODSUMOWANIA))


def zbuduj_kontrole(ws, agregaty: dict, ceny_srednie: dict,
                    kategorie: Counter) -> None:
    """Te same agregaty, ale jako WARTOŚCI policzone w Pythonie."""
    ws.append(KOLUMNY_PODSUMOWANIA)

    ws["A1"].comment = None      # bez komentarza - trzymamy się prostoty
    ws["H1"] = "Wartości policzone w Pythonie."
    ws["H1"].font = KONTROLA_FONT
    ws["H2"] = ("Porównaj z arkuszem Podsumowanie, jeśli podejrzewasz, "
                "że dane zmieniły się po wygenerowaniu.")
    ws["H2"].font = KONTROLA_FONT

    for indeks, region in enumerate(REGIONY + ["RAZEM"], start=2):
        dane = agregaty[region]
        ws.cell(row=indeks, column=1, value=region)
        ws.cell(row=indeks, column=2, value=dane["suma"])
        ws.cell(row=indeks, column=3, value=dane["srednia"])
        ws.cell(row=indeks, column=4, value=dane["liczba"])
        ws.cell(row=indeks, column=5, value=ceny_srednie[region])
        ws.cell(row=indeks, column=6, value=None)   # ZGODNE liczy Podsumowanie

        for kolumna in (1, 2, 3, 4, 5, 6):
            ws.cell(row=indeks, column=kolumna).font = KONTROLA_FONT

    for indeks in range(2, len(REGIONY) + 3):
        ws.cell(row=indeks, column=2).number_format = FORMAT_KWOTA
        ws.cell(row=indeks, column=3).number_format = FORMAT_KWOTA
        ws.cell(row=indeks, column=4).number_format = FORMAT_LICZBA
        ws.cell(row=indeks, column=5).number_format = FORMAT_CENA

    # blok walidacji - wartości, nie formuły
    start = len(REGIONY) + 4
    ws.cell(row=start, column=1, value="Walidacja danych wejściowych").font = Font(bold=True)
    ws.cell(row=start + 1, column=1, value="Kategoria problemu")
    ws.cell(row=start + 1, column=2, value="Liczba").font = Font(bold=True)
    ws.cell(row=start + 1, column=1).font = Font(bold=True)

    for offset, (kategoria, ile) in enumerate(sorted(kategorie.items()), start=2):
        ws.cell(row=start + offset, column=1, value=kategoria).font = KONTROLA_FONT
        ws.cell(row=start + offset, column=2, value=ile).font = KONTROLA_FONT

    suma_problemow = start + len(kategorie) + 2
    ws.cell(row=suma_problemow, column=1, value="RAZEM problemów").font = Font(bold=True)
    ws.cell(row=suma_problemow, column=2, value=sum(kategorie.values())).font = Font(bold=True)

    for kolumna, szerokosc in zip("ABCDEF", (14, 18, 20, 20, 16, 12)):
        ws.column_dimensions[kolumna].width = szerokosc
    ws.column_dimensions["H"].width = 46

    styl_naglowka(ws, len(KOLUMNY_PODSUMOWANIA))


def zbuduj_metodyke(ws, liczba_wierszy: int, kategorie: Counter) -> None:
    ws.column_dimensions["A"].width = 120
    tresci = [
        ("Metodyka raportu", True),
        ("", False),
        ("Jakie elementy są FORMULAMI i dlaczego:", True),
        ("  - Kolumna Sprzedaż w arkuszu Dane (=Dn*En). Użytkownik może poprawić "
         "ilość lub cenę i natychmiast zobaczyć nową wartość. To jest funkcja pliku, "
         "a nie sposób liczenia.", False),
        ("  - Cała tabela w arkuszu Podsumowanie (SUMIF, COUNTIF, AVERAGEIF). "
         "Podsumowanie ma się przeliczać, gdy zmienią się dane wejściowe.", False),
        ("  - Kolumna ZGODNE w arkuszu Podsumowanie. To formuła porównująca dwie "
         "niezależne wartości - jej zadaniem jest wykrycie rozbieżności.", False),
        ("", False),
        ("Jakie elementy są WARTOŚCIAMI i dlaczego:", True),
        ("  - Cały arkusz Kontrola. Te same agregaty policzone w Pythonie. "
         "Muszą być wartościami, bo ich zadaniem jest istnienie NIEZALEŻNIE "
         "od tego, czy ktokolwiek policzy formuły.", False),
        ("  - Blok Walidacja w arkuszu Kontrola. Liczba wykrytych problemów "
         "to wynik pracy Pythona - nie formuła.", False),
        ("  - Kolumna Problem w arkuszu Dane. Opis błędu jest tekstem, "
         "nie wynikiem obliczeń Excela.", False),
        ("", False),
        ("Co się stanie z Podsumowaniem w Excelu sprzed 365:", True),
        ("  Formuły SUMIF, COUNTIF i AVERAGEIF istnieją od dziesięcioleci, "
         "więc zadziałają. Gdybyśmy użyli UNIQUE/FILTER/XLOOKUP - dostalibyśmy "
         "#NAME?. Świadomie ich nie użyliśmy.", False),
        ("", False),
        ("Co się stanie, gdy inny program odczyta plik bez otwierania go w Excelu:", True),
        ("  data_only=True zwróci None dla wszystkiego, co jest formułą: "
         "wszystkich komórek w arkuszu Dane (Sprzedaż) oraz w Podsumowaniu. "
         "Dlatego agregaty powtórzono jako WARTOŚCI w arkuszu Kontrola. "
         "Program czytający plik powinien czytać Kontrolę, nie Podsumowanie.", False),
        ("", False),
        ("O co chodzi z fullCalcOnLoad:", True),
        ("  Skoroszyt ma ustawione fullCalcOnLoad=True. Oznacza to, że Excel po "
         "otwarciu pliku przeliczy wszystko od nowa, zamiast ufać wartościom "
         "z pliku. Bez tego użytkownik mógłby zobaczyć puste komórki albo "
         "sekwencję przeliczania.", False),
        ("", False),
        ("Gdzie leży logika biznesowa:", True),
        ("  Cała logika - walidacja danych wejściowych, wykrywanie niepoprawnych "
         "ilości i cen, wyliczanie agregatów, kategoryzacja problemów - leży "
         "w Pythonie (funkcje waliduj, policz_sprzedaz, agreguj). Formuły w arkuszu "
         "Podsumowanie NIE liczą niczego, czego Python nie policzyłby sam. "
         "Ich rolą jest danie użytkownikowi ŻYWEGO narzędzia: gdy zmieni dane "
         "wejściowe, zobaczy nowe podsumowanie bez uruchamiania naszego skryptu. "
         "Arkusz Kontrola jest zabezpieczeniem: pokazuje, co system policzył "
         "niezależnie, i pozwala wykryć rozbieżność.", False),
    ]

    for indeks, (tekst, pogrubienie) in enumerate(tresci, start=1):
        komorka = ws.cell(row=indeks, column=1, value=tekst)
        komorka.alignment = Alignment(wrap_text=True, vertical="top")
        if pogrubienie:
            komorka.font = Font(bold=True, size=12)
        if tekst:
            ws.row_dimensions[indeks].height = 30 if len(tekst) > 100 else 18


def zbuduj_ustawienia(ws, liczba_wierszy: int) -> None:
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 42

    ustawienia = [
        ("Wersja raportu", WERSJA_RAPORTU),
        ("Data wygenerowania", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("Autor", AUTOR),
        ("Liczba wierszy danych", liczba_wierszy),
        ("Arkusz danych", "Dane"),
        ("Arkusz prezentacji", "Podsumowanie"),
        ("Arkusz kontroli", "Kontrola"),
        ("Regiony", ", ".join(REGIONY)),
        ("Produkty", ", ".join(PRODUKTY)),
        ("Próg minimalnej ceny", 0.01),
        ("Próg maksymalnej ilości", 1000),
        ("Uwagi", "Regiony i produkty przekazywane do formuł są brane z tego arkusza."),
    ]

    ws.append(["Parametr", "Wartość"])
    for nazwa, wartosc in ustawienia:
        ws.append([nazwa, wartosc])

    styl_naglowka(ws, 2)


# ----------------------------------------------------------------------
# ORKIESTRACJA
# ----------------------------------------------------------------------
def main(liczba_wierszy: int = 200) -> int:
    rekordy = generuj_dane(liczba_wierszy)
    agregaty = agreguj(rekordy)
    ceny_srednie = srednia_cena(rekordy)

    wb = Workbook()

    ws_dane = wb.active
    ws_dane.title = "Dane"
    kategorie = zbuduj_dane(ws_dane, rekordy)

    ws_pod = wb.create_sheet("Podsumowanie")
    zbuduj_podsumowanie(ws_pod, liczba_wierszy, agregaty, ceny_srednie)

    ws_kontrola = wb.create_sheet("Kontrola")
    zbuduj_kontrole(ws_kontrola, agregaty, ceny_srednie, kategorie)

    ws_metodyka = wb.create_sheet("Metodyka")
    zbuduj_metodyke(ws_metodyka, liczba_wierszy, kategorie)

    ws_ustawienia = wb.create_sheet("Ustawienia")
    zbuduj_ustawienia(ws_ustawienia, liczba_wierszy)

    # Przeliczenie przy otwarciu - kluczowe, bo plik ma formuły bez cache
    if wb.calculation is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True

    wb.properties.title = "Raport sprzedaży według regionów"
    wb.properties.creator = AUTOR
    wb.properties.description = (
        f"Raport w wersji {WERSJA_RAPORTU}. Podsumowanie w formułach, "
        f"kontrola w wartościach. Wygenerowano {datetime.now():%Y-%m-%d %H:%M}."
    )

    wb.active = 1                      # otwiera się na Podsumowaniu

    cel = OUTPUT / "06_cw4_raport.xlsx"
    wb.save(cel)

    print("=" * 84)
    print("RAPORT WYGENEROWANY")
    print("=" * 84)
    print(f"  plik:              {cel}")
    print(f"  wierszy danych:    {liczba_wierszy}")
    print(f"  problemów danych:  {sum(kategorie.values())}")
    for kategoria, ile in sorted(kategorie.items()):
        print(f"    - {kategoria:<20} {ile}")

    print()
    print("  Agregaty policzone w Pythonie (arkusz Kontrola):")
    print(f"  {'Region':<12} {'Suma':>14} {'Średnia':>14} {'Liczba':>8} {'Śr. cena':>10}")
    print("  " + "-" * 62)
    for region in REGIONY + ["RAZEM"]:
        d = agregaty[region]
        print(f"  {region:<12} {d['suma']:>14,.2f} {d['srednia']:>14,.2f} "
              f"{d['liczba']:>8} {ceny_srednie[region]:>10,.2f}")

    print()
    print("  Oczekiwany wynik kolumny ZGODNE:")
    print("    Wszystkie wiersze powinny pokazać 'OK' po otwarciu w Excelu,")
    print("    O ILE nikt nie zmienił danych w arkuszu Dane.")
    print("    Jeśli zobaczysz 'SPRAWDŹ' bez zmian w danych - to sygnał,")
    print("    że coś się rozjechało w formułach albo w agregacji w Pythonie.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

**Kluczowe decyzje projektowe i dlaczego właśnie tak:**

- **Python liczy wszystko, a formuły nic nie wnoszą nowego.** To najważniejsza obserwacja w całym rozwiązaniu. Funkcje `waliduj`, `policz_sprzedaz`, `agreguj` i `srednia_cena` są **całkowicie niezależne** od openpyxl — nie importują arkusza, nie dotykają komórek, nie wiedzą o Excelu. Można je przetestować jedną linijką (`assert agreguj([...])["Północ"]["suma"] == 123.45`) i użyć w CLI, w API, w CSV. Formuły w arkuszu `Podsumowanie` są **prezentacją tych samych liczb** dla człowieka, który chce mieć żywe narzędzie.

- **Dwa arkusze z tymi samymi liczbami — `Podsumowanie` i `Kontrola`.** To nie duplikacja, to **redundancja kontrolna**. Jej celem jest wykrycie sytuacji, w której wartości w pliku rozjeżdżają się z prawdą — na przykład dlatego, że ktoś zmienił dane wejściowe w arkuszu `Dane`, ale nie otworzył pliku ponownie, żeby się przeliczyło. Kolumna `ZGODNE` zamienia tę kontrolę w formułę, którą **użytkownik** może zobaczyć bez uruchamiania Pythona. Wykorzystujemy tu fakt, że formuła **porównująca** jest bezpieczna — nie musi być aktualna, żeby była przydatna.

- **`Dane!$B$2:$B$201` budowane programowo.** Litera kolumny pochodzi z `get_column_letter(KOLUMNY_DANE.index("Region") + 1)`, a numer ostatniego wiersza z `liczba_wierszy + 1`. Dzięki temu zmiana liczby wierszy danych nie wymaga edycji ani jednej formuły — a zmiana kolejności kolumn w `KOLUMNY_DANE` **też nie**, bo wszystko liczy się z listy. To jest ten rodzaj „konfiguracji deklaratywnej", o którym będzie moduł 26, tylko w prostszej formie.

- **`fullCalcOnLoad=True`.** Bez tego plik z formułami bez cache mógłby otworzyć się z pustymi komórkami w `Podsumowaniu`. Z tym — Excel przelicza wszystko przed pokazaniem. To nie jest kosmetyka: to element poprawnego działania raportu. Wyjaśnienie jest też w arkuszu `Metodyka`, bo **użytkownik ma prawo wiedzieć, czemu ufa**.

- **`random.Random(42)` zamiast globalnego `random.seed()`.** Instancja generatora jest **lokalna** dla funkcji `generuj_dane`. Nie zmienia globalnego stanu modułu `random`, więc rozwiązanie jest bezpieczne, gdyby ktoś użył go w większym systemie, w którym inne moduły też losują. To drobiazg, ale dokładnie ten rodzaj drobiazgu, który odróżnia kod „działa u mnie" od kodu produkcyjnego.

- **Arkusz `Ustawienia` jako dane, nie jako kod.** Regiony, produkty, progi walidacji — wszystko w komórkach. To jeszcze nie pełna konfiguracja deklaratywna (bo funkcje nadal mają je zaszyte), ale pokazuje kierunek. W module 26 zobaczysz, jak to wyciągnąć do YAML-a i uczynić naprawdę wymiennym.

- **Brak `read_only` i `write_only`.** Przy 200 wierszach niepotrzebne. Ale w komentarzu warto zapisać, że dla 200 000 wierszy struktura byłaby inna: `Dane` generowane w `write_only`, a `Podsumowanie` — już nie, bo `write_only` nie pozwala wracać do arkusza. To realne wyzwanie projektowe (moduł 19).

**Czego ten szkic jeszcze nie robi** (a w wersji produkcyjnej powinien): zapisu atomowego z kopią zapasową (moduł 18), sanitacji danych wejściowych od użytkownika (moduł 21), testów (moduł 20 — choć funkcje są już do tego gotowe, co było celem), wyniesienia konfiguracji do YAML (moduł 26). Ale najważniejsza decyzja — **gdzie jest logika biznesowa** — jest już podjęta i uzasadniona, a to jest sedno zadania.

</details>

## 6. Typowe błędy i pułapki

**1. „Wpisałem `ws["C1"] = "SUM(A1:B1)"` i Excel pokazuje to jako napis, a nie liczy" (objaw) → brak znaku `=` na początku; openpyxl traktuje tekst bez `=` jak zwykły string i zapisuje go jako tekst (przyczyna) → dodaj `=`: `ws["C1"] = "=SUM(A1:B1)"`. Sprawdź też, czy nie wkleiłeś przypadkiem apostrofu z przodu (niektóre narzędzia dodają go „dla bezpieczeństwa") (naprawa).**

Idealne do wykrycia jedną linijką: `if komorka.data_type != "f": print("to nie jest formuła")`. Jeśli po zapisie komórka nie ma `data_type == 'f'`, to znaczy, że openpyxl nie rozpoznał w niej formuły — i cała reszta dyskusji o cache i `data_only` jest bezprzedmiotowa.

**2. „Wpisałem `=SUMA(A1;B1)` i Excel pokazuje `#NAME?`" (objaw) → w pliku `.xlsx` formuły są zapisywane w wariancie en-US: angielskie nazwy funkcji i przecinek jako separator argumentów; lokalizacja to sprawa wyświetlania, nie formatu pliku. `SUMA` i średnik nie istnieją w gramatyce pliku (przyczyna) → pisz `=SUM(A1,B1)`, `=IF(...)`, `=AVERAGE(...)`; nie kopiuj formuł z paska formuły polskiego Excela do kodu; pamiętaj też o **kropce dziesiętnej** w liczbach wstawianych do formuł (naprawa).**

Szczególnie zdradliwy wariant tego błędu: `f"=ROUND({wartosc},2)"` przy polskim locale. Wartość `1,5` (bo tak wygląda w UI) wstawi przecinek do formuły i powstanie `=ROUND(1,5,2)` — **trzy argumenty**. Excel pokaże błąd składni, a Ty będziesz szukać winy w Excelu.

**3. „Wypisałem `=SUM(A1:B1)` w kodzie, a `data_only=True` zwraca `None`" (objaw) → nie ma pamięci podręcznej. openpyxl nie liczy formuł — zapisuje tylko element `<f>` bez elementu `<v>`. Nie ma czego odczytać (przyczyna) → otwórz plik w Excelu (albo LibreOffice z włączonym przeliczaniem przy wczytywaniu), zapisz, i dopiero wtedy wczytaj z `data_only=True`. Alternatywnie policz to samo sam w Pythonie (naprawa).**

Do tego błędu dochodzi jeszcze jedno źródło: **przepuszczenie pliku przez openpyxl kasuje cache**. Nawet jeśli plik miał cache, to po `load_workbook(...)` + `save(...)` w trybie domyślnym cache zniknie — bo openpyxl zapisze formuły bez wartości. To jest konsekwencja modelu z modułu 03 i jeden z punktów „co ginie przy zapisie" z modułu 18.

**4. „Wygenerowałem plik z formułami, otworzyłem w Excelu, zapisałem — i wszystkie formuły zamieniły się na liczby" (objaw) → skoroszyt był wczytany z `data_only=True`, więc w modelu nie było ani jednej formuły — tylko same wartości; `save()` zapisał model taki, jaki był. **Nie ma ostrzeżenia, bo nie ma czego ostrzegać** (przyczyna) → **nigdy nie łącz `load_workbook(data_only=True)` z `save()` na tym samym obiekcie.** Jeśli potrzebujesz wartości — wczytaj je w osobnej sesji, wyciągnij i zamknij. If potrzebujesz zapisać zmiany — otwórz plik drugi raz w trybie domyślnym. Napisz strażnika (`policz_formuly()` + `bezpieczny_zapis()` z sekcji 2.4) (naprawa).**

Nie ma sposobu, żeby odzyskać formuły z pliku, w którym ich nie ma — a jeśli nadpisałeś oryginał (bo „przecież działa"), to nie ma nawet z czego odbudować. To najdroższy błąd w całym module.

**5. „Skopiowałem formułę w dół i wszystkie odwołania zostały takie same — Excel pokazuje bezsensowne wyniki" (objaw) → openpyxl nie tłumaczy referencji przy przypisaniu tekstu do komórki; nie zna pojęcia „kopiowanie komórki". Przepisuje tekst dosłownie (przyczyna) → generuj formułę z właściwymi numerami (`f"=SUM(A{i}:B{i})"`) albo użyj `Translator(formuła, origin=...).translate_formula(cel)`. To **nie jest** błąd openpyxl — to jego kontrakt: robi dokładnie to, o co prosisz (naprawa).**

Ten błąd jest o tyle zdradliwy, że **objawia się dopiero w Excelu**. W kodzie wszystko wygląda dobrze — formuła jest, komórka jest. Dopiero użytkownik zauważa, że każdy wiersz pokazuje ten sam wynik.

**6. „Zmieniłem nazwę arkusza i formuły w innych arkuszach pokazują `#REF!`" (objaw) → openpyxl nie ma grafu zależności i nie śledzi odwołań w formułach; `ws.title = "Nowa nazwa"` zmienia tylko nazwę, nie treść formuł, w których stara nazwa jest zapisana jako tekst (przyczyna) → traktuj nazwy arkuszy jak **publiczny kontrakt** — ustal je raz i nie zmieniaj. Jeśli musisz zmienić, przejdź po wszystkich arkuszach i w formułach zamień stary fragment na nowy (ostrożnie, bo zamiana tekstu może trafić w coś innego). Zbuduj to jako funkcję `zmien_nazwe_arkusza(wb, stara, nowa)` (naprawa).**

Powiązany przypadek, jeszcze gorszy: **usunięcie arkusza** (`del wb["Dane 2026"]`) pozostawia formuły wskazujące na nieistniejący arkusz. openpyxl nawet tego nie zauważy — błąd objawi się dopiero w Excelu.

**7. „Komórka z formułą dynamiczną `=UNIQUE(...)` pokazuje `#NAME?`, chociaż mam Excel 365" (objaw) → w pliku `.xlsx` nowsze funkcje są zapisywane z prefiksem (`_xlfn.UNIQUE`, `_xlfn.XLOOKUP`, część z `_xlws.`) z powodu zamkniętego słownika nazw w oryginalnej specyfikacji formatu; openpyxl nie dodaje tego prefiksu automatycznie (przyczyna) → sprawdź empirycznie, jak Excel zapisuje tę funkcję: utwórz formułę ręcznie, zapisz plik, rozpakuj ZIP i zobacz `xl/worksheets/sheet1.xml`. Wpisz w kodzie **dokładnie** to samo. To jedyna metoda dająca pewność dla Twojej wersji (naprawa).**

Alternatywa, która jest zwykle lepsza: **policz to w Pythonie**. `UNIQUE` to `sorted(set(...))`, `FILTER` to list comprehension. Wynik zapisany jako wartości działa w każdej wersji Excela, jest widoczny natychmiast i nie wymaga metadanych rozlania, których openpyxl i tak nie generuje.

**8. „Zapisałem formułę tablicową i Excel zgłasza, że plik jest uszkodzony" (objaw) → najczęściej `ArrayFormula` została przypisana do **wielu** komórek (zamiast raz, w lewej-górnej, z `ref` obejmującym cały zakres), albo `ref` nie zgadza się z rzeczywistym umieszczeniem, albo zakresy dwóch formuł tablicowych na siebie nachodzą (przyczyna) → przypisz `ArrayFormula` **dokładnie raz**, do komórki lewej-górnej, z `ref` wskazującym na pełny zakres. Sprawdź, czy nie przypisujesz tej samej instancji obiektu do kilku komórek (naprawa).**

Druga wersja tego błędu: pomylenie `ArrayFormula` z **formułą dynamiczną**. Formuła dynamiczna to zwykły `str` z prefiksem `_xlfn.` — **nie** opakowuj jej w `ArrayFormula`, bo Excel dostanie informację, że to formuła CSE obejmująca zakres, i nie rozleje jej poprawnie.

**9. „Po wczytaniu pliku z formułą tablicową mój kod wywalił się z `AttributeError: 'ArrayFormula' object has no attribute 'startswith'" (objaw) → w komórce z formułą tablicową `cell.value` **nie jest** `str`, a obiektem `ArrayFormula`; kod, który zakładał, że każda formuła to tekst zaczynający się od `=`, nie ma prawa działać (przyczyna) → sprawdzaj typ przed użyciem: `if isinstance(komorka.value, str)` / `elif isinstance(komorka.value, ArrayFormula)`. Albo sprawdzaj `data_type == 'f'` i dopiero potem sięgaj po `value`. Uwaga: `ArrayFormula.text` często **nie zawiera** znaku `=` — sprawdź w swojej wersji (naprawa).**

Ogólna zasada z tego błędu: **`cell.value` nie zawsze zwraca typ z listy z modułu 05.** Formuła, komentarz w scalonej komórce, hiperlink — to wszystko miejsca, w których openpyxl może zwrócić obiekt. Zawsze sprawdzaj typ, gdy nie jesteś pewien.

**10. „Użytkownik wpisał komentarz zaczynający się od `=` i teraz plik uruchamia makro/link" (objaw) → tekst pochodzący od użytkownika, zaczynający się od `=`, `+`, `-` albo `@`, zostaje zapisany jako **formuła**, a Excel go wykona (przyczyna) → **sanitizuj wszystko, co pochodzi spoza Twojego kodu**, zanim trafi do komórki: usuń znak wyzwalający albo ustaw `data_type = "s"` po przypisaniu wartości. Wymuszaj typ tekstowy na kolumnach, które są identyfikatorami (`number_format = "@"`). Szczegóły i pełna funkcja — moduł 21 (naprawa).**

To nie jest tylko kwestia bezpieczeństwa, ale też **poprawności**: użytkownik wpisuje `-5`, jeśli chodziło o liczbę, a Ty zapisujesz formułę `=-5` — która akurat zadziała. Ale `+48 512 345 678` zapisane jako formuła pokaże `#NAME?` albo coś jeszcze dziwniejszego. Sanitizacja chroni jednocześnie przed atakiem i przed śmieciami w raporcie.

**11. „Ustawiłem `fullCalcOnLoad=True`, ale `data_only=True` nadal zwraca `None`" (objaw) → `fullCalcOnLoad` to **prośba do programu, który otworzy plik** (Excela, LibreOffice), żeby przeliczył formuły. To nie jest mechanizm openpyxl i nie sprawia, że Python cokolwiek policzy (przyczyna) → zrozum podział ról: `fullCalcOnLoad` mówi Excelowi „policz", `data_only=True` mówi openpyxl „przeczytaj to, co już policzone". Jeśli potrzebujesz wartości w Pythonie bez otwierania Excela — policz je sam (naprawa).**

Warto to powtórzyć trzy razy, bo to najczęstsze nieporozumienie w całym module: **openpyxl nie liczy formuł. Nigdy. Nie ma tam silnika.**

**12. „Formuła jest zapisana, ale Excel po otwarciu pokazuje komunikat o przeliczaniu / puste komórki" (objaw) → plik ma formuły bez cache i **nie** ma flagi pełnego przeliczenia; Excel może próbować użyć wartości z pliku, których nie ma, i pokazać pustkę, albo przeliczyć tylko część (przyczyna) → ustaw `wb.calculation.fullCalcOnLoad = True` **przed każdym** `save()` w plikach, które zawierają formuły. Pamiętaj o defensywnym `if wb.calculation is None` — przy wczytanym pliku bez `calcPr` ten atrybut może nie istnieć (naprawa).**

Powiązana pułapka: **sprzeczne ustawienia przeliczania w pliku z jednej strony, a oczekiwanie z drugiej.** Jeśli Twój system odczytuje plik zaraz po wygenerowaniu (np. wysyła go mailem), a odbiorcą jest program — formuły są bezużyteczne, choćbyś ustawił wszystko po mistrzowsku. Zapamiętaj tabelę decyzyjną z sekcji 2.18.

**13. „Wygenerowałem 50 000 wierszy z formułami i plik otwiera się pół minuty" (objaw) → każda formuła to dodatkowy łańcuch znaków w pliku plus obowiązek przeliczenia po stronie Excela; przy dziesiątkach tysięcy formuł koszt rośnie liniowo, a Excel przelicza je przy każdym otwarciu (przyczyna) → dla dużych zbiorów pisz **wartości** policzone w Pythonie; formuły zarezerwuj dla kluczowych komórek podsumowania (kilkadziesiąt, nie kilkadziesiąt tysięcy). Rozważ `write_only=True` przy generowaniu (naprawa, szczegóły w module 19).**

Analogia: **formuła to instrukcja dla kogoś innego, a nie Twój obliczony wynik.** Jeśli wygenerujesz 50 000 instrukcji, ktoś je wszystkie wykona — przy każdym otwarciu pliku, za każdym razem. Excel sobie z tym radzi, ale użytkownik czeka.

**14. „Mam plik z odwołaniem do zewnętrznego skoroszytu i po przepuszczeniu przez openpyxl zobaczyłem `#REF!`" (objaw) → odwołania do zewnętrznych plików (linki) to osobne części pliku `.xlsx`; openpyxl nie zawsze zachowuje je w pełni, a brak wpisu w liście linków powoduje, że Excel nie umie rozwiązać odwołania (przyczyna) → wczytaj plik z `keep_links=True` (domyślne) i **sprawdź efekt**; jeśli zewnętrzne odwołania są w Twoim pliku istotne — nie przepuszczaj pliku przez openpyxl (reguła z modułu 18) albo przekształć odwołania zewnętrzne w wartości przed zapisem. Diagnostyka: sprawdź listę linków na obiekcie skoroszytu (atrybut wewnętrzny — sprawdź nazwę w swojej wersji) (naprawa).**

Szersza zasada, do której wróci moduł 18: **im „bogatszy" plik, tym większe ryzyko, że przepuszczenie go przez openpyxl coś zgubi.** Formuły z odwołaniami zewnętrznymi to jeden z tych przypadków.

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Formuła to dla openpyxl tekst zaczynający się od `=`. Nic więcej.** Metoda `_bind_value()` sprawdza pierwszy znak, ustawia `data_type = 'f'` i zapamiętuje tekst. **Nie parsuje. Nie waliduje. Nie sprawdza nazw funkcji.** W pliku formuła zapisuje się jako `<f>SUM(A1:B1)</f>` — bez znaku `=` i **bez elementu `<v>`**, czyli bez wyniku. Ten brakujący element jest przyczyną wszystkiego, co dzieje się dalej.

2. **openpyxl nie liczy i nie będzie liczyć.** To biblioteka formatu pliku, nie arkusz kalkulacyjny. Nie ma tam `SUM`, `VLOOKUP` ani `IF`. Wniosek praktyczny: **jeśli Twój program potrzebuje liczby, policz ją w Pythonie.** Formuła to instrukcja dla człowieka, który otworzy plik — a nie sposób na uzyskanie danych w kodzie. Tabela decyzyjna z sekcji 2.18 to najważniejsza rzecz w tym module.

3. **`data_only=True` czyta **cache**, nie liczy.** Cache to element `<v>` zapisany obok formuły przez program, który umiał liczyć (Excel, LibreOffice). Jeśli plik powstał z openpyxl, cache **nie ma** → dostajesz `None`. Jeśli plik przeszedł przez openpyxl, cache **ginie**. Jeśli ktoś zmienił dane wejściowe bez przeliczenia, cache **kłamie po cichu**. W trybie `data_only=True` model nie zawiera **ani jednej** formuły — i dlatego **`data_only=True` + `save()` = bezpowrotna utrata formuł**. To najdroższy błąd tego modułu; broń się zasadą „inna sesja do odczytu, inna do zapisu".

4. **Formuła to kontrakt z konkretnym programem — i tylko z nim.** W pliku `.xlsx` obowiązuje wariant en-US: **angielskie nazwy funkcji, przecinek jako separator, kropka dziesiętna**. Nowe funkcje dynamiczne wymagają prefiksu (`_xlfn.UNIQUE`), którego openpyxl nie dodaje automatycznie — sprawdź empirycznie, jak zapisuje je Twój Excel. Referencje są w notacji A1 i **openpyxl ich nie tłumaczy przy kopiowaniu** — albo generujesz formułę z właściwymi numerami, albo używasz `Translator`. A nazwy arkuszy i zakresy, na których opierają się formuły, to **publiczny interfejs** — zmiana nazwy arkusza łamie odwołania po cichu.

5. **Najlepszy wzorzec to hybryda: formuły dla człowieka, wartości dla systemu, obie w jednym pliku.** Arkusz z żywymi formułami (`SUMIF`, `AVERAGEIF`) daje użytkownikowi narzędzie — może zmienić dane i zobaczyć nowe sumy. Równoległy arkusz z tymi samymi liczbami policzonymi w Pythonie daje **kontrolę**: pokazuje, co system policzył niezależnie, i pozwala wykryć rozbieżność. Formuła porównująca te dwa światy jest jedynym miejscem, gdzie formuła ma sens „na wyjściu" — bo jej zadaniem nie jest liczenie, a **wykrywanie rozbieżności**. To wzorzec, który rozbudujemy w module 26 jako „Excel to warstwa prezentacji, nie silnik biznesowy".

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
from pathlib import Path
from zipfile import ZipFile
import shutil
import subprocess

from openpyxl import Workbook, load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, quote_sheetname, absolute_coordinate
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# ==================================================================
# 2. ZAPIS FORMUL
# ==================================================================
ws["C1"] = "=SUM(A1:B1)"                       # formuła: tekst z '=' na początku
ws["C2"] = f"=SUM(A2:A{ostatni})"              # interpolacja - działa, to zwykły tekst
ws["C3"] = "SUM(A1:B1)"                        # ❌ TO NIE JEST FORMULA. To napis.
ws["C4"] = "=SUMA(A1;B1)"                      # ❌ po polsku i ze średnikiem NIE ZADZIAŁA
ws["C5"] = "=SUM(A1,B1)"                       # ✅ angielska nazwa, przecinek

ws["C1"].number_format = "#,##0.00"            # format to osobna warstwa (moduł 10)

# Międzyarkuszowo (nazwy ze spacjami -> apostrofy):
nazwa = "Dane 2026"
ws["D1"] = f"=SUM({quote_sheetname(nazwa)}!$B$2:$B$100)"

# Liczby w formule: KROPKA dziesiętna, przecinek jako separator:
ws["D2"] = "=ROUND(A1*1.23,2)"                 # ✅
# ws["D3"] = "=ROUND(A1*1,23;2)"               # ❌ to da trzy argumenty!

# ==================================================================
# 3. ODCZYT - DWA TRYBY, DWA ŚWIATY
# ==================================================================
# TRYB DOMYŚLNY: literały formuł
wb = load_workbook("plik.xlsx")                 # data_only=False
k = wb.active["C1"]
print(k.value, k.data_type)                     # '=SUM(A1:B1)' 'f'

# TRYB WARTOŚCI: cache
wb_w = load_workbook("plik.xlsx", data_only=True)
print(wb_w.active["C1"].value)                  # 30 ALBO None (gdy brak cache)

# UWAGA: to są DWIE OSOBNE sesje. Nie mieszaj ich.
# NIGDY: load_workbook(..., data_only=True) -> ... -> save()  <- ZNISZCZYSZ FORMULY

# ==================================================================
# 4. AUDYT PLIKU - CZY SĄ FORMULY, CZY JEST CACHE
# ==================================================================
def policz_formuly(sciezka) -> int:
    """Ile komórek z formułami ma plik? Tani odczyt, read_only."""
    wb = load_workbook(sciezka, read_only=True)
    try:
        return sum(
            1
            for ws in wb.worksheets
            for wiersz in ws.iter_rows()
            for komorka in wiersz
            if komorka.data_type == "f"
        )
    finally:
        wb.close()


def porownaj_dwa_swiaty(sciezka):
    """Lista (arkusz, adres, formula, wartosc_z_cache)."""
    wb_f = load_workbook(sciezka)
    wb_w = load_workbook(sciezka, data_only=True)
    wynik = []
    try:
        for nazwa in wb_f.sheetnames:
            ws_f, ws_w = wb_f[nazwa], wb_w[nazwa]
            for wiersz in ws_f.iter_rows():
                for komorka in wiersz:
                    if komorka.data_type == "f":
                        wynik.append((
                            nazwa, komorka.coordinate,
                            komorka.value, ws_w[komorka.coordinate].value,
                        ))
        return wynik
    finally:
        wb_f.close()
        wb_w.close()


# Strażnik przed katastrofą:
def bezpieczny_zapis(wb, cel, *, zrodlo_mialo_formuly: bool) -> None:
    if zrodlo_mialo_formuly:
        raise RuntimeError(
            "Skoroszyt wczytano w trybie data_only=True, a źródło miało formuły. "
            "Zapis zniszczyłby formuły. Wczytaj ponownie bez data_only."
        )
    wb.save(cel)


# ==================================================================
# 5. WYMUSZANIE PRZELICZENIA PRZY OTWARCIU
# ==================================================================
def wymus_przeliczenie(wb) -> None:
    """Prosi program otwierający plik o pełne przeliczenie. Sprawdź w swojej wersji."""
    if wb.calculation is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True
# W pliku: <calcPr calcId="..." fullCalcOnLoad="1"/> w xl/workbook.xml
# To NIE sprawia, że openpyxl liczy!

# ==================================================================
# 6. TRANSLATOR - PRZESUWANIE REFERENCJI
# ==================================================================
tlumacz = Translator("=SUM($A2:B2)*$D$1", origin="C2")
tlumacz.translate_formula("C3")      # '=SUM($A3:B3)*$D$1'
tlumacz.translate_formula("D4")      # '=SUM($A4:C4)*$D$1'
# Referencje z $ zostają, względne się przesuwają - jak w Excelu.
# Metody do przesuwania przy insert/delete wierszy: sprawdź w dokumentacji.

# Bez Translatora - świadome generowanie z właściwymi numerami:
for i in range(2, 101):
    ws.cell(row=i, column=3, value=f"=SUM(A{i}:B{i})")

# ==================================================================
# 7. NAZWY ZDEFINIOWANE W FORMULACH
# ==================================================================
konf = wb.create_sheet("Konfiguracja")
konf["A1"], konf["B1"] = "Stawka VAT", 0.23

odniesienie = f"{quote_sheetname(konf.title)}!{absolute_coordinate('B1')}"
wb.defined_names.add(DefinedName(name="STAWKA_VAT", attr_text=odniesienie))

ws["B1"] = "=A1*STAWKA_VAT"          # używamy nazwy w formule
ws["B2"] = "=SUM(SPRZEDAZ_MIESIECZNA)"

# Odczyt definicji:
for nazwa in wb.defined_names:
    print(nazwa, "->", wb.defined_names[nazwa].attr_text)

# ==================================================================
# 8. FORMULA TABLICOWA (ArrayFormula)
# ==================================================================
# Przypisz RAZ, do lewej-górnej komórki, z ref obejmującym CAŁY zakres:
ws["D1"] = ArrayFormula(ref="D1:D10", text="=A1:A10*B1:B10")

# Odczyt - value NIE jest str:
k = ws["D1"]
if isinstance(k.value, ArrayFormula):
    print(k.value.text, k.value.ref)   # text często BEZ '=' - sprawdź wersję

# ==================================================================
# 9. FORMULY DYNAMICZNE (Excel 365+)
# ==================================================================
# W pliku wymagają prefiksu. SPRAWDŹ EMPIRYCZNIE w swojej wersji:
ws["F1"] = "=_xlfn.UNIQUE(A2:A100)"
ws["H1"] = '=_xlfn.XLOOKUP(A2,Klienci!A:A,Klienci!B:B,"nie znaleziono")'
ws["J1"] = '=_xlfn.FILTER(A2:B100,B2:B100>1000,"brak danych")'
# openpyxl NIE generuje metadanych rozlania (spill).
# Alternatywa, która zawsze działa: policz w Pythonie.
unikalne = sorted(set(v for v in kolumna if v is not None))
for i, wartosc in enumerate(unikalne, start=1):
    ws.cell(row=i, column=9, value=wartosc)

# ==================================================================
# 10. TOKENIZER - INSPEKCJA FORMULY
# ==================================================================
tokeny = Tokenizer("=SUM(A1,B1)+IF(C1>0,1,0)").items
for t in tokeny:
    print(t.type, t.subtype, repr(t.value))
# type: 'OPERAND' | 'FUNC' | 'OP_PRE' | 'OP_IN' | 'OP_POST' | 'WSPACE' | 'SEP' | 'ARRAY'

# Wyodrębnienie nazw funkcji (lepsze niż regex):
funkcje = [t.value.rstrip("(").upper() for t in tokeny if t.type == "FUNC"]

# ==================================================================
# 11. BEZPIECZEŃSTWO - '=' W DANYCH UŻYTKOWNIKA
# ==================================================================
ZNAKI_WYZWALAJACE = ("=", "+", "-", "@")

def jako_tekst(value) -> str:
    """Nie pozwala, żeby treść użytkownika stała się formułą."""
    tekst = "" if value is None else str(value)
    if tekst.startswith(ZNAKI_WYZWALAJACE):
        return "'" + tekst
    return tekst

# Niskopoziomowy trik, gdy MUSISZ zapisać tekst zaczynający się od '=':
komorka = ws.cell(row=1, column=1)
komorka.value = "=SUM(A1:B1)"     # openpyxl ustawi data_type='f'
komorka.data_type = "s"           # <- nadpisujemy na tekst. Sprawdź w swojej wersji.

# ==================================================================
# 12. LICZENIE BEZ EXCELA (LibreOffice, headless)
# ==================================================================
# soffice --headless --convert-to xlsx --outdir output/przel output/plik.xlsx
# (LibreOffice musi mieć włączone przeliczanie przy wczytywaniu)
# Wynik wczytaj z data_only=True. Uwaga na różnice w wynikach - nie do rozliczeń.

# ==================================================================
# 13. RĘCZNE ZAGLĄDANIE DO PLIKU (diagnostyka)
# ==================================================================
with ZipFile("plik.xlsx") as archiwum:
    xml = archiwum.read("xl/worksheets/sheet1.xml").decode("utf-8")
print("formuł w XML:", xml.count("<f>"))       # ile elementów <f>
print("wyników w XML:", xml.count("<v>"))      # ile elementów <v> (cache)
# Formuła bez cache: <f>...</f> bez <v>
# Formuła z cache:   <f>...</f><v>...</v>

# ==================================================================
# 14. TABELA DECYZYJNA - WARTOŚCI CZY FORMULY
# ==================================================================
# WARTOŚCI:  plik czyta program | dokument archiwalny | duże zbiory
#            determinizm | brak Excela po drodze | bezpieczeństwo
# FORMULY:   użytkownik ma edytować dane i widzieć nowe sumy
#            wymóg audytowalności | żywe narzędzie, nie raport
# HYBRYDA:   formuły w arkuszu prezentacji + wartości w arkuszu kontroli
#            + jedna formuła porównująca oba światy
# ZAWSZE:    ustaw `fullCalcOnLoad` = True, gdy w pliku są formuły

# ==================================================================
# 15. SZKIELET SKRYPTU Z FORMULAMI (wzorzec do kopiowania)
# ==================================================================
# 1. KONFIGURACJA: nazwy kolumn, formaty, progi - na górze pliku
# 2. FUNKCJE CZYSTE: walidacja i agregacja w Pythonie (bez openpyxl!)
# 3. BUDOWA ARKUSZA DANYCH: wartości + formuły per wiersz
# 4. BUDOWA ARKUSZA PREZENTACJI: formuły zbudowane f-stringiem
# 5. BUDOWA ARKUSZA KONTROLI: te same liczby jako WARTOŚCI
# 6. wymus_przeliczenie(wb)  <- zawsze, gdy są formuły
# 7. wb.properties.description <- wersja, data, autor
# 8. wb.save(NOWA_NAZWA)     <- nigdy do pliku źródłowego
# 9. AUDYT PO ZAPISIE: policz_formuly() i porównaj z oczekiwaniem
```

## 9. Co dalej

Nauczyłeś się najważniejszego rozróżnienia w całej automatyzacji Excela: **formuła to instrukcja, wartość to wynik** — i to, że openpyxl operuje wyłącznie na instrukcjach, nigdy na wynikach. Teraz wiesz, dlaczego `data_only=True` zwraca `None`, czemu nie wolno łączyć go z `save()` i dlaczego prawdziwa logika biznesowa nie może mieszkać w formule, jeśli Twój program ma cokolwiek z niej wyczytać.

To, co wydaje się tu techniczne, jest w istocie **decyzją architektoniczną**: gdzie jest źródło prawdy o liczbie. Formuła jest źródłem prawdy **tylko dla człowieka, który patrzy na ekran**. Dla systemu prawdą jest ta wartość, którą policzył Python. Trzymanie się tej zasady oszczędzi Ci tysięcy godzin debugowania w kolejnych latach — a początkowe „a może by tak formułą" wraca do ludzi dopiero po pierwszym poważnym błędzie produkcyjnym.

W **module 07** wejdziemy w **siatkę**: wiersze, kolumny i zakresy. Tam czekają trzy tematy, które są bezpośrednio powiązane z tym modułem:

- **`max_row` i `max_column` kłamią.** Rozmiar modelu rośnie nie tylko od danych, ale od formatowania i od „dotkniętych" komórek. Zobaczysz, jak to wykryć i jak znaleźć naprawdę ostatni wiersz danych — a to jest warunek poprawnego budowania **zakresów w formułach**, takich jak `Dane!$F$2:$F$201`, które dziś pisałeś na podstawie założenia.
- **`ws.append()` i jego pułapka.** Najszybszy sposób dopisywania wierszy, ale stringi traktuje jako iterowalne (`append("ABC")` zapisze trzy komórki). Przy generowaniu setek tysięcy wierszy z formułami to narzędzie pierwszej potrzeby — i pierwsze źródło pomyłek.
- **`insert_rows()` i `delete_rows()` nie aktualizują formuł.** Dokładnie ten sam problem, o którym mówiliśmy w sekcji 2.8 przy kopiowaniu: openpyxl nie ma grafu zależności. Wstawienie wiersza nie przesunie `=SUM(A1:A10)` na `=SUM(A1:A11)`. Zobaczysz, jak to obejść — częściowo przez `ws.move_range(..., translate=True)`.

Przygotuj do modułu 07:

- **plik z tego modułu** (`output/06_dwa_swiaty.xlsx`) oraz **jego kopię**, w której dopiszesz kilka formuł w nowym wierszu — posłuży do testowania, co się dzieje z referencjami przy `insert_rows()`,
- **plik, w którym `max_row` jest zaniżony albo zawyżony.** Zbuduj go tak: w arkuszu z danymi w wierszach 1–10 sformatuj (np. pogrub) komórki w wierszach 11–50, ale nic do nich nie wpisuj. Potem sprawdź `ws.max_row`. Zobaczysz liczbę, która nie odpowiada danym — i to jest dokładnie ten problem, który w module 07 rozwiązujemy,
- **odpowiedź na pytanie:** w których miejscach w `06_cw4_raport.xlsx` zbudowałeś zakres formuły „na podstawie założenia" o liczbie wierszy, a w których na podstawie faktycznego rozmiaru danych? Powinieneś umieć wskazać palcem linijkę, w której przyjąłeś `liczba_wierszy + 1`, i wyjaśnić, co by się stało, gdyby Excel dodał do pliku wiersz z formatowaniem powyżej tego zakresu.