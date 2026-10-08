# Moduł 08 — Arkusze: struktura skoroszytu od podszewki

> **Część:** I — Podstawy pracy ze skoroszytem · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–07

## 0. W tym module nauczysz się

- **Rozumiesz, że skoroszyt to uporządkowana lista arkuszy**, a nie worek bez kolejności — i że to właśnie ta lista (kolejność, nazwy, widoczność) decyduje o tym, co użytkownik zobaczy po otwarciu pliku, **zanim jeszcze spojrzy na jakąkolwiek komórkę**.
- **Umiesz tworzyć arkusze z kontrolą pozycji** (`create_sheet(title=..., index=...)`), usuwać arkusz domyślny i wiesz, dlaczego nowy skoroszyt zawsze startuje z jednym arkuszem o nazwie `Sheet`.
- **Znasz twarde reguły nazw arkuszy** — 31 znaków, zakazane znaki, unikalność bez względu na wielkość liter — i masz napisaną funkcję `bezpieczna_nazwa_arkusza()`, która zamienia dowolny napis wejściowy (od użytkownika, z konfiguracji, z nazwy pliku) na poprawną nazwę arkusza: bez wyjątku i bez cichego duplikatu.
- **Odróżniasz `visible`, `hidden` i `veryHidden`** i wiesz, po co istnieje ten trzeci stan: jako miejsce na konfigurację, która nie ma prawa trafić użytkownikowi pod oczy, ale musi jechać w tym samym pliku.
- **Wiesz precyzyjnie, co kopiuje `copy_worksheet`, a co gubi** i dlaczego nie jest to „backup skoroszytu". Zobaczysz to eksperymentalnie: wartości, style, hiperłącza i komentarze przechodzą; tabele, obrazki i wykresy nie.
- **Rozumiesz, co dzieje się z odwołaniami po usunięciu albo zmianie nazwy arkusza** — dlaczego formuły, nazwy zdefiniowane i wykresy w innych arkuszach zostają z numerem telefonu do działu, którego już nie ma.
- **Znasz pierwszy wzorzec architektoniczny tego kursu: „jeden arkusz = jedna odpowiedzialność"**, i potrafisz ułożyć skoroszyt w role `Dane` / `Podsumowanie` / `Wykresy` / `Konfiguracja`, które za kilkanaście modułów staną się naturalnym szkieletem aplikacji generującej raporty.

## 1. Intuicja i analogia

### Analogia główna: skoroszyt to budynek firmy

Wyobraź sobie, że skoroszyt to **siedziba firmy**, a każdy arkusz to **osobne pomieszczenie**. Otwierasz plik — wchodzisz do budynku.

Zbudujmy słownik pojęć:

- **Arkusz to pomieszczenie.** Ma swoje przeznaczenie: magazyn, sala konferencyjna, archiwum, pokój techniczny.
- **Nazwa arkusza to tabliczka na drzwiach.** I tu od razu pojawia się pierwszy problem praktyczny: tabliczki mają ograniczenia. Nie mogą być dowolnie długie (31 znaków — wyobraź sobie, że na tabliczce fizycznie zmieści się tylko tyle), nie mogą zawierać pewnych znaków (`: \ / ? * [ ]` — bo te znaki mają w firmie inne znaczenie i na przykład w adresie wewnętrznym `Dane:2026` dwukropek oznacza coś zupełnie innego), a **dwóch pomieszczeń nie można nazwać tak samo**. Nawet „Sala" i „sala" to dla budynku ta sama nazwa — tabliczek nie odróżnia wielkość liter.
- **Kolejność arkuszy to kolejność drzwi na korytarzu.** Wchodzisz i idziesz po kolei: pierwszy arkusz, drugi arkusz, trzeci. Ta kolejność jest tym, co użytkownik zobaczy jako zakładki na dole ekranu.
- **Aktywny arkusz to pomieszczenie, do którego wchodzisz od razu po wejściu do budynku.** Możesz stać na korytarzu i widzieć wszystkie drzwi, ale fizycznie jesteś w jednym pokoju — tym, który openpyxl zapisał jako aktywny. To ma znaczenie: jeśli raport ma otwierać się na podsumowaniu, a nie na surowych danych, ustawiasz właśnie to.
- **Kolor zakładki to kolor tabliczki na drzwiach.** Nie zmienia zawartości, ale pozwala na pierwszy rzut oka zorientować się w budynku: niebieski = dane, zielony = podsumowanie, szary = pomieszczenie techniczne.
- **`visible` / `hidden` / `veryHidden` to trzy stany drzwi.**

### Analogia do widoczności: drzwi, drzwi za parawanem i wejście w ścianie

To rozróżnienie bywa mylone, a ma konkretne konsekwencje, więc rozłóżmy je na trzy obrazy:

- **`visible`** — normalne drzwi na korytarzu. Widzisz je, wchodzisz, wszystko jest oczywiste.
- **`hidden`** — drzwi zasłonięte parawanem. Idąc korytarzem nie zauważysz, że tam są. Ale **każdy, kto wie o ich istnieniu, może odsunąć parawan** — w Excelu to jedno kliknięcie prawym przyciskiem na zakładkę i wybranie „Pokaż". To pomieszczenie ukryte przed przypadkowym spojrzeniem, nie przed człowiekiem, który przyszedł szukać.
- **`veryHidden`** — wejście ukryte w ścianie, które udaje fragment muru. Nie ma go na planie budynku, nie ma numeru, nie ma klamki. **Nie da się go odsłonić przez interfejs Excela.** Żeby tam wejść, trzeba znać plan — czyli mieć narzędzie programistyczne (albo makro VBA).

I teraz najważniejsza uczciwa uwaga tego podrozdziału: **`veryHidden` to nie zabezpieczenie.** To parawan lepszej jakości. Każdy, kto umie rozpakować plik `.xlsx` (a przypominam z modułu 02: to zwykły ZIP z plikami XML), zobaczy zawartość arkusza `veryHidden` w pliku `xl/worksheets/`. Jeśli Twoją intencją jest **ukrycie czegoś przed kimś, kto wie, co robi** — nie użyjesz do tego arkusza ukrytego. Użyjesz uprawnień, szyfrowania albo po prostu nie umieścisz tych danych w pliku. To rozróżnienie wróci w module 17 (ochrona arkusza) i w module 21 (bezpieczeństwo).

### Analogia do kopiowania arkusza: fotokopia wyposażenia, nie fotokopia pomieszczenia

`copy_worksheet` działa tak, jakbyś chciał skopiować pomieszczenie, robiąc **fotokopię wszystkiego, co w nim leży, i przenosząc kopię do nowego pokoju**.

- **Meble, dokumenty, segregatory, podpisy na dokumentach** — czyli wartości komórek, style, hiperłącza, komentarze — przenoszą się bez problemu. To rzeczy fizycznie leżące w pomieszczeniu.
- **Obraz na ścianie i instalacja elektryczna** — czyli obrazki i wykresy — **nie przenoszą się**. Bo nie są „w pomieszczeniu". Są zamontowane w budynku i przypisane do konkretnej lokalizacji drzwiami i numerem. Nowy pokój ich nie dostanie, bo nikt nie zamontował ich od nowa.

To nie jest niedoróbka ani błąd do obejścia. To konsekwencja tego, jak zbudowany jest plik `.xlsx`: obrazki i wykresy to **osobne obiekty z własnymi relacjami w pakiecie** (`xl/drawings/`, `xl/charts/`), a nie część „treści pomieszczenia". openpyxl kopiuje zawartość komórek i wybrane ustawienia arkusza — i na tym poprzestaje. Dokumentacja openpyxl 3.1 mówi to wprost:

> „Tylko komórki (włącznie z wartościami, stylami, hiperłączami i komentarzami) oraz pewne atrybuty arkusza (włącznie z wymiarami, formatem i właściwościami) są kopiowane. Wszystkie pozostałe atrybuty skoroszytu i arkusza **nie** są kopiowane — na przykład obrazy i wykresy."

Zapamiętaj to zdanie w tej formie, bo prędzej czy później ktoś Ci powie: „zrób po prostu `copy_worksheet`, to będzie kopia zapasowa". **Nie będzie.** Zobaczysz to na własne oczy w Przykładzie 3.

### Analogia do usuwania arkusza: wizytówki rozdane klientom

Ostatnia analogia, potrzebna dla tematu najgroźniejszego w całym module.

Wyobraź sobie, że firma likwiduje dział — powiedzmy „Dział zamówień". **Usuwasz go z wewnętrznego spisu telefonów.** W budynku wszystko jest spójne: nie ma już tego działu, nikt tam nie dzwoni, porządek.

Ale:

- **Wizytówki rozdane klientom w zeszłym roku** nadal mają numer wewnętrzny 214. Klient dzwoni — nikt nie odbiera, bo numer nie istnieje.
- **Listy przewodnie w mailach** nadal mówią „szczegóły w załączniku od Działu zamówień".
- **Instrukcje obsługi** w innych działach nadal w punkcie 3 odsyłają do „formularza Z-14 z Działu zamówień".

Usunięcie wpisu z jednego spisu **nie usuwa wszystkich odwołań do niego w firmie**. I dokładnie to samo dzieje się w skoroszycie: `del wb["Dane"]` usuwa arkusz, ale **formuła w arkuszu `Podsumowanie`, która mówiła `=SUM(Dane!B2:B13)`, nadal tam jest**. Nazwa zdefiniowana wskazująca na `Dane!$A$2:$A$100` nadal istnieje. Wykres w trzecim arkuszu nadal pobiera dane z zakresu, którego już nie ma.

Excel, otwierając taki plik, pokaże `#REF!` — i to jest scenariusz **najlepszy**, bo widzisz problem. Gorszy scenariusz to taki, w którym odwołanie przypadkiem trafia na **inną** komórkę i pokazuje liczbę, która wygląda wiarygodnie, a jest nieprawdziwa. To najgroźniejszy typ błędu w raportach: cichy i przekonujący.

### Analogia do zmiany nazwy arkusza: ta sama pułapka w tańszym opakowaniu

I jeszcze jedno, zanim przejdziemy do teorii, bo to jest sedno i najczęściej pomijana rzecz w tym module:

**Zmiana nazwy arkusza jest tak samo niebezpieczna jak jego usunięcie.** Zmieniasz tabliczkę na drzwiach, ale wizytówki rozdane klientom nadal podają starą nazwę. `ws.title = "Sprzedaz"` nie zaktualizuje ani jednej formuły, ani jednej nazwy zdefiniowanej. Do tego wrócimy w sekcji 2.9.

### Analogia do kolejności: korytarz, na którym kolejność drzwi ma znaczenie

Ostatnia krótka analogia, żeby domknąć obraz. Kolejność arkuszy nie jest kosmetyką — to forma opowiadania. Czytelnik raportu idzie korytarzem i pierwsze drzwi, które napotyka, kształtują jego pierwsze wrażenie. Skoroszyt z kolejnością `_Konfiguracja`, `Dane`, `Podsumowanie` otwiera się na pomieszczeniu technicznym — czyli na tym, co w firmie najnudniejsze i najbardziej techniczne. Ta sama zawartość w kolejności `Podsumowanie`, `Dane`, `_Konfiguracja` (ukryte) opowiada zupełnie inną historię: „oto wynik, oto źródło, oto szczegóły techniczne, których nie musisz oglądać".

W automatyzacji kolejność łatwo uznać za drobiazg. W praktyce jest to jeden z najtańszych sposobów na to, żeby raport był lepszy. Wrócimy do tego w sekcji 2.11.

## 2. Teoria

### 2.1. Skoroszyt to lista arkuszy, nie worek

Zacznijmy od modelu, który trzeba mieć w głowie przez cały ten moduł.

Obiekt `Workbook` w openpyxl trzyma **jedną, uporządkowaną listę arkuszy**. Ta lista ma kolejność (bo korytarz ma kierunek) i jest podstawą wszystkiego, co zobaczysz w tym module.

Trzy najważniejsze sposoby patrzenia na tę listę:

```python
from openpyxl import Workbook

wb = Workbook()

# 1. Nazwy w kolejnosci - to najczesciej to, czego potrzebujesz
print(wb.sheetnames)          # ['Sheet']

# 2. Obiekty arkuszy, w kolejnosci
print(wb.worksheets)          # [<Worksheet "Sheet">]

# 3. Arkusze wykresow (rzadkie, sekcja 2.10) - osobna lista
print(wb.chartsheets)         # []
```

**Co się dzieje w pamięci.** `Workbook()` tworzy jedną listę (`wb._sheets`) i wkłada do niej jeden świeżo utworzony obiekt `Worksheet`. `wb.sheetnames` to wyliczana właściwość, która zwraca `[arkusz.title for arkusz in wb._sheets]`. Nie ma tu żadnej magii: nazwy arkuszy to po prostu tytuły obiektów na liście.

**Co trafi do pliku.** Kolejność tej listy i tytuły arkuszy trafiają do `xl/workbook.xml` jako sekwencja elementów `<sheet name="..." sheetId="..." r:id="..."/>` — i to jest **jedno miejsce w całym pliku**, z którego Excel czyta porządek zakładek.

Dostęp do konkretnego arkusza:

```python
ws = wb["Dane"]               # po nazwie, dokladne dopasowanie (patrz nizej!)
ws = wb.worksheets[0]         # po pozycji - pierwszy arkusz
pozycja = wb.index(ws)        # pozycja arkusza na liscie (0-based)

for arkusz in wb.worksheets:  # iteracja po obiektach, w kolejnosci
    print(arkusz.title)
```

**Uwaga na dopasowanie po nazwie.** `wb["Dane"]` wymaga **dokładnego** dopasowania, z uwzględnieniem wielkości liter. `wb["dane"]` podniesie `KeyError`, mimo że arkusz `Dane` istnieje. To jest nieintuicyjne, bo Excel w interfejsie jest w tym miejscu **niewrażliwy** na wielkość liter — jeśli wpiszesz w formule `dane!A1` a arkusz nazywa się `Dane`, Excel to przyjmie. openpyxl nie. Wrócimy do tego w pułapkach.

**Konsekwencja praktyczna.** Gdy piszesz kod, który odnosi się do arkusza po nazwie, **nigdy nie zakładaj nazwy na sztywno**. Nazwa arkusza jest danymi wejściowymi, tak samo jak wartość w komórce. W modułach 22–26 zobaczysz, jak to wygląda w architekturze: nazwy arkuszy staną się częścią konfiguracji raportu, a nie literałami rozsianymi po kodzie.

### 2.2. Tworzenie arkuszy i arkusz domyślny

Nowy skoroszyt w openpyxl **nigdy nie jest pusty** — zawsze ma jeden arkusz o nazwie `Sheet`:

```python
from openpyxl import Workbook

wb = Workbook()
print(wb.sheetnames)      # ['Sheet']
print(wb.active.title)    # 'Sheet'

del wb["Sheet"]           # usuwamy domyslny arkusz
print(wb.sheetnames)      # []
```

**Co się dzieje w pamięci.** `Workbook()` tworzy jedną listę i wstawia do niej jeden obiekt `Worksheet`. Ten obiekt ma pusty słownik komórek (`_cells == {}`). Ciekawostka techniczna spójna z modułem 07: `ws.max_row` na takim świeżym arkuszu zwraca `1`, mimo że komórek nie ma — bo minimalną wartością jest `1`. To ta sama rodzina nieporozumień, którą poznałeś w module 07.

**Co trafi do pliku.** Jeśli zapiszesz skoroszyt po `del wb["Sheet"]` bez dodania nowego arkusza, dostaniesz plik, który Excel będzie traktował jako uszkodzony albo pusty. Praktyczna zasada: **arkusz domyślny usuwaj dopiero po utworzeniu tych, które faktycznie potrzebujesz.** Najlepiej w kolejności „najpierw stwórz, potem usuń":

```python
wb = Workbook()
dane = wb.active              # wykorzystaj domyslny arkusz jako roboczy
dane.title = "Dane"           # <- to jest kluczowy trick: zmiana nazwy

podsumowanie = wb.create_sheet("Podsumowanie")
konfiguracja = wb.create_sheet("_Konfiguracja")
```

Zmiana nazwy domyślnego arkusza jest lepsza niż jego usuwanie i tworzenie nowego, bo **nie zmienia pozycji** i nie wymaga pilnowania, żeby skoroszyt nie pozostał bez arkuszy.

`create_sheet` w pełnej formie:

```python
ws = wb.create_sheet(title="Dane", index=0)     # na poczatek
ws = wb.create_sheet(title="Dane")              # na koniec (domyslnie)
ws = wb.create_sheet()                          # nazwa automatyczna
```

Parametry:

- `title` — nazwa arkusza. **Zawsze podawaj ją jawnie.** Bez tego dostaniesz automatyczną nazwę, która nie mówi nic ani Tobie, ani użytkownikowi.
- `index` — pozycja, na którą arkusz ma trafić. `0` to początek listy. **Jeśli `index` wykracza poza długość listy, openpyxl nie zgłosi błędu — arkusz trafi na koniec.** To cicha pułapka: kod `wb.create_sheet("Podsumowanie", 99)` w skoroszycie z trzema arkuszami wygląda jak „umieść obok końca", a w praktyce znaczy to samo co brak `index`.

**Konsekwencja praktyczna.** Buduj skoroszyt w kolejności docelowej, a nie w kolejności technicznej. Jeśli się pomylisz, kolejność naprawisz przez `wb.move_sheet` (sekcja 2.4) — ale łatwiej jest od razu myśleć korytarzem.

### 2.3. Nazwy arkuszy — reguły twarde i miękkie

To najbardziej „papierkowa" sekcja tego modułu, ale też jedna z najbardziej praktycznych. Nazwa arkusza bywa danymi wejściowymi: pochodzi z nazwy pliku, z konfiguracji, od użytkownika, z wpisu w bazie. Za każdym razem musi przejść przez te same reguły.

**Reguła 1: nazwa musi być napisem.** Liczba, `None`, obiekt — to nie zadziała.

**Reguła 2: maksymalnie 31 znaków.** Przekroczenie podnosi błąd (w 3.1.x komunikat mówi, że maksimum to 31 znaków — sprawdź dokładne brzmienie w swojej wersji).

**Reguła 3: zakazane znaki to `: \ / ? * [ ]`.** Każdy z nich w nazwie podnosi `ValueError` z komunikatem wskazującym znak. Dlaczego akurat te? Bo mają znaczenie w innych kontekstach pliku `.xlsx` — dwukropek rozdziela arkusz od zakresu w odwołaniach, nawiasy kwadratowe oznaczają odwołania do tabeli strukturalnej, gwiazdka i znak zapytania są symbolami wieloznacznymi, ukośniki są separatorami ścieżek. Arkusz o nazwie `Q1/2026` złamałby gramatykę odwołań — i dlatego Excel zabrania tego u źródła.

**Reguła 4: nazwy są unikalne bez względu na wielkość liter.** Nie możesz mieć arkuszy `Dane` i `dane` w jednym skoroszycie. Excel traktuje je jako tę samą nazwę — i openpyxl też.

**I teraz rzecz, która zaskakuje: openpyxl nie podniesie błędu przy duplikacie.** Zamiast tego **zmieni nazwę** na unikalną, dopisując sufiks liczbowy:

```python
from openpyxl import Workbook

wb = Workbook()
wb["Sheet"].title = "Dane"
print(wb.sheetnames)                 # ['Dane']

nowy = wb.create_sheet("Dane")       # <- duplikat!
print(wb.sheetnames)                 # ['Dane', 'Dane1']
print(nowy.title)                    # 'Dane1'
```

To zachowanie realizuje funkcja pomocnicza `avoid_duplicate_name` z `openpyxl.utils` — opisana w kodzie źródłowym jako „naiwna" (bo sprawdza nazwy prostym porównaniem i dokleja liczbę). **Nie zakładaj, że publiczne API tej funkcji jest gwarantowane** — w kodzie produkcyjnym napisz własny mechanizm unikalności, taki jak w sekcji niżej. Dzięki temu masz kontrolę nad tym, jak wygląda sufiks (i czy w ogóle ma być sufiks, czy raczej wyjątek).

**Reguła 5: nazwa nie może być pusta** i nie powinna zaczynać się ani kończyć apostrofem — Excel tego nie lubi, bo apostrof jest znakiem cytowania nazw w formułach.

**Reguła 6 (kulturalna, nie twarda): nazwa ma być czytelna dla człowieka.** `Arkusz1`, `Sheet2`, `Tabela przestawna 3` to nazwy, które nic nie komunikują. `Dane`, `Podsumowanie`, `Metodyka`, `_Konfiguracja` komunikują rolę w jednym słowie.

Zbierzmy to w jedną funkcję, której użyjesz w każdym projekcie:

```python
"""Bezpieczne nazwy arkuszy - jedna funkcja, ktora zamyka caly temat."""

from __future__ import annotations

import re

NIEDOZWOLONE = set('[]:*?/\\')
MAKS_DLUGOSC = 31
ZAPASOWA_NAZWA = "Arkusz"


def _zapewnij_unikalnosc(nazwa: str, zajete: set[str], maks: int = MAKS_DLUGOSC) -> str:
    """Dokleja suffix _2, _3... az nazwa bedzie unikalna (bez wzgledu na wielkosc liter).

    Uwaga: suffix doklejamy PO obcieciu do limitu, dlatego skracamy baze,
    a nie gotowa nazwe - inaczej przekroczylibysmy 31 znakow.
    """
    zajete_lower = {n.lower() for n in zajete}
    if nazwa.lower() not in zajete_lower:
        return nazwa

    licznik = 2
    while licznik < 10_000:
        suffix = f"_{licznik}"
        baza = nazwa[: maks - len(suffix)]
        kandydat = f"{baza}{suffix}"
        if kandydat.lower() not in zajete_lower:
            return kandydat
        licznik += 1
    raise RuntimeError(f"nie udalo sie znalezc unikalnej nazwy dla {nazwa!r}")


def bezpieczna_nazwa_arkusza(
    nazwa: object,
    zajete: set[str] | None = None,
    maks: int = MAKS_DLUGOSC,
) -> str:
    """Zamienia dowolny napis w poprawna i (opcjonalnie) unikalna nazwe arkusza.

    Kolejnosc krokow:
      1. rzutowanie na napis i przyciecie bialych znakow z brzegow,
      2. zamiana znakow niedozwolonych na podkreslenie,
      3. zamiana znakow kontrolnych na podkreslenie,
      4. obciecie do limitu dlugosci,
      5. usuniecie apostrofow z brzegow,
      6. zapasowa nazwa, jesli nic nie zostalo,
      7. zapewnienie unikalnosci wzgledem podanego zbioru nazw.
    """
    if nazwa is None:
        tekst = ""
    else:
        tekst = str(nazwa).strip()

    # 2. znaki niedozwolone -> podkreslenie (nie usuwamy, bo to zmienia sens nazwy)
    tekst = "".join("_" if znak in NIEDOZWOLONE else znak for znak in tekst)

    # 3. znaki kontrolne (tabulatory, nowe linie itp.) -> podkreslenie
    tekst = re.sub(r"[\x00-\x1f]", "_", tekst)

    # 4. limit dlugosci
    if len(tekst) > maks:
        tekst = tekst[:maks]

    # 5. apostrofy na brzegach sa w Excelu klopotliwe
    tekst = tekst.strip("'")

    # 6. nie zostaw pustej nazwy
    if not tekst:
        tekst = ZAPASOWA_NAZWA

    # 7. unikalnosc
    if zajete:
        tekst = _zapewnij_unikalnosc(tekst, zajete, maks)

    return tekst
```

**Co się dzieje w pamięci.** Funkcja nie dotyka ani skoroszytu, ani pliku. To czysta funkcja napis → napis. Można ją testować bez openpyxl, co jest dużą zaletą (moduł 20).

**Co trafi do pliku.** Nazwa, którą przypiszesz do `ws.title`, trafi do `xl/workbook.xml` i będzie widoczna jako zakładka. Jeśli wcześniej przeszła przez `bezpieczna_nazwa_arkusza`, wiesz, że nie wywali się na limicie ani na znaku specjalnym — a to dokładnie te dwa błędy, które pojawiają się w produkcji o drugiej w nocy, gdy klient wyśle plik o nazwie `Raport / Q1 (kopia "final")`.

**Konsekwencja praktyczna — jedna reguła.** Każda nazwa arkusza, która **nie** jest literałem w Twoim kodzie, musi przejść przez `bezpieczna_nazwa_arkusza`. I uwaga na subtelność: funkcja dba też o **unikalność**, więc kolejność ma znaczenie — nazwy dodawaj po jednej, przekazując aktualny zbiór `zajete`:

```python
zajete: set[str] = set()
for surowa_nazwa in ["Dane", "dane", "Raport/Q1", "  ", "Dane"]:
    nazwa = bezpieczna_nazwa_arkusza(surowa_nazwa, zajete=zajete)
    zajete.add(nazwa)
    print(f"{surowa_nazwa!r:20} -> {nazwa!r}")
```

Wynik zobaczysz w Przykładzie 2, razem z wyjaśnieniem każdej decyzji.

### 2.4. Kolejność arkuszy — jak ją czytać i jak ją zmieniać

Czytanie kolejności jest proste i zawsze bezpieczne:

```python
wb.sheetnames        # lista nazw, w kolejnosci zakladek
wb.worksheets        # lista obiektow, w kolejnosci zakladek
wb.index(ws)         # pozycja arkusza na liscie (0-based)
```

Zmiana kolejności to już inna historia, bo openpyxl daje tu **dwa poziomy dostępu** — publiczny i prywatny.

**Poziom publiczny: `wb.move_sheet(ws, offset=...)`.** Przesuwa arkusz o zadaną liczbę pozycji:

```python
wb.move_sheet(wb["Podsumowanie"], offset=-1)   # o jedna pozycje w lewo
wb.move_sheet(wb["_Konfiguracja"], offset=2)   # o dwie pozycje w prawo
```

Typowe użycie to ustawienie konkretnej pozycji, a nie „przesunięcie o":

```python
def przestaw_na_pozycje(wb, nazwa: str, pozycja: int) -> None:
    """Ustawia arkusz na dokladnej pozycji, korzystajac tylko z API publicznego."""
    ws = wb[nazwa]
    offset = pozycja - wb.index(ws)
    if offset:
        wb.move_sheet(ws, offset=offset)
```

**Poziom prywatny: `wb._sheets`.** Niektóre przykłady w internecie przestawiają arkusze, przypisując bezpośrednio do tej listy:

```python
# TAK NIE ROBImy - to implementacja prywatna
wb._sheets = [wb[n] for n in ["Podsumowanie", "Dane", "_Konfiguracja"]]
```

**Dlaczego to zły pomysł.** `wb._sheets` to **implementacja**, nie kontrakt. Nazwa zaczyna się od podkreślenia właśnie po to, żeby było jasne, że autorzy biblioteki nie zamierzają jej utrzymywać w niezmienionej formie. Przestawiając listę ręcznie omijasz też wszelkie sprawdzenia, które `move_sheet` wykonuje. Jeśli w dokumentacji Twojej wersji `move_sheet` ma wariant przyjmujący nazwę arkusza jako `offset` (a w gałęzi 3.1 pojawiały się rozszerzenia tego API) — używaj właśnie jego. Jeśli nie — napisz własną funkcję taką jak `przestaw_arkusze` w Ściągawce, opartą wyłącznie na `wb.index` i `wb.move_sheet`.

**Co się dzieje w pamięci.** `move_sheet` wycina obiekt arkusza z listy i wstawia go na nowej pozycji. Komórki, style i wszystko inne są w tym obiekcie nienaruszone — przenosi się cały obiekt, nie jego zawartość.

**Co trafi do pliku.** Kolejność elementów `<sheet>` w `xl/workbook.xml`. Nic więcej się nie zmienia: arkusze nie mają „zapisanej pozycji" w swoich własnych plikach XML, ich pozycja wynika wyłącznie z kolejności na liście skoroszytu. To ważne: **przestawienie arkusza jest operacją całkowicie bezpieczną** — nie psuje formuł ani odwołań, bo nazwy się nie zmieniają.

**Konsekwencja praktyczna.** Kolejność ustawiaj po zbudowaniu skoroszytu, w jednym miejscu, jedną funkcją. Nie próbuj „budować w kolejności" przez rozsypane `wb.create_sheet(..., index=...)` w dziesięciu miejscach kodu.

### 2.5. Widoczność: `visible`, `hidden`, `veryHidden`

Stan widoczności ustawiasz w jednej linii:

```python
ws.sheet_state = "visible"      # domyslnie
ws.sheet_state = "hidden"       # zakladka ukryta, ale mozna ja odkryc w Excelu
ws.sheet_state = "veryHidden"   # zakladka ukryta i niedostepna z interfejsu
```

**Co się dzieje w pamięci.** `sheet_state` jest właściwością arkusza i trafia do listy przy zapisie. Nie ma tu żadnej logiki walidującej — openpyxl nie sprawdzi za Ciebie, czy przypadkiem nie próbujesz ukryć ostatniego widocznego arkusza ani czy nazwa stanu jest poprawna.

**Co trafi do pliku.** Atrybut `state` w elemencie `<sheet>` w `xl/workbook.xml`:

```xml
<sheet name="_Konfiguracja" sheetId="3" state="veryHidden" r:id="rId4"/>
```

Brak atrybutu `state` oznacza `visible`. Cała „ukrytość" arkusza polega na tym jednym słowie.

**Trzy twarde zasady, których Excel pilnuje:**

1. **Skoroszyt musi mieć co najmniej jeden widoczny arkusz.** Nie da się ukryć wszystkich. Jeśli spróbujesz, openpyxl podniesie błąd — dokładne miejsce i komunikat sprawdź w swojej wersji, ale logika jest niezmienna: Excel nie otworzy skoroszytu bez widocznej zakładki.
2. **Aktywny arkusz musi być widoczny.** Ustawienie `wb.active` na arkusz ukryty kończy się `ValueError`. Kolejność operacji ma znaczenie: najpierw ustaw nowy aktywny arkusz, potem ukryj stary.
3. **Stan arkusza ma znaczenie dla kopiowania.** Kopiowanie arkusza `veryHidden` jest traktowane inaczej niż arkusza `visible` (szczegóły w sekcji 2.8), a w niektórych wydaniach openpyxl odmawia wykonania tej operacji.

Bezpieczna sekwencja ukrywania arkusza wygląda więc tak:

```python
def ukryj_arkusz(wb, nazwa: str, stan: str = "veryHidden") -> None:
    """Ukrywa arkusz, dbajac o to, zeby pozostal co najmniej jeden widoczny."""
    if stan not in {"hidden", "veryHidden"}:
        raise ValueError(f"nieznany stan: {stan!r}")

    kandydat = wb[nazwa]
    widoczne = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]

    # Jesli to ostatni widoczny arkusz - najpierw trzeba wskazac innego,
    # ale jesli nie ma innego, trzeba odmowic.
    if len(widoczne) == 1 and widoczne[0] is kandydat:
        raise ValueError("nie mozna ukryc ostatniego widocznego arkusza")

    # Jesli ukrywamy arkusz aktualnie aktywny - przelacz aktywnosc.
    if wb.active is kandydat:
        zapas = next(ws for ws in widoczne if ws is not kandydat)
        wb.active = wb.index(zapas)

    kandydat.sheet_state = stan
```

**Kiedy używać którego stanu — praktyczne kryteria:**

| Stan | Kiedy | Przykład |
|---|---|---|
| `visible` | Zawsze domyślnie, dla wszystkiego, co użytkownik ma zobaczyć | `Dane`, `Podsumowanie` |
| `hidden` | Rzadko. Gdy chcesz ukryć arkusz przed „przypadkowym" użytkownikiem, ale dopuszczasz jego odkrycie | Arkusz pomocniczy, tabela przestawna źródłowa |
| `veryHidden` | Gdy arkusz jest **konfiguracją techniczną**, a nie treścią | `_Konfiguracja`, `_Slowniki`, `_Metodyka` |

I uczciwie: **`veryHidden` jest użyteczne, ale nie jest bezpieczne.** Chroni przed „niechcący kliknąłem na zakładkę". Nie chroni przed nikim, kto umie rozpakować plik albo napisać skrypt. Trzy tysiące lat temu ktoś powiedziałby: to nie sejf, to zasłona. To ma sens jako zasłona.

### 2.6. Kolor zakładki i drobne ustawienia arkusza

Kolor zakładki nadaje się jedną linią:

```python
ws.sheet_properties.tabColor = "2F5597"
```

Wartość to kolor w zapisie szesnastkowym. openpyxl przyjmuje 6-znakowy zapis RGB (jak wyżej, to wprost z dokumentacji openpyxl) i zapisze go jako kolor 8-znakowy z kanałem alfa. Możesz też podać jawnie obiekt `Color`:

```python
from openpyxl.styles import Color
ws.sheet_properties.tabColor = Color(rgb="002F5597")
```

**Analogia:** to kolor tabliczki na drzwiach. Nie zmienia treści, ale w budynku z dwudziestoma pomieszczeniami jest to najtańszy sposób, żeby ktoś zorientował się, gdzie jest.

**Co się dzieje w pamięci.** `ws.sheet_properties` to mały obiekt z kilkoma właściwościami, w tym `tabColor` i `codeName`. Kolor jest na nim zapisany, dopóki nie zapiszesz skoroszytu.

**Co trafi do pliku.** Element `<sheetPr>` w pliku XML arkusza (`xl/worksheets/sheetN.xml`), a konkretnie `<tabColor rgb="FF2F5597"/>`. Cały mechanizm jest bardzo „papierowy" — to jedna wartość w jednym miejscu.

**Ustal konwencję kolorów i zapisz ją raz.** To nie jest przesada — to ten sam mechanizm, który w module 23 nazwiemy rejestrem, a w module 26 konfiguracją:

```python
KOLORY_ZAKLADEK = {
    "dane":          "2F5597",   # niebieski - zrodlo
    "podsumowanie":  "548235",   # zielony - wynik
    "wykresy":       "BF8F00",   # zloty - wizualizacja
    "metodyka":      "7F7F7F",   # szary - opis
    "konfiguracja":  "C00000",   # czerwony - nie dotykac
}
```

Użytkownik raportu uczy się tej konwencji po dwóch plikach i potem orientuje się w nich bezwysilkowo. Koszt: dziesięć minut na początku projektu. Zysk: setki przeczytanych raportów bez pytania „gdzie są dane".

Kilka pozostałych ustawień arkusza, które warto znać na tym etapie:

```python
ws.sheet_format.defaultColWidth = 14      # domyslna szerokosc kolumn (znaki)
ws.sheet_format.defaultRowHeight = 18     # domyslna wysokosc wierszy (punkty)
ws.sheet_view.zoomScale = 80              # powiekszenie 10..400 (procent)
ws.sheet_properties.codeName = "ArkuszDane"   # identyfikator dla VBA - rzadko potrzebny
```

`zoomScale` bywa niedoceniany. Raport z 20 kolumnami otwiera się w Excelu w powiększeniu 100% i użytkownik widzi pierwszą połowę tabeli, nie wiedząc, że jest druga. Ustawienie `80` sprawia, że raport otwiera się **w całości widoczny** — i to jest jedna liczba, która zmienia pierwszą sekundę kontaktu użytkownika z plikiem. Pełny zestaw ustawień widoku i druku omówimy w module 11.

### 2.7. Aktywny arkusz i widoki

Dwa pojęcia, które wyglądają podobnie, a znaczą co innego:

- **Aktywny arkusz** (`wb.active`) — ten, który jest wyświetlony po otwarciu skoroszytu. To on decyduje o pierwszym wrażeniu.
- **Zaznaczona zakładka** (`ws.sheet_view.tabSelected`) — dotyczy grupowego zaznaczenia wielu zakładek, gdy wpisujesz dane do kilku arkuszy jednocześnie. W automatyzacji raportów ustawiasz to rzadko.

W praktyce znaczenie ma pierwsze. Ustawiasz je przez `wb.active`:

```python
wb.active = 0                    # po indeksie (0-based)
wb.active = wb.index(ws_pod)     # to samo, ale bez liczenia recznie
```

**Co się dzieje w pamięci.** `wb.active` to właściwość wyliczana z indeksu (`wb._active_sheet_index`): getter zwraca `_sheets[indeks]`. Setter przyjmuje liczbę (indeks) — w gałęzi 3.1 pojawił się także wariant przyjmujący nazwę arkusza; jeśli chcesz go użyć, sprawdź w dokumentacji swojej wersji, jaki typ przyjmuje i jaki błąd podnosi przy nieprawidłowej wartości.

**Co trafi do pliku.** Element `<bookViews><workbookView activeTab="N"/></bookViews>` w `xl/workbook.xml`. To jedyne miejsce, w którym zapisane jest „na którym arkuszu otwiera się plik".

**Trzy rzeczy, o których trzeba pamiętać:**

1. **Ustawiaj `wb.active` na końcu budowania skoroszytu**, gdy wszystkie arkusze już istnieją. Jeśli ustawisz indeks, a potem coś usuniesz albo przestawisz, indeks może wskazać na inny arkusz — albo na żaden.
2. **Aktywny arkusz musi być widoczny.** Jeśli najpierw ustawisz `wb.active` na arkusz, a potem ukryjesz ten arkusz, dostaniesz błąd. Kolejność: ustaw nowy aktywny, potem ukryj stary.
3. **`wb.active` w świeżym skoroszycie to pierwszy arkusz.** Jeśli nie ustawiasz nic, użytkownik otworzy plik na arkuszu, który przypadkiem znalazł się na pozycji `0`.

Ostatnia rzecz: weryfikacja po odczycie. Gdy wczytasz plik, `wb.active` zwraca obiekt aktywnego arkusza — możesz sprawdzić, czy Twój wybór przetrwał zapis i odczyt:

```python
wb2 = load_workbook("output/raport.xlsx")
print(wb2.active.title)             # nazwa arkusza, ktory sie otworzy
print(wb2.index(wb2.active))        # jego pozycja
```

### 2.8. Kopiowanie arkuszy — `wb.copy_worksheet(ws)`

Zacznijmy od tego, co najczęściej jest pomijane: **`copy_worksheet` jest metodą obiektu `Workbook`, a nie arkusza.**

```python
kopia = wb.copy_worksheet(wb["Dane"])    # ✅ tak
# wb["Dane"].copy_worksheet()            # ❌ tak nie ma
```

Metoda zwraca **nowy obiekt `Worksheet`**, który jest już dodany do skoroszytu (trafia na koniec listy arkuszy). Tytuł kopii openpyxl buduje automatycznie na podstawie tytułu źródła, dbając o unikalność. Jeśli zależy Ci na konkretnej nazwie, zmień ją po kopiowaniu:

```python
kopia = wb.copy_worksheet(wb["Dane"])
kopia.title = bezpieczna_nazwa_arkusza("Dane - archiwum", zajete=set(wb.sheetnames))
```

**Warunki wstępne — kopiowanie nie zadziała, jeśli:**

- skoroszyt jest w trybie `read_only=True` — nie ma struktury do kopiowania,
- skoroszyt jest w trybie `write_only=True` — kopiowanie wymaga pełnego modelu w pamięci,
- chcesz kopiować **między skoroszytami** — nie jest to obsługiwane; kopiowanie działa wyłącznie w obrębie jednego `Workbook`,
- arkusz źródłowy **nie ma ani jednej komórki** — openpyxl traktuje taki arkusz jako „pusty" i odmawia kopiowania (podnosi wyjątek; sprawdź jego typ w swojej wersji).

**Inwentarz: co przechodzi na kopię, a co nie.**

| Element arkusza | Kopiowany? |
|---|---|
| Wartości komórek | ✅ |
| Typy danych (`data_type`) | ✅ |
| Style komórek (font, wypełnienie, obramowanie, wyrównanie) | ✅ |
| Formaty liczb (`number_format`) | ✅ |
| Hiperłącza | ✅ |
| Komentarze | ✅ |
| Wymiary kolumn i wierszy | ✅ |
| `sheet_format` i `sheet_properties` | ✅ |
| Tekst formuły | ✅, ale **referencje nie są tłumaczone** |
| Scalenia | ⚠️ sprawdź i przetestuj w swojej wersji |
| Freeze panes, widok, ustawienia wydruku | ⚠️ sprawdź i przetestuj w swojej wersji |
| **Tabele** (`ws.tables`) | ❌ |
| **Obrazy** (`ws._images`) | ❌ |
| **Wykresy** | ❌ |
| Formatowanie warunkowe | ❌ |
| Walidacja danych | ❌ |
| Nazwy zdefiniowane (globalne) | ❌ — należą do skoroszytu, nie do arkusza |
| Makra VBA | ❌ — należą do skoroszytu |

Wiersze z ✅ wynikają wprost z dokumentacji („tylko komórki … oraz pewne atrybuty arkusza, włącznie z wymiarami, formatem i właściwościami"). Wiersze z ⚠️ to obszar, który zmieniał się między wydaniami openpyxl — **przetestuj w swojej wersji**, zanim na tym oprzesz produkt.

**Dlaczego obrazki i wykresy nie idą z kopią?** Bo w strukturze pliku nie są „treścią arkusza", a **osobnymi obiektami z własnymi relacjami**. Obrazek to plik w `xl/media/` plus wpis w `xl/drawings/drawingN.xml` plus relacja w `xl/worksheets/_rels/sheetN.xml.rels`. Wykres to `xl/charts/chartN.xml` i podobny łańcuch relacji. Skopiowanie takiego łańcucha wymagałoby skopiowania i przepisania relacji — a `copy_worksheet` zajmuje się tylko komórkami i wybranymi ustawieniami arkusza. To ta sama granica, która w module 07 sprawiała, że formuły nie aktualizowały się przy `insert_rows`: openpyxl nie buduje grafu powiązań, więc nie zarządza zależnościami.

**Jak sprawdzić, ile obrazków ma arkusz?** Nie ma do tego publicznego API. W praktyce używa się prywatnego `ws._images`:

```python
print(len(ws._images))       # liczba obrazkow w arkuszu (API prywatne!)
```

To kolejny przykład granicy omówionej w module 03: openpyxl udostępnia to, co wynika z formatu pliku, a nie zawsze to, czego chciałbyś do diagnostyki. Używaj `_images` w skryptach diagnostycznych i testach, ale nie w logice biznesowej — prywatne atrybuty mogą się zmienić.

**Antywzorzec: `copy_worksheet` jako „backup".** Widziałem ten pomysł w kilku projektach: „zanim zacznę modyfikować arkusz, zrobię sobie jego kopię na wszelki wypadek". Trzy powody, dla których to nie działa:

1. Kopia **nie zawiera** obrazków, wykresów, tabel ani formatowania warunkowego — czyli dokładnie tych elementów, które najczęściej chce się odzyskać.
2. Kopia zostaje **w wynikowym pliku**. Twój „backup" jedzie do klienta jako dodatkowa zakładka.
3. To nie jest kopia **skoroszytu**. Skopiowanie arkusza nie zabezpiecza ani nazw zdefiniowanych, ani ustawień skoroszytu, ani makr.

Prawdziwa kopia zapasowa to **kopia pliku na dysku** (`shutil.copy2(źródło, cel)`) wykonana przed otwarciem go w openpyxl. To rozwiązanie omówimy w module 18 jako element wzorca bezpiecznego zapisu.

### 2.9. Usuwanie arkusza i konsekwencje dla odwołań

Dwie równoważne formy usunięcia:

```python
del wb["Dane"]        # po nazwie
wb.remove(wb["Dane"]) # po obiekcie
```

**Co się dzieje w pamięci.** Obiekt arkusza zniknie z listy skoroszytu razem z całym słownikiem komórek. W gałęzi 3.1 openpyxl dokłada starań, żeby posprzątać także **nazwy zdefiniowane lokalne** dla tego arkusza (te z przypisaniem do konkretnego arkusza przez `localSheetId`) — to fragment API, który w 3.1 był przebudowywany, więc szczegóły sprawdź w swojej wersji.

**Czego openpyxl nie posprząta — i to jest cały problem:**

- **Formuł w innych arkuszach, które odwołują się do usuniętego arkusza.**
- **Globalnych nazw zdefiniowanych**, których `attr_text` wskazuje na usunięty arkusz.
- **Wykresów w innych arkuszach** pobierających dane z usuniętego arkusza.
- **Odwołań zewnętrznych** w innych plikach wskazujących na ten arkusz.

Spójrzmy na to w kodzie:

```python
"""Usuniecie arkusza: co zostaje w pliku."""

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

OUTPUT = Path(__file__).resolve().parent.parent / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "08_usuwanie.xlsx"

wb = Workbook()

ws_dane = wb.active
ws_dane.title = "Dane"
ws_dane["A1"] = "Wartość"
ws_dane["A2"] = 10
ws_dane["A3"] = 20

ws_pod = wb.create_sheet("Podsumowanie")
ws_pod["A1"] = "Suma"
ws_pod["B1"] = "=SUM(Dane!A2:A3)"                       # formula do arkusza "Dane"

# Nazwa zdefiniowana (globalna) wskazujaca na arkusz "Dane"
wb.defined_names.add(DefinedName("MojaSuma", attr_text="Dane!$A$2:$A$3"))

print("PRZED USUNIECIEM")
print("  arkusze          :", wb.sheetnames)
print("  formula w B1     :", ws_pod["B1"].value)
print("  nazwa MojaSuma   :", wb.defined_names["MojaSuma"].attr_text)

del wb["Dane"]

print()
print("PO USUNIECIU ARKUSZA 'Dane'")
print("  arkusze          :", wb.sheetnames)
print("  formula w B1     :", ws_pod["B1"].value)       # nadal '=SUM(Dane!A2:A3)'
print("  nazwa MojaSuma   :", wb.defined_names["MojaSuma"].attr_text)
print()
print("  Nic nie zostalo posprzatane. Excel pokaze #REF! przy otwarciu.")

wb.save(CEL)

# Sprawdzamy, co jest w pliku po zapisie i odczycie
wb2 = load_workbook(CEL)
try:
    ws = wb2["Podsumowanie"]
    print()
    print("PO ZAPISIE I ODCZYCIE")
    print("  arkusze          :", wb2.sheetnames)
    print("  formula w B1     :", ws["B1"].value)
    print("  nazwy zdefiniowane:", list(wb2.defined_names))
    print("  nazwa MojaSuma   :", wb2.defined_names["MojaSuma"].attr_text)
finally:
    wb2.close()
```

**Co się dzieje w pamięci.** `del wb["Dane"]` usuwa obiekt arkusza z listy. Nic więcej. Formuła w `Podsumowanie` to zwykły napis `'=SUM(Dane!A2:A3)'` — napisy nie mają pojęcia, do czego się odnoszą, więc pozostają nietknięte. `DefinedName` to osobny obiekt z polem tekstowym `attr_text` — też nie ma pojęcia, że cel zniknął.

**Co trafi do pliku.** Dokładnie to, co widzisz w wydruku powyżej: skoroszyt z jednym arkuszem, formułą wskazującą na nieistniejący arkusz i nazwą zdefiniowaną wskazującą na nieistniejący zakres. Excel pokaże `#REF!` — a jeśli `Dane` było nazwą, którą później utworzysz od nowa, odwołanie „ożyje" i zacznie wskazywać na nowe dane. To drugi, gorszy scenariusz: **cichy błąd, który zniknął sam i przywrócił się w nieprzewidywalnym momencie.**

### 2.9.1. Zmiana nazwy arkusza — ta sama pułapka w tańszym opakowaniu

To najważniejsza rzecz, którą trzeba zapamiętać z tego modułu, a bardzo rzadko ktoś ją podkreśla:

```python
ws_dane.title = "Sprzedaz"      # zmieniamy nazwe arkusza z "Dane" na "Sprzedaz"
```

Po tej operacji:

- `wb.sheetnames` pokaże `['Sprzedaz', 'Podsumowanie']` — wszystko wygląda spójnie,
- formuła w `Podsumowanie` **nadal** brzmi `=SUM(Dane!A2:A3)`,
- nazwa zdefiniowana **nadal** brzmi `Dane!$A$2:$A$3`.

**Zmiana nazwy jest więc operacją tak samo ryzykowną jak usunięcie.** Excel w interfejsie załatwia to automatycznie: gdy zmieniasz nazwę arkusza przez UI, Excel przechodzi przez wszystkie formuły i aktualizuje odwołania. openpyxl tego nie robi — bo tak jak nie ma grafu zależności dla `insert_rows` (moduł 07), tak nie ma go dla nazw arkuszy. Zmienia się jeden napis w `xl/workbook.xml`.

**Dwa sposoby, żeby sobie z tym poradzić:**

1. **Ustal nazwy raz i nie zmieniaj ich.** To brzmi jak unikanie problemu, ale jest najpoważniejszym rozwiązaniem: nazwy arkuszy są częścią kontraktu raportu. Jeśli musisz je zmieniać, to znaczy, że raport zmienia interfejs — i powinien być nową wersją (moduł 26).
2. **Zbuduj skoroszyt od nowa.** Jeśli naprawdę musisz zmienić nazwy, przejdź po wszystkich formułach i nazwach zdefiniowanych i przepisz je programowo — analogicznie jak w module 07 dochodziłeś do `move_range(translate=True)`. To kosztowne, ale wykonalne i w pełni kontrolowalne.

### 2.10. Chartsheet — arkusz, który sam jest wykresem

Excel zna dwa rodzaje zakładek: **arkusz** (Worksheet, siatka komórek) i **arkusz wykresu** (Chartsheet) — zakładkę, która zawiera wyłącznie wykres, bez siatki komórek.

openpyxl obsługuje oba:

```python
cs = wb.create_chartsheet(title="Wykres sprzedaży")
print(wb.chartsheets)          # [<Chartsheet "Wykres sprzedaży">]
```

Różnice względem `Worksheet` są fundamentalne: **Chartsheet nie ma komórek**. Nie ma `cs["A1"]`, nie ma `max_row`, nie ma `iter_rows`. To nie „arkusz z wykresem" — to sam wykres, który zajmuje całą zakładkę.

**Dlaczego w automatyzacji raportów używa się go rzadko:**

1. **Nie da się na nim umieścić niczego poza wykresem.** A raport niemal zawsze potrzebuje tytułu, źródła, daty, przypisu — czyli komórek.
2. **Nie działa w trybach `read_only` i `write_only`** — tak jak wiele innych „złożonych" elementów.
3. **Jest słabiej pokryty przez API** niż zwykłe arkusze i wykresy. To rzadko uczęszczana część openpyxl, więc szczegóły (np. sposób dodawania wykresu do chartsheeta) sprawdź w dokumentacji swojej wersji, zanim na nich oprzesz produkt.
4. **Jest zaskakujący dla użytkownika.** Człowiek szukający danych nie spodziewa się, że zakładka „Wykres" nie ma tabeli.

**Kiedy ma sens:** gdy naprawdę budujesz prezentację „jeden wykres na stronę" i zakładki są slajdami. W praktyce generowania raportów biznesowych **wykresy osadzone w zwykłym arkuszu** (moduł 15) są właściwym wyborem w 95% przypadków: można je obok danych umieścić, opisać komórkami, wydrukować razem z tabelą i skopiować formułą.

**Uczciwa rada na tym etapie kursu:** wiedz, że `Chartsheet` istnieje, ale nie używaj go, dopóki nie zobaczysz modułu 15 i nie będziesz pewien, że naprawdę go potrzebujesz.

### 2.11. Pierwszy wzorzec kursu: jeden arkusz = jedna odpowiedzialność

To jest moment, w którym po raz pierwszy dotykamy architektury — i to nie na siłę, a dlatego, że bez tego wzorca wszystko, co robiłeś w modułach 04–07, rozmnoży się w niekontrolowany sposób.

**Problem w jednym zdaniu:** jeśli jeden arkusz jednocześnie przechowuje dane wejściowe, podsumowanie, wykres, ustawienia i objaśnienie, to **nie da się go kopiować, ukrywać, chronić ani kasować** — bo każda z tych operacji wymaga wyodrębnienia jednej odpowiedzialności z reszty.

**Rozwiązanie:** rozdziel role. Cztery archetypy, które w praktyce pokrywają niemal wszystkie raporty:

| Rola arkusza | Typowa nazwa | `sheet_state` | Kolor | Co w nim jest |
|---|---|---|---|---|
| Dane | `Dane`, `Dane 2026`, `Sprzedaz` | `visible` | niebieski | Czysta tabela: nagłówki, wiersze. Zero komentarzy „na marginesie", zero ozdób |
| Podsumowanie | `Podsumowanie`, `KPI` | `visible` | zielony | Liczby zagregowane, formuły lub wartości policzone w Pythonie |
| Wizualizacja | `Wykresy` | `visible` | złoty | Wykresy osadzone, ewentualnie pasek danych |
| Konfiguracja | `_Konfiguracja`, `_Slowniki` | `veryHidden` | czerwony | Pary klucz–wartość, mapowania, wersja raportu |

**Dlaczego to działa.** Każdy arkusz ma jedną odpowiedzialność, więc każdy można traktować jako osobny byt: `Dane` można podmienić na inne źródło bez ruszania podsumowania; `Podsumowanie` można skopiować do innego pliku (pamiętając o ograniczeniach z sekcji 2.8); `_Konfiguracja` można ukryć i wyczyścić przed wysyłką na zewnątrz (moduł 21). **Wszystko to wynika z jednej decyzji podjętej na początku.**

**Konsekwencje praktyczne, które zobaczysz już teraz:**

- **Nie umieszczaj dwóch tabel na jednym arkuszu.** To znaczy: nie buduj arkusza `Dane`, w którym wiersze 1–50 to sprzedaż, wiersze 55–80 to koszty, a wiersze 85–100 to „notatki". Kiedy przyjdzie nowy wymaganie „dodaj kolumnę do sprzedaży", nie będziesz wiedział, gdzie kończy się sprzedaż.
- **Nie mieszaj ról.** Arkusz `Podsumowanie`, w którym obok KPI stoi tabela pomocnicza z danymi wejściowymi, jest arkuszem, który nie da się ani chronić, ani ukryć, ani skopiować.
- **Nazwy zaczynające się od podkreślenia są sygnałem.** `_Konfiguracja` to konwencja czytelna w tydzień: podkreślenie mówi „to nie jest treść".
- **Kolejność na korytarzu opowiada historię.** `Podsumowanie`, `Dane`, `_Konfiguracja` (ukryte) czyta się samo. `_Konfiguracja`, `Dane`, `Salda 2024 kopia`, `Podsumowanie_final_v3` — nie.

W modułach 22–26 ten prosty wzorzec rozrośnie się w pełną architekturę: role arkuszy staną się klasami w kodzie, ich konfiguracja — strukturą danych, a „jeden arkusz = jedna odpowiedzialność" zamieni się w zasadę SRP stosowaną do warstwy prezentacji. Na razie wystarczy, że będziesz tego pilnował ręcznie w każdym skoroszycie, który budujesz.

## 3. Przykłady krok po kroku

### Przykład 1 — Skoroszyt z rolami arkuszy, konfiguracją i ustawionym widokiem

Zbudujemy raport, który od razu pokazuje wszystkie elementy tego modułu: trzy role, kolor zakładek, ukryty arkusz konfiguracyjny i aktywny arkusz na podsumowaniu.

```python
"""Raport z podzialem na role arkuszy: Dane / Podsumowanie / _Konfiguracja.

Uruchom:  python examples/08_role_arkuszy.py
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "08_role_arkuszy.xlsx"

KOLOR_DANE = "2F5597"
KOLOR_PODSUMOWANIE = "548235"
KOLOR_KONFIGURACJA = "C00000"

DANE_SPRZEDAZY = [
    ("Styczeń", 120_000, 78_000),
    ("Luty", 98_500, 71_200),
    ("Marzec", 143_700, 88_400),
    ("Kwiecień", 131_200, 82_900),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()

    # --- Arkusz 1: DANE ---------------------------------------------------
    # Zamiast usuwac domyslny arkusz i tworzyc nowy - zmieniamy mu nazwe.
    # To nie zmienia pozycji i nie zostawia skoroszytu bez arkusza.
    ws_dane = wb.active
    ws_dane.title = "Dane"

    ws_dane.append(["Miesiąc", "Sprzedaż", "Koszt"])
    for miesiac, sprzedaz, koszt in DANE_SPRZEDAZY:
        ws_dane.append([miesiac, sprzedaz, koszt])

    for komorka in ws_dane["A1:C1"][0]:
        komorka.font = komorka.font.copy(bold=True)

    # --- Arkusz 2: PODSUMOWANIE -------------------------------------------
    ws_pod = wb.create_sheet("Podsumowanie")
    ws_pod["A1"] = "Wskaźnik"
    ws_pod["B1"] = "Wartość"
    ws_pod["A2"] = "Liczba miesięcy"
    ws_pod["B2"] = len(DANE_SPRZEDAZY)
    ws_pod["A3"] = "Sprzedaż razem"
    # Formula odnoszaca sie do arkusza "Dane" - budujemy ja z literalow,
    # ale w module 26 zastapimy to budowaniem programowym.
    ws_pod["B3"] = "=SUM(Dane!B2:B5)"
    ws_pod["A4"] = "Koszt razem"
    ws_pod["B4"] = "=SUM(Dane!C2:C5)"
    ws_pod["A5"] = "Marża"
    ws_pod["B5"] = "=(B3-B4)/B3"
    ws_pod["B5"].number_format = "0.0%"

    # --- Arkusz 3: KONFIGURACJA (ukryty gleboko) --------------------------
    ws_konf = wb.create_sheet("_Konfiguracja")
    ws_konf.append(["Klucz", "Wartość"])
    ws_konf.append(["Wersja raportu", "1.0"])
    ws_konf.append(["Waluta", "PLN"])
    ws_konf.append(["Źródło danych", "ERP / moduł sprzedaży"])
    ws_konf.append(["Wygenerowano", "2026-10-07"])
    ws_konf.sheet_state = "veryHidden"          # <- niewidoczny z interfejsu Excela

    # --- Kolor zakladek (konwencja rol) -----------------------------------
    ws_dane.sheet_properties.tabColor = KOLOR_DANE
    ws_pod.sheet_properties.tabColor = KOLOR_PODSUMOWANIE
    ws_konf.sheet_properties.tabColor = KOLOR_KONFIGURACJA

    # --- Widok: raport otwiera sie na podsumowaniu ------------------------
    ws_pod.sheet_view.zoomScale = 90
    wb.active = wb.index(ws_pod)

    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        print("=" * 74)
        print("CO JEST W PLIKU")
        print("=" * 74)
        print(f"  arkusze (kolejnosc)     : {wb.sheetnames}")
        print(f"  aktywny arkusz          : {wb.active.title!r}")
        print(f"  pozycja aktywnego       : {wb.index(wb.active)}")
        print()
        for ws in wb.worksheets:
            kolor = ws.sheet_properties.tabColor
            kolor_txt = kolor.rgb if kolor is not None else "brak"
            print(f"  {ws.title:<16} state={ws.sheet_state:<11} "
                  f"tabColor={kolor_txt:<10} wymiary={ws.dimensions}")
        print()
        print("  Uwaga: arkusz '_Konfiguracja' jest w pliku i da sie go odczytac,")
        print("  mimo ze w Excelu nie ma go na liscie zakladek. veryHidden to")
        print("  zaslona, nie zabezpieczenie.")
    finally:
        wb.close()


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W momencie `wb.save` skoroszyt trzyma trzy obiekty `Worksheet` na jednej liście, w kolejności `['Dane', 'Podsumowanie', '_Konfiguracja']`, z indeksem aktywnym `1` i jednym stanem `veryHidden`. Wszystkie te informacje są „papierowe" — to atrybuty obiektów, nie zawartość komórek.

**Co trafi do pliku.** Do `xl/workbook.xml` trafi sekwencja trzech elementów `<sheet>` z atrybutami `name` i `state="veryHidden"` na trzecim, oraz jeden element `<workbookView activeTab="1"/>`. Kolory zakładek trafią do `xl/worksheets/sheetN.xml` jako `<tabColor>`. Nic więcej.

**Czego ten przykład uczy najbardziej:** że **struktura skoroszytu to warstwa informacji niezależna od danych**. Możesz mieć identyczne dane w dwóch plikach, a one będą zupełnie różne w odbiorze, bo różni je kolejność, nazwy, kolory, widoczność i to, który arkusz jest aktywny. To warstwa, którą w module 22 nazwiemy „prezentacją" — i którą w raportach biznesowych projektuje się świadomie, nie przypadkiem.

### Przykład 2 — `bezpieczna_nazwa_arkusza` w akcji

Sprawdzimy funkcję z sekcji 2.3 na zestawie wejść, które realnie trafiają do skryptów: nazwy z plików, nazwy od użytkownika, duplikaty.

```python
"""Bezpieczne nazwy arkuszy - test funkcji na problematycznych wejsciach.

Uruchom:  python examples/08_bezpieczne_nazwy.py
"""

from __future__ import annotations

import re

MAKS_DLUGOSC = 31
ZAPASOWA_NAZWA = "Arkusz"
NIEDOZWOLONE = set('[]:*?/\\')


def _zapewnij_unikalnosc(nazwa: str, zajete: set[str], maks: int = MAKS_DLUGOSC) -> str:
    zajete_lower = {n.lower() for n in zajete}
    if nazwa.lower() not in zajete_lower:
        return nazwa
    licznik = 2
    while licznik < 10_000:
        suffix = f"_{licznik}"
        baza = nazwa[: maks - len(suffix)]
        kandydat = f"{baza}{suffix}"
        if kandydat.lower() not in zajete_lower:
            return kandydat
        licznik += 1
    raise RuntimeError(f"nie udalo sie znalezc unikalnej nazwy dla {nazwa!r}")


def bezpieczna_nazwa_arkusza(
    nazwa: object,
    zajete: set[str] | None = None,
    maks: int = MAKS_DLUGOSC,
) -> str:
    tekst = "" if nazwa is None else str(nazwa).strip()
    tekst = "".join("_" if znak in NIEDOZWOLONE else znak for znak in tekst)
    tekst = re.sub(r"[\x00-\x1f]", "_", tekst)
    if len(tekst) > maks:
        tekst = tekst[:maks]
    tekst = tekst.strip("'")
    if not tekst:
        tekst = ZAPASOWA_NAZWA
    if zajete:
        tekst = _zapewnij_unikalnosc(tekst, zajete, maks)
    return tekst


PRZYPADKI: list[tuple[str, object]] = [
    ("nazwa zwyczajna",            "Dane"),
    ("nazwa z data",               "Dane 2026"),
    ("ukosnik w nazwie pliku",     "Raport/Q1"),
    ("dwukropek",                  "Podsumowanie: 2026"),
    ("gwiazdka i nawias",          "Arkusz*[tmp]"),
    ("same biale znaki",           "     "),
    ("napis pusty",                ""),
    ("None",                       None),
    ("liczba zamiast napisu",      2026),
    ("apostrof na brzegach",       "'Q1 2026'"),
    ("za dluga nazwa",             "Bardzo długa nazwa arkusza przekraczająca limit"),
    ("duplikat wielkimi literami", "DANE"),
    ("duplikat dokladny",          "Dane"),
    ("znowu za dluga (duplikat)",  "Bardzo długa nazwa arkusza przekraczająca limit"),
]


def main() -> None:
    zajete: set[str] = set()

    print("=" * 108)
    print(f"{'przypadek':<28} {'wejscie':<50} -> wynik")
    print("=" * 108)

    for opis, surowa in PRZYPADKI:
        wynik = bezpieczna_nazwa_arkusza(surowa, zajete=zajete)
        zajete.add(wynik)
        wejscie_repr = repr(surowa)
        if len(wejscie_repr) > 47:
            wejscie_repr = wejscie_repr[:44] + "..."
        print(f"{opis:<28} {wejscie_repr:<50} -> {wynik!r}")

    print()
    print("=" * 108)
    print("WNIOSKI")
    print("=" * 108)
    print("  1. Znaki : \\ / ? * [ ] zamieniamy na '_' - nie usuwamy, bo to zmienia sens.")
    print("  2. Limit 31 znakow obowiazuje PRZED doklejeniem sufiksu unikalnosci.")
    print("  3. Unikalnosc jest bez wzgledu na wielkosc liter (DANE == dane).")
    print("  4. None, napis pusty i same biale znaki dostaja nazwe zapasowa.")

    print()
    print("Zbior uzytych nazw:")
    for nazwa in sorted(zajete):
        print(f"  {nazwa!r} (dlugosc: {len(nazwa)})")
    print()
    print(f"  Maksymalna dlugosc w zbiorze: {max(len(n) for n in zajete)} "
          f"(limit: {MAKS_DLUGOSC})")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Nic w openpyxl — funkcja operuje wyłącznie na napisach i zbiorze `zajete`. Dlatego można ją przetestować bez tworzenia pliku, co jest bardzo praktyczne (moduł 20).

**Co trafi do pliku.** Nic bezpośrednio. Ale wynik tej funkcji trafi do `ws.title` → do `xl/workbook.xml`. Warto prześledzić wyjątki w kodzie źródłowym openpyxl, żeby zobaczyć, przed czym ta funkcja chroni:

```python
from openpyxl import Workbook

wb = Workbook()

# Przekroczenie 31 znakow - ValueError
try:
    wb.create_sheet("Bardzo długa nazwa arkusza przekraczająca limit")
except ValueError as blad:
    print("limit:", blad)

# Znak niedozwolony - ValueError
try:
    wb.create_sheet("Raport/Q1")
except ValueError as blad:
    print("znak:", blad)

# Duplikat - BRAK wyjatku, cicha zmiana nazwy
print(wb.create_sheet("Dane").title)
print(wb.create_sheet("dane").title)
print(wb.sheetnames)
```

Wydruk z tego fragmentu pokaże dokładnie to, co opisaliśmy w teorii: dwa różne typy błędów na pierwsze dwa przypadki i **ciszę** przy duplikacie. Cisza jest najgroźniejsza — bo nie dowiesz się o niej z logu.

**Czego ten przykład uczy najbardziej:** że walidacja nazw arkusza to **osobny, czysty etap przetwarzania danych wejściowych**, który należy wykonać **przed** dotknięciem openpyxl. To wzorzec, który powtórzy się w module 21 (sanitacja danych) i w module 26 (walidacja na granicy warstwy).

### Przykład 3 — `copy_worksheet`: co przechodzi, a co nie

Ten przykład ma jeden cel: **udowodnić eksperymentalnie**, że kopia arkusza nie jest kopią arkusza.

```python
"""copy_worksheet: co sie kopiuje, a co nie.

Uruchom:  python examples/08_kopiowanie.py
Wymaga Pillow (opcjonalnie) do czesci z obrazkiem: pip install pillow
"""

from __future__ import annotations

import base64
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

CEL = OUTPUT / "08_kopiowanie.xlsx"
OBRAZEK = OUTPUT / "08_logo.png"

# Minimalny, poprawny PNG (1x1 piksel) - zapisujemy go na dysk, zeby nie
# wymagac zadnych zewnetrznych plikow w repozytorium kursu.
PNG_1X1_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def przygotuj_obrazek() -> None:
    if not OBRAZEK.exists():
        OBRAZEK.write_bytes(base64.b64decode(PNG_1X1_B64))
        print(f"  utworzono plik obrazka: {OBRAZEK}")


def wstaw_obrazek(ws) -> bool:
    """Zwraca True, jesli obrazek zostal wstawiony (wymaga Pillow)."""
    try:
        from openpyxl.drawing.image import Image
    except ImportError:
        print("  [pomijam obrazek] brak Pillow - zainstaluj: pip install pillow")
        return False
    try:
        obraz = Image(str(OBRAZEK))
    except ImportError:
        print("  [pomijam obrazek] brak Pillow - zainstaluj: pip install pillow")
        return False
    obraz.width, obraz.height = 64, 64
    ws.add_image(obraz, "E2")
    return True


def zbuduj(cel: Path) -> tuple[Workbook, object]:
    wb = Workbook()
    ws = wb.active
    ws.title = "Zrodlo"

    ws.append(["Pozycja", "Kwota"])
    ws.append(["Widzet", 12.5])
    ws.append(["Wkręt", 340.0])

    # --- elementy, ktore Maja przejsc na kopie ---------------------------
    naglowek = Font(bold=True, color="FFFFFFFF")
    tlo = PatternFill(fill_type="solid", fgColor="FF2F5597")
    for komorka in ws["A1:B1"][0]:
        komorka.font = naglowek
        komorka.fill = tlo

    ws["B2"].number_format = "#,##0.00"
    ws["A2"].comment = Comment("Pozycja referencyjna", "Kurs openpyxl")
    ws["C1"] = "Dokumentacja"
    ws["C1"].hyperlink = "https://openpyxl.readthedocs.io/"

    # --- element, ktory NIE przejdzie ------------------------------------
    ma_obrazek = wstaw_obrazek(ws)

    wb.save(cel)
    return wb, ws, ma_obrazek


def porownaj(zrodlo, kopia) -> None:
    print("=" * 78)
    print("PORÓWNANIE ARKUSZA ZRODLOWEGO I KOPII (stan w pamieci)")
    print("=" * 78)

    def tak_nie(warunek: bool) -> str:
        return "TAK" if warunek else "NIE"

    wiersze = [
        ("wartosc A2",        zrodlo["A2"].value,      kopia["A2"].value),
        ("wartosc B3",        zrodlo["B3"].value,      kopia["B3"].value),
        ("A1 pogrubione",     zrodlo["A1"].font.bold,  kopia["A1"].font.bold),
        ("A1 tlo (rgb)",      zrodlo["A1"].fill.fgColor.rgb,
                              kopia["A1"].fill.fgColor.rgb),
        ("B2 number_format",  zrodlo["B2"].number_format,
                              kopia["B2"].number_format),
        ("A2 komentarz",      zrodlo["A2"].comment is not None,
                              kopia["A2"].comment is not None),
        ("C1 hiperlacz",      getattr(zrodlo["C1"].hyperlink, "target", None),
                              getattr(kopia["C1"].hyperlink, "target", None)),
    ]
    for etykieta, a, b in wiersze:
        zgodne = a == b
        print(f"  {etykieta:<22} zrodlo={a!r:<34} kopia={b!r:<34} "
              f"{'OK' if zgodne else 'ROZNICA'}")

    # Obrazki sa w atrybucie prywatnym _images - nie ma publicznego API.
    obrazy_zrodlo = len(getattr(zrodlo, "_images", []))
    obrazy_kopia = len(getattr(kopia, "_images", []))
    print()
    print(f"  liczba obrazkow (_images): zrodlo={obrazy_zrodlo}  kopia={obrazy_kopia}"
          f"  -> {'KOPIA NIE MA OBRAZKA' if obrazy_kopia < obrazy_zrodlo else 'brak roznicy'}")

    # Tabele - na tym etapie kursu wiemy tylko, ze istnieje slownik ws.tables
    print(f"  liczba tabel (ws.tables) : zrodlo={len(zrodlo.tables)}  "
          f"kopia={len(kopia.tables)}")


def main() -> None:
    przygotuj_obrazek()

    print("=" * 78)
    print("1. BUDUJEMY ARKUSZ ZRODLOWY I KOPIUJEMY GO")
    print("=" * 78)
    wb, ws_zrodlo, ma_obrazek = zbuduj(CEL)
    print(f"  arkusze przed kopiowaniem: {wb.sheetnames}")

    kopia = wb.copy_worksheet(ws_zrodlo)
    print(f"  arkusze po kopiowaniu    : {wb.sheetnames}")
    print(f"  tytul kopii nadany przez openpyxl: {kopia.title!r}")
    print()

    porownaj(ws_zrodlo, kopia)
    print()
    print("  Wniosek: komorki, style, formaty, komentarze i hiperlacze przechodza.")
    print("           Obrazki NIE przechodza - mimo ze sa 'w' arkuszu zrodlowym.")

    print()
    print("=" * 78)
    print("2. KONTROLA PO ZAPISIE I ODCZYCIE")
    print("=" * 78)
    wb.save(CEL)
    wb.close()

    wb2 = load_workbook(CEL)
    try:
        print(f"  arkusze w pliku: {wb2.sheetnames}")
        for nazwa in wb2.sheetnames:
            ws = wb2[nazwa]
            obrazy = len(getattr(ws, "_images", []))
            print(f"    {nazwa!r:<20} obrazy={obrazy}  komentarz A2: "
                  f"{'jest' if ws['A2'].comment else 'brak'}  "
                  f"hiperlacz C1: {'jest' if ws['C1'].hyperlink else 'brak'}")
        if not ma_obrazek:
            print()
            print("  (bez Pillow obrazek nie zostal wstawiony w ogole -")
            print("   zainstaluj pillow i uruchom przyklad ponownie,")
            print("   zeby zobaczyc roznice 1 vs 0 obrazkow)")
    finally:
        wb2.close()

    print()
    print("=" * 78)
    print("3. WARUNKI, W KTORYCH KOPIOWANIE NIE DZIALA")
    print("=" * 78)

    # a) arkusz bez ani jednej komorki
    wb3 = Workbook()
    pusty = wb3.create_sheet("Pusty")
    try:
        wb3.copy_worksheet(pusty)
        print("  a) pusty arkusz: skopiowany bez bledu")
    except Exception as blad:      # noqa: BLE001 - chcemy zobaczyc typ
        print(f"  a) pusty arkusz: {type(blad).__name__}: {blad}")
    wb3.close()

    # b) tryb read_only
    wb4 = load_workbook(CEL, read_only=True)
    try:
        wb4.copy_worksheet(wb4["Zrodlo"])
        print("  b) read_only: skopiowany bez bledu")
    except Exception as blad:      # noqa: BLE001
        print(f"  b) read_only: {type(blad).__name__}: {blad}")
    finally:
        wb4.close()

    print()
    print("  Wniosek: copy_worksheet to operacja wymagajaca pelnego modelu")
    print("           w pamieci i niepustego arkusza.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `copy_worksheet` tworzy nowy obiekt `Worksheet`, po czym iteruje po słowniku komórek źródła i dla każdej komórki tworzy odpowiednik w kopii: przypisuje wartość, typ danych, styl (kopiowany jako obiekt) oraz — jeśli są — hiperłącze i komentarz. **Nic poza tym.** Obrazek w źródle siedzi w `ws_zrodlo._images`, do którego kopiujący kod w ogóle nie zagląda.

**Co trafi do pliku.** Dwa arkusze w `xl/workbook.xml`, dwa pliki `xl/worksheets/sheetN.xml`, ale **tylko jeden** wpis w `xl/media/` i jeden `xl/drawings/drawingN.xml` — należący do arkusza źródłowego. Widać to na poziomie pakietu: kopia ma komórki, ale nie ma własnego rysunku.

**Czego ten przykład uczy najbardziej:** że „skopiuj arkusz" w openpyxl znaczy „skopiuj **komórki i wybrane ustawienia**", a nie „skopiuj **wszystko, co jest w arkuszu**". To ta sama rodzina ograniczeń, którą poznałeś w module 07 przy `insert_rows` (formuły się nie aktualizują) i którą w module 18 zobaczysz jako pełny inwentarz tego, co ginie przy przepuszczaniu pliku przez openpyxl.

### Przykład 4 — Usunięcie i zmiana nazwy arkusza: co zostaje w pliku

```python
"""Usuniecie i zmiana nazwy arkusza - wplyw na formuly i nazwy zdefiniowane.

Uruchom:  python examples/08_usuwanie_i_zmiana_nazwy.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)


def zbuduj() -> Workbook:
    wb = Workbook()

    ws_dane = wb.active
    ws_dane.title = "Dane"
    ws_dane["A1"] = "Wartość"
    ws_dane["A2"] = 10
    ws_dane["A3"] = 20

    ws_pod = wb.create_sheet("Podsumowanie")
    ws_pod["A1"] = "Suma"
    ws_pod["B1"] = "=SUM(Dane!A2:A3)"

    wb.defined_names.add(DefinedName("MojaSuma", attr_text="Dane!$A$2:$A$3"))
    return wb


def raport(wb: Workbook, tytul: str) -> None:
    ws_pod = wb["Podsumowanie"]
    nazwa = wb.defined_names["MojaSuma"].attr_text
    print(f"  {tytul}")
    print(f"    arkusze            : {wb.sheetnames}")
    print(f"    formula w B1       : {ws_pod['B1'].value!r}")
    print(f"    nazwa 'MojaSuma'   : {nazwa!r}")


def main() -> None:
    print("=" * 78)
    print("EKSPERYMENT A: USUNIECIE ARKUSZA, DO KTÓREGO ODNOSZA SIE ODNIESIENIA")
    print("=" * 78)
    wb = zbuduj()
    raport(wb, "przed usunieciem arkusza 'Dane':")

    del wb["Dane"]
    print()
    raport(wb, "po 'del wb[\"Dane\"]':")
    print()
    print("  Nic nie zostalo posprzatane. Excel pokaze #REF!.")

    cel_a = OUTPUT / "08_usuwanie.xlsx"
    wb.save(cel_a)
    wb.close()

    wb_a = load_workbook(cel_a)
    try:
        print()
        print(f"  po zapisie i odczycie pliku {cel_a.name}:")
        print(f"    arkusze            : {wb_a.sheetnames}")
        print(f"    formula w B1       : {wb_a['Podsumowanie']['B1'].value!r}")
        print(f"    nazwa 'MojaSuma'   : {wb_a.defined_names['MojaSuma'].attr_text!r}")
    finally:
        wb_a.close()

    print()
    print("=" * 78)
    print("EKSPERYMENT B: ZMIANA NAZWY ARKUSZA (bez jego usuwania)")
    print("=" * 78)
    wb2 = zbuduj()
    raport(wb2, "przed zmiana nazwy:")

    wb2["Dane"].title = "Sprzedaz"
    print()
    raport(wb2, "po 'ws.title = \"Sprzedaz\"':")
    print()
    print("  Nazwa arkusza zmienila sie w spisie arkuszy,")
    print("  ale formula i nazwa zdefiniowana nadal wskazuja na 'Dane'.")
    print("  Efekt w Excelu: #REF! - mimo ze dane sa w pliku, tylko pod inna nazwa.")

    print()
    print("=" * 78)
    print("EKSPERYMENT C: RENOWACJA - PRZEPISANIE ODNIESIEN PO ZMIANIE NAZWY")
    print("=" * 78)
    wb3 = zbuduj()
    wb3["Dane"].title = "Sprzedaz"

    # Ręczne przepisanie odwolan - jestesmy wlasnym "Excelem".
    ws_pod = wb3["Podsumowanie"]
    stara_formula = ws_pod["B1"].value
    ws_pod["B1"] = stara_formula.replace("Dane!", "Sprzedaz!")

    stara_nazwa = wb3.defined_names["MojaSuma"].attr_text
    nowa_nazwa = stara_nazwa.replace("Dane!", "Sprzedaz!")
    # W 3.1.x DefinedName jest obiektem - modyfikujemy atrybut attr_text.
    wb3.defined_names["MojaSuma"].attr_text = nowa_nazwa

    raport(wb3, "po zmianie nazwy I recznym przepisaniu odwolan:")
    print()
    print("  To dziala, ale jest kosztowne: trzeba przejsc po WSZYSTKICH")
    print("  formulach i nazwach. W module 18 zobaczysz, dlaczego")
    print("  w praktyce lepiej jest budowac skoroszyt od nowa.")

    print()
    print("=" * 78)
    print("EKSPERYMENT D: USUNIECIE ARKUSZA, KTORY JEST AKTYWNY")
    print("=" * 78)
    wb4 = zbuduj()
    wb4.active = 0                       # aktywny = 'Dane'
    print(f"  aktywny przed usunieciem: {wb4.active.title!r}")
    del wb4["Dane"]
    aktywny = wb4.active
    print(f"  aktywny po usunieciu    : "
          f"{aktywny.title if aktywny is not None else None!r}")
    print("  Uwaga: nie zakladaj, ze openpyxl przelaczy aktywnosc sensownie.")
    print("  Zasada: przed usunieciem arkusza ustaw jako aktywny inny, widoczny arkusz.")
    wb4.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W eksperymencie A usuwamy obiekt z listy arkuszy, zostawiając nietknięty obiekt `DefinedName` i nietknięty napis formuły. W eksperymencie B zmieniamy jeden atrybut `_title` na obiekcie arkusza — nic więcej. W eksperymencie C robimy to, co zrobiłby Excel: przechodzimy po odwołaniach i przepisujemy je ręcznie.

**Co trafi do pliku.** W A i B — skoroszyt z odwołaniami do nieistniejącej nazwy. Excel pokaże `#REF!` i (w wielu wersjach) zaproponuje naprawę przy otwarciu. W C — poprawny plik, ale zapłaciliśmy za to ręczną renowacją, która **nie skaluje się**: przy dwudziestu arkuszach i pięćdziesięciu nazwach to już nie jest „przepisanie jednego napisu".

**Czego ten przykład uczy najbardziej:** że **odwołania w openpyxl nie mają właściciela**. Nikt nie pilnuje ich spójności — ani w module 07 przy `insert_rows`, ani tutaj przy usuwaniu i zmianie nazwy arkusza. Konsekwencja architektoniczna, którą zapamiętaj do modułu 26: **jeśli potrzebujesz raportu z nowymi nazwami arkuszy, zbuduj go od nowa, a nie przekształcaj istniejącego.** Budowanie od nowa jest operacją deterministyczną i testowalną; przekształcanie jest operacją zależną od tego, co jest w pliku — a tego nie kontrolujesz.

### Przykład 5 — Kolejność, widoczność i widok: porządek skoroszytu jako decyzja

```python
"""Porzadkowanie skoroszytu: kolejnosc, widocznosc, aktywny arkusz.

Uruchom:  python examples/08_porzadek.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "08_porzadek.xlsx"

KOLEJNOSC_DOCELOWA = ["Podsumowanie", "Dane", "Wykresy", "_Konfiguracja"]

KOLORY = {
    "Dane": "2F5597",
    "Podsumowanie": "548235",
    "Wykresy": "BF8F00",
    "_Konfiguracja": "C00000",
}


def przestaw_arkusze(wb, kolejnosc: list[str]) -> None:
    """Ustawia arkusze w podanej kolejnosci, uzywajac TYLKO API publicznego."""
    nieznane = [n for n in kolejnosc if n not in wb.sheetnames]
    if nieznane:
        raise KeyError(f"arka nie istnieja: {nieznane}")

    for pozycja, nazwa in enumerate(kolejnosc):
        ws = wb[nazwa]
        offset = pozycja - wb.index(ws)
        if offset:
            wb.move_sheet(ws, offset=offset)


def zbuduj_w_chaosie(cel: Path) -> None:
    """Budujemy skoroszyt w 'zlej' kolejnosci - tak, jak wychodzi z kodu napisanego na szybko."""
    wb = Workbook()

    ws_konf = wb.active
    ws_konf.title = "_Konfiguracja"          # <- przypadkiem pierwszy arkusz!
    ws_konf.append(["Klucz", "Wartość"])
    ws_konf.append(["Wersja", "1.0"])
    ws_konf.append(["Waluta", "PLN"])

    ws_dane = wb.create_sheet("Dane")
    ws_dane.append(["Miesiąc", "Sprzedaż"])
    ws_dane.append(["Styczeń", 120_000])
    ws_dane.append(["Luty", 98_500])

    ws_wyk = wb.create_sheet("Wykresy")
    ws_wyk["A1"] = "Miejsce na wykresy (moduł 15)"
    ws_wyk["A1"].font = ws_wyk["A1"].font.copy(italic=True)

    ws_pod = wb.create_sheet("Podsumowanie")
    ws_pod["A1"] = "Sprzedaż razem"
    ws_pod["B1"] = "=SUM(Dane!B2:B3)"

    wb.save(cel)
    wb.close()


def popraw(cel: Path) -> None:
    """Wczytujemy plik, porzadkujemy go i zapisujemy ponownie."""
    wb = load_workbook(cel)
    try:
        print(f"  kolejnosc przed : {wb.sheetnames}")

        # 1. kolejnosc
        przestaw_arkusze(wb, KOLEJNOSC_DOCELOWA)
        print(f"  kolejnosc po    : {wb.sheetnames}")

        # 2. kolory zakladek wg roli
        for nazwa, kolor in KOLORY.items():
            wb[nazwa].sheet_properties.tabColor = kolor

        # 3. widocznosc - konfiguracja ma nie rzucac sie w oczy
        wb["_Konfiguracja"].sheet_state = "veryHidden"

        # 4. widok - raport otwiera sie na podsumowaniu, w pomniejszeniu
        ws_pod = wb["Podsumowanie"]
        ws_pod.sheet_view.zoomScale = 90
        wb.active = wb.index(ws_pod)

        wb.save(cel)
    finally:
        wb.close()


def sprawdz(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        print("  --- stan po porzadkach ---")
        for idx, ws in enumerate(wb.worksheets):
            znacznik_aktywnego = "<-- AKTYWNY" if ws is wb.active else ""
            kolor = ws.sheet_properties.tabColor
            print(f"    [{idx}] {ws.title:<16} state={ws.sheet_state:<11} "
                  f"zoom={ws.sheet_view.zoomScale:<5} "
                  f"tabColor={kolor.rgb if kolor else 'brak':<10} {znacznik_aktywnego}")
    finally:
        wb.close()


def main() -> None:
    zbuduj_w_chaosie(CEL)
    print(f"Zapisano chaos: {CEL}\n")

    print("=" * 78)
    print("PRZED PORZADKAMI (tak wyglada skoroszyt zbudowany 'na szybko')")
    print("=" * 78)
    wb = load_workbook(CEL)
    try:
        print(f"  kolejnosc : {wb.sheetnames}")
        print(f"  aktywny   : {wb.active.title!r}  "
              f"(uzytkownik otwiera plik na pomieszczeniu technicznym!)")
    finally:
        wb.close()

    print()
    print("=" * 78)
    print("PORZADKOWANIE")
    print("=" * 78)
    popraw(CEL)

    print()
    print("=" * 78)
    print("PO PORZADKACH")
    print("=" * 78)
    sprawdz(CEL)

    print()
    print("  Kolejnosc, kolory, widocznosc i aktywny arkusz to jedyne zmiany.")
    print("  Zaden komorek nie dotknelismy - a plik czyta sie zupelnie inaczej.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `przestaw_arkusze` wykonuje sekwencję `move_sheet` ustawiając kolejne arkusze na docelowych pozycjach. Kluczowe: po każdej operacji indeksy się zmieniają, ale algorytm jest odporny, bo liczy offset jako `pozycja_docelowa - pozycja_bieżąca` **w momencie wykonania** — a nie z góry ustaloną listę przesunięć. To ten sam schemat, który w module 07 stosowałeś przy `delete_rows` (idź od dołu, licz pozycje na bieżąco).

**Co trafi do pliku.** Nowa kolejność elementów `<sheet>` w `xl/workbook.xml`; atrybuty `state="veryHidden"` i kolory `<tabColor>` w plikach arkuszy; element `<workbookView activeTab="0"/>`; `zoomScale="90"` w `<sheetView>`. **Żadna komórka nie została zmieniona** — a plik czyta się zupełnie inaczej.

**Czego ten przykład uczy najbardziej:** że **struktura skoroszytu jest osobną warstwą projektową**. Ten sam zestaw danych można podać w skoroszycie, który otwiera się na arkuszu technicznym i wygląda jak bałagan, albo w skoroszycie, który otwiera się na podsumowaniu z kolorowymi zakładkami i czyta się jak raport. Koszt różnicy: jedna funkcja porządkująca na trzydzieści linii kodu.

## 4. Anatomia API

| Metoda / klasa / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | Nowy skoroszyt z jednym arkuszem `Sheet` | — | `wb.active` to ten arkusz |
| `wb.sheetnames` | Lista nazw arkuszy, w kolejności zakładek | — | Główny sposób czytania struktury |
| `wb.worksheets` | Lista obiektów `Worksheet`, w kolejności | — | Bez chartsheetów |
| `wb.chartsheets` | Lista obiektów `Chartsheet` | — | Rzadko używane (sekcja 2.10) |
| `wb["Nazwa"]` | Arkusz po nazwie | `str` | **Dopasowanie dokładne** — `wb["dane"]` dla `Dane` podniesie `KeyError` |
| `wb.worksheets[i]` | Arkusz po pozycji | `int`, 0-based | Bezpieczniejsze niż `wb[i]` przy niejasnej intencji |
| `wb.index(ws)` | Pozycja arkusza na liście | obiekt `Worksheet` | 0-based; podstawa do obliczania offsetów |
| `wb.create_sheet(title=, index=)` | Tworzy arkusz | `title` — nazwa; `index` — pozycja (0-based) | `index` poza zakresem → arkusz trafia na koniec, **bez błędu** |
| `wb.remove(ws)` | Usuwa arkusz | obiekt `Worksheet` | Nie sprząta formuł ani nazw globalnych |
| `del wb["Nazwa"]` | To samo, po nazwie | `str` | Ekwiwalent `wb.remove` |
| `wb.copy_worksheet(ws)` | Kopiuje arkusz w obrębie skoroszytu | obiekt `Worksheet` | Metoda `Workbook`! Nie kopiuje tabel, obrazków, wykresów |
| `wb.move_sheet(ws, offset=)` | Przesuwa arkusz o `offset` pozycji | `offset` — liczba (ujemna = w lewo) | API publiczne; rozszerzenia z nazwą arkusza sprawdź w swojej wersji |
| `wb.active` | Aktywny arkusz (getter/setter) | `int` (indeks); w 3.1.x także `str` | Wartość musi wskazywać **widoczny** arkusz |
| `wb.defined_names` | Kolekcja nazw zdefiniowanych | `.add(DefinedName)`, `[nazwa]`, `in` | Nie jest czyszczona przy usuwaniu arkusza |
| `ws.title` | Nazwa arkusza | `str`, max 31 znaków | Zmiana **nie** aktualizuje odwołań w formułach |
| `ws.sheet_state` | Widoczność | `"visible"`, `"hidden"`, `"veryHidden"` | Co najmniej jeden arkusz musi zostać `visible` |
| `ws.sheet_properties.tabColor` | Kolor zakładki | 6- lub 8-znakowy hex albo `Color` | „Tanie, a poprawia UX" |
| `ws.sheet_properties.codeName` | Identyfikator arkusza dla VBA | `str` | Rzadko potrzebne |
| `ws.sheet_view.zoomScale` | Powiększenie | `int` 10–400 | Ustaw 80–90 dla szerokich raportów |
| `ws.sheet_view.tabSelected` | Zakładka zaznaczona (grupa) | `bool` | Rzadko używane w automatyzacji |
| `ws.sheet_format.defaultColWidth` | Domyślna szerokość kolumn | liczba (znaki) | Dotyczy kolumn bez własnego ustawienia |
| `ws.sheet_format.defaultRowHeight` | Domyślna wysokość wierszy | liczba (punkty) | Jak wyżej |
| `ws._images` | Lista obrazków arkusza | — | **API prywatne** — do diagnostyki i testów, nie do logiki biznesowej |
| `ws.tables` | Słownik tabel arkusza | — | Pełne omówienie w module 12 |
| `wb.create_chartsheet(title=)` | Tworzy arkusz wykresu | `title` | Brak komórek; sprawdź szczegóły w swojej wersji |
| `openpyxl.utils.avoid_duplicate_name` | Dokleja numer do duplikatu nazwy | `(names, value)` | Funkcja pomocnicza openpyxl; nie traktuj jej jako gwarantowanego API |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — skoroszyt z trzema rolami i ustawionym widokiem.**

Napisz `examples/08_cw1.py`, który:

1. Tworzy skoroszyt i **zmienia nazwę arkusza domyślnego** na `Dane` (nie usuwa go i nie tworzy nowego).
2. Wypełnia `Dane` tabelką: nagłówki `Kwartał`, `Sprzedaż`, `Koszt` w wierszu 1, oraz cztery wiersze danych (`Q1`–`Q4`) z dowolnymi liczbami.
3. Tworzy arkusz `Podsumowanie` i wpisuje w nim:
   - wiersz nagłówków `Wskaźnik` / `Wartość`,
   - `Sprzedaż razem` z formułą `=SUM(Dane!B2:B5)`,
   - `Koszt razem` z formułą `=SUM(Dane!C2:C5)`,
   - `Marża` z formułą `=(B2-B3)/B2` i formatem liczby `0.0%` (użyj `ws["B4"].number_format = "0.0%"`).
4. Tworzy arkusz `_Konfiguracja` i wpisuje w nim trzy pary `klucz` / `wartość`: wersję raportu, walutę i źródło danych.
5. Ustawia `_Konfiguracja.sheet_state = "veryHidden"`.
6. Nadaje kolory zakładek z **jednego słownika** `KOLORY_ZAKLADEK` zdefiniowanego na początku pliku (nie rozsypuj literałów po kodzie).
7. Ustawia `zoomScale = 90` na arkuszu `Podsumowanie` i ustawia go jako aktywny — **w tej kolejności**: najpierw widok i aktywność, potem nic więcej, na końcu zapis.
8. Zapisuje plik jako `output/08_cw1.xlsx`, a następnie **wczytuje go ponownie** i wypisuje tabelkę:

   ```
   pozycja | nazwa | state | tabColor | zoom | aktywny
   ```

9. Na końcu wypisuje jedno zdanie weryfikacyjne: czy `wb.active.title == "Podsumowanie"` i czy liczba arkuszy w pliku wynosi `3`.

W pliku `output/08_cw1_wnioski.md` odpowiedz:

- **Dlaczego zmiana nazwy arkusza domyślnego jest lepsza niż `del wb["Sheet"]` + `create_sheet`?** Podaj dwa powody.
- Co by się stało, gdybyś **przed** ustawieniem `wb.active` ukrył `_Konfiguracja`, a potem próbował ustawić ją jako aktywną? Sprawdź to eksperymentalnie w osobnym skrypcie i zapisz komunikat błędu.
- Dlaczego `_Konfiguracja` ma `state="veryHidden"`, a nie `hidden`? Odpowiedz w jednym zdaniu, ale tak, żeby uwzględniało to, co wiesz o odczycie pliku `.xlsx` z modułu 02.

**Zadanie 2 — `bezpieczna_nazwa_arkusza` na realnych danych.**

Napisz `examples/08_cw2.py`, który:

1. Definiuje `bezpieczna_nazwa_arkusza` (możesz skopiować z sekcji 2.3).
2. Przygotowuje **listę dziesięciu** surowych nazw pochodzących z „prawdziwego" źródła. Wymyśl je tak, żeby każda testowała **inną** regułę:

   - nazwa poprawna, krótka,
   - nazwa z ukośnikiem,
   - nazwa z dwukropkiem,
   - nazwa z nawiasem kwadratowym i gwiazdką,
   - nazwa z pytajnikiem,
   - nazwa z ukośnikiem wstecznym,
   - nazwa pusta (napis `""`),
   - nazwa `None`,
   - nazwa mająca 40 znaków,
   - duplikat nazwy pierwszej, ale zapisany wielkimi literami.

3. Przetwarza je **po kolei**, dodając każdą wynikową nazwę do zbioru `zajete` i tworząc w skoroszycie odpowiedni arkusz. Dzięki temu zobaczysz, jak działa zapewnianie unikalności w praktyce.
4. Zapisuje skoroszyt jako `output/08_cw2.xlsx`.
5. Wczytuje go ponownie i wypisuje listę `wb.sheetnames`.
6. Dla każdej nazwy wypisuje jej długość i sprawdza asercjami:
   - długość ≤ 31,
   - brak znaków z zestawu `[]:*?/\\`,
   - unikalność bez względu na wielkość liter.

W pliku `output/08_cw2_wnioski.md` odpowiedz:

- **Które z dziesięciu wejść spowodowałyby wyjątek, gdyby nie przeszły przez Twoją funkcję?** Wypisz je i podaj typ błędu, jaki podnosi openpyxl. **Sprawdź to eksperymentalnie** w osobnym miejscu skryptu (użyj `try/except` i wypisz komunikaty).
- Które wejście **nie** spowodowałoby wyjątku, a mimo to jest niebezpieczne? Dlaczego?
- Dlaczego funkcja zamienia znaki niedozwolone na `_`, a nie usuwa ich? Podaj scenariusz, w którym usunięcie prowadziłoby do dwóch różnych nazw wejściowych dających tę samą nazwę wyjściową (kolizja).

### 🟡 Warsztat

**Zadanie 3 — dowód eksperymentalny: `copy_worksheet` gubi obrazki.**

To ćwiczenie jest bezpośrednim rozwinięciem Przykładu 3, ale z naciskiem na **metodę dowodu**, a nie na wynik.

Napisz `examples/08_cw3.py`, który:

1. Buduje arkusz `Zrodlo`, w którym:
   - są dane w `A1:B4` (nagłówek + trzy wiersze),
   - nagłówek ma pogrubienie i wypełnienie,
   - `B2` ma format `#,##0.00`,
   - `A2` ma komentarz,
   - `C1` ma hiperłącze,
   - `E2` ma obrazek PNG (użyj obrazka zapisanego w `output/08_cw3_logo.png` — możesz użyć tego samego 1×1 PNG zakodowanego w base64, co w Przykładzie 3, i zapisać go na dysk).
2. Skopiuje arkusz: `kopia = wb.copy_worksheet(ws_zrodlo)`.
3. **Zbuduje tabelę porównawczą w konsoli** — dla każdego z sześciu elementów (wartość, styl czcionki, wypełnienie, format liczby, komentarz, hiperłącze, obrazek) wypisze: `element | zrodlo | kopia | zgodne?`. **Zrób to jedną pętlą po liście krotek** `(opis, wartość_ze_zrodla, wartość_w_kopii)`, a nie siedmioma linijkami `print` — w ten sposób ćwiczenie uczy Cię jednocześnie wzorca tabelarycznego wypisu.
4. Zapisuje plik, wczytuje go, i **powtarza porównanie na wczytanych danych** (to kluczowy krok: dowód musi dotyczyć pliku, nie stanu w pamięci).
5. Wypisuje **jednoznaczny werdykt**: `copy_worksheet kopiuje: [lista]`, `copy_worksheet gubi: [lista]`.
6. Na końcu wykonuje **trzy kontrolowane próby niepowodzenia** i wypisuje dla każdej typ wyjątku oraz komunikat:
   - kopiowanie pustego arkusza,
   - kopiowanie w trybie `read_only=True`,
   - kopiowanie w trybie `write_only=True` (utwórz `Workbook(write_only=True)`, dodaj arkusz przez `create_sheet()`, spróbuj skopiować).

W pliku `output/08_cw3_wnioski.md` odpowiedz:

- **Dlaczego obrazek nie został skopiowany** — wyjaśnij to na poziomie **struktury pakietu `.xlsx`**, odwołując się do tego, co wiesz z modułu 02 o `xl/drawings/` i `xl/media/`.
- **Gdyby obrazek jednak został skopiowany**, co musiałoby się stać z plikiem w `xl/media/`? Czy byłby to ten sam plik, czy kopia? Uzasadnij i — jeśli potrafisz — sprawdź eksperymentalnie, rozpakowując wynikowy plik i listując zawartość `xl/media/`.
- **Twoja koleżanka mówi: „zrobię `copy_worksheet` jako backup przed modyfikacją".** Napisz jej odpowiedź w maksymalnie trzech zdaniach, wymieniając **konkretne** elementy, które w tym „backupie" znikną.

### 🔴 Wyzwanie

**Zadanie 4 — `uporzadkuj_skoroszyt(wb, kolejnosc)` z pełną walidacją.**

Napisz `examples/08_cw4.py` z funkcją:

```python
def uporzadkuj_skoroszyt(
    wb,
    kolejnosc: list[str],
    *,
    reszta_na_koniec: bool = True,
    aktywuj: str | None = None,
    ukryj: list[str] | None = None,
) -> dict:
    ...
```

**Wymagania funkcjonalne:**

- **Walidacja wejścia `kolejnosc`:**
  - `kolejnosc` nie może być pusta → `ValueError`,
  - `kolejnosc` nie może zawierać duplikatów (również duplikatów różniących się tylko wielkością liter) → `ValueError` z wypisaniem duplikatu,
  - każda nazwa z `kolejnosc` musi istnieć w `wb.sheetnames` → `KeyError` z listą brakujących.

- **Przestawianie:**
  - arkusze z `kolejnosc` ustawiane są dokładnie w podanej kolejności, od pozycji 0,
  - jeśli `reszta_na_koniec=True`, arkusze nieujęte w `kolejnosc` trafiają na koniec **w kolejności, w jakiej były** w `wb.sheetnames` przed operacją,
  - jeśli `reszta_na_koniec=False`, arkusze nieujęte mają pozostać na swoich pozycjach **względem siebie**, ale nie mogą wejść w obszar zajęty przez `kolejnosc`.

- **Widoczność (parametr `ukryj`):**
  - każda nazwa z `ukryj` zostaje ustawiona na `sheet_state="veryHidden"`,
  - funkcja **musi odmówić** (podnosząc `ValueError`) sytuacji, w której po operacji nie pozostałby **ani jeden** arkusz widoczny,
  - jeśli arkusz, który ma być ukryty, jest aktualnie aktywny, funkcja nie może zostawić skoroszytu z ukrytym aktywnym arkuszem — musi wybrać inny (jawnie) i wypisać w raporcie, że to zrobiła.

- **Aktywny arkusz (parametr `aktywuj`):**
  - jeśli podany, musi istnieć (`KeyError`) i musi być widoczny po zakończeniu operacji (`ValueError`),
  - jeśli `None` — funkcja wybiera pierwszy arkusz z `kolejnosc`, który po operacji pozostaje widoczny.

- **Raport zwrotny** — słownik z polami:
  - `kolejnosc_przed`: lista nazw przed operacją,
  - `kolejnosc_po`: lista nazw po operacji,
  - `aktywny`: nazwa aktywnego arkusza po operacji,
  - `ukryte`: lista faktycznie ukrytych nazw,
  - `ostrzezenia`: lista napisów (np. `"arkusz 'X' byl aktywny i zostal ukryty; nowy aktywny: 'Y'"`).

**Wymagania jakościowe:**

- Funkcja używa **wyłącznie API publicznego** (`wb.sheetnames`, `wb.index`, `wb.move_sheet`, `wb.active`, `ws.sheet_state`). **Zero odwołań do `wb._sheets`.**
- Funkcja **nie modyfikuje** skoroszytu, jeśli walidacja wykryje błąd — czyli: **najpierw cała walidacja, potem wszystkie zmiany.** To ten sam wzorzec, którego użyłeś w module 07 przy usuwaniu wierszy (najpierw zbierz decyzje, potem działaj).

**Sześć testów, każdy na świeżo zbudowanym skoroszycie z arkuszami `['_Konfiguracja', 'Dane', 'Wykresy', 'Podsumowanie', 'Notatki']`:**

1. **Szczęśliwa ścieżka:** `kolejnosc=["Podsumowanie", "Dane", "Wykresy"]`, `ukryj=["_Konfiguracja"]`, `aktywuj="Podsumowanie"`. Oczekiwane: kolejność `["Podsumowanie", "Dane", "Wykresy", "_Konfiguracja", "Notatki"]` (reszta na koniec, w pierwotnej kolejności), `_Konfiguracja` jako `veryHidden`, aktywny `Podsumowanie`.
2. **Duplikat w `kolejnosc`:** `kolejnosc=["Dane", "dane"]` → `ValueError`.
3. **Nieistniejąca nazwa:** `kolejnosc=["Nie ma takiego"]` → `KeyError` z listą brakujących.
4. **Próba ukrycia wszystkiego:** `ukryj=["_Konfiguracja", "Dane", "Wykresy", "Podsumowanie", "Notatki"]` → `ValueError`, i **dodatkowo sprawdzenie, że skoroszyt nie został zmieniony** (porównaj `wb.sheetnames` i stany `sheet_state` przed i po nieudanym wywołaniu).
5. **`aktywuj` na arkuszu, który będzie ukryty:** `aktywuj="_Konfiguracja"`, `ukryj=["_Konfiguracja"]` → `ValueError`.
6. **Ukrywamy aktualnie aktywny arkusz** bez podania `aktywuj`: skoroszyt startuje z aktywnym `_Konfiguracja`, wywołanie `ukryj=["_Konfiguracja"]`. Oczekiwane: sukces, `_Konfiguracja` ukryta, nowy aktywny to pierwszy widoczny arkusz z `kolejnosc`, w `ostrzezeniach` jeden wpis.

Po każdym teście wypisz `test N: OK` albo `test N: BŁĄD - <szczegóły>`, a na końcu `wynik: X/6`.

Zapisz raport do `output/08_cw4_raport.txt`.

W pliku `output/08_cw4_wnioski.md` odpowiedz:

1. **Dlaczego reguła „najpierw cała walidacja, potem wszystkie zmiany" jest w tej funkcji ważniejsza niż wydajność?** Odwołaj się do testu 4.
2. **Dlaczego funkcja nie używa `wb._sheets`?** Podaj dwa powody — jeden praktyczny i jeden dotyczący utrzymania kodu.
3. **Co by się stało, gdyby walidacja i zmiany były przeplatane** (np. najpierw przestawiamy pierwsze dwa arkusze, a potem wykrywamy, że trzeci nie istnieje)? Opisz stan skoroszytu, jaki zastałby użytkownik, i nazwij to jednym słowem (podpowiedź: to słowo pojawi się w module 26).

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""uporzadkuj_skoroszyt: walidacja, przestawianie, widocznosc, aktywny arkusz.

Uruchom:  python examples/08_cw4.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

ARKUSZE_STARTOWE = ["_Konfiguracja", "Dane", "Wykresy", "Podsumowanie", "Notatki"]


# ==================================================================
# WALIDACJA - wszystko przed jakakolwiek zmiana
# ==================================================================
def _waliduj(wb, kolejnosc, ukryj, aktywuj) -> None:
    """Podnosi wyjatek, jesli operacja jest niemozliwa. Nic nie zmienia."""
    if not kolejnosc:
        raise ValueError("parametr 'kolejnosc' nie moze byc pusty")

    # duplikaty - rowniez rozniace sie tylko wielkoscia liter
    widziane: dict[str, str] = {}
    for nazwa in kolejnosc:
        klucz = nazwa.lower()
        if klucz in widziane:
            raise ValueError(
                f"duplikat w 'kolejnosc': {nazwa!r} powtarza {widziane[klucz]!r} "
                f"(nazwy arkuszy sa unikalne bez wzgledu na wielkosc liter)"
            )
        widziane[klucz] = nazwa

    # istnienie arkuszy
    istniejace = set(wb.sheetnames)
    brakujace = [n for n in kolejnosc if n not in istniejace]
    if brakujace:
        raise KeyError(f"arkusze nie istnieja w skoroszycie: {brakujace}")

    if ukryj:
        brakujace_ukryj = [n for n in ukryj if n not in istniejace]
        if brakujace_ukryj:
            raise KeyError(f"arkusze do ukrycia nie istnieja: {brakujace_ukryj}")

        # co zostanie widoczne?
        do_ukrycia = set(ukryj)
        widoczne_po = [n for n in istniejace if n not in do_ukrycia]
        if not widoczne_po:
            raise ValueError(
                "operacja ukrylaby WSZYSTKIE arkusze - "
                "Excel wymaga co najmniej jednego widocznego arkusza"
            )

    if aktywuj is not None:
        if aktywuj not in istniejace:
            raise KeyError(f"arkusz do aktywacji nie istnieje: {aktywuj!r}")
        if ukryj and aktywuj in set(ukryj):
            raise ValueError(
                f"arkusz {aktywuj!r} mialby byc jednoczesnie aktywny i ukryty"
            )


# ==================================================================
# OPERACJE - dopiero po pelnej walidacji
# ==================================================================
def _przestaw(wb, kolejnosc: list[str], reszta_na_koniec: bool) -> None:
    """Ustawia arkusze z 'kolejnosc' na poczatku, w podanej kolejnosci."""
    przed = list(wb.sheetnames)

    if reszta_na_koniec:
        reszta = [n for n in przed if n not in set(kolejnosc)]
    else:
        reszta = [n for n in przed if n not in set(kolejnosc)]

    cel = list(kolejnosc) + reszta

    # Algorytm: dla kazdej pozycji docelowej przesun arkusz na nia.
    # Offset liczymy na biezaco, wiec zmiana indeksow nas nie gubi.
    for pozycja, nazwa in enumerate(cel):
        ws = wb[nazwa]
        offset = pozycja - wb.index(ws)
        if offset:
            wb.move_sheet(ws, offset=offset)


def _ustaw_widocznosc(wb, ukryj: list[str] | None, ostrzezenia: list[str]) -> list[str]:
    """Ustawia wskazane arkusze na veryHidden, dbajac o aktywny i o widocznosc."""
    if not ukryj:
        return []

    do_ukrycia = list(dict.fromkeys(ukryj))       # bez duplikatow, zachowana kolejnosc
    faktycznie_ukryte: list[str] = []

    # Jesli ktorys z ukrywanych jest aktywny - najpierw wybierz inny, widoczny.
    aktywny = wb.active
    if aktywny is not None and aktywny.title in set(do_ukrycia):
        kandydaci = [ws for ws in wb.worksheets
                     if ws.sheet_state == "visible" and ws.title not in set(do_ukrycia)]
        if not kandydaci:
            # walidacja powinna byla to wychwycic - ale nie zakladamy cudow
            raise ValueError("brak widocznego arkusza, na ktory mozna przelaczyc aktywnosc")
        nowy = kandydaci[0]
        wb.active = wb.index(nowy)
        ostrzezenia.append(
            f"arkusz {aktywny.title!r} byl aktywny i zostal ukryty; "
            f"nowy aktywny: {nowy.title!r}"
        )

    for nazwa in do_ukrycia:
        ws = wb[nazwa]
        if ws.sheet_state != "veryHidden":
            ws.sheet_state = "veryHidden"
            faktycznie_ukryte.append(nazwa)

    return faktycznie_ukryte


def _ustaw_aktywny(wb, kolejnosc: list[str], aktywuj: str | None) -> str | None:
    """Ustawia aktywny arkusz. Bez parametru wybiera pierwszy widoczny z 'kolejnosc'."""
    if aktywuj is not None:
        wb.active = wb.index(wb[aktywuj])
        return aktywuj

    for nazwa in kolejnosc:
        ws = wb[nazwa]
        if ws.sheet_state == "visible":
            wb.active = wb.index(ws)
            return nazwa

    # ostatecznosc: pierwszy widoczny w calym skoroszycie
    for ws in wb.worksheets:
        if ws.sheet_state == "visible":
            wb.active = wb.index(ws)
            return ws.title
    return None


# ==================================================================
# GLOWNA FUNKCJA
# ==================================================================
def uporzadkuj_skoroszyt(
    wb,
    kolejnosc: list[str],
    *,
    reszta_na_koniec: bool = True,
    aktywuj: str | None = None,
    ukryj: list[str] | None = None,
) -> dict:
    """Porzadkuje skoroszyt: kolejnosc, widocznosc i aktywny arkusz.

    Wzorzec: NAJPIERW cala walidacja, POTEM wszystkie zmiany.
    Dzieki temu nieudane wywolanie nie zostawia skoroszytu w stanie posrednim.
    """
    kolejnosc_przed = list(wb.sheetnames)
    ostrzezenia: list[str] = []

    # --- 1. WALIDACJA (moze podniesc wyjatek - nic jeszcze nie zmieniamy) ---
    _waliduj(wb, kolejnosc, ukryj, aktywuj)

    # --- 2. ZMIANY ---
    _przestaw(wb, kolejnosc, reszta_na_koniec)
    ukryte = _ustaw_widocznosc(wb, ukryj, ostrzezenia)
    aktywny = _ustaw_aktywny(wb, kolejnosc, aktywuj)

    return {
        "kolejnosc_przed": kolejnosc_przed,
        "kolejnosc_po": list(wb.sheetnames),
        "aktywny": aktywny,
        "ukryte": ukryte,
        "ostrzezenia": ostrzezenia,
    }


# ==================================================================
# TESTY
# ==================================================================
def zbuduj_skoroszyt(aktywny_index: int = 0):
    wb = Workbook()
    wb.active.title = ARKUSZE_STARTOWE[0]
    for nazwa in ARKUSZE_STARTOWE[1:]:
        ws = wb.create_sheet(nazwa)
        ws["A1"] = f"zawartosc arkusza {nazwa}"
    wb.active = aktywny_index
    return wb


def stan(wb) -> dict:
    return {
        "kolejnosc": list(wb.sheetnames),
        "stany": {ws.title: ws.sheet_state for ws in wb.worksheets},
        "aktywny": wb.active.title if wb.active is not None else None,
    }


def test_1_szczesliwa_sciezka():
    wb = zbuduj_skoroszyt()
    raport = uporzadkuj_skoroszyt(
        wb,
        ["Podsumowanie", "Dane", "Wykresy"],
        ukryj=["_Konfiguracja"],
        aktywuj="Podsumowanie",
    )
    oczekiwana_kolejnosc = ["Podsumowanie", "Dane", "Wykresy", "_Konfiguracja", "Notatki"]
    ok = (
        raport["kolejnosc_po"] == oczekiwana_kolejnosc
        and wb["_Konfiguracja"].sheet_state == "veryHidden"
        and raport["aktywny"] == "Podsumowanie"
    )
    return ok, f"kolejnosc_po={raport['kolejnosc_po']}, aktywny={raport['aktywny']!r}"


def test_2_duplikat_w_kolejnosci():
    wb = zbuduj_skoroszyt()
    przed = stan(wb)
    try:
        uporzadkuj_skoroszyt(wb, ["Dane", "dane"])
    except ValueError as blad:
        return stan(wb) == przed, f"ValueError zgodnie z oczekiwaniem: {blad}"
    return False, "oczekiwano ValueError"


def test_3_nieistniejaca_nazwa():
    wb = zbuduj_skoroszyt()
    przed = stan(wb)
    try:
        uporzadkuj_skoroszyt(wb, ["Nie ma takiego"])
    except KeyError as blad:
        return stan(wb) == przed, f"KeyError zgodnie z oczekiwaniem: {blad}"
    return False, "oczekiwano KeyError"


def test_4_ukrycie_wszystkiego():
    wb = zbuduj_skoroszyt()
    przed = stan(wb)
    wszystkie = list(wb.sheetnames)
    try:
        uporzadkuj_skoroszyt(wb, wszystkie, ukryj=wszystkie)
    except ValueError as blad:
        niezmieniony = stan(wb) == przed
        opis = f"ValueError: {blad}; skoroszyt niezmieniony: {niezmieniony}"
        return niezmieniony, opis
    return False, "oczekiwano ValueError"


def test_5_aktywny_i_ukryty_jednoczesnie():
    wb = zbuduj_skoroszyt()
    przed = stan(wb)
    try:
        uporzadkuj_skoroszyt(wb, ["Dane"], aktywuj="_Konfiguracja", ukryj=["_Konfiguracja"])
    except ValueError as blad:
        return stan(wb) == przed, f"ValueError zgodnie z oczekiwaniem: {blad}"
    return False, "oczekiwano ValueError"


def test_6_ukrycie_aktywnego():
    wb = zbuduj_skoroszyt(aktywny_index=0)      # aktywny = '_Konfiguracja'
    raport = uporzadkuj_skoroszyt(
        wb,
        ["Podsumowanie", "Dane", "Wykresy", "Notatki"],
        ukryj=["_Konfiguracja"],
    )
    ok = (
        wb["_Konfiguracja"].sheet_state == "veryHidden"
        and raport["aktywny"] == "Podsumowanie"
        and len(raport["ostrzezenia"]) == 1
    )
    return ok, (f"aktywny={raport['aktywny']!r}, "
                f"ostrzezenia={raport['ostrzezenia']}")


def main() -> int:
    testy = [
        ("1 - szczesliwa sciezka", test_1_szczesliwa_sciezka),
        ("2 - duplikat w kolejnosci", test_2_duplikat_w_kolejnosci),
        ("3 - nieistniejaca nazwa", test_3_nieistniejaca_nazwa),
        ("4 - ukrycie wszystkiego", test_4_ukrycie_wszystkiego),
        ("5 - aktywny i ukryty jednoczesnie", test_5_aktywny_i_ukryty_jednoczesnie),
        ("6 - ukrycie aktywnego", test_6_ukrycie_aktywnego),
    ]

    linie = ["Raport z testow uporzadkuj_skoroszyt", "=" * 72, ""]
    przeszlo = 0

    for etykieta, funkcja in testy:
        try:
            ok, opis = funkcja()
        except Exception as blad:      # noqa: BLE001
            ok, opis = False, f"wyjatek: {type(blad).__name__}: {blad}"
        status = "OK   " if ok else "BLAD "
        if ok:
            przeszlo += 1
        linia = f"test {etykieta:<38} {status} {opis}"
        print(linia)
        linie.append(linia)

    podsumowanie = f"\nwynik: {przeszlo}/{len(testy)}"
    print(podsumowanie)
    linie.append(podsumowanie)

    (OUTPUT / "08_cw4_raport.txt").write_text("\n".join(linie) + "\n", encoding="utf-8")
    return 0 if przeszlo == len(testy) else 1


if __name__ == "__main__":
    raise SystemExit(main())
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **Funkcja `_waliduj` nic nie zmienia — i to jest cała jej wartość.** Wszystkie cztery warunki (niepusta lista, brak duplikatów, istnienie arkuszy, pozostanie widocznego arkusza, spójność `aktywuj` z `ukryj`) sprawdzane są **przed** pierwszym `move_sheet`. Dzięki temu test 4 może porównać stan skoroszytu przed i po nieudanym wywołaniu i dostać równość. Gdyby walidacja była przeplatana ze zmianami, taki test byłby niemożliwy do napisania — i to jest sygnał, że wzorzec jest dobry: **poprawnie zaprojektowana funkcja daje się przetestować tanio.**

- **`_przestaw` liczy offset na bieżąco, nie z góry.** To ten sam wzorzec, który stosowałeś w module 07 przy usuwaniu wierszy (`sorted(..., reverse=True)`) i w Przykładzie 5. Kolejność operacji zmienia indeksy, więc pozycje trzeba liczyć w momencie wykonania. To uniwersalna zasada pracy z każdą listą, którą modyfikujesz w trakcie przetwarzania.

- **`ukryj` przechodzi przez `dict.fromkeys`, żeby zachować kolejność i usunąć duplikaty.** Subtelność, ale ma znaczenie: powtórzone nazwy w `ukryj` nie powinny powodować podwójnych wpisów w raporcie „faktycznie ukryte". `dict.fromkeys(lista)` to prosty, a mało znany idiom „usuń duplikaty, zachowaj kolejność pierwszeństwa".

- **`_ustaw_widocznosc` przestawia aktywność, zanim cokolwiek ukryje.** Ustawienie `sheet_state = "veryHidden"` na aktywnym arkuszu kończy się błędem, więc kolejność jest nieprzypadkowa. Funkcja dodatkowo **zapisuje ostrzeżenie** — a nie tylko robi to po cichu. To wzorzec, który w module 18 nazwiemy „raportem utraty": jeśli operacja musi zrobić coś, czego użytkownik może nie oczekiwać, zapisz to, żeby mógł to zobaczyć.

- **`_ustaw_aktywny` ma trzy poziomy zapasowe.** Pierwszy to jawny parametr (jeśli podany), drugi to pierwszy widoczny arkusz z `kolejnosc`, trzeci to pierwszy widoczny w całym skoroszycie. Ten trzeci poziom teoretycznie nigdy nie powinien się przydać (bo walidacja zapewnia istnienie widocznego arkusza), ale jego obecność oznacza, że funkcja **nie może** zostawić skoroszytu w stanie bez aktywnego arkusza. Zasada: kod obronny w miejscach, których naruszenie kończy się plikiem nie do otwarcia.

- **Zero odwołań do `wb._sheets`.** To nie jest ostrożność estetyczna. Używanie prywatnej listy oznacza, że aktualizujesz strukturę **bez żadnej walidacji** — i jeśli w przyszłej wersji openpyxl lista przestanie być zwykłą listą (np. stanie się obiektem z własną logiką), Twój kod przestanie działać w sposób, którego nie przewidziałeś. API publiczne istnieje po to, żeby ta granica była jasna.

- **Raport zwrotny jako słownik, a nie lista napisów.** Dzięki temu funkcja jest użyteczna nie tylko dla człowieka (wypis), ale i dla kodu, który ją wywoła (np. log audytowy w module 26 albo test asercji). To wzorzec DTO, do którego wrócimy w module 23.

**Czego ten szkic jeszcze nie robi, a warto dopisać w produkcji:** obsługi arkuszy wykresów (`wb.chartsheets`) w parametrze `kolejnosc` — na razie funkcja działa tylko na `Worksheet`. Gdybyś chciał objąć też chartsheety, wystarczy zamienić `wb.worksheets` na `wb.worksheets + wb.chartsheets` i pamiętać, że `Chartsheet` nie ma `sheet_state` w tej samej formie (sprawdź w dokumentacji swojej wersji). Dodatkowo: parametr `kolorowe_zakladki: dict[str, str]`, który nada kolory w tym samym przebiegu — to naturalne rozszerzenie, które w module 23 nazwiemy już „budowniczym konfiguracji".

</details>

## 6. Typowe błędy i pułapki

**1. „Utworzyłem dwa arkusze o tej samej nazwie i openpyxl nie zgłosił błędu" (objaw) → openpyxl nie podnosi wyjątku przy duplikacie nazwy — automatycznie zmienia nazwę na unikalną, dopisując sufiks liczbowy, korzystając z pomocniczej funkcji `avoid_duplicate_name`. `wb.create_sheet("Dane")` przy istniejącym `Dane` da arkusz `Dane1` (przyczyna) → **jeśli nazwy arkuszy mają dla Ciebie znaczenie, buduj je sam przez `bezpieczna_nazwa_arkusza` i sprawdzaj `wb.sheetnames` po utworzeniu. Nigdy nie zakładaj, że `ws.title` równa się temu, co podałeś w `create_sheet` (naprawa).**

To najniebezpieczniejszy błąd w tym module, bo jest cichy. Nie zobaczysz wyjątku, nie zobaczysz ostrzeżenia (a jeśli je zobaczysz, to zależy od wersji) — po prostu po dwóch tygodniach okaże się, że formuła `=SUM(Dane!B2:B13)` wskazuje na inny arkusz niż zamierzone.

**2. „`wb["dane"]` podnosi `KeyError`, a arkusz `Dane` przecież istnieje" (objaw) → `wb["Nazwa"]` dopasowuje nazwy **dokładnie**, z uwzględnieniem wielkości liter. Excel w interfejsie jest w tym miejscu niewrażliwy — i stąd nieporozumienie (przyczyna) → albo używaj dokładnych nazw (najlepiej: pobranych z `wb.sheetnames`), albo napisz własną funkcję wyszukiwania:

```python
def arkusz_po_nazwie(wb, nazwa: str):
    """Wyszukiwanie nazwy arkusza bez wzgledu na wielkosc liter."""
    klucz = nazwa.strip().lower()
    for ws in wb.worksheets:
        if ws.title.lower() == klucz:
            return ws
    raise KeyError(f"nie ma arkusza o nazwie {nazwa!r}; dostepne: {wb.sheetnames}")
```

Jeśli używasz własnej funkcji, **pamiętaj, że wtedy Twój kod jest łagodniejszy niż Excel** — a to oznacza, że formuły, które napiszesz z tej nazwy, mogą nie zadziałać. Lepiej wymuszać dokładne nazwy (naprawa).

**3. „Ukryłem wszystkie arkusze i teraz zapis wywala błąd" (objaw) → Excel wymaga co najmniej jednego widocznego arkusza w skoroszycie, a openpyxl tego pilnuje — ustawienie `wb.active` na ukryty arkusz albo zapis skoroszytu bez widocznego arkusza kończy się wyjątkiem; dokładne miejsce i komunikat sprawdź w swojej wersji (przyczyna) → **ustal kolejność: najpierw aktywny, potem ukrywanie.** I sprawdzaj przed ukryciem, czy zostanie coś widocznego — dokładnie tak, jak w funkcji `ukryj_arkusz` z sekcji 2.5 (naprawa).**

**4. „Ustawiłem `wb.active` na arkusz, a potem go ukryłem i teraz mam błąd" (objaw) → `wb.active` musi wskazywać **widoczny** arkusz, a operacja ukrycia nie przełącza aktywności automatycznie (przyczyna) → kolejność operacji ma znaczenie. Najpierw wskaż nowy aktywny arkusz (inny niż ten, który ukrywasz), potem ukryj. Jeśli nie masz innego arkusza — najpierw go utwórz (naprawa).**

**5. „`copy_worksheet` nie skopiował obrazków / wykresów / tabel" (objaw) → dokumentacja openpyxl mówi to wprost: kopiowane są **tylko komórki (wartości, style, hiperłącza, komentarze) oraz pewne atrybuty arkusza (wymiary, format, właściwości)**; wszystko inne, w tym obrazy i wykresy, nie jest kopiowane. Tabele również nie — dokumentacja wspomina, że openpyxl nie zarządza zależnościami takimi jak formuły, tabele czy wykresy przy kopiowaniu (przyczyna) → **jeśli chcesz skopiować arkusz „w całości", skopiuj plik na dysku** (`shutil.copy2`) i modyfikuj kopię. Jeśli chcesz skopiować pojedynczy obrazek — dodaj go programowo do kopii przez `add_image` (moduł 16). Dla wykresów — dodaj wykres programowo (moduł 15). Dla tabel — dodaj tabelę (moduł 12). To nie obejście, tylko poprawny model pracy (naprawa).**

**6. „`wb.copy_worksheet` wywala `IndexError`, mimo że arkusz istnieje" (objaw) → openpyxl traktuje arkusz bez ani jednej komórki jako „pusty" i odmawia skopiowania; typ wyjątku i komunikat sprawdź w swojej wersji, ale zachowanie jest spójne — pustego arkusza się nie kopiuje (przyczyna) → upewnij się, że w arkuszu jest co najmniej jedna komórka (np. nagłówek), **zanim** go skopiujesz. Jeśli naprawdę potrzebujesz pustej kopii, użyj `wb.create_sheet` i przenieś do niej ustawienia ręcznie (naprawa).**

**7. „Próbuję kopiować arkusz między dwoma skoroszytami i nie działa" (objaw) → `copy_worksheet` działa **tylko w obrębie jednego `Workbook`**; kopiowanie między skoroszytami nie jest obsługiwane (przyczyzna) → masz dwa realne wyjścia: (a) wczytaj oba pliki, przekopiuj dane komórka po komórce w pętli po `iter_rows`, a style przenieś ręcznie — albo (b) potraktuj skoroszyt docelowy jako **szablon** i generuj go od nowa na podstawie danych ze źródła, zamiast kopiować strukturę. Rozwiązanie (b) jest zwykle lepsze architektonicznie — i to jest dokładnie wzorzec „Prototype" z modułu 23 (naprawa).**

**8. „Nazwa arkusza `Raport / Q1` lub `Podsumowanie: 2026` wywala skrypt" (objaw) → znaki `: \ / ? * [ ]` są niedozwolone w nazwach arkuszy — openpyxl podnosi `ValueError` wskazując znak. Przekroczenie 31 znaków podnosi `ValueError` z komunikatem o limicie (przyczyna) → **nigdy nie przekazuj surowej nazwy do `create_sheet` ani `ws.title`.** Przepuść ją przez `bezpieczna_nazwa_arkusza` z sekcji 2.3. Dotyczy to zwłaszcza nazw pochodzących z plików, konfiguracji i od użytkownika (naprawa).**

**9. „Usunąłem arkusz (albo zmieniłem jego nazwę) i teraz Excel pokazuje `#REF!` w innym arkuszu" (objaw) → openpyxl nie śledzi odwołań między arkuszami: ani formuł, ani nazw zdefiniowanych, ani wykresów. Usunięcie albo zmiana nazwy arkusza nie aktualizuje niczego poza nazwami lokalnymi arkusza, które `remove` stara się posprzątać — a i to sprawdź w swojej wersji (przyczyna) → **nigdy nie usuwaj i nie zmieniaj nazwy arkusza w skoroszycie, w którym są formuły albo nazwy zdefiniowane, chyba że masz plan ich renowacji.** Wzorce, o których mówiliśmy w module 07, powtarzają się tu w całości: albo ręczna renowacja (przejdź po formułach i nazwach i przepisz je), albo — bezpieczniej — zbuduj skoroszyt od nowa z nowymi nazwami. Zobacz Przykład 4 (naprawa).**

**10. „Przestawiam arkusze przez `wb._sheets = [...]`, bo tak było w internecie" (objaw) → `wb._sheets` to implementacja prywatna, a nie część publicznego API. Wykorzystywanie jej oznacza, że opierasz się na szczególe, którego autorzy biblioteki nie zobowiązali się utrzymać. Przestawiając listę ręcznie pomijasz też wszelkie sprawdzenia, które `move_sheet` wykonuje (przyczyna) → używaj `wb.move_sheet(ws, offset=...)`, `wb.index(ws)` i `wb.sheetnames`. Napisz jedną funkcję przestawiającą (jak `przestaw_arkusze` z Przykładu 5) i używaj jej wszędzie (naprawa).**

**11. „`wb.create_sheet("Podsumowanie", 99)` — chciałem dodać arkusz na końcu, ale myślałem, że to błąd" (objaw) → `index` poza zakresem listy nie podnosi wyjątku: `list.insert` wstawia element na koniec, gdy indeks jest za duży. Tak więc `index=99` w skoroszycie z trzema arkuszami działa jak brak `index` (przyczyna) → **jeśli chcesz umieścić arkusz na końcu, po prostu pomiń `index`.** Jeśli podajesz `index`, podawaj wartość, którą świadomie wybrałeś z zakresu `0..len(wb.sheetnames)`. A jeśli kod ma być jawny, dopisz walidację:

```python
def utworz_na_pozycji(wb, nazwa: str, pozycja: int):
    if not 0 <= pozycja <= len(wb.sheetnames):
        raise ValueError(
            f"pozycja {pozycja} poza zakresem 0..{len(wb.sheetnames)}"
        )
    return wb.create_sheet(bezpieczna_nazwa_arkusza(nazwa, set(wb.sheetnames)), pozycja)
```

(naprawa).

**12. „Skopiowałem arkusz i teraz mam dwa razy te same dane w raporcie" (objaw) → `copy_worksheet` nie jest operacją tymczasową — kopia zostaje w skoroszycie **na stałe** i trafia do pliku wynikowego. Nie ma mechanizmu „ukrytej kopii roboczej" (przyczyna) → **kopie robocze rób poza skoroszytem.** Jeśli potrzebujesz „wersji roboczej" do porównania, zapisz dwa pliki i porównaj je programowo (moduł 20 pokaże Ci gotowe narzędzia). Jeśli kopia ma zostać w pliku, nadaj jej **jawną, znaczącą nazwę** i włącz ją do modelu raportu jako osobną rolę arkusza (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Skoroszyt to uporządkowana lista arkuszy — i ta lista jest osobną warstwą informacji.** Kolejność, nazwy, kolory zakładek, widoczność i aktywny arkusz nie zmieniają ani jednej wartości w komórkach, a całkowicie zmieniają sposób, w jaki plik jest czytany. Traktuj strukturę skoroszytu jako świadomą decyzję projektową podejmowaną raz, a nie jako efekt uboczny kolejności, w jakiej napisałeś kod. W praktyce: buduj arkusze w jednym miejscu, przestawiaj jedną funkcją, ustawiaj widok na końcu.

2. **Nazwa arkusza jest danymi wejściowymi, a nie literałem.** Ma twarde reguły (31 znaków, znaki `: \ / ? * [ ]`, unikalność bez względu na wielkość liter), spotyka się z danymi, które przychodzą z plików, konfiguracji i od użytkowników, i **openpyxl nie podniesie błędu przy duplikacie** — zmieni nazwę po cichu. Konsekwencja: każdą nazwę pochodzącą z zewnątrz przepuszczaj przez jedną funkcję walidującą (`bezpieczna_nazwa_arkusza`), pisaną raz i używaną wszędzie.

3. **`visible` / `hidden` / `veryHidden` to trzy różne narzędzia, a nie trzy stopnie tej samej rzeczy.** `hidden` ukrywa przed przypadkowym spojrzeniem (można odsłonić). `veryHidden` ukrywa przed interfejsem (wymaga narzędzia) — ale **nie przed kimkolwiek, kto potrafi rozpakować ZIP**. Arkusz `_Konfiguracja` z `veryHidden` to dobre miejsce na wersję raportu i słowniki, ale nie na hasła, klucze API ani dane osobowe. Zawsze przy tym pilnuj dwóch twardych zasad Excela: co najmniej jeden arkusz widoczny i aktywny arkusz widoczny.

4. **`copy_worksheet` kopiuje komórki i wybrane ustawienia — nie arkusz.** Przechodzą wartości, style, formaty liczb, hiperłącza i komentarze. **Nie przechodzą** tabele, obrazki, wykresy, formatowanie warunkowe, walidacja ani nazwy zdefiniowane. Kopia nie działa między skoroszytami, nie działa w `read_only`/`write_only` i nie działa dla pustego arkusza. Dlatego `copy_worksheet` **nie jest backupem** — backupem jest skopiowanie pliku na dysku. A jeśli naprawdę chcesz „skopiować wszystko", to w openpyxl znaczy „zbudować od nowa".

5. **Usunięcie i zmiana nazwy arkusza to operacje na odwołaniach — nie tylko na strukturze.** openpyxl nie ma grafu zależności: formuła `=SUM(Dane!B2:B13)`, nazwa zdefiniowana `Dane!$A$2:$A$100` i wykres pobierający dane z arkusza `Dane` **pozostaną nietknięte**, nawet jeśli `Dane` przestanie istnieć albo zmieni nazwę. Excel pokaże `#REF!` (dobry scenariusz) albo — gorzej — odwołanie trafi na inne dane i pokaże wiarygodną, ale fałszywą liczbę. Wniosek architektoniczny, który wróci w module 26: **jeśli raport ma nową strukturę, zbuduj go od nowa; nie przekształcaj istniejącego.**

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Color, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# ==================================================================
# 2. CZYTANIE STRUKTURY SKOROSZYTU
# ==================================================================
wb.sheetnames                      # ['Dane', 'Podsumowanie', '_Konfiguracja'] - kolejnosc!
wb.worksheets                      # lista obiektow Worksheet, w kolejnosci
wb.chartsheets                     # lista obiektow Chartsheet (rzadkie)
wb.index(ws)                       # pozycja arkusza (0-based)
ws = wb["Dane"]                    # DOKLADNE dopasowanie nazwy (wielkosc liter ma znaczenie!)
ws = wb.worksheets[0]              # pierwszy arkusz po pozycji

# ==================================================================
# 3. TWORZENIE I USUWANIE
# ==================================================================
wb = Workbook()                    # nowy skoroszyt ma JEDEN arkusz o nazwie 'Sheet'
ws = wb.active
ws.title = "Dane"                  # <- lepsze niz del + create_sheet

ws = wb.create_sheet("Podsumowanie")            # na koniec
ws = wb.create_sheet("Podsumowanie", 0)         # na poczatek
# index poza zakresem -> arkusz trafia na koniec, BEZ bledu

del wb["Dane"]                     # usuniecie po nazwie
wb.remove(wb["Dane"])              # to samo po obiekcie

# ==================================================================
# 4. KOLEJNOSC - TYLKO API PUBLICZNE
# ==================================================================
wb.move_sheet(ws, offset=-1)       # przesun o jedna pozycje w lewo
wb.move_sheet(ws, offset=2)        # przesun o dwie pozycje w prawo

def przestaw_arkusze(wb, kolejnosc: list[str]) -> None:
    """Ustawia arkusze dokladnie w podanej kolejnosci. Wylacznie API publiczne."""
    nieznane = [n for n in kolejnosc if n not in wb.sheetnames]
    if nieznane:
        raise KeyError(f"arkusze nie istnieja: {nieznane}")
    for pozycja, nazwa in enumerate(kolejnosc):
        ws = wb[nazwa]
        offset = pozycja - wb.index(ws)      # offset liczony NA BIEZACO
        if offset:
            wb.move_sheet(ws, offset=offset)

# NIE ROB TAK: wb._sheets = [...] - to implementacja prywatna

# ==================================================================
# 5. NAZWY ARKUSZY - WALIDACJA
# ==================================================================
# Reguly: max 31 znakow, zakazane : \ / ? * [ ], unikalne bez wzgledu
# na wielkosc liter, nie moga byc puste.
# Duplikat NIE podnosi bledu - openpyxl cicho dokleja numer.

NIEDOZWOLONE = set('[]:*?/\\')
MAKS_DLUGOSC = 31

def _zapewnij_unikalnosc(nazwa: str, zajete: set[str], maks: int = MAKS_DLUGOSC) -> str:
    zajete_lower = {n.lower() for n in zajete}
    if nazwa.lower() not in zajete_lower:
        return nazwa
    licznik = 2
    while licznik < 10_000:
        suffix = f"_{licznik}"
        kandydat = f"{nazwa[: maks - len(suffix)]}{suffix}"
        if kandydat.lower() not in zajete_lower:
            return kandydat
        licznik += 1
    raise RuntimeError(f"brak unikalnej nazwy dla {nazwa!r}")

def bezpieczna_nazwa_arkusza(nazwa, zajete=None, maks: int = MAKS_DLUGOSC) -> str:
    import re
    tekst = "" if nazwa is None else str(nazwa).strip()
    tekst = "".join("_" if znak in NIEDOZWOLONE else znak for znak in tekst)
    tekst = re.sub(r"[\x00-\x1f]", "_", tekst)
    tekst = tekst[:maks].strip("'") or "Arkusz"
    return _zapewnij_unikalnosc(tekst, zajete, maks) if zajete else tekst

# Uzycie:
zajete: set[str] = set()
for surowa in ["Dane", "Raport/Q1", "dane", "Podsumowanie: 2026"]:
    nazwa = bezpieczna_nazwa_arkusza(surowa, zajete=zajete)
    zajete.add(nazwa)
    print(f"{surowa!r} -> {nazwa!r}")

# ==================================================================
# 6. WIDOCZNOSC
# ==================================================================
ws.sheet_state = "visible"         # domyslnie - widoczna zakladka
ws.sheet_state = "hidden"          # ukryta, ale odkrywalna przez uzytkownika Excela
ws.sheet_state = "veryHidden"      # ukryta takze przed interfejsem Excela (NIE sejf!)

# Twarde zasady Excela:
#   - co najmniej jeden arkusz musi pozostac 'visible'
#   - aktywny arkusz musi byc widoczny
# Kolejnosc operacji: najpierw nowy aktywny, POTEM ukrycie starego.

def ukryj_arkusz(wb, nazwa: str, stan: str = "veryHidden") -> None:
    if stan not in {"hidden", "veryHidden"}:
        raise ValueError(f"nieznany stan: {stan!r}")
    kandydat = wb[nazwa]
    widoczne = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]
    if len(widoczne) == 1 and widoczne[0] is kandydat:
        raise ValueError("nie mozna ukryc ostatniego widocznego arkusza")
    if wb.active is kandydat:
        zapas = next(ws for ws in widoczne if ws is not kandydat)
        wb.active = wb.index(zapas)        # NAJPIERW aktywny
    kandydat.sheet_state = stan            # POTEM ukrycie

# ==================================================================
# 7. KOLOR ZAKLADKI, WIDOK, AKTYWNY ARKUSZ
# ==================================================================
ws.sheet_properties.tabColor = "2F5597"                 # 6-znakowy hex
ws.sheet_properties.tabColor = Color(rgb="002F5597")    # albo obiekt Color

ws.sheet_view.zoomScale = 90        # 10..400 - ustaw 80-90 dla szerokich raportow
ws.sheet_view.tabSelected = True    # zaznaczona zakladka (grupa) - rzadko potrzebne

wb.active = wb.index(ws)            # aktywny arkusz po indeksie (bezpieczna forma)
# w 3.1.x spotkasz tez wariant z nazwa arkusza - sprawdz swoja wersje

# ==================================================================
# 8. KOPIOWANIE ARKUSZA - CO PRZECHODZI, A CO NIE
# ==================================================================
kopia = wb.copy_worksheet(wb["Dane"])      # METODA SKOROSZYTU, nie arkusza!
kopia.title = bezpieczna_nazwa_arkusza("Dane - archiwum", set(wb.sheetnames))

# PRZECHODZI: wartosci, typy danych, style, number_format, hiperlacza,
#             komentarze, wymiary kolumn/wierszy, sheet_format, sheet_properties
# NIE PRZECHODZI: tabele, obrazy, wykresy, formatowanie warunkowe,
#                 walidacja danych, nazwy zdefiniowane, makra VBA
# NIE DZIALA: miedzy skoroszytami, w read_only, w write_only, dla pustego arkusza

print(len(ws._images))   # liczba obrazkow - API PRYWATNE, tylko do diagnostyki

# PRAWDZIWY BACKUP = kopia pliku na dysku:
# import shutil; shutil.copy2(zrodlo, f"{zrodlo}.bak")

# ==================================================================
# 9. USUWANIE I ZMIANA NAZWY - KONSEKWENCJE
# ==================================================================
# WZORZEC: najpierw ustaw nowy aktywny arkusz, potem usun stary.
nowy_aktywny = next(ws for ws in wb.worksheets if ws.title != "DoUsuniecia")
wb.active = wb.index(nowy_aktywny)
del wb["DoUsuniecia"]

# CO ZOSTAJE NIETKNIETE po usunieciu / zmianie nazwy arkusza:
#   - formuly w innych arkuszach wskazujace na ten arkusz
#   - globalne nazwy zdefiniowane (wb.defined_names) wskazujace na ten arkusz
#   - wykresy i odsylacze zewnetrzne
# Excel pokaze #REF! albo - gorzej - wskaze inne dane.

# Nazwy zdefiniowane - API 3.1.x:
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names.add(DefinedName("MojaSuma", attr_text="Dane!$A$2:$A$3"))
print(wb.defined_names["MojaSuma"].attr_text)
print("MojaSuma" in wb.defined_names)
print(list(wb.defined_names))

# ==================================================================
# 10. SKOROSZYT W ROLACH - WZORZEC STARTOWY PROJEKTU
# ==================================================================
KOLORY_ZAKLADEK = {
    "dane":         "2F5597",   # niebieski - zrodlo
    "podsumowanie": "548235",   # zielony   - wynik
    "wykresy":      "BF8F00",   # zloty     - wizualizacja
    "metodyka":     "7F7F7F",   # szary     - opis
    "konfiguracja": "C00000",   # czerwony  - nie dotykac
}

ROLE_ARKUSZY = {
    "Dane":           ("dane",         "visible"),
    "Podsumowanie":   ("podsumowanie", "visible"),
    "Wykresy":        ("wykresy",      "visible"),
    "_Konfiguracja":  ("konfiguracja", "veryHidden"),
}

# ==================================================================
# 11. CHECKLISTA PRZED ZAPISEM SKOROSZYTU
# ==================================================================
# [ ] Czy arkusze sa w docelowej kolejnosci (wb.sheetnames)?
# [ ] Czy nazwy sa poprawne i unikalne (max 31 znakow, bez : \ / ? * [ ])?
# [ ] Czy co najmniej jeden arkusz pozostaje 'visible'?
# [ ] Czy wb.active wskazuje WIDOCZNY arkusz, na ktorym ma sie otworzyc raport?
# [ ] Czy kolory zakladek sa ustawione wg jednej konwencji z jednego slownika?
# [ ] Czy arkusze techniczne sa ukryte (i czy wiem, ze to nie zabezpieczenie)?
# [ ] Czy nie usuwalem ani nie zmienialem nazwy arkusza z formulami/nazwami?
# [ ] Czy otworzylem wynikowy plik w Excelu i sprawdzilem #REF! ?
```

## 9. Co dalej

Nauczyłeś się zarządzać **strukturą** skoroszytu: jego listą arkuszy, ich nazwami, kolejnością, widocznością i rolą. To trzecia warstwa tego kursu — po typach danych (moduły 04–06) i geometrii arkusza (moduł 07). Dołożyłeś do modelu mentalnego trzy rzeczy, które będą procentować przez całą resztę kursu:

- **„Struktura skoroszytu to decyzja projektowa".** W module 22 zobaczysz, jak ta sama myśl rozwinie się w podział na trzy światy (dane / prezentacja / zapis). W module 26 role arkuszy staną się klasami w kodzie, a `KOLORY_ZAKLADEK` — elementem konfiguracji raportu. Już teraz masz zalążek tego wzorca na kilku linijkach.
- **„Nazwy i odwołania nie mają właściciela".** To ta sama rodzina problemów, którą poznałeś w module 07 przy `insert_rows`. W module 12 zobaczysz ją w wersji „tabela z zakresem, który się nie zaktualizuje". W module 15 w wersji „wykres pobierający dane z arkusza po refaktorze". A w module 18 — jako pełny inwentarz tego, co trzeba sprawdzić po przepuszczeniu pliku przez openpyxl.
- **„`veryHidden` to zasłona, nie sejf".** W module 17 zobaczysz to samo przy ochronie hasłem. W module 21 staniemy się ostrożni wobec wszystkiego, co przychodzi z zewnątrz — włącznie z ukrytymi arkuszami, które ktoś zostawił w pliku od klienta. To jest ta sama rozmowa o bezpieczeństwie, prowadzona trzy razy w coraz ostrzejszym tonie.

W **module 09** zajmiemy się **stylem komórki** — i to jest moduł, na który możesz się przygotować, robiąc jedno ćwiczenie myślowe. W module 07 zobaczyłeś, że `ws.max_row` rośnie od formatowania, a w tym module, że `copy_worksheet` kopiuje style. Ale **czym właściwie jest styl w openpyxl?** Odpowiedź jest zaskakująca: to nie zbiór „ustawień na komórce", a **niemutowalny, współdzielony obiekt**, do którego komórka ma tylko **odwołanie**.

Zanim przejdziesz dalej, spróbuj — **i to jest Twoje zadanie przygotowawcze do modułu 09** — zgadnąć, dlaczego poniższy kod **nie pogrubi tekstu**:

```python
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
ws["A1"] = "Nagłówek"

ws["A1"].font.bold = True      # <- dlaczego to nic nie robi?

print(ws["A1"].font.bold)      # <- co tu zostanie wypisane?
```

Zapisz swoją hipotezę, a potem — zanim przeczytasz moduł 09 — **uruchom ten kod i sprawdź**, co naprawdę się stało i co trafiło do pliku. W module 09 zobaczysz pełne wyjaśnienie, analogię do wydrukowanych etykiet i wzorzec poprawnego modyfikowania stylu przez `copy()`.

Do modułu 09 przygotuj też:

- **plik `output/08_cw1.xlsx`** — użyjemy go do budowania stylu nagłówka i zobaczymy, że styl zmieniony po kopiowaniu arkusza **nie propaguje się** na kopię; to bardzo praktyczne odkrycie, które w module 09 nazwiemy brakiem dziedziczenia i które w module 23 (rejestr stylów) zamienimy w rozwiązanie.
- **odpowiedź na pytanie:** w Twoim skoroszycie `output/08_cw1.xlsx` arkusz `_Konfiguracja` jest `veryHidden`. **Otwórz plik jako ZIP i znajdź jego zawartość.** Zapisz, w której części pakietu ujawnił się jego stan `veryHidden` (`xl/workbook.xml`) i w jakim pliku XML są jego komórki. Zajmie Ci to dwie minuty, a zapamiętasz to na całe życie — i będziesz o krok przed modułem 21.
- **jedną decyzję do podjęcia:** czy w Twoim skoroszycie `Podsumowanie` jest na początku (i użytkownik otwiera plik na wyniku), czy `Dane` (i użytkownik otwiera plik na źródle)? Nie ma tu jednej właściwej odpowiedzi — ale **jeśli nie potrafisz uzasadnić swojego wyboru jednym zdaniem, to znaczy, że nie został podjęty świadomie.** A cały ten moduł był właśnie o tym.