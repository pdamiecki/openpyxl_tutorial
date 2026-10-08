Zanim napiszę moduł, zweryfikuję dokładne API tabel w openpyxl 3.1.x — to obszar, gdzie w poradnikach krąży najwięcej nieistniejących metod i nieaktualnych obejść (m.in. słynny błąd z filtrami tabel i sposób nadawania nazw kolumn ze spacjami).

Potwierdzone kluczowe fakty (m.in. fix #1156 z 3.1.0). Pobiorę pełną stronę dokumentacji tabel i źródło modułu, żeby zweryfikować `ws.tables`, nadawanie nazw kolumn ze spacjami oraz walidację `displayName`.

Teraz zweryfikuję zachowanie tabel w trybach `read_only`/`write_only` oraz to, co dokładnie dzieje się przy duplikatach i nakładających się zakresach — to sedno pułapek tego modułu.

Ostatnia weryfikacja: obsługa tabel w trybie `write_only` (czy `add_table` istnieje i jak przechowuje tabele) oraz dokładne zachowanie `add_table` przy duplikatach.

Doskonale — mam potwierdzenie kluczowych mechanizmów (m.in. że `add_table` sprawdza duplikaty nazw **bez rozróżniania wielkości liter** i ostrzega w `write_only`). Zweryfikuję jeszcze, skąd openpyxl bierze nazwy kolumn tabeli przy zapisie — to decyduje o poprawności kilku przykładów.

Ostatnia weryfikacja — mechanizm zapisu tabel w 3.x (nadawanie `id`, relacje, nazwy kolumn z nagłówków). Chcę mieć pewność co do dwóch rzeczy, o których piszę wprost.

# Moduł 12 — Tabele: nazwane obiekty arkusza, nie tylko zakresy z kolorem

> **Część:** II — Formatowanie · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–11

## 0. W tym module nauczysz się

- **Zrozumiesz, czym jest Tabela (obiekt `ListObject`) i dlaczego nie jest tym samym co zakres z formatowaniem.** Zobaczysz, że tabela to osobny byt w pliku — nie fragment arkusza, tylko dołączona do niego część z własną nazwą, zakresem, kolumnami i stylem. To ta sama kategoria wiedzy co „XLSX to ZIP z XML" z modułu 02, ale na poziomie pojedynczego obiektu.
- **Nauczysz się tworzyć tabele programowo:** `Table(displayName=..., ref=...)`, `TableStyleInfo`, `ws.add_table()`. Zrozumiesz, które elementy tabela dostaje automatycznie, a które musisz jej dać.
- **Poznasz reguły nazywania tabel** — i dowiesz się, że openpyxl sprawdza **tylko jedną** z nich (spacje). Przekonasz się, dlaczego nazwa `Tbl1` przechodzi przez openpyxl, ale Excel odmawia otwarcia pliku, i nauczysz się generować nazwy bezpieczne.
- **Zrozumiesz, skąd biorą się nazwy kolumn tabeli.** To nieoczywiste: openpyxl **nie czyta ich z obiektu `Table`**, tylko z **komórek nagłówka w arkuszu** — i to dopiero w momencie zapisu. Zobaczysz, co z tego wynika i dlaczego w trybie `write_only` trzeba je nadać ręcznie.
- **Opanujesz structured references** (`=SUM(SprzedazTabela[Kwota])`) jako mechanizm „odwołania po tożsamości, nie po położeniu" — i zrozumiesz, dlaczego nazwa tabeli jest **kontraktem**, a nie dekoracją.
- **Nauczysz się odróżniać tabelę od autofiltra** (moduł 11) i wybierać świadomie między nimi. Zobaczysz, kiedy tabela pomaga, a kiedy przeszkadza.
- **Poznasz realne ograniczenia tabel w openpyxl:** brak automatycznego rozszerzania `ref` przy programowym dopisywaniu wierszy, brak sprawdzania nakładających się zakresów, niedostępność w `read_only`, ograniczona obsługa wiersza podsumowania.

## 1. Intuicja i analogia

### 1.1. Analogia główna: stos kartek kontra segregator

W module 11 poznałeś `ws.auto_filter.ref = "A1:H100"`. To był **zakres** z nałożonym sitkiem. Teraz poznasz coś jakościowo innego.

Wyobraź sobie, że na biurku leży **stos luźnych kartek**. Każda kartka to wiersz danych, a wszystkie razem tworzą kolumny od A do H. Kartki leżą w kolejności i można je przeglądać. To jest **zakres** (`A1:H100`).

Teraz wyobraź sobie **segregator**. Ma:

- **nazwę na grzbiecie** — na przykład `Sprzedaz2026`. Możesz powiedzieć „przynieś segregator Sprzedaz2026" i wszyscy wiedzą, o co chodzi. Przy stosie kartek musiałbyś powiedzieć „ten stos od kratki A1 do H100" — czyli opisać **położenie**, nie **tożsamość**.
- **etykiety przegród** — to nagłówki kolumn. Każda przegroda ma swoją nazwę („Klient", „Kwota", „Region").
- **zakładki-filtry** przyklejone do każdej przegrody — możesz przez nie przesiewać („pokaż tylko region Północ").
- **indeks** — dzięki któremu możesz się odwołać do zawartości po nazwie przegrody, a nie po jej numerze.

I najważniejsza różnica — ta, którą trzeba zapamiętać:

> **Gdy dokładasz kartkę do segregatora, segregator sam się rozszerza.** Wkładasz kartkę, a indeks, filtry i odwołania już wiedzą, że jest nowa.
>
> **Gdy dokładasz kartkę do stosu, stos jest większy — ale każde odwołanie „od A1 do H100" nadal kończy się na H100.** Nowa kartka leży poza zakresem.

Ta różnica to cała istota tabeli. Jest jednak jedno „ale", które omówię dokładnie w punkcie 2.12 i które jest najważniejszą pułapką tego modułu: **segregator rozszerza się sam, gdy kartkę wkłada człowiek — ale gdy kartkę wkłada program (openpyxl), trzeba mu powiedzieć.** Do tego wrócimy.

### 1.2. Analogia: nazwa tabeli to nazwa zmiennej

Pomyśl o swoim kodzie. Nigdy nie piszesz:

```python
zamowienia[3][7]     # co to za kolumna? trzeba policzyć...
```

gdy możesz napisać:

```python
zamowienie.kwota_netto     # wiadomo, co to jest
```

Nazwa tabeli robi dokładnie to samo dla Excela. Zamiast pisać formułę odwołującą się do **położenia**:

```excel
=SUMA(D2:D100)
```

możesz napisać formułę odwołującą się do **tożsamości**:

```excel
=SUMA(Sprzedaz2026[Kwota])
```

I teraz konsekwencja, która pokazuje, po co to jest. Wyobraź sobie, że ktoś **wstawia nową kolumnę** przed kolumną D (na przykład „Waluta"). Formuła `=SUMA(D2:D100)` teraz sumuje **Walutę**, nie Kwotę — i nikt tego nie zauważy, bo wynik będzie wyglądał wiarygodnie. Formuła `=SUMA(Sprzedaz2026[Kwota])` nadal sumuje Kwotę, bo odwołuje się po nazwie, nie po literze kolumny.

A teraz przełóż to na język architektury (i to jest zapowiedź modułu 26):

> **Nazwa tabeli jest kontraktem z użytkownikiem.** Kontraktem, czyli interfejsem, na którym opierają się formuły, tabele przestawne, wykresy i makra. Zmiana nazwy tabeli to jak zmiana nazwy publicznej funkcji w bibliotece, z której korzysta pięć innych programów: wszystko, co się do niej odwoływało, przestaje działać.

Dlatego do nazywania tabel podchodzi się tak samo jak do nazywania API: **raz, świadomie, i potem się tego trzyma.**

### 1.3. Analogia: filtry tabeli to zakładki przyklejone do segregatora

Idziesz do sklepu kupić segregator. I okazuje się, że **nie możesz go kupić bez zakładek**. Producent przykleił je na stałe. Możesz co najwyżej schować zakładkę, ale segregator zawsze pozostanie „segregatorem z zakładkami".

Tak działa tabela Excela z wierszem nagłówka. Dokumentacja mówi to wprost:

> „Filtry będą dodawane automatycznie do tabel, które mają wiersze nagłówka. Nie jest możliwe utworzenie tabeli z wierszem nagłówka bez filtrów."

To jest **kontrast wobec modułu 11**. Tam autofiltr był **opcjonalną nakładką**, którą mogłeś dołożyć albo nie. Tutaj filtr jest **częścią konstrukcji** — tabela z nagłówkiem ma filtr zawsze.

Ma to dwie praktyczne konsekwencje:

1. **Nie musisz (i nie możesz) „dodawać" filtra do tabeli** — on tam jest od początku.
2. **Jeśli nie chcesz filtra**, jedynym wyjściem jest tabela bez wiersza nagłówka (`headerRowCount = False`). Ale taka tabela ma inne problemy (zobaczysz w 2.4) i rzadko jest tym, czego chcesz.

Analogia do modułu 11 zostaje więc mocna: **autofiltr to sitko nałożone na stos kartek; tabela to segregator, w którym sitko jest wbudowane w każdą przegrodę.**

### 1.4. Analogia: tabela w pliku to kartka z opisem dołączona do teczki

W module 02 nauczyłeś się, że `.xlsx` to teczka (`ZIP`) pełna dokumentów (`XML`). Tabela w tej metaforze to **oddzielna kartka z opisem segregatora**, dołączona do teczki — plus **spinka** w arkuszu, która mówi „ten prostokąt komórek należy do opisu numer 1".

To znaczy, że definicja tabeli **nie jest w arkuszu**. Jest w osobnym dokumencie wewnątrz teczki. Jeśli otworzysz `.xlsx` jako ZIP (jak w module 02) i usuniesz tę kartkę, tabela zniknie — chociaż dane w arkuszu zostaną nietknięte. To dowód, że tabela jest **warstwą opisu**, a nie warstwą danych.

Wrócę do tego w 2.1 z konkretnymi ścieżkami.

## 2. Teoria

### 2.1. Czym właściwie jest tabela i gdzie mieszka

**Definicja robocza:** tabela to **nazwany zakres z wierszem nagłówka, automatycznymi filtrami, listą kolumn i przypisanym stylem** — zapisany w pliku jako osobny obiekt.

Zapamiętaj te cztery składniki: nazwa, zakres, kolumny, styl. Każdy z nich ma swoje miejsce w modelu openpyxl:

| Składnik | W obiekcie `Table` | Uwagi |
|---|---|---|
| Nazwa | `displayName` (i `name`) | To samo co „tożsamość" z analogii 1.2 |
| Zakres | `ref` | Prostokąt komórek, np. `"A1:E5"` — **musi obejmować nagłówki** |
| Kolumny | `tableColumns` | Lista obiektów `TableColumn` |
| Styl | `tableStyleInfo` | Obiekt `TableStyleInfo` |

A gdzie to mieszka w archiwum? W module 02 poznałeś strukturę pakietu. Tabela dodaje do niej **trzy nowe ślady**:

1. **`xl/tables/table1.xml`, `table2.xml`, …** — właściwa definicja tabeli. To ta „kartka z opisem segregatora" z analogii 1.4. Zawiera nazwę, zakres, kolumny, styl, filtr.
2. **`xl/worksheets/_rels/sheet1.xml.rels`** — relacja: „arkusz 1 odwołuje się do części tabeli `table1.xml`". To „spinka".
3. **`xl/worksheets/sheet1.xml`** — element `<tableParts>`, który wymienia identyfikatory relacji do tabel.

Schematycznie, w `sheet1.xml`:

```xml
<worksheet ...>
  <sheetData>...</sheetData>
  <tableParts count="1">
    <tablePart r:id="rId1"/>
  </tableParts>
</worksheet>
```

A w `table1.xml`:

```xml
<table xmlns="..." id="1" displayName="Sprzedaz2026" ref="A1:E5" headerRowCount="1">
  <autoFilter ref="A1:E5"/>
  <tableColumns count="5">
    <tableColumn id="1" name="Miesiąc"/>
    <tableColumn id="2" name="Sprzedaż"/>
    ...
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium9" showRowStripes="1"/>
</table>
```

**Co się dzieje w pamięci.** Kolekcja tabel arkusza to `ws._tables` — obiekt klasy `TableList`, który **dziedziczy po zwykłym słowniku Pythona** (`dict`) i kluczuje tabele po nazwie. Publiczny dostęp masz przez property `ws.tables`. To ważne, bo `TableList` **nadpisuje tylko trzy metody** (`add`, `get`, `items`) i zostawia resztę zachowań zwykłego słownika — z konsekwencjami, o których opowiem w 2.8.

**Co trafi do pliku.** Przy zapisie openpyxl serializuje każdy obiekt `Table` do osobnego `tableN.xml`, zakłada relacje i wstawia `<tableParts>`. **Identyfikatory tabel i relacji nadaje openpyxl przy zapisie** — Ty się tym nie przejmujesz. Ale to, co openpyxl zapisze, zależy od tego, co jest w obiekcie (nazwa, zakres, kolumny, styl).

### 2.2. Tabela kontra zakres — pełne porównanie

To zestawienie jest najważniejszą tabelą tego modułu. Wracaj do niego, ilekroć zastanawiasz się, czy użyć tabeli.

| Cecha | Zakres (z autofiltrem z modułu 11) | Tabela (`Table`) |
|---|---|---|
| **Czym jest** | Zbiór komórek opisany położeniem | **Obiekt** z nazwą i tożsamością |
| **Nazwa** | Brak (tylko adres typu `A1:H100`) | `displayName`, np. `Sprzedaz2026` |
| **Odwołanie w formule** | `=SUMA(D2:D100)` (po położeniu) | `=SUMA(Sprzedaz2026[Kwota])` (po nazwie) |
| **Rozszerzanie przy wpisywaniu w Excelu przez człowieka** | Nie | **Tak** — tabela rośnie sama |
| **Rozszerzanie przy dopisywaniu przez openpyxl** | Nie dotyczy | **Nie!** — trzeba poprawić `ref` ręcznie (2.12) |
| **Filtry** | Opcjonalne | **Zawsze obecne**, gdy jest wiersz nagłówka |
| **Style** | Trzeba nadać ręcznie (moduł 09) | Gotowe schematy kolorów (`TableStyleInfo`) |
| **Wiersz podsumowania** | Nie | Tak (`totalsRowCount`) — obsługa w openpyxl ograniczona |
| **Wykrywanie przez odbiorcę** | Użytkownik nie „widzi" granic zakresu | Użytkownik widzi „tabelę" i może się do niej odwołać po nazwie |
| **Gdzie w pliku** | W `sheetN.xml` (komórki + `<autoFilter>`) | Osobny `xl/tables/tableN.xml` + relacja |
| **Kiedy używać** | Jednorazowy, statyczny filtr nad danymi | Dane rosnące, na których stoją formuły i raporty przestawne |
| **Kiedy nie używać** | — | Ogromne pliki, dane tymczasowe, wielopoziomowe nagłówki (2.11) |

**Reguła kciuka:**

> Jeśli dane są **statyczne** i tylko je prezentujesz — wystarczy zakres z autofiltrem.
>
> Jeśli dane **rosną** i inne arkusze, formuły albo tabele przestawne się do nich odwołują — użyj tabeli.

### 2.3. Nazwa tabeli — reguły, których openpyxl NIE sprawdza

To najbardziej „podstępna" część tego modułu i miejsce, w którym traci się najwięcej czasu na szukaniu błędów, bo objaw pojawia się **u odbiorcy** (Excel mówi „plik wymaga naprawy"), a nie u Ciebie.

Zacznijmy od tego, co openpyxl **sprawdza**. W źródłach openpyxl 3.1.x jest klasa deskryptora nazwy tabeli (jej kod nie jest długi, więc przytoczę istotę):

```python
class TableNameDescriptor(String):
    """Table names cannot have spaces in them"""
    def __set__(self, instance, value):
        if value is not None and " " in value:
            raise ValueError("Table names cannot have spaces")
        super().__set__(instance, value)
```

Czyli: **jeśli nazwa zawiera spację, openpyxl od razu rzuca `ValueError`** — przy tworzeniu obiektu `Table`. To jest głośny, dobry błąd: dowiesz się o nim natychmiast.

**Ale to wszystko.** openpyxl nie sprawdza niczego więcej. W szczególności **nie sprawdza**:

- czy nazwa **nie wygląda jak adres komórki** (np. `Tbl1`, `ABC123`),
- czy nazwa **nie zawiera polskich znaków**,
- czy nazwa **nie zaczyna się od cyfry**,
- czy nazwa **nie zawiera myślnika** albo innych znaków specjalnych.

To znaczy, że openpyxl pozwoli Ci utworzyć `Table(displayName="Tbl1", ...)` — i **plik się zapisze**. Problem pojawi się dopiero, gdy otworzysz go w Excelu.

Zobaczmy to na konkretach:

```python
from openpyxl.worksheet.table import Table

# --- Te nazwy openpyxl PRZYJMIE bez mrugnięcia okiem ---

# 1. Nazwa wygladajaca jak ADRES KOMORKI - zle!
#    "TBL" to poprawna nazwa kolumny Excela (kolumna TBL istnieje,
#    bo Excel ma kolumny do XFD = 16384; TBL to 13584).
#    "TBL1" to wiec poprawny adres komorki -> Excel odrzuci taka nazwe tabeli.
t1 = Table(displayName="Tbl1", ref="A1:C4")          # ❌ Excel: "plik wymaga naprawy"

# 2. Inny przyklad adresu: ABC123 (ABC = kolumna 731, poprawna)
t2 = Table(displayName="ABC123", ref="A1:C4")        # ❌ to tez wyglada jak adres

# 3. Nazwa z polskim znakiem - ryzyko!
t3 = Table(displayName="Sprzedaż2026", ref="A1:C4")  # ⚠️ openpyxl przepusci, Excel - sprawdz!

# 4. Nazwa z myslnikiem - ryzyko!
t4 = Table(displayName="Tabela-1", ref="A1:C4")      # ⚠️ openpyxl przepusci, Excel - sprawdz!

# --- Ta nazwa jest bezpieczna ---
t5 = Table(displayName="Sprzedaz2026", ref="A1:C4")  # ✅ bez polskich znakow, bez myslnika
t6 = Table(displayName="tbl_Sprzedaz_2026", ref="A1:C4")  # ✅ prefiks tbl_ - najbezpieczniej
```

Dlaczego `Tbl1` to problem akurat teraz, a nie w openpyxl? Bo Excel traktuje nazwy tabel **tak samo jak nazwy definiowane** (defined names) — a jedną z reguł nazw definiowanych jest „nazwa nie może wyglądać jak adres komórki". `TBL1` **jest** poprawnym adresem (kolumna `TBL` mieści się w zakresie Excela), więc Excel odmawia. `Tabela1` jest bezpieczne, bo `TABELA` nie jest poprawną nazwą kolumny (kolumny kończą się na `XFD`).

**Pełne reguły bezpiecznej nazwy tabeli** (zebrane w jedno miejsce, bo tego szukają wszyscy):

1. Musi **zaczynać się od litery, podkreślnika (`_`) lub backslasha (`\`)** — nie od cyfry.
2. Może zawierać **litery, cyfry, podkreślniki, backslashe, kropki**.
3. **Nie może zawierać spacji** (openpyxl to sprawdza).
4. **Nie może zawierać myślnika** ani innych znaków specjalnych.
5. **Nie może wyglądać jak adres komórki** (`A1`, `Tbl1`, `ABC123`).
6. **Nie może być równa `R` ani `C`** (kolizja z notacją wiersz/kolumna w makrach).
7. Maksymalnie 255 znaków.
8. **Unikalna w całym skoroszycie, bez rozróżniania wielkości liter** (`Tabela1` i `TABELA1` to ta sama nazwa — zobacz 2.10).
9. **Polskie znaki: technicznie Excel je przyjmuje, ale unikaj** — różne wersje Excela i inne programy mogą zachować się nieprzewidywalnie wokół znaków diakrytycznych w nazwach obiektów.

Praktyczny generator bezpiecznych nazw — miej taką funkcję w projekcie:

```python
"""Generator bezpiecznych nazw tabeli."""

from __future__ import annotations

import re
import unicodedata

# Znaki dozwolone w nazwie tabeli (po normalizacji)
_NIEDOZWOLONE = re.compile(r"[^A-Za-z0-9_.]")
# Czy nazwa wyglada jak adres komorki (max 3 litery + cyfry)?
_WYGLADA_JAK_ADRES = re.compile(r"^[A-Za-z]{1,3}\d+$")
# Poprawne maksymalne kolumny Excela: XFD (16384)
_MAKS_KOLUMNA = 16384


def _kolumna_na_numer(litery: str) -> int:
    """Zamienia litery kolumny na numer (A=1, Z=26, AA=27, XFD=16384)."""
    numer = 0
    for znak in litery.upper():
        numer = numer * 26 + (ord(znak) - ord("A") + 1)
    return numer


def czy_adres_komorki(nazwa: str) -> bool:
    """Sprawdza, czy nazwa wyglada jak poprawny adres komorki Excela."""
    dopasowanie = _WYGLADA_JAK_ADRES.match(nazwa)
    if not dopasowanie:
        return False
    # Wyciagnij litery i sprawdz, czy kolumna miesci sie w zakresie Excela
    litery = re.match(r"^[A-Za-z]+", nazwa).group(0)
    return _kolumna_na_numer(litery) <= _MAKS_KOLUMNA


def bezpieczna_nazwa_tabeli(propozycja: str, prefiks: str = "tbl_") -> str:
    """Zamienia dowolna propozycje nazwy tabeli w nazwe bezpieczna dla Excela.

    - usuwa polskie znaki (zamienia na odpowiedniki ASCII),
    - usuwa niedozwolone znaki (spacje, myslniki itd.),
    - dodaje prefiks, zeby nazwa nie wygladala jak adres komorki,
    - pilnuje, zeby nazwa nie zaczynala sie od cyfry.
    """
    # 1. Rozloz znaki diakrytyczne na odpowiedniki ASCII
    #    (ą -> a, ś -> s, ż -> z ...)
    znormalizowane = unicodedata.normalize("NFKD", propozycja)
    ascii_only = znormalizowane.encode("ascii", "ignore").decode("ascii")

    # 2. Usun wszystko poza dozwolonymi znakami; spacje i myslniki -> podkreslnik
    oczyszczone = _NIEDOZWOLONE.sub("_", ascii_only)

    # 3. Sklej powtorzone podkreslniki i przytnij z brzegow
    oczyszczone = re.sub(r"_+", "_", oczyszczone).strip("_")

    # 4. Zbuduj nazwe z prefiksem (prefiks gwarantuje, ze nie jest adresem)
    nazwa = f"{prefiks}{oczyszczone}" if oczyszczone else prefiks.rstrip("_")

    # 5. Gdyby mimo wszystko wygladala jak adres - dodaj podkreslnik
    if czy_adres_komorki(nazwa):
        nazwa = f"{nazwa}_"

    return nazwa[:255]


if __name__ == "__main__":
    for propozycja in [
        "Sprzedaż 2026",       # spacja + polski znak
        "Tbl1",                # wyglada jak adres
        "Raport-Kwartalny-Q1", # myslniki
        "2026_sprzedaz",       # zaczyna sie od cyfry
        "ŚĆĘŁÓŻ",              # same polskie znaki
    ]:
        print(f"  {propozycja!r:<24} -> {bezpieczna_nazwa_tabeli(propozycja)!r}")
```

Wynik:

```
  'Sprzedaż 2026'          -> 'tbl_Sprzedaz_2026'
  'Tbl1'                   -> 'tbl_Tbl1'
  'Raport-Kwartalny-Q1'    -> 'tbl_Raport_Kwartalny_Q1'
  '2026_sprzedaz'          -> 'tbl_2026_sprzedaz'
  'ŚĆĘŁÓŻ'                 -> 'tbl_SCELOZ'
```

**Konsekwencja praktyczna:** buduj nazwy tabel **programowo**, przez funkcję, nigdy nie wpisuj ich „z ręki" w 30 miejscach. To ta sama zasada co rejestr stylów z modułu 09 i rejestr formatów z modułu 10 — **jedno źródło prawdy dla nazw**.

### 2.4. Struktura tabeli: nagłówek, kolumny, filtr

Trzy atrybuty, które decydują o tym, czym tabela jest:

**`headerRowCount`** — ile wierszy nagłówka ma tabela. Domyślnie `1`. Może być `0` (tabela bez nagłówka) albo większą liczbą, ale w praktyce liczba większa niż 1 rzadko działa dobrze (Excel traktuje nagłówek tabeli jako **jeden** wiersz). Jeśli ustawisz `headerRowCount = 0`:

- tabela **nie dostaje filtrów** (bo filtry wymagają nagłówka),
- `_initialise_columns()` **nie nada** kolumnom nazw z komórek (bo nie ma wiersza nagłówka),
- tabela staje się „nazwanym zakresem" bez nagłówków — rzadko przydatna.

**`tableColumns`** — lista obiektów `TableColumn`, po jednym na kolumnę. Każda ma:

- `id` — numer kolumny (bezwzględny numer kolumny Excela, nie 1..n! Zobacz 2.5),
- `name` — nazwa kolumny (to, czego używasz w structured references),
- oraz szereg pól opcjonalnych: `totalsRowFunction`, `totalsRowLabel`, `calculatedColumnFormula` itd.

**`autoFilter`** — obiekt `AutoFilter` z zakresem. **Ustawiony domyślnie**, gdy tabela ma nagłówek. Nie da się go nie mieć.

### 2.5. Skąd biorą się nazwy kolumn — mechanizm, który zaskakuje

To jest miejsce, w którym intuicja zawodzi najwięcej osób. Przeczytaj uważnie.

Wyobraź sobie, że tworzysz tabelę tak (jak w dokumentacji openpyxl):

```python
ws.append(["Fruit", "2011", "2012"])   # naglowki kolumn w arkuszu
ws.append(["Apples", 10000, 5000])     # dane
ws.append(["Pears", 2000, 3000])

tab = Table(displayName="Table1", ref="A1:C3")
```

Zauważ: **nie podałeś nigdzie nazw kolumn w obiekcie `Table`**. A jednak po zapisie plik ma tabelę z kolumnami `Fruit`, `2011`, `2012`. Skąd openpyxl je wziął?

**Z komórek arkusza — i to dopiero przy zapisie.** Mechanizm, potwierdzony w źródłach openpyxl, jest taki:

1. W momencie zapisu openpyxl sprawdza, czy `table.tableColumns` jest puste.
2. Jeśli puste, wywołuje wewnętrznie `_initialise_columns()`, które tworzy po jednym `TableColumn` na każdą kolumnę zakresu `ref`.
3. Następnie **czyta wiersz nagłówka z arkusza** (pierwszy wiersz zakresu `ref`) i przepisuje wartości tych komórek do `column.name`.
4. Jeśli któraś komórka nagłówka **nie jest napisem** (na przykład jest liczbą albo `None`), openpyxl emituje ostrzeżenie: *„File may not be readable: column headings must be strings."*

**Wniosek, który trzeba zapamiętać:**

> **Nazwy kolumn tabeli nie są własnością obiektu `Table` — pochodzą z komórek nagłówka arkusza.** Obiekt `Table` przechowuje je tylko jako kopię zapisaną przy zapisie.

A `_initialise_columns()` — jak nazywa kolumny, zanim odczyta komórki? Tworzy `TableColumn(id=idx, name="Column{idx}")`, gdzie **`idx` to bezwzględny numer kolumny**, wyliczony z `range_boundaries(ref)`. To znaczy: tabela zaczynająca się w kolumnie C dostanie tymczasowe nazwy `Column3`, `Column4`, `Column5` (nie `Column1`, `Column2`, ...). Zwykle nie ma to znaczenia, bo i tak zaraz nazwy zostaną nadpisane wartościami z nagłówków — ale jeśli w trybie `write_only` zapomnisz nadać nazwy, w Excelu zobaczysz kolumny o nazwach `Column5`, `Column6`... I to jest dokładnie ten objaw, po którym rozpoznasz ten problem.

**Konsekwencja praktyczna nr 1:** nagłówki w arkuszu **muszą być napisami**. Jeśli masz w nagłówku rok jako liczbę (`2026`), openpyxl ostrzeże i Excel może uznać plik za uszkodzony. Napraw to, zapisując nagłówki jako napisy (`"2026"`).

**Konsekwencja praktyczna nr 2:** jeśli chcesz nadać nazwy kolumn **świadomie** (np. inne niż w nagłówkach arkusza — nie zalecam, ale bywa potrzebne), możesz je ustawić ręcznie:

```python
tab = Table(displayName="Tbl1", ref="A1:C3")
tab._initialise_columns()                      # utworz liste kolumn
for kolumna, nazwa in zip(tab.tableColumns, ["Produkt", "Rok_2026", "Rok_2027"]):
    kolumna.name = nazwa
ws.add_table(tab)
```

⚠️ **Uwaga:** `_initialise_columns` ma podkreślnik na początku, co w Pythonie jest **umownym znakiem „to jest wewnętrzne"**. Dokumentacja openpyxl sama go używa w przykładzie dla trybu `write_only`, więc jest to „półpubliczne" — ale traktuj je ostrożnie: w nowszych wersjach biblioteki jego zachowanie może się zmienić. **Sprawdź w dokumentacji swojej wersji**, jeśli budujesz coś trwałego. To wzorzec, który powróci w module 18: podkreślnik oznacza „wiem, że dotykam wnętrza, i mam do tego powód".

**Konsekwencja praktyczna nr 3 — najważniejsza dla `write_only`:** w trybie `write_only` openpyxl **nie ma dostępu do komórek** (bo pisze strumieniowo i niczego nie trzyma w pamięci). Więc **nie może odczytać nagłówków**. Dlatego:

- `ws.add_table()` w `write_only` **emituje ostrzeżenie**: *„In write-only mode you must add table columns manually"*,
- musisz albo nadać kolumny ręcznie (`_initialise_columns` + pętla), albo ustawić `headerRowCount = False`.

### 2.6. Structured references — odwołania po tożsamości

Wracamy do analogii 1.2. Structured references to język odwołań do tabeli po nazwie i nazwie kolumny. Oto podstawowe formy:

| Forma | Znaczenie |
|---|---|
| `Sprzedaz2026[Kwota]` | Cała kolumna „Kwota" tabeli |
| `SUM(Sprzedaz2026[Kwota])` | Suma tej kolumny |
| `Sprzedaz2026[@Kwota]` | Wartość „Kwota" z **bieżącego wiersza** |
| `Sprzedaz2026[[#Nagłówki],[Kwota]]` | Komórka nagłówka kolumny „Kwota" |
| `Sprzedaz2026[#Wszystkie]` | Cała tabela z nagłówkami |
| `Sprzedaz2026[#Dane]` | Tylko dane (bez nagłówków i sum) |
| `Sprzedaz2026[[#Dane],[Kwota]]` | Dane z jednej kolumny |
| `SUMA(Sprzedaz2026[@[Kwota netto]:[Kwota brutto]])` | Zakres między dwiema kolumnami w bieżącym wierszu |

**Dwie przestrogi, obie ważne:**

⚠️ **Przestroga 1 — openpyxl nie waliduje formuł.** Zapisujesz napis. Jeśli zrobisz literówkę w nazwie kolumny, openpyxl tego nie zauważy. Excel przy otwarciu pokaże `#REF!` albo błąd. To ta sama zasada co w module 06: **openpyxl zapisuje przepis, nie sprawdza, czy danie wyjdzie.**

⚠️ **Przestroga 2 — specyfikatory lokalizacyjne.** Formuła jest przechowywana w pliku w **kanonicznej (angielskiej) formie**, a Excel tłumaczy ją przy wyświetlaniu na język interfejsu. Dlatego specyfikatory pisz **w formie kanonicznej**: `[#All]`, `[#Headers]`, `[#Data]`, `[#Totals]`, `[@]`. Nazwy kolumn pozostają takie, jakie są w nagłówkach (mogą być po polsku). **Sprawdź to w swojej wersji Excela**, bo zachowanie przy nietypowych formach bywa różne — to obszar, w którym warto eksperymentować, a nie zakładać.

**Jak zapisać structured reference w openpyxl:**

```python
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()
ws = wb.active
ws.title = "Dane"

ws.append(["Miesiąc", "Sprzedaż", "Koszt", "Marża"])
for miesiac, sprzedaz, koszt in [("Sty", 100, 60), ("Lut", 120, 70), ("Mar", 90, 80)]:
    ws.append([miesiac, sprzedaz, koszt])

tab = Table(displayName="Sprzedaz2026", ref="A1:D4")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tab)

# Kolumna "Marza" jest 4. kolumna -> wpisujemy formuly per wiersz
ws["D2"] = "=Sprzedaz2026[@Sprzedaż]-Sprzedaz2026[@Koszt]"
ws["D3"] = "=Sprzedaz2026[@Sprzedaż]-Sprzedaz2026[@Koszt]"
ws["D4"] = "=Sprzedaz2026[@Sprzedaż]-Sprzedaz2026[@Koszt]"

# Podsumowanie w innej komorce - poza zakresem tabeli
ws["F2"] = "=SUM(Sprzedaz2026[Sprzedaż])"
ws["F3"] = "=SUM(Sprzedaz2026[Koszt])"

wb.save("output/12_structured_refs.xlsx")
```

**Co trafi do pliku:** formuły zostaną zapisane **jako tekst** (openpyxl nie oblicza ich, moduł 06). Excel policzy je przy otwarciu. Odwołanie `Sprzedaz2026[Sprzedaż]` działa nawet wtedy, gdy ktoś wstawi nową kolumnę przed „Sprzedaż" — bo opiera się na nazwie kolumny, nie na literze.

⚠️ **Ważna uwaga o zakresie tabeli a formułach.** Structured reference typu `Sprzedaz2026[@...]` **musi być wewnątrz zakresu tabeli `ref`** (wiersz bieżący) albo poza nim, ale wskazując na kolumnę tabeli. Jeśli wpiszesz formułę z odwołaniem do tabeli w wierszu, który **nie należy** do zakresu `ref`, Excel zgłosi błąd. To pokazuje, że **`ref` musi być poprawny i kompletny** — inaczej psuje nie tylko wygląd, ale i logikę.

### 2.7. Style tabel — gotowe schematy kolorów

W module 09 musiałeś ręcznie budować `Font`, `PatternFill`, `Border` i nakładać je na każdą komórkę. Tabela dostaje styl **za darmo** — wystarczy nazwać schemat.

```python
from openpyxl.worksheet.table import TableStyleInfo

styl = TableStyleInfo(
    name="TableStyleMedium9",       # nazwa schematu
    showFirstColumn=False,          # wyroznij pierwsza kolumne (zwykle False)
    showLastColumn=False,           # wyroznij ostatnia kolumne (zwykle False)
    showRowStripes=True,            # pasy na wierszach (zebra) - czytelnosc
    showColumnStripes=False,        # pasy na kolumnach (rzadko)
)
tab.tableStyleInfo = styl
```

**Jakie nazwy stylów są poprawne?** Openpyxl definiuje trzy rodziny schematów. Ich pełna lista w 3.1.x to:

| Rodzina | Zakres numerów | Charakter |
|---|---|---|
| `TableStyleLight1` … `TableStyleLight21` | 21 schematów | Subtelne, jasne tła |
| `TableStyleMedium1` … `TableStyleMedium28` | 28 schematów | Wyraźne, kolorowe (te najczęściej się wybiera) |
| `TableStyleDark1` … `TableStyleDark11` | 11 schematów | Mocne, ciemne tła |

Nazwy są **dokładnie** takie, jak w Excelu — te, które widzisz w galerii stylów tabeli. Jeśli użyjesz nazwy spoza tych trzech rodzin, Excel może zignorować styl (nie wyrzuci błędu, po prostu tabela nie będzie miała nadanego wyglądu).

**Trzy uwagi praktyczne:**

1. **`showRowStripes=True` to najtańsza poprawa czytelności długiej tabeli.** Naprzemienne jasne wiersze pomagają wzrokowi nie zgubić się w wierszu — dokładnie jak linijka pod tekstem.
2. **Nie mieszaj stylu tabeli z ręcznym `PatternFill` na tych samych komórkach bez zrozumienia priorytetów.** Excel ma złożoną kolejność pierwszeństwa między formatowaniem bezpośrednim komórki, stylem tabeli, a formatowaniem warunkowym (moduły 13–14). Jeśli malujesz komórki ręcznie i jednocześnie nadajesz styl tabeli, wynik może Cię zaskoczyć. **Wybierz jedno podejście i sprawdź efekt wzrokowo.**
3. **Styl tabeli nie działa bez tabeli.** To brzmi oczywisto, ale zdarza się pomylić: `TableStyleInfo` przypisane do obiektu, który **nie został dodany** przez `ws.add_table()`, nie zrobi nic.

### 2.8. Kolekcja `ws.tables` — i jej trzy niespodzianki

`ws.tables` to obiekt `TableList`, który **dziedziczy po `dict`**. To znaczy, że działa jak słownik, ale openpyxl nadpisał w nim tylko trzy metody. Efekt jest taki, że API wygląda znajomo, ale w trzech miejscach zachowuje się inaczej, niż podpowiada intuicja (i w jednym miejscu inaczej, niż podpowiada **dokumentacja**).

**Podstawowe operacje:**

```python
# Liczba tabel w arkuszu
len(ws.tables)                             # np. 2

# Dostep po nazwie (jak w slowniku)
tab = ws.tables["Sprzedaz2026"]

# Iteracja po obiektach Table
for tabela in ws.tables.values():
    print(tabela.displayName, tabela.ref)

# Usuniecie tabeli (sam wpis; dane w komorkach zostaja!)
del ws.tables["Sprzedaz2026"]
```

**Niespodzianka 1 — `items()` zwraca `(nazwa, zakres)`, nie `(nazwa, obiekt)`.**

To jest **niezgodne z zachowaniem zwykłego słownika**, w którym `items()` daje pary klucz-wartość, gdzie wartością jest obiekt. Tutaj openpyxl nadpisał `items()` tak, że zwraca **zakres jako napis**:

```python
>>> ws.tables.items()
[("Sprzedaz2026", "A1:D10")]     # <- drugi element to NAPIS z zakresem, nie obiekt Table
```

Więc:

```python
for nazwa, obiekt_lub_zakres in ws.tables.items():
    # Uwaga! obiekt_lub_zakres to NAPIS "A1:D10", nie obiekt Table!
    ...
```

Jeśli chcesz **obiekty**, użyj `.values()`. Jeśli chcesz **pary nazwa-obiekt**, zbuduj je sam:

```python
# Chcesz pary (nazwa, OBIEKT)? Zbuduj sam:
for nazwa in ws.tables:
    tabela = ws.tables[nazwa]          # to jest obiekt Table
    print(nazwa, tabela.ref, tabela.column_names)
```

Ten szczegół jest źródłem błędów typu `AttributeError: 'str' object has no attribute 'ref'` u osób, które napisały `for name, table in ws.tables.items()` i potem wołały `table.ref`.

**Niespodzianka 2 — wyszukiwanie po zakresie wymaga `get(table_range=...)`.**

Dokumentacja openpyxl zawiera przykład `ws.tables["A1:D10"]`, sugerujący, że można pobrać tabelę po zakresie przez nawiasy kwadratowe. **Ale `TableList` nie nadpisuje `__getitem__`** — dziedziczy je ze zwykłego słownika, gdzie kluczem jest **nazwa**. Więc:

```python
>>> ws.tables["A1:D10"]
KeyError: 'A1:D10'          # ⚠️ tablica nie ma klucza "A1:D10"
```

Poprawny sposób wyszukania tabeli po zakresie to nadpisana metoda `get`:

```python
tabela = ws.tables.get(table_range="A1:D10")     # ✅ to dziala
```

⚠️ **Uważaj na pułapkę argumentu:** `get` przyjmuje `name` **jako pierwszy parametr pozycyjny**. Więc `ws.tables.get("A1:D10")` **nie** szuka po zakresie — szuka po **nazwie** `"A1:D10"` i zwróci `None`. Musisz podać jawnie `table_range=`.

To jest dokładnie ta kategoria, o której mówi moduł 03: **dokumentacja i kod nie zawsze mówią to samo.** Gdy coś „nie działa mimo że jest w dokumentacji", zajrzyj do źródła — tak jak zrobiliśmy tutaj.

**Niespodzianka 3 — `del ws.tables[...]` usuwa opis, nie dane.**

Usunięcie tabeli z kolekcji usuwa **definicję** tabeli (jej nazwę, styl, filtr, status obiektu). **Komórki w arkuszu zostają nietknięte.** To potwierdza tezę z 2.1: tabela to warstwa opisu, nie danych. Analogia: **zdjąłeś etykietę z segregatora** — kartki nadal leżą tam, gdzie leżały.

### 2.9. Ograniczenia, których openpyxl NIE pilnuje za Ciebie

Zebrałem tu wszystko, co openpyxl **przepuści**, a co Excel odrzuci albo zepsuje. To jest miejsce, w którym zapłacisz najwięcej czasu, jeśli nie będziesz ostrożny.

**A) Jak wygląda dokładnie to, co openpyxl sprawdza przy `add_table`?**

Warto zobaczyć to raz w źródle, bo jest pouczające:

```python
def add_table(self, table):
    """
    Check for duplicate name in definedNames and other worksheet tables
    before adding table.
    """
    if self.parent._duplicate_name(table.name):
        raise ValueError("Table with name {0} already exists".format(table.name))
    if not hasattr(self, "_get_cell"):
        warn("In write-only mode you must add table columns manually")
    self._tables.add(table)
```

Czyli przy dodawaniu tabeli openpyxl sprawdza **dwie rzeczy**:

1. **Czy nazwa nie powtarza się w skoroszycie** — pyta `wb._duplicate_name(table.name)`. Jeśli tak, rzuca `ValueError("Table with name ... already exists")`.
2. **Czy nie jesteśmy w trybie `write_only`** — jeśli tak (nie ma `_get_cell`), tylko **ostrzega**.

**I nic więcej.** Nie sprawdza nakładania się zakresów, nie sprawdza, czy nagłówki są napisami, nie sprawdza, czy zakres nie jest pusty.

**B) Nakładające się zakresy — openpyxl milczy.**

```python
ws.add_table(Table(displayName="Tabela_A", ref="A1:D10"))
ws.add_table(Table(displayName="Tabela_B", ref="C1:F10"))   # nachodzi na A!
```

openpyxl **przyjmie obie**. Zapisze plik. Problem pojawi się przy otwarciu w Excelu: dwie tabele na przecinających się komórkach to dla Excela błąd. Dlatego napisanie własnego **detektora nakładania się** jest warte zachodu — zrobimy to w ćwiczeniu 🔴.

Dlaczego openpyxl tego nie sprawdza? Bo sprawdzenie nakładania się prostokątów dla każdej pary tabel w każdym arkuszu to praca, którą biblioteka świadomie zostawia Tobie. To ta sama filozofia co w module 07 z `max_row`: **openpyxl robi dokładnie to, o co go poprosisz, i nie zgaduje Twoich intencji.**

**C) Tabela w scalonym zakresie — nie rób tego.**

Scalanie komórek (moduł 11) usuwa wartości i zamienia komórki na `MergedCell`. Tabela na takim obszarze ma niespójny nagłówek i Excel się gubi. **Nigdy nie umieszczaj tabeli w scalonym zakresie** i nigdy nie scalaj komórek wewnątrz tabeli.

**D) Tabela z wierszem nagłówka i zerem wierszy danych — Excel odrzuca.**

```python
ws.append(["Produkt", "Sztuki"])       # tylko naglowek
tabela = Table(displayName="Tbl1", ref="A1:B1")   # ⚠️ naglowek bez danych!
```

Tabela o zakresie `A1:B1` (albo `A1:B2` przy pustych danych) **zostanie zapisana**, ale Excel potraktuje ją jako uszkodzoną. **Reguła: zakres tabeli musi zawierać nagłówek plus co najmniej jeden wiersz danych.** Zabezpieczaj się warunkiem:

```python
if ws.max_row > 1:                  # przynajmniej naglowek + 1 wiersz danych
    ws.add_table(tabela)
```

To ekstremalnie częsty błąd w generatorach raportów: gdy dane wejściowe są puste (np. zero zamówień w danym miesiącu), tabela powstaje z samym nagłówkiem.

**E) Zduplikowane nazwy kolumn — Excel odrzuca (bez rozróżniania wielkości liter).**

Jeśli masz dwa nagłówki „Kwota" albo „Kwota" i „KWOTA", openpyxl je przepuści, a Excel uzna plik za uszkodzony. **Nagłówki kolumn tabeli muszą być unikalne**, i to bez rozróżniania wielkości liter. Dotyczy to zwłaszcza danych z pandas, gdzie duplikaty kolumn są możliwe.

**F) Nagłówki niebędące napisami — ostrzeżenie i ryzyko.**

Jak w 2.5: nagłówek, który jest liczbą, `datetime` albo `None`, wywoła ostrzeżenie openpyxl i może popsuć plik. **Rzutuj nagłówki na napisy.**

**G) Kropka w nazwie kolumny — ryzyko.**

Do structured references i do Excela lepiej, żeby nazwy kolumn nie zawierały kropek ani znaków `[`, `]`, `#`, `'`. Te znaki mają specjalne znaczenie w składni odwołań strukturalnych.

### 2.10. Unikalność nazw — cały skoroszyt, bez wielkości liter

Sprawdzenie duplikatów działa **na poziomie skoroszytu**, nie arkusza — i **ignoruje wielkość liter**. Potwierdza to test w źródłach openpyxl:

```python
def test_duplicate_table_name(self, Workbook, Table):
    wb = Workbook()
    ws = wb.create_sheet()
    ws.add_table(Table(displayName="Table1", ref="A1:D10"))
    assert True == wb._duplicate_name("Table1")     # ta sama nazwa
    assert True == wb._duplicate_name("TABLE1")     # inna wielkosc liter - TEZ duplikat
```

**Konsekwencje praktyczne:**

1. **Nie możesz mieć tabeli `Dane` w arkuszu 1 i tabeli `Dane` w arkuszu 2.** Druga próba `add_table` rzuci `ValueError`.
2. **`dane` i `DANE` to ta sama nazwa.** Nazwy tabel są porównywane bez rozróżniania wielkości liter.
3. **Nazwy tabel kolidują z nazwami definiowanymi.** `add_table` pyta `wb._duplicate_name`, a to sprawdza **zarówno tabele, jak i defined names**. Więc nie możesz mieć tabeli o nazwie, która jest już użyta jako nazwa definiowana — i odwrotnie.
4. **Konsekwencja dla generowania:** jeśli generujesz tabele w pętli, **buduj nazwę z licznikiem albo z nazwy arkusza**, żeby uniknąć kolizji:

```python
for indeks, (region, dane) in enumerate(regiony.items(), start=1):
    ws = wb.create_sheet(region)
    ...
    nazwa = bezpieczna_nazwa_tabeli(f"{region}_{indeks}")    # np. tbl_Polnoc_1
    ...
```

5. **Błąd duplikatu możesz zobaczyć nie tylko przy `add_table`, ale i przy `load_workbook`.** Jeśli wczytujesz plik, w którym (na skutek błędu w innym narzędziu) są dwie tabele o tej samej nazwie, `load_workbook` rzuci `ValueError: Table with name ... already exists` już na etapie czytania. To jest bardzo frustrujące, bo plik otwiera się w Excelu i w LibreOffice — a openpyxl go nie wczyta. Jedynym wyjściem jest naprawa pliku innym narzędziem (albo ręczne usunięcie jednej definicji z archiwum). **To argument za tym, żeby pilnować unikalności nazw po swojej stronie** — bo błąd, który popełnisz Ty, zamknie Ci drogę do wczytania pliku w przyszłości.

### 2.11. Kiedy tabela przeszkadza

Tabela nie zawsze jest właściwym rozwiązaniem. Trzy sytuacje, w których lepiej jej unikać:

**A) Ogromne pliki.** Tabela to dodatkowy obiekt z listą kolumn i relacjami. Na plikach rzędu setek tysięcy wierszy narzut jest niewielki w porównaniu z samymi danymi, ale **filtry i styl tabeli powodują dodatkową pracę przy renderowaniu w Excelu**. Jeśli generujesz plik, który ma być tylko magazynem danych (do wczytania przez program, nie przez człowieka), tabela nic nie daje — a dodaje zależności (relacje, osobne części w archiwum). W module 19 zobaczysz to w kontekście wydajności.

**B) Tabele tymczasowe i arkusze robocze.** Jeśli tworzysz arkusz pomocniczy, który i tak zostanie usunięty albo ukryty, tabela nie ma sensu. Filtr i nazwa nie są nikomu potrzebne.

**C) Dane z wieloma poziomami nagłówków.** Tabela Excela wymaga **jednego** wiersza nagłówka. Jeśli Twój raport ma dwa wiersze nagłówka (np. „2026" scalone nad „Q1/Q2/Q3/Q4") — tabela tego nie obsłuży. Zostaje zwykły zakres z ręcznym formatowaniem (moduły 09–11). To dokładnie ten przypadek, w którym `headerRowCount = 2` kusi, ale nie działa dobrze: Excel i tak traktuje nagłówek tabeli jako jeden wiersz.

**D) Tabela jako źródło prawdy w pipeline.** Jeśli dane mają być tylko przetworzone i wyeksportowane, **tabela jest warstwą prezentacji dla człowieka**. Nie buduj logiki na jej istnieniu — to zapowiedź modułu 26 (warstwy) i modułu 25 (Specification).

### 2.12. „Segregator sam się rozszerza" — ale nie wtedy, gdy pisze program

To najważniejsza praktyczna korekta analogii z 1.1 i najczęstsza przyczyna „dlaczego Excel nie widzi moich nowych wierszy".

**Jak działa rozszerzanie w Excelu.** Gdy **człowiek** w Excelu kliknie w komórkę bezpośrednio pod tabelą i zacznie pisać, Excel automatycznie wciąga nowy wiersz do tabeli. Mechanizm działa w **interfejsie** — to Excel reaguje na akcję użytkownika.

**Co robi openpyxl.** Openpyxl **nie ma pojęcia o interfejsie**. Gdy piszesz:

```python
ws.append(["Nowy", 100, 50, 50])      # dopisuje wiersz na koncu arkusza
wb.save("plik.xlsx")
```

openpyxl **nie wie**, że ten wiersz powinien należeć do tabeli. `Table.ref` pozostaje bez zmian (`A1:D4`), a nowy wiersz jest **poza** zakresem tabeli. Po otwarciu w Excelu:

- tabela nadal obejmuje tylko stary zakres,
- nowy wiersz ma dane, ale **nie jest częścią tabeli** (nie ma filtra, nie ma stylu, nie jest widziany w structured references),
- formuła `=SUM(Sprzedaz2026[Sprzedaż])` **nie uwzględni** nowego wiersza.

**Rozwiązanie: rozszerz `ref` ręcznie.**

```python
from openpyxl.utils import get_column_letter

# Po dopisaniu wierszy zaktualizuj zakres tabeli
tabela = ws.tables["Sprzedaz2026"]
ostatnia_litera = get_column_letter(ws.max_column)
tabela.ref = f"A1:{ostatnia_litera}{ws.max_row}"
```

**Analogia:** segregator rozszerza się sam, gdy kartkę wkłada **człowiek** (bo wie, że to segregator). Ale gdy kartkę wkłada **program** — program widzi tylko „kartka na biurku obok segregatora". Musisz mu powiedzieć: „ta kartka też należy do segregatora". To jest praca, którą w interaktywnym Excelu wykonuje interfejs, a w automatyzacji musisz wykonać sam.

To wspaniały przykład szerszej prawdy z modułu 03:

> **Openpyxl nie odtwarza tego, co robi Excel — odtwarza format pliku.** Wszystko, co w Excelu dzieje się „samo" (autokorekta, rozszerzanie tabel, inteligentne wypełnianie), jest funkcją **interfejsu**, nie **formatu**. I dlatego w automatyzacji trzeba to zrobić ręcznie.

**Konsekwencja architektoniczna (zapowiedź modułu 24 — Facade):** zamiast rozrzucać ten kod po skrypcie, warto mieć jedną funkcję „dopisz wiersz do tabeli", która **jednocześnie** dopisuje dane i rozszerza `ref`. Napiszemy ją w ćwiczeniu 🟡.

## 3. Przykłady krok po kroku

### Przykład 1 — podstawowa tabela z filtrem i stylem (🟢)

Zacznijmy od kompletnego, uruchamialnego przykładu — wraz z funkcją weryfikacji, bo (jak w module 11) zawsze chcesz zobaczyć, co naprawdę trafiło do pliku.

```python
"""Podstawowa tabela: nazwa, styl, filtr, weryfikacja.

Uruchom:  python examples/12_tabela_podstawowa.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "12_tabela_podstawowa.xlsx"

# --- dane testowe ---------------------------------------------------
NAGLOWKI = ["Miesiąc", "Sprzedaż", "Koszt", "Marża"]

DANE = [
    ("Styczeń", 120_000, 78_000),
    ("Luty",     98_500, 70_920),
    ("Marzec",  143_700, 122_145),
    ("Kwiecień", 131_200, 76_096),
    ("Maj",      156_800, 94_080),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaż"

    # --- naglowek jako NAPISY (2.5!) -----------------------------
    ws.append(NAGLOWKI)
    for miesiac, sprzedaz, koszt in DANE:
        ws.append([miesiac, sprzedaz, koszt, sprzedaz - koszt])

    # --- utworzenie tabeli ---------------------------------------
    ostatnia_litera = get_column_letter(len(NAGLOWKI))
    zakres = f"A1:{ostatnia_litera}{ws.max_row}"     # A1:D6

    tabela = Table(displayName="Sprzedaz2026", ref=zakres)

    # Styl: paski na wierszach to najtansza poprawa czytelnosci (2.7)
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    # UWAGA: dodawaj przez ws.add_table (sprawdza duplikaty nazw - 2.9A)
    ws.add_table(tabela)

    # --- formaty liczb (modul 10) --------------------------------
    for wiersz in range(2, ws.max_row + 1):
        for kolumna in (2, 3, 4):
            ws.cell(row=wiersz, column=kolumna).number_format = '#,##0 "zł"'

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Sprzedaż"]
        print("=" * 84)
        print("WERYFIKACJA PO ZAPISIE I ODCZYCIE")
        print("=" * 84)
        print(f"  liczba tabel w arkuszu:  {len(ws.tables)}")

        for nazwa, tabela in ((n, ws.tables[n]) for n in ws.tables):
            print(f"  nazwa tabeli:            {nazwa!r}")
            print(f"  zakres (ref):            {tabela.ref!r}")
            print(f"  liczba wierszy naglowka: {tabela.headerRowCount}")
            print(f"  nazwy kolumn:            {tabela.column_names}")
            print(f"  filtr obecny:            {tabela.autoFilter is not None}")
            print(f"  styl:                    {tabela.tableStyleInfo.name if tabela.tableStyleInfo else None}")
            print(f"  pokaz paskow:            {tabela.tableStyleInfo.showRowStripes if tabela.tableStyleInfo else None}")

        print()
        print(f"  .items() zwraca:         {ws.tables.items()}")
        print("  ^ drugi element to ZAKRES (napis), nie obiekt Table (2.8!)")
    finally:
        wb.close()

    print()
    print("=" * 84)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 84)
    print("  1. Zakres A1:D6 ma w rogu znacznik filtra (strzalki w naglowkach).")
    print("  2. Kliknij w dowolna komorke tabeli -> wstazka 'Projekt tabeli'.")
    print("     Excel pokaze NAZWE tabeli: Sprzedaz2026.")
    print("  3. Wiersze maja naprzemienne jasne tlo (paski).")
    print("  4. Wpisz cokolwiek w wiersz 7 (bezposrednio pod tabela)")
    print("     -> Excel SAM rozszerzy tabele (to robi INTERFEJS, nie plik!).")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Kluczowe punkty do przemyślenia:**

1. **Nagłówki jako napisy.** `ws.append(NAGLOWKI)` — wszystkie są `str`. To warunek poprawnego pliku (2.5).
2. **`ref` obejmuje nagłówki.** `A1:D6` — wiersz 1 to nagłówek, wiersze 2–6 to dane. To nie opcja, to wymóg: „ref musi obejmować nagłówki".
3. **`displayName` bezpieczne.** `Sprzedaz2026` — bez spacji, bez polskich znaków, nie wygląda jak adres komórki.
4. **`ws.add_table` zamiast `ws._tables.add`.** Ta pierwsza sprawdza duplikaty w całym skoroszycie (2.9A). Jeśli dodasz tabelę bezpośrednio do `_tables`, ominiesz ten test i sam sobie zrobisz krzywdę.
5. **Nie nadałem nazw kolumn ręcznie.** Nie musiałem — openpyxl odczyta je z wiersza nagłówka przy zapisie (2.5). Weryfikacja pokazuje, że `tabela.column_names` po odczycie zwraca `['Miesiąc', 'Sprzedaż', 'Koszt', 'Marża']`.

**Co trafi do pliku.** W archiwum znajdziesz:

- `xl/tables/table1.xml` — definicja tabeli `Sprzedaz2026` z zakresem, czterema kolumnami, filtrem i stylem,
- relację w `xl/worksheets/_rels/sheet1.xml.rels`,
- `<tableParts>` w `xl/worksheets/sheet1.xml`.

Możesz to zobaczyć, otwierając plik jako ZIP (jak w module 02) albo takim kodem:

```python
import zipfile

with zipfile.ZipFile("output/12_tabela_podstawowa.xlsx") as archiwum:
    for nazwa in archiwum.namelist():
        if "table" in nazwa.lower():
            print(nazwa)
# xl/tables/table1.xml
```

### Przykład 2 — dopisanie 200 wierszy i rozszerzenie `ref` (🟡)

To ćwiczenie pokazuje w praktyce korektę z 2.12. Zrobimy je **dwa razy**: raz tak, jak się nie robi (zapominając o `ref`) i raz poprawnie — żeby zobaczyć różnicę.

```python
"""Dopisywanie wierszy do tabeli: dlaczego 'sam sie rozszerza' to nieprawda.

Uruchom:  python examples/12_dopisywanie.py
"""

from __future__ import annotations

import random
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

LOSOWE = random.Random(42)     # staly seed -> powtarzalnosc


def _nowe_wiersze(ile: int):
    """Generuje 'ile' wierszy danych sprzedazowych."""
    for indeks in range(1, ile + 1):
        sprzedaz = LOSOWE.randint(5_000, 50_000)
        koszt = LOSOWE.randint(2_000, sprzedaz)
        yield (f"Faktura-{1000 + indeks:04d}", sprzedaz, koszt, sprzedaz - koszt)


def _zaloz_tabele(ws) -> None:
    """Tworzy tabele nad naglowkiem + pierwszym wierszem danych."""
    ws.append(["Numer", "Sprzedaż", "Koszt", "Marża"])
    ws.append(next(_nowe_wiersze(1)))
    tabela = Table(displayName="Faktury", ref=f"A1:D{ws.max_row}")
    tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True)
    ws.add_table(tabela)


def wariant_zly(cel: Path) -> None:
    """Dopisuje wiersze, ale NIE rozszerza ref - klasyczny blad."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Zły"
    _zaloz_tabele(ws)

    for wiersz in _nowe_wiersze(200):
        ws.append(wiersz)          # dane sa dopisane...

    # ...ale tabela nadal obejmuje tylko pierwszy wiersz danych!
    # tabela.ref pozostaje "A1:D2"

    wb.save(cel)
    wb.close()


def wariant_dobry(cel: Path) -> None:
    """Dopisuje wiersze ORAZ rozszerza ref tabeli."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Dobry"
    _zaloz_tabele(ws)

    for wiersz in _nowe_wiersze(200):
        ws.append(wiersz)

    # TU JEST CALA ROZNICA: reczne rozszerzenie zakresu tabeli (2.12)
    dopisz_wiersze_do_tabeli(ws, "Faktury", ile=200)

    wb.save(cel)
    wb.close()


def dopisz_wiersze_do_tabeli(ws, nazwa_tabeli: str, ile: int) -> None:
    """Rozszerza ref tabeli tak, by objal biezacy zakres danych arkusza.

    Zalozone: dane sa juz wpisane do arkusza, a tabela zaczyna sie w A1.
    """
    tabela = ws.tables[nazwa_tabeli]
    ostatnia_litera = get_column_letter(ws.max_column)
    nowy_ref = f"A1:{ostatnia_litera}{ws.max_row}"
    tabela.ref = nowy_ref


def porownaj(zly: Path, dobry: Path) -> None:
    print("=" * 84)
    print("POROWNANIE: ZAPOMNIANE ref KONTRA ROZSZERZONE ref")
    print("=" * 84)
    for opis, sciezka in (("ZAPOMNIANE", zly), ("ROZSZERZONE", dobry)):
        wb = load_workbook(scieka := sciezka)
        try:
            ws = wb.active
            tabela = ws.tables["Faktury"]
            wiersze_w_arkuszu = ws.max_row
            # zakres tabeli -> ile wierszy obejmuje?
            od, do = tabela.ref.split(":")
            ostatni_wiersz_tabeli = int("".join(ch for ch in do if ch.isdigit()))
            print(f"  {opis:<14} wierszy w arkuszu: {wiersze_w_arkuszu:>4} | "
                  f"ref tabeli: {tabela.ref!r:<12} | "
                  f"wierszy w tabeli: {ostatni_wiersz_tabeli}")
        finally:
            wb.close()
    print()
    print("  WNIOSEK: w wariancie 'ZAPOMNIANE' 201 wierszy danych istnieje w arkuszu,")
    print("  ale tabela obejmuje tylko 2. Nowe wiersze są POZA tabelą:")
    print("  brak filtra, brak stylu, nie widzi ich =SUM(Faktury[Sprzedaż]).")
    print()
    print("  Otwórz oba pliki w Excelu i sprawdź:")
    print("   - zakładka 'Zły': filtr i styl obejmują tylko pierwszy wiersz danych.")
    print("   - zakładka 'Dobry': filtr i styl obejmują wszystkie 201 wierszy.")


def main() -> None:
    zly = OUTPUT / "12_dopisywanie_zly.xlsx"
    dobry = OUTPUT / "12_dopisywanie_dobry.xlsx"
    wariant_zly(zly)
    wariant_dobry(dobry)
    porownaj(zly, dobry)


if __name__ == "__main__":
    main()
```

**Czego uczy najbardziej:** że **rozszerzanie tabeli to obowiązek programu, nie właściwość formatu.** To dokładnie ta lekcja z 2.12 — openpyxl odtwarza format pliku, a nie zachowanie interfejsu Excela.

I druga rzecz: funkcja `dopisz_wiersze_do_tabeli` to **zalążek dobrej abstrakcji**. Zauważ, że ma dokładnie jeden powód, by się zmienić (sposób wyliczania zakresu tabeli) i jedno zadanie. To zapowiedź wzorca **Facade** z modułu 24 — funkcja, za którą schowasz szczegóły uprzykrzające życie.

### Przykład 3 — wykrycie i naprawa nakładających się tabel (🔴)

Ćwiczenie z briefu, wykonane uczciwie: najpierw **tworzymy** plik z dwiema nakładającymi się tabelami (openpyxl to przepuści!), potem **detektorem** je wykrywamy, na koniec **naprawiamy**.

```python
"""Wykrywanie i naprawa nakladajacych sie tabel.

Uruchom:  python examples/12_nakladanie.py
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "12_nakladanie.xlsx"


@dataclass(frozen=True)
class Prostokat:
    """Prostokat opisany bezwzglednymi indeksami kolumn i wierszy."""
    min_col: int
    min_row: int
    max_col: int
    max_row: int

    @classmethod
    def z_zakresu(cls, zakres: str) -> "Prostokat":
        """Buduje prostokat z napisu zakresu, np. 'C1:F10'."""
        min_col, min_row, max_col, max_row = range_boundaries(zakres)
        return cls(min_col, min_row, max_col, max_row)

    def nachodzi_na(self, inny: "Prostokat") -> bool:
        """Czy dwa prostokaty maja wspolne komorki?

        Warunek rozlacznosci: jeden jest calkowicie po lewej, prawej,
        powyzej albo ponizej drugiego. Jesli zaden nie zachodzi - nachodza.
        """
        rozlaczne = (
            self.max_col < inny.min_col
            or inny.max_col < self.min_col
            or self.max_row < inny.min_row
            or inny.max_row < self.min_row
        )
        return not rozlaczne


def utworz_plik_z_nakladaniem(cel: Path) -> None:
    """Tworzy plik z DWOMA nachodzacymi tabelami.

    UWAGA: openpyxl NIE sprawdza nakladania (2.9B) - plik sie zapisze,
    ale Excel uzna go za uszkodzony.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Nakladanie"

    # Naglowek + dane w A1:F10
    ws.append(["A", "B", "C", "D", "E", "F"])
    for wiersz in range(2, 11):
        for kolumna in range(1, 7):
            ws.cell(row=wiersz, column=kolumna, value=wiersz * 10 + kolumna)

    styl = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)

    t_a = Table(displayName="Tabela_A", ref="A1:D10")
    t_a.tableStyleInfo = styl
    ws.add_table(t_a)

    # Ta tabela nachodzi na A1:D10 w obszarze C1:F10
    t_b = Table(displayName="Tabela_B", ref="C1:F10")
    t_b.tableStyleInfo = styl
    ws.add_table(t_b)          # openpyxl MILCZY (2.9B)

    wb.save(cel)
    wb.close()


@dataclass
class Konflikt:
    arkusz: str
    tabela_a: str
    zakres_a: str
    tabela_b: str
    zakres_b: str
    zakres_wspolny: str


def wykryj_nakladania(sciezka: Path) -> list[Konflikt]:
    """Znajduje wszystkie pary tabel o nachodzacych sie zakresach."""
    wb = load_workbook(sciezka)
    konflikty: list[Konflikt] = []
    try:
        for ws in wb.worksheets:
            tabele = [(nazwa, ws.tables[nazwa]) for nazwa in ws.tables]
            for indeks_a in range(len(tabele)):
                for indeks_b in range(indeks_a + 1, len(tabele)):
                    nazwa_a, tab_a = tabele[indeks_a]
                    nazwa_b, tab_b = tabele[indeks_b]
                    prost_a = Prostokat.z_zakresu(tab_a.ref)
                    prost_b = Prostokat.z_zakresu(tab_b.ref)
                    if prost_a.nachodzi_na(prost_b):
                        # Wylicz czesc wspolna
                        wspolny = Prostokat(
                            min_col=max(prost_a.min_col, prost_b.min_col),
                            min_row=max(prost_a.min_row, prost_b.min_row),
                            max_col=min(prost_a.max_col, prost_b.max_col),
                            max_row=min(prost_a.max_row, prost_b.max_row),
                        )
                        from openpyxl.utils import get_column_letter
                        zakres_wspolny = (
                            f"{get_column_letter(wspolny.min_col)}{wspolny.min_row}:"
                            f"{get_column_letter(wspolny.max_col)}{wspolny.max_row}"
                        )
                        konflikty.append(Konflikt(
                            arkusz=ws.title,
                            tabela_a=nazwa_a, zakres_a=tab_a.ref,
                            tabela_b=nazwa_b, zakres_b=tab_b.ref,
                            zakres_wspolny=zakres_wspolny,
                        ))
    finally:
        wb.close()
    return konflikty


def napraw(sciezka: Path) -> None:
    """Naprawia plik: sprowadza nachodzace tabele do jednej.

    Strategia: zostaw pierwsza, usun pozostale nachodzace.
    (W realnym projekcie: zdecyduj swiadomie, ktora tabela jest wlasciwa!)
    """
    wb = load_workbook(sciezka)
    try:
        for ws in wb.worksheets:
            tabele = [(n, ws.tables[n]) for n in ws.tables]
            do_usuniecia: set[str] = set()
            for i in range(len(tabele)):
                for j in range(i + 1, len(tabele)):
                    nazwa_a, tab_a = tabele[i]
                    nazwa_b, tab_b = tabele[j]
                    if nazwa_a in do_usuniecia:
                        continue
                    if Prostokat.z_zakresu(tab_a.ref).nachodzi_na(
                            Prostokat.z_zakresu(tab_b.ref)):
                        # Zachowujemy wieksza tabele, usuwamy mniejsza
                        obszar_a = (tab_a.ref, _liczba_komorek(tab_a.ref))
                        obszar_b = (tab_b.ref, _liczba_komorek(tab_b.ref))
                        mniejsza = nazwa_b if obszar_b[1] < obszar_a[1] else nazwa_a
                        do_usuniecia.add(mniejsza)
            for nazwa in do_usuniecia:
                del ws.tables[nazwa]        # usuwa DEFINICJE, nie dane (2.8)
        wb.save(sciezka)
    finally:
        wb.close()


def _liczba_komorek(zakres: str) -> int:
    min_col, min_row, max_col, max_row = range_boundaries(zakres)
    return (max_col - min_col + 1) * (max_row - min_row + 1)


def main() -> None:
    utworz_plik_z_nakladaniem(CEL)
    print(f"Utworzono plik z celowo nakladajacymi sie tabelami: {CEL}")

    konflikty = wykryj_nakladania(CEL)
    print("\n" + "=" * 84)
    print("WYKRYTE KONFLIKTY")
    print("=" * 84)
    for k in konflikty:
        print(f"  [{k.arkusz}] {k.tabela_a} {k.zakres_a}  <->  "
              f"{k.tabela_b} {k.zakres_b}   (wspolne: {k.zakres_wspolny})")

    napraw(CEL)
    konflikty_po = wykryj_nakladania(CEL)
    print("\n" + "=" * 84)
    print("PO NAPRAWIE")
    print("=" * 84)
    print(f"  liczba konfliktow: {len(konflikty_po)} (oczekiwane: 0)")

    wb = load_workbook(CEL)
    try:
        ws = wb.active
        print(f"  pozostale tabele:  {[n for n in ws.tables]}")
        print(f"  dane w A2 sa nienaruszone: {ws['A2'].value}")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
```

**Czego uczy najbardziej:** że **openpyxl nie jest strażnikiem poprawności Twojego pliku**. Przepuści nakładające się tabele, tabele z pustymi danymi, tabele z duplikatami kolumn — bo jego zadaniem jest **wierny zapis tego, co mu każesz**, a nie pilnowanie, czy to ma sens. To dokładnie ta filozofia z 2.9.

I druga lekcja, architektoniczna: `Prostokat` to mały, samodzielny **obiekt wartości** (value object) z metodą `nachodzi_na`. Nie ma w nim niczego o openpyxl poza klasową metodą `z_zakresu`. To znaczy, że **logikę nakładania możesz przetestować bez żadnego pliku** — czysta funkcja na dwóch prostokątach. To zapowiedź modułu 20 (testowanie) i 26 (oddzielenie domeny od infrastruktury): **im mniej openpyxl w Twojej logice, tym łatwiej ją sprawdzić.**

**Uwaga o `napraw`:** prawdziwy „naprawiacz" nie może zgadywać, którą tabelę zachować — to **decyzja biznesowa**. W przykładzie użyłem heurystyki „zachowaj większą", ale w realnym projekcie napisałbyś raport konfliktów i kazał człowiekowi zdecydować. To ta sama zasada co w module 11 przy `print_area`: **narzędzie ma raportować, a nie decydować za człowieka w sprawach, których nie rozumie.**

**Dlaczego to wykrywamy u siebie, a nie polegamy na Excelu:** bo gdy Excel powie „plik wymaga naprawy", użytkownik zobaczy komunikat, a Ty nie. Wykrycie konfliktu **przed wysłaniem pliku** to różnica między „działa" a „klient dzwoni". Napisz taki detektor raz i uruchamiaj go w pipeline.

### Przykład 4 — odczyt tabeli z istniejącego pliku i praca z kolumnami

Ostatni przykład: jak **czytać** tabele z cudzego pliku (i jak radzić sobie z tym, że `items()` zwraca zakres zamiast obiektu).

```python
"""Odczyt tabel z istniejacego pliku.

Uruchom:  python examples/12_odczyt_tabel.py
    (wymaga pliku z Przykladu 1: output/12_tabela_podstawowa.xlsx)
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

ROOT = Path(__file__).resolve().parent.parent
ZRODLO = ROOT / "output" / "12_tabela_podstawowa.xlsx"


def wypisz_tabele(sciezka: Path) -> None:
    wb = load_workbook(sciezka)
    try:
        for ws in wb.worksheets:
            if not len(ws.tables):
                continue
            print(f"\n=== Arkusz: {ws.title} (tabel: {len(ws.tables)}) ===")

            # UWAGA: .values() daje OBIEKTY Table; .items() daje (nazwa, ZAKRES)!
            for tabela in ws.tables.values():
                print(f"\n  Tabela: {tabela.displayName}")
                print(f"    zakres:       {tabela.ref}")
                print(f"    kolumny:      {tabela.column_names}")
                print(f"    naglowek:     {tabela.headerRowCount} wiersz(y)")

                # Odczyt danych z zakresu tabeli (modul 07)
                min_col, min_row, max_col, max_row = range_boundaries(tabela.ref)
                print("    dane:")
                # Wiersz min_row to naglowek -> dane od min_row+1
                for wiersz in ws.iter_rows(
                    min_row=min_row + 1, max_row=max_row,
                    min_col=min_col, max_col=max_col, values_only=True,
                ):
                    print(f"      {wiersz}")
    finally:
        wb.close()


def wyszukaj_po_zakresie(sciezka: Path, zakres: str) -> None:
    """Pokazuje poprawny sposob wyszukania tabeli po zakresie (2.8)."""
    wb = load_workbook(sciezka)
    try:
        ws = wb.active
        # ZLE: ws.tables["A1:D6"]  ->  KeyError (2.8, niespodzianka 2)
        # DOBRZE: metoda get z jawnym argumentem table_range
        tabela = ws.tables.get(table_range=zakres)
        if tabela is None:
            print(f"  Nie znaleziono tabeli o zakresie {zakres!r}")
        else:
            print(f"  Tabela w zakresie {zakres!r}: {tabela.displayName}")
    finally:
        wb.close()


def main() -> None:
    if not ZRODLO.exists():
        print(f"Brak pliku {ZRODLO}. Uruchom najpierw 12_tabela_podstawowa.py.")
        return

    wypisz_tabele(ZRODLO)
    print()
    print("=" * 70)
    wyszukaj_po_zakresie(ZRODLO, "A1:D6")


if __name__ == "__main__":
    main()
```

**Czego uczy najbardziej:** trzech rzeczy z sekcji 2.8, wszystkich w działaniu — że `.values()` daje obiekty, `.items()` daje napisy, a wyszukiwanie po zakresie wymaga `get(table_range=...)`. To dokładnie ta sama lekcja co w module 07 z `max_row`: **API wygląda znajomo, ale w szczegółach różni się od intuicji i od dokumentacji.**

## 4. Anatomia API

| Metoda / klasa / atrybut | Co robi | Parametry / wartości | Uwagi |
|---|---|---|---|
| `Table(displayName, ref)` | Tworzy definicję tabeli | `displayName` — nazwa; `ref` — zakres (musi obejmować nagłówki) | Konstruktor z `openpyxl.worksheet.table` |
| `Table.headerRowCount` | Liczba wierszy nagłówka | `int`, domyślnie `1` | `0` = tabela bez nagłówka i bez filtrów |
| `Table.tableColumns` | Lista kolumn | `list[TableColumn]` | Zwykle pusta — openpyxl wypełnia przy zapisie (2.5) |
| `Table.column_names` (property) | Nazwy kolumn | `list[str]` | Wygodny skrót odczytowy |
| `Table.autoFilter` | Filtr tabeli | `AutoFilter` | **Ustawiany automatycznie**, gdy jest nagłówek |
| `Table.sortState` | Stan sortowania | `SortState` | Rzadko potrzebne; **sprawdź w dokumentacji swojej wersji** |
| `Table.totalsRowCount` / `totalsRowShown` | Wiersz podsumowania | `int` / `bool` | Obsługa w openpyxl **ograniczona** — licz sam w Pythonie |
| `Table.tableStyleInfo` | Styl tabeli | `TableStyleInfo` | Bez tego tabela jest bezstylowa |
| `Table.id` | Identyfikator tabeli | `int` | **Zarządzany przez openpyxl przy zapisie** — nie ustawiaj ręcznie |
| `Table.path` (property) | Ścieżka w archiwum | np. `/xl/tables/table1.xml` | Do diagnostyki pakietu |
| `Table._initialise_columns()` | Tworzy listę kolumn | — | **Półpubliczne** (podkreślnik). Nazwy tymczasowe `Column{N}` |
| `TableStyleInfo(name, ...)` | Styl wizualny tabeli | `name`, `showFirstColumn`, `showLastColumn`, `showRowStripes`, `showColumnStripes` | Poprawne nazwy: `TableStyleLight1..21`, `TableStyleMedium1..28`, `TableStyleDark1..11` |
| `TableColumn(id, name)` | Kolumna tabeli | `id` — **bezwzględny** numer kolumny Excela; `name` — nazwa kolumny | Zawiera też `totalsRowFunction`, `calculatedColumnFormula` |
| `ws.add_table(tabela)` | Dodaje tabelę do arkusza | obiekt `Table` | Sprawdza duplikat nazwy (bez wielkości liter!) w całym skoroszycie |
| `ws.tables` | Kolekcja tabel arkusza | `TableList` (dziedziczy po `dict`) | Klucz = nazwa tabeli |
| `ws.tables[nazwa]` | Tabela po nazwie | — | **Nie działa po zakresie!** (2.8) |
| `ws.tables.get(table_range=...)` | Tabela po zakresie | `name=` lub `table_range=` | **Uwaga:** `get("A1:D6")` szuka po NAZWIE |
| `ws.tables.values()` | Iteracja po obiektach | — | To, czego zwykle chcesz |
| `ws.tables.items()` | Pary `(nazwa, ref)` | — | **Zwraca napisy ref, nie obiekty!** (2.8) |
| `del ws.tables[nazwa]` | Usuwa definicję tabeli | — | **Dane w komórkach zostają** |
| `len(ws.tables)` | Liczba tabel | — | |
| `wb._duplicate_name(nazwa)` | Sprawdza duplikat | — | Metoda „półpubliczna"; sprawdza tabele **i** defined names, bez wielkości liter |
| `range_boundaries(zakres)` | Rozbija zakres na liczby | `"A1:D6"` → `(1, 1, 4, 6)` | Z `openpyxl.utils.cell` — kluczowe do geometrii |
| `get_column_letter(n)` | Numer kolumny → litera | `4` → `"D"` | Do budowania zakresów |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — utwórz tabelę z filtrem i stylem.**

Napisz `examples/12_cw1.py`, który tworzy `output/12_cw1.xlsx`:

1. Arkusz o nazwie `Produkty` z nagłówkiem: `Nazwa`, `Cena netto`, `VAT`, `Cena brutto` (wszystkie napisy).
2. Osiem wierszy danych wygenerowanych w pętli (ceny losowe, `seed` ustalony).
3. Kolumna `VAT` i `Cena brutto` jako **formuły** (`=C2*0.23`, `=B2+C2`), z formatem `#,##0.00 "zł"`.
4. **Tabela** o nazwie `Produkty2026` obejmująca nagłówek i wszystkie dane, ze stylem `TableStyleMedium9` i paskami na wierszach.
5. Tabelę dodaj przez `ws.add_table`.
6. Zamrożenie `"A2"` i wyłączone linie siatki (moduł 11).

Następnie dodaj funkcję, która **wczytuje plik ponownie** i wypisuje:

- liczbę tabel w arkuszu,
- `displayName`, `ref`, `headerRowCount` każdej tabeli,
- `column_names`,
- czy `autoFilter` jest ustawiony,
- nazwę stylu i wartość `showRowStripes`.

W pliku `output/12_cw1_wnioski.md` odpowiedz:

1. **Skąd openpyxl wziął nazwy kolumn tabeli?** Nie nadałeś ich w obiekcie `Table`. W którym momencie i z jakiego miejsca zostały odczytane? (Odsyłam do 2.5.)
2. **Dlaczego `column_names` po odczycie pliku zwraca dokładnie te same napisy co nagłówki w arkuszu?** Co by się stało, gdybyś w nagłówku wpisał liczbę zamiast napisu?
3. **Czym różni się `ws.tables.items()` od `ws.tables.values()`?** Wypisz oba wyniki i wyjaśnij, dlaczego pętla `for nazwa, tabela in ws.tables.items(): print(tabela.ref)` zakończy się błędem.
4. **Czy `TableStyleInfo` działa, jeśli nie dodasz tabeli przez `ws.add_table`?** Sprawdź eksperymentalnie (usuń wywołanie `add_table`, zapisz, otwórz) i opisz wynik.
5. **Dlaczego w nazwie tabeli nie ma spacji ani polskich znaków?** Odsyłam do 2.3 — napisz dwa konkretne zagrożenia.

**Zadanie 2 — rozpoznaj różnicę tabela kontra zakres.**

Napisz `examples/12_cw2.py`, który tworzy **dwa arkusze w jednym skoroszycie**:

- Arkusz `Zakres`: dane w `A1:D6` (nagłówek + 5 wierszy) z **autofiltrem** (`ws.auto_filter.ref`), bez tabeli.
- Arkusz `Tabela`: te same dane jako **tabela** o nazwie `TabelaZadan`.

Następnie:

1. W obu arkuszach dopisz 3 nowe wiersze przez `ws.append`.
2. Zapisz plik.
3. Wczytaj plik i wypisz dla obu arkuszy: zakres autofiltra (`ws.auto_filter.ref`) oraz zakres tabeli (`ws.tables[...].ref`).

W pliku `output/12_cw2_wnioski.md`:

1. **Czy w arkuszu `Zakres` nowe wiersze znalazły się w autofiltrze?** Dlaczego?
2. **Czy w arkuszu `Tabela` nowe wiersze znalazły się w tabeli?** Dlaczego?
3. **Otwórz plik w Excelu.** W arkuszu `Tabela` wpisz ręcznie coś w wiersz bezpośrednio pod tabelą — czy Excel rozszerzył tabelę sam? **Dlaczego program tego nie zrobił, a człowiek w Excelu tak?** (To pytanie o 2.12 — napisz odpowiedź w 3–4 zdaniach.)
4. **Która konstrukcja wymaga mniej pracy przy programowym dopisywaniu wierszy: zakres czy tabela?** Uzasadnij.
5. **W jakim scenariuszu wybrałbyś zakres, a w jakim tabelę?** Podaj po dwa przykłady.

### 🟡 Warsztat

**Zadanie 3 — funkcja „dopisz wiersze do tabeli".**

To zadanie zamienia wiedzę z 2.12 w **wielokrotnego użytku narzędzie**.

**Część A.** Napisz `examples/12_cw3.py` z klasą `TabelaDanych`, która opakowuje tabelę i pilnuje spójności `ref`:

```python
"""Bezpieczne opakowanie tabeli: dopisywanie wierszy bez zapominania o ref."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


class TabelaDanych:
    """Opakowuje tabele Excela i dba o spojnosc zakresu z danymi.

    Korzystajac z tej klasy, NIE dopisujesz wierszy przez ws.append bezposrednio
    - zawsze przez .dopisz(), zeby zakres tabeli byl aktualizowany automatycznie.
    """

    def __init__(self, ws, nazwa: str, naglowki: list[str], styl: str = "TableStyleMedium9"):
        self.ws = ws
        self.nazwa = nazwa
        self._naglowki = list(naglowki)
        self._wiersz_start = ws.max_row + 1      # tabela zaczyna sie tu
        ws.append(self._naglowki)
        self._tabela = Table(displayName=nazwa, ref=self._aktualny_ref())
        self._tabela.tableStyleInfo = TableStyleInfo(name=styl, showRowStripes=True)
        ws.add_table(self._tabela)

    def _aktualny_ref(self) -> str:
        ostatnia = get_column_letter(len(self._naglowki))
        return f"A{self._wiersz_start}:{ostatnia}{self.ws.max_row}"

    def dopisz(self, wiersz: list) -> None:
        """Dopisuje wiersz danych i ROZSZERZA zakres tabeli."""
        if len(wiersz) != len(self._naglowki):
            raise ValueError(
                f"Wiersz ma {len(wiersz)} wartosci, a tabela {len(self._naglowki)} kolumn"
            )
        self.ws.append(wiersz)
        self._tabela.ref = self._aktualny_ref()    # TU jest sedno!

    @property
    def liczba_wierszy_danych(self) -> int:
        return self.ws.max_row - self._wiersz_start
```

**Wymagania:**

1. `TabelaDanych.__init__` tworzy nagłówek, tabelę i pilnuje, żeby zakres był poprawny.
2. `dopisz(wiersz)` dopisuje dane **i rozszerza `ref`** — nie da się zapomnieć.
3. `dopisz` waliduje liczbę kolumn (bo tabela z niepasującą liczbą kolumn to błąd).
4. Dodaj metodę `ustaw_format_kolumny(nazwa_kolumny, kod_formatu)`, która znajduje numer kolumny po nazwie i nadaje format wszystkim wierszom danych (moduł 10). **To użyteczny skrót** — spróbuj go napisać, korzystając z `self._naglowki.index(nazwa) + 1`.

**Część B.** Użyj `TabelaDanych`, żeby zbudować plik `output/12_cw3.xlsx` z **jedną** tabelą, do której dopiszesz 500 wierszy w pętli.

**Część C.** Dodaj funkcję `weryfikuj(sciezka)`, która sprawdza **warunek niezmienny**: liczba wierszy danych w tabeli **musi być równa** liczbie dopisanych wierszy. Zwróć `True`/`False` i wypisz wynik.

W pliku `output/12_cw3_wnioski.md`:

1. **Dlaczego opakowanie tabeli w klasę jest lepsze niż rozrzucenie `ws.append` i ręcznego `tabela.ref = ...` po całym skrypcie?** Podaj trzy argumenty (jeden o błędach, jeden o testowaniu, jeden o czytelności kodu).
2. **Twoja klasa ma prywatne pole `_tabela`.** Co by się stało, gdyby ktoś z zewnątrz zmienił `_tabela.ref` na wartość niepasującą do danych? Czy Twoja klasa by to wykryła? Jak byś to zabezpieczył?
3. **Metoda `dopisz` waliduje liczbę kolumn.** Wymyśl jeszcze dwa warunki, które warto walidować przy dopisywaniu wiersza do tabeli, i uzasadnij.
4. **Gdybyś chciał dopisywać wiersze z bazy danych (kursor zwracający tysiące wierszy), czy Twoja klasa by się sprawdziła?** Co byś zmienił? (To zapowiedź modułu 19.)
5. **To zadanie jest prototypem wzorca.** Który wzorzec z modułu 23 lub 24 realizuje ta klasa? Uzasadnij w dwóch zdaniach.

### 🔴 Wyzwanie

**Zadanie 4 — audyt tabel w skoroszycie.**

Napisz `examples/12_cw4.py` implementujący **audytor tabel**, który znajduje wszystkie dziewięć klas problemów omówionych w 2.9 i podaje sposób naprawy.

```python
"""Audytor tabel: wykrywa problemy, ktore openpyxl przepuszcza."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


@dataclass
class Problem:
    waga: str          # "BLAD" albo "OSTRZEZENIE"
    arkusz: str
    tabela: str | None
    opis: str          # co jest zle
    naprawa: str       # co zrobic


def audytuj_tabele(sciezka: Path) -> list[Problem]:
    """Sprawdza wszystkie tabele w pliku pod katem dziewieciu klas problemow.

    Wykrywa:
      1. Nazwa tabeli ze spacja (openpyxl by nie zapisal, ale plik mogl przyjsc z zewnatrz).
      2. Nazwa wygladajaca jak adres komorki (np. Tbl1, ABC123).
      3. Nazwa z polskimi znakami lub myslnikiem.
      4. Nazwy tabel powtarzajace sie w skoroszycie (bez wielkosci liter).
      5. Tabele o nachodzacych sie zakresach.
      6. Tabele z naglowkiem, ale bez wierszy danych (ref = sam naglowek).
      7. Zduplikowane nazwy kolumn (bez wielkosci liter).
      8. Naglowki kolumn, ktore nie sa napisami.
      9. Tabele w scalonych zakresach.
    """
    ...


def raport_tekstowy(problemy: list[Problem]) -> str:
    """Formatuje liste problemow jako czytelny raport."""
    ...
```

**Wymagania funkcjonalne:**

1. **Dziewięć testów**, po jednym na każdą klasę problemu. Dla każdej napisz **osobny plik testowy** (z celowo wstrzykniętym problemem) i sprawdź, że audyt go wykrywa:

   - nazwa ze spacją — **trudne**, bo `Table(displayName="Zła Nazwa")` rzuci `ValueError` przy tworzeniu. **Musisz utworzyć taki plik, manipulując XML-em archiwum** (rozpakuj, popraw `displayName` w `table1.xml`, spakuj). Jeśli to zbyt trudne, **udokumentuj, dlaczego nie da się tego zrobić przez openpyxl** i oznacz test jako pominięty. To pouczające samo w sobie.
   - nazwa jak adres: `Tbl1` (utwórz przez openpyxl — przejdzie!).
   - nazwa z polskim znakiem: `Sprzedaż` (przejdzie przez openpyxl).
   - duplikat nazwy: **też trudne** — próba `add_table` z duplikatem rzuci `ValueError`. Utwórz dwie tabele w **różnych arkuszach** o **różnych** nazwach, a potem... nie, też się nie da. **Dokumentuj ograniczenie.**
   - nakładające się zakresy (jak w Przykładzie 3).
   - tabela bez danych: `ref="A1:C1"`.
   - zduplikowane kolumny: nagłówki `["Kwota", "Kwota"]`.
   - nagłówek niebędący napisem: nagłówek z liczbą `2026`.
   - tabela w scalonym zakresie: `ws.merge_cells("A1:B1")` + tabela.

2. **`audytuj_tabele` musi używać `range_boundaries`** do wykrywania nakładania i tabel bez danych. Do sprawdzenia „wygląda jak adres komórki" użyj logiki z 2.3 (możesz skopiować `czy_adres_komorki`).

3. **Sprawdzenie nagłówków niebędących napisami** wymaga odczytania **pierwszego wiersza zakresu `ref`** i sprawdzenia typu każdej komórki przez `isinstance(cell.value, str)`. Pamiętaj, że `ref` obejmuje nagłówek.

4. **Sprawdzenie tabeli w scalonym zakresie** wymaga porównania zakresu tabeli z `ws.merged_cells.ranges` — użyj tego samego testu nakładania co do tabel.

5. **`raport_tekstowy`** tworzy czytelny raport:

```
=== AUDYT TABEL: 12_cw4.xlsx ===

Arkusz: Nakladanie  (tabel: 2)
  [BLAD] Tabela_A <-> Tabela_B
     opis:    zakresy A1:D10 i C1:F10 nachodzą na siebie (wspólne: C1:D10)
     naprawa: zostaw jedną tabelę albo rozdziel zakresy

Arkusz: Pusta  (tabel: 1)
  [BLAD] PustaTabela (A1:C1)
     opis:    tabela ma nagłówek, ale zero wierszy danych
     naprawa: nie dodawaj tabeli, gdy ws.max_row == 1 (warunek w kodzie)

PODSUMOWANIE: 2 błędy, 0 ostrzeżeń
```

6. **Zapisz raport** do `output/12_cw4_raport.txt` i wypisz `wynik: X/9` (ile z dziewięciu klas problemów udało się pokryć testem).

W pliku `output/12_cw4_wnioski.md`:

1. **Które z dziewięciu problemów dało się wstrzyknąć przez openpyxl, a których nie?** Zrób tabelę: problem | da się przez openpyxl | dlaczego. To pytanie o granicę między „co openpyxl przepuszcza" a „co blokuje".
2. **Dlaczego to, że openpyxl blokuje duplikat nazwy przy `add_table`, jest jednocześnie dobre i złe?** (Dobre — chroni Cię. Złe — blokuje **wczytanie** cudzego pliku, który ma duplikat.) Opisz scenariusz, w którym to boli.
3. **Jak wykryłeś tabelę w scalonym zakresie?** Wyjaśnij, dlaczego użyłeś tego samego testu nakładania co do tabel — i dlaczego to nie jest przypadek.
4. **Twój audyt sprawdza nagłówki przez `isinstance(cell.value, str)`.** Co zrobiłbyś z nagłówkiem, który jest `None` (pusta komórka nagłówka)? Czy to błąd, czy ostrzeżenie? Uzasadnij.
5. **Wyobraź sobie, że ten audyt ma być uruchamiany automatycznie po każdym wygenerowaniu raportu.** Gdzie go wstawisz — przed zapisem czy po? Jakie to ma konsekwencje dla obsługi błędów? To zapowiedź modułu 20 (testy) i 26 (pipeline).

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Audytor tabel - kluczowe fragmenty."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


WZORZEC_ADRESU = re.compile(r"^[A-Za-z]{1,3}\d+$")
MAKS_KOLUMNA = 16384


def _kolumna_na_numer(litery: str) -> int:
    numer = 0
    for znak in litery.upper():
        numer = numer * 26 + (ord(znak) - ord("A") + 1)
    return numer


def _czy_adres_komorki(nazwa: str) -> bool:
    if not WZORZEC_ADRESU.match(nazwa):
        return False
    litery = re.match(r"^[A-Za-z]+", nazwa).group(0)
    return _kolumna_na_numer(litery) <= MAKS_KOLUMNA


def _czy_ascii(nazwa: str) -> bool:
    """Czy nazwa sklada sie wylacznie ze znakow ASCII?"""
    try:
        nazwa.encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def _zakresy_nachodza(zakres_a: str, zakres_b: str) -> str | None:
    """Zwraca zakres wspolny, jesli nachodza; inaczej None."""
    a = range_boundaries(zakres_a)
    b = range_boundaries(zakres_b)
    min_col = max(a[0], b[0])
    min_row = max(a[1], b[1])
    max_col = min(a[2], b[2])
    max_row = min(a[3], b[3])
    if min_col > max_col or min_row > max_row:
        return None
    return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


def audytuj_tabele(sciezka: Path) -> list[Problem]:
    wb = load_workbook(sciezka)
    problemy: list[Problem] = []
    nazwy_widziane: dict[str, str] = {}     # nazwa.casefold() -> "arkusz/tabela"
    try:
        for ws in wb.worksheets:
            tabele = [(n, ws.tables[n]) for n in ws.tables]

            # --- 6. Tabela bez danych -------------------------------
            for nazwa, tabela in tabele:
                _, min_row, _, max_row = range_boundaries(tabela.ref)
                if tabela.headerRowCount and max_row <= min_row:
                    problemy.append(Problem(
                        "BLAD", ws.title, nazwa,
                        f"tabela ma naglowek, ale zero wierszy danych ({tabela.ref})",
                        "nie dodawaj tabeli, gdy ws.max_row == 1",
                    ))

            # --- 8. Naglowki niebędące napisami ----------------------
            for nazwa, tabela in tabele:
                if not tabela.headerRowCount:
                    continue
                min_col, min_row, max_col, _ = range_boundaries(tabela.ref)
                for kolumna in range(min_col, max_col + 1):
                    wartosc = ws.cell(row=min_row, column=kolumna).value
                    if not isinstance(wartosc, str):
                        problemy.append(Problem(
                            "BLAD", ws.title, nazwa,
                            f"naglowek w kolumnie {get_column_letter(kolumna)} "
                            f"nie jest napisem (jest {type(wartosc).__name__})",
                            "rzutuj naglowki na str() przed utworzeniem tabeli",
                        ))

            # --- 7. Zduplikowane nazwy kolumn -----------------------
            for nazwa, tabela in tabele:
                widziane_kolumny: dict[str, int] = {}
                for kolumna in tabela.tableColumns:
                    klucz = (kolumna.name or "").casefold()
                    if klucz in widziane_kolumny:
                        problemy.append(Problem(
                            "BLAD", ws.title, nazwa,
                            f"zduplikowana nazwa kolumny: {kolumna.name!r} "
                            f"(Excel odrzuci plik)",
                            "nadaj unikalne nazwy kolumn przed dodaniem tabeli",
                        ))
                    widziane_kolumny[klucz] = 1

            # --- 1/2/3/4. Nazwy tabel -------------------------------
            for nazwa, tabela in tabele:
                # 1. spacja - openpyxl by nie zapisal, ale plik moze przyjsc z zewnatrz
                if " " in (tabela.displayName or ""):
                    problemy.append(Problem(
                        "BLAD", ws.title, nazwa,
                        "nazwa tabeli zawiera spacje",
                        "zmien nazwe - Excel nie akceptuje spacji w nazwach tabel",
                    ))
                # 2. wyglada jak adres
                if _czy_adres_komorki(tabela.displayName or ""):
                    problemy.append(Problem(
                        "BLAD", ws.title, nazwa,
                        f"nazwa {tabela.displayName!r} wyglada jak adres komorki",
                        "dodaj prefiks, np. tbl_ (Excel odrzuci taka nazwe)",
                    ))
                # 3. polskie znaki / myslnik
                if not _czy_ascii(tabela.displayName or "") or "-" in (tabela.displayName or ""):
                    problemy.append(Problem(
                        "OSTRZEZENIE", ws.title, nazwa,
                        f"nazwa {tabela.displayName!r} zawiera znaki ryzykowne (polskie/myslnik)",
                        "uzyj wylacznie liter ASCII, cyfr i podkreslnika",
                    ))
                # 4. duplikat w skoroszycie (bez wielkosci liter)
                klucz = (tabela.displayName or "").casefold()
                if klucz in nazwy_widziane:
                    problemy.append(Problem(
                        "BLAD", ws.title, nazwa,
                        f"nazwa powtarza sie z {nazwy_widziane[klucz]}",
                        "nadaj unikalna nazwe w calym skoroszycie (bez wielkosci liter)",
                    ))
                else:
                    nazwy_widziane[klucz] = f"{ws.title}/{nazwa}"

            # --- 5. Nakładanie się tabel ----------------------------
            for i in range(len(tabele)):
                for j in range(i + 1, len(tabele)):
                    nazwa_a, tab_a = tabele[i]
                    nazwa_b, tab_b = tabele[j]
                    wspolny = _zakresy_nachodza(tab_a.ref, tab_b.ref)
                    if wspolny:
                        problemy.append(Problem(
                            "BLAD", ws.title, f"{nazwa_a} <-> {nazwa_b}",
                            f"zakresy {tab_a.ref} i {tab_b.ref} nachodza "
                            f"(wspolne: {wspolny})",
                            "zostaw jedna tabele albo rozdziel zakresy",
                        ))

            # --- 9. Tabela w scalonym zakresie ----------------------
            for nazwa, tabela in tabele:
                for scalony in ws.merged_cells.ranges:
                    if _zakresy_nachodza(tabela.ref, str(scalony)):
                        problemy.append(Problem(
                            "BLAD", ws.title, nazwa,
                            f"tabela nachodzi na scalony zakres {scalony}",
                            "nie umieszczaj tabeli w scalonym zakresie",
                        ))
    finally:
        wb.close()
    return problemy
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **Ten sam test nakładania dla tabel i dla scalonych zakresów.** To nie przypadek, a **reuse abstrakcji**: dwa prostokąty nachodzą na siebie albo nie — niezależnie od tego, czy jednym z nich jest tabela, drugim scalony zakres, czy czymkolwiek innym. Wydzielenie `_zakresy_nachodza(zakres_a, zakres_b)` do jednej funkcji obsługuje **pięć** z dziewięciu testów (nakładanie, scalony zakres) i jest gotowym kandydatem do testów jednostkowych. To wzorzec: **gdy ta sama logika przyda się w wielu miejscach, wydziel ją jako funkcję czystą** — taką, która nie dotyka pliku. Zapowiedź modułu 20 (testy) i 25 (Specification).

- **Problem 1 (spacja w nazwie) jest nieosiągalny przez `Table`** — bo deskryptor `TableNameDescriptor` **rzuca `ValueError` przy tworzeniu obiektu**. To znaczy, że takiego pliku **nie da się utworzyć przez openpyxl** — ale **da się go wczytać**, jeśli przyjdzie z zewnątrz (Excel jest bardziej tolerancyjny? nie — Excel też nie akceptuje spacji; ale narzędzie trzecie mogło wygenerować taką nazwę w XML). Dlatego audytor ma sens: **sprawdza pliki przychodzące, nie tylko wychodzące.** Test pomijamy, ale wpisujemy problem do raportu.

- **Problem 4 (duplikat w skoroszycie)** — z tego samego powodu: `add_table` blokuje utworzenie duplikatu. Ale **przy wczytaniu cudzego pliku z duplikatem `load_workbook` sam rzuci `ValueError`** — więc audytor i tak nie dobiegnie do końca. To znaczy, że problem duplikatu trzeba rozwiązać **innym narzędziem** (naprawa archiwum) — i to jest pouczające: **openpyxl broni się przed duplikatem tak skutecznie, że sam nie potrafi wczytać pliku, który go ma.** Zapisz to w rozwiązaniu jako świadome ograniczenie audytu.

- **`_czy_adres_komorki` liczy numer kolumny** — i porównuje z maksymalną kolumną Excela (16384 = `XFD`). Bez tego warunku każda nazwa w stylu `Rok2026` byłaby fałszywie oznaczona jako adres (bo `Rok` nie jest poprawną kolumną). To pokazuje subtelność: **walidacja musi znać granice domeny, którą sprawdza.** Analogia: nie odrzucasz tablic rejestracyjnych dlatego, że wyglądają jak słowa — sprawdzasz, czy mają właściwy format i czy nie kolidują z czymś innym.

- **Nagłówek `None` (pusty) klasyfikujemy jako `BLAD`** — bo Excel odrzuca tabelę z pustym nagłówkiem kolumny. To jest ta sama kategoria co nagłówek liczbowy: „nagłówki muszą być napisami". W audycie rozróżniamy tylko `BLAD`/`OSTRZEZENIE`; dla pustego nagłówka `BLAD` jest właściwe, bo plik nie otworzy się poprawnie.

- **Rozróżnienie `BLAD`/`OSTRZEZENIE`** jest istotne: `BLAD` = Excel odrzuci plik (nazwa jak adres, duplikat kolumn, nakładanie). `OSTRZEZENIE` = plik się otworzy, ale jest ryzyko (polskie znaki). **Nigdy nie oznaczaj ostrzeżenia jako błędu** — audyt, który krzyczy o rzeczach nieszkodliwych, przestaje być czytany. To dokładnie ta sama lekcja co w module 11 przy `audyt_druku`: „audyt, który nie potrafi odróżnić znanego błędu od fałszywego alarmu, jest bezwartościowy".

**Dlaczego to wykrywamy u siebie, a nie polegamy na Excelu:** bo gdy Excel powie „plik wymaga naprawy", użytkownik zobaczy komunikat, a Ty nie. Wykrycie konfliktu **przed wysłaniem pliku** to różnica między „działa" a „klient dzwoni". Ten audyt to prototyp, który w module 26 stanie się częścią pipeline'u: **wygeneruj → zaudytuj → dopiero potem wyślij.**

</details>

## 6. Typowe błędy i pułapki

**1. „`Table(displayName="Moja Tabela", ...)` rzuca `ValueError`" (objaw) → **nazwa tabeli zawiera spację.** Deskryptor nazwy tabeli w openpyxl sprawdza to jako **jedyną** rzecz i rzuca `ValueError("Table names cannot have spaces")` już przy tworzeniu obiektu (2.3) (przyczyna) → **usuń spacje**, zastąp podkreślnikami: `"Moja_Tabela"` albo `bezpieczna_nazwa_tabeli("Moja Tabela")` → `"tbl_Moja_Tabela"`. To błąd „dobrego rodzaju": głośny i natychmiastowy. Ale uwaga — to tylko jedna z ośmiu reguł; pozostałe openpyxl przepuści (naprawa).**

**2. „Openpyxl zapisał plik bez błędu, ale Excel mówi «plik wymaga naprawy» i usuwa tabelę" (objaw) → **nazwa tabeli wygląda jak adres komórki** (np. `Tbl1`, `ABC123`). Excel traktuje nazwy tabel jak nazwy definiowane i **odrzuca** te, które wyglądają jak adres komórki (`TBL1` jest poprawnym adresem: kolumna `TBL` = 13584 ≤ 16384) (przyczyna) → **dodaj prefiks**, np. `tbl_`: `tbl_Tbl1`, `tbl_ABC123`, `tbl_Sprzedaz`. To najczęstsza przyczyna „repair prompt" w plikach generowanych przez openpyxl — obok duplikatów kolumn. Zapamiętaj: **`Tabela1` jest bezpieczne, `Tbl1` nie** (naprawa).**

**3. „`load_workbook` rzuca `ValueError: Table with name ... already exists`, mimo że plik otwiera się w Excelu i LibreOffice" (objaw) → **w pliku są dwie tabele o tej samej nazwie** (albo tabela koliduje z nazwą definiowaną). Openpyxl sprawdza duplikaty nazw **przy wczytywaniu** i **rzuca wyjątek** zamiast załadować plik (2.10). Nazwy są porównywane **bez rozróżniania wielkości liter**, więc `Dane` i `DANE` to duplikat (przyczyna) → (a) **po swojej stronie nigdy nie twórz duplikatów** — buduj nazwy z prefiksem i licznikiem (`tbl_Mazowieckie_1`, `tbl_Mazowieckie_2`), (b) jeśli plik przyszedł z zewnątrz i ma duplikat, napraw go **innym narzędziem** — rozpakuj archiwum i usuń jedną definicję z `xl/tables/tableN.xml` plus jej relację (moduł 18). To bolesna lekcja: **openpyxl tak skutecznie broni się przed duplikatem, że sam nie potrafi wczytać pliku, który go ma** (naprawa).**

**4. „Tabela jest w pliku, ale obejmuje tylko nagłówek — nowe wiersze są poza nią" (objaw) → **zapomniałeś rozszerzyć `Table.ref` przy programowym dopisywaniu.** Openpyxl nie wie, że nowe wiersze należą do tabeli — rozszerzanie tabeli to funkcja **interfejsu Excela**, nie formatu pliku (2.12, analogia 1.1). Zakres `ref` pozostaje bez zmian (przyczyna) → **po dopisaniu wierszy zaktualizuj zakres ręcznie**: `tabela.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"`. Najlepiej schowaj to w funkcji albo klasie „dopisz wiersz do tabeli" (ćwiczenie 🟡) — wtedy nie da się zapomnieć. Weryfikacja: `tabela.ref` po odczycie musi obejmować ostatni wiersz `ws.max_row` (naprawa).**

**5. „Tabela ma nagłówek, ale Excel odrzuca plik jako uszkodzony (narzędzie generuje raport z pustymi danymi)" (objaw) → **`ref` tabeli obejmuje nagłówek, ale zero wierszy danych** (np. `A1:C1` albo `A1:C2` z pustymi danymi). Excel wymaga tabeli z nagłówkiem **plus co najmniej jednym wierszem danych**. To ekstremalnie częsty błąd w generatorach raportów: gdy dane wejściowe są puste (zero zamówień w miesiącu), tabela powstaje z samym nagłówkiem (przyczyna) → **zabezpiecz się warunkiem**: `if ws.max_row > 1: ws.add_table(tabela)`. Jeśli danych może nie być, **dodaj choć jeden wiersz „(brak danych)"** — to jest wybór biznesowy, ale ratuje plik. To ta sama rodzina problemów co puste strony w `print_area` z modułu 11: **granice zakresu muszą być wyliczone z realnych danych, nie z założenia** (naprawa).**

**6. „Excel mówi «usunięto tabelę, bo nagłówki się nie zgadzają»" / „tabela ma kolumny `Column5`, `Column6`" (objaw) → dwie powiązane przyczyny. Pierwsza: **komórki nagłówka nie są napisami** (liczba, data, `None`) — openpyxl ostrzega *„File may not be readable: column headings must be strings"*, ale zapisuje, a Excel odrzuca. Druga: **w trybie `write_only`** openpyxl nie ma dostępu do komórek, więc **nie odczyta nagłówków** i zostawi tymczasowe nazwy `Column{N}` (gdzie `N` to **bezwzględny** numer kolumny!) (2.5) (przyczyna) → (a) **rzutuj nagłówki na napisy** przed `append`, (b) w `write_only` **nadaj nazwy kolumn ręcznie**: `tabela._initialise_columns()` i pętla ustawiająca `column.name`, albo utwórz tabelę bez nagłówka (`headerRowCount = False`). Pamiętaj, że nazwy kolumn **pochodzą z komórek arkusza**, nie z Twojego kodu — to najważniejszy wniosek z 2.5 (naprawa).**

**7. „Dwie tabele na tym samym obszarze — openpyxl nie protestuje, ale Excel odrzuca plik" (objaw) → **openpyxl nie sprawdza nakładania się zakresów** (2.9B). Sprawdza tylko duplikaty nazw. Dwie tabele na przecinających się komórkach są dla openpyxl dwoma poprawnymi obiektami (przyczyna) → **napisz własny detektor nakładania** (Przykład 3 i ćwiczenie 🔴) i uruchamiaj go przed zapisem, jeśli tylko w Twoim kodzie mogłyby powstać nachodzące tabele (generowanie w pętli, dziedziczenie szablonu). Test: dla każdej pary tabel sprawdź, czy prostokąty nachodzą (naprawa).**

**8. „Kolumna o nazwie `Kwota` jest zdublowana i Excel odrzuca plik" (objaw) → **zduplikowane nazwy kolumn, bez rozróżniania wielkości liter.** Jeśli Twoje nagłówki mają dwa razy `Kwota` albo `Kwota` i `KWOTA`, Excel uzna plik za uszkodzony. Dotyczy to zwłaszcza danych z pandas, gdzie duplikaty kolumn są możliwe (przyczyna) → **przed utworzeniem tabeli ujednolicz nazwy kolumn** — dopisuj licznik do duplikatów: `Kwota`, `Kwota2`, `Kwota3`. To ta sama zasada co przy nazwach tabel: **unikalność bez wielkości liter.** Napisz funkcję `unikalne_naglowki(list[str])` i używaj jej zawsze (naprawa).**

**9. „W trybie `read_only` `ws.tables` nie działa / rzuca błąd" (objaw) → **`ReadOnlyWorksheet` nie ma kolekcji tabel.** Openpyxl w trybie `read_only` (moduł 19) czyta plik strumieniowo, komórka po komórce, i **nie wczytuje definicji tabel** — to znane ograniczenie (zgłoszenie #843 w openpyxl), które nie zostało rozwiązane w 3.1.x (przyczyna) → **jeśli musisz czytać tabele, użyj trybu normalnego** (`load_workbook` bez `read_only`). Jeśli plik jest ogromny i musisz użyć `read_only`, **przeczytaj definicję tabeli osobno — rozpakowując archiwum i parsując `xl/tables/tableN.xml`** (moduł 02 + 18). To pokazuje, że tryby wydajnościowe mają cenę: **`read_only` kupuje pamięć za cenę funkcji** (naprawa).**

**10. „Nie mogę dodać tabeli — `add_table` wyrzuca `ValueError` w `write_only`" (objaw) → ⚠️ **To nie jest błąd!** W `write_only` `add_table` **tylko ostrzega**: *„In write-only mode you must add table columns manually"*. Jeśli widzisz `ValueError`, to **prawie na pewno duplikat nazwy**, nie tryb `write_only` (2.9A) (przyczyna) → (a) sprawdź, czy nazwa tabeli nie powtarza się w skoroszycie (i to bez wielkości liter), (b) jeśli to naprawdę `write_only` z ostrzeżeniem — **nadaj nazwy kolumn ręcznie** (`_initialise_columns` + pętla). Nie myl `ValueError` z ostrzeżeniem: pierwsze przerywa program, drugie idzie na `stderr` (naprawa).**

**11. „Chcę, żeby formuła `=SUM(Tabela1[Kwota])` uwzględniała nowe wiersze, ale nie uwzględnia" (objaw) → **struktura odwołania jest poprawna, ale `ref` tabeli nie obejmuje nowych wierszy** (punkt 4) albo **formuła jest wpisana poza zakresem tabeli `ref`** — strukturalne odwołanie do wiersza bieżącego (`[@...]`) musi być **w wierszu należącym do tabeli** (2.6) (przyczyna) → (a) rozszerz `ref` przy każdym dopisywaniu wierszy, (b) upewnij się, że formuła jest w wierszu wewnątrz zakresu tabeli. I pamiętaj: **openpyxl nie oblicza formuł** (moduł 06) — nie sprawdzisz wyniku w Pythonie, tylko w Excelu. Weryfikacja: otwórz plik, kliknij w tabelę i sprawdź, czy podsumowanie uwzględnia ostatni wiersz (naprawa).**

**12. „Tabela nadpisuje moje ręczne formatowanie komórek / kolor się nie pojawia" (objaw) → **konflikt między stylem tabeli a formatowaniem bezpośrednim komórek.** Openpyxl zapisuje i jedno, i drugie, ale Excel ma własną kolejność pierwszeństwa (formatowanie bezpośrednie, styl tabeli, formatowanie warunkowe) (przyczyna) → **wybierz jedno podejście.** Jeśli chcesz wygląd tabeli — użyj `TableStyleInfo` i nie maluj komórek ręcznie. Jeśli chcesz własne formatowanie — nie nadawaj `tableStyleInfo` (albo nadaj `TableStyleInfo` tylko z paskami, bez koloru). **Nie mieszaj i sprawdzaj wzrokowo** — bo priorytety w Excelu są złożone i wersjozależne (naprawa).**

**13. „Tabela przesuwa się / wygląda dziwnie po scaleniu komórek w jej obrębie" (objaw) → **tabela na scalonym zakresie albo scalenie wewnątrz tabeli.** Scalanie (moduł 11) usuwa wartości i zamienia komórki na `MergedCell`. Tabela w takim obszarze ma niespójny nagłówek i Excel się gubi (przyczyna) → **nigdy nie scalaj komórek wewnątrz tabeli** ani nie umieszczaj tabeli na scalonym zakresie. Jeśli raport wymaga scalonego nagłówka nad tabelą — umieść go **poza** zakresem tabeli `ref`, w osobnym wierszu powyżej (naprawa).**

**14. „`ws.print_area = "Tabela1"` nie działa / openpyxl ostrzega" (objaw) → **openpyxl nie potrafi rozwiązać dynamicznego obszaru wydruku opartego na nazwie tabeli.** Excel pozwala ustawić `print_area` na nazwę tabeli (obszar „sam się rozszerza"), ale to definicja dynamiczna — openpyxl jej **nie rozwiąże** i **emituje ostrzeżenie** przy próbie wczytania (przyczyna) → **użyj zakresu tabeli**: `ws.print_area = ws.tables["Tabela1"].ref`. Jeśli chcesz zachować dynamikę (obszar rośnie razem z tabelą), **aktualizuj `print_area` po każdym rozszerzeniu tabeli** — tak samo jak `ref` (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Tabela to obiekt, nie zakres.** Ma nazwę, zakres, listę kolumn i styl — i mieszka jako **osobna część** w archiwum (`xl/tables/tableN.xml`), połączona z arkuszem relacją i elementem `<tableParts>`. To dowód, że tabela jest **warstwą opisu**, nie warstwą danych: usunięcie definicji tabeli nie usuwa ani jednej komórki. Konsekwencja: `del ws.tables["X"]` usuwa tabelę, dane zostają.

2. **Nazwa tabeli jest kontraktem, a nie dekoracją — i dlatego ma surowe reguły.** Odwołania strukturalne (`=SUM(Sprzedaz2026[Kwota])`) działają po **tożsamości**, nie po położeniu, więc zmiana nazwy psuje wszystko, co się do niej odwoływało. Openpyxl sprawdza **tylko jedną** z ośmiu reguł (spacje) — resztę (`Tbl1` wyglądające jak adres, polskie znaki, myślniki) przepuszcza, a Excel odrzuca. **Reguła: buduj nazwy programowo, z prefiksem `tbl_`, przez jedną funkcję.**

3. **Nazwy kolumn tabeli pochodzą z komórek nagłówka arkusza — nie z obiektu `Table`.** Openpyxl odczytuje je **przy zapisie** (jeśli nie nadałeś ich jawnie) i ostrzega, gdy nagłówek nie jest napisem. Konsekwencja: nagłówki **muszą być napisami** (`str(2026)`, nie `2026`), muszą być **unikalne bez wielkości liter**, a w trybie `write_only` (bez dostępu do komórek) **trzeba je nadać ręcznie** — inaczej kolumny dostaną nazwy `Column{N}`.

4. **„Segregator sam się rozszerza" tylko wtedy, gdy kartkę wkłada człowiek.** Rozszerzanie tabeli to funkcja **interfejsu Excela**, nie formatu pliku — więc openpyxl tego nie robi. Po każdym programowym dopisaniu wierszy **musisz zaktualizować `Table.ref` ręcznie**. To najważniejsza praktyczna korekta analogii z 1.1 i jedna z najczęstszych przyczyn „dlaczego Excel nie widzi moich nowych wierszy". **Reguła: schowaj dopisywanie wierszy w funkcji, która zawsze rozszerza `ref`.**

5. **Openpyxl nie jest strażnikiem poprawności Twojego pliku — więc Ty musisz nim być.** Przepuści nakładające się tabele, tabele bez danych, duplikaty kolumn, nazwy wyglądające jak adresy. Excel odrzuci je dopiero przy otwarciu — **u odbiorcy, nie u Ciebie**. To ta sama kategoria co ostrzeżenia o utracie danych z modułu 03 i puste strony z modułu 11. **Reguła: napisz audytor tabel (ćwiczenie 🔴) i uruchamiaj go przed wysłaniem pliku.** Wygeneruj → zaudytuj → dopiero potem wyślij.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORT
# ==================================================================
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

# ==================================================================
# 2. NAJPROSTSZA TABELA (3 linie)
# ==================================================================
ws.append(["Miesiąc", "Sprzedaż", "Koszt"])     # naglowki jako NAPISY!
for wiersz in dane:
    ws.append(wiersz)

tabela = Table(displayName="Sprzedaz2026", ref=f"A1:C{ws.max_row}")
tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tabela)                            # ZAWSZE przez add_table!
# ⚠️ ref MUSI obejmowac naglowki
# ⚠️ tabela MUSI miec >= 1 wiersz danych (inaczej Excel odrzuci)

# ==================================================================
# 3. NAZWA TABELI - REGULY (2.3)
# ==================================================================
# ✅ Sprzedaz2026, tbl_Sprzedaz_2026, Dane_1
# ❌ Moja Tabela         (spacja -> ValueError od razu)
# ❌ Tbl1                (wyglada jak adres -> Excel odrzuca!)
# ❌ ABC123              (wyglada jak adres -> Excel odrzuca!)
# ❌ Tabela-1            (myslnik)
# ❌ Sprzedaż2026        (polski znak - ryzyko)
# ❌ 2026_dane           (zaczyna sie od cyfry)
# ⚠️ unikalna w CALYM skoroszycie, BEZ wielkosci liter
#    ("Dane" i "DANE" to TA SAMA nazwa!)

# ==================================================================
# 4. STYL TABELI (2.7)
# ==================================================================
styl = TableStyleInfo(
    name="TableStyleMedium9",   # poprawne: Light1-21, Medium1-28, Dark1-11
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,        # paski - najtansza poprawa czytelnosci
    showColumnStripes=False,
)
tabela.tableStyleInfo = styl
# ⚠️ styl dziala TYLKO gdy tabela jest dodana przez ws.add_table
# ⚠️ nie mieszaj stylu tabeli z recznym PatternFill bez zrozumienia priorytetow

# ==================================================================
# 5. NAZWY KOLUMN - skad sie biora (2.5)
# ==================================================================
# DOMYSLNIE: openpyxl czyta je z wiersza naglowka arkusza PRZY ZAPISIE
# ⚠️ naglowki MUSZA byc napisami (str), inaczej ostrzezenie + ryzyko

# RECZNIE (potrzebne w write_only):
tabela._initialise_columns()                # polpubliczne (podkreslnik!)
for kolumna, nazwa in zip(tabela.tableColumns, ["Produkt", "Rok", "Kwota"]):
    kolumna.name = nazwa
# ⚠️ tymczasowe nazwy to "Column{N}", gdzie N = BEZWZGLEDNY numer kolumny

# ==================================================================
# 6. TABELA BEZ NAGLOWKA (rzadko!)
# ==================================================================
tabela.headerRowCount = 0      # brak naglowka => BRAK FILTROW
# ⚠️ filtry sa automatyczne przy naglowku i NIE da sie ich pominac

# ==================================================================
# 7. DOPISYWANIE WIERSZY - ROZSZERZAJ ref! (2.12)
# ==================================================================
ws.append(["Nowy", 100, 50])

# ...i OBOWIAZKOWO rozszerz zakres tabeli:
tabela = ws.tables["Sprzedaz2026"]
tabela.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
# ⚠️ BEZ TEGO nowe wiersze sa POZA tabela (bez filtra, bez stylu,
#    niewidoczne dla =SUM(Sprzedaz2026[Sprzedaż]))

# ==================================================================
# 8. KOLEKCJA TABEL (2.8) - uwaga na niespodzianki!
# ==================================================================
len(ws.tables)                          # liczba tabel w arkuszu
tab = ws.tables["Sprzedaz2026"]         # po NAZWIE (dziala)
for nazwa in ws.tables:                 # iteracja po nazwach
    print(nazwa, ws.tables[nazwa].ref)

for tabela in ws.tables.values():       # ✅ OBIEKTY Table
    print(tabela.displayName, tabela.column_names)

ws.tables.items()                       # ⚠️ [(nazwa, ZAKRES-napis)] NIE obiekty!
tab = ws.tables.get(table_range="A1:D6")  # ✅ po zakresie
# ⚠️ ws.tables["A1:D6"] -> KeyError! (TableList nie nadpisuje __getitem__)
# ⚠️ ws.tables.get("A1:D6") -> szuka po NAZWIE, nie po zakresie!

del ws.tables["Sprzedaz2026"]           # usuwa DEFINICJE; dane ZOSTAJA

# ==================================================================
# 9. STRUKTURED REFERENCES (2.6)
# ==================================================================
ws["D2"] = "=Sprzedaz2026[@Sprzedaż]-Sprzedaz2026[@Koszt]"   # biezacy wiersz
ws["F2"] = "=SUM(Sprzedaz2026[Sprzedaż])"                     # cala kolumna
ws["F3"] = "=SUM(Sprzedaz2026[[#Data],[Koszt]])"              # dane kolumny
# ⚠️ openpyxl NIE waliduje formul - literowka => #REF! w Excelu
# ⚠️ formuly [@...] MUSZA byc w wierszu nalezycacym do zakresu ref

# ==================================================================
# 10. TABELA JAKO OBSZAR WYDRUKU (2.9/przykład)
# ==================================================================
# openpyxl NIE rozwiaze print_area = nazwa tabeli (ostrzezenie!)
ws.print_area = ws.tables["Sprzedaz2026"].ref    # ✅ uzyj zakresu
# ⚠️ aktualizuj recznie przy kazdym rozszerzeniu tabeli

# ==================================================================
# 11. AUDYT PRZED WYSYLKA (cwiczenie 🔴)
# ==================================================================
# Sprawdz dla kazdej tabeli:
#   □ nazwa bez spacji, bez polskich znakow, bez myslnika
#   □ nazwa NIE wyglada jak adres komorki (Tbl1, ABC123)
#   □ nazwa unikalna w skoroszycie BEZ wielkosci liter
#   □ zakresy dwoch tabel NIE nachodza na siebie
#   □ tabela ma naglowek + >= 1 wiersz danych (ws.max_row > 1)
#   □ naglowki kolumn sa NAPISAMI
#   □ nazwy kolumn unikalne BEZ wielkosci liter
#   □ tabela NIE jest w scalonym zakresie
#   □ ref tabeli obejmuje wszystkie dopisane wiersze

# ==================================================================
# 12. SZYBKA DIAGNOSTYKA (co sprawdzic w Excelu)
# ==================================================================
# Wstazka "Projekt tabeli" -> nazwa tabeli, styl, opcje
# Ctrl+End                   -> gdzie Excel mysli, ze koncza sie dane
# Klik w ostatni wiersz danych -> czy nalezy do tabeli (widac strzalke filtra)
# Nurek: plik .xlsx jako ZIP -> xl/tables/table1.xml
```

## 9. Co dalej

Zamknąłeś drugi duży blok części II. Umięsz już: formatować pojedyncze komórki (moduły 09–10), układać cały arkusz i przygotować go do druku (moduł 11) oraz tworzyć nazwane, samorozszerzające się tabele, na których mogą stać formuły (ten moduł). Trzy rzeczy, które zabierasz ze sobą:

- **Nawyk nazywania.** Zamiast `A1:H100` masz `Sprzedaz2026`. To nie kosmetyka — to zmiana z odwołania **po położeniu** na odwołanie **po tożsamości**, dokładnie jak w kodzie zamiana `lista[3][7]` na `zamowienie.kwota_netto`.
- **Świadomość, że openpyxl nie pilnuje poprawności pliku.** Przepuści nakładające się tabele, puste tabele, nazwy wyglądające jak adresy. Excel odrzuci je u odbiorcy. **Napisz audytor i uruchamiaj go przed wysłaniem pliku** — to najcenniejsze narzędzie, jakie wyniesiesz z tego modułu.
- **Zrozumienie, gdzie kończy się format a zaczyna interfejs.** „Segregator sam się rozszerza" to prawda dla człowieka w Excelu i nieprawda dla programu. Ta granica — między **zapisem stanu** a **zachowaniem aplikacji** — będzie wracać w module 18 (modyfikacja plików) i 26 (architektura).

**W module 13** wchodzimy w **formatowanie warunkowe** — i to jest skok jakościowy, który zaskoczy Cię tak samo jak tabele. Zobaczysz:

- **dlaczego formatowanie warunkowe to nie „efekt zapisany w komórce", lecz „reguła mieszkająca w arkuszu"** — z konsekwencją, że **koloru z formatowania warunkowego nie odczytasz** przez `cell.fill` (to ten sam paradoks co tabela jako warstwa opisu w 2.1);
- **czym jest `DifferentialStyle` (dxf)** i dlaczego reguła niesie tylko *różnicę* względem stylu komórki — analogia do poprawki długopisem na wydruku;
- jak działają **priorytety reguł** (`rule.priority`, `stopIfTrue`) i dlaczego bez `stopIfTrue` „pierwsza pasująca reguła" nie znaczy „jedyna";
- **największą pułapkę:** formuła w regule jest zapisywana z punktu widzenia **lewej-górnej komórki zakresu** i używa odniesień względnych/absolutnych zgodnie z intencją — analogia do wzoru na kalce;
- jak zbudować `CellIsRule` (podświetl wartości > próg) i `FormulaRule` (podświetl **cały wiersz**, gdy w kolumnie D jest „Pilne").

**Zanim przejdziesz dalej, wykonaj jedno ćwiczenie obserwacyjne:**

1. **Uruchom `examples/12_tabela_podstawowa.py` i `examples/12_dopisywanie.py`.** Otwórz oba pliki w Excelu i **porównaj wariant „zły" z „dobrym"**: w złym zaznacz kilka ostatnich wierszy danych i sprawdź, czy mają strzałkę filtra i styl tabeli. Zobaczysz na własne oczy różnicę, którą opisuje 2.12 — i to jest lekcja, którą zapamiętasz lepiej niż z tekstu.
2. **Otwórz `output/12_dopisywanie_dobry.xlsx` jako archiwum ZIP** (zmień rozszerzenie albo użyj `zipfile`). Zajrzyj do `xl/tables/table1.xml` i porównaj go z `xl/tables/table1.xml` z wariantu „złego". **Zobaczysz dokładnie, jak `ref` wygląda w XML** — i to połączy teorię z 2.1 i 2.12 w jedną konkretną rzecz na dysku.
3. **Wpisz w Excelu ręcznie nowy wiersz bezpośrednio pod tabelą** w „dobrym" pliku i zapisz. Potem **wczytaj ten plik openpyxl i sprawdź `tabela.ref`.** Zobaczysz, że Excel zapisał rozszerzony zakres — a to znaczy, że **Twój audytor z ćwiczenia 🔴 musi być uruchamiany nie tylko na plikach, które tworzysz, ale i na tych, które przychodzą z zewnątrz.** Plik, który przeszedł przez ręce człowieka, może mieć inną strukturę niż ten, który wygenerowałeś — i to jest dokładnie tematem modułu 18 (modyfikacja istniejących plików).