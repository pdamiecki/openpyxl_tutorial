Zanim napiszę moduł, zweryfikuję trzy rzeczy, które w materiałach o formatach liczb są najczęściej podawane błędnie: status `Cell.is_date` w 3.1.x, dopuszczalne wartości `Cell.data_type` oraz kody locale (w briefie jest `[$-417]` obok „Polska `[$-415]`" — to wymaga sprawdzenia).

`is_date` **istnieje** w 3.1.x (nie zostało usunięte) — dobrze, że sprawdziłem. Sprawdzę jeszcze dwie rzeczy: automatyczny format daty nadawany przez openpyxl oraz kody locale.

Sprawdzę jeszcze dokładną tablicę wbudowanych formatów openpyxl oraz realne zachowanie kodu locale przy nazwach miesięcy.

# Moduł 10 — Formaty liczb i daty: wartość to nie to, co widzisz

> **Część:** II — Formatowanie · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–09

## 0. W tym module nauczysz się

- **Zrozumiesz fundamentalne rozdzielenie wartości od reprezentacji.** W komórce Excela siedzi jedna liczba. To, co widzi człowiek — „19,99 zł", „35,0%", „2:30:00" — to **warstwa wyświetlania**, całkowicie odłączona od danych. Moduł 09 mówił o pięciu „pieczątkach" wyglądu; tutaj poznasz szósty element stylu i przekonasz się, że rządzi się zupełnie innymi prawami niż pozostałe.
- **Nauczysz się czytać i pisać kody formatów Excela** — od `0.00` przez `#,##0.00 "zł"` aż po `[$-415]dd mmmm yyyy`. Rozbierzesz kod formatu na cztery sekcje (dodatnie; ujemne; zero; tekst) i poznasz każdy znak specjalny: `#`, `0`, `,`, `_`, `*`, `@`, `\`, `[...]`.
- **Zrozumiesz, dlaczego `0,05` z formatem `0%` daje `5%`, a `5` z tym samym formatem daje `500%`** — i dlaczego to najczęstsza pomyłka w raportach finansowych.
- **Poznasz mechanizm dat i godzin od podszewki**: datę w pliku `.xlsx` to liczba dni od epoki, godzinę — ułamek doby. Zobaczysz, jak openpyxl automatycznie nadaje format na podstawie typu Pythona, i **kiedy tego nie robi**.
- **Nauczysz się zapisywać czasy trwania przekraczające 24 godziny** (`[h]:mm:ss`) — czyli jak zrobić raport nadgodzin, który nie „gubi" połowy godzin, oraz zrozumiesz, dlaczego zwykły format `h:mm:ss` to robi.
- **Opanujesz odczytywanie tego, co openpyxl naprawdę myśli o komórce**: `cell.data_type` (`n`, `s`, `d`, `b`, `f`, `e`) oraz `cell.is_date` — wraz z dokładnym kodem źródłowym heurystyki, która za tym stoi, i jej realnymi pułapkami (np. format `#,##0 szt.` uznany za datę!).
- **Zbudujesz centralny rejestr formatów** — jeden słownik, z którego korzysta cały raport, żeby w jednym pliku nie było „trzech odcieni waluty".
- **Zobaczysz, jak format liczby trafia do `xl/styles.xml`** jako `<numFmt numFmtId="164" .../>` — i dlaczego własne formaty zaczynają się dokładnie od identyfikatora 164.

## 1. Intuicja i analogia

### 1.1. Analogia główna: metka cenowa

To najważniejsza analogia tego modułu. Wróćmy na chwilę do sklepu.

Na półce leży **towar**. Na nim przyklejona jest **metka**: „19,99 zł".

Zauważ, że w tym zdaniu są **dwie zupełnie różne rzeczy**:

- **Towar** — istnieje fizycznie, ma masę, cenę w systemie magazynowym, kod kreskowy. System magazynowy zna go jako `19.99` (bez „zł", bez przecinka, bez formatowania).
- **Metka** — kawałek papieru, na którym ktoś **wydrukował** cenę w konkretnym formacie: z symbolem waluty, z przecinkiem dziesiętnym, w konkretnej czcionce i kolorze.

I teraz kluczowe zdanie, które jest tezą całego tego modułu:

> **Zerwanie metki i wydrukowanie nowej nie zmienia towaru ani jego ceny. Zmienia tylko to, co widzi klient.**

Możesz wydrukować metkę „19,99 zł", albo „19.99 PLN", albo „~19,99~ *okazja*", albo „19,99 €" — a system magazynowy nadal będzie znał **jedną, niezmienną liczbę**: `19.99`.

Dokładnie tak działa `number_format` w Excelu:

```python
ws["B2"].value = 19.99                    # towar
ws["B2"].number_format = '#,##0.00 "zł"'  # metka
```

Po zapisaniu pliku **w komórce `B2` siedzi jedna liczba: `19.99`**. Tak zapisuje ją `xl/worksheets/sheet1.xml`. Jest jeszcze **drugi** zapis — w komórce `B2` jest też **numer stylu**, a w tym stylu zapisany jest kod formatu `#,##0.00 "zł"`. Excel, rysując ten arkusz na ekranie, bierze liczbę i przepuszcza ją przez kod formatu.

**Konsekwencja praktyczna numer jeden:** jeśli ktoś w Excelu zobaczy „19,99 zł" i spróbuje to „naprawić" przez zmianę formatu na `0.00`, zobaczy `19.99`. Bo wartość się nie zmieniła — zmieniła się tylko metka.

**Konsekwencja praktyczna numer dwa — ostrzeżenie:** zmiana formatu **nie zaokrągla i nie przelicza niczego**. Jeśli w komórce jest `0.3456789`, a Ty ustawisz format `0.0%`, Excel **wyświetli** `34,6%` (zaokrąglając do wyświetlania), ale w pliku nadal będzie `0.3456789`, a wszystkie formuły będą liczyć na pełnej wartości. Analogia: metka zaokrągla cenę do groszy, ale **kasa wie, że towar kosztuje 19,9876 zł** — i taką kwotę pomnoży przez ilość.

To jest ta sama rodzina problemów co „puste krzesło" z modułu 05. Formatowanie to liczba warstw pomiędzy danymi a człowiekiem.

### 1.2. Druga analogia: formularz urzędowy

Skoro metka mówi „jak wygląda", potrzebujemy drugiej analogii na „jak postępować w skomplikowanych przypadkach".

Wyobraź sobie urzędowy **formularz z rubrykami**:

```
Kwota dodatnia: |____________|
Kwota ujemna:   |____________|
Kwota zero:     |____________|
Tekst:          |____________|
```

Cztery rubryki. I teraz zasada urzędu: **wypełnij pierwszą rubrykę. Jeśli liczba jest ujemna, wypełnij drugą, pierwszą zostaw puste. Jeśli liczba to zero, wypełnij trzecią.** Nigdy nie wypełnia się wszystkich czterech naraz — urzędnik wybiera **tę właściwą** w zależności od wartości.

To **dokładnie** jest struktura kodu formatu Excela: **cztery sekcje oddzielone średnikami, w kolejności `dodatnie; ujemne; zero; tekst`**. Excel patrzy na wartość i używa tylko jednej z czterech sekcji.

```python
ws["B2"].number_format = '#,##0.00 "zł";(#,##0.00)'   # dodatnie normalnie, ujemne w nawiasie
```

Ten kod formatu ma **dwie** sekcje, nie cztery. I to też jest w porządku — brakujące sekcje mają zachowanie domyślne. To jak formularz, w którym wypełniono tylko dwa pola i przyjęto regulamin, że „reszta według standardu".

### 1.3. Trzecia analogia: zegarek i stoper

To analogia potrzebna wyłącznie dla czasów trwania, ale warta własnego akapitu, bo ratuje całe raporty.

Na ręce masz **zegarek**: pokazuje `14:35`. Zegarek to cała doba podzielona na 24 godziny i po `23:59` skacze na `00:00`. Ma sens jako „która jest godzina".

Na bieżni masz **stoper**: pokazuje `01:14:35` i jeśli stoper chodzi 30 godzin, pokaże `30:00:00`. Stoper **nie jest** podzielony na doby. Mierzy **upływ czasu**.

W Excelu obie rzeczy to **ta sama liczba** — ułamek doby (o tym za chwilę). Różni je **format**:

| Sytuacja | Format | Wynik |
|---|---|---|
| „która jest godzina" | `h:mm:ss` | `2:30:00` |
| „ile trwało" | `[h]:mm:ss` | `26:30:00` |

**Nawiasy kwadratowe wokół `h` są tym, co zamienia zegarek w stoper.** Bez nich Excel „zawija" godziny po 24. Z nimi — liczy dalej. Jeśli kiedykolwiek zrobiłeś tabelę nadgodzin i połowa godzin „ginęła", to była właśnie ta jedna para nawiasów.

Zapamiętaj obraz: **`[h]` = stoper, `h` = zegarek.** Nawiasy kwadratowe mówią „nie zawijaj na dobie".

### 1.4. Czwarta analogia: `@` jako przezroczysta rura

Ostatnia, krótka. Zastanów się, co ma zrobić Excel, gdy w komórce z formatem liczbowym jest **tekst**. Bo format `#,##0.00` mówi „pokaż liczbę z dwoma miejscami po przecinku" — a tu jest napis „brak danych".

Odpowiedź Excela jest w czwartej sekcji formatu. Znak **`@`** to **przezroczysta rura na tekst**: „cokolwiek tu wpadnie, przepuść bez zmian". Bez tej sekcji Excel przy nietekstowych sekcjach i tak pokaże tekst, ale z sekcją możesz mu **coś dopisać**, np. `@" (do weryfikacji)"`.

Analogia: to jak **okienko w kopercie**. Wkładasz kartkę, okienko ją przepuszcza — a Ty możesz nadrukować na kopercie coś wokół okienka.

## 2. Teoria

### 2.1. Fundament: wartość i reprezentacja to dwa światy

Zbierzmy w jednym miejscu to, co rozsypane było w poprzednich modułach.

**Świat wartości** — to, co trafia do `xl/worksheets/sheetN.xml` w elemencie `<v>`. Jedna liczba. Bez waluty. Bez przecinka. Bez koloru.

**Świat reprezentacji** — to, co trafia do `xl/styles.xml` w elemencie `<numFmt formatCode="..."/>`, podpiętego pod styl komórki (atrybut `numFmtId` w `<xf>`).

Komórka w XML wygląda więc tak:

```xml
<c r="B2" s="4">
  <v>19.99</v>
</c>
```

A gdzieś dalej, w `styles.xml`:

```xml
<numFmts count="1">
  <numFmt numFmtId="164" formatCode="#,##0.00 &quot;zł&quot;"/>
</numFmts>
<cellXfs count="5">
  ...
  <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  ...
</cellXfs>
```

Cztery warstwy pośrednie: komórka wskazuje styl numer 4, styl numer 4 wskazuje format `numFmtId=164`, ten format ma kod `#,##0.00 "zł"`, a wartość to `19.99`. **Excel składa to na ekranie dopiero w momencie rysowania.**

**Identyfikator 164 to nie przypadek.** W źródle `openpyxl/styles/numbers.py` znajdziesz stałą:

```python
BUILTIN_FORMATS_MAX_SIZE = 164
```

Identyfikatory od 0 do 163 są **zarezerwowane dla formatów wbudowanych**. Twój własny kod formatu zawsze dostanie numer 164 lub wyższy. To ta sama rodzina faktów co zarezerwowane wypełnienia `fillId=0` i `fillId=1` z modułu 09 — szczegóły formatu pliku, o których nikt nie mówi, a które widać w każdym raporcie diagnostycznym.

### 2.2. `number_format` to zwykły napis — i to jest jego siła oraz jego słabość

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ws["A1"] = 19.99
ws["A1"].number_format = '#,##0.00 "zł"'

print(type(ws["A1"].number_format).__name__)   # str
print(ws["A1"].number_format)                  # #,##0.00 "zł"
print(ws["A1"].value)                          # 19.99 - wartosc nietknieta
```

**Co się dzieje w pamięci.** Do atrybutu `number_format` komórki trafia zwykły napis Pythona. Nie ma tu żadnego obiektu, żadnego walidatora składni (o jednym wyjątku za chwilę), żadnego parsowania. **W przeciwieństwie do `Font` czy `PatternFill` nie ma tu „pieczątki"** — jest napis przypisany do komórki.

**Co trafi do pliku.** Jeśli ten napis jest jednym ze znanych formatów wbudowanych, openpyxl podepnie odpowiedni `numFmtId` z zakresu 0–163. Jeśli nie — utworzy wpis `<numFmt>` z własnym identyfikatorem od 164 w górę.

**Jedyny wyjątek od „żadnej walidacji":** w źródle znajdziesz `NumberFormatDescriptor`:

```python
class NumberFormatDescriptor(String):
    def __set__(self, instance, value):
        if value is None:
            value = FORMAT_GENERAL
        super(NumberFormatDescriptor, self).__set__(instance, value)
```

Czyli: **jeśli przypiszesz `None`, dostaniesz `"General"`**. Nie `None`, nie błąd — po prostu format ogólny. To zachowanie jest sensowne (pusta wartość nie ma sensu jako kod formatu), ale warto je znać, bo debugując „czemu ten format się nie zapisał", możesz odkryć, że przekazałeś `None` z jakiegoś słownika.

**A teraz najważniejsze zdanie tego podrozdziału:**

> **openpyxl NIE waliduje kodu formatu.** Możesz wpisać `"to jest bezsensowny napis"` i openpyxl to zapisze bez mrugnięcia okiem. Excel przy otwarciu pokaże ten tekst dosłownie w komórce albo potraktuje go jako format nietypowy.

To jest kategoria **„openpyxl potrafi tylko częściowo"** z modułu 03: API istnieje, wywołanie przechodzi, plik powstaje — ale **ani openpyxl, ani Ty nie macie gwarancji, że Excel to zrozumie**. Dlatego w tym module tyle miejsca zajmuje nauka **składni** — bo biblioteka Ci jej nie sprawdzi, musisz ją znać sam.

### 2.3. Formaty wbudowane — pełna tablica

W `openpyxl.styles.numbers` znajdziesz słownik `BUILTIN_FORMATS`. To **dokładna kopia** tego, co Excel ma wpisane na stałe. Warto go znać, bo:

1. Pozwala sprawdzić, czy Twój kod formatu jest już wbudowany (`is_builtin(fmt)`).
2. Wyjaśnia, dlaczego pewne formaty „kosztują zero" — są zapisywane jako liczba, nie jako napis.
3. Jest to gotowa lista poprawnych kodów, na której możesz się wzorować.

| ID | Kod formatu | Do czego |
|---|---|---|
| 0 | `General` | Ogólny |
| 1 | `0` | Liczba całkowita |
| 2 | `0.00` | Dwa miejsca dziesiętne |
| 3 | `#,##0` | Tysiące, bez groszy |
| 4 | `#,##0.00` | Tysiące, z groszami |
| 5 | `"$"#,##0_);("$"#,##0)` | Waluta USD |
| 6 | `"$"#,##0_);[Red]("$"#,##0)` | Waluta USD, ujemne na czerwono |
| 7 | `"$"#,##0.00_);("$"#,##0.00)` | Waluta USD z groszami |
| 8 | `"$"#,##0.00_);[Red]("$"#,##0.00)` | Jak wyżej, ujemne na czerwono |
| 9 | `0%` | Procent |
| 10 | `0.00%` | Procent z dwoma miejscami |
| 11 | `0.00E+00` | Notacja naukowa |
| 12 | `# ?/?` | Ułamek zwykły |
| 13 | `# ??/??` | Ułamek zwykły, dwa znaki mianownika |
| 14 | `mm-dd-yy` | Data (US) |
| 15 | `d-mmm-yy` | Data |
| 16 | `d-mmm` | Dzień i miesiąc |
| 17 | `mmm-yy` | Miesiąc i rok |
| 18 | `h:mm AM/PM` | Godzina 12-godzinna |
| 19 | `h:mm:ss AM/PM` | Godzina z sekundami, 12-godzinna |
| 20 | `h:mm` | Godzina 24-godzinna |
| 21 | `h:mm:ss` | Godzina z sekundami |
| 22 | `m/d/yy h:mm` | Data i godzina |
| 37 | `#,##0_);(#,##0)` | Księgowy, bez symbolu |
| 38 | `#,##0_);[Red](#,##0)` | Księgowy, ujemne na czerwono |
| 39 | `#,##0.00_);(#,##0.00)` | Księgowy z groszami |
| 40 | `#,##0.00_);[Red](#,##0.00)` | Jak wyżej, ujemne czerwone |
| 41–44 | `_(* #,##0_);...` | Formaty **księgowe** z wyrównaniem symbolu |
| 45 | `mm:ss` | Minuty i sekundy |
| 46 | `[h]:mm:ss` | **Czas trwania przekraczający 24 h** |
| 47 | `mmss.0` | — |
| 48 | `##0.0E+0` | Notacja naukowa |
| 49 | `@` | **Tekst** |

Zwróć uwagę na dwie rzeczy:

- **Dziury w numeracji (23–36).** Format ma zarezerwowane miejsca, których openpyxl nie wypełnia — to obszary zależne od lokalizacji w Excelu. Nie próbuj „odtworzyć" tych numerów.
- **Format 49 to `@` czyli tekst.** Osobny, wbudowany format dla danych tekstowych. To wyjaśnia, dlaczego w kodach pocztowych tak często mówi się „ustaw format tekstowy" — istnieje na to gotowy wpis.

**Wbudowane kody w praktyce:**

```python
from openpyxl.styles.numbers import (
    FORMAT_DATE_DATETIME, FORMAT_DATE_TIMEDELTA, FORMAT_PERCENTAGE,
    FORMAT_TEXT, FORMAT_NUMBER_COMMA_SEPARATED1, builtin_format_id, is_builtin,
)

print(FORMAT_DATE_DATETIME)      # yyyy-mm-dd h:mm:ss
print(FORMAT_DATE_TIMEDELTA)     # [hh]:mm:ss
print(FORMAT_TEXT)               # @
print(is_builtin('#,##0.00'))    # True - to format wbudowany (id 4)
print(builtin_format_id('#,##0.00'))   # 4
print(is_builtin('#,##0.00 "zł"'))     # False - wlasny format
```

**Zaleta używania wbudowanych:** plik ma mniejszy `styles.xml` (brak sekcji `<numFmts>`), bo wystarczy numer. **Wada:** format wbudowany wygląda inaczej w każdym locale — użytkownik z angielskim systemem zobaczy `1,234.56`, a polskim `1 234,56`. Jeśli chcesz **pełnej kontroli nad wyglądem**, używaj własnego kodu (i wtedy `numFmtId >= 164`).

### 2.4. Anatomia kodu formatu

#### 2.4.1. Cztery sekcje

```text
sekcja1 ; sekcja2 ; sekcja3 ; sekcja4
  │         │          │         └── tekst
  │         │          └──────────── zero
  │         └─────────────────────── wartości ujemne
  └───────────────────────────────── wartości dodatnie
```

Zasady:

- **Excel używa dokładnie jednej sekcji** — tej, która pasuje do wartości.
- Brakujące sekcje mają zachowanie domyślne: jeśli jest jedna sekcja, dotyczy wszystkich liczb. Jeśli dwie — pierwsza dodatnich i zera, druga ujemnych. Jeśli trzy — zero dostaje trzecią.
- **Pierwsze trzy sekcje dotyczą liczb. Czwarta dotyczy tekstu.** To rozdzielenie jest ważne: `@` w czwartej sekcji nie ma nic wspólnego z liczbami.

```python
# Cztery sekcje: dodatnie | ujemne | zero | tekst
ws["A1"].number_format = '#,##0.00;-#,##0.00;"—";@" [tekst]"'
#   1 234.56   |  -1 234.56   |   —   |  abc [tekst]
```

#### 2.4.2. Znaki specjalne — pełny słownik

| Znak | Znaczenie | Przykład |
|---|---|---|
| `0` | **Obowiązkowe** miejsce cyfry. Brakującą cyfrę pokaże jako zero. | `00` → `7` pokaże `07` |
| `#` | **Opcjonalne** miejsce cyfry. Brakującej nie pokaże. | `0.0#` → `1.5` pokaże `1.5` |
| `?` | Miejsce cyfry, ale zarezerwowane (spacja zamiast cyfry) — wyrównanie dziesiętne | `???.??` |
| `.` | Separator dziesiętny (w zapisie formatu Excela to zawsze kropka) | `0.00` |
| `,` | **Kontekst zależy od pozycji!** Patrz niżej. | `#,##0` lub `#,##0,` |
| `%` | Pomnóż przez 100 i dopisz znak procenta | `0%` |
| `E+` / `E-` | Notacja naukowa | `0.00E+00` |
| `/` | Ułamek zwykły | `# ?/?` |
| `"..."` | **Tekst dosłowny** | `0.00 "zł"` |
| `\` | Escape — następny znak jako tekst | `0.00 \z\l` |
| `_` | Wstaw spację szerokości następnego znaku | `_)` |
| `*` | Powtarzaj następny znak, aż wypełni szerokość kolumny | `*-` |
| `@` | Miejsce na wartość tekstową | `@" szt."` |
| `[Kolor]` | Kolor tekstu: `[Red]`, `[Blue]`, `[Green]`, `[Black]`, `[White]`, `[Cyan]`, `[Magenta]`, `[Yellow]` | `[Red]#,##0` |
| `[$-xxx]` | Kod lokalizacji (LCID) | `[$-415]dd mmmm yyyy` |
| `[hh]`, `[mm]`, `[ss]` | **Czas trwania** — nie zawijaj na dobie | `[h]:mm:ss` |
| `[>100]` itp. | Warunek | `[>100]0.0;[Red]0.00` |

#### 2.4.3. Przecinek: najdziwniejszy znak w całym systemie

Przecinek ma **trzy różne znaczenia** zależnie od pozycji:

```text
# , ##0        -> separator tysiecy          (1234567 -> 1,234,567)
0 . 00 ,       -> podziel przez 1000        (1234567 -> 1234.57)
# , ##0 . 00 ,  -> separator + podzial
```

Konkretnie:

| Kod | Wartość 1234567 wyświetli się jako | Dlaczego |
|---|---|---|
| `#,##0` | `1 234 567` | przecinek w środku = separator tysięcy |
| `#,##0,` | `1 235` | przecinek **na końcu** = podziel przez 1000 |
| `#,##0,,` | `1` | dwa przecinki na końcu = podziel przez 1 000 000 |
| `#,##0,"tys."` | `1 235 tys.` | dzielenie + tekst |

**Dzielenie przez tysiąc** to bardzo praktyczna funkcja w raportach: chcesz pokazać budżet w tysiącach, ale nie chcesz zmieniać wartości w danych (bo formuły mają liczyć na pełnych kwotach). Analogia: to **skala na mapie**. Mapa nie zmniejsza terenu — tylko ty podajesz ją w skali 1:1000.

```python
# Wartosc pozostaje pelna, wyswietlanie idzie w tysiacach.
ws["B2"].value = 12_345_678
ws["B2"].number_format = '#,##0,"tys."'      # wyswietli: 12 346 tys.
```

**Uwaga o separatorach:** w kodzie formatu zawsze piszesz `.` jako separator dziesiętny i `,` jako separator tysięcy — **niezależnie od locale**. Excel sam podmienia je na lokalne znaki przy wyświetlaniu. To znaczy: **nie wpisujesz w kodzie polskiego przecinka dziesiętnego.** Kod `0,00` nie znaczy „zero z dwoma miejscami po przecinku po polsku" — znaczy „podziel przez 1000 dwa razy z zerami". To realna pułapka dla osób, które myślą po polsku i piszą `# ##0,00`.

#### 2.4.4. `_` i `*` — wyrównywanie i wypełnianie

Te dwa znaki rzadko są potrzebne, ale pojawiają się w formatach wbudowanych, więc warto je rozumieć.

```python
# "_(" znaczy: zrob puste miejsce o szerokosci "("
ws["A1"].number_format = '#,##0_)'      # liczby wyrownane do prawej jakby z nawiasem
ws["A2"].number_format = '(#,##0_)'     # ujemne w nawiasie, wyrownane

# "*" powtarza znak do konca szerokosci kolumny
ws["A3"].number_format = '#,##0*-'      # kreski dopelniajace do krawedzi
```

**Analogia `_`:** to jak **pusty kafelek w drukowanym formularzu**. Nie wpisujesz tam nic, ale kafelek ma rozmiar i dzięki temu kolumny cyfr w całej tabeli się schodzą. Zobacz format wbudowany 42:

```python
'_(("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)'
```

Cała ta plątanina to nic innego jak **sposób na wyrównanie kwot do prawej tak, żeby symbol waluty był zawsze na tej samej pozycji**, niezależnie od długości liczby. Tak wygląda „format księgowy".

**Analogia `*`:** to jak **linia kropkowana do końca wiersza w spisie treści**. Wypełnia przestrzeń do krawędzi. Ma zastosowanie głównie estetyczne.

### 2.5. Formatowanie warunkowe **wewnątrz** kodu formatu

To inna rzecz niż formatowanie warunkowe z modułów 13–14. Tutaj warunek siedzi w samym kodzie formatu i dotyczy tylko wyświetlania treści tej komórki.

```python
# Jesli wartosc > 100 -> format z jedna cyfra po przecinku
# w przeciwnym razie -> format z dwiema cyframi i na czerwono
ws["A1"].number_format = '[>100]0.0;[Red]0.00'
```

Składnia: `[warunek]` przed sekcją. Obsługiwane operatory to `>`, `<`, `>=`, `<=`, `=` oraz `<>`.

⚠️ **Uczciwe ostrzeżenie o ograniczeniach.** Excel pozwala na **maksymalnie dwa** warunki w kodzie formatu (plus sekcja domyślna dla reszty). To nie jest mechanizm, w którym można zbudować piętnaście reguł. Dodatkowo:

- openpyxl **nie waliduje** tych warunków — wpisujesz napis i modlisz się, żeby Excel go przełknął.
- Dokumentacja Excela jest w tym miejscu oszczędna, a zachowanie zależy od wersji.

**Praktyczna rada:** jeśli potrzebujesz więcej niż jednego warunku albo złożonej logiki — **użyj formatowania warunkowego z modułu 13/14**, a nie tego mechanizmu. Warunki w kodzie formatu są przydatne w jednym wąskim przypadku: gdy chcesz, żeby **sama reprezentacja liczbowa** się zmieniała (np. duże kwoty pokazywać w milionach, a małe normalnie), a nie kolor.

**Kolor zamiast warunku** — to natomiast jest w pełni niezawodne i szeroko stosowane:

```python
# Ujemne na czerwono, dodatnie normalnie - klasyka sprawozdan
ws["A1"].number_format = '#,##0.00;[Red]-#,##0.00'
```

Zwróć uwagę: `[Red]` działa tylko w tych sekcjach, w których go postawisz. Kolor **nie zapisuje się w komórce** — tak samo jak formatowanie warunkowe z modułu 13, kolor z kodu formatu pojawia się dopiero, gdy Excel narysuje komórkę. Odczytując plik przez openpyxl, **nie znajdziesz tego koloru** w `cell.font.color`. Jeśli kiedyś będziesz próbował wykryć „czerwone liczby" w pliku od klienta i nie znajdziesz ich w stylu komórki — to prawdopodobnie ten mechanizm.

### 2.6. Daty i czas: liczba, która udaje kalendarz

#### 2.6.1. Excel nie ma typu daty

To zdanie może zaskoczyć, ale jest prawdziwe i wyjaśnia połowę problemów z datami.

> **W pliku `.xlsx` data to liczba.** Nie kalendarz, nie „1 marca 2026", tylko liczba zmiennoprzecinkowa: **liczba dni od epoki**.

Przykłady (dla epoki 1900):

| Data w Excelu | Liczba w pliku |
|---|---|
| 1900-01-01 | `1` |
| 2026-03-01 | `46082` |
| 2026-03-01 12:00 | `46082.5` |
| 2026-03-01 06:00 | `46082.25` |
| 2026-03-01 18:00 | `46082.75` |

Część całkowita to **dzień**, część ułamkowa to **pora dnia** — ułamek doby. `0.5` to południe, `0.25` to 6:00, `0.75` to 18:00.

**Analogia:** wyobraź sobie, że ktoś kazał Ci zapisywać daty w zeszycie wyłącznie jako liczby. „1 marca 2026, godzina 6:00" zapisujesz jako `46082.25`. Cały system działa, dopóki wszyscy wiedzą, że to „dni od 1 stycznia 1900". Excel właśnie tak robi — a format `yyyy-mm-dd` to **instrukcja dla człowieka, jak rozkodować tę liczbę**.

To dlatego w module 05 mówiliśmy, że data w Excelu to liczba — teraz widzisz pełny mechanizm.

#### 2.6.2. Epoka, czyli skąd liczyć dni

```python
from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

wb = Workbook()
print(wb.epoch)      # 1899-12-30 00:00:00  (Windows, domyslnie)

wb.epoch = CALENDAR_MAC_1904
print(wb.epoch)      # 1904-01-01 00:00:00  (Mac)
```

Dwie rzeczy do zapamiętania:

1. **Domyślna epoka to nie 1900-01-01, a 1899-12-30.** Wynika to z historycznej korekty Excela (i z fikcyjnego 29 lutego 1900 — patrz niżej). Jeśli samodzielnie przeliczasz daty na liczby, **nie zgaduj epoki** — używaj `wb.epoch`.
2. **Pliki z Maca mają inną epokę** (`1904`). Excel ma tryb „skoroszyt od 1904", pozostały po starych Macintoshu. Jeśli wczytasz plik z Maca i wszystkie daty przesuną się o 4 lata — to właśnie to. openpyxl odczytuje epokę z pliku, więc zwykle nie musisz tego obsługiwać ręcznie, ale **musisz to sprawdzić, jeśli sam przeliczasz daty**.

#### 2.6.3. Fikcyjny 29 lutego 1900 — słynny błąd

Excel traktuje rok 1900 jako przestępny — choć **nie był** (rok podzielny przez 100, ale nie przez 400). W numeracji Excela:

| Serial | Excel pokazuje |
|---|---|
| 59 | 1900-02-28 |
| **60** | **1900-02-29** ← nie istnieje! |
| 61 | 1900-03-01 |

To słynny błąd, który Excel zachowuje celowo dla zgodności z Lotusem 1-2-3. Ma to dwie konsekwencje dla Ciebie:

- **Daty sprzed 1 marca 1900 są przesunięte o jeden dzień.** openpyxl obsługuje to poprawnie przy konwersji (odejmuje dzień dla wartości poniżej 61), ale jeśli sam konwertujesz, musisz o tym pamiętać.
- **Nigdy nie przeliczaj dat samodzielnie, jeśli możesz użyć biblioteki.** To jest dokładnie ta klasa problemu, w której ręczna implementacja wydaje się prosta i zawsze gdzieś się rozjeżdża.

#### 2.6.4. Konwersja: biblioteka kontra własna arytmetyka

openpyxl udostępnia narzędzia:

```python
from datetime import date, datetime

from openpyxl.utils.datetime import from_excel, to_excel

# Data -> serial
print(to_excel(date(2026, 3, 1)))            # 46082
print(to_excel(datetime(2026, 3, 1, 12, 0))) # 46082.5

# Serial -> data
print(from_excel(46082))                     # 2026-03-01 00:00:00
print(from_excel(46082.5))                   # 2026-03-01 12:00:00
```

⚠️ **Sygnatury tych funkcji zmieniały się między wersjami** (parametr `epoch`, parametr `timedelta`). **Sprawdź w dokumentacji swojej wersji**, zanim ich użyjesz w kodzie produkcyjnym. W praktyce rzadko potrzebujesz ich ręcznie — openpyxl konwertuje daty automatycznie przy przypisywaniu wartości.

#### 2.6.5. Automatyczne nadanie formatu — i kiedy go NIE ma

To jest jeden z tych mechanizmów, które oszczędzają czas, dopóki nie zaczną Cię zaskakiwać.

W źródle `openpyxl/cell/cell.py` znajdziesz:

```python
if dt == 'd':
    if not is_date_format(self.number_format):
        self.number_format = get_time_format(t)
```

Czytając to po polsku: **jeśli przypisujesz wartość typu czasowego (`datetime`, `date`, `time`, `timedelta`) i komórka nie ma jeszcze formatu daty — openpyxl nada jej format zależny od typu.**

Konkretne przypisania:

| Typ Pythona | Format nadany automatycznie |
|---|---|
| `datetime` | `yyyy-mm-dd h:mm:ss` |
| `date` | `yyyy-mm-dd` |
| `time` | `h:mm:ss` |
| `timedelta` | `[hh]:mm:ss` |

```python
from datetime import date, datetime, time, timedelta

from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ws["A1"] = date(2026, 3, 1)
ws["A2"] = datetime(2026, 3, 1, 14, 30)
ws["A3"] = time(14, 30)
ws["A4"] = timedelta(hours=26, minutes=30)

for adres in ("A1", "A2", "A3", "A4"):
    print(f"{adres}: value={ws[adres].value!r}  format={ws[adres].number_format!r}")
```

Wynik:

```
A1: value=datetime.datetime(2026, 3, 1, 0, 0)  format='yyyy-mm-dd'
A2: value=datetime.datetime(2026, 3, 1, 14, 30)  format='yyyy-mm-dd h:mm:ss'
A3: value=datetime.time(14, 30)  format='h:mm:ss'
A4: value=datetime.timedelta(days=1, seconds=8100)  format='[hh]:mm:ss'
```

Zwróć uwagę na linię `A1`: przypisałeś `date`, a dostałeś `datetime`. To normalne — openpyxl sprowadza daty do `datetime`, bo Excel nie ma osobnego typu dla daty bez godziny.

**Kiedy automat NIE zadziała — pułapka.** Warunek `if not is_date_format(self.number_format)` sprawia, że **jeśli komórka ma już format daty, openpyxl go NIE nadpisze**. To zachowanie jest celowe (jest na nie osobny test w repozytorium openpyxl o nazwie `test_not_overwrite_time_format`), ale ma niezręczną konsekwencję:

```python
ws["B1"].number_format = "mmm-yy"      # format daty: miesiac-rok
ws["B1"] = date(2026, 3, 1)

print(ws["B1"].number_format)          # mmm-yy  <- NIE zostal nadpisany
```

**Po co to komu?** Bo pozwala **ustawić format przed wartością**:

```python
# Kolejnosc 1: format najpierw, wartosc potem - pelna kontrola
ws["D2"].number_format = "yyyy-mm-dd"
ws["D2"] = date(2026, 3, 1)

# Kolejnosc 2: wartosc najpierw, format potem - tez dziala
ws["D3"] = date(2026, 3, 1)
ws["D3"].number_format = "yyyy-mm-dd"
```

W obu kolejnościach dostaniesz to samo, więc **nie ma tu pułapki w sensie „nie działa"**. Jest natomiast pułapka **odwrotna, groźniejsza**: jeśli komórka ma format daty z **innego powodu** (np. odziedziczyła go ze stylu nazwanego albo z formatowania warunkowego — nie, warunkowe nie wpływa na `number_format`; ale ze stylu nazwanego tak), to przypisanie daty **nie poprawi** formatu, i możesz zobaczyć datę wyświetloną jako liczbę. Diagnostyka: sprawdź `cell.number_format` **przed** przypisaniem wartości.

#### 2.6.6. `is_date_format` — heurystyka, którą warto znać na wylot

openpyxl musi jakoś zgadnąć, czy dany kod formatu „jest datą". Robi to funkcją:

```python
def is_date_format(fmt):
    if fmt is None:
        return False
    fmt = fmt.split(";")[0]              # tylko pierwsza sekcja!
    fmt = STRIP_RE.sub("", fmt)          # usun teksty w cudzyslowie i [...] locale
    return re.search(r"(?<![_\\])[dmhysDMHYS]", fmt) is not None
```

Rozłóżmy to na czynniki pierwsze, bo każdy fragment ma praktyczne znaczenie:

1. **`fmt.split(";")[0]`** — patrzy **tylko na pierwszą sekcję**. Format `0.00;yyyy-mm-dd` zostanie uznany za niedatowy (bo pierwsza sekcja to liczba). Ma to sens, bo warunek „czy to data" rozstrzyga konwencja dla wartości dodatnich.
2. **`STRIP_RE.sub("", fmt)`** — usuwa **teksty w cudzysłowie** (`"..."`) oraz **nawiasy kwadratowe inne niż `[h]`, `[m]`, `[s]`**. Czyli `"zł"`, `[$-415]`, `[Red]` są usuwane przed sprawdzaniem.
3. **`re.search(r"(?<![_\\])[dmhysDMHYS]", fmt)`** — szuka **pierwszego wystąpienia** którejkolwiek z liter `d`, `m`, `h`, `y`, `s` (małe lub duże), **które nie jest poprzedzone podkreśleniem `_` ani backslashem `\`**.

**I tu jest pułapka, którą warto zobaczyć na własne oczy.** Bo litera `s` (sekundy) i litera `m` (miesiące/minuty) występują w zwykłych słowach:

```python
from openpyxl import Workbook
from openpyxl.styles.numbers import is_date_format

wb = Workbook()
ws = wb.active

# Format WALUTOWY, ale bez cudzyslowu wokol "szt."
ws["A1"] = 1200
ws["A1"].number_format = '#,##0 szt.'
print(is_date_format('#,##0 szt.'))    # True!  <- bo 's' od slowa "szt."

print(ws["A1"].is_date)                # True   <- komorka uwazana za date

# To samo, ale w cudzyslowie - litera 's' jest w tekscie, wiec zostaje usunieta
print(is_date_format('#,##0 "szt."'))  # False
```

**To jest realna, reprodukowalna pułapka.** Format `#,##0 szt.` zostanie przez openpyxl uznany za format daty, bo zawiera literę `s`. Konsekwencje:

- `cell.is_date` zwróci `True` dla zwykłej liczby.
- Jeśli później **wczytasz** ten plik i będziesz konwertować wartości na daty „bo `is_date` mówi, że to data" — dostaniesz bzdury.
- Excel sam wyświetli to prawdopodobnie poprawnie (bo `s` bez kontekstu godzinowego jest w formatach Excela ignorowane), ale **openpyxl myśli inaczej niż Excel**.

**Wniosek praktyczny, jeden i konkretny:**

> **Tekst w kodzie formatu zawsze zamykaj w cudzysłowie.** `'#,##0 "szt."'`, nie `'#,##0 szt.'`.

Ta jedna zasada: (a) eliminuje heurystyczne pomyłki openpyxl, (b) czyni kod formatu jednoznacznym dla Excela, (c) chroni przed sytuacją, w której Twoje słowo tekstowe zawiera `d`, `m`, `h`, `y` albo `s` i zaczyna być interpretowane jako format daty.

#### 2.6.7. Mikrosekundy i precyzja czasu

```python
from datetime import datetime

ws["A1"] = datetime(2026, 3, 1, 14, 30, 45, 123456)
print(ws["A1"].number_format)     # yyyy-mm-dd h:mm:ss
```

Domyślny format **nie pokazuje mikrosekund**. Wartość `123456` mikrosekund nadal tam jest (jako część ułamka doby), ale reprezentacja pokaże `14:30:45`.

Jeśli naprawdę potrzebujesz mikrosekund:

```python
ws["A1"].number_format = "hh:mm:ss.000"      # tysieczne sekundy
```

⚠️ Ale uwaga na dwa ograniczenia:

1. **Rozdzielczość liczby zmiennoprzecinkowej Excela.** Excel przechowuje daty jako liczby zmiennoprzecinkowe podwójnej precyzji. Doba ma 86 400 sekund, więc teoretyczna rozdzielczość to około $10^{-5}$ sekundy — ale w praktyce po kilku operacjach arytmetycznych precyzja się rozjeżdża. **Nie buduj na tym systemu pomiaru czasu.**
2. **`datetime` ze strefą czasową (`tzinfo`) nie jest obsługiwany.** openpyxl nie zapisuje informacji o strefie — o tym szerzej poniżej.

#### 2.6.8. Strefy czasowe — uczciwie o ograniczeniu

```python
from datetime import datetime, timezone

dt_aware = datetime(2026, 3, 1, 14, 30, tzinfo=timezone.utc)
ws["A1"] = dt_aware          # przypisanie przejdzie...
```

openpyxl **przyjmie** taką wartość, ale przy zapisie strefa zostanie **utracona** — Excel nie ma koncepcji strefy czasowej w komórce. W efekcie:

- **zapiszesz** `14:30` bez informacji, że to UTC,
- **przy odczycie** dostaniesz `datetime` **bez `tzinfo`** (`naive`),
- ktoś, kto porówna odczytaną wartość z `datetime` świadomym strefy, **zawsze** dostanie `False` przy porównaniu (bo Python nie porównuje naiwnych ze świadomymi).

**Zasada:** jeśli w Twoim systemie czas ma strefę, **konwertuj na UTC i zapisuj jako naiwny `datetime` w UTC**, a informację o strefie umieść w nagłówku kolumny albo w arkuszu `_meta`. Nigdy nie zakładaj, że strefa „przejdzie" przez plik.

### 2.7. Lokalizacja formatów: `[$-xxx]` i nazwy miesięcy

#### 2.7.1. Kod LCID w kwadracie

Kod formatu może zaczynać się od **kodu lokalizacji** w nawiasie kwadratowym:

```text
[$-415]dd mmmm yyyy
```

`415` to identyfikator lokalizacji (LCID) w zapisie szesnastkowym. Pełny zapis to `0x0415`; Excel pomija wiodące zero. Oto najważniejsze kody:

| Kod | Język |
|---|---|
| `[$-405]` | Czeski |
| `[$-407]` | Niemiecki |
| `[$-409]` | Angielski (US) |
| `[$-410]` | Włoski |
| `[$-414]` | Norweski |
| **`[$-415]`** | **Polski** ← to jest ten właściwy |
| `[$-416]` | Portugalski |
| `[$-417]` | Łotewski |
| `[$-419]` | Rosyjski |
| `[$-421]` | Indonezyjski |

⚠️ **Uwaga, którą warto zapamiętać, bo łatwo ją pomylić:** `415` to **polski**, a `417` to **łotewski**. W internecie krąży sporo materiałów, w których te dwie wartości są mieszane. Jeśli Twój format `[$-417]dd mmmm yyyy` pokazuje „marts" zamiast „marzec" — wiesz już, dlaczego.

**Jak to działa w praktyce:**

```python
from datetime import date

ws["A1"] = date(2026, 3, 1)
ws["A1"].number_format = "[$-415]dd mmmm yyyy"
# Excel pokaze: 01 marzec 2026   (nazwa miesiaca po polsku)
```

To działa, bo LCID mówi Excelowi, **w jakim języku mają być nazwy miesięcy i dni tygodnia** w tym konkretnym formacie. To bardzo przydatne — możesz wygenerować jeden plik, w którym kolumna dat jest po polsku, a inna po angielsku:

```python
ws["A1"].number_format = "[$-415]dd mmmm yyyy"    # 01 marzec 2026
ws["A2"].number_format = "[$-409]dd mmmm yyyy"    # 01 March 2026
```

#### 2.7.2. Dlaczego `mmmm` pokazuje nazwę w „cudzym" języku

I tu rzecz, która zaskakuje wszystkich:

> **Nazwa miesiąca w formacie `mmmm` zależy od ustawień systemowych odbiorcy (i/lub języka interfejsu Excela), a nie od języka Twojego kodu Pythona.**

Konsekwencje:

- **Ty generujesz plik w Polsce** i widzisz `marzec`.
- **Odbiorca w Wielkiej Brytanii** otwiera ten sam plik i widzi `March`.
- **Odbiorca z angielskim Excelem na polskim systemie** zobaczy zależnie od tego, co wygrywa w jego konfiguracji — i to jest miejsce, w którym „działa u mnie" bywa prawdą lokalną.

**To nie jest błąd openpyxl.** To architektura: plik `.xlsx` nie zawiera tekstu „marzec" — zawiera liczbę `46082` i instrukcję `mmmm`. **Tekst powstaje na komputerze odbiorcy, w momencie rysowania komórki.** Analogia: wysyłasz przepis kulinarny, a nie gotowe danie — a kucharz użyje lokalnych nazw składników.

**Jak uzyskać pewność co do języka nazwy miesiąca?** Masz dwie drogi:

1. **Kod LCID** (`[$-415]mmmm`) — najbliższy ideału, ale wymaga sprawdzenia w docelowym Excelu odbiorcy.
2. **Zapis tekstu, nie daty** — jeśli raport musi wyglądać **identycznie** wszędzie, zapisz `"marzec 2026"` jako `str`. Tracisz wtedy możliwość sortowania po dacie i wykonywania obliczeń. To świadomy kompromis.

**Rekomendacja:** dla raportów wewnętrznych używaj formatów dat z LCID. Dla dokumentów, które muszą wyglądać identycznie na całym świecie (np. faktury eksportowe) — albo używaj ISO 8601 (`yyyy-mm-dd`, który nie ma nazw miesięcy i jest jednoznaczny), albo zapisuj tekst.

### 2.8. Czasy trwania przekraczające 24 godziny

Wracamy do stopera z analogii 1.3, tym razem z kodem.

**Problem:** pracownik w marcu przepracował 26 godzin nadgodzin. Zapisujesz wynik jako `datetime` — i nie możesz, bo `datetime` to punkt w czasie „któraś godzina", a nie „upłynęło 26 godzin". Potrzebujesz `timedelta`.

```python
from datetime import timedelta

from openpyxl import Workbook

wb = Workbook()
ws = wb.active

# Nadgodziny: 26 godzin 30 minut
ws["A1"] = timedelta(hours=26, minutes=30)
print(ws["A1"].number_format)      # [hh]:mm:ss  <- nawiasy sa!

# Suma tygodniowa: 40 godzin 15 minut
ws["A2"] = timedelta(hours=40, minutes=15)

# Suma - formuła
ws["A3"] = "=SUM(A1:A2)"
ws["A3"].number_format = "[h]:mm:ss"

wb.save("output/10_nadgodziny.xlsx")
```

**Co się dzieje w pamięci.** `timedelta` to obiekt Pythona przechowujący dni, sekundy i mikrosekundy. openpyxl konwertuje go na **liczbę** — ułamek doby. `timedelta(hours=26, minutes=30)` to $26{,}5 / 24 = 1{,}104166...$ dnia. Do pliku trafia więc liczba `1.1041666666666667`.

**Co trafi do pliku.** Liczba **plus** kod formatu `[hh]:mm:ss`. I ten kod jest całym sekretem: nawiasy kwadratowe mówią Excelowi „**nie zawijaj na dobie**". Excel mnoży ułamek przez 24 i pokazuje `26:30:00`.

**Co by było bez nawiasów:**

```python
ws["B1"] = timedelta(hours=26, minutes=30)
ws["B1"].number_format = "hh:mm:ss"      # ❌ zle!
# Excel pokaze: 02:30:00   <- 26 godzin "zawiniete" na dobe, +2:30
```

To jest **najczęstszy błąd w raportach nadgodzin** — i najbardziej kosztowny, bo plik wygląda poprawnie, liczby wyglądają wiarygodnie, a dane są nieprawdziwe. Nikt tego nie zauważy, dopóki ktoś nie zsumuje godzin ręcznie.

**Trzy warianty nawiasów:**

| Format | Znaczenie | Wynik dla 26,5 h |
|---|---|---|
| `[h]:mm:ss` | Godziny bez ograniczenia (minimalna liczba cyfr) | `26:30:00` |
| `[hh]:mm:ss` | Godziny, minimum dwie cyfry | `26:30:00` |
| `[h]:mm` | Bez sekund | `26:30` |
| `[m]:ss` | Minuty bez ograniczenia | `1590:00` |

Format `[m]:ss` bywa użyteczny, gdy mierzysz czasy krótkie, ale liczne — np. łączny czas obsługi klientów, który może przekroczyć 60 minut w jednej komórce.

**Przekonwertowanie na godziny dziesiętne.** Częsta potrzeba biznesowa: „26:30" to 26,5 godziny i chcesz to widzieć jako liczbę, żeby policzyć wynagrodzenie. Rozwiązanie: policz to samodzielnie w Pythonie, zamiast liczyć na magię formatu.

```python
from datetime import timedelta

def godziny_dziesietne(delta: timedelta) -> float:
    """Zamienia timedelta na godziny dziesietne (26:30 -> 26.5)."""
    return delta.total_seconds() / 3600


# Jesli masz wartosc z komorki (ułamek doby), przelicz tak:
def z_ulamka_doby(wartosc: float) -> float:
    """Ułamek doby -> godziny dziesietne."""
    return wartosc * 24
```

**To jest ważny wzorzec architektoniczny, który wróci w module 26:** jeśli biznes potrzebuje **liczby do dalszych obliczeń**, **policz ją w Pythonie i zapisz jako liczbę**. Format `[h]:mm:ss` jest dla **człowieka**, nie dla kolejnych formuł. Excel „rozumie" ten format, więc formuła `=SUM()` na komórkach z `[h]:mm:ss` zadziała — ale jeśli chcesz liczyć stawki, rób to na wartości dziesiętnej.

### 2.9. Jak openpyxl klasyfikuje komórkę: `data_type` i `is_date`

To sekcja diagnostyczna. Bez niej debugowanie formatów jest zgadywaniem.

#### 2.9.1. `cell.data_type` — pełna lista wartości

W źródle znajdziesz:

```python
TYPE_STRING = 's'
TYPE_FORMULA = 'f'
TYPE_NUMERIC = 'n'
TYPE_BOOL = 'b'
TYPE_NULL = 'n'
TYPE_INLINE = 'inlineStr'
TYPE_ERROR = 'e'
TYPE_FORMULA_CACHE_STRING = 'str'
```

| `data_type` | Znaczenie | Kiedy się pojawia |
|---|---|---|
| `'n'` | liczba | `int`, `float`, `Decimal`, **oraz pusta komórka** (bo `TYPE_NULL = 'n'`) |
| `'s'` | tekst (shared string) | `str`, `bytes`, `CellRichText` |
| `'d'` | data/czas | `datetime`, `date`, `time`, `timedelta` |
| `'b'` | wartość logiczna | `True` / `False` |
| `'f'` | formuła | tekst zaczynający się od `=` (dłuższy niż 1 znak) |
| `'e'` | błąd | jedna z wartości `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` |
| `'inlineStr'` | tekst wpisany w arkuszu | przy odczycie z plików zawierających stringi inline |
| `'str'` | tekst z pamięci podręcznej formuły | wartość tekstowa formuły zapisana przez Excel |

**Trzy fakty, które wynikają wprost z tego kodu:**

**Fakt 1: pusta komórka ma `data_type == 'n'`.** Ustawiasz `cell.value = None`, dostajesz typ liczbowy. To konsekwencja tego, że `TYPE_NULL` i `TYPE_NUMERIC` mają tę samą literę `'n'`. Konsekwencja praktyczna: **`cell.data_type == 'n'` nie znaczy „w komórce jest liczba"** — znaczy „komórka nie ma typu tekstowego ani błędu". Musisz sprawdzić wartość osobną asercją.

**Fakt 2: błędy Excela są wykrywane automatycznie z napisu.** Fragment kodu:

```python
elif value in ERROR_CODES:
    self.data_type = 'e'
```

Więc:

```python
ws["A1"] = "#N/A"
print(ws["A1"].data_type)     # 'e'
print(ws["A1"].value)         # #N/A
```

Wpisujesz **napis**, a dostajesz **typ błędu**. To wyjaśnia, co mówiliśmy w module 05: błędy czytane z pliku przychodzą jako napisy (bo `value` to nadal `str`), ale openpyxl je **klasyfikuje** jako błędy. Wykorzystaj to w walidacji: `if cell.data_type == 'e'` złapie wszystkie siedem kodów błędów.

**Fakt 3: `=` na początku napisu robi z niego formułę.** Fragment:

```python
if len(value) > 1 and value.startswith("="):
    self.data_type = 'f'
```

Zwróć uwagę na **`len(value) > 1`**. Samotny znak `=` (długość 1) **pozostaje tekstem**. To zachowanie jest celowe i ma osobny test w openpyxl (`test_not_formula`: `cell.value = "=" → data_type == 's'`). Praktyczna konsekwencja: chcąc wpisać tekst „= koszt", **musisz** ustawić `data_type` ręcznie albo poprzedzić tekst apostrofem — i to jest dokładnie zapowiedź modułu 21 (bezpieczeństwo, formula injection).

**Czwarty fakt — nieoczywisty:** gdy przypisujesz string, openpyxl wywołuje `check_string()`, która sprawdza **niedozwolone znaki kontrolne**:

```python
ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
```

Jeśli Twój tekst zawiera znak o kodzie 0x0B (pionowy tabulator), dostaniesz `IllegalCharacterError` przy zapisie. To najczęstsza przyczyna „nie wiem, dlaczego mój raport się nie zapisuje" przy danych importowanych z systemów zewnętrznych. Rozwiązaniem jest sanityzacja tekstu przed przypisaniem (moduł 21).

#### 2.9.2. `cell.is_date` — i jej dokładne znaczenie

`is_date` **istnieje w openpyxl 3.1.x** (nie zostało usunięte — mimo że pojawia się w changelogu hasło „Remove deprecated methods from Cell", dotyczy ono innych metod). Jej implementacja to dosłownie:

```python
@property
def is_date(self):
    """True if the value is formatted as a date

    :type: bool
    """
    return self.data_type == 'd' or (
        self.data_type == 'n' and is_date_format(self.number_format)
    )
```

Rozłóżmy to zdanie, bo jest kluczowe:

- Jeśli `data_type == 'd'` (czyli wartość to `datetime`/`date`/`time`/`timedelta`) → **`True`**.
- Jeśli `data_type == 'n'` (liczba) **i** kod formatu wygląda na datę (według heurystyki z 2.6.6) → **`True`**.
- W każdym innym przypadku → **`False`**.

**I teraz najważniejsza uwaga o `is_date`: ona nie patrzy na wartość.** Patrzy na **typ** i na **format**.

```python
ws["A1"] = 12345                                  # zwykla liczba
ws["A1"].number_format = "yyyy-mm-dd"             # format daty
print(ws["A1"].is_date)                           # True
print(ws["A1"].value)                             # 12345 - liczba, nie data
```

`is_date` mówi: „**gdyby Excel miał to narysować, narysowałby datę**". To stwierdzenie o reprezentacji, nie o wartości.

**Praktyczna rekomendacja:** `is_date` używaj **tylko w diagnostyce i logach**. Do decyzji o konwersji wartości używaj `isinstance(value, (datetime, date, time))` — to jednoznaczne i nie zależy od heurystyki formatu.

#### 2.9.3. Trzy funkcje diagnostyczne, które warto znać

```python
from openpyxl.styles.numbers import is_builtin, is_date_format, is_datetime, is_timedelta_format

print(is_date_format("yyyy-mm-dd"))       # True
print(is_date_format("0.00"))             # False
print(is_timedelta_format("[h]:mm:ss"))   # True
print(is_datetime("yyyy-mm-dd h:mm:ss"))  # 'datetime'
print(is_datetime("yyyy-mm-dd"))          # 'date'
print(is_datetime("h:mm"))                # 'time'
print(is_builtin("[h]:mm:ss"))            # True  (to format wbudowany, id 46)
```

`is_datetime` zwraca napis `'datetime'`, `'date'` albo `'time'` (albo `None`, jeśli to nie format daty). Zwróć uwagę: rozpoznaje po obecności liter `d`/`y` (data) i `h`/`s` (czas). `is_timedelta_format` sprawdza natomiast obecność nawiasów `[h]`, `[m]`, `[s]` — czyli dokładnie tego mechanizmu, o którym mówiliśmy w 2.8.

### 2.10. Pieniądze, `Decimal` i problem precyzji

#### 2.10.1. Dlaczego `float` to zły typ na pieniądze

Klasyczny przykład:

```python
print(0.1 + 0.2)          # 0.30000000000000004
print(0.1 + 0.2 == 0.3)   # False
```

`float` to liczba zmiennoprzecinkowa **binarna** — nie potrafi dokładnie zapisać `0.1` (tak jak nie potrafisz zapisać $1/3$ dokładnie w systemie dziesiętnym). Sumowanie tysięcy takich wartości prowadzi do groszy, które „znikają" albo się „pojawiają". W raporcie finansowym to niedopuszczalne.

**Analogia:** to jak odmierzać 0,1 litra **kubkiem wyskalowanym w trzecich częściach litra**. Za każdym razem masz odrobinę za mało albo za dużo, a po stu nalaniach różnica jest widoczna.

#### 2.10.2. Co robi openpyxl

`Decimal` **jest akceptowany** — występuje w `NUMERIC_TYPES`:

```python
from decimal import Decimal

from openpyxl import Workbook

wb = Workbook()
ws = wb.active

ws["A1"] = Decimal("19.99")
print(ws["A1"].data_type)     # 'n'
print(ws["A1"].value)         # 19.99
```

⚠️ **Ale**. Excel nie ma typu dziesiętnego. Na poziomie pliku wartość jest reprezentowana jako liczba tekstowa w XML, a **przy ponownym odczycie dostajesz `float`**. Oznacza to, że:

- **precyzja `Decimal` nie jest przez plik gwarantowana**,
- obieg `Decimal → plik → odczyt` zwraca `float`.

Dokładne zachowanie przy zapisie (`str(value)` kontra konwersja przez `float`) **sprawdź w dokumentacji swojej wersji openpyxl** — jeśli budujesz system, w którym precyzja groszy jest krytyczna, oprzyj się na własnej reprezentacji, a nie na typie komórki.

#### 2.10.3. Wzorzec, który naprawdę działa: grosze jako `int`

Jeśli prowadzisz obliczenia pieniężne, **przenieś arytmetykę do Pythona i pracuj na liczbach całkowitych**:

```python
"""Pieniadze bez bledow zaokraglenia - grosze jako int."""

from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class Kwota:
    """Kwota pieniezna przechowywana w groszach (int)."""
    grosze: int

    @classmethod
    def z_pln(cls, wartosc: str | Decimal) -> "Kwota":
        # Konwersja przez Decimal, z zaokragleniem do groszy
        return cls(int((Decimal(str(wartosc)) * 100).quantize(Decimal("1"))))

    @property
    def jako_pln(self) -> Decimal:
        return Decimal(self.grosze) / 100

    def __add__(self, inna: "Kwota") -> "Kwota":
        return Kwota(self.grosze + inna.grosze)


# Arytmetyka na int - zero bledow
suma = Kwota.z_pln("0.10") + Kwota.z_pln("0.20")
print(suma.grosze)         # 30
print(suma.jako_pln)       # 0.30
```

**Do komórki zapisujesz wartość dziesiętną** (bo tak musi wyglądać w Excelu), ale **liczysz na `int`** i dopiero na końcu konwertujesz:

```python
ws["B2"] = float(kwota.jako_pln)          # do Excela idzie liczba
ws["B2"].number_format = '#,##0.00 "zł"'  # wyswietlanie
```

To jest zapowiedź wzorca architektonicznego z modułu 26: **logika biznesowa (pieniądze) żyje w Pythonie; Excel jest warstwą prezentacji.** Nigdy nie pozwól, żeby Excel był jedynym miejscem, w którym liczą się pieniądze — bo wtedy nie masz testów, nie masz walidacji i nie wiesz, czy wynik jest poprawny.

### 2.11. Wzorzec: centralny rejestr formatów

To bezpośrednie rozwinięcie rejestru stylów z modułu 09. Poprzednio rejestr trzymał **pieczątki** (`Font`, `Fill`). Teraz dodajemy do niego **kody formatów** — i to jest naturalne miejsce, bo `number_format` jest szóstym elementem tego samego zestawu.

**Problem:** w kodzie raportu na 800 linii napiszesz `'#,##0.00 "zł"'` kilkanaście razy, a przy okazji: raz `'#,##0.00 zł'` (bez cudzysłowu), raz `'#,##0.00 "PLN"'`, raz `'# ##0,00 zł'` (polski przecinek, czyli błąd). Klient dostaje raport z **trzema wariantami tej samej waluty** i pyta, dlaczego kwoty wyglądają niespójnie.

**Rozwiązanie:**

```python
"""Centralny rejestr formatow - jedno zrodlo prawdy o reprezentacji."""

from __future__ import annotations

# ----------------------------------------------------------------------
# FORMATY - jedyne miejsce, gdzie pojawiaja sie kody formatow
# ----------------------------------------------------------------------
FORMAT = {
    # --- liczby ---
    "liczba":            "#,##0",
    "liczba_2":          "#,##0.00",
    # --- waluta ---
    "waluta_pln":        '#,##0.00 "zł"',
    "waluta_pln_ksieg":  '#,##0.00 "zł";(#,##0.00 "zł")',
    "waluta_pln_tys":    '#,##0, "tys. zł"',
    "waluta_eur":        '#,##0.00 "€"',
    # --- procenty ---
    "procent":           "0%",
    "procent_1":         "0.0%",
    "procent_2":         "0.00%",
    # --- daty i czas ---
    "data_iso":          "yyyy-mm-dd",
    "data_pl":           "[$-415]dd.mm.yyyy",
    "data_pl_dluga":     "[$-415]dd mmmm yyyy",
    "miesiac_rok":       "[$-415]mmmm yyyy",
    "data_czas":         "yyyy-mm-dd hh:mm",
    "godzina":           "hh:mm",
    "godzina_sek":       "hh:mm:ss",
    "czas_trwania":      "[h]:mm:ss",          # NIE zawija po 24 h!
    "czas_trwania_min":  "[h]:mm",
    # --- tekst i identyfikatory ---
    "tekst":             "@",
    "kod_pocztowy":      "@",
    "pesel_nip":         "@",
}


def ustaw_format(komorka, nazwa: str) -> None:
    """Ustawia format na komorce na podstawie nazwy z rejestru.

    Rzuca KeyError z lista dostepnych nazw, jesli nazwy nie ma.
    """
    if nazwa not in FORMAT:
        raise KeyError(f"nie ma formatu {nazwa!r}; dostepne: {sorted(FORMAT)}")
    komorka.number_format = FORMAT[nazwa]
```

**Użycie:**

```python
ustaw_format(ws["B2"], "waluta_pln")
ustaw_format(ws["C2"], "procent_1")
ustaw_format(ws["D2"], "data_pl")
ustaw_format(ws["E2"], "czas_trwania")
```

**Dlaczego to jest ważne — cztery powody:**

1. **Spójność.** Zmiana `FORMAT["waluta_pln"]` w jednym miejscu zmienia wszystkie kwoty w całym raporcie.
2. **Poprawność składni.** Kody formatów są trudne w pisaniu (cudzysłowy, nawiasy, LCID). Definiujesz je raz, a nie kilkanaście razy pod presją czasu.
3. **Testowalność.** Możesz napisać test, który dla każdego formatu w rejestrze sprawdza, że openpyxl go akceptuje i że `is_date_format` daje oczekiwany wynik. To nie jest możliwe, gdy formaty są rozsypane. (Testy — moduł 20.)
4. **Bezpieczeństwo heurystyki.** Ponieważ wszystkie kody są w jednym miejscu, możesz przejrzeć je świadomie i upewnić się, że każdy tekst jest w cudzysłowie (pułapka `#,##0 szt.` z 2.6.6 nie przejdzie przez code review).

**W modułach 23 i 26** ten słownik stanie się elementem konfiguracji raportu czytanej z pliku YAML/TOML. Ale uwaga — granica: **konfiguracja opisuje „co" (jak wygląda kwota), kod opisuje „jak" (jak nałożyć format)**. Słownik to konfiguracja; funkcja `ustaw_format` to kod.

## 3. Przykłady krok po kroku

### Przykład 1 — pięć formatów na jednej liczbie: dowód, że wartość i reprezentacja są rozłączne

Zaczynamy od eksperymentu, który jest sedno tego modułu w jednym pliku.

```python
"""JEDNA wartosc, PIEC roznych reprezentacji.

Uruchom:  python examples/10_wartosc_vs_reprezentacja.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "10_wartosc_vs_reprezentacja.xlsx"

WARTOSC = 1234.5678

FORMATY = [
    ("General",                 "General"),
    ("liczba calkowita",        "0"),
    ("dwa miejsca",             "0.00"),
    ("tysiace z groszami",      "#,##0.00"),
    ("waluta PLN",              '#,##0.00 "zł"'),
    ("procent",                 "0.0%           <- UWAGA!"),
    ("w tysiacach",             '#,##0, "tys."'),
    ("notacja naukowa",         "0.00E+00"),
    ("czas trwania",            "[h]:mm:ss"),
    ("tekst (nie ma sensu)",    "@"),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Formaty"

    ws["A1"] = "opis"
    ws["B1"] = "kod formatu"
    ws["C1"] = "wartość (identyczna!)"
    ws["D1"] = "co widzi człowiek"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22

    for indeks, (opis, kod) in enumerate(FORMATY, start=2):
        ws.cell(row=indeks, column=1, value=opis)
        ws.cell(row=indeks, column=2, value=kod)
        komorka = ws.cell(row=indeks, column=3, value=WARTOSC)
        komorka.number_format = kod

    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Formaty"]
        print("=" * 84)
        print("CO NAPRAWDE JEST W PLIKU (po zapisie i odczycie)")
        print("=" * 84)
        print(f"  {'opis':<22} {'format':<26} {'value':<12} {'data_type'} {'is_date'}")
        print(f"  {'-'*22} {'-'*26} {'-'*12} {'-'*9} {'-'*7}")

        for wiersz in range(2, len(FORMATY) + 2):
            opis = ws.cell(row=wiersz, column=1).value
            komorka = ws.cell(row=wiersz, column=3)
            print(
                f"  {opis:<22} {komorka.number_format:<26} "
                f"{komorka.value!r:<12} {komorka.data_type:<9} {komorka.is_date}"
            )
    finally:
        wb.close()

    print()
    print("=" * 84)
    print("WNIOSEK")
    print("=" * 84)
    print("  Kolumna 'value' jest IDENTYCZNA w kazdym wierszu: 1234.5678.")
    print("  Rozni sie TYLKO format - czyli to, co zobaczy czlowiek.")
    print()
    print("  Formatowanie NIE zmienia wartosci. Nie zaokragla. Nie przelicza.")
    print("  Kolejne formuly w Excelu beda liczyc na 1234.5678, nie na tym,")
    print("  co widac w kolumnie D.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Kluczowe fragmenty do przemyślenia:**

1. **Kolumna `value` w weryfikacji jest identyczna w każdym wierszu.** To jest dowód. Nie ma znaczenia, który format wybierzesz — liczba w pliku to zawsze `1234.5678`.
2. **Wiersz „procent"** to celowy przykład pułapki. Wartość `1234.5678` z formatem `0.0%` wyświetli się jako `123456,8%` — bo Excel pomnoży przez 100. Chciałeś `1234,6%`? Wtedy w komórce powinno być `12.345678`. To dokładnie ta pułapka, o której mówimy w ćwiczeniu 🟡.
3. **Wiersz „tekst"** pokazuje, że format `@` na liczbie nie robi nic sensownego — bo liczba nie jest tekstem. Czwarta sekcja formatu (`@`) dotyczy wyłącznie wartości tekstowych.
4. **Kolumna `is_date`** pokaże `False` dla wszystkich oprócz wiersza `czas trwania` — tam `[h]:mm:ss` **zawiera `h` i `s`**, więc heurystyka `is_date_format` zwróci `True`. Świetna ilustracja tego, że `is_date` jest heurystyką, a nie prawdą objawioną.

**Co trafi do pliku.** W `xl/numFmts` znajdziesz `<numFmt numFmtId="164" .../>` dla pierwszego **nie**wbudowanego formatu i kolejne numery dla następnych. Format `General` zostanie zapisany jako `numFmtId="0"` (wbudowany), `0` jako `numFmtId="1"`, `0.00` jako `2`, `#,##0.00` jako `4`, `0.0%` jako `9`, `0.00E+00` jako `11`, `@` jako `49`. Tylko trzy formaty (`#,##0.00 "zł"`, `#,##0, "tys."`, `[h]:mm:ss`) — a właściwie dwa, bo `[h]:mm:ss` też jest wbudowany (id 46) — dostaną własne identyfikatory ≥ 164.

**Zajrzyj do `styles.xml` po uruchomieniu tego przykładu.** Zobaczysz na własne oczy różnicę między „format wbudowany, jeden bajt" a „format własny, osobny wpis".

### Przykład 2 — anty-przykład procentu i jego naprawa

Zbudujmy świadomie błędny raport, a potem go naprawmy.

```python
"""Procent: 0.05 kontra 5 - dwa rozne swiaty.

Uruchom:  python examples/10_procent.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "10_procent.xlsx"

# Male dane: marza jako ulamek i jako wartosc "procentowa"
DANE_ULAMKI = [
    ("styczeń", 120_000, 0.35),      # 0.35 = 35%
    ("luty",     98_500, 0.28),
    ("marzec",  143_700, 0.15),
    ("kwiecień", 131_200, 0.42),
]


def zbuduj_bledny(cel: Path) -> None:
    """Wersja BLEDNA: wartosc jest juz w procentach, a format mnozy przez 100."""
    wb = Workbook()
    ws = wb.active
    ws.title = "BŁĄD"

    ws.append(["Miesiąc", "Sprzedaż", "Marża"])
    for miesiac, sprzedaz, marza in DANE_ULAMKI:
        ws.append([miesiac, sprzedaz, marza * 100])      # <- 35 zamiast 0.35

    for wiersz in range(2, 6):
        ws.cell(row=wiersz, column=3).number_format = "0.0%"

    ws["C7"] = "=SUM(C2:C5)"
    ws["C7"].number_format = "0.0%"

    wb.save(cel)
    wb.close()


def zbuduj_poprawny(cel: Path) -> None:
    """Wersja POPRAWNA: wartosc jako ulamek, format robi reszte."""
    wb = Workbook()
    ws = wb.active
    ws.title = "POPRAWNIE"

    ws.append(["Miesiąc", "Sprzedaż", "Marża"])
    for miesiac, sprzedaz, marza in DANE_ULAMKI:
        ws.append([miesiac, sprzedaz, marza])            # <- 0.35

    for wiersz in range(2, 6):
        ws.cell(row=wiersz, column=3).number_format = "0.0%"

    ws["C7"] = "=SUM(C2:C5)"
    ws["C7"].number_format = "0.0%"

    wb.save(cel)
    wb.close()


def main() -> None:
    print("=" * 88)
    print("PROCENT: dwa swiaty zapisu")
    print("=" * 88)
    print()
    print("  WARIANT A (BLEDNY): do komorki idzie 35, format 0.0%")
    print("    Excel wyswietli:  3500,0%")
    print("    Bo format '0.0%' sam mnozy wartosc przez 100.")
    print()
    print("  WARIANT B (POPRAWNY): do komorki idzie 0.35, format 0.0%")
    print("    Excel wyswietli:  35,0%")
    print("    Bo to format odpowiada za znak % i za pomnozenie.")
    print()
    print("  Format '0.0%' to instrukcja: 'wez wartosc, pomnóz przez 100,")
    print("  dopisz znak procenta'. Wiec wartosc musi byc ULAMKIEM.")
    print()

    zbuduj_bledny(OUTPUT / "10_procent_bledny.xlsx")
    zbuduj_poprawny(OUTPUT / "10_procent_poprawny.xlsx")
    print(f"  Zapisano: {OUTPUT / '10_procent_bledny.xlsx'}")
    print(f"  Zapisano: {OUTPUT / '10_procent_poprawny.xlsx'}")
    print()
    print("  Otwórz oba pliki w Excelu i porownaj kolumne 'Marża' oraz sume.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W wariancie błędnym w komórkach siedzą liczby `35`, `28`, `15`, `42`. Format `0.0%` mnoży je przez 100 przy rysowaniu. W wariancie poprawnym w komórkach siedzą `0.35`, `0.28` itd., a format daje `35,0%`.

**Co trafi do pliku.** Ten sam kod formatu w obu przypadkach (`0.0%`, wbudowany, `numFmtId="9"`). Różnią się **wartości w `<v>`**. To jest kolejny dowód na rozłączenie dwóch światów.

**Czego ten przykład uczy najbardziej:** że format `%` **nie jest dekoracją** — jest **operatorem**. Zmienia to, co widzi człowiek, w sposób mnożący. Dlatego decyzja „co jest w komórce: 0.35 czy 35" to decyzja **biznesowa i architektoniczna**, nie kosmetyczna. I dlatego **każdy raport finansowy powinien mieć jeden, ustalony standard** — albo wszystkie procenty jako ułamki, albo wszystkie jako liczby, ale nigdy mieszanka.

### Przykład 3 — cztery sekcje w praktyce: raport z ujemnymi, zerami i tekstem

```python
"""Cztery sekcje kodu formatu: dodatnie; ujemne; zero; tekst.

Uruchom:  python examples/10_cztery_sekcje.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "10_cztery_sekcje.xlsx"

# Wartosci testowe: dodatnia, ujemna, zero, tekst, pusta
WARTOSCI = [
    ("dodatnia", 1234.5678),
    ("ujemna", -1234.5678),
    ("zero", 0),
    ("tekst", "brak danych"),
    ("pusta", None),
]

# Trzy warianty tego samego zestawu sekcji
FORMATY = [
    (
        "1 sekcja (domyslna dla liczb)",
        "#,##0.00",
    ),
    (
        "2 sekcje (dodatnie ; ujemne)",
        '#,##0.00 "zł";(#,##0.00 "zł")',
    ),
    (
        "3 sekcje (dodatnie ; ujemne ; zero)",
        '#,##0.00 "zł";(#,##0.00 "zł");"—"',
    ),
    (
        "4 sekcje (dodatnie ; ujemne ; zero ; tekst)",
        '#,##0.00 "zł";(#,##0.00 "zł");"—";@" (do weryfikacji)"',
    ),
    (
        "kolor w sekcji ujemnej",
        '#,##0.00 "zł";[Red](#,##0.00 "zł")',
    ),
]


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cztery sekcje"

    # Naglowek
    ws["A1"] = "typ wartości"
    ws["B1"] = "wartość w komórce"
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 20

    # Kolumny z wariantami formatow
    for indeks, (opis, kod) in enumerate(FORMATY, start=3):
        komorka_opis = ws.cell(row=1, column=indeks, value=opis)
        komorka_opis.alignment = komorka_opis.alignment.copy(wrap_text=True)
        ws.cell(row=2, column=indeks, value=kod)
        ws.column_dimensions[komorka_opis.column_letter].width = 30

    ws.row_dimensions[1].height = 34

    # Dane
    for przesuniecie, (typ, wartosc) in enumerate(WARTOSCI):
        wiersz = 3 + przesuniecie
        ws.cell(row=wiersz, column=1, value=typ)
        ws.cell(row=wiersz, column=2, value=wartosc)

        for kolumna, (_, kod) in enumerate(FORMATY, start=3):
            komorka = ws.cell(row=wiersz, column=kolumna, value=wartosc)
            komorka.number_format = kod

    # Kluczowa uwaga zapisana w pliku, zeby odbiorca wiedzial, co widzi.
    ws["B10"] = "UWAGA: komorki w wierszu 'pusta' nie maja wartosci -"
    ws["B11"] = "format nie tworzy danych z niczego. Pusta komorka to pusta komorka."

    wb.save(CEL)
    print(f"Zapisano: {CEL}")
    print()
    print("=" * 84)
    print("CO ZOBACZYSZ W EXCELU")
    print("=" * 84)
    print("  wiersz 'dodatnia'   ->  1 234,57 zł")
    print("  wiersz 'ujemna'     ->  (1 234,57 zł)  albo czerwone -1 234,57 zł")
    print("  wiersz 'zero'       ->  —    (tylko w wersji z 3. sekcja)")
    print("  wiersz 'tekst'      ->  brak danych (do weryfikacji)")
    print("                         ^ tylko w wersji z 4. sekcja (@)")
    print("  wiersz 'pusta'      ->  pusto - format nie ma na co dzialac")
    print()
    print("  Zwroc uwage: w wersji '1 sekcja' tekst i zero pokaza sie")
    print("  zwyczajnie. Kazda kolejna sekcja dodaje regule dla kolejnego")
    print("  przypadku. Dlatego wlasnie sekcje sa ODDZIELONE SREDNIKIEM.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W kolumnach C–G są **te same pięć wartości**. Różnią się tylko kody formatów. To jeszcze raz ilustruje rozdział wartości od reprezentacji, ale tym razem w wymiarze „jak obsłużyć przypadki brzegowe".

**Co trafi do pliku.** Każdy z czterech niepowtarzalnych kodów trafi do `xl/styles.xml` jako osobny `<numFmt>` (z wyjątkiem `#,##0.00`, który jest wbudowany jako id 4). Zauważ, że formaty z 3. i 4. sekcją są **dłuższe**, ale w pliku różnią się tylko treścią `formatCode`.

**Czego ten przykład uczy najbardziej:** czwarta sekcja (`@`) jest jedynym sposobem na sformatowanie **tekstu** w komórce. Trzecia sekcja (zero) jest jedynym sposobem na pokazanie zera jako myślnika. Oba są powszechnie używane w sprawozdaniach finansowych, gdzie „0,00 zł" wygląda jak błąd, a „—" wygląda jak profesjonalizm.

I jeszcze jedna lekcja, ważna: **wiersz „pusta"**. Format nie tworzy danych. Pusta komórka pozostanie pusta bez względu na to, jaki format jej ustawisz.

### Przykład 4 — nadgodziny: czasy trwania powyżej 24 godzin

To ćwiczenie 🔴 z briefu, wykonane porządnie.

```python
"""Nadgodziny: czasy trwania powyzej 24 godzin.

Uruchom:  python examples/10_nadgodziny.py
"""

from __future__ import annotations

from datetime import timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "10_nadgodziny.xlsx"

# (pracownik, pon, wt, sr, czw, pt) - kazdy dzien jako timedelta
EWIDENCJA = [
    ("Kowalska Anna",  timedelta(hours=2, minutes=15), timedelta(hours=3),
     timedelta(hours=1, minutes=45), timedelta(0), timedelta(hours=2, minutes=30)),
    ("Nowak Piotr",    timedelta(hours=4), timedelta(hours=2, minutes=30),
     timedelta(hours=5, minutes=15), timedelta(hours=1), timedelta(hours=0, minutes=45)),
    ("Wiśniewski Jan", timedelta(hours=8, minutes=30), timedelta(hours=6),
     timedelta(hours=7, minutes=15), timedelta(hours=9, minutes=30),
     timedelta(hours=5)),
    ("Zielińska Ewa",  timedelta(hours=1), timedelta(hours=1, minutes=30),
     timedelta(hours=2), timedelta(hours=2, minutes=15), timedelta(hours=3, minutes=45)),
]

DNI = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek"]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Nadgodziny"

    # --- naglowek ------------------------------------------------------
    ws["A1"] = "Ewidencja nadgodzin — marzec 2026"
    ws.merge_cells("A1:G1")
    ws["B1"].alignment = ws["B1"].alignment.copy(horizontal="left", vertical="center")
    ws["B1"].font = ws["B1"].font.copy(size=14, bold=True)
    ws.row_dimensions[1].height = 24

    naglowki = ["Pracownik", *DNI, "Razem"]
    for indeks, tekst in enumerate(naglowki, start=1):
        komorka = ws.cell(row=3, column=indeks, value=tekst)
        komorka.font = komorka.font.copy(bold=True)

    # --- dane ----------------------------------------------------------
    for przesuniecie, (pracownik, *dni) in enumerate(EWIDENCJA):
        wiersz = 4 + przesuniecie
        ws.cell(row=wiersz, column=1, value=pracownik)

        for indeks, delta in enumerate(dni):
            komorka = ws.cell(row=wiersz, column=2 + indeks, value=delta)
            # KRYTYCZNE: nawiasy kwadratowe = "nie zawijaj na dobie"
            komorka.number_format = "[h]:mm"

        # Suma tygodniowa jako formula (Excel ja policzy przy otwarciu)
        kolumna_sumy = 2 + len(dni)
        suma = ws.cell(
            row=wiersz,
            column=kolumna_sumy,
            value=f"=SUM(B{wiersz}:F{wiersz})",
        )
        suma.number_format = "[h]:mm"
        suma.font = suma.font.copy(bold=True)

    # --- wiersz podsumowania -------------------------------------------
    wiersz_sumy = 4 + len(EWIDENCJA)
    ws.cell(row=wiersz_sumy, column=1, value="RAZEM").font = \
        ws.cell(row=wiersz_sumy, column=1).font.copy(bold=True)

    for kolumna in range(2, 8):
        litera = ws.cell(row=3, column=kolumna).column_letter
        komorka = ws.cell(
            row=wiersz_sumy,
            column=kolumna,
            value=f"=SUM({litera}4:{litera}{wiersz_sumy - 1})",
        )
        komorka.number_format = "[h]:mm"
        komorka.font = komorka.font.copy(bold=True)

    # --- kolumna kontrolna: godziny dziesietne (policzone w Pythonie) ---
    # To NIE jest formula - to zwykla liczba, ktora policzylismy sami.
    ws.cell(row=3, column=8, value="Razem (godz. dzies.)").font = \
        ws.cell(row=3, column=8).font.copy(bold=True)

    for przesuniecie, (_, *dni) in enumerate(EWIDENCJA):
        wiersz = 4 + przesuniecie
        total = sum(dni, timedelta(0))
        komorka = ws.cell(
            row=wiersz,
            column=8,
            value=round(total.total_seconds() / 3600, 2),
        )
        komorka.number_format = "0.00"

    for kolumna, szerokosc in zip("ABCDEFGH", (20, 13, 13, 13, 13, 13, 13, 20)):
        ws.column_dimensions[kolumna].width = szerokosc

    ws.freeze_panes = "B4"
    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    """Weryfikacja: czy format [h]:mm naprawde jest w pliku."""
    wb = load_workbook(cel)
    try:
        ws = wb["Nadgodziny"]
        print("=" * 88)
        print("WERYFIKACJA PO ZAPISIE I ODCZYCIE")
        print("=" * 88)
        print(f"  {'komorka':<9} {'value':<28} {'format':<12} {'data_type'} is_date")
        print(f"  {'-'*9} {'-'*28} {'-'*12} {'-'*9} {'-'*7}")

        for adres in ("B4", "C4", "G4", "B7", "G7", "H4"):
            komorka = ws[adres]
            print(
                f"  {adres:<9} {str(komorka.value):<28} "
                f"{komorka.number_format:<12} {komorka.data_type:<9} {komorka.is_date}"
            )

        print()
        print("  B4  = 2:15 nadgodzin Kowalskiej w poniedziałek")
        print("  B7  = 8:30 nadgodzin Wiśniewskiego w poniedziałek")
        print("  G4  = suma tygodniowa (formula) - Excel policzy przy otwarciu")
        print("  H4  = 9.5 (godziny dziesietne - policzone w Pythonie)")
    finally:
        wb.close()

    print()
    print("=" * 88)
    print("NAJWAZNIEJSZY PUNKT")
    print("=" * 88)
    print("  Wszystkie komorki z czasem maja format '[h]:mm' - z NAWIASAMI.")
    print("  Bez nich Excel pokazalby np. 26 godzin jako 02:00 (zawiniete po dobie).")
    print()
    print("  Zmien format w B4 na 'h:mm' i porownaj - zobaczysz roznice")
    print("  dopiero przy wartosciach przekraczajacych 24 h (np. suma tygodniowa).")
    print()
    print("  Kolumna H to wersja 'dla obliczen': liczby dziesietne, ktore mozna")
    print("  od razu pomnozyc przez stawke godzinowa.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Każda `timedelta` jest konwertowana przez openpyxl na ułamek doby. `timedelta(hours=26, minutes=30)` → `1.104166...`. W komórce suma tygodniowa dla Wiśniewskiego to `timedelta(hours=36, minutes=15)` → ponad 1,5 dnia.

**Co trafi do pliku.** Do `<v>` trafiają liczby (ułamki doby), a kody formatów `[h]:mm` do `styles.xml` jako **jeden wspólny wpis** — mimo że użyto ich w 30 komórkach. To deduplikacja z modułu 09 w praktyce.

**Czego ten przykład uczy najbardziej:** że `[h]:mm` i `h:mm` to **nie dwa warianty estetyczne jednego formatu** — to dwa różne mechanizmy o różnych wynikach. I że w realnym raporcie potrzebujesz **obu reprezentacji naraz**: `[h]:mm` dla człowieka (bo tak myśli o nadgodzinach) i liczby dziesiętnej dla dalszych obliczeń (bo tak liczy się wynagrodzenie).

**Uwaga o `timedelta` i `data_type`.** W wydruku weryfikacji zobaczysz, co openpyxl uznaje za `data_type` dla `timedelta`. Wewnętrzna klasyfikacja `timedelta` jest nieoczywista — w repozytorium openpyxl istnieje test na ten temat oznaczony jako `xfail` (czyli „zachowanie odbiega od oczekiwanego"). **Nie opieraj logiki na `data_type` przy `timedelta`** — przy `timedelta` sprawdzaj `is_timedelta_format(cell.number_format)` albo po prostu `isinstance` wartości. **Sprawdź też w swojej wersji openpyxl**, jak dokładnie zachowuje się `timedelta` przy odczycie.

### Przykład 5 — diagnostyka formatów w pliku i test rejestru

Domykamy moduł weryfikacją na poziomie pakietu `.xlsx`.

```python
"""Diagnostyka formatow liczb w pliku .xlsx.

Uruchom:  python examples/10_diagnostyka_formatow.py
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# Kody wbudowane, ktore chcemy rozpoznawac w diagnostyce
WBUDOWANE_NAZWY = {
    "0": "General", "1": "0", "2": "0.00", "3": "#,##0", "4": "#,##0.00",
    "9": "0%", "10": "0.00%", "11": "0.00E+00", "12": "# ?/?", "13": "# ??/??",
    "14": "mm-dd-yy", "15": "d-mmm-yy", "16": "d-mmm", "17": "mmm-yy",
    "18": "h:mm AM/PM", "19": "h:mm:ss AM/PM", "20": "h:mm", "21": "h:mm:ss",
    "22": "m/d/yy h:mm", "37": "#,##0_);(#,##0)", "38": "#,##0_);[Red](#,##0)",
    "39": "#,##0.00_);(#,##0.00)", "40": "#,##0.00_);[Red](#,##0.00)",
    "45": "mm:ss", "46": "[h]:mm:ss", "47": "mmss.0", "48": "##0.0E+0", "49": "@",
}

BUILTIN_MAX = 164


def czytaj_formats(sciezka: Path) -> dict:
    """Zwraca informacje o formatach liczb w pliku .xlsx."""
    with zipfile.ZipFile(sciezka) as archiwum:
        try:
            xml = archiwum.read("xl/styles.xml")
        except KeyError:
            return {"blad": "brak xl/styles.xml"}

    korzen = ElementTree.fromstring(xml)

    # Wlasne formaty (numFmtId >= 164)
    wlasne: list[tuple[str, str]] = []
    sekcja = korzen.find("m:numFmts", NS)
    if sekcja is not None:
        for element in sekcja:
            num_id = element.get("numFmtId")
            kod = element.get("formatCode")
            wlasne.append((num_id, kod))

    # Ile stylow uzywa formatow nie-General?
    cell_xfs = korzen.find("m:cellXfs", NS)
    uzyte_ids: dict[str, int] = {}
    if cell_xfs is not None:
        for xf in cell_xfs:
            num_fmt_id = xf.get("numFmtId", "0")
            if num_fmt_id != "0":
                uzyte_ids[num_fmt_id] = uzyte_ids.get(num_fmt_id, 0) + 1

    return {
        "wlasne_formaty": wlasne,
        "liczba_wlasnych": len(wlasne),
        "uzyte_numFmtId": uzyte_ids,
        "liczba_zestawow_stylu": len(list(cell_xfs)) if cell_xfs is not None else 0,
    }


def opis_id(num_fmt_id: str) -> str:
    """Zwraca czytelny opis identyfikatora formatu."""
    if num_fmt_id in WBUDOWANE_NAZWY:
        return f"WBUDOWANY: {WBUDOWANE_NAZWY[num_fmt_id]}"
    if num_fmt_id.isdigit() and int(num_fmt_id) >= BUILTIN_MAX:
        return "WLASNY (>=164)"
    return "nieznany / zarezerwowany"


def main() -> None:
    print("=" * 92)
    print("DIAGNOSTYKA FORMATOW LICZB W PLIKACH .XLSX")
    print("=" * 92)
    print()

    pliki = sorted(OUTPUT.glob("10_*.xlsx"))
    if not pliki:
        print("  Brak plikow. Uruchom najpierw inne przyklady z tego modulu.")
        return

    for plik in pliki:
        dane = czytaj_formats(plik)
        print(f"  --- {plik.name} ---")

        if "blad" in dane:
            print(f"      {dane['blad']}")
            print()
            continue

        print(f"      liczba zestawow stylu (cellXfs): {dane['liczba_zestawow_stylu']}")
        print(f"      liczba wlasnych formatow (>=164): {dane['liczba_wlasnych']}")

        if dane["wlasne_formaty"]:
            print("      wlasne formaty:")
            for num_id, kod in dane["wlasne_formaty"]:
                print(f"        numFmtId={num_id:<5} formatCode={kod!r}")
        else:
            print("      wlasne formaty: brak (wszystkie uzyte sa wbudowane)")

        if dane["uzyte_numFmtId"]:
            print("      identyfikatory uzyte przez style:")
            for num_id, ile in sorted(dane["uzyte_numFmtId"].items(), key=lambda x: int(x[0])):
                print(f"        numFmtId={num_id:<5} ({ile}x)  -> {opis_id(num_id)}")
        print()

    print("=" * 92)
    print("JAK TO CZYTAC")
    print("=" * 92)
    print("  * Format 'General' ma numFmtId=0 i NIE jest zapisywany jako <numFmt>.")
    print("  * Identi 0-163 sa ZAREZERWOWANE dla formatow wbudowanych Excela.")
    print("  * Twoje wlasne kody startuja od 164 (BUILTIN_FORMATS_MAX_SIZE).")
    print("  * Jeden kod formatu uzyty w 1000 komorek = JEDEN wpis w <numFmts>.")
    print("    Deduplikacja dziala tak samo jak przy stylach (modul 09).")
    print()
    print("  Jesli w raporcie widzisz TRZY rozne kody dla waluty - to znaczy,")
    print("  ze ktos nie uzywal centralnego rejestru formatow (sekcja 2.11).")


if __name__ == "__main__":
    main()
```

**Czego ten przykład uczy najbardziej:** że **format liczby jest sprawdzalny na poziomie pliku**, dokładnie tak samo jak style z modułu 09. Potrafisz policzyć, ile unikalnych kodów formatów naprawdę jest w Twoim raporcie — i jeśli odpowiedź brzmi „jedenaście", a Ty pamiętasz, że definiowałeś cztery, to właśnie znalazłeś problem.

To umiejętność, która w module 18 (modyfikacja istniejących plików) stanie się podstawą audytu „co przechodzi, a co ginie".

## 4. Anatomia API

| Element | Co robi | Parametry / wartość | Uwagi |
|---|---|---|---|
| `cell.number_format` | Kod formatu **reprezentacji** | `str`, np. `'#,##0.00 "zł"'` | **Nie walidowany!** `None` → `"General"` |
| `cell.value` | Wartość w komórce | `int`, `float`, `Decimal`, `str`, `bool`, `datetime`, `date`, `time`, `timedelta`, `None` | Format **nie zmienia** wartości |
| `cell.data_type` | Klasyfikacja zawartości | `'n'`, `'s'`, `'d'`, `'b'`, `'f'`, `'e'`, `'inlineStr'`, `'str'` | Pusta komórka → `'n'`; błędy Excela → `'e'` |
| `cell.is_date` | Czy komórka „jest datą" | `bool` | **Istnieje w 3.1.x.** Nie patrzy na wartość, tylko na typ i format |
| `cell.base_date` | Data bazowa dla seriali | `datetime` | Bierze epokę ze skoroszytu; sprawdź w swojej wersji |
| `wb.epoch` | Epoka skoroszytu | `datetime` | `1899-12-30` (Windows) lub `1904-01-01` (Mac) |
| `openpyxl.utils.datetime.to_excel(dt, epoch=...)` | Data Pythona → serial Excela | `datetime`/`date` | ⚠️ Sygnatura zmieniała się między wersjami — sprawdź |
| `openpyxl.utils.datetime.from_excel(v, epoch=...)` | Serial → `datetime` | liczba | Jak wyżej |
| `CALENDAR_WINDOWS_1900`, `CALENDAR_MAC_1904` | Stałe epok | `datetime` | `openpyxl.utils.datetime` |
| `openpyxl.styles.numbers.BUILTIN_FORMATS` | Słownik formatów wbudowanych | `dict[int, str]` | Dokładna kopia tego, co ma Excel |
| `BUILTIN_FORMATS_MAX_SIZE` | Granica własnych formatów | `164` | Identyfikatory ≥ 164 to Twoje formaty |
| `is_date_format(fmt)` | Heurystyka „czy to format daty" | `str` → `bool` | Patrzy tylko na **pierwszą sekcję**; usuwa teksty i `[locale]`; szuka `[dmhys]` |
| `is_timedelta_format(fmt)` | Czy format to czas trwania | `str` → `bool` | Szuka `[h]`, `[m]`, `[s]` |
| `is_datetime(fmt)` | Klasyfikuje format | `str` → `'datetime'` / `'date'` / `'time'` / `None` | Opiera się na obecności `d`/`y` (data) i `h`/`s` (czas) |
| `is_builtin(fmt)` | Czy kod jest wbudowany | `str` → `bool` | |
| `builtin_format_id(fmt)` | Identyfikator formatu wbudowanego | `str` → `int` / `None` | |
| `builtin_format_code(index)` | Kod formatu wbudowanego po ID | `int` → `str` / `None` | |
| `FORMAT_TEXT`, `FORMAT_PERCENTAGE`, `FORMAT_DATE_DATETIME`, `FORMAT_DATE_TIMEDELTA`, … | Gotowe stałe kodów | `str` | `openpyxl.styles.numbers` |
| `get_time_format(t)` | Format domyślny dla typu | typ Pythona → `str` | Używane wewnętrznie przy przypisywaniu wartości |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — pięć formatów, jedna wartość (z weryfikacją).**

Napisz `examples/10_cw1.py`, który:

1. Tworzy skoroszyt z arkuszem `Formaty`.
2. W komórkach `B2:B6` umieszcza **pięć razy tę samą wartość**: `1234.5678`.
3. Nadaje im **pięć różnych formatów**, po jednym na komórkę:
   - `'#,##0'`,
   - `'#,##0.00'`,
   - `'#,##0.00 "zł"'`,
   - `'#,##0, "tys."'`,
   - `'0.00E+00'`.
4. W kolumnie `A` wpisuje **opis** każdego formatu (po polsku), a w kolumnie `C` — **kod formatu**.
5. Zapisuje plik jako `output/10_cw1.xlsx`.
6. **Wczytuje plik ponownie** i wypisuje dla każdej komórki: `value`, `number_format`, `data_type`, `is_date`.

W pliku `output/10_cw1_wnioski.md` odpowiedz:

1. **Czy kolumna `value` jest identyczna we wszystkich pięciu komórkach?** Jeśli tak — co to dowodzi?
2. **Które z użytych formatów są wbudowane, a które nie?** Sprawdź funkcją `is_builtin()` i wypisz wynik. Dla wbudowanych podaj `builtin_format_id()`.
3. **Czym różni się `'#,##0.00 "zł"'` od `'#,##0.00 zł'`** (bez cudzysłowu)? Sprawdź oba kody funkcją `is_date_format()` i wyjaśnij wynik w kontekście pułapki z sekcji 2.6.6.
4. **Co pokaże format `'#,##0, "tys."'`?** Opisz słownie i uzasadnij, gdzie wartość zostałaby podzielona.

**Zadanie 2 — jedna liczba, różne prawdy: procent.**

Napisz `examples/10_cw2.py`, który tworzy **dwa** pliki:

- `output/10_cw2_bledny.xlsx` — w komórce `A1` jest `35`, format `"0.0%"`,
- `output/10_cw2_poprawny.xlsx` — w komórce `A1` jest `0.35`, format `"0.0%"`.

Otwórz oba w Excelu i opisz w `output/10_cw2_wnioski.md`:

1. Co wyświetla każdy plik?
2. **Dlaczego wartość `35` z formatem `0.0%` daje `3500,0%`?** Wyjaśnij mechanizm, nie tylko fakt.
3. Masz w arkuszu 40 kolumn z marżami. Które podejście wybierzesz i dlaczego? Podaj **dwa** argumenty (jeden o formułach, jeden o testach).
4. Jak **wykryłbyś** w cudzym pliku, że ktoś pomieszał oba podejścia? Opisz procedurę (podpowiedź: formaty liczb + zakresy wartości + `value`).

### 🟡 Warsztat

**Zadanie 3 — rejestr formatów z testem.**

Rozbuduj rejestr z sekcji 2.11 do wersji z testem automatycznym.

**Część A — rejestr.** Napisz `examples/10_cw3.py`, który zawiera:

1. Słownik `FORMAT` z **co najmniej piętnastoma** wpisami, w tym obowiązkowo:
   - `"waluta_pln"` — `'#,##0.00 "zł"'`,
   - `"waluta_pln_ksieg"` — z sekcją dla ujemnych w nawiasie,
   - `"waluta_pln_tys"` — dzielenie przez tysiąc (`#,##0,`),
   - `"procent"`, `"procent_1"`, `"procent_2"`,
   - `"data_iso"`, `"data_pl"`, `"data_pl_dluga"` (z LCID 415), `"data_czas"`,
   - `"godzina"`, `"godzina_sek"`, `"czas_trwania"`, `"czas_trwania_min"`,
   - `"tekst"`, `"kod_pocztowy"`, `"pesel_nip"`.
2. Funkcję `ustaw_format(komorka, nazwa)` — rzucającą `KeyError` z listą dostępnych nazw.
3. Funkcję `ustaw_format_do_zakresu(ws, zakres, nazwa)` — używającą `range_boundaries` z `openpyxl.utils.cell` i zwracającą liczbę sformatowanych komórek.
4. Skrypt demonstracyjny budujący arkusz:
   - kolumna z kwotami (`waluta_pln`),
   - kolumna z marżami jako ułamkami (`procent_1`),
   - kolumna z datami (`data_pl`),
   - kolumna z czasami trwania powyżej 24 h (`czas_trwania`),
   - kolumna z kodami pocztowymi jako tekst (`kod_pocztowy`) — **z zerami wiodącymi, np. `"00-950"`**,
   - kolumna z wartościami ujemnymi, żeby pokazać sekcję z nawiasem.

**Część B — test.** Napisz funkcję `testuj_rejestr()` (albo plik `examples/10_cw3_test.py`), która:

1. Dla **każdego** wpisu w `FORMAT` sprawdza, że nie jest pusty i jest typu `str`.
2. Dla **każdego** wpisu wypisuje: nazwę, kod, `is_builtin(kod)`, `is_date_format(kod)`, `is_timedelta_format(kod)`, `is_datetime(kod)`.
3. Sprawdza, że **formaty dat są rozpoznawane jako daty**, a **formaty liczbowe nie** — wypisuje ostrzeżenie, jeśli któryś format liczbowy ma `is_date_format() == True` (to sygnał pułapki z 2.6.6!).
4. Sprawdza, że wszystkie teksty w kodach formatów są **w cudzysłowie** — heurystyka: jeśli kod zawiera litery `z`, `ł`, `s`, `t`, `n` poza cudzysłowami i jednocześnie nie jest formatem daty, wypisuje ostrzeżenie.

W pliku `output/10_cw3_wnioski.md` odpowiedz:

1. **Ile formatów z Twojego rejestru jest wbudowanych, a ile własnych?** Podaj konkretne liczby i wypisz identyfikatory.
2. **Który z Twoich formatów ma `is_date_format() == True`, mimo że datą nie jest?** Jeśli żaden — uzasadnij, dlaczego udało Ci się tego uniknąć. Jeśli jakiś — wyjaśnij, jak go naprawisz.
3. **Uruchom `10_diagnostyka_formatow.py` na pliku z części A.** Ile wpisów `<numFmt>` powstało? Czy zgadza się to z liczbą własnych formatów z rejestru? **Wyjaśnij każdą różnicę.**
4. **Co się stanie, gdy ktoś zmieni `FORMAT["waluta_pln"]` z `'#,##0.00 "zł"'` na `'#,##0 "zł"'`?** Ile miejsc w kodzie musisz zmienić? Ile komórek zmieni wygląd? Wypisz odpowiedź — to jest miara wartości rejestru.
5. **Dlaczego kody pocztowe i numery PESEL mają format `@`?** Co się dzieje z `"00-950"` i z `"01234567890"`, jeśli zapiszesz je jako liczby? Sprawdź eksperymentalnie i opisz wynik. Zwróć uwagę: **format `@` nie naprawi tego, co już zostało zapisane jako liczba.**

### 🔴 Wyzwanie

**Zadanie 4 — biblioteka formatów z detekcją anomalii i raportem audytowym.**

Napisz `examples/10_cw4.py` implementujący klasę `FormatLibrary` **oraz** narzędzie do audytu cudzych plików.

```python
class FormatLibrary:
    """Rejestr formatow raportowych z walidacja i audytem."""

    def __init__(self, formaty: dict[str, str] | None = None):
        ...

    def ustaw(self, komorka, nazwa: str) -> None:
        """Ustawia format na komorce. Rzuca KeyError z lista dostepnych."""
        ...

    def ustaw_zakres(self, ws, zakres: str, nazwa: str) -> int:
        """Ustawia format na zakresie. Zwraca liczbe komorek."""
        ...

    def waliduj(self) -> dict:
        """Sprawdza spojnosc rejestru. Zwraca raport z ostrzezeniami."""
        ...

    def podejrzane_daty(self) -> list[tuple[str, str]]:
        """Zwraca formaty, ktore is_date_format() blednie uznaje za daty."""
        ...


def audyt_pliku(sciezka: Path) -> dict:
    """Analizuje ISTNIEJACY plik .xlsx pod katem formatow liczb.

    Zwraca raport:
      - liczba unikalnych kodow formatow w pliku,
      - kody formatow i ile razy wystepuja,
      - wykryte ANOMALIE (np. wiele wariantow waluty, formaty uznane
        za daty mimo ze nimi nie sa, formaty dat bez [h] w kolumnach czasowych),
      - liste komorek z wartosciami wygladajacymi na daty/liczby niewlasciwie.
    """
    ...
```

**Wymagania funkcjonalne:**

1. **Rejestr jako parametr konstruktora** — jeśli nie podano, użyj domyślnego. Klasa **nie może** mieć literałów kodu formatu w ciele metod.
2. **`waliduj()`** sprawdza: (a) czy żaden format liczbowy nie jest fałszywie rozpoznawany jako data, (b) czy teksty w kodach są w cudzysłowie, (c) czy formaty czasów trwania faktycznie zawierają `[h]`, `[m]` lub `[s]`, (d) czy nie ma dwóch różnych kodów o tej samej „roli" (np. dwa warianty waluty).
3. **`podejrzane_daty()`** zwraca listę par `(nazwa, kod)` dla formatów, gdzie `is_date_format(kod)` jest `True`, a nazwa nie zawiera `data`/`czas`/`miesiac`/`godzina`.
4. **`audyt_pliku()`** działa na **cudzym** pliku: wczytuje go przez `load_workbook`, przechodzi po komórkach i zbiera `number_format`, liczy wystąpienia, wykrywa anomalie.

**Anomalie, które musi wykryć `audyt_pliku()`:**

- **Wiele wariantów waluty.** Jeśli w pliku jest więcej niż jeden kod zawierający symbol waluty (`zł`, `PLN`, `€`, `$`) — zgłoś to jako „raport używa N wariantów waluty".
- **Formaty uznane za daty, choć nimi nie są.** Użyj `is_date_format()` na każdym znalezionym kodzie.
- **Procenty poza zakresem.** Dla komórek z formatem `%` sprawdź, czy wartość mieści się w rozsądnym zakresie (np. `-2` do `20`) — wartości poza tym zakresem to najprawdopodobniej pomylenie ułamka z liczbą procentową.
- **Puste komórki z formatem daty.** Jeśli `data_type == 'n'` i `value is None`, a format to data — zwróć ostrzeżenie (to częsty efekt uboczny stosowania formatów do zakresów „na zapas").

**Pięć testów:**

1. `ustaw` z nieistniejącą nazwą → `KeyError` z listą dostępnych.
2. `ustaw_zakres(ws, "A1:E10", "waluta_pln")` zwraca `50`.
3. `waliduj()` na poprawnym rejestrze zwraca pustą listę ostrzeżeń.
4. `podejrzane_daty()` wykrywa celowo wprowadzony format `'#,##0 szt.'` (bez cudzysłowu).
5. `audyt_pliku()` na pliku z Przykładu 1 wykrywa co najmniej jedną anomalię (podpowiedź: format `'0.0%'` na wartości `1234.5678`).

Na końcu wypisz tabelę `test N: OK / BŁĄD` i podsumowanie `wynik: X/5`. Zapisz raport do `output/10_cw4_raport.txt`.

W pliku `output/10_cw4_wnioski.md` odpowiedz:

1. **Jak wykryłbyś w cudzym pliku, że kolumna z czasem używa `h:mm:ss` zamiast `[h]:mm:ss`?** Opisz algorytm. Podpowiedź: szukaj w kolumnie wartości, które przy `[h]` byłyby inne niż przy `h`.
2. **Dlaczego `audyt_pliku` nie może po prostu porównać pliku z „wzorcem"?** Czym różni się audyt cudzego pliku od testowania własnego generatora?
3. **Jak Twój `waliduj()` wykrywa „dwa warianty waluty"?** Opisz heurystykę i jej ograniczenia — co, jeśli ktoś ma w jednym raporcie walutę polską i euro, bo tak ma być?
4. **Gdybyś miał dodać wykrywanie kolejnej anomalii — np. identyfikatorów zapisanych jako liczby zamiast tekstu — gdzie w kodzie musiałbyś to dopisać?** Wypisz miejsca. To miara otwartości klasy.

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""FormatLibrary + audyt cudzego pliku."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree

from openpyxl import load_workbook
from openpyxl.styles.numbers import (
    is_builtin,
    is_date_format,
    is_timedelta_format,
)
from openpyxl.utils.cell import range_boundaries

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

FORMATY_DOMYSLNE = {
    "liczba":           "#,##0",
    "liczba_2":         "#,##0.00",
    "waluta_pln":       '#,##0.00 "zł"',
    "waluta_pln_ksieg": '#,##0.00 "zł";(#,##0.00 "zł")',
    "waluta_pln_tys":   '#,##0, "tys. zł"',
    "waluta_eur":       '#,##0.00 "€"',
    "procent":          "0%",
    "procent_1":        "0.0%",
    "procent_2":        "0.00%",
    "data_iso":         "yyyy-mm-dd",
    "data_pl":          "[$-415]dd.mm.yyyy",
    "data_pl_dluga":    "[$-415]dd mmmm yyyy",
    "data_czas":        "yyyy-mm-dd hh:mm",
    "godzina":          "hh:mm",
    "godzina_sek":      "hh:mm:ss",
    "czas_trwania":     "[h]:mm:ss",
    "czas_trwania_min": "[h]:mm",
    "tekst":            "@",
    "kod_pocztowy":     "@",
    "pesel_nip":        "@",
}

# Slowa kluczowe w nazwach formatow, ktore usprawiedliwiaja bycie "data"
SLOWA_DATY = ("data", "czas", "miesiac", "godzina", "dzien", "rok")


class FormatLibrary:
    """Rejestr formatow z walidacja i detekcja anomalii."""

    def __init__(self, formaty: dict[str, str] | None = None):
        self._formaty = dict(formaty or FORMATY_DOMYSLNE)

    # ------------------------------------------------------------------
    # STOSOWANIE
    # ------------------------------------------------------------------
    def ustaw(self, komorka, nazwa: str) -> None:
        if nazwa not in self._formaty:
            dostepne = sorted(self._formaty)
            raise KeyError(f"nie ma formatu {nazwa!r}; dostepne: {dostepne}")
        komorka.number_format = self._formaty[nazwa]

    def ustaw_zakres(self, ws, zakres: str, nazwa: str) -> int:
        if nazwa not in self._formaty:
            dostepne = sorted(self._formaty)
            raise KeyError(f"nie ma formatu {nazwa!r}; dostepne: {dostepne}")

        kod = self._formaty[nazwa]
        min_col, min_row, max_col, max_row = range_boundaries(zakres)

        licznik = 0
        for wiersz in range(min_row, max_row + 1):
            for kolumna in range(min_col, max_col + 1):
                ws.cell(row=wiersz, column=kolumna).number_format = kod
                licznik += 1
        return licznik

    # ------------------------------------------------------------------
    # WALIDACJA wlasnego rejestru
    # ------------------------------------------------------------------
    def _czy_tekst_w_cudzyslowie(self, kod: str) -> bool:
        """Heurystyka: czy literaly tekstowe sa w cudzyslowie.

        Sprawdza, czy poza cudzyslowami i nawiasami [..] nie ma
        liter, ktore Excel/openpyxl moga zinterpretowac jako format.
        """
        bez_cudzyslowow = kod
        while '"' in bez_cudzyslowow:
            poczatek = bez_cudzyslowow.find('"')
            koniec = bez_cudzyslowow.find('"', poczatek + 1)
            if koniec == -1:
                return False                      # niezamkniety cudzyslow
            bez_cudzyslowow = (
                bez_cudzyslowow[:poczatek] + bez_cudzyslowow[koniec + 1:]
            )

        # usun rowniez nawiasy [..]
        import re
        bez_nawiasow = re.sub(r"\[[^\]]*\]", "", bez_cudzyslowow)

        # Znaki, ktore w formacie sa "reserved" - ich obecnosc poza
        # cudzyslowem moze byc bledem (np. 's' od 'szt.')
        return not any(z in bez_nawiasow for z in "dmhys")

    def waliduj(self) -> dict:
        """Zwraca raport z ostrzezeniami o niespojnosci rejestru."""
        ostrzezenia: list[str] = []

        waluty: list[str] = []
        for nazwa, kod in self._formaty.items():
            # (a) formaty liczbowe blednie rozpoznawane jako daty
            if is_date_format(kod) and not any(
                slowo in nazwa.lower() for slowo in SLOWA_DATY
            ):
                ostrzezenia.append(
                    f"format {nazwa!r} = {kod!r} jest rozpoznawany jako DATA,"
                    f" choc nazwa na to nie wskazuje (heurystyka is_date_format)"
                )

            # (b) teksty poza cudzyslowem
            if not self._czy_tekst_w_cudzyslowie(kod):
                ostrzezenia.append(
                    f"format {nazwa!r} = {kod!r} zawiera litere formatu"
                    f" (d/m/h/y/s) poza cudzyslowem - zamknij tekst w \"...\""
                )

            # (c) czas trwania musi miec nawiasy
            if "czas_trwania" in nazwa and not is_timedelta_format(kod):
                ostrzezenia.append(
                    f"format {nazwa!r} = {kod!r} nie ma nawiasow [h]/[m]/[s]"
                    f" - czasy powyzej 24 h zostana zawiniete!"
                )

            # (d) zbierz waluty do wykrycia duplikatow
            if any(symbol in kod for symbol in ("zł", "PLN", "€", "EUR", "$", "USD")):
                waluty.append(f"{nazwa}={kod}")

        if len(waluty) > 2:
            ostrzezenia.append(
                f"rejestr definiuje {len(waluty)} wariantow waluty: {waluty}."
                f" Upewnij sie, ze to zamierzone."
            )

        return {"ostrzezenia": ostrzezenia, "liczba_formatow": len(self._formaty)}

    def podejrzane_daty(self) -> list[tuple[str, str]]:
        """Formaty, ktore is_date_format() blednie uznaje za daty."""
        return [
            (nazwa, kod)
            for nazwa, kod in self._formaty.items()
            if is_date_format(kod)
            and not any(slowo in nazwa.lower() for slowo in SLOWA_DATY)
        ]


# ----------------------------------------------------------------------
# AUDYT CUDZEGO PLIKU
# ----------------------------------------------------------------------
def _czy_waluta(kod: str) -> bool:
    return any(symbol in kod for symbol in ("zł", "PLN", "€", "EUR", "$", "USD"))


def audyt_pliku(sciezka: Path) -> dict:
    """Analizuje istniejacy plik .xlsx pod katem formatow liczb."""
    wb = load_workbook(sciezka)
    try:
        licznik_formatow: dict[str, int] = {}
        podejrzane: list[str] = []
        procenty_poza_zakresem: list[str] = []
        puste_z_formatem_daty: list[str] = []

        for ws in wb.worksheets:
            for wiersz in ws.iter_rows():
                for komorka in wiersz:
                    kod = komorka.number_format
                    if kod in (None, "General"):
                        continue
                    licznik_formatow[kod] = licznik_formatow.get(kod, 0) + 1

                    # format wyglada na date, ale nazwa nie ma nic wspolnego
                    if is_date_format(kod) and not is_timedelta_format(kod):
                        pass          # sama data jest OK - nie zglaszamy

                    # procenty poza rozsadnym zakresem
                    if "%" in kod and isinstance(komorka.value, (int, float)):
                        if not (-2 <= komorka.value <= 20):
                            procenty_poza_zakresem.append(
                                f"{ws.title}!{komorka.coordinate} = "
                                f"{komorka.value} ({kod})"
                            )

                    # pusta komorka z formatem daty
                    if (
                        is_date_format(kod)
                        and komorka.value is None
                        and komorka.data_type == "n"
                    ):
                        puste_z_formatem_daty.append(
                            f"{ws.title}!{komorka.coordinate} ({kod})"
                        )

        warianty_walut = sorted(k for k in licznik_formatow if _czy_waluta(k))

        anomalie: list[str] = []
        if len(warianty_walut) > 1:
            anomalie.append(
                f"plik uzywa {len(warianty_walut)} wariantow waluty: {warianty_walut}"
            )
        if procenty_poza_zakresem:
            anomalie.append(
                f"{len(procenty_poza_zakresem)} komorek z procentami poza"
                f" sensownym zakresem (prawdopodobnie pomylone ulamki z %)"
            )
        if puste_z_formatem_daty:
            anomalie.append(
                f"{len(puste_z_formatem_daty)} pustych komorek z formatem daty"
                f" (formatowanie 'na zapas')"
            )

        return {
            "plik": str(sciezka),
            "liczba_unikalnych_formatow": len(licznik_formatow),
            "formaty": licznik_formatow,
            "warianty_walut": warianty_walut,
            "procenty_poza_zakresem": procenty_poza_zakresem,
            "puste_z_formatem_daty": puste_z_formatem_daty,
            "anomalie": anomalie,
        }
    finally:
        wb.close()
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **Rejestr wstrzykiwany przez konstruktor.** `FormatLibrary(formaty=...)` pozwala użyć tej samej klasy w dwóch raportach o **różnych** paletach formatów. Wersja z rejestrem w klasie na sztywno byłaby nieużywalna w drugim projekcie. To zapowiedź **wstrzykiwania zależności** z modułu 26.

- **`ustaw` i `ustaw_zakres` rzucają `KeyError` z listą dostępnych nazw.** Zamiast cichego „nic się nie stało", użytkownik od razu dowiaduje się, co ma do dyspozycji. To dokładnie ten sam wzorzec, co w `StyleLibrary` z modułu 09 — **błąd, który mówi, jak go naprawić**, jest warty dziesięć razy więcej niż błąd, który mówi tylko „nie działa".

- **`_czy_tekst_w_cudzyslowie` — heurystyka świadomie nieidealna.** Sprawdza, czy po usunięciu tekstów w cudzysłowie i nawiasów `[...]` nie została żadna litera ze zbioru `dmhys`. To **ta sama litera, którą szuka `is_date_format`** — i to jest celowe, bo chodzi o wykrycie dokładnie tej samej klasy problemu. Ograniczenie: fałszywe alarmy dla formatów zawierających **prawdziwe** znaki formatu daty bez cudzysłowu (np. `yy` w kodzie daty) — dlatego funkcja jest używana **tylko dla formatów nienazwanych jako daty**.

- **`waliduj()` zwraca raport, a nie rzuca wyjątkiem.** To wzorzec z modułu 09: rozdzielamy „szybką ścieżkę produkcyjną" (ustawienie formatu) od „wolnej ścieżki diagnostycznej" (walidacja rejestru). Walidacja nie może przerywać generowania raportu — ma informować.

- **`audyt_pliku` używa `load_workbook` bez `read_only`.** Bo audyt potrzebuje dostępu do `number_format` każdej komórki, a w trybie `read_only` jest to możliwe, ale ograniczone. **Wersja produkcyjna powinna sprawdzić, czy `read_only` wystarczy** — przy plikach rzędu setek tysięcy wierszy pełny odczyt jest kosztowny. To zapowiedź modułu 19.

- **Pięć rodzajów anomalii, każda w formie inspekcji, nie porównania ze wzorcem.** Dlaczego tak, a nie „porównaj z oczekiwanym"? Bo **audyt cudzego pliku to nie testowanie własnego generatora**. Nie masz wzorca — masz plik, o którym nie wiesz nic poza tym, że powstał w Excelu. Więc szukasz **wewnętrznych niespójności**: dwóch wariantów tej samej waluty, procentów w niemożliwym zakresie, formatów bez pokrycia w danych. To ta sama filozofia, którą w module 18 zastosujemy do wykrywania utraconych funkcji pliku.

**Czego szukać w wynikach:** uruchom `audyt_pliku` na pliku z Przykładu 1 (`10_wartosc_vs_reprezentacja.xlsx`). Powinien zgłosić:
1. warianty waluty — bo w pliku jest `'#,##0.00 "zł"'` (jedyny wariant, więc **nie** zgłosi),
2. **procenty poza zakresem** — bo komórka z `0.0%` ma wartość `1234.5678`, czyli 123456,8% — to jest właśnie ta anomalia, którą chcemy złapać,
3. prawdopodobnie kilkanaście pustych komórek z formatem daty (bo formaty były nadane „na zapas" do wierszy z opisami).

Ten trzeci punkt jest bardzo pouczający: pokazuje, że **nadanie formatu komórkom, które nigdy nie dostaną wartości, jest marnotrawstwem i szumem diagnostycznym**. To zapowiedź tematu wydajności z modułu 19.

**Co jeszcze warto dopisać w wersji produkcyjnej:** funkcję `porownaj_z_rejestrem(audyt, library)`, która odpowiada na pytanie „czy cudzy plik używa tych samych kodów formatów, co mój rejestr?" — i zwraca listę różnic. To narzędzie do migracji starych raportów na nowy standard (moduł 26).

</details>

## 6. Typowe błędy i pułapki

**1. „Marża wyszła 3500% zamiast 35%" (objaw) → w komórce jest `35`, a format `"0.0%"` mnoży wartość przez 100 przy rysowaniu. Format `%` w Excelu to **operator**, nie dekoracja: mówi „pomnóż przez 100 i dopisz znak procenta". Więc wartość procentowa musi być **ułamkiem** (`0.35`), a nie liczbą „już w procentach" (`35`) (przyczyna) → ustal w zespole **jeden standard** i trzymaj się go. Zalecany: **procenty zawsze jako ułamki**, bo tego oczekują formuły (`0.35 * 1000` daje sens, `35 * 1000` nie) i tego oczekują biblioteki statystyczne. Jeśli plik przychodzi z zewnątrz z procentami jako liczby — **przekonwertuj przy imporcie**, dzieląc przez 100, a nie „napraw" formatem. Weryfikacja: dla każdej komórki z formatem `%` sprawdź, czy wartość mieści się w sensownym zakresie (np. $-2$ do $20$) — patrz `audyt_pliku` w zadaniu 4 (naprawa).**

**2. „Wpisałem format `#,##0 szt.` i `cell.is_date` zwraca `True` dla zwykłej liczby" (objaw) → heurystyka `is_date_format` w openpyxl usuwa teksty **w cudzysłowie** i szuka liter `d`, `m`, `h`, `y`, `s` poza nimi. W kodzie `#,##0 szt.` litera `s` (od „szt.") **nie jest w cudzysłowie**, więc jest znaleziona i openpyxl uznaje format za datę. Konsekwencje: `is_date` kłamie, a jeśli później konwertujesz wartości na daty „bo to data" — dostajesz bzdury (przyczyna) → **zawsze zamykaj tekst w kodzie formatu w cudzysłów**: `'#,##0 "szt."'`. Sprawdź swój rejestr funkcją `is_date_format()` i szukaj wpisów, które są `True`, choć nazwa nie mówi o dacie. Ta jedna zasada („tekst zawsze w cudzysłowie") eliminuje całą klasę problemów: heurystyczne pomyłki, niejednoznaczności dla Excela i sytuacje, w których Twoje słowo tekstowe (`szt.`, `dni`, `mies.`, `hyb.`) zaczyna być interpretowane jako format daty (naprawa).**

**3. „Nadgodziny wychodzą za małe — pracownik miał 26 godzin, a widać 2:00" (objaw) → użyty format to `h:mm:ss` (zegarek), a nie `[h]:mm:ss` (stoper). Excel domyślnie **zawija** godziny po 24, tak jak zegarek po północy. Nawiasy kwadratowe mówią „nie zawijaj na dobie" (przyczyna) → **dla każdej kolumny czasów trwania używaj `[h]:mm:ss` albo `[h]:mm`**. Dotyczy to nadgodzin, czasów obsługi, czasów trwania procesów, długości rozmów, sumarycznych czasów w logach. Reguła: jeśli kolumna opisuje **upływ czasu** (a nie „któraś godzina"), format **musi** mieć nawiasy. Weryfikacja: sprawdź, czy suma w kolumnie przekracza 24 godziny — jeśli tak, a wyświetlana wartość jest mniejsza, format jest zły. W diagnostyce pomoże `is_timedelta_format(cell.number_format)` (naprawa).**

**4. „`number_format = 0,00` dało dziwne wyniki" (objaw) → w **kodzie formatu** Excela separator dziesiętny to **zawsze kropka**, a separator tysięcy to **zawsze przecinek** — niezależnie od locale. Excel sam podmienia je na lokalne znaki przy wyświetlaniu. Wpisanie polskiego `0,00` nie znaczy „zero z dwoma miejscami po przecinku" — przecinek **na końcu** ma zupełnie inne znaczenie: **dzieli przez tysiąc**. Kod `0,00` oznacza zatem „podziel przez tysiąc dwa razy" (przyczyna) → **w kodach formatów pisz zawsze `0.00` i `#,##0`**, korzystając z angielskiej notacji, nawet jeśli pracujesz po polsku. Excel wyświetli to poprawnie jako `1 234,56` w polskim locale. Jeśli chcesz wymusić konkretne separatory niezależnie od systemu odbiorcy — musisz użyć kodu LCID, np. `[$-415]#,##0.00`, i sprawdzić wynik w docelowym Excelu (naprawa).**

**5. „Data wyświetla się jako liczba (np. 46082), mimo że przypisałem `datetime`" (objaw) → prawdopodobnie komórka **miała już ustawiony format, który nie jest formatem daty**, a openpyxl nadaje domyślny format daty **tylko wtedy, gdy** `is_date_format(number_format)` zwraca `False`. Jeśli format komórki pochodzi ze stylu nazwanego albo z szablonu i jest np. `General`, a Ty najpierw **wczytałeś** plik i modyfikujesz istniejącą komórkę, możesz trafić na format, który „wygląda jak General, a nie jest" (przyczyna) → **ustawiaj format jawnie, nie licz na automat**: `komorka.number_format = "yyyy-mm-dd"` po przypisaniu wartości. Diagnostyka: wypisz `cell.number_format` **przed** i **po** przypisaniu wartości i porównaj. Pamiętaj też o drugiej kolejności — jeśli **najpierw** ustawisz format daty, a potem wartość, automat **nie** nadpisze Twojego formatu (i to jest zachowanie pożądane) (naprawa).**

**6. „Nazwa miesiąca w formacie `mmmm` wyświetla się po angielsku, mimo że piszę po polsku" (objaw) → nazwy miesięcy i dni tygodnia **nie są zapisane w pliku**. W pliku jest liczba (`46082`) i instrukcja (`mmmm`). Tekst powstaje **na komputerze odbiorcy**, w języku jego systemu/interfejsu Excela — nie w języku Twojego kodu Pythona (przyczyna) → użyj **kodu LCID** w formacie: `[$-415]dd mmmm yyyy` dla polskiego (`415` = polski!), `[$-409]` dla angielskiego. **Uważaj na mylenie kodów:** `415` to polski, `417` to łotewski — w internecie krąży sporo błędnych przykładów. Jeśli dokument musi wyglądać **identycznie** wszędzie, użyj ISO 8601 (`yyyy-mm-dd`, bez nazw miesięcy) albo zapisz tekst zamiast daty (naprawa).**

**7. „Kod pocztowy `00-950` zapisał się jako `-950`" / „PESEL stracił wiodące zero" (objaw) → zapisanie wartości jako **liczby** usuwa zera wiodące — bo liczba `950` i liczba `00950` to ta sama liczba `950`. Nie chodzi o format: `#` i `0` w kodzie formatu decydują tylko o tym, ile cyfr **wyświetlić**, a nie o tym, ile ich jest w wartości. Jeśli wartość nie ma zer wiodących, żaden format ich nie doda — może co najwyżej **dopisać** zera przez `0` (np. format `00000` pokaże `950` jako `00950`), ale to nie odtworzy prawdziwego kodu pocztowego (przyczyna) → **identyfikatory zawsze jako tekst.** Ustaw format `"@"` i przypisz wartość jako `str`: `komorka.value = "00-950"`. Format `@` mówi „to jest tekst, wyświetl dosłownie". **Kluczowe ostrzeżenie:** format `@` **nie naprawi** wartości już zapisanej jako liczba — `cell.number_format = "@"` na komórce z wartością `950` nadal pokaże `950`, a nie `00950`. Musisz poprawić **wartość** (`cell.value = "00950"`), a nie format (naprawa).**

**8. „Format `#,##0.00 "zł"` pokazuje się w komórce dosłownie jako tekst, zamiast formatować liczbę" (objaw) → dwie możliwe przyczyny. Po pierwsze: **wartość komórki nie jest liczbą** (jest napisem `"1234.56"`), więc Excel traktuje ją jako tekst i wyświetla tak, jak jest — formaty liczbowe nie mają na co działać. Po drugie: **kod formatu jest niepoprawny składniowo**, np. brakuje cudzysłowu wokół tekstu albo użyto znaku nieobsługiwanego przez daną wersję Excela — wtedy Excel pokazuje kod jako surowy tekst (przyczyna) → sprawdź najpierw `type(cell.value)` (musi być `int`, `float` lub `Decimal`, nie `str`); następnie sprawdź, czy każdy literalny tekst jest w cudzysłowie. Pamiętaj, że **openpyxl nie waliduje kodu formatu** — przyjmie dowolny napis. Jeśli wartość ma typ tekstowy, a ma być liczbą, napraw **wartość** (konwersja przy imporcie), nie format (naprawa).**

**9. „Ustawiłem `cell.number_format = None` i format nie zniknął" (objaw) → `NumberFormatDescriptor` w openpyxl zamienia `None` na `"General"`. Czyli nie dostajesz pustego formatu, tylko format ogólny, który jednak wyświetla liczby „normalnie". W większości przypadków efekt jest zbliżony, ale nie identyczny — np. dla bardzo małych lub bardzo dużych liczb `General` może przełączyć się na notację naukową (przyczyna) → jeśli chcesz powrócić do zachowania domyślnego, ustaw jawnie `komorka.number_format = "General"`. Jeśli chcesz, żeby komórka **nie miała** żadnego formatu — wiedz, że takiego stanu się nie da uzyskać przez `None`; musisz ustawić styl „Normal" (naprawa).**

**10. „Data przesunęła się o cztery lata przy wczytaniu pliku od klienta z Maca" (objaw) → Excel zna **dwie epoki**: Windows (`1899-12-30`, domyślna) i Mac 1904 (`1904-01-01`). Pliki utworzone w trybie „skoroszyt od 1904" mają inną numerację — ta sama data ma o 1462 dni więcej/mniej. Jeśli samodzielnie przeliczasz daty, dostajesz przesunięcie o cztery lata (przyczyna) → **nie przeliczaj dat samodzielnie.** openpyxl odczytuje epokę z pliku do `wb.epoch` i używa jej przy konwersji. Jeśli piszesz własną konwersję — użyj `wb.epoch`, nigdy nie zakładaj `1899-12-30` na sztywno. Uwaga też na pokrewny problem: **fikcyjny 29 lutego 1900** (serial `60`), który Excel utrzymuje dla zgodności z Lotusem — daty sprzed 1 marca 1900 są w numeracji Excela o dzień przesunięte (naprawa).**

**11. „Data z pliku nie równa się dacie w Pythonie" (objaw) → `datetime` świadomy strefy czasowej (`tzinfo`) **nie jest obsługiwany** przez openpyxl — strefa zostaje utracona przy zapisie, a przy odczycie wartość jest **naiwna**. Porównanie naiwnego `datetime` ze świadomym **zawsze** zwraca `False` w Pythonie (i nie rzuca wyjątku! — tylko `==` daje `False`, a `-` podnosi `TypeError`) (przyczyna) → konwertuj wszystko na **UTC i zapisuj jako naiwne `datetime`**, a informację o strefie umieść w nagłówku kolumny albo w arkuszu `_meta`. Przy porównaniach zawsze sprowadzaj obie strony do tego samego stanu: `odczytana.replace(tzinfo=timezone.utc)` albo `.astimezone()` (naprawa).**

**12. „`timedelta` zachowuje się dziwnie — `is_date` zwraca coś nieoczekiwanego" (objaw) → klasyfikacja `timedelta` w openpyxl jest nieoczywista: `TIMEDELTA` jest w `TIME_TYPES`, więc dostaje `data_type == 'd'`, ale zapisuje się jako liczba (ułamek doby). W repozytorium openpyxl istnieje **test oznaczony jako `xfail`** dla tego przypadku, co jest jawnym przyznaniem, że zachowanie odbiega od oczekiwanego (przyczyna) → **nie opieraj logiki na `data_type` przy `timedelta`.** Sprawdzaj wartość przez `isinstance(value, timedelta)` albo format przez `is_timedelta_format(cell.number_format)`. **Sprawdź też w swojej wersji openpyxl**, jak dokładnie `timedelta` zachowuje się przy odczycie — to jedno z tych miejsc, gdzie wersja ma znaczenie (naprawa).**

**13. „Raport ma trzy różne odcienie tej samej waluty" (objaw) → kody formatów są rozsypane po kodzie: `'#,##0.00 "zł"'`, `'#,##0.00 zł'`, `'# ##0,00 zł'` w różnych miejscach, pisane w pośpiechu i bez sprawdzania. Dodatkowo jeden z nich (`#,##0.00 zł` bez cudzysłowu) wpadnie w heurystyczną pułapkę `is_date_format` (przyczyna) → **centralny rejestr formatów** (`FORMAT = {...}` z sekcji 2.11) plus funkcja `ustaw_format(komorka, nazwa)`. Wszystkie kwoty w raporcie przechodzą przez jedną nazwę, a nie przez literał. Sprawdź wynik funkcją diagnostyczną z Przykładu 5: jeśli w `<numFmts>` jest więcej niż jeden kod zawierający `zł`, masz niespójność. Docelowo zero formatów jako literały w kodzie raportu — wszystkie z rejestru (naprawa).**

**14. „Formatowanie nie zaokrągla kwot, a ja potrzebowałem, żeby zaokrąglało" (objaw) → format **nigdy** nie zmienia wartości — to tylko wyświetlanie. Komórka z wartością `19.9876` i formatem `0.00` pokaże `19,99`, ale w pliku nadal jest `19.9876`, a formuły liczą na pełnej wartości. Jeśli potem pomnożysz przez 1000, wynik będzie oparty na `19.9876`, nie na `19.99` (przyczyna) → jeśli naprawdę potrzebujesz **zaokrąglonej wartości**, użyj funkcji Excela (`=ROUND(A1,2)`) albo zaokrąglij w Pythonie **przed** przypisaniem: `komorka.value = round(wartosc, 2)`. Pamiętaj przy tym, że `round()` w Pythonie używa zaokrąglania bankierskiego (do parzystej), a Excel — arytmetycznego (połówki w górę) — jeśli to ma znaczenie dla pieniędzy, użyj `Decimal` z jawnym trybem zaokrąglania (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Wartość i reprezentacja to dwa odrębne światy, połączone jednym numerem.** W komórce pliku `.xlsx` siedzi liczba (`<v>19.99</v>`), a w `styles.xml` — kod formatu. Excel składa je na ekranie dopiero w momencie rysowania. **Format nigdy nie zmienia wartości**: nie zaokrągla, nie przelicza, nie konwertuje typu. Możesz mieć pięć komórek z identyczną liczbą `1234.5678` i pięć zupełnie różnych obrazów na ekranie. Z tego jednego zdania wynikają wszystkie konsekwencje tego modułu — włącznie z tym, że „naprawianie" danych przez zmianę formatu jest niemożliwe.

2. **Kod formatu to formularz z czterema rubrykami i zbiorem znaków o konkretnych znaczeniach.** Sekcje (`dodatnie; ujemne; zero; tekst`) są rozdzielone średnikami i Excel używa **dokładnie jednej** — tej pasującej do wartości. Znaki mają precyzyjne znaczenie: `#` (opcjonalna cyfra), `0` (obowiązkowa), `_` (spacja), `*` (wypełnienie), `@` (tekst), `[Red]` (kolor), `[h]` (nie zawijaj po dobie), `[$-415]` (polska lokalizacja). **Najbardziej zdradliwy jest przecinek**, który znaczy „separator tysięcy" w środku kodu, a „podziel przez tysiąc" na końcu. I najważniejsza reguła praktyczna: **separatorem dziesiętnym w kodzie formatu jest zawsze kropka, nawet gdy piszesz po polsku.**

3. **Tekst w kodzie formatu zawsze zamykaj w cudzysłowie — i nie jest to kwestia estetyki, a poprawności.** `'#,##0 "szt."'` i `'#,##0 szt.'` **nie są** równoważne. Heurystyka `is_date_format` w openpyxl usuwa teksty w cudzysłowie i szuka liter `d`, `m`, `h`, `y`, `s` poza nimi — więc `szt.` bez cudzysłowu zostanie znalezione przez `s` i openpyxl uzna format za **format daty**, mimo że nim nie jest. Konsekwencje: `cell.is_date` kłamie, a jeśli później oprzesz na tym konwersje, dostaniesz bzdury. Ta jedna zasada eliminuje całą klasę problemów i sprawia, że kody formatów są jednoznaczne zarówno dla openpyxl, jak i dla Excela.

4. **Daty to liczby, a `[h]` to jedyna rzecz, która czyni z zegarka stoper.** Data w Excelu to liczba dni od epoki (`1899-12-30` dla Windows, `1904-01-01` dla Maka), a godzina to ułamek doby — dlatego `2026-03-01 06:00` to `46082.25`. To wyjaśnia, dlaczego daty mają dwa typy błędów (epoka Maca i fikcyjny 29 lutego 1900) i dlaczego **nigdy nie należy przeliczać ich ręcznie**. Osobno: **czasy trwania wymagają formatu z nawiasami** (`[h]:mm:ss`) — bez nich Excel zawija godziny po dobie i raport nadgodzin pokazuje `2:30` zamiast `26:30`. To najbardziej kosztowny cichy błąd w raportach kadrowo-płacowych, bo plik wygląda poprawnie.

5. **Raport bez centralnego rejestru formatów zawsze skończy się niespójnością.** Trzy odcienie waluty, dwa warianty procentów, procenty jako ułamki w jednej kolumnie i jako liczby w drugiej, formaty dat uznawane za coś innego przez openpyxl — to wszystko skutki rozproszenia. Rozwiązanie jest jedno i sprawdzalne: **jeden słownik `FORMAT` z nazwami logicznymi (`"waluta_pln"`, `"procent_1"`, `"czas_trwania"`), jedna funkcja `ustaw_format(komorka, nazwa)` i zero literałów kodów formatu w pozostałej części kodu.** Sprawdzalna miara jakości: liczba wpisów w `<numFmts>` w `xl/styles.xml` powinna być **mała, stabilna i zgodna z twoim rejestrem** — bez względu na to, ile komórek sformatujesz.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. FUNDAMENT - wartosc kontra reprezentacja
# ==================================================================
ws["B2"] = 19.99                            # WARTOŚĆ (liczba w pliku)
ws["B2"].number_format = '#,##0.00 "zł"'    # REPREZENTACJA (kod w styles.xml)
# Format NIE zmienia wartości. Nie zaokrągla. Nie przelicza.
# Do komórki w kodzie formatu trafia ZWYKŁY NAPIS - nie ma tu obiektu stylu.

# None -> "General" (NumberFormatDescriptor):
ws["B2"].number_format = None               # -> "General"

# ==================================================================
# 2. FORMATY WBUDOWANE - najczęściej używane identyfikatory
# ==================================================================
#  0 General      1 0          2 0.00       3 #,##0      4 #,##0.00
#  9 0%          10 0.00%     11 0.00E+00  12 # ?/?     13 # ??/??
# 14 mm-dd-yy    20 h:mm      21 h:mm:ss   22 m/d/yy h:mm
# 37-44 księgowe 45 mm:ss     46 [h]:mm:ss 48 ##0.0E+0  49 @
# Identyfikatory 0-163 = WBUDOWANE. Własne kody startują od 164!

from openpyxl.styles.numbers import (
    is_builtin, builtin_format_id, builtin_format_code, BUILTIN_FORMATS,
)
is_builtin('#,##0.00')          # True
builtin_format_id('#,##0.00')   # 4

# ==================================================================
# 3. ANATOMIA KODU FORMATU
# ==================================================================
# dodatnie ; ujemne ; zero ; tekst
'#,##0.00 "zł";(#,##0.00 "zł");"—";@" (do weryfikacji)"'

# ZNAKI:
#   0   -> obowiazkowe miejsce cyfry        00  -> 7 pokaze jako 07
#   #   -> opcjonalne miejsce cyfry
#   ?   -> miejsce zarezerwowane (wyrownanie dziesietne)
#   .   -> separator dziesietny  (ZAWSZE kropka w kodzie!)
#   ,   -> separator tysiecy LUB dzielenie przez 1000 na KONCU
#   %   -> pomnóz przez 100 i dopisz znak procenta
#   E+  -> notacja naukowa
#   ".."-> tekst doslowny        ZAWSZE ZAMYKAJ TEKST W CUDZYSLOWIE!
#   \   -> escape nastepnego znaku
#   _   -> spacja szerokosci nastepnego znaku
#   *   -> powtarzaj znak do konca szerokosci kolumny
#   @   -> miejsce na wartosc TEKSTOWA
#   [Red] [Blue] [Green] [Black] [White] [Cyan] [Magenta] [Yellow]
#   [$-415] -> lokalizacja (415 = polski!)   417 = lotewski, NIE polski!
#   [h] [m] [s] -> czas trwania, NIE zawijaj na dobie
#   [>100] -> warunek (max 2 warunki w jednym kodzie)

# Przykłady:
'#,##0.00 "zł"'                  # 1 234,56 zł
'#,##0.00 "zł";(#,##0.00 "zł")'  # ujemne w nawiasie
'#,##0, "tys."'                  # dzielenie przez 1000
'0.0%'                           # 0.35 -> 35,0%   |  35 -> 3500,0%  (!)
'@" (do weryfikacji)"'           # tekst + dopisek
'[Red]#,##0.00'                  # zawsze czerwone
'#,##0.00;[Red]-#,##0.00'        # ujemne na czerwono

# ==================================================================
# 4. DATY I CZAS
# ==================================================================
from datetime import date, datetime, time, timedelta

# Wartosc = liczba dni od epoki; godzina = ulamek doby.
# 2026-03-01 06:00  ->  46082.25

# openpyxl nadaje format AUTOMATYCZNIE, ale TYLKO gdy komorka nie ma
# jeszcze formatu daty:
ws["A1"] = date(2026, 3, 1)              # -> format 'yyyy-mm-dd'
ws["A2"] = datetime(2026, 3, 1, 14, 30)  # -> format 'yyyy-mm-dd h:mm:ss'
ws["A3"] = time(14, 30)                  # -> format 'h:mm:ss'
ws["A4"] = timedelta(hours=26, minutes=30)  # -> format '[hh]:mm:ss'

# Kolejnosc: format PRZED wartoscia = pelna kontrola (automat nie nadpisze)
ws["C2"].number_format = "yyyy-mm-dd"
ws["C2"] = date(2026, 3, 1)

# Format daty po polsku / angielsku:
ws["D2"].number_format = "[$-415]dd mmmm yyyy"   # 01 marzec 2026
ws["D3"].number_format = "[$-409]dd mmmm yyyy"   # 01 March 2026

# CZASY TRWANIA - nawiasy kwadratowe sa OBOWIAZKOWE!
ws["E2"].number_format = "[h]:mm:ss"    # 26:30:00 (stoper - NIE zawija)
ws["E3"].number_format = "h:mm:ss"      # 02:30:00 (zegarek - ZAWIJA!)  ❌

# Epoka skoroszytu:
print(wb.epoch)      # 1899-12-30 (Windows) lub 1904-01-01 (Mac)
# NIE przeliczaj dat recznie - uzyj wb.epoch albo openpyxl.utils.datetime
from openpyxl.utils.datetime import from_excel, to_excel
# ⚠️ Sygnatury zmienialy sie miedzy wersjami - sprawdz w swojej wersji

# ==================================================================
# 5. DATA_TYPE I IS_DATE - diagnostyka
# ==================================================================
# data_type:  'n' liczba (TAKZE PUSTA!),  's' tekst,  'd' data/czas,
#             'b' bool,  'f' formula,  'e' blad,  'inlineStr',  'str'
#
# Komorka pusta ma data_type == 'n' (TYPE_NULL = TYPE_NUMERIC = 'n')
# Napis zaczynajacy sie od '=' (dluzszy niz 1) -> data_type 'f'
# Napis rowny '#N/A', '#DIV/0!' itd. -> data_type 'e'
# Samotny '=' -> pozostaje 's' (tekst)

# is_date - ISTNIEJE w 3.1.x. Implementacja:
#   return self.data_type == 'd' or (
#       self.data_type == 'n' and is_date_format(self.number_format))
# ⚠️ is_date NIE patrzy na wartosc - tylko na typ i format!

from openpyxl.styles.numbers import is_date_format, is_datetime, is_timedelta_format

is_date_format("yyyy-mm-dd")      # True
is_date_format("0.00")            # False
is_date_format('#,##0 szt.')      # TRUE (!) bo 's' poza cudzyslowem
is_date_format('#,##0 "szt."')    # False - tekst usuniety
is_timedelta_format("[h]:mm:ss")  # True
is_datetime("yyyy-mm-dd h:mm:ss") # 'datetime'

# ==================================================================
# 6. WZORZEC: CENTRALNY REJESTR FORMATOW
# ==================================================================
FORMAT = {
    "liczba":           "#,##0",
    "liczba_2":         "#,##0.00",
    "waluta_pln":       '#,##0.00 "zł"',
    "waluta_pln_ksieg": '#,##0.00 "zł";(#,##0.00 "zł")',
    "waluta_pln_tys":   '#,##0, "tys. zł"',
    "procent":          "0%",
    "procent_1":        "0.0%",
    "data_iso":         "yyyy-mm-dd",
    "data_pl":          "[$-415]dd.mm.yyyy",
    "data_pl_dluga":    "[$-415]dd mmmm yyyy",
    "data_czas":        "yyyy-mm-dd hh:mm",
    "godzina":          "hh:mm",
    "godzina_sek":      "hh:mm:ss",
    "czas_trwania":     "[h]:mm:ss",
    "czas_trwania_min": "[h]:mm",
    "tekst":            "@",
    "kod_pocztowy":     "@",
}

def ustaw_format(komorka, nazwa: str) -> None:
    """Ustawia format z rejestru. Rzuca KeyError z lista dostepnych nazw."""
    if nazwa not in FORMAT:
        raise KeyError(f"nie ma formatu {nazwa!r}; dostepne: {sorted(FORMAT)}")
    komorka.number_format = FORMAT[nazwa]

def ustaw_format_zakresu(ws, zakres: str, nazwa: str) -> int:
    from openpyxl.utils.cell import range_boundaries
    kod = FORMAT[nazwa]
    min_col, min_row, max_col, max_row = range_boundaries(zakres)
    licznik = 0
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).number_format = kod
            licznik += 1
    return licznik

# ==================================================================
# 7. PIENIĄDZE - grosze jako int
# ==================================================================
from decimal import Decimal

# float NIE nadaje sie na pieniadze:  0.1 + 0.2 != 0.3
# Decimal jest akceptowany (NUMERIC_TYPES), ale plik nie gwarantuje
# precyzji - po odczycie dostajesz float.
grosze = int((Decimal("19.99") * 100).quantize(Decimal("1")))   # 1999
do_excela = float(Decimal(grosze) / 100)                        # 19.99
ws["B2"] = do_excela
ws["B2"].number_format = '#,##0.00 "zł"'

# ==================================================================
# 8. CZEGO NIE DA SIĘ ZROBIĆ / NA CO UWAŻAĆ
# ==================================================================
# ❌ Zmienic wartosci przez format        -> format to tylko wyswietlanie
# ❌ Zaokraglic kwoty przez format        -> uzyj ROUND() albo round() w Pythonie
# ❌ Strefy czasowej (tzinfo)             -> gubiona przy zapisie; zapisuj UTC naive
# ❌ Uzyc h:mm:ss dla czasow > 24 h       -> zawija! Uzyj [h]:mm:ss
# ⚠️ openpyxl NIE waliduje kodu formatu   -> bledny kod = Excel pokaze tekst
# ⚠️ Kropka jako separator dziesietny w kodzie -> ZAWSZE, nawet po polsku
# ⚠️ Tekst w kodzie formatu               -> ZAWSZE w cudzyslowie "..."
# ⚠️ Nazwa miesiaca zalezy od systemu odbiorcy, nie od Twojego kodu

# ==================================================================
# 9. DIAGNOSTYKA - ile formatow naprawde jest w pliku
# ==================================================================
import zipfile
from xml.etree import ElementTree

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def policz_formats(sciezka) -> dict:
    with zipfile.ZipFile(sciezka) as z:
        korzen = ElementTree.fromstring(z.read("xl/styles.xml"))

    sekcja = korzen.find("m:numFmts", NS)
    wlasne = []
    if sekcja is not None:
        wlasne = [(e.get("numFmtId"), e.get("formatCode")) for e in sekcja]

    cell_xfs = korzen.find("m:cellXfs", NS)
    uzyte = set()
    if cell_xfs is not None:
        for xf in cell_xfs:
            uzyte.add(xf.get("numFmtId", "0"))

    return {
        "wlasne_formaty": wlasne,           # numFmtId >= 164
        "liczba_wlasnych": len(wlasne),
        "uzyte_numFmtId": sorted(uzyte, key=int),
    }
# Format "General" (id 0) NIE jest zapisywany jako <numFmt>.
# Jeden kod uzyty w 1000 komorek = JEDEN wpis (deduplikacja).
```

## 9. Co dalej

Masz za sobą drugą połowę tematu formatowania. Moduł 09 dał Ci **pieczątki** (jak wygląda tekst, tło, ramka), moduł 10 dołożył **szósty element** — jak wygląda **wartość**. Razem masz pełny obraz tego, co można zrobić z pojedynczą komórką. Trzy rzeczy, które zabierasz ze sobą dalej:

- **Rozdzielenie wartości od reprezentacji jako odruch.** Od teraz, gdy zobaczysz w raporcie „3500%", twoje pierwsze pytanie nie brzmi „jak naprawić format?", a „**co jest w komórce**?". Bo w większości przypadków problem nie leży w formacie, tylko w tym, co ktoś do niego włożył.
- **Cudzysłów w kodach formatów jako nawyk.** To drobiazg, który wygląda jak pedanteria, dopóki `#,##0 szt.` nie zamieni twojej kolumny liczb w kolumnę „dat" dla openpyxl.
- **Nawiasy w formatach czasów trwania jako kontrola jakości.** Jeśli twój raport ma kolumnę z sumami godzin, **sprawdź, czy suma przekracza 24** — i jeśli przekracza, a wyświetla się mniejsza, masz błąd formatu, którego nikt inny nie zauważy.

**W module 11** wyjdziemy poza pojedynczą komórkę i zajmiemy się **wyglądem całego arkusza**: szerokości kolumn, wysokości wierszy, scalaniem, zamrażaniem okien, autofiltrem, a przede wszystkim **drukiem i eksportem do PDF**. To moduł praktyczny — zobaczysz:

- dlaczego `ws.column_dimensions["B"].font = ...` **nie działa** na istniejące komórki (i dlaczego to ograniczenie samego formatu pliku, nie openpyxl),
- jak działa `freeze_panes` (i dlaczego `"B2"` zamraża inny obszar niż `"A3"`),
- jak przygotować arkusz do druku na A4 poziomo z **powtarzanym nagłówkiem na każdej stronie** (`print_title_rows`),
- dlaczego `fitToWidth`/`fitToHeight` **nie działają** bez trzeciego ustawienia w `sheet_properties.pageSetUpPr`,
- jak zbudować nagłówek i stopkę druku z numeracją stron (`&P`, `&N`, `&D`) — czyli ten dziwny „język" w stopkach Excela,
- i jak scalanie komórek **niszczy dane** (to powtórzenie ostrzeżenia z modułu 09, tym razem z pełnym wyjaśnieniem mechanizmu).

**Zanim przejdziesz dalej, wykonaj dwa przygotowania:**

1. **Uruchom `examples/10_diagnostyka_formatow.py`** na katalogu `output/` — po tym module powinno tam być kilka plików z Przykładów 1–5. Zapisz sobie **liczbę wpisów `<numFmt>`** w każdym. W module 11 dojdzie do nich **wymiar arkusza** (szerokości, wysokości, ustawienia druku), a w module 12 — **tabele**. Zobaczysz, że `xl/styles.xml` i `xl/worksheets/sheetN.xml` rosną w różnym tempie i z różnych powodów.

2. **Zajrzyj do jednego ze swoich starych raportów** — takiego, który powstał przed tym kursem. Uruchom na nim `audyt_pliku` z zadania 4 albo prostszy skrypt z Przykładu 5. Zadaj sobie trzy pytania:
   - **Ile unikalnych kodów formatów jest w tym pliku?** Jeśli odpowiedź to „jedenaście", a ty definiowałeś trzy — właśnie widzisz koszt braku rejestru.
   - **Czy kolumna z czasem ma nawiasy `[h]`?** Jeśli nie i wartości przekraczają 24 godzin — znalazłeś cichy błąd.
   - **Czy są komórki z formatem daty i pustą wartością?** Jeśli tak — ktoś sformatował „na zapas" zakres, który nie został wypełniony. To zapowiedź modułu 19: formatowanie komórek, które nie istnieją w sensie danych, to marnotrawstwo, które w dużych plikach boli.