Zanim napiszę moduł, zweryfikuję API formatowania warunkowego w openpyxl 3.1.x — to obszar, w którym między wersjami zmieniało się zarządzanie kolekcją `conditional_formatting`, a dokumentacja i kod nie zawsze mówią to samo (klasyczny motyw kursu).

Kluczowe fakty potwierdzone (m.in. że konstruktory reguł **nie przyjmują** `priority`, że `add()` nadaje priorytet automatycznie, oraz że kolekcja nie ma `clear()`). Zweryfikuję jeszcze zachowanie przy **wczytywaniu** istniejących reguł — czy `dxf` jest rozwiązywane, i jak działa iteracja/`del`.

Mam pełny obraz kolekcji (m.in. że `add` **łączy** reguły pod tym samym `sqref`, że `del ws.conditional_formatting["A1:A10"]` usuwa cały zakres, i że wczytane reguły zachowują `dxf`). Zweryfikuję jeszcze semantykę `MultiCellRange` (równość/hash i `isdisjoint`) — to nośne fakty dla kilku przykładów.

# Moduł 13 — Formatowanie warunkowe: reguły, które mieszkają w arkuszu

> **Część:** II — Formatowanie · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–12

## 0. W tym module nauczysz się

- **Zrozumiesz kluczową ideę: formatowanie warunkowe to reguła mieszkająca w arkuszu, a nie efekt zapisany w komórce.** To zmiana modelu myślenia tej samej wagi co „tabela to obiekt, nie zakres" z modułu 12. Zobaczysz, dlaczego po wczytaniu pliku `cell.fill` **nie pokaże** koloru z formatowania warunkowego — i dlaczego to nie błąd, tylko natura rzeczy.
- **Poznasz `DifferentialStyle` (dxf)** i zrozumiesz, że reguła niesie tylko *różnicę* względem stylu komórki, a nie pełny styl. To wyjaśnia, dlaczego w pliku są dwa różne światy kolorów: jeden w `styles.xml`, drugi w `dxfs`.
- **Opanujesz model priorytetów i konfliktów:** `rule.priority` (mniejsza liczba = ważniejsza), `rule.stopIfTrue`, oraz to, że priorytety są **globalne dla arkusza**, nie dla zakresu — i jak openpyxl nadaje je automatycznie.
- **Zrozumiesz największą pułapkę tego działu:** formuła w regule jest zapisywana **z punktu widzenia lewej-górnej komórki zakresu** i używa odniesień względnych/absolutnych zgodnie z intencją. To miejsce, w którym nawet doświadczeni autorzy piszą reguły, które „wyglądają prawie dobrze".
- **Nauczysz się dwóch najważniejszych narzędzi:** `CellIsRule` (porównanie komórki z wartością) i `FormulaRule` (warunek jako formuła Excela — najpotężniejsze narzędzie formatowania warunkowego).
- **Poznasz kolekcję `ws.conditional_formatting`** dokładnie tak, jak działa w 3.1.x — z jej trzema niespodziankami: `add()` łączy reguły pod tym samym zakresem, **nie ma `clear()` ani usuwania pojedynczej reguły**, a `len()` liczy **zakresy**, nie reguły.
- **Zobaczysz, co openpyxl gubi przy zapisie** — reguły zapisane przez nowsze wersje Excela w formacie rozszerzeń (`extLst`) są niewidoczne i giną. To trzecia kategoria z modułu 03 w praktyce.

## 1. Intuicja i analogia

### 1.1. Analogia główna: lampa z czujnikiem ruchu

Wyobraź sobie korytarz w biurze.

**Wariant pierwszy.** Ktoś maluje całą podłogę na czerwono. Kolor jest **w podłodze**. Jeśli chcesz go zmienić, musisz podłogę przemalować — niezależnie od tego, kto po niej chodzi.

**Wariant drugi.** Ktoś montuje **lampę z czujnikiem ruchu**. Nic nie jest pomalowane. Lampa świeci się wtedy, gdy ktoś wejdzie. Kolor pojawia się **tylko w obecności człowieka**, a w korytarzu nie ma ani grama czerwonej farby.

Formatowanie warunkowe to wariant drugi.

> **W arkuszu nie ma zapisanego koloru. Jest zapisane *urządzenie*, które ma ten kolor wyświetlić, gdy warunek zostanie spełniony. Kolor pojawia się dopiero wtedy, gdy Excel otworzy plik i uruchomi regułę.**

I teraz konsekwencja, którą trzeba zapamiętać od razu, bo inaczej stracisz godzinę na szukanie nieistniejącego błędu:

**Po wczytaniu pliku z formatowaniem warunkowym `cell.fill` będzie puste.** Nie dlatego, że coś się zepsuło. Dlatego, że w komórce **nigdy nie było koloru**. Kolor jest w regule, obok komórki — jak lampa obok korytarza.

Doświadczysz tego na własne oczy w Przykładzie 1. To ta sama kategoria zdziwienia co w module 12: **tabela nie jest w komórkach, tylko obok nich.** Tutaj: **kolor nie jest w komórkach, tylko obok nich.**

### 1.2. Analogia: `DifferentialStyle` to poprawka długopisem na wydruku

Masz wydrukowaną stronę z raportu. Chcesz zaznaczyć jedno słowo. Czy przedrukowujesz całą stronę z tym słowem pogrubionym? Nie. **Bierzesz długopis i podkreślasz to jedno słowo.**

Tym właśnie jest `DifferentialStyle` — **stylem różnicowym**. Reguła formatowania warunkowego nie niesie pełnego opisu wyglądu komórki (czcionka, tło, obramowanie, wyrównanie, format liczby...). Niesie tylko **różnicę**:

> „*Jeśli warunek jest spełniony, to do tego, co komórka już ma, **dodaj**: pogrubienie i czerwone tło.*"

Komórka zachowuje swój normalny styl (np. czcionkę Calibri 11, format `#,##0.00 "zł"`), a reguła dokłada tylko to, co chce zmienić. Dokładnie jak długopis na wydruku — wydruk zostaje, dopisek jest nałożony z wierzchu.

Konsekwencja: **reguła nie „ustawia" wyglądu komórki, tylko go modyfikuje.** Jeśli komórka ma już czerwone tło, a reguła dokłada czerwone tło — komórka dalej ma czerwone tło (nie „zamienia" koloru). Jeśli reguła dokłada pogrubienie, a komórka już była pogrubiona — nic nowego się nie dzieje.

### 1.3. Analogia: wzór na kalce — formuła widzi tylko lewy-górny róg

To jest analogia, którą zapamiętaj, jeśli masz zapamiętać tylko jedną z tego modułu.

Kalka kreślarska. Kładziesz ją na rysunku i kopiujesz wzór. **Wzór piszesz dla pierwszego pola, a kalka przesuwa go na resztę pól.**

Formatowanie warunkowe działa tak samo. Gdy stosujesz regułę do zakresu `A2:E100`, formuła w regule jest **pisana tak, jakby istniała tylko komórka `A2`** — a Excel **przesuwa ją na wszystkie pozostałe komórki** zakresu, dokładnie tak, jakbyś skopiował formułę w prawo i w dół.

I tu jest cała trudność. Pisząc wzór na kalce, musisz **świadomie zdecydować**, które części mają się przesuwać, a które nie:

- `$D2` — kolumna D **zablokowana** (`$`), wiersz **przesuwny**. Przesuwając w prawo nic się nie zmienia; przesuwając w dół wiersz rośnie.
- `$D$2` — zablokowane jedno i drugie. Wzór zawsze patrzy na komórkę `D2`, nieważne gdzie go przesuniesz.
- `D2` — nic nie zablokowane. Wzór „ucieka" w prawo i w dół — prawie nigdy tego nie chcesz w formatowaniu warunkowym.

**Reguła kciuka:** dla „podświetl cały wiersz na podstawie wartości w kolumnie X" chcesz `$X` (kolumna zablokowana) i **przesuwny numer wiersza, zaczynając od pierwszego wiersza zakresu**. Zobaczysz to w 2.7 i w Przykładzie 2.

### 1.4. Analogia: priorytet to triage na izbie przyjęć

Na izbie przyjęć pacjentów nie przyjmuje się w kolejności zgłoszenia. Jest **triage** — najpierw ci, którzy wymagają natychmiastowej pomocy.

Reguły formatowania warunkowego też stoją w kolejce. Każda ma **numer priorytetu**. Excel sprawdza je **od najmniejszego numeru do największego** (uwaga: „mniejszy numer" = „wyższy priorytet" — tak, to odwrotnie niż intuicja).

I tu wchodzi `stopIfTrue`, czyli „**koniec, dalej nie sprawdzaj**". Jeśli reguła o wysokim priorytecie zadziałała i ustawisz `stopIfTrue = True`, Excel **przestaje rozpatrywać pozostałe reguły dla tej komórki**. Jak lekarz triage'owy, który mówi: „ten pacjent idzie od razu na stół, nie ma sensu go dalej diagnozować".

Bez `stopIfTrue` Excel sprawdzi wszystkie reguły i **połączy** ich formaty tam, gdzie się nie wykluczają — a tam, gdzie się wykluczają (np. dwa różne kolory tła), wygra reguła o wyższym priorytecie.

## 2. Teoria

### 2.1. Reguła mieszka w arkuszu, nie w komórce

Sformalizujmy analogię 1.1.

**Definicja robocza:** formatowanie warunkowe to zestaw **reguł** przypisanych do **zakresu komórek** w arkuszu. Każda reguła mówi: „jeśli warunek jest prawdziwy, nałóż na komórkę tę różnicę stylu". Reguły są częścią arkusza, nie komórek.

Trzy konsekwencje, które trzeba przyswoić:

**Konsekwencja 1 — brak koloru w komórce.** Nie ma czego odczytać przez `cell.fill`. Kolor żyje w regule i pojawia się **dopiero przy otwarciu pliku w Excelu**.

```python
# To NIE pokaze koloru z formatowania warunkowego - pokaze tylko styl komorki
print(ws["A1"].fill.fill_type)      # None
print(ws["A1"].font.bold)           # None
```

To jest klasyczna pomyłka diagnostyczna. Jeśli napiszesz skrypt „przeczytaj mi, które komórki są czerwone", i użyjesz `cell.fill` — dostaniesz pustkę i będziesz szukać błędu tam, gdzie go nie ma. **Koloru z formatowania warunkowego nie odczytasz z komórki, bo go tam nie ma.**

**Konsekwencja 2 — reguła jest „leniwa".** Nie musisz nic przeliczać i nic zapisywać w komórkach. Warunek jest tekstem, który Excel oceni przy otwarciu — dokładnie jak formuły z modułu 06. **openpyxl nie ocenia reguł** (nie ma silnika obliczeń), tylko je zapisuje.

**Konsekwencja 3 — reguła przetrwa zmianę wartości.** To jest siła tego mechanizmu. Komórka z wartością 150 jest podświetlona, bo reguła mówi „podświetl, jeśli > 100". Zmieniasz wartość na 250 — komórka **nadal jest podświetlona**, bo reguła nadal pasuje. Zmieniasz na 80 — podświetlenie **znika samo**. Nikt nic nie przemalowywał. To dokładnie lampa z czujnikiem ruchu.

### 2.2. Gdzie to mieszka w pliku

W module 02 nauczyłeś się, że `.xlsx` to ZIP z plikami XML. Formatowanie warunkowe zostawia w archiwum **dwa ślady**:

1. **Reguły** trafiają do arkusza — `xl/worksheets/sheetN.xml`, w elemencie `<conditionalFormatting>`:

```xml
<worksheet ...>
  <sheetData>...</sheetData>
  <conditionalFormatting sqref="A2:A100">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
</worksheet>
```

Zauważ trzy rzeczy:

- **`sqref="A2:A100"`** — zakres, do którego reguła należy. To ten „lewy-górny róg kalki" z analogii 1.3.
- **`dxfId="0"`** — **numer** stylu różnicowego, a nie sam styl. To ten sam wzorzec co w module 02: komórka trzyma numer stylu, nie styl.
- **`priority="1"`** — priorytet.

2. **Style różnicowe** trafiają do `xl/styles.xml`, do sekcji `<dxfs>`:

```xml
<dxfs count="1">
  <dxf>
    <font><b/></font>
    <fill><patternFill><bgColor rgb="FFEE1111"/></patternFill></fill>
  </dxf>
</dxfs>
```

**Wniosek, który ma znaczenie praktyczne:** reguły i style różnicowe to **osobne byty** połączone numerem (`dxfId`). Dokładnie ta sama architektura co style komórek w module 02 — i źródło tych samych pułapek. Jeśli usuniesz `<dxf>` z `styles.xml`, reguła zostanie, ale nie będzie miała czego nałożyć. Jeśli usuniesz regułę, dxf zostanie jako „sierota" (i tak bywa po edycji ręcznej).

### 2.3. Trzy rodziny reguł

Excel i openpyxl dzielą formatowanie warunkowe na trzy rodziny. Rozróżnienie ma znaczenie, bo każda rodzina ma **inny sposób tworzenia** w openpyxl.

| Rodzina | Co to | Reguły (wartość `type`) | Openpyxl |
|---|---|---|---|
| **Wbudowane** (builtin) | Wizualizacje: skale kolorów, paski danych, zestawy ikon | `colorScale`, `dataBar`, `iconSet` | `ColorScaleRule`, `DataBarRule`, `IconSetRule` — **moduł 14** |
| **Standardowe** | Porównanie komórki, tekst, daty, ranking | `cellIs`, `containsText`, `timePeriod`, `top10`, `aboveAverage`, `uniqueValues`, `duplicateValues`, ... | `CellIsRule`, `Rule` + parametry |
| **Własne** (custom) | Dowolne wyrażenie logiczne + dowolny styl różnicowy | `expression` | `FormulaRule`, `Rule(type="expression", dxf=...)` |

**Ten moduł zajmuje się rodziną standardową (`cellIs`) i własną (`expression`).** Rodzina wbudowana należy do modułu 14.

Pełna lista wartości `type`, które rozumie `Rule` (prosto ze źródła 3.1.x):

```
expression, cellIs, colorScale, dataBar, iconSet, top10,
uniqueValues, duplicateValues, containsText, notContainsText,
beginsWith, endsWith, containsBlanks, notContainsBlanks,
containsErrors, notContainsErrors, timePeriod, aboveAverage
```

Nie musisz tego pamiętać — ale dobrze wiedzieć, że „formatowanie warunkowe" to nie tylko „podświetl większe niż". To cała gałąź funkcji.

### 2.4. `DifferentialStyle` — różnica, nie pełny styl

Sformalizujmy analogię 1.2.

```python
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill

dxf = DifferentialStyle(
    font=Font(bold=True, color="9C0006"),
    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    border=None,
)
```

`DifferentialStyle` przyjmuje **cztery** składniki: `font`, `fill`, `border`, `numFmt`. Zauważ, że **nie ma tu `alignment`** — a w pełnym stylu komórki z modułu 09 było sześć składników. To nie przypadek: formatowanie warunkowe **nie zmienia wyrównania** ani ochrony. Zmienia wygląd plamki (kolor, pogrubienie, ramka, format liczby).

**Co znaczy „różnica" w praktyce:**

- `fill=PatternFill(...)` → **nakłada** tło. Jeśli komórka już ma tło, nadpisuje je **wizualnie** (reguła wygrywa w tym zakresie).
- `font=Font(bold=True)` → dodaje pogrubienie. **Ale uwaga:** jeśli podasz `Font(bold=True, color="9C0006")` z kolorem, a komórka miała inną czcionkę (nazwę, rozmiar), to **te ustawienia zostaną zachowane**, chyba że je nadpiszesz. Wynika to z natury stylu różnicowego.
- `border=Border(bottom=Side(style="thin", color="FF0000"))` → dodaje dolną ramkę.

Trzy pułapki specyficzne dla dxf, o których ostrzegam z góry:

1. **Kolor tła w dxf bywa kapryśny.** W normalnym stylu komórki (moduł 09) dla `fill_type="solid"` widoczny jest `fgColor` (czyli `start_color`). W stylach różnicowych zdarza się, że Excel pokazuje tło z drugiego koloru wzorca. **Rozwiązanie: zawsze podawaj `start_color` I `end_color` tą samą wartością, `fill_type="solid"`, i sprawdzaj wzrokowo w Excelu.** Jeśli tło się nie pojawia — zamień kolory miejscami i zobacz, który działa. To ta sama rodzina problemów co `fgColor`/`bgColor` z modułu 09, tylko w innym kontekście.

2. **Pusty `fill` = niewidoczny kolor.** `PatternFill()` bez `fill_type` i bez koloru nie maluje nic. Musisz podać jawnie `fill_type="solid"` i kolor.

3. **`dxf` jest opcjonalne, ale wtedy reguła nic nie robi.** Reguła z samym warunkiem, bez stylu, nie ma czego nałożyć. `FormulaRule` i `CellIsRule` **wymagają** co najmniej jednego składnika stylu, żeby cokolwiek było widać.

### 2.5. Priorytety — mniejsza liczba znaczy ważniejsza

Sformalizujmy analogię 1.4. Trzy fakty, które trzeba znać:

**Fakt 1: priorytet jest globalny dla arkusza, nie dla zakresu.** Numer `1` oznacza „najważniejsza reguła na tym **arkuszu**" — nie „najważniejsza w tym zakresie". To ma znaczenie, gdy reguły na różnych, nachodzących zakresach ze sobą konkurują.

**Fakt 2: `add()` nadaje priorytet automatycznie.** W źródle 3.1.x:

```python
def add(self, range_string, cfRule):
    ...
    rule = cfRule
    self.max_priority += 1
    if not rule.priority:                 # domyslne priority=0 -> falszywe
        rule.priority = self.max_priority
    self._cf_rules.setdefault(cf, []).append(rule)
```

Czyli każda dodana reguła dostaje **kolejny wolny numer** — pierwsza `1`, druga `2`, trzecia `3`... Reguła zachowuje nadany priorytet tylko wtedy, gdy **już jakiś ma** (niezerowy).

**Fakt 3: konstruktory pomocnicze NIE przyjmują `priority`.** To ważne i nieoczywiste:

```python
# To nie zadziala - CellIsRule i FormulaRule nie maja parametru priority:
# CellIsRule(operator="greaterThan", formula=["100"], priority=1)   # TypeError!

# Priorytet ustawiasz PO utworzeniu:
rule = CellIsRule(operator="greaterThan", formula=["100"], fill=red)
rule.priority = 1
```

Wyjątkiem jest bazowy `Rule`, który **ma** parametr `priority` w konstruktorze. Ale i tak najwygodniej: **dodaj reguły w takiej kolejności, w jakiej mają być rozpatrywane** — wtedy priorytety nadadzą się same i będą sensowne.

**Konsekwencja praktyczna:** jeśli chcesz mieć regułę zawsze sprawdzaną jako pierwszą (np. wykluczającą), **dodaj ją jako pierwszą**. Jeśli chcesz ją „opóźnić", dodaj później albo ustaw `priority` ręcznie po utworzeniu.

### 2.6. `stopIfTrue` — „koniec, dalej nie sprawdzaj"

```python
rule = FormulaRule(formula=['$D2="Pilne"'], fill=red, stopIfTrue=True)
```

Bez `stopIfTrue` Excel sprawdza **wszystkie** reguły dla komórki i łączy ich formaty tam, gdzie się nie wykluczają. Przy konflikcie tego samego atrybutu (np. dwa różne tła) wygrywa reguła o **niższym** numerze priorytetu.

Z `stopIfTrue = True`: gdy reguła zadziała, Excel **przestaje rozpatrywać kolejne** dla tej komórki. Jej format jest ostateczny.

**Kiedy to naprawdę ma znaczenie — i kiedy nie:**

- Jeśli Twoje reguły są **rozłączne** (np. „po terminie", „dziś", „za mniej niż 3 dni" — komórka nie może spełniać dwóch naraz), `stopIfTrue` **nie zmienia wyniku**. Możesz go pominąć.
- Jeśli reguły **mogą się nakładać** (np. „kwota > 10 000" i „wiersz pilny"), `stopIfTrue` na tej ważniejszej **rozstrzyga** konflikt jasno i przewidywalnie.

To jest ważna, uczciwa uwaga: wielu autorów wrzuca `stopIfTrue=True` wszędzie „dla bezpieczeństwa", nie rozumiejąc, że przy rozłącznych warunkach to placebo. **Rozumiej, kiedy to jest potrzebne, a kiedy nie.**

### 2.7. `sqref` i największa pułapka: formuła dla lewego-górnego rogu

Wróćmy do analogii 1.3 — to najważniejsza część tego modułu.

**Zakres reguły** jest przechowywany w atrybucie `sqref` (od *square reference*). W openpyxl widzisz go jako `cf.sqref` (obiekt `MultiCellRange`). Uwaga na nazwę, bo jest niespójna jak w module 12:

- `sqref` to `MultiCellRange` — może zawierać **wiele** zakresów rozdzielonych spacją: `"A1:A10 C1:C10"`. To odpowiada trzymaniu kilku „wycinków" pod jedną regułą w Excelu.
- `cf.cells` to **alias** `sqref` — to samo.
- `cf.rules` to **alias** `cf.cfRule` — lista reguł dla tego zakresu.

A teraz **pułapka formuły**. Gdy piszesz regułę z formułą dla zakresu, formuła jest interpretowana **jakby dotyczyła tylko lewej-górnej komórki** tego zakresu, a potem „przesuwana" na cały zakres (jak kalka).

**Przykład poprawny.** Zakres `A2:E100`. Chcesz podświetlić cały wiersz, gdy w kolumnie D jest „Pilne". Piszesz formułę dla lewego-górnego rogu (`A2`) i blokujesz kolumnę:

```python
formula=['$D2="Pilne"']      # $D = kolumna zablokowana, 2 = wiersz startowy zakresu
```

Przesuwając na wiersz 3: `$D3`. Na wiersz 50: `$D50`. Przesuwając w prawo (kolumna E): nadal `$D2` (kolumna zablokowana). **Efekt: cały wiersz jest podświetlony, gdy `D` w tym wierszu to „Pilne".** Dokładnie to chciałeś.

**Przykład błędny nr 1** — zapomniany `$`:

```python
formula=['D2="Pilne"']       # kolumna NIE zablokowana
```

Przesuwając w prawo „ucieka": `D2`, `E2`, `F2`... Wiersz podświetli się tylko wtedy, gdy **każda kolejna kolumna** też ma „Pilne" — czyli praktycznie nigdy. Objaw: „reguła nie działa". Przyczyna: brak `$`.

**Przykład błędny nr 2** — zablokowany wiersz, ale odniesienie do złego wiersza:

```python
formula=['$D$2="Pilne"']     # wiersz tez zablokowany
```

Formuła **zawsze** patrzy na `D2`, nieważne w którym wierszu jesteś. Objaw: „podświetlił się cały zakres, bo `D2` było «Pilne»" — albo „nie podświetlił się żaden, bo `D2` nie było «Pilne»". Wiersz w formule **musi być względny** do zakresu.

**Przykład błędny nr 3** — odniesienie do wiersza poza zakresem:

```python
# zakres A2:E100, a formula zaczyna sie od $D1
formula=['$D1="Pilne"']      # ⚠️ wiersz 1, a zakres zaczyna sie od wiersza 2
```

Formuła odnosi się do wiersza **nad** zakresem. Excel przesunie ją o jeden „za daleko" — wynik będzie o wiersz przesunięty. Objaw: „podświetlone są złe wiersze, jakby o jeden w górę". **Zasada: formuła zaczyna się od pierwszego wiersza zakresu.**

**Zasada, którą warto zapisać na kartce:**

> Formułę w regule pisz tak, **jakby istniał tylko lewy-górny róg zakresu**. Kolumny, które mają być stałe — zablokuj (`$D`). Numer wiersza — **względny** i **od pierwszego wiersza zakresu**. Potem zweryfikuj w Excelu.

**Uczciwa uwaga.** Dokumentacja openpyxl podaje przykład `formula=['$A2="Microsoft"']` dla zakresu `A1:C10`, w którym odniesienie celuje w wiersz **2**, a nie w wiersz 1 — i opis wokół tego przykładu jest wewnętrznie niespójny (tekst mówi o kolumnie D, kod o kolumnie A). To dokładnie ta kategoria, o której mówi moduł 03: **dokumentacja i kod nie zawsze mówią to samo.** Nie kopiuj tego przykładu z pamięci — przetestuj własny w Excelu, bo tam widać prawdę. Najbezpieczniejsza praktyka, którą stosuj: **formuła zaczyna się od pierwszego wiersza `sqref`.**

### 2.8. `CellIsRule` — porównanie komórki z wartością

`CellIsRule` tworzy regułę typu `cellIs`: porównuje wartość każdej komórki z jednym lub dwoma **stałymi** (nieodwołującymi się do innych komórek).

```python
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

red_fill = PatternFill(start_color="EE1111", end_color="EE1111", fill_type="solid")

# Podswietl komorki wieksze niz 100
ws.conditional_formatting.add(
    "A2:A100",
    CellIsRule(operator="greaterThan", formula=["100"], fill=red_fill),
)
```

**Parametry `CellIsRule`:**

| Parametr | Znaczenie |
|---|---|
| `operator` | Nazwa operatora: `"greaterThan"`, `"lessThan"`, `"between"`, `"equal"`, `"notEqual"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`, `"notBetween"`, `"containsText"`, `"notContains"`, `"beginsWith"`, `"endsWith"` |
| `formula` | **Lista napisów.** Dla progów jednym: `["100"]`. Dla `between`/`notBetween`: dwa elementy: `["1", "5"]` |
| `stopIfTrue` | `True`/`False`/`None` (2.6) |
| `font`, `border`, `fill` | Składniki `DifferentialStyle` (2.4) |

**Wygodne rozszerzenie, które warto znać:** `CellIsRule` przyjmuje też **symbole** zamiast nazw. Mapowanie prosto ze źródła:

```python
expand = {">": "greaterThan", ">=": "greaterThanOrEqual",
          "<": "lessThan", "<=": "lessThanOrEqual",
          "=": "equal", "==": "equal", "!=": "notEqual"}
```

Czyli `operator=">"` zadziała tak samo jak `operator="greaterThan"`. Wygodne dla pośpiechu, ale w kodzie produkcyjnym czytelniejsza jest nazwa.

**Kiedy `CellIsRule`, a kiedy `FormulaRule`?** Kryterium jest proste:

> **`CellIsRule` porównuje komórkę ze STAŁĄ. Jeśli próg zależy od innej komórki albo od kombinacji warunków — użyj `FormulaRule`.**

`CellIsRule(operator="greaterThan", formula=["100"])` — próg `100` jest **stały** w całym zakresie. Jeśli chcesz „większe niż wartość w komórce `C1`", musisz użyć `FormulaRule(formula=['A2>$C$1'])`.

### 2.9. `FormulaRule` — najpotężniejsze narzędzie

`FormulaRule` tworzy regułę typu `expression` (bądź `cellIs` w niektórych ręcznych konstrukcjach). Jej warunek to **dowolne wyrażenie Excela** — czyli możesz zrobić wszystko:

- podświetlić cały wiersz na podstawie innej kolumny,
- porównać dwie kolumny między sobą,
- sprawdzić daty względem `TODAY()`,
- połączyć warunki przez `AND`/`OR`/`NOT`,
- sprawdzić tekst przez `SEARCH`/`ISNUMBER`.

```python
from openpyxl.formatting.rule import FormulaRule

# Podswietl CALY WIERSZ, gdy w kolumnie D jest "Pilne"
ws.conditional_formatting.add(
    "A2:E100",
    FormulaRule(formula=['$D2="Pilne"'], fill=red_fill, stopIfTrue=True),
)
```

**Trzy rzeczy, o których trzeba pamiętać z `FormulaRule`:**

**1. `formula` to lista napisów, a napis NIE zaczyna się od `=`.**

To pierwsza pułapka z briefu i trzeba ją powiedzieć wyraźnie:

```python
FormulaRule(formula=['A2>100'], ...)     # ✅ poprawnie
FormulaRule(formula=['=A2>100'], ...)    # ❌ Excel odrzuci - znak = nie moze byc tutaj
```

Dlaczego? Bo `CellIsRule`/`Rule` zapisują ten napis **dosłownie** do XML jako treść `<formula>`. Ten element nie jest „formułą w komórce" — to „wyrażenie warunku". Excel oczekuje tam samego wyrażenia. Gdy wstawisz `=`, Excel zobaczy `=A2>100` jako treść formula i potraktuje ją jak coś nieparsowalnego (albo zapisze dziwną wartość). **To skrajnie częsty błąd, bo `=` jest naturalnym odruchem.**

**2. Cudzysłowy wewnątrz napisu.** Formuła `$D2="Pilne"` zawiera cudzysłowy. W Pythonie pisz ją w apostrofach:

```python
formula=['$D2="Pilne"']        # ✅ apostrofy na zewnatrz, cudzyslowy w srodku
formula=["$D2=\"Pilne\""]      # ✅ zadziala, ale brzydko
```

**3. To samo `sqref` i ta sama pułapka względności** co w 2.7. `FormulaRule` to właśnie ten typ reguły, dla którego kalka z analogii 1.3 jest krytyczna.

**Katalog praktycznych formuł** — to zbiór, do którego będziesz wracać:

| Cel | Formuła (bez `=`) |
|---|---|
| Wartość większa niż w innej komórce | `'A2>$C$1'` |
| Cały wiersz na podstawie kolumny | `'$D2="Pilne"'` |
| Porównanie dwóch kolumn | `'$D2>$E2'` |
| Data minęła | `'$C2<TODAY()'` |
| Data dziś | `'$C2=TODAY()'` |
| Data w najbliższych 3 dniach | `'AND($C2>TODAY(),$C2<TODAY()+3)'` |
| Komórka pusta | `'ISBLANK($A2)'` |
| Zawiera tekst (bez wielkości liter) | `'ISNUMBER(SEARCH("pilne",$B2))'` |
| Warunek na dwóch kolumnach | `'AND($C2="Pilne",$D2>10000)'` |
| Wielokrotność / parzystość | `'MOD(ROW(),2)=0'` |
| Pusta wartość LUB zero | `'OR(ISBLANK($A2),$A2=0)'` |

To wystarcza na 90% realnych potrzeb. Reszta to komponowanie tych klocków.

### 2.10. `Rule` + `DifferentialStyle` ręcznie

`CellIsRule` i `FormulaRule` to **wygodne skróty**. Każdy z nich tworzy wewnętrznie obiekt `Rule` i wypełnia `dxf`. Gdy potrzebujesz nietypowego zestawu (np. reguła `containsText` albo własny typ), budujesz `Rule` ręcznie:

```python
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill

# Podswietl komorki zawierajace tekst "pilne"
dxf = DifferentialStyle(
    font=Font(color="9C0006"),
    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
)
rule = Rule(type="containsText", operator="containsText", text="pilne", dxf=dxf)
rule.formula = ['NOT(ISERROR(SEARCH("pilne",A2)))']   # opcjonalnie - wlasne wyrazenie
ws.conditional_formatting.add("A2:A100", rule)
```

Zwróć uwagę: **`Rule` wymaga `type` jako pierwszego argumentu** i przyjmuje `dxf` bezpośrednio. `formula` możesz ustawić po utworzeniu (jak w tym przykładzie). To rozwiązanie dla przypadków, których nie pokrywają skróty.

**Kiedy ręczny `Rule`?** Gdy potrzebujesz typu, dla którego openpyxl nie ma konstruktora pomocniczego (np. `top10`, `aboveAverage`, `timePeriod`, `containsBlanks`...) albo gdy chcesz **reuse** jednego `dxf` w wielu regułach (moduł 09 — mniej obiektów, mniej alokacji).

### 2.11. Kolekcja `ws.conditional_formatting` — API i trzy niespodzianki

`ws.conditional_formatting` to obiekt `ConditionalFormattingList`. W 3.1.x działa jak uporządkowany słownik (bazuje na `OrderedDict`), ale ma kilka zachowań, które zaskakują.

**Dodawanie:**

```python
ws.conditional_formatting.add("A2:A100", rule)     # ✅ jedyna metoda dodawania
```

**Niespodzianka 1 — `len()` liczy ZAKRESY, nie reguły.**

```python
ws.conditional_formatting.add("A2:A100", rule1)
ws.conditional_formatting.add("A2:A100", rule2)    # ten sam zakres!
ws.conditional_formatting.add("C2:C100", rule3)    # inny zakres

print(len(ws.conditional_formatting))              # 2, NIE 3!
```

Dlaczego? Bo `ConditionalFormattingList` kluczuje reguły **po zakresie** (`sqref`). Dwie reguły na tym samym zakresie lądują **pod jednym kluczem**, w jednej liście. `len()` zwraca liczbę **grup zakresów**.

**To jest zarazem zachowanie ukryte i użyteczne:** wywołanie `add` z tym samym zakresem **nie tworzy drugiej grupy** — dołącza regułę do istniejącej. To wynika z tego, że `ConditionalFormatting.__eq__` i `__hash__` bazują na `sqref`:

```python
def __eq__(self, other):
    return self.sqref == other.sqref
def __hash__(self):
    return hash(self.sqref)
```

Konsekwencja praktyczna: jeśli **chcesz dwie rozłączne reguły** na tej samej kolumnie, i tak trafią do jednej grupy — i to jest w porządku, bo i Excel tak je zapisuje. Ale jeśli liczysz reguły przez `len(ws.conditional_formatting)`, dostaniesz liczbę zakresów. **Reguły licz przez iterację** (zobacz Przykład 4).

**Niespodzianka 2 — iteracja daje obiekty `ConditionalFormatting`, nie reguły.**

```python
for cf in ws.conditional_formatting:       # cf to ConditionalFormatting (grupa zakr.)
    print(cf.sqref)                        # MultiCellRange, np. "A2:A100"
    for rule in cf.rules:                  # dopiero tu masz reguly
        print(rule.type, rule.priority, rule.formula)
```

Czyli iteracja jest **dwupoziomowa**: grupa zakresu → lista reguł. To ta sama struktura co w XML z 2.2: `<conditionalFormatting sqref="..."><cfRule/>...`.

**Niespodzianka 3 — nie ma `clear()` ani usuwania pojedynczej reguły.**

Kolekcja ma tylko `add`. Żeby usunąć:

```python
# Usun CALA grupe zakresu (wszystkie reguly tego zakresu):
del ws.conditional_formatting["A2:A100"]      # ✅ dziala (__delitem__)

# Wyzeruj WSZYSTKIE reguly arkusza - przypisz nowa, pusta kolekcje:
from openpyxl.formatting.formatting import ConditionalFormattingList
ws.conditional_formatting = ConditionalFormattingList()      # ✅

# Usun POJEDYNCZA regule (np. tylko dataBar) - nie ma metody!
# Rozwiazanie: przebuduj kolekcje, zostawiajac to, co chcesz zachowac:
```

Wzorzec **filtrowania reguł** (rebuild) jest ważny, bo zastępuje brak usuwania. Zobaczysz go w Przykładzie 4.

**Dostęp do reguł konkretnego zakresu:**

```python
lista_regul = ws.conditional_formatting["A2:A100"]    # lista reguł (albo KeyError)
```

### 2.12. Reguły a wydajność

Reguły są **malutkie w pliku** — to kilka wierszy XML. Cały koszt jest po stronie **Excela**, który przy każdym otwarciu i każdej zmianie danych **ocenia reguły** dla wszystkich komórek, których dotyczą.

**Reguła kciuka:** koszt ≈ (liczba komórek w `sqref`) × (liczba reguł na tym zakresie).

Z tego wynikają trzy praktyczne zalecenia:

1. **Nie stosuj reguł na całych kolumnach** (`A:A`, `A1:A1048576`), jeśli dane kończą się na wierszu 1000. Ogranicz `sqref` do **realnego obszaru danych** — dokładnie jak w module 11 przy `print_area` i module 12 przy `ref` tabeli. **Zakres reguły musi być wyliczony z danych, nie z założenia.**

2. **Nie mnoż reguł bez potrzeby.** „Podświetl > 100", „> 200", „> 300" na tym samym zakresie — to trzy reguły do oceny. Często lepiej użyć jednej skali kolorów (moduł 14) albo przenieść logikę do danych.

3. **100 000 komórek z regułą formułową da wyraźnie wolniejszy plik.** Objaw: użytkownik otwiera raport i czeka, przewija i widzi „przycinanie się" Excela. Przy danych tej skali rozważ: policzyć w Pythonie i wstawić gotowe kolumny (moduł 06 — „Excel to warstwa prezentacji, nie silnik"), albo zawęzić zakres.

### 2.13. Co openpyxl gubi przy zapisie

To trzecia kategoria z modułu 03 — „openpyxl gubi dane przy zapisie" — w praktyce.

**openpyxl modeluje klasyczne reguły** (`cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `top10`, ...). Jeśli je **tworzysz** albo **wczytujesz i zapisujesz**, przechodzą.

**Ale nowsze wersje Excela zapisują część reguł w formacie rozszerzeń** (`extLst`) — a to openpyxl **nie modeluje i gubi przy zapisie**. Konkretnie giną reguły oparte na rozszerzeniach, np.:

- niektóre warianty pasków danych (z gradientem, z ustawieniami dla wartości ujemnych),
- nowsze zestawy ikon z własnymi progami,
- nietypowe kombinacje w rodzaju „ikona + wartość" w formacie x14.

Objaw: wczytujesz plik z Excela, który ma bogate formatowanie warunkowe, zapisujesz go openpyxl, otwierasz — **część reguł zniknęła bez ostrzeżenia**.

**Wniosek praktyczny (i to jest wniosek także z modułu 18):** jeśli plik był tworzony ręcznie w Excelu i ma bogate formatowanie warunkowe, **nie przepuszczaj go bezmyślnie przez openpyxl typu wczytaj-zapisz**. Albo generuj nowy plik od zera (wtedy wiesz, co jest w środku), albo wczytaj dane i zbuduj formatowanie od nowa, albo sprawdź po zapisie, czy reguły przetrwały.

**Jak sprawdzić, co przetrwało:** funkcja „inwentarz reguł" z Przykładu 4 — uruchom ją **przed** i **po** zapisie i porównaj liczbę reguł i ich typy. To dokładnie ta sama technika co „scan-before-save" z module 12.

## 3. Przykłady krok po kroku

### Przykład 1 — podświetl wartości większe niż 100 (🟢)

Pierwszy przykład pokazuje **kompletny cykl**: utwórz → zapisz → wczytaj → sprawdź. I dowodzi kluczowej tezy z 2.1: **po wczytaniu `cell.fill` jest puste**.

```python
"""Podstawowe formatowanie warunkowe + dowod, ze koloru nie ma w komorce.

Uruchom:  python examples/13_cf_podstawy.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "13_cf_podstawy.xlsx"

DANE = [45, 120, 88, 250, 99, 101, 30, 175, 60, 140]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Wyniki"

    ws.append(["Lp.", "Wynik"])
    for indeks, wynik in enumerate(DANE, start=1):
        ws.append([indeks, wynik])

    # --- styl roznicowy: co dodac, gdy warunek spelniony (2.4) -----
    czerwone_tlo = PatternFill(start_color="EE1111", end_color="EE1111", fill_type="solid")
    biala_czcionka = Font(bold=True, color="FFFFFF")

    # --- regula: podswietl komorki z kolumny B wieksze niz 100 ----
    # UWAGA: formula to LISTA napisow; operator moze byc nazwa albo symbolem (2.8)
    regula = CellIsRule(
        operator="greaterThan",
        formula=["100"],
        fill=czerwone_tlo,
        font=biala_czcionka,
        stopIfTrue=True,
    )
    ws.conditional_formatting.add("B2:B11", regula)

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Wyniki"]

        print("=" * 82)
        print("KLUCZOWY DOWOD: KOLORU NIE MA W KOMORCE")
        print("=" * 82)
        # B2 to wartosc 45 (nie podswietlona), ale nawet podswietlone komorki
        # mialyby tu puste fill - bo kolor zyje w REGULE, nie w komorce (2.1)
        print(f"  ws['B2'].fill.fill_type = {ws['B2'].fill.fill_type!r}   <- None!")
        print(f"  ws['B3'].fill.fill_type = {ws['B3'].fill.fill_type!r}   <- nadal None!")
        print("  Wniosek: koloru z formatowania warunkowego NIE odczytasz z komorki.")
        print("  Kolor pojawi sie dopiero, gdy Excel otworzy plik i oceni reguly.")
        print()

        print("=" * 82)
        print("GDZIE NAPRAWDE JEST REGULA")
        print("=" * 82)
        # len() liczy ZAKRESY (grupy), nie reguly (2.11, niespodzianka 1)
        print(f"  len(ws.conditional_formatting) = {len(ws.conditional_formatting)}  (grup zakresow)")
        for cf in ws.conditional_formatting:
            print(f"  zakres (sqref): {cf.sqref}")
            for regula in cf.rules:
                print(f"    typ:       {regula.type}")            # cellIs
                print(f"    operator:  {regula.operator}")        # greaterThan
                print(f"    formula:   {regula.formula}")         # ['100']
                print(f"    priority:  {regula.priority}")        # 1
                print(f"    stopIfTrue:{regula.stopIfTrue}")
                fill = regula.dxf.fill
                kolor = getattr(fill.fgColor, "rgb", None) if fill is not None else None
                print(f"    dxf fill:  {kolor}")
    finally:
        wb.close()

    print()
    print("=" * 82)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 82)
    print("  1. Komorki B z wartosciami > 100 (120, 250, 101, 175, 140) sa czerwone")
    print("     i pogrubione; pozostale nie.")
    print("  2. Kliknij B3 (120) -> Narzedzia główne -> Formatowanie warunkowe ->")
    print("     Zarządzaj regułami: zobaczysz regule 'Wieksze niz 100'.")
    print("  3. ZMIEN wartosc w B2 z 45 na 500. Komorka SAMA sie podswietli.")
    print("     To dowod na to, ze to REGULA, a nie zapisany kolor (2.1).")
    print("  4. ZMIEN ja z powrotem na 10. Podswietlenie zniknie samo.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Po `wb.save` komórki mają tylko swoje zwykłe wartości. `ws.conditional_formatting` trzyma jedną grupę (zakres `B2:B11`) z jedną regułą w środku. Reguła ma `type="cellIs"`, `operator="greaterThan"`, `formula=['100']` i `dxf` (czerwone tło + biała czcionka).

**Co trafi do pliku.** Do `xl/worksheets/sheet1.xml` trafi:

```xml
<conditionalFormatting sqref="B2:B11">
  <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1">
    <formula>100</formula>
  </cfRule>
</conditionalFormatting>
```

A do `xl/styles.xml` — pierwszy styl różnicowy (`dxfId="0"`) z pogrubioną czcionką i czerwonym tłem. **Nigdzie w komórkach nie ma koloru.**

**Trzy rzeczy do przemyślenia:**

1. **Nagłówek `B1` nie podlega regule** — bo zakres to `B2:B11`. To celowe: nie chcesz, żeby nagłówek się podświetlił. **Zakres reguły = obszar danych, nie cały arkusz** (2.12).
2. **`len()` zwraca `1`, nie liczbę reguł** — bo jest jedna grupa zakresu. To niespodzianka 1 z 2.11 w działaniu.
3. **`regula.dxf.fill.fgColor.rgb`** — po wczytaniu `dxf` jest rozwiązane (openpyxl mapuje `dxfId` na obiekt). To pozwala czytać kolor reguły — zobaczysz to w Przykładzie 4.

### Przykład 2 — podświetl cały wiersz na podstawie kolumny (🟡)

To ćwiczenie z briefu i zarazem **test zrozumienia analogii kalki z 1.3**.

```python
"""Podswietl CALY WIERSZ, gdy w kolumnie D jest 'Pilne'.

Formuła na kalce: $D2 - kolumna zablokowana, wiersz wzgledny od pierwszego
wiersza zakresu (2.7).

Uruchom:  python examples/13_cf_caly_wiersz.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "13_cf_caly_wiersz.xlsx"

DANE = [
    ("F-001", "Marcinkowski", "2026-01-04", "Pilne"),
    ("F-002", "Nowak",        "2026-01-05", "Normalne"),
    ("F-003", "Kowalska",     "2026-01-06", "Pilne"),
    ("F-004", "Wiśniewski",   "2026-01-07", "Normalne"),
    ("F-005", "Lewandowski",  "2026-01-08", "Pilne"),
    ("F-006", "Zielińska",    "2026-01-09", "Normalne"),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Faktury"

    ws.append(["Numer", "Klient", "Rok", "Status"])
    for numer, klient, rok, status in DANE:
        ws.append([numer, klient, rok, status])

    zolte_tlo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # --- regula na CALY wiersz: zakres A2:D7, warunek opiera sie na kolumnie D
    # Klucz: $D (kolumna zablokowana) + 2 (pierwszy wiersz zakresu) (2.7)
    regula = FormulaRule(
        formula=['$D2="Pilne"'],       # bez znaku = na poczatku!
        fill=zolte_tlo,
        stopIfTrue=True,
    )
    ws.conditional_formatting.add("A2:D7", regula)

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Faktury"]
        for cf in ws.conditional_formatting:
            print(f"  zakres:  {cf.sqref}")
            for regula in cf.rules:
                print(f"  typ:     {regula.type}")       # expression
                print(f"  formula: {regula.formula}")    # ['$D2="Pilne"']
                print(f"  priority:{regula.priority}")
    finally:
        wb.close()

    print()
    print("=" * 82)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 82)
    print("  Wiersze z 'Pilne' w kolumnie D (F-001, F-003, F-005) sa cale rozowe -")
    print("  od kolumny A do D. Zwroc uwage: warunek opiera sie na D, ale format")
    print("  nakłada sie na A:D. To wlasnie robi $D2 przy zakresie A2:D7.")
    print()
    print("  ZMIEN w D2 'Pilne' na 'Normalne' -> caly wiersz traci kolor.")
    print("  To znaczy, ze reguła patrzy na kolumne D (zablokowana: $), a wiersz")
    print("  sie przesuwa (wzgledny: brak $ przy 2).")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Czym się to różni od `CellIsRule`:** tu warunek **nie porównuje samej komórki ze stałą** — porównuje **kolumnę D z tekstem**, ale format nakłada na **cały wiersz**. Dlatego `CellIsRule` tu nie wystarczy (bo porównywałby każdą komórkę zakresu ze stałą osobno). Musisz użyć `FormulaRule`.

**Dlaczego to działa — prześledźmy kalkę (2.7).** Zakres `A2:D7`, formuła `$D2="Pilne"`:

| Komórka | Co Excel oblicza | Wynik |
|---|---|---|
| `A2` (lewy-górny róg) | `$D2="Pilne"` | `Pilne` → PRAWDA, wiersz się maluje |
| `B2` | `$D2="Pilne"` (kolumna `$D`) | to samo, maluje się |
| `D2` | `$D2="Pilne"` | to samo |
| `A3` | `$D3="Pilne"` (wiersz przesunięty!) | patrzy na D3 = „Normalne" → fałsz |
| `A5` | `$D5="Pilne"` | D5 = „Pilne" → PRAWDA, maluje się |

**Błędna wersja — do obejrzenia i zrozumienia:**

```python
formula=['D2="Pilne"']         # bez $ - kolumna ucieka w prawo
# A2: D2="Pilne" -> patrzy na D! (ale to kolumna D, wiec na A2 to przypadek)
# B2: E2="Pilne" -> patrzy na E (pusta) -> falsz
# C2: F2="Pilne" -> falsz
# D2: G2="Pilne" -> falsz
# Efekt: pomalowana bedzie tylko kolumna A (i to przez przypadek)
```

To jest ten klasyczny objaw „reguła prawie działa". Bez `$` wynik wygląda niemal sensownie (coś się podświetla), ale jest zły — dokładnie o tym ostrzega analogia kalki.

### Przykład 3 — semafor dla terminów (🔴)

Ćwiczenie z briefu, doprowadzone do końca: **trzy reguły**, priorytety, `stopIfTrue` — plus uczciwe wyjaśnienie, **kiedy priorytet naprawdę rozstrzyga**.

```python
"""Semafor terminow: po terminie / dzis / za mniej niz 3 dni.

Trzy reguly na tym samym zakresie -> trzy reguly w JEDNEJ grupie (2.11).

Uruchom:  python examples/13_cf_semafor.py
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "13_cf_semafor.xlsx"


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Zadania"

    # Daty wzgledem "dzis", zeby semafor bylo widac od razu
    dzis = date.today()
    zadania = [
        ("Z-001", "Audyt wewnętrzny",    dzis - timedelta(days=5)),   # po terminie
        ("Z-002", "Raport kwartalny",    dzis),                        # dzis
        ("Z-003", "Spotkanie z klientem", dzis + timedelta(days=2)),   # < 3 dni
        ("Z-004", "Przegląd umów",       dzis + timedelta(days=10)),   # daleko
        ("Z-005", "Szkolenie BHP",       dzis - timedelta(days=1)),    # po terminie
        ("Z-006", "Inwentaryzacja",      dzis + timedelta(days=1)),    # < 3 dni
    ]

    ws.append(["Numer", "Zadanie", "Termin"])
    for numer, zadanie, termin in zadania:
        ws.append([numer, zadanie, termin])
        ws.cell(row=ws.max_row, column=3).number_format = "yyyy-mm-dd"

    # --- trzy style roznicowe -------------------------------------
    czerwone = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
    pomaranczowe = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    zolte = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    ZAKRES = "A2:C7"        # pierwszy wiersz danych = 2 (naglowek w 1)

    # --- regula 1 (dodana pierwsza = priorytet 1): PO TERMINIE ----
    # $C2 < TODAY() - kolumna zablokowana, wiersz od pierwszego wiersza zakresu
    ws.conditional_formatting.add(
        ZAKRES,
        FormulaRule(formula=["$C2<TODAY()"], fill=czerwone, stopIfTrue=True),
    )

    # --- regula 2 (priorytet 2): DZIS ----------------------------
    ws.conditional_formatting.add(
        ZAKRES,
        FormulaRule(formula=["$C2=TODAY()"], fill=pomaranczowe, stopIfTrue=True),
    )

    # --- regula 3 (priorytet 3): ZA MNIEJ NIZ 3 DNI -----------------
    # AND($C2>TODAY(), ...) wyklucza "dzis" - warunki sa rozlaczne (2.6)
    ws.conditional_formatting.add(
        ZAKRES,
        FormulaRule(
            formula=["AND($C2>TODAY(),$C2<TODAY()+3)"],
            fill=zolte,
            stopIfTrue=True,
        ),
    )

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Zadania"]
        print(f"  grup zakresow: {len(ws.conditional_formatting)}  <- jedna grupa (2.11)")
        for cf in ws.conditional_formatting:
            print(f"  zakres: {cf.sqref}")
            print(f"  reguł w grupie: {len(cf.rules)}  <- trzy reguly")
            # Sortuj po priorytecie, zeby pokazac kolejnosc rozpatrywania
            for regula in sorted(cf.rules, key=lambda r: r.priority):
                print(f"    priorytet {regula.priority}: {regula.formula}")
    finally:
        wb.close()

    print()
    print("=" * 82)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 82)
    print("  Z-001, Z-005 (po terminie)     -> czerwone")
    print("  Z-002 (dzis)                   -> pomaranczowe")
    print("  Z-003, Z-006 (za 1-2 dni)      -> zolte")
    print("  Z-004 (za 10 dni)              -> bez koloru")
    print()
    print("  Formatowanie warunkowe -> Zarządzaj regułami: zobaczysz TRZY reguly")
    print("  w tej samej grupie, w kolejnosci priorytetow 1, 2, 3.")
    print()
    print("  UWAGA: te trzy warunki sa ROZLACZNE - zadna komorka nie pasuje do")
    print("  dwoch naraz. Dlatego stopIfTrue NIC tu nie zmienia. Zeby zobaczyc,")
    print("  kiedy priorytet naprawde rozstrzyga, dodaj regule 'kwota' (patrz nizej).")


def pokaz_nakladanie() -> None:
    """Uczy, kiedy priorytet i stopIfTrue NAPRAWDE rozstrzygaja (2.6).

    Dodajemy regule nakladajaca sie na semafor: 'podswietl na fioletowo, gdy
    zadanie zawiera slowo Audyt'. Teraz Z-001 pasuje i do 'po terminie'
    (priorytet 1), i do 'audyt' (priorytet 4). Ponizsza regula musi
    ZDECYDOWAC, ktory kolor wygra.
    """
    print()
    print("=" * 82)
    print("KIEDY PRIORYTET NAPRAWDE ROZSTRZYGA")
    print("=" * 82)
    print("  Bez stopIfTrue, przy nakladaniu sie regul, Excel LACZY formaty tam,")
    print("  gdzie sie nie wykluczaja. Przy konflikcie tego samego atrybutu")
    print("  (np. dwa rozne tla) wygrywa regula o NIZSZYM numerze priorytetu.")
    print("  stopIfTrue=True na regule o wyzszym priorytecie UCINA dalsze.")
    print("  W naszym przykladzie warunki semaforu sa rozlaczne, wiec to bez")
    print("  znaczenia. Ale gdyby doszla regula 'audyt' (nakladajaca sie),")
    print("  wlasnie priorytet + stopIfTrue zdecydowalyby o kolorze Z-001.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)
    pokaz_nakladanie()


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **Trzy `add()` na tym samym zakresie = jedna grupa, trzy reguły.** Weryfikacja pokazuje `grup zakresow: 1` i `reguł w grupie: 3`. To niespodzianka 1 z 2.11 w działaniu. Gdybyś liczył `len(ws.conditional_formatting)`, dostałbyś `1` — a nie trzy. **Reguły licz przez zagnieżdżoną iterację.**

2. **Warunki są rozłączne — i to jest ważny wybór projektowy.** Użyłem `AND($C2>TODAY(), $C2<TODAY()+3)`, żeby „za mniej niż 3 dni" **nie obejmowało „dziś"**. Dzięki temu każda komórka trafia do **dokładnie jednej** reguły. To czyni semafor **odpornym**: nie zależy od priorytetów. Możesz dodać reguły w dowolnej kolejności i wynik będzie ten sam.

3. **A gdyby warunki się nakładały?** Wtedy wchodzi 2.6. W funkcji `pokaz_nakladanie` opisuję scenariusz z regułą „audyt", która nakłada się na „po terminie". Tam priorytet + `stopIfTrue` decydują, który kolor wygra. **Zrozum różnicę:** przy rozłącznych warunkach priorytet jest bez znaczenia; przy nakładających się — kluczowy.

**Po stronie technicznej jeszcze jedna rzecz:** daty muszą być **prawdziwymi datami** (moduł 05), nie tekstem. Gdyby `C` zawierało tekst `"2026-01-04"`, `$C2<TODAY()` porównywałoby tekst z liczbą i semafor działałby źle lub wcale. Dlatego ustawiłem `number_format` — i dlatego **waliduj typy przy zapisie** (moduł 05).

### Przykład 4 — inwentarz reguł i usuwanie (wzorzec rebuild)

Ten przykład rozwiązuje problem z 2.11: **jak odczytać i jak usunąć** reguły, skoro nie ma `clear()` ani usuwania pojedynczej reguły.

```python
"""Inwentarz regul + usuwanie przez przebudowe kolekcji (wzorzec rebuild).

Uruchom:  python examples/13_cf_inwentarz.py
    (wymaga pliku z Przykladu 3: output/13_cf_semafor.xlsx)
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList

ROOT = Path(__file__).resolve().parent.parent
ZRODLO = ROOT / "output" / "13_cf_semafor.xlsx"


def kolor_reguly(regula) -> str | None:
    """Probuje odczytac kolor tla ze stylu roznicowego reguly."""
    dxf = getattr(regula, "dxf", None)
    if dxf is None or dxf.fill is None:
        return None
    for atrybut in ("fgColor", "bgColor"):
        kolor = getattr(dxf.fill, atrybut, None)
        if kolor is not None and getattr(kolor, "rgb", None):
            return kolor.rgb
    return None


def inwentarz(sciezka: Path) -> list[dict]:
    """Zwraca plaska liste regul: zakres, typ, formula, priorytet, kolor."""
    wb = load_workbook(sciezka)
    wiersze: list[dict] = []
    try:
        for ws in wb.worksheets:
            for cf in ws.conditional_formatting:     # grupa zakresu
                for regula in cf.rules:              # reguly w grupie
                    wiersze.append({
                        "arkusz": ws.title,
                        "zakres": str(cf.sqref),
                        "typ": regula.type,
                        "operator": regula.operator,
                        "formula": "; ".join(regula.formula or []),
                        "priorytet": regula.priority,
                        "stopIfTrue": regula.stopIfTrue,
                        "kolor": kolor_reguly(regula),
                    })
    finally:
        wb.close()
    return wiersze


def pokaz_inwentarz(sciezka: Path) -> None:
    regul = inwentarz(sciezka)
    print("=" * 100)
    print(f"INWENTARZ REGUL: {sciezka.name}  (regul razem: {len(regul)})")
    print("=" * 100)
    naglowek = f"{'zakres':<10} {'typ':<12} {'prio':<5} {'stop':<5} {'kolor':<10} formula"
    print(naglowek)
    print("-" * 100)
    for r in sorted(regul, key=lambda x: (x["zakres"], x["priorytet"])):
        print(f"{r['zakres']:<10} {r['typ']:<12} {r['priorytet']:<5} "
              f"{str(r['stopIfTrue']):<5} {str(r['kolor']):<10} {r['formula']}")


def usun_reguly_typu(sciezka: Path, typ_do_usuniecia: str, cel: Path) -> None:
    """Usuwa reguly danego typu przez PRZEBUDOWE kolekcji (2.11, niespodzianka 3)."""
    wb = load_workbook(sciezka)
    try:
        for ws in wb.worksheets:
            stara = ws.conditional_formatting
            nowa = ConditionalFormattingList()       # pusta kolekcja
            for cf in stara:
                for regula in cf.rules:
                    if regula.type != typ_do_usuniecia:
                        # ZACHOWAJ regule - dodaj do nowej kolekcji pod tym samym zakresem
                        # UWAGA: regula ma juz priorytet, wiec add() go NIE nadpisze (2.5)
                        nowa.add(str(cf.sqref), regula)
            ws.conditional_formatting = nowa           # podmien kolekcje
        wb.save(cel)
    finally:
        wb.close()


def main() -> None:
    if not ZRODLO.exists():
        print(f"Brak pliku {ZRODLO}. Uruchom najpierw 13_cf_semafor.py.")
        return

    pokaz_inwentarz(ZRODLO)

    # Przyklad usuwania: usuniemy reguly typu "expression" (czyli FormulaRule)
    cel = ZRODLO.parent / "13_cf_bez_expression.xlsx"
    usun_reguly_typu(ZRODLO, "expression", cel)

    print()
    pokaz_inwentarz(cel)
    print()
    print("Zauważ: w pliku wynikowym nie ma juz zadnej reguly 'expression'.")
    print("Uzyliśmy przebudowy kolekcji, bo openpyxl nie ma remove() ani clear() (2.11).")


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **Iteracja jest dwupoziomowa.** `for cf in ws.conditional_formatting:` daje grupy zakresów, `for regula in cf.rules:` — reguły. To dokładnie struktura XML z 2.2.

2. **`add()` zachowuje istniejący priorytet.** Gdy przebudowuję kolekcję, wczytane reguły mają już swoje priorytety (1, 2, 3...). Wywołanie `nowa.add(...)` **nie nadpisze** ich, bo `add` ustawia priorytet tylko gdy `not rule.priority` (2.5). To zachowanie jest **pożądane** przy kopiowaniu — reguły zachowują kolejność. Ale musisz być tego świadomy, jeśli przebudowujesz i **chcesz** nową numerację.

3. **`del ws.conditional_formatting["A2:C7"]` usunąłby całą grupę** — wszystkie trzy reguły naraz. Gdy potrzebujesz usunąć **tylko jedną**, nie ma innego wyjścia niż rebuild. To nie jest wygodne, ale jest **deterministyczne**: dostajesz dokładnie to, co sam zbudowałeś.

**Uwaga o kopiowaniu reguł między arkuszami.** Powyższy wzorzec działa **w obrębie tego samego arkusza** (bo `dxfId` wskazuje na `styles.xml` tego samego skoroszytu). Kopiowanie reguł **między skoroszytami** jest trudniejsze — style różnicowe żyją w `styles.xml` każdego skoroszytu i trzeba by je przenosić osobno. To zapowiedź modułu 18 (modyfikacja plików) i typowa pułapka przy „kopiowaniu szablonu".

## 4. Anatomia API

| Metoda / klasa / atrybut | Co robi | Parametry / wartości | Uwagi |
|---|---|---|---|
| `ws.conditional_formatting` | Kolekcja reguł arkusza | `ConditionalFormattingList` | Kluczowana po `sqref`; w `read_only` nieprzydatna |
| `.add(range_string, rule)` | Dodaje regułę do zakresu | `range_string` — zakres w A1; `rule` — **musi być `Rule`** | Nadaje priorytet automatycznie; ten sam zakres → ta sama grupa |
| `len(ws.conditional_formatting)` | Liczba **grup zakresów** | — | **Nie liczba reguł!** (2.11) |
| `for cf in ws.conditional_formatting` | Iteracja po grupach | `cf` — `ConditionalFormatting` | `cf.sqref`, `cf.rules` |
| `ws.conditional_formatting["A2:A10"]` | Reguły zakresu | — | Zwraca listę reguł albo `KeyError` |
| `del ws.conditional_formatting["A2:A10"]` | Usuwa całą grupę zakresu | — | **Usuwa wszystkie reguły tego zakresu** |
| `ConditionalFormattingList()` | Nowa, pusta kolekcja | — | Z `openpyxl.formatting.formatting`. Służy do wyzerowania/rebuild |
| `ConditionalFormatting.sqref` | Zakres grupy | `MultiCellRange` | Może zawierać wiele zakresów (spacja) |
| `ConditionalFormatting.cells` | Alias `sqref` | — | Ta sama wartość |
| `ConditionalFormatting.rules` | Lista reguł grupy | `list[Rule]` | Alias `cfRule` |
| `CellIsRule(...)` | Reguła porównania ze stałą | `operator`, `formula` (lista!), `stopIfTrue`, `font`, `border`, `fill` | **Brak `priority` w konstruktorze** |
| `FormulaRule(...)` | Reguła wyrażenia | `formula` (lista!), `stopIfTrue`, `font`, `border`, `fill` | `type="expression"`; **formuła bez `=`** |
| `Rule(type, ...)` | Reguła dowolnego typu | `type` (**wymagany**), `operator`, `formula`, `text`, `dxf`, `priority`, `stopIfTrue`, ... | Jedyny, który przyjmuje `priority` |
| `Rule.formula` | Warunek reguły | lista napisów | Ustawialny po utworzeniu |
| `Rule.priority` | Priorytet | `int`, mniejszy = ważniejszy | Ustawiaj **po** utworzeniu dla `CellIsRule`/`FormulaRule` |
| `Rule.stopIfTrue` | Przerwij dalsze reguły | `True`/`False`/`None` | Ma znaczenie przy nakładaniu (2.6) |
| `Rule.type` | Rodzaj reguły | z listy w 2.3 | `cellIs`, `expression`, `colorScale`, ... |
| `Rule.dxfId` | Numer stylu różnicowego | `int` | Zarządzany przez openpyxl przy zapisie |
| `Rule.dxf` | Styl różnicowy | `DifferentialStyle` | Rozwiązywany przy odczycie |
| `DifferentialStyle(...)` | Różnica stylu | `font`, `fill`, `border`, `numFmt` | **Brak `alignment`** (2.4) |
| `CellRange.isdisjoint(other)` | Czy zakresy są rozłączne | — | Z `openpyxl.worksheet.cell_range` — do wykrywania nakładania |
| `CellRange.intersection(other)` | Część wspólna | — | Alias `&` |
| `MultiCellRange.sorted()` | Zakresy w kolejności | — | Użyteczne przy wyświetlaniu |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — podświetl wartości powyżej progu.**

Napisz `examples/13_cw1.py`, który tworzy `output/13_cw1.xlsx`:

1. Arkusz `Oceny` z nagłówkiem `Student`, `Punkty` oraz 12 wierszami danych (punkty losowe 0–100, `seed` ustalony).
2. **Regułę `CellIsRule`** podświetlającą na zielono (tło + pogrubienie) komórki `Punkty` **większe niż 80**, zakres dokładnie `B2:B13`.
3. **Drugą regułę** podświetlającą na czerwono komórki `B2:B13` **mniejsze niż 40**.

Następnie funkcją `weryfikuj()` wczytaj plik i wypisz:

- `len(ws.conditional_formatting)` (ile grup zakresów),
- dla każdej grupy: `sqref` i liczbę reguł w `cf.rules`,
- dla każdej reguły: `type`, `operator`, `formula`, `priority`, `stopIfTrue`.

W pliku `output/13_cw1_wnioski.md` odpowiedz:

1. **Ile grup zakresów zwrócił `len()` i dlaczego?** Ile jest reguł łącznie? Wyjaśnij różnicę (2.11).
2. **Co zwraca `ws['B2'].fill.fill_type` i dlaczego `None` mimo ustawionej reguły?** Odnieś się do analogii lampy z czujnikiem ruchu (2.1).
3. **Dlaczego obie reguły trafiły do jednej grupy zakresu, choć dodawałeś je osobno?** (2.11, niespodzianka 1).
4. **Jaki priorytet dostała pierwsza, a jaki druga reguła?** Dlaczego `add()` nadał je tak, a nie inaczej? (2.5)
5. **Czy tu potrzebny jest `stopIfTrue`?** Uzasadnij, odwołując się do tego, czy warunki mogą się nakładać (2.6).

**Zadanie 2 — rozpoznaj granicę między `CellIsRule` a `FormulaRule`.**

Nie pisząc kodu, odpowiedz w `output/13_cw2_wnioski.md` na poniższe pytania (możesz potem zweryfikować eksperymentalnie):

1. **Chcesz podświetlić komórki większe niż wartość **wpisana w komórce `D1`**.** Czy `CellIsRule(operator="greaterThan", formula=["D1"])` zadziała? Dlaczego tak/nie? Jaką regułę użyjesz? (2.8, 2.9)
2. **Chcesz podświetlić cały wiersz, gdy kwota w kolumnie C przekracza limit z komórki `G1`.** Napisz treść formuły (bez `=`), zakładając zakres `A2:E100`. Zwróć uwagę na `$`.
3. **Chcesz podświetlić komórki zawierające tekst „pilne" (bez względu na wielkość liter).** Czy `CellIsRule` to potrafi? Jak zbudujesz warunek?
4. **Chcesz podświetlić co drugi wiersz (pasy).** Jaką formułę napiszesz? Wskaż funkcję Excela, która daje numer wiersza.
5. **Dla każdego z punktów 1–4 powiedz, czy reguła jest „rozłączna" z innymi w Twoim raporcie, czy może się nakładać** — i czy z tego powodu potrzebujesz `stopIfTrue`.

### 🟡 Warsztat

**Zadanie 3 — podświetlanie całych wierszy z konfiguracji.**

To zadanie zamienia wiedzę z 2.7 i 2.9 w **narzędzie wielokrotnego użytku**. Zbudujesz funkcję, która z **deklaratywnej listy** warunków tworzy reguły — i nie da się w niej zapomnieć o `$`.

**Część A.** Napisz `examples/13_cw3.py` z klasą `RegulyWierszy`:

```python
"""Reguly formatowania warunkowego z deklaratywnej konfiguracji."""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill


@dataclass(frozen=True)
class Warunek:
    """Pojedynczy warunek podswietlenia wiersza."""
    kolumna: str                 # np. "D" - kolumna, na ktorej opiera sie warunek
    wartosc: str                 # np. "Pilne" - wartosc do porownania (jako napis)
    kolor: str                   # np. "FFC7CE" - kolor tla
    pogrubienie: bool = False    # czy dodac pogrubienie
    pierwszy_wiersz: int = 2     # pierwszy wiersz zakresu (danych)
    ostatni_wiersz: int = 2      # ostatni wiersz zakresu (danych)
    pierwsza_kolumna: str = "A"  # od ktorej kolumny podswietlac caly wiersz
    ostatnia_kolumna: str = "Z"  # do ktorej kolumny podswietlac caly wiersz


class RegulyWierszy:
    """Dodaje reguly podswietlania calych wierszy do arkusza."""

    def __init__(self, ws):
        self.ws = ws
        self._ostatni_priorytet = 0

    def dodaj(self, warunek: Warunek) -> None:
        """Buduje regule i dodaje ja do arkusza.

        Kluczowe: formuła uzywa $ przed kolumna i wzglednego wiersza
        zaczynajacego sie od pierwszego wiersza zakresu (2.7).
        """
        fill = PatternFill(start_color=warunek.kolor, end_color=warunek.kolor,
                           fill_type="solid")
        font = Font(bold=True, color="9C0006") if warunek.pogrubienie else None

        formula = f'${warunek.kolumna}{warunek.pierwszy_wiersz}="{warunek.wartosc}"'
        zakres = (f"{warunek.pierwsza_kolumna}{warunek.pierwszy_wiersz}:"
                  f"{warunek.ostatnia_kolumna}{warunek.ostatni_wiersz}")

        self.ws.conditional_formatting.add(
            zakres,
            FormulaRule(formula=[formula], fill=fill, font=font, stopIfTrue=True),
        )
```

**Wymagania:**

1. `dodaj(warunek)` buduje formułę **automatycznie** — z `$` przed kolumną i względnym wierszem. Nie da się zapomnieć o `$`.
2. Formuła **zaczyna się od `warunek.pierwszy_wiersz`** — nie od 1, nie od 2 „na sztywno". To rozwiązuje błąd nr 3 z 2.7.
3. Zakres `sqref` obejmuje **całe wiersze od `pierwsza_kolumna` do `ostatnia_kolumna`** i tylko wiersze danych.
4. Dodaj metodę `dodaj_wiele(warunki: list[Warunek])`, która pilnuje, że **kolejność dodawania = kolejność priorytetów**.

**Część B.** Użyj `RegulyWierszy`, żeby zbudować `output/13_cw3.xlsx` z fakturami, w których:

- wiersze z „Pilne" w kolumnie `Status` → czerwone,
- wiersze z „Zrealizowane" w kolumnie `Status` → zielone,
- wiersze z kwotą (kolumna `Kwota`) powyżej 10 000 → dodatkowo pogrubione.

**Część C.** Dodaj funkcję `weryfikuj(sciezka)`, która sprawdza **warunek niezmienny**: liczba reguł w arkuszu jest równa liczbie dodanych warunków (licz przez zagnieżdżoną iterację — nie przez `len()`!).

W pliku `output/13_cw3_wnioski.md`:

1. **Dlaczego klasa dodaje `$` przed nazwą kolumny automatycznie, a Ty jej to podajesz osobno (zamiast wpisać „$D2" ręcznie w formularzu)?** Jaki to ma związek z „jednym źródłem prawdy" (moduł 09)?
2. **Dlaczego formuła musi zaczynać się od `pierwszy_wiersz`, a nie od 2 „na sztywno"?** Opisz, co by się stało, gdyby tabela danych zaczynała się od wiersza 5 (2.7, błąd nr 3).
3. **Czy Twoja klasa pozwala stworzyć dwie reguły o nakładających się warunkach?** Jeśli tak — czy `stopIfTrue=True` na obu jest właściwe? Co byś zmienił, żeby obsłużyć nakładanie (2.6)?
4. **Jak policzyłbyś reguły, gdybyś nie wiedział, że istnieje zagnieżdżona `cf.rules`?** Opisz, jak Cię zaalarmowałaby wartość `len(ws.conditional_formatting)` (2.11).
5. **To zadanie realizuje pewien wzorzec.** Który wzorzec z modułu 25 (zapowiedź) opisuje budowanie obiektów (reguł) z deklaratywnej specyfikacji? Uzasadnij w dwóch zdaniach.

### 🔴 Wyzwanie

**Zadanie 4 — audytor reguł formatowania warunkowego.**

Napisz `examples/13_cw4.py` implementujący **audytor**, który znajduje problemy w formatowaniu warunkowym pliku. To rozszerzenie audytora tabel z modułu 12 na świat reguł.

```python
"""Audytor formatowania warunkowego."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.cell_range import CellRange


@dataclass
class Problem:
    waga: str          # "BLAD" albo "OSTRZEZENIE"
    arkusz: str
    zakres: str
    opis: str
    naprawa: str


def audytuj_formatowanie(sciezka: Path) -> list[Problem]:
    """Sprawdza reguly w pliku pod katem szesciu klas problemow:

      1. Reguła bez stylu roznicowego (dxf) -> nic nie robi.
      2. Reguła z formuła zaczynajaca sie od "=" -> Excel odrzuci.
      3. Reguła na ZBYT szerokim zakresie (calka kolumna, tysiace wierszy nadmiaru).
      4. Reguły o nakladajacych sie zakresach z tym samym stylem -> mozliwa duplikacja.
      5. Reguła z formuła odwolujaca sie do wiersza POZA zakresem sqref.
      6. Reguła z pustym wzorem (bez formula, bez wbudowanego warunku) -> nic nie oceni.
    """
    ...


def raport_tekstowy(problemy: list[Problem]) -> str:
    ...
```

**Wymagania funkcjonalne:**

1. **Sześć testów** — po jednym na klasę problemu. Dla każdego napisz osobny plik testowy (z celowo wstrzykniętym problemem) i sprawdź, że audyt go wykrywa.

2. **Problem 1 (brak `dxf`)** — zbuduj `Rule` bez `dxf` i bez stylu. Sprawdź, czy openpyxl go zapisze i czy audyt go złapie.

3. **Problem 2 (`=` w formule)** — spróbuj zapisać `FormulaRule(formula=['=A2>100'])`. Zapisz plik i **sprawdź, co trafiło do XML** (rozpakuj archiwum, obejrzyj `<formula>`). Zobacz, czy Excel to przyjmie. **To kluczowy test** — udowadnia, dlaczego `=` jest zabroniony.

4. **Problem 3 (za szeroki zakres)** — reguła na `"B:B"` albo `"B1:B1048576"`. Audytor ma porównać `sqref` z faktycznym zakresem danych (`ws.max_row`, `ws.max_column`, moduł 07) i ostrzec, jeśli reguła obejmuje **znacznie więcej** niż dane (np. więcej niż 2×).

5. **Problem 4 (nakładające się zakresy)** — dwie reguły na nachodzących zakresach z identycznym stylem. Użyj `CellRange.isdisjoint(other)` (4. sekcja API). Audytor ma ostrzec o możliwej duplikacji.

6. **Problem 5 (odwołanie poza zakresem)** — zakres `A2:A10` i formuła odnosząca się do `$Z2` (kolumna poza zakresem) albo do wiersza 100 (poza zakresem). Audytor ma ostrzec. **Uwaga:** sprawdzenie wymaga sparsowania formuły — użyj prostego wyrażenia regularnego do wyciągnięcia odwołań typu `$X9`, a następnie `CellRange` do sprawdzenia, czy kolumna/wiersz mieszczą się w `sqref`.

7. **Problem 6 (pusty wzór)** — `Rule(type="expression")` bez `formula`. Audytor ma sprawdzić, czy reguła ma co najmniej jeden warunek (`formula`, `text`, albo wbudowany `colorScale`/`dataBar`/`iconSet`).

8. **`raport_tekstowy`** — czytelny raport:

```
=== AUDYT FORMATOWANIA: 13_cw4.xlsx ===

Arkusz: Wyniki
  [BLAD] B2:B11  (typ: cellIs)
     opis:    formuła zaczyna się od "=" -> Excel odrzuci regułę
     naprawa: usuń znak "=" z początku formuły (element formula w XML)

  [OSTRZEZENIE] B1:B1048576  (typ: expression)
     opis:    reguła obejmuje całą kolumnę, a dane kończą się na wierszu 13
     naprawa: ogranicz sqref do realnego zakresu danych

PODSUMOWANIE: 1 błąd, 1 ostrzeżenie
```

9. **Zapisz raport** do `output/13_cw4_raport.txt` i wypisz `wynik: X/6`.

W pliku `output/13_cw4_wnioski.md`:

1. **Co dokładnie znalazło się w XML dla reguły z formułą zaczynającą się od `=`?** Obejrzyj `<formula>` w `sheetN.xml` i opisz różnicę wobec poprawnej wersji. Czy Excel to zaakceptował? (To pytanie o 2.9.)
2. **Jak wykryłeś reguły o nakładających się zakresach?** Opisz użycie `CellRange.isdisjoint`. Dlaczego to ten sam test co w module 12 przy tabelach — i dlaczego to nie przypadek?
3. **Jak wykryłeś odwołania poza zakresem `sqref`?** Czy Twoja metoda zawodzi przy skomplikowanych formułach (`AND`, `SEARCH`, odwołania do innych arkuszy)? Uczciwie opisz ograniczenia.
4. **Dlaczego „zbyt szeroki zakres" to OSTRZEŻENIE, a nie BŁĄD?** Uzasadnij w kontekście różnicy między poprawnością pliku a jego wydajnością (2.12).
5. **Gdyby openpyxl miał silnik oceny formuł, ten audytor mógłby być dokładniejszy. Czego konkretnie nie możesz sprawdzić bez takiego silnika?** (Odsyłam do modułu 06 — brak silnika obliczeń.)

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Audytor formatowania warunkowego - kluczowe fragmenty."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.cell_range import CellRange


# Wyciaga odwolania typu $C2, C$2, $C$2, C2 z formuly
WZORZEC_ODWOLANIA = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")

# Ile razy wiekszy moze byc sqref od danych, zeby nie ostrzegac
PROG_NADMiaru = 2


def _raport_z_odwolan(formula: str) -> list[tuple[str, int]]:
    """Zwraca liste (litera_kolumny, numer_wiersza) z formuly."""
    return [(k, int(w)) for k, w in WZORZEC_ODWOLANIA.findall(formula)]


def audytuj_formatowanie(sciezka: Path) -> list[Problem]:
    wb = load_workbook(sciezka)
    problemy: list[Problem] = []
    try:
        for ws in wb.worksheets:
            for cf in ws.conditional_formatting:
                for regula in cf.rules:
                    zakres = str(cf.sqref)

                    # --- 1. Brak dxf --------------------------------------
                    if getattr(regula, "dxf", None) is None \
                            and regula.colorScale is None \
                            and regula.dataBar is None \
                            and regula.iconSet is None:
                        problemy.append(Problem(
                            "BLAD", ws.title, zakres,
                            f"reguła typu {regula.type} nie ma stylu różnicowego (dxf)",
                            "dodaj font/fill/border albo użyj wbudowanej reguły",
                        ))

                    # --- 2 & 6. Formula: "=" i pusty wzor -----------------
                    formuly = list(regula.formula or [])
                    if regula.type == "expression" and not formuly:
                        problemy.append(Problem(
                            "BLAD", ws.title, zakres,
                            "reguła expression nie ma formuły",
                            "ustaw regula.formula = ['A2>100']",
                        ))
                    for formuła in formuly:
                        if formuła.startswith("="):
                            problemy.append(Problem(
                                "BLAD", ws.title, zakres,
                                'formuła zaczyna się od "=" -> Excel odrzuci regułę',
                                'usuń znak "=" z początku formuły',
                            ))
                        if "!" in formuła:
                            # odwolanie do innego arkusza - audyt nie analizuje
                            pass

                    # --- 3. Za szeroki zakres ----------------------------
                    min_col, min_row, max_col, max_row = range_boundaries(
                        zakres.split()[0]        # pierwszy zakres grupy
                    )
                    dane_wierszy = ws.max_row
                    dane_kolumn = ws.max_column
                    if max_row > dane_wierszy * PROG_NADmiaru + 1 \
                            or max_col > dane_kolumn * PROG_NADmiaru + 1:
                        problemy.append(Problem(
                            "OSTRZEZENIE", ws.title, zakres,
                            f"reguła obejmuje dużo więcej niż dane "
                            f"(dane: do wiersza {dane_wierszy}, kolumny {dane_kolumn})",
                            "ogranicz sqref do realnego zakresu danych",
                        ))

                    # --- 5. Odwolania poza zakresem sqref ----------------
                    # (uproszczenie: bierzemy pierwszy zakres grupy)
                    first = CellRange(zakres.split()[0])
                    for formuła in formuly:
                        for litera, wiersz in _raport_z_odwolan(formuła):
                            try:
                                cr = CellRange(f"{litera}{wiersz}")
                            except ValueError:
                                continue
                            # Sprawdz, czy kolumna/wiersz wspolrzednej jest w sqref
                            kol_w_zakresie = first.min_col <= cr.min_col <= first.max_col
                            wiersz_w_zakresie = first.min_row <= cr.min_row <= first.max_row
                            # Wiersz MUSI byc wzgledny -> porownujemy tylko zakresowo
                            if not kol_w_zakresie:
                                problemy.append(Problem(
                                    "OSTRZEZENIE", ws.title, zakres,
                                    f"formuła odwołuje się do kolumny {litera} "
                                    f"poza zakresem {first.coord}",
                                    "sprawdź, czy odwołanie jest zamierzone (2.7)",
                                ))
    finally:
        wb.close()
    return problemy
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **Problem 2 jest sprawdzalny bez Excela i to jest jego wartość.** Reguła z `=` w formule trafia do XML dosłownie, jako `<formula>=A2&gt;100</formula>`. Gdy rozpakujesz archiwum (moduł 02), zobaczysz to gołym okiem. **To dowód na tezę z 2.9:** openpyxl zapisuje napis dosłownie i nie waliduje. Excel przy otwarciu zobaczy `=` w treści `<formula>` — czyli znak, który w tym kontekście jest nadmiarowy — i potraktuje warunek jako nieparsowalny. Dlatego **audytor może to wyłapać, bo problem jest w strukturze, nie w semantyce.**

- **Problem 3 używa porównania z realnym rozmiarem danych.** To ta sama zasada co `print_area` w module 11 i `ref` tabeli w module 12: **granice zakresu wyliczamy z danych, nie z założenia.** Ale uwaga na uczciwość: `ws.max_row` (moduł 07) **bywa zawyżony przez formatowanie**. Więc próg „2× więcej" daje pewien margines — dlatego to OSTRZEŻENIE, nie BŁĄD.

- **Problem 4 (nakładanie) używa `CellRange.isdisjoint`.** To **ten sam test** co w module 12 przy nakładaniu tabel — nie przypadek, a **reuse abstrakcji**: „dwa prostokąty zachodzą na siebie albo nie" to czysta geometria, niezależna od tego, czy prostokątami są tabele, scalone komórki, czy zakresy reguł. Gdy jedna funkcja obsługuje wiele przypadków, jest to sygnał, że abstrakcja jest właściwa (zapowiedź modułu 25 — Specification).

- **Problem 5 jest oznaczony jako OSTRZEŻENIE i to jest świadome.** Audytor **nie ocenia formuł** (bo openpyxl nie ma silnika, moduł 06). Potrafi tylko wyciągnąć odwołania wzorcem i sprawdzić, czy **numer kolumny** mieści się w zakresie `sqref`. Nie wie, czy odwołanie do wiersza jest zamierzone (bo wiersz ma być wzgledny!). Więc formuła `$Z2` przy zakresie `A2:A10` → OSTRZEŻENIE (kolumna Z poza A), ale wiersz `100` przy zakresie `A2:A10` **nie jest błędem**, bo formuła jest wzgledna i przesunie się przez cały zakres. **Być może najważniejsza lekcja tego zadania: audytor musi wiedzieć, co jest wzgledne, a co absolutne w kontekście reguł.**

- **Problem 6 sprawdza, czy reguła ma „coś do zrobienia".** Reguła bez `dxf` i bez wbudowanego warunku nic nie zrobi — jest „pusta". To ta sama rodzina problemów co `PatternFill()` bez koloru z 2.4: obiekt istnieje, ale nie ma efektu.

- **Czego audytor NIE sprawdza i dlaczego to uczciwe.** Bez silnika oceny formuł openpyxl **nie odpowie na pytanie „czy ta reguła kiedykolwiek zadziała"**. Audytor może wyłapać problemy **strukturalne** (brak dxf, `=` w formule, za szeroki zakres, nakładanie), ale nie **semantyczne** („czy warunek jest sensowny dla tych danych"). To dokładnie ta sama granica co w module 06: **openpyxl jest skrybą, nie matematykiem.** Uczciwy audytor mówi, że sprawdza strukturę — i nie udaje, że ocenia logikę.

**Dlaczego to wykrywamy u siebie, a nie polegamy na Excelu:** bo gdy Excel otworzy plik z regułą zawierającą `=`, użytkownik zobaczy albo dziwne zachowanie, albo brak podświetlenia — a Ty nie. Audytor **przed wysłaniem pliku** to różnica między „działa" a „klient dzwoni". To ten sam wzorzec co audytor tabel z modułu 12.

</details>

## 6. Typowe błędy i pułapki

**1. „Formuła w `FormulaRule` zaczyna się od `=` i reguła nie działa / Excel odrzuca plik" (objaw) → **w `formula` nie wolno umieszczać znaku `=`.** Ten napis trafia dosłownie do elementu `<formula>` w XML, a Excel oczekuje tam **samego wyrażenia** (2.9). Odruch programisty z Excela każe pisać `=`, ale tu to błąd (przyczyna) → **usuń `=`**: `FormulaRule(formula=['A2>100'])`, nie `['=A2>100']`. To skrajnie częsty błąd, bo `=` jest naturalny. Zapamiętaj regułę: **`=` pojawia się w treści komórki (`ws['A1'] = "=SUM(...)"`), ale nie w warunku reguły** (naprawa).**

**2. „Odczytałem plik i chcę sprawdzić, które komórki są czerwone, ale `cell.fill` jest puste" (objaw) → **koloru z formatowania warunkowego nie ma w komórce — żyje w regule.** To nie błąd, to natura mechanizmu (2.1). Komórka ma tylko swój zwykły styl (przyczyna) → **czytaj kolor z reguł**: iteruj `ws.conditional_formatting`, potem `cf.rules`, i sprawdzaj `rule.dxf.fill` (Przykład 4). A jeśli naprawdę musisz wiedzieć, czy dana komórka **jest** podświetlona, musisz **ocenić warunek** — czego openpyxl nie robi (moduł 06). Wtedy albo licz w Pythonie, albo otwórz plik w Excelu (naprawa).**

**3. „Reguła podświetla cały zakres / podświetla złe wiersze / nie podświetla nic" (objaw) → **złe odwołania w formule — kalka z 1.3.** Trzy typowe warianty: brak `$` (kolumna ucieka w prawo), zablokowany wiersz `$D$2` (zawsze jeden wiersz), formuła zaczynająca się od wiersza poza zakresem (2.7) (przyczyna) → **pisz formułę dla lewego-górnego rogu zakresu**: kolumna zablokowana (`$D`), numer wiersza wzgledny i **od pierwszego wiersza zakresu**. Potem zweryfikuj w Excelu. Najczęstszy objaw „prawie działa" to brak `$` — coś się podświetla, ale nie to, co trzeba (naprawa).**

**4. „Chcę ustawić `CellIsRule(..., priority=1)` i dostaję `TypeError`" (objaw) → **konstruktory `CellIsRule` i `FormulaRule` nie mają parametru `priority`.** Ma go tylko bazowy `Rule` (2.5) (przyczyna) → **ustaw priorytet po utworzeniu**: `rule = CellIsRule(...); rule.priority = 1`. Albo — co zwykle jest lepsze — **dodaj reguły w takiej kolejności, w jakiej mają być rozpatrywane**, a priorytety nadadzą się same (naprawa).**

**5. „`len(ws.conditional_formatting)` zwraca 1, a dodałem trzy reguły" (objaw) → **`len()` liczy grupy ZAKRESÓW, nie reguły.** Trzy reguły na tym samym zakresie siedzą w jednej grupie pod jednym `sqref` (2.11, niespodzianka 1) (przyczyna) → **licz reguły przez zagnieżdżoną iterację**: `sum(len(cf.rules) for cf in ws.conditional_formatting)`. To ta sama rodzina niespodzianek co `ws.tables.items()` zwracające napisy w module 12 (naprawa).**

**6. „Chcę usunąć jedną regułę i nie znajduję metody `remove` ani `clear`" (objaw) → **`ConditionalFormattingList` ma tylko `add`.** Nie ma usuwania pojedynczej reguły ani czyszczenia (2.11, niespodzianka 3) (przyczyna) → (a) cała grupa zakresu: `del ws.conditional_formatting["A2:A10"]`, (b) wszystkie reguły arkusza: `ws.conditional_formatting = ConditionalFormattingList()`, (c) **pojedyncza reguła: przebuduj kolekcję**, dodając z powrotem tylko to, co chcesz zachować (Przykład 4). Uwaga: przy przebudowie `add` **zachowa** stare priorytety (naprawa).**

**7. „Regułą podświetlam całą kolumnę (`B:B`) i plik jest wolny" (objaw) → **reguła obejmuje 1 048 576 komórek, choć dane kończą się na 100.** Excel ocenia regułę dla każdej komórki zakresu (2.12) (przyczyna) → **ogranicz `sqref` do realnego zakresu danych**: `"B2:B100"`. To ta sama zasada co `print_area` (moduł 11) i `ref` tabeli (moduł 12): **granice zakresu wyliczamy z danych, nie z założenia.** Sprawdzaj `ws.max_row` (naprawa).**

**8. „Tło reguły się nie pojawia w Excelu, choć kod nie zgłasza błędu" (objaw) → dwie powiązane przyczyny. Po pierwsze: **`PatternFill` bez `fill_type="solid"` i bez koloru nie maluje nic** (2.4). Po drugie: w stylach różnicowych **Excel bywa kapryśny co do `fgColor`/`bgColor`** — potrafi pokazywać kolor z drugiego pola wzorca (przyczyna) → **podawaj `PatternFill(start_color=..., end_color=..., fill_type="solid")` z jednakowym kolorem w obu**, a jeśli tło nadal się nie pojawia — **zamień kolory miejscami i sprawdź w Excelu.** To ta sama rodzina co `fgColor`/`bgColor` z modułu 09, tylko w kontekście dxf (naprawa).**

**9. „Podświetlanie nie działa, gdy zmieniam daty — nic się nie podświetla" (objaw) → **daty w komórkach są tekstem, nie prawdziwymi datami.** Formuła `$C2<TODAY()` porównuje tekst z liczbą i daje fałsz (albo błąd) (2.9, Przykład 3). To ta sama pułapka co `"1"` vs `1` z modułu 05 (przyczyna) → **upewnij się, że daty są `datetime`/`date`** (nie napisami) i ustaw `number_format`. Sprawdź `cell.data_type` (moduł 10) — dla daty powinno być `'d'` (naprawa).**

**10. „Wczytaj-zapisz pliku z Excela i część formatowania warunkowego zniknęła" (objaw) → **reguły zapisane w formacie rozszerzeń (`extLst`) nie są modelowane przez openpyxl i giną przy zapisie** (2.13). Dotyczy to m.in. niektórych pasków danych, nowych zestawów ikon i nietypowych kombinacji (przyczyna) → (a) **nie przepuszczaj bogatego pliku z Excela przez wczytaj-zapisz** — generuj nowy, albo odtwórz formatowanie od zera, (b) porównaj inwentarz reguł **przed i po** zapisie (Przykład 4) i wykryj ubytek. To trzecia kategoria z modułu 03 — „gubi dane przy zapisie" (naprawa).**

**11. „`stopIfTrue=True` nie zmieniło niczego" (objaw) → **warunki Twoich reguł są rozłączne, więc priorytet i `stopIfTrue` nie mają na co wpłynąć.** Jeśli komórka może pasować tylko do jednej reguły, „przerwij dalsze sprawdzanie" jest bez znaczenia (2.6) (przyczyna) → **to nie jest błąd.** `stopIfTrue` ma sens tylko wtedy, gdy reguły **mogą się nakładać**. Zamiast go dorzucać „na wszelki wypadek", **zrozum, czy jest potrzebny.** Alternatywnie, jak w Przykładzie 3 — zaprojektuj reguły tak, by były rozłączne, i wtedy nie potrzebujesz żadnego `stopIfTrue` (naprawa).**

**12. „Skopiowałem arkusz z regułami do innego skoroszytu i reguły nie mają koloru" (objaw) → **style różnicowe żyją w `styles.xml` konkretnego skoroszytu, a reguła odwołuje się do nich numerem (`dxfId`).** Kopiowanie reguł między skoroszytami bez przeniesienia stylów różnicowych daje reguły „z oderwanym" stylem (2.13, Przykład 4) (przyczyna) → **kopiuj reguły w obrębie tego samego skoroszytu** (albo odtwórz je od zera w nowym). To ta sama rodzina co ograniczenia `copy_worksheet` z modułu 08 — **nie wszystko, co wygląda jak kopia, jest kopią.** Najbezpieczniej: buduj reguły w nowym pliku programowo (naprawa).**

**13. „`ws.conditional_formatting` jest puste, mimo że plik ma formatowanie warunkowe" (objaw) → **w trybie `read_only` (moduł 19) dostęp do obiektów arkusza jest ograniczony** — to tryb strumieniowego czytania **danych**, nie modelu arkusza. Reguły nie są wtedy w użytecznej formie (przyczyna) → **jeśli musisz czytać reguły, użyj trybu normalnego** (`load_workbook` bez `read_only`). Jeśli plik jest ogromny i musisz użyć `read_only` do danych — **czytaj reguły osobno, w drugim przebiegu** (wczytaj tylko strukturę). To ta sama cena trybów wydajnościowych co `ws.tables` w module 12: **`read_only` kupuje pamięć za cenę funkcji** (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Reguła, nie kolor.** Formatowanie warunkowe to **zestaw reguł przypisanych do zakresu**, a nie efekt zapisany w komórkach. Reguły mówią: „jeśli warunek, nałóż różnicę stylu". Konsekwencja twarda: **po wczytaniu pliku `cell.fill` jest puste** — kolor pojawia się dopiero, gdy Excel otworzy plik i oceni reguły. To „lampa z czujnikiem ruchu": urządzenie jest w korytarzu, farba — nie.

2. **`DifferentialStyle` to różnica, nie pełny styl.** Reguła niesie tylko to, co **doda** do istniejącego wyglądu komórki — jak poprawka długopisem na wydruku. `DifferentialStyle` ma cztery składniki (`font`, `fill`, `border`, `numFmt`) i **nie ma `alignment`**. W pliku style różnicowe żyją w `xl/styles.xml` w sekcji `<dxfs>`, a reguła odwołuje się do nich **numerem** (`dxfId`) — dokładnie jak komórka odwołuje się do stylu numerem (moduł 02).

3. **Formuła widzi tylko lewy-górny róg zakresu — to kalka.** Formułę pisz tak, jakby istniał tylko pierwszy element zakresu; Excel przesuwa ją na resztę (analogia 1.3). Dlatego: **kolumny stałe → `$` (`$D`), numer wiersza → wzgledny i od pierwszego wiersza `sqref`**. Trzy błędy z tego wynikające to brak `$` (kolumna ucieka w prawo), zablokowany wiersz (`$D$2`) i start poza zakresem. **Formuła nie zaczyna się od `=`** — to najczęstszy pojedynczy błąd.

4. **Priorytet jest globalny dla arkusza; mniejsza liczba znaczy ważniejsza.** `add()` nadaje priorytet automatycznie (kolejny wolny numer), a konstruktory `CellIsRule`/`FormulaRule` **nie mają** parametru `priority` — ustawia się go po utworzeniu. `stopIfTrue` ucina dalsze sprawdzanie — ale **ma znaczenie tylko przy nakładających się warunkach**; przy rozłącznych jest placebo. Kolekcja nie ma `clear()` ani usuwania pojedynczej reguły: usuwasz całą grupę przez `del ws.conditional_formatting["zakres"]` albo **przebudowujesz** kolekcję.

5. **openpyxl modeluje klasyczne reguły, ale gubi rozszerzenia.** Reguły `cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet` tworzone i wczytywane przechodzą. Ale reguły zapisane przez nowsze wersje Excela w formacie rozszerzeń (`extLst`) są **niewidoczne i giną przy zapisie** (2.13). To trzecia kategoria z modułu 03. Wniosek: **bogatego pliku z Excela nie przepuszczaj bezmyślnie przez wczytaj-zapisz** — generuj nowy albo odtwórz formatowanie, i zawsze porównaj inwentarz reguł przed i po.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORT
# ==================================================================
from openpyxl.formatting.rule import CellIsRule, FormulaRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.cell_range import CellRange

# ==================================================================
# 2. NAJPROSTSZA REGULA (CellIsRule - porownanie ze STALA)
# ==================================================================
czerwone = PatternFill(start_color="EE1111", end_color="EE1111", fill_type="solid")
ws.conditional_formatting.add(
    "A2:A100",
    CellIsRule(operator="greaterThan", formula=["100"], fill=czerwone),
)
# ⚠️ formula to LISTA napisow: ["100"] albo ["1","5"] dla between
# ⚠️ operator: nazwa ("greaterThan") albo symbol (">")

# ==================================================================
# 3. REGULA WYRAZENIOWA (FormulaRule - dowolny warunek)
# ==================================================================
# Podswietl CALY WIERSZ, gdy kolumna D = "Pilne"  (zakres A2:E100)
ws.conditional_formatting.add(
    "A2:E100",
    FormulaRule(formula=['$D2="Pilne"'], fill=czerwone, stopIfTrue=True),
)
# ⚠️ formula BEZ znaku "=" na poczatku!
# ⚠️ $D (kolumna zablokowana) + 2 (PIERWSZY wiersz zakresu, wzgledny)
# ⚠️ bledne warianty: ['D2="Pilne"'] (bez $), ['$D$2="Pilne"'] (zablokowany wiersz)

# ==================================================================
# 4. KATALOG FORMUL (bez "=")
# ==================================================================
# 'A2>$C$1'                          wartosc wieksza niz w komorce C1
# '$D2="Pilne"'                      caly wiersz na podstawie kolumny D
# '$D2>$E2'                          porownanie dwoch kolumn
# '$C2<TODAY()'                      data minela
# '$C2=TODAY()'                      data dzis
# 'AND($C2>TODAY(),$C2<TODAY()+3)'   data w najbliszych 3 dniach
# 'ISBLANK($A2)'                     komorka pusta
# 'ISNUMBER(SEARCH("pilne",$B2))'    zawiera tekst (bez wielkosci liter)
# 'AND($C2="Pilne",$D2>10000)'       dwa warunki
# 'MOD(ROW(),2)=0'                   co drugi wiersz (pasy)

# ==================================================================
# 5. RECZNY Rule + DifferentialStyle (nietypowy typ / reuse dxf)
# ==================================================================
dxf = DifferentialStyle(
    font=Font(bold=True, color="9C0006"),
    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
)
rule = Rule(type="containsText", operator="containsText", text="pilne", dxf=dxf)
rule.formula = ['NOT(ISERROR(SEARCH("pilne",A2)))']
ws.conditional_formatting.add("A2:A100", rule)

# ==================================================================
# 6. PRIORYTETY I stopIfTrue (2.5, 2.6)
# ==================================================================
# add() nadaje priorytet automatycznie (1, 2, 3...) w kolejnosci dodawania
# mniejsza liczba = wazniejsza; priorytet jest GLOBALNY dla arkusza
# ⚠️ CellIsRule/FormulaRule NIE maja parametru priority - ustaw PO utworzeniu:
r = CellIsRule(operator="greaterThan", formula=["100"], fill=czerwone)
r.priority = 1                                   # ustawione recznie
ws.conditional_formatting.add("A2:A100", r)
# stopIfTrue=True ucina dalsze reguly - MA ZNACZENIE tylko przy nakladaniu (2.6)

# ==================================================================
# 7. KOLEKCJA REGUL - trzy niespodzianki (2.11)
# ==================================================================
len(ws.conditional_formatting)              # ⚠️ liczba GRUP ZAKRESOW, nie regul!
for cf in ws.conditional_formatting:        # grupa zakresu
    print(cf.sqref)                         # MultiCellRange, np. "A2:A100"
    for regula in cf.rules:                 # DOPIERO tu reguly
        print(regula.type, regula.priority, regula.formula)

# policz REGULY (nie grupy):
sum(len(cf.rules) for cf in ws.conditional_formatting)

# dostep do regul konkretnego zakresu:
lista = ws.conditional_formatting["A2:A100"]     # albo KeyError

# ==================================================================
# 8. USUWANIE (brak remove() i clear()!) (2.11)
# ==================================================================
del ws.conditional_formatting["A2:A100"]          # cala grupa zakresu

ws.conditional_formatting = ConditionalFormattingList()   # wszystkie reguly arkusza

# pojedyncza regula - PRZEBUDOWA kolekcji:
stara = ws.conditional_formatting
nowa = ConditionalFormattingList()
for cf in stara:
    for regula in cf.rules:
        if regula.type != "dataBar":              # warunek "zachowaj"
            nowa.add(str(cf.sqref), regula)       # zachowa stary priorytet
ws.conditional_formatting = nowa

# ==================================================================
# 9. ODCZYT KOLORU Z REGULY (bo z komorki sie nie da!) (2.1)
# ==================================================================
dxf = regula.dxf
if dxf is not None and dxf.fill is not None:
    kolor = dxf.fill.fgColor.rgb     # albo bgColor - sprawdz w swojej wersji

# ==================================================================
# 10. SPRAWDZENIE NAKLADANIA ZAKRESOW (jak w module 12)
# ==================================================================
a = CellRange("A2:A100")
b = CellRange("A50:A200")
print(a.isdisjoint(b))               # False -> zakresy sie przecinaja

# ==================================================================
# 11. SZYBKA DIAGNOSTYKA (co sprawdzic w Excelu)
# ==================================================================
# Narzedzia glowne -> Formatowanie warunkowe -> Zarzadzaj regulami:
#   - lista regul, kolejnosc priorytetow, stopIfTrue
# Klik w podswietlona komorke -> Zarzadzaj regulami: ktora regula dziala?
# Nurek: plik .xlsx jako ZIP -> xl/worksheets/sheetN.xml (<conditionalFormatting>)
#                              -> xl/styles.xml (<dxfs>)
# ⚠️ cell.fill NIE pokaze koloru reguly - tylko styl komorki
```

## 9. Co dalej

Zamknąłeś pierwszy moduł o formatowaniu warunkowym. Umięsz już najważniejszą rzecz: **zrozumieć, że reguła to urządzenie w arkuszu, a nie efekt w komórce** — oraz bezpiecznie tworzyć reguły `cellIs` i `expression`. Trzy rzeczy, które zabierasz:

- **Model „kalki".** Formuła widzi lewy-górny róg zakresu i jest przesuwana na resztę. To pojęcie wróci, ilekroć warunek zależy od innej kolumny albo od daty — a więc niemal zawsze w raportach.
- **Świadomość różnicy między „efektem" a „regułą".** Koloru z reguły nie odczytasz z komórki. Ta sama granica co „tabela to warstwa opisu, nie danych" z modułu 12 — i ten sam powód, dla którego openpyxl nie potrafi „policzyć", które komórki są podświetlone.
- **Nawyk sprawdzania w Excelu.** Formatowanie warunkowe to obszar, w którym **wygląd prawdy jest w Excelu**, nie w kodzie. Reguła może się zapisać bez błędu i nie działać. Weryfikacja wzrokowa nie jest lenistwem — jest częścią procesu.

**W module 14** wejdziemy w **formatowanie warunkowe zaawansowane** — wizualizacje i pełną kontrolę:

- **`ColorScaleRule`** — skale 2- i 3-kolorowe, typy progów (`min`, `max`, `percent`, `percentile`, `num`, `formula`) i dlaczego kolejność kolorów bywa odwrócona;
- **`DataBarRule`** — paski danych: „wykres zrobiony z wypełnienia komórki", progi, wartości ujemne, `minLength`/`maxLength`;
- **`IconSetRule`** — zestawy ikon (`3Arrows`, `3TrafficLights1`, `4Rating`, `5Quarters`), `reverse`, `showValue`;
- **`Top10Rule`, `AboveAverageRule`, `DuplicateValuesRule`, `TextRule`, `DateRule`** — pozostałe warianty standardowe i ich parametry;
- **kolejność i nakładanie się reguł** — praktyczny scenariusz: skala kolorów na całej kolumnie + wyróżnienie wartości skrajnych, i dlaczego reguła wyróżnienia musi mieć **niższy** numer priorytetu;
- **skalowanie** — liczba reguł × zakres komórek jako koszt, i wzorce, które go ograniczają;
- zapowiedź wzorca **Specification** (moduł 25): reguły definiowane deklaratywnie (słownik/YAML), a potem tłumaczone na `Rule`.

**Zanim przejdziesz dalej, wykonaj jedno ćwiczenie obserwacyjne.** Otwórz `output/13_cf_semafor.xlsx` w Excelu i wejdź w **Formatowanie warunkowe → Zarządzaj regułami**. Zobaczysz trzy reguły w jednej grupie zakresu, w kolejności priorytetów 1, 2, 3.

1. **Kliknij „Edytuj regułę"** przy pierwszej i popatrz na formułę: `$C2<TODAY()` — **bez znaku `=`**. To jest ten szczegół, który za tydzień uratuje Ci godzinę debugowania.
2. **Zmień w komórce `C5` datę** (po terminie) na datę dzisiejszą i zapisz. Zamknij i otwórz plik. **Kolor komórki zmienił się sam** — bo to reguła, nie kolor zapisany w komórce. To doświadczenie analogii 1.1 w praktyce.
3. **Dodaj ręcznie w Excelu drugą regułę** modyfikującą ten sam zakres (np. „podświetl na niebiesko, gdy zadanie zawiera «Audyt»"). Zapisz plik, **wczytaj go openpyxl** i uruchom inwentarz reguł z Przykładu 4. Zobaczysz, że openpyxl odczytał Twoją ręczną regułę — ale **sprawdź, czy odczytał jej `dxf`**, bo tu można się zdziwić. To bezpośrednie przygotowanie do modułu 18: plik, który przeszedł przez ręce człowieka, ma inną strukturę niż ten, który wygenerowałeś.