Zanim napiszę moduł, zweryfikuję sygnatury klas reguł zaawansowanych w źródle openpyxl 3.1.x — to obszar, w którym brief zawiera co najmniej jedną niepewną pozycję (`DifferentialStyle(name=...)`), a konstruktory mają zaskakujące parametry.

Dokumentacja potwierdzona. Teraz pobiorę samo źródło klas reguł i `DifferentialStyle` przez CDN (szybszy i pewniejszy od surowego GitHuba).

Fetch surowych plików zawodzi. Spróbuję inaczej: wyszukiwanie źródeł sygnatur oraz widok GitHub.

Mam sygnatury. Teraz pobiorę pełne źródło (aby ustalić, które klasy faktycznie istnieją — brief wymienia `Top10Rule`, `DateRule` itd., czego muszę zweryfikować) oraz `DifferentialStyle`.

Mam pełną prawdę o API — i brief zawiera kilka nieścisłości, które muszę sprostować (m.in. brak klas `Top10Rule`/`DateRule`, brak `name` w `DifferentialStyle`). Zweryfikuję jeszcze jedną rzecz, od której zależy odczyt reguł: czy po wczytaniu pliku `regula.dxf` jest rozwiązywane.

Spróbuję potwierdzić zachowanie `dxf` przy odczycie — to wpływa na przykład „inwentarza reguł”.

Weryfikacja zakończona. Kluczowe ustalenia ze źródła `openpyxl` 3.1.x, które **prostują brief** (i które wykorzystam w module):

- **Nie istnieją** klasy `Top10Rule`, `AboveAverageRule`, `BelowAverageRule`, `DuplicateValuesRule`, `TextRule`, `DateRule`. openpyxl 3.1.x ma tylko **pięć** fabryk: `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule`, `IconSetRule`. Wszystko inne buduje się bazowym `Rule`.
- `DifferentialStyle` **nie ma** parametru `name` — ma `font, numFmt, fill, alignment, border, protection`.
- `ColorScaleRule` / `DataBarRule` / `IconSetRule` **nie przyjmują** `dxf` ani `stopIfTrue` (wbudowane).
- Po wczytaniu pliku `regula.dxf` **jest** rozwiązywane z `dxfId` (potwierdzone w parserze), a dla reguł wbudowanych `dxf` jest `None`, za to wypełnione są `colorScale`/`dataBar`/`iconSet`.

---

# Moduł 14 — Formatowanie warunkowe zaawansowane: wizualizacje w komórce i pełna kontrola nad regułami

> **Część:** II — Formatowanie · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–13

## 0. W tym module nauczysz się

- **Zrozumiesz różnicę między dwoma światami reguł:** *standardowymi* (malują komórkę przez styl różnicowy `dxf` — znasz je z modułu 13) i *wbudowanymi* (rysują **w komórce mikro-wykres**: skalę kolorów, pasek danych albo ikonę). Zobaczysz, dlaczego wbudowane **nie mają `dxf`** i dlaczego nie przyjmują `font`/`fill`.
- **Opanujesz `ColorScaleRule`** — skale 2- i 3-kolorowe oraz sześć typów progów (`min`, `max`, `num`, `percent`, `percentile`, `formula`). Zrozumiesz, na czym naprawdę polega różnica między `percent` a `percentile` — najczęstsze nieporozumienie w całym formacie.
- **Opanujesz `DataBarRule`** — paski danych, progi, `minLength`/`maxLength`, `showValue` — i usłyszysz wprost, czego openpyxl **nie potrafi** (nowsze warianty pasków, obsługa wartości ujemnych poza podstawową specyfikacją).
- **Opanujesz `IconSetRule`** — wszystkie 17 zestawów ikon, `showValue`, `reverse`, oraz pułapkę „szarych ikon", która psuje czytelność raportów.
- **Dowiesz się, że openpyxl 3.1.x NIE ma klas `Top10Rule`, `AboveAverageRule`, `DateRule`** ani podobnych. Zamiast tego nauczysz się budować te reguły bazowym `Rule` — i to jest wiedza, która odróżnia kurs rzetelny od kursu kopiującego nieistniejące API.
- **Zrozumiesz nakładanie się reguł** na konkretnym scenariuszu: skala kolorów na całej kolumnie **plus** wyróżnienie wartości skrajnych. Zobaczysz, dlaczego reguła wyróżnienia musi mieć **niższy numer priorytetu** i jak to sprawdzić eksperymentalnie.
- **Poznasz `DifferentialStyle` w pełnej, sześcioelementowej postaci** (a nie czteroelementowej, jak upraszczałem w module 13) oraz dowiesz się, dlaczego **nie da się** użyć `NamedStyle` w formatowaniu warunkowym.
- **Nauczysz się czytać reguły po wczytaniu pliku** i **diagnozować rozbieżności** między modelem openpyxl a tym, co widać w Excelu.
- **Zobaczysz wzorce skalowania** (koszt ≈ liczba reguł × rozmiar zakresu) i **zapowiedź wzorca Specification** — reguły budowane z deklaratywnej konfiguracji.

## 1. Intuicja i analogia

### 1.1. Wbudowane formatowanie to mikro-wykres, nie kolor

W module 13 mówiliśmy o regułach, które **malują** komórkę: czerwone tło, pogrubienie, ramka. Dziś wchodzimy w coś jakościowo innego.

Wyobraź sobie **pasek postępu** na stronie internetowej. Gdy strona się ładuje, nie widzisz liczby „73%". Widzisz **pasek**, który wypełnił się w trzech czwartych. Liczba gdzieś tam jest, ale twoje oko czyta **kształt**.

Wbudowane formatowanie warunkowe robi dokładnie to — **zamienia komórkę w mikro-wykres**:

- **Skala kolorów** (`colorScale`) to **mapa cieplna**. Jak mapa pogodowa, na której temperaturę pokazuje się kolorem: zimne regiony na niebiesko, gorące na czerwono. W komórce: mała wartość → jeden kolor, duża → drugi, środek → trzeci. Nie ma tu żadnej ikony ani paska — jest **gradient koloru tła**.
- **Pasek danych** (`dataBar`) to **ten pasek postępu wpisany w komórkę**. Wartość zajmuje część szerokości komórki, proporcjonalnie do swojego miejsca w zakresie.
- **Zestaw ikon** (`iconSet`) to **sygnalizacja świetlna**. Zielone światło = dobrze, czerwone = źle. Albo strzałka w górę / w dół. Albo gwiazdki jak w ocenach.

> **Kluczowa różnica wobec modułu 13:** reguła standardowa (np. `CellIsRule`) **maluje** komórkę i dlatego potrzebuje **stylu różnicowego** (`dxf`) — mówi „co nałożyć". Reguła wbudowana **rysuje w komórce kształt albo gradient** i dlatego potrzebuje **własnych danych** — progów i kolorów/ikon. To dwa różne światy.

Z tej różnicy płyną trzy konsekwencje praktyczne, od których zależeć będzie reszta modułu:

**Konsekwencja 1 — wbudowana reguła nie ma `dxf`.** Jeśli po wczytaniu pliku zobaczysz `regula.dxf is None`, ale `regula.colorScale` niepuste — to nie błąd. To znaczy, że masz przed sobą regułę wbudowaną, która stylu różnicowego w ogóle nie potrzebuje. To ważne, bo w module 13 nauczyłeś się czytać kolor z `rule.dxf` — przy regułach wbudowanych ta droga prowadzi donikąd, kolor siedzi gdzie indziej.

**Konsekwencja 2 — `ColorScaleRule`, `DataBarRule` i `IconSetRule` nie przyjmują `font`, `fill`, `border` ani `dxf`.** Nie ma tu czego „dokładać", bo kolor jest częścią samej wizualizacji. To najczęstsze źródło `TypeError` przy pierwszym kontakcie:

```python
# To NIE zadziala - TypeError: unexpected keyword argument 'fill'
# ColorScaleRule(start_type="min", end_type="max", fill=czerwone)
```

**Konsekwencja 3 — `DataBarRule` i `IconSetRule` nie przyjmują też `stopIfTrue`.** To ma znaczenie: przy wbudowanych regułach nie da się w prosty sposób „ucinać dalszego sprawdzania" tym parametrem. Zobaczysz w 2.9, jak sobie z tym radzić (przez ostrożne projektowanie zakresów i kolejności).

### 1.2. Progi (cfvo) to linijka z podziałkami

Skąd reguła wbudowana wie, jaki kolor narysować? Z **progów**. Każdy próg mówi: „do tego miejsca jeden kolor, od tego miejsca drugi".

Wyobraź sobie **linijkę**. Kreski na linijce to progi. Ale tu jest sztuczka: **od czego mierzysz kreski?** To pytanie prowadzi do sześciu typów progów, których musisz się nauczyć raz na zawsze. Cztery najważniejsze:

- **`min` / `max`** — „skrajna wartość w zakresie". Próg, który **sam się dopasowuje**: zawsze na minimum i na maksimum. Analogia: „od początku linijki do jej końca", niezależnie od tego, jak długa jest linijka. To najbezpieczniejszy wybór dla skal kolorów i pasków, bo nie musisz znać danych.
- **`num`** — „konkretna liczba". Próg na stałe, np. `val=100`. Analogia: kreska dokładnie na 100 milimetrach. Niezależny od danych.
- **`percent`** — „procent **zakresu** (od min do max)". Próg na `50` oznacza „w połowie drogi między najmniejszą a największą wartością". Analogia: „kreska w połowie **pudełka**", a nie „w połowie wartości".
- **`percentile`** — „**percentyl**", czyli próg **rankingowy**. `percentile=80` oznacza „wartość, poniżej której znajduje się 80% danych". Analogia: „przy tej ocenie jesteś lepszy od 80% klasy" — to mówi o **pozycji w rankingu**, nie o odległości od minimum.
- **`formula`** — „wartość wyliczona formułą". Próg wskazujący na komórkę albo wyrażenie (np. `A1`). Używane rzadko; do skomplikowanych przypadków.

**Różnica `percent` vs `percentile` to jedno z najbardziej mylących miejsc w całym formacie OOXML.** Wrócimy do niej szczegółowo w 2.3. Na razie zapamiętaj obrazowo: **`percent` patrzy na szerokość pudełka, `percentile` patrzy na kolejność w kolejce.**

### 1.3. Nakładanie reguł: warstwy farby

W module 13 poznałeś triage — kolejność rozpatrywania reguł. Dziś dodajemy drugi obraz, bo w module 14 reguły nakładają się znacznie częściej.

Wyobraź sobie, że malujesz ścianę **kilkoma warstwami farby**. Gdy dwie warstwy kładą ten sam kolor w tym samym miejscu, wygrywa ta **położona później** — ta z wierzchu. W Excelu numeracja jest odwrócona (mniejszy numer = wyżej), ale mechanizm ten sam: **wygrywa reguła o wyższym priorytecie, czyli o niższym numerze.**

Typowy scenariusz tego modułu:

- **Reguła 1** (priorytet 1): „wyróżnij wartości skrajne" — konkretny, solidny kolor tła.
- **Reguła 2** (priorytet 2): skala kolorów na całej kolumnie — gradient tła.

Obie chcą ustawić **tło** komórki. Warstwa z priorytetem 1 leży na wierzchu, więc na wartościach skrajnych zobaczysz **solidny kolor**, a na pozostałych — **gradient**. Gdybyś dał skali priorytet 1, gradient przesłoniłby wyróżnienie i wyglądałoby to jak „wyróżnienie nie działa". Dokładnie o tym mówi brief: **reguła wyróżnienia musi mieć niższy numer priorytetu.**

### 1.4. Deklaratywna konfiguracja to formularz zamówienia

Analogia na zapowiedź wzorca Specification (moduł 25).

Wyobraź sobie restaurację. Klient **nie gotuje**. Klient wypełnia **formularz zamówienia**: „danie: zupa; ostrość: średnia; dodatki: grzanki". Kuchnia czyta formularz i realizuje.

W module 14 zbudujesz coś analogicznego: **opis reguł w zwykłej liście słowników** (albo w YAML), a osobny kod tłumaczy ten opis na obiekty `Rule` i wywołania `ws.conditional_formatting.add(...)`. Konfiguracja mówi **co**, kod wie **jak**. Zobaczysz to w wyzwaniu 🔴. Na razie trzymaj w głowie obraz: **formularz zamówienia → kuchnia → danie.**

## 2. Teoria

### 2.1. Trzy rodziny reguł — przypomnienie i ważne sprostowanie

W module 13 podzieliliśmy reguły na trzy rodziny. Wróćmy do tego podziału, bo teraz wchodzimy w dwie z nich naraz:

| Rodzina | Co robi | `type` | Jak tworzyć w openpyxl |
|---|---|---|---|
| **Standardowe** | Malują komórkę przez `dxf` | `cellIs`, `expression`, `top10`, `aboveAverage`, `duplicateValues`, `containsText`, `timePeriod`, ... | `CellIsRule`, `FormulaRule`, **albo bazowy `Rule`** |
| **Wbudowane** | Rysują w komórce mikro-wykres | `colorScale`, `dataBar`, `iconSet` | `ColorScaleRule`, `DataBarRule`, `IconSetRule` |

**Sprostowanie, które muszę wyłożyć wprost.** Brief do tego modułu (i wiele tutoriali w internecie) wymienia klasy `Top10Rule`, `AboveAverageRule`, `BelowAverageRule`, `DuplicateValuesRule`, `TextRule`, `ContainsText`, `DateRule`, `TimePeriod`. **Te klasy nie istnieją w openpyxl 3.1.x.** Łatwo to sprawdzić — zajrzyj do źródła `openpyxl/formatting/rule.py`: są tam dokładnie trzy funkcje fabryczne dla rodzin standardowych i wbudowanych:

```python
CellIsRule(operator, formula, stopIfTrue, font, border, fill)
FormulaRule(formula, stopIfTrue, font, border, fill)
ColorScaleRule(start_type, start_value, start_color, mid_type, mid_value, mid_color,
               end_type, end_value, end_color)
DataBarRule(start_type, start_value, end_type, end_value, color, showValue,
            minLength, maxLength)
IconSetRule(icon_style, type, values, showValue, percent, reverse)
```

I to wszystko. **Nie ma `Top10Rule`.** Nie ma `DateRule`. Nie ma `DuplicateValuesRule`. Wszystkie te zachowania realizujesz przez bazowy `Rule` z odpowiednim `type` — co pokażę szczegółowo w 2.7.

> **Dlaczego to ważne i dlaczego jest dobrą lekcją.** Gdybyś kopiował kurs z internetu i wpisał `from openpyxl.formatting.rule import Top10Rule`, dostałbyś `ImportError`. To jest **dokładnie ta kategoria z modułu 03**: „nie wierz dokumentacji ani tutorialowi — zweryfikuj w źródle swojej wersji". Tutaj weryfikacja wygląda tak: zajrzyj do `rule.py` i policz funkcje. Jest ich pięć.

A zarazem to dobra wiadomość: **bazowy `Rule` pokrywa wszystkie 18 typów** z listy dozwolonych wartości. Musisz tylko wiedzieć, który parametr ustawić dla którego typu. To opanujesz w 2.7.

### 2.2. `FormatObject` — pojedynczy próg

Zanim wejdziemy w `ColorScaleRule`, `DataBarRule` i `IconSetRule`, musisz poznać cegiełkę wspólną: **`FormatObject`** (w formacie OOXML nazywa się `cfvo` — *conditional formatting value object*).

```python
from openpyxl.formatting.rule import FormatObject

FormatObject(type="min")                # prog: minimum zakresu
FormatObject(type="max")                # prog: maksimum zakresu
FormatObject(type="num", val=100)       # prog: liczba 100
FormatObject(type="percent", val=50)    # prog: 50% zakresu (polowa drogi min..max)
FormatObject(type="percentile", val=80) # prog: 80. percentyl
FormatObject(type="formula", val="A1")  # prog: wartosc z komorki A1
```

**Parametry `FormatObject`:**

| Parametr | Znaczenie |
|---|---|
| `type` | Jeden z: `"num"`, `"percent"`, `"max"`, `"min"`, `"formula"`, `"percentile"`. **Wymagany.** |
| `val` | Wartość progu. **Liczba** dla `num`/`percent`/`percentile`; **napisy** (odwołanie do komórki) dla `formula`; **`None`** dla `min`/`max` |
| `gte` | „większe lub równe" — flaga dopasowania. Zwykle `None` (rzadko potrzebna) |

**Pułapka `val`.** Wewnętrznie openpyxl używa `ValueDescriptor`: jeśli `type == "formula"` **albo** wartość wygląda jak adres komórki (`"A1"`), oczekiwany typ to `str`; inaczej `float`. Czyli:

```python
FormatObject(type="num", val=100)        # ✅ liczba
FormatObject(type="num", val="100")      # ❌ TypeError - "100" nie jest adresem komorki
FormatObject(type="formula", val="A1")   # ✅ adres komorki jako formula
FormatObject(type="min")                 # ✅ bez val
```

Czyli dla progów liczbowych podawaj **liczby**, nie napisy. Dla `num`/`percent`/`percentile` — liczby. Dla `formula` — napis będący adresem albo wyrażeniem.

**Kiedy używasz `FormatObject` bezpośrednio?** Rzadko. Fabryki `ColorScaleRule`, `DataBarRule` i `IconSetRule` budują obiekty `FormatObject` za ciebie z prostszych parametrów. Ale gdy potrzebujesz **mieszanych typów progów** w jednym zestawie (np. pierwszy próg `min`, drugi `percentile`, trzeci `num`) — musisz zbudować `ColorScale` ręcznie, o czym w 2.4.

### 2.3. `percent` vs `percentile` vs `num` — rozłóżmy to raz na zawsze

To najważniejsza podsekcja teoretyczna tego modułu. Weźmy konkretne dane: `[10, 20, 30, 40, 500]`.

- **`min`** = `10`, **`max`** = `500`. Zakres = 490.
- **`percent=50`** = `10 + 0,5 × 490 = 255`. Połowa **drogi** między min a max.
- **`percentile=50`** = `30`. Mediana — wartość, poniżej której leży 50% danych. A medianą jest `30`, bo trzy z pięciu liczb są mniejsze lub równe 30.
- **`num=100`** = `100`. Po prostu sto.

Popatrz, jak bardzo `255` i `30` się różnią. Jedna przerwa `500` ciągnie `percent` daleko w górę, ale `percentile` prawie nie drgnie, bo liczy się **pozycja**, nie odległość.

**Praktyczna wskazówka wyboru progu:**

| Chcesz… | Użyj |
|---|---|
| Skala od najniższej do najwyższej wartości w danych | `min` / `max` |
| Kolor zależny od udziału w przedziale (np. „pełny pasek przy 80% zakresu") | `percent` |
| Progów odpornych na pojedyncze skrajne wartości (outliery) | `percentile` |
| Stałego, z góry znanego progu (np. budżet 1 000 000) | `num` |
| Progu wyliczanego gdzieś indziej w arkuszu | `formula` |

**Kiedy `percentile` jest lepszy od `percent`?** Zawsze, gdy w danych są **outliery**. Wyobraź sobie pensje w firmie: 90 osób zarabia 3–8 tys., a prezes 200 tys. Jeśli zrobisz pasek danych od `min` do `max`, wszystkie paski poza prezesem będą maleńkie — bo prezes rozciągnął skalę. Jeśli użyjesz `end_type="percentile", end_value=90`, skala „nasyci się" na 90. percentylu i paski znów będą czytelne. To bardzo praktyczny trik — pokażę go w Przykładzie 2.

### 2.4. `ColorScaleRule` — skala kolorów

Skala kolorów to gradient tła. Może mieć **dwa kolory** (od jednego do drugiego) albo **trzy** (z kolorem środkowym — dwie skale gradientowe).

**Sygnatura (to funkcja fabryczna, nie klasa!):**

```python
ColorScaleRule(start_type=None, start_value=None, start_color=None,
               mid_type=None, mid_value=None, mid_color=None,
               end_type=None, end_value=None, end_color=None)
```

**Reguły używania:**

1. **Musisz podać `start_type` i `end_type`.** Bez nich openpyxl zgłosi `ValueError("Both start and end colors must be set")` (a dokładniej: „At least two cfvo are needed").
2. **Kolory muszą być w tej samej liczbie co typy.** Dwa typy → dwa kolory. Trzy typy → trzy kolory.
3. **`mid_*` podajesz tylko dla skali 3-kolorowej.** Kolejność typów w XML musi być rosnąca: start (najmniejsza wartość), mid (środek), end (największa). openpyxl buduje `FormatObject` w kolejności, w jakiej podasz parametry — więc **pisz je w kolejności od najmniejszej do największej wartości progu**.
4. **Kolory podawaj jako napis** w formacie `RRGGBB` (6 znaków) albo `AARRGGBB` (8 znaków, z kanałem alfa). Najbezpieczniej: 8 znaków, np. `"FF63BE7B"`. Fabryka sama opakuje napis w `Color`. Możesz też przekazać obiekt `Color` — zadziała tak samo.

**Dwa przykłady — 2 i 3 kolory:**

```python
from openpyxl.formatting.rule import ColorScaleRule

# Skala 2-kolorowa: min -> max (od zoltozielonego do czerwonego)
ws.conditional_formatting.add(
    "B2:B11",
    ColorScaleRule(start_type="min", start_color="FF63BE7B",
                   end_type="max", end_color="FFF8696B"),
)

# Skala 3-kolorowa: min -> srodek na 50. percentylu -> max
ws.conditional_formatting.add(
    "C2:C11",
    ColorScaleRule(
        start_type="min",      start_color="FFF8696B",   # czerwony (dol)
        mid_type="percentile", mid_value=50,
        mid_color="FFFFEB84",                            # zolty (srodek)
        end_type="max",        end_color="FF63BE7B",     # zielony (gora)
    ),
)
```

**Najczęstsza pułapka: odwrócone kolory.** Ludzie intuicyjnie kojarzą „dużo = dobrze = zielone", ale w Excelu domyślny gradient idzie „mało = czerwone, dużo = zielone" tylko wtedy, gdy **sam tak go zdefiniujesz**. Kolejność ma znaczenie: `start_color` dotyczy **najmniejszej** wartości, `end_color` — **największej**. Jeśli zobaczysz w raporcie, że najgorsze wyniki są zielone, to znaczy, że zamieniłeś kolory miejscami. To trywialne, a zdarza się bez przerwy.

**Skala 3-kolorowa bez `mid_*`** nie zadziała — dostaniesz skalę 2-kolorową albo błąd. openpyxl wymaga spójnej liczby progów i kolorów.

**Kiedy ręcznie, a kiedy fabryka?** Fabryka wystarcza w 95% przypadków. Ręczne budowanie `ColorScale` + `FormatObject` potrzebujesz, gdy progi mają **różne typy** (np. start `min`, mid `num=50`, end `percentile=90`) w jednej skali — fabryka też to potrafi, bo typy podajesz osobno (`start_type`, `mid_type`, `end_type`), więc w praktyce fabryka wystarcza niemal zawsze.

### 2.5. `DataBarRule` — pasek w komórce

Pasek danych to **wykres zrobiony z wypełnienia komórki** — wartość rośnie, pasek się wydłuża.

**Sygnatura:**

```python
DataBarRule(start_type=None, start_value=None, end_type=None,
            end_value=None, color=None, showValue=None,
            minLength=None, maxLength=None)
```

| Parametr | Znaczenie |
|---|---|
| `start_type`, `start_value` | Typ i wartość dolnego progu (skąd pasek startuje) |
| `end_type`, `end_value` | Typ i wartość górnego progu (gdzie pasek jest pełny) |
| `color` | Kolor paska — napis `RRGGBB` albo `AARRGGBB` |
| `showValue` | `True`/`False` — czy pokazać liczbę obok paska (domyślnie w Excelu: `True`) |
| `minLength` | Minimalna długość paska w procentach (0–100) |
| `maxLength` | Maksymalna długość paska w procentach (0–100) |

**Dwa progi wystarczą** — w odróżnieniu od skali kolorów, gdzie było pojęcie koloru środkowego. `DataBarRule` buduje dokładnie dwa `FormatObject` (`start` i `end`), więc **nie ma tu progu środkowego**.

**Przykład — pasek z nasyceniem na 90. percentylu** (odporny na outliery — trik z 2.3):

```python
from openpyxl.formatting.rule import DataBarRule

ws.conditional_formatting.add(
    "B2:B100",
    DataBarRule(
        start_type="num", start_value=0,        # pasek startuje od zera
        end_type="percentile", end_value=90,    # pelny pasek na 90. percentylu
        color="FF638EC6",
        showValue=True,
    ),
)
```

Jeśli chcesz, żeby skala była względna do danych, użyj `start_type="min"` i `end_type="max"`. Jeśli chcesz, żeby pasek startował od zera — `start_type="num", start_value=0`.

**Czego openpyxl NIE potrafi w paskach danych — powiem wprost.** Dokumentacja 3.1 stwierdza: *„Currently, openpyxl supports the DataBars as defined in the original specification. Borders and directions were added in a later extension."* W praktyce oznacza to:

- **Nie ustawisz** nowszych wariantów pasków: gradient vs pełny kolor, kolor obramowania, kierunek paska, **kolor osi i obsługa wartości ujemnych** (dla liczb ujemnych Excel rysuje pasek od osi — openpyxl nie da ci kontroli nad kolorem tej osi).
- Pasek na wartościach ujemnych nadal się pojawi (Excel go narysuje), ale **nie masz nad nim pełnej kontroli** — dostajesz zachowanie domyślne.

To trzecia kategoria z modułu 03 — **„openpyxl potrafi tylko częściowo"**. Jeśli raport wymaga pasków z pełną kontrolą nad wartościami ujemnymi, rozważ inny generator (np. `xlsxwriter`) albo traktuj to jako znane ograniczenie i sprawdź wynik w Excelu.

**Pułapka z dokumentacji.** Przykład w oficjalnej dokumentacji podaje `showValue="None"` — jako **napis**. To błąd w dokumentacji; `showValue` to `Bool`. Jeśli skopiujesz, Excel może potraktować napis „None" jako prawdę. **Podawaj `True`/`False`.** To kolejny dowód na tezę z modułu 03 — dokumentacja i kod nie zawsze mówią to samo.

**Pułapka `minLength`/`maxLength`.** To liczby całkowite oznaczające **procent szerokości komórki** (0–100), nie piksele. Domyślnie Excel używa około 10 i 90. Nie wszystkie wersje Excela honorują je identycznie — sprawdź w swojej wersji.

### 2.6. `IconSetRule` — sygnalizacja świetlna

Zestaw ikon wyświetla w komórce ikonę (strzałka, światło, flaga, gwiazdka) zależną od progu.

**Sygnatura:**

```python
IconSetRule(icon_style=None, type=None, values=None,
            showValue=None, percent=None, reverse=None)
```

| Parametr | Znaczenie |
|---|---|
| `icon_style` | **Nazwa zestawu ikon** (np. `"3TrafficLights1"`, `"5Arrows"`) |
| `type` | **Typ progów** (`"percent"`, `"num"`, `"percentile"`, `"min"`, `"max"`, `"formula"`). Wspólny dla wszystkich progów! |
| `values` | **Lista wartości progów.** Liczba elementów musi się zgadzać z liczbą ikon w zestawie |
| `showValue` | `True`/`False` — czy pokazać liczbę obok ikony |
| `percent` | Flaga formatu (procenty). Zwykle zostaw `None` |
| `reverse` | `True` — odwróć kolejność ikon (duże wartości = ikona „dolna") |

**⚠️ Największa pułapka tego zestawu parametrów.** Łatwo pomylić `icon_style` (nazwa ikon) z `type` (typ progu). To **dwie różne rzeczy**:

```python
# icon_style = nazwa ikon; type = typ progow; values = progi
IconSetRule("5Arrows", "percent", [0, 20, 40, 60, 80])
#            ^^^^^^^^  ^^^^^^^^^   ^^^^^^^^^^^^^^^^^^^
#            ikony      typ progow  lista wartosci (5 szt. dla 5 ikon)
```

Jeśli zamienisz je miejscami, dostaniesz `ValidationError` albo pusty zestaw w Excelu.

**Wszystkie 17 zestawów ikon** (dokładna lista z openpyxl 3.1.x) i **liczba wymaganych progów**:

| Zestaw | Ikony | Progi |
|---|---|---|
| `3Arrows`, `3ArrowsGray`, `3Flags` | 3 | 3 |
| `3TrafficLights1`, `3TrafficLights2` | 3 | 3 |
| `3Signs`, `3Symbols`, `3Symbols2` | 3 | 3 |
| `4Arrows`, `4ArrowsGray`, `4RedToBlack` | 4 | 4 |
| `4Rating`, `4TrafficLights` | 4 | 4 |
| `5Arrows`, `5ArrowsGray` | 5 | 5 |
| `5Rating`, `5Quarters` | 5 | 5 |

Zauważ **„warianty z/bez tła"**, o których mówi brief. Excel dzieli ikony na *rimmed* (z ramką/obramowaniem) i *unrimmed* (bez ramki). W nazwach openpyxl widać to najczęściej jako warianty numerowane:

- `3TrafficLights1` — z ramką; `3TrafficLights2` — bez ramki,
- `3Symbols` — w kółkach; `3Symbols2` — bez kółek.

**Nie ma tu pełnej systematyczności** — nazwy typu `3Arrows` vs `3ArrowsGray` różnią się **kolorem ikon** (kolorowe vs szare), a nie ramką. Dlatego **zawsze obejrzyj efekt w Excelu**, zamiast zakładać, co robi dana nazwa.

**Liczba progów musi się zgadzać z liczbą ikon.** To pułapka, która kończy się „uszkodzonym plikiem":

```python
IconSetRule("3TrafficLights1", "percent", [0, 33, 67])   # ✅ 3 progi dla 3 ikon
IconSetRule("3TrafficLights1", "percent", [0, 50])       # ❌ 2 progi dla 3 ikon!
```

Gdy liczba się nie zgadza, Excel przy otwarciu może zaproponować „naprawę" pliku. Zawsze `len(values) == liczba ikon`.

**Pierwszy próg zwykle startuje od zera.** Dla `type="percent"` pierwszy próg to najczęściej `0`. Jeśli zaczniesz od `10`, wartości poniżej 10 **nie dostaną żadnej ikony** — komórka pozostanie pusta. Często to niezamierzone.

**`showValue` a „szare ikony" — pułapka z briefu.** Zestawy z `...Gray` (`3ArrowsGray`, `5ArrowsGray`, `4ArrowsGray`) mają ikony w odcieniach szarości. Jeśli dodatkowo ustawisz `showValue=False` (tylko ikona, bez liczby), dostaniesz komórkę, w której jest **sama szara strzałka bez liczby i bez tła** — wygląda to jak wyłączona, niedziałająca komórka. Klasyczny efekt „raport wygląda na zepsuty". Zasady zdrowego rozsądku:

- **Nie łącz `...Gray` z `showValue=False`** — albo pokaż liczbę, albo użyj kolorowego zestawu.
- Pamiętaj, że **zestaw ikon nie maluje tła**. Jeśli chcesz i ikonę, i kolor tła, potrzebujesz **drugiej reguły** (standardowej, z `fill`). To kolejny powód, dlaczego reguły się nakładają — i dlaczego wracamy do 2.9.

**`reverse`.** Ustaw `reverse=True`, gdy większa wartość ma dostawać „gorszą" ikonę (np. większe opóźnienie → czerwona strzałka). Zamiast tego możesz odwrócić listę progów, ale `reverse` jest jaśniejsze w intencji.

**`percent` a `type`.** Fabryka `IconSetRule` stosuje **ten sam `type` do wszystkich progów**. To znaczy: **nie zrobisz mieszanych typów progów** w jednym zestawie ikon przez fabrykę (np. pierwszy `min`, reszta `percent`). Gdy tego potrzebujesz — buduj `IconSet` ręcznie z listy `FormatObject` (analogicznie jak w 2.4). W praktyce dla zestawów ikon prawie zawsze używa się `type="percent"` i równych przedziałów.

**Praktyczna uwaga o `percent`.** Parametr `percent` na poziomie `IconSet` jest w formacie OOXML flagą mówiącą, czy progi są procentami. W większości przypadków, gdy używasz `type="percent"`, możesz zostawić `percent=None` (Excel wywnioskuje z typów progów). Jeśli w Twojej wersji Excel nie stosuje ikon, spróbuj `percent=True`. **Sprawdź w dokumentacji swojej wersji** — to jeden z tych parametrów, gdzie zachowanie bywa nieoczywiste.

### 2.7. Reguły bez klasy fabrycznej — bazowy `Rule`

Teraz najważniejsza część praktyczna: **rodziny standardowe, dla których NIE ma fabryki.** Budujesz je bazowym `Rule`, ustawiając `type` i odpowiednie atrybuty.

Przypomnijmy pełny konstruktor `Rule` (ze źródła 3.1.x):

```python
Rule(type, dxfId=None, priority=0, stopIfTrue=None, aboveAverage=None,
     percent=None, bottom=None, operator=None, text=None, timePeriod=None,
     rank=None, stdDev=None, equalAverage=None, formula=(),
     colorScale=None, dataBar=None, iconSet=None, extLst=None, dxf=None)
```

`type` jest **wymagany** (pozycyjny). Reszta to atrybuty — i każdy `type` używa innego podzbioru. Oto **katalog** najważniejszych:

#### `top10` — TOP N / dolne N

Podświetla pierwsze N wartości (albo N procent).

```python
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import PatternFill

dxf_top = DifferentialStyle(
    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
)

# Najwyzsze 10 wartosci
rule = Rule(type="top10", rank=10, percent=False, bottom=False, dxf=dxf_top)
ws.conditional_formatting.add("B2:B100", rule)

# Najwyzsze 10 PROCENT (a nie 10 szt.)
rule = Rule(type="top10", rank=10, percent=True, bottom=False, dxf=dxf_top)

# Najnizsze 5 wartosci
rule = Rule(type="top10", rank=5, percent=False, bottom=True, dxf=dxf_top)
```

| Atrybut | Znaczenie |
|---|---|
| `rank` | Ile pozycji (albo ile procent) |
| `percent` | `True` — `rank` to procent, `False` — to liczba sztuk |
| `bottom` | `True` — dolne N zamiast górnych N |

#### `aboveAverage` / `belowAverage` — powyżej/poniżej średniej

`aboveAverage` obsługuje **oba** kierunki — jest jeden typ, nie dwa:

```python
# Powyzej sredniej
rule = Rule(type="aboveAverage", aboveAverage=True, dxf=dxf_top)

# Ponizej sredniej
rule = Rule(type="aboveAverage", aboveAverage=False, dxf=dxf_top)

# Powyzej sredniej + rowna sredniej + odchylenie standardowe
rule = Rule(type="aboveAverage", aboveAverage=True, equalAverage=True, stdDev=1, dxf=dxf_top)
```

| Atrybut | Znaczenie |
|---|---|
| `aboveAverage` | `True` — powyżej; `False` — poniżej |
| `equalAverage` | `True` — wlicz wartość równą średniej |
| `stdDev` | Odchylenia standardowe zamiast prostej średniej |

**To wyjaśnia, dlaczego nie ma `BelowAverageRule`** — bo `aboveAverage=False` załatwia sprawę jednym typem.

#### `duplicateValues` / `uniqueValues` — duplikaty / unikaty

Bez dodatkowych atrybutów — sam `type` i `dxf`:

```python
rule_dup = Rule(type="duplicateValues", dxf=dxf_top)
ws.conditional_formatting.add("A2:A100", rule_dup)

rule_uniq = Rule(type="uniqueValues", dxf=dxf_top)
```

#### `containsText` i warianty tekstowe

Tu wchodzi `operator`, `text` **i** `formula`. Uwaga: dla reguł tekstowych Excel oczekuje **formuły** (bo sam warunek tekstowy nie wystarcza do ewaluacji w OOXML).

```python
rule = Rule(
    type="containsText",
    operator="containsText",
    text="pilne",
    dxf=dxf_top,
)
rule.formula = ['NOT(ISERROR(SEARCH("pilne",A2)))']
ws.conditional_formatting.add("A2:F100", rule)
```

| `operator` | Kiedy użyć |
|---|---|
| `containsText` | Zawiera tekst |
| `notContains` | Nie zawiera |
| `beginsWith` | Zaczyna się od |
| `endsWith` | Kończy się na |

Katalog formuł dla reguł tekstowych (poprawne pary `operator` ↔ `formula`):

| Sytuacja | `operator` | `formula` |
|---|---|---|
| Zawiera „abc" | `containsText` | `NOT(ISERROR(SEARCH("abc",A2)))` |
| Nie zawiera „abc" | `notContains` | `ISERROR(SEARCH("abc",A2))` |
| Zaczyna się od „abc" | `beginsWith` | `LEFT(A2,3)="abc"` |
| Kończy się na „abc" | `endsWith` | `RIGHT(A2,3)="abc"` |

**Ważne:** `SEARCH` jest **niewrażliwe na wielkość liter** (`"PILNE"` i `"pilne"` zadziałają tak samo). Jeśli chcesz porównanie wrażliwe na wielkość liter, użyj `FIND` zamiast `SEARCH`.

#### `containsBlanks`, `notContainsBlanks`, `containsErrors`, `notContainsErrors`

```python
# Puste komorki
rule = Rule(type="containsBlanks", dxf=dxf_top)
rule.formula = ['LEN(TRIM(A2))=0']
ws.conditional_formatting.add("A2:A100", rule)

# Komorki z bledami (#N/A, #DIV/0! itd.)
rule = Rule(type="containsErrors", dxf=dxf_top)
rule.formula = ['ISERROR(A2)']
```

#### `timePeriod` — okresy czasu

Tu jest najciekawiej, bo masz **dwa sposoby**.

**Sposób 1 — przez atrybut `timePeriod`** (Excel rozumie to sam):

```python
rule = Rule(type="timePeriod", timePeriod="today", dxf=dxf_top)
ws.conditional_formatting.add("C2:C100", rule)
```

Dozwolone wartości `timePeriod` (dokładna lista z 3.1.x): `today`, `yesterday`, `tomorrow`, `last7Days`, `thisMonth`, `lastMonth`, `nextMonth`, `thisWeek`, `lastWeek`, `nextWeek`. **To tyle — nie ma tu `thisYear`/`nextYear`.**

**Sposób 2 — przez `FormulaRule`** (moduł 13), z własną logiką daty. Bezpieczniejszy, bo nie zależy od tego, jak Excel interpretuje `timePeriod`:

```python
from openpyxl.formatting.rule import FormulaRule

# Data minela (bezpieczna, przewidywalna)
ws.conditional_formatting.add(
    "C2:C100",
    FormulaRule(formula=["$C2<TODAY()"], fill=czerwone, stopIfTrue=True),
)
```

**Rekomendacja:** dla prostych przypadków (`today`, `yesterday`, `tomorrow`) użyj `Rule(type="timePeriod", ...)` — jest czytelniej. Dla „za mniej niż N dni", „w tym kwartale" albo gdy chcesz mieć pewność co do logiki — **użyj `FormulaRule` z modułu 13**. To ta sama motywacja co w module 06: nie licz na magię, powiedz wprost, czego chcesz.

#### Podsumowanie katalogu `Rule`

| Zachowanie | `type` | Kluczowe atrybuty | Fabryka? |
|---|---|---|---|
| Porównanie ze stałą | `cellIs` | `operator`, `formula` | `CellIsRule` ✅ |
| Wyrażenie własne | `expression` | `formula` | `FormulaRule` ✅ |
| TOP N | `top10` | `rank`, `percent`, `bottom` | ❌ ręcznie |
| Powyżej/poniżej średniej | `aboveAverage` | `aboveAverage`, `equalAverage`, `stdDev` | ❌ ręcznie |
| Duplikaty | `duplicateValues` | — | ❌ ręcznie |
| Unikaty | `uniqueValues` | — | ❌ ręcznie |
| Tekst | `containsText`, `beginsWith`, `endsWith`, `notContainsText` | `operator`, `text`, `formula` | ❌ ręcznie |
| Puste komórki | `containsBlanks`, `notContainsBlanks` | `formula` | ❌ ręcznie |
| Błędy | `containsErrors`, `notContainsErrors` | `formula` | ❌ ręcznie |
| Okres czasu | `timePeriod` | `timePeriod` | ❌ ręcznie |
| Skala kolorów | `colorScale` | `colorScale` (obiekt) | `ColorScaleRule` ✅ |
| Pasek danych | `dataBar` | `dataBar` (obiekt) | `DataBarRule` ✅ |
| Zestaw ikon | `iconSet` | `iconSet` (obiekt) | `IconSetRule` ✅ |

**Pięć fabryk. Trzynaście typów.** To zdanie zapamiętaj — oszczędzi ci `ImportError`.

### 2.8. `DifferentialStyle` w pełnej postaci

W module 13 pokazałem `DifferentialStyle` jako zestaw **czterech** składników (`font`, `fill`, `border`, `numFmt`) i napisałem, że nie ma tam `alignment`. Zaglądając do źródła 3.1.x, widzimy **pełną, sześcioelementową sygnaturę** — dokładnie taką samą jak styl komórki z modułu 09:

```python
DifferentialStyle(font=None, numFmt=None, fill=None,
                  alignment=None, border=None, protection=None, extLst=None)
```

**Sprostowanie i lekcja.** Cztery składniki to **uproszczenie praktyczne** — to te, które faktycznie ma sens ustawiać w regułach. Ale model strukturalny przyjmuje sześć, bo w formacie OOXML `<dxf>` może zawierać dokładnie te same elementy co zwykły styl komórki. **Zapisuję to jako sprostowanie, bo to dokładnie ta kategoria z modułu 03: „zweryfikuj w źródle".** Mój własny uproszczony opis w module 13 był właśnie tym, przed czym ostrzegam.

**Które składniki mają sens w praktyce?**

| Składnik | Sens w CF? | Uwagi |
|---|---|---|
| `font` | ✅ Tak | Pogrubienie, kolor, kursywa — najczęściej używane |
| `fill` | ✅ Tak | Tło — najczęściej używane (pamiętaj o `PatternFill(start_color=, end_color=, fill_type="solid")`) |
| `border` | ✅ Tak | Ramki wokół komórek |
| `numFmt` | 🟡 Częściowo | Formalnie działa, ale UI Excela dla CF zwykle nie oferuje zmiany formatu liczby — sprawdź w Excelu |
| `alignment` | 🟡 Ostrożnie | Struktura to przyjmuje, ale UI Excela dla formatowania warunkowego zwykle nie ustawia wyrównania |
| `protection` | 🟡 Ostrożnie | Jak wyżej — rzadko używane |

**Rekomendacja:** skup się na `font`, `fill`, `border`. Reszta jest w strukturze, ale ich zachowanie w Excelu bywa nieoczywiste — traktuj jak eksperyment, nie jak pewnik.

**Dlaczego NIE `NamedStyle`?** To ważna pułapka. Kuszące byłoby zrobić:

```python
# To NIE zadziala!
from openpyxl.styles import NamedStyle
styl = NamedStyle(name="Naglowek", font=Font(bold=True))
wb.add_named_style(styl)
# rule = ... ; rule.dxf = styl   # ❌ BŁĄD - reguła wymaga DifferentialStyle
```

**`NamedStyle` to nie `DifferentialStyle`.** To dwa różne byty, przechowywane w różnych miejscach pliku i stosowane różnymi mechanizmami. Reguła potrafi przyjąć tylko `DifferentialStyle`. Jeśli masz gotowy `NamedStyle` (moduł 09) i chcesz go użyć w regule, musisz **przekonwertować** go ręcznie — zbudować nowy `DifferentialStyle` z tych samych składników. To znana, udokumentowana w społeczności ograniczenie.

**Ciekawostka praktyczna (i powód do optymizmu): deduplikacja dxf.** `DifferentialStyleList.append()` (to lista stylów różnicowych w skoroszycie) **sprawdza, czy identyczny styl już istnieje, i nie dodaje duplikatu**. To dokładnie ten sam mechanizm współdzielenia, o którym mówiłem w module 09 („pieczątka projektowana raz, przybijana w wielu miejscach"). Czyli jeśli dziesięć reguł używa identycznego „czerwonego tła + pogrubienia", w pliku będzie **jeden** `<dxf>`, do którego wszystkie się odwołują. Dobra wiadomość dla rozmiaru pliku — i jeszcze jeden powód, by używać wspólnych, buforowanych stylów (moduł 09).

### 2.9. Nakładanie się reguł i priorytety w praktyce

Wróćmy do analogii warstw farby (1.3). Przyjrzymy się najczęstszemu realnemu konfliktowi: **skala kolorów + wyróżnienie skrajnych**.

**Scenariusz.** Masz kolumnę wyników `B2:B100`. Chcesz:

1. **Mapę cieplną** — cała kolumna w gradiencie od czerwonego (niskie) do zielonego (wysokie).
2. **Dodatkowe wyróżnienie** — najgorsze (np. poniżej progu) wartości na solidny, mocny czerwony, żeby „krzyczały".

Obie reguły chcą ustawić **tło** komórki. Więc wygrywa ta z **niższym numerem priorytetu**.

Rozwiązanie: **najpierw dodaj regułę wyróżnienia, potem skalę kolorów.** Wtedy wyróżnienie dostanie priorytet 1, a skala 2.

```python
# Regula 1 - dodana PIERWSZA -> priorytet 1 -> lezy na wierzchu
ws.conditional_formatting.add(
    "B2:B100",
    FormulaRule(formula=["$B2<60"], fill=mocne_czerwone, stopIfTrue=True),
)

# Regula 2 - dodana DRUGA -> priorytet 2 -> pod spodem
ws.conditional_formatting.add(
    "B2:B100",
    ColorScaleRule(start_type="min", start_color="FFF8696B",
                   end_type="max", end_color="FF63BE7B"),
)
```

**Co zobaczysz w Excelu:** większość komórek w gradiencie, ale te poniżej 60 — na solidny czerwony. Jeśli zamienisz kolejność, gradient przesłoni wyróżnienie i pomyślisz, że „wyróżnienie nie działa". **To najczęstsze źródło frustracji przy łączeniu reguł wbudowanych i standardowych.**

**Uwaga o `stopIfTrue`.** `ColorScaleRule` **nie przyjmuje** `stopIfTrue` (2.1, konsekwencja 3). Ale reguła **wyróżnienia** (która jest `FormulaRule`) przyjmuje — i tutaj `stopIfTrue=True` ma sens: gdy wyróżnienie zadziała, nie ma po co nakładać gradientu na tę komórkę.

**Kolejny realny konflikt: dwie reguły standardowe zmieniające to samo.** Jeśli dwie reguły ustawiają różne tło i obie pasują, wygrywa niższy priorytet. Jeśli ustawiają **różne atrybuty** (jedna tło, druga pogrubienie), **oba się nałożą** — bo warstwy farby nie kolidują, gdy malują co innego.

**`DataBar` a `fill`.** Pasek danych i zwykłe wypełnienie tła **mogą współistnieć** — pasek jest rysowany na wierzchu tła (albo obok). Jeśli więc chcesz i pasek, i kolor tła zależny od innego warunku — zwykle to zadziała. Ale kolejność wpływa na to, co jest widoczne przy niskich wartościach.

**Jak sprawdzić faktyczne priorytety po dodaniu?** Wczytaj plik i przejrzyj (Przykład 4 pokazuje pełny inwentarz):

```python
for cf in ws.conditional_formatting:
    for regula in sorted(cf.rules, key=lambda r: r.priority):
        print(regula.priority, regula.type)
```

**Reguła kciuka dla priorytetów w module 14:**

> **Bardziej szczegółowa reguła zwykle powinna mieć niższy numer priorytetu niż bardziej ogólna.** Wyróżnienie skrajnych (szczegółowe) przed skalą kolorów (ogólne). Wyjątek przed regułą. Reguła „konkretna wartość" przed „zakresem".

### 2.10. Odczyt reguł po wczytaniu pliku i diagnostyka

Umiem już tworzyć reguły. Teraz najważniejsza umiejętność inżynierska: **odczytać je z istniejącego pliku i zrozumieć, co widzę.**

**Co działa po `load_workbook` (potwierdzone w źródle):**

Po wczytaniu pliku parser robi mniej więcej to:

```python
cf = ConditionalFormatting.from_tree(element)
for rule in cf.rules:
    if rule.dxfId is not None:
        rule.dxf = self.differential_styles[rule.dxfId]   # <-- ROZWIAZANIE dxf!
    self.ws.conditional_formatting.add(cf.sqref, rule)
```

Czyli:

- **Dla reguł standardowych** (`cellIs`, `expression`...) `regula.dxf` **jest rozwiązywane** z `dxfId` na obiekt ze `styles.xml`. Możesz czytać `regula.dxf.fill`, `regula.dxf.font` itd. To potwierdza to, co robiliśmy w module 13.
- **Dla reguł wbudowanych** (`colorScale`, `dataBar`, `iconSet`) `dxfId` jest `None`, więc `regula.dxf` pozostaje `None`. Ale wypełnione są **obiekty wbudowane**:

```python
regula.colorScale        # obiekt ColorScale
regula.colorScale.color  # lista Color
regula.colorScale.cfvo   # lista FormatObject (progi)

regula.dataBar           # obiekt DataBar
regula.dataBar.color     # Color
regula.dataBar.cfvo      # lista FormatObject
regula.dataBar.minLength, .maxLength, .showValue

regula.iconSet           # obiekt IconSet
regula.iconSet.iconSet   # nazwa zestawu, np. "3TrafficLights1"
regula.iconSet.cfvo      # lista FormatObject
regula.iconSet.showValue, .reverse, .percent
```

**To jest klucz do diagnostyki:** chcesz kolor skali kolorów — czytaj `regula.colorScale.color`, **nie** `regula.dxf`. Chcesz kolor zwykłej reguły — czytaj `regula.dxf.fill`.

**Dlaczego wczytane reguły mogą NIE odpowiadać temu, co widać w Excelu?** Cztery przyczyny, które trzeba znać:

1. **Reguły rozszerzeń (`extLst`) są niewidoczne dla openpyxl.** Nowsze wersje Excela zapisują część reguł w formacie rozszerzeń — openpyxl ich nie modeluje, więc ich nie odczyta i zgubi przy zapisie (2.11). Objaw: Excel pokazuje podświetlenie, a po wczytaniu openpyxl widzisz mniej reguł w `ws.conditional_formatting`.
2. **Bogate warianty pasków danych i zestawów ikon.** Excel 2019+ ma warianty pasków (gradient, kontrolki dla wartości ujemnych) i reguły, których openpyxl nie zna w pełni. Po wczytaniu zobaczysz uproszczoną wersję albo nic.
3. **Pliki tworzone przez starsze narzędzia** (openpyxl 2.x, `xlsxwriter`, Apache POI) mogą zapisywać struktury nieco inne niż Excel — openpyxl odczyta to, co rozumie, resztę pominie.
4. **Kluczowanie po `sqref`.** Kolekcja łączy reguły o tym samym zakresie w jedną grupę (moduł 13, 2.11). Jeśli ręcznie edytowałeś plik w Excelu, może się okazać, że spodziewasz się trzech grup, a jest jedna z siedmioma regułami.

**Praktyczny wniosek diagnostyczny:** zanim zdiagnozujesz „reguła zniknęła", **wypisz inwentarz** (Przykład 4) i porównaj z tym, co widzisz w Excelu. Jeśli inwentarz jest krótszy niż rzeczywistość — masz do czynienia z regułą rozszerzenia.

**Nurek do pliku.** Jeśli inwentarz nie wystarcza, rozpakuj `.xlsx` (moduł 02) i zajrzyj do `xl/worksheets/sheetN.xml`. Zobaczysz `<conditionalFormatting sqref="...">` z `<cfRule>` w środku. Dla reguł wbudowanych są tam `<colorScale>`, `<dataBar>` albo `<iconSet>` z `<cfvo>`. Dla reguł standardowych — `dxfId` odsyłający do `<dxfs>` w `styles.xml`. **To najlepszy sposób, by zobaczyć prawdę**, gdy Excel i openpyxl się nie zgadzają.

### 2.11. Formatowanie warunkowe a tabele i a `data_only`

Dwie rzeczy, o które pytają wszyscy, bo wydają się groźne, a wcale nie są.

**Formatowanie warunkowe a `data_only` — zero wpływu.** Reguły **nie zmieniają wartości** komórek. To tylko instrukcje rysowania. `data_only=True` (moduł 06) czyta **wartości z pamięci podręcznej** — a reguły nie mają z tym nic wspólnego. Możesz więc:

- wczytać plik z `data_only=True` i mieć normalne wartości,
- wczytać go z `data_only=False` i mieć formuły,
- w obu przypadkach `ws.conditional_formatting` będzie **identyczne**.

Innymi słowy: **`data_only` nie ma nic do formatowania warunkowego.** Reguły są obojętne na to, czy czytasz wartości, czy formuły. To uspokajające i warto to wiedzieć, bo ludzie kombinują z `data_only`, myśląc, że „odblokują" kolory.

**Formatowanie warunkowe a tabele (`Table`).** Tabela (moduł 12) ma **własny styl** (`TableStyleInfo` — paseczki, kolor nagłówka). Ten styl żyje **w tabeli**, a reguły żyją **w arkuszu**. To dwie niezależne warstwy:

- Warstwa tabeli = baza (tło, paski),
- Warstwa reguł = nakłada się **na wierzch** i zwykle wygrywa tam, gdzie ustawia ten sam atrybut.

Praktyczne konsekwencje:

1. **Reguła na zakresie tabeli zadziała** — nie ma konfliktu technicznego. Excel nałoży kolor z reguły na tło tabeli.
2. **Zakres reguły ≠ `ref` tabeli — typowa pułapka.** To ta sama rodzina co w module 12: jeśli tabela rozszerza się o nowe wiersze (bo dopisałeś dane), a `sqref` reguły **nie** został rozszerzony, nowe wiersze nie dostaną formatowania. Musisz trzymać `sqref` i `ref` w zgodzie — albo rozszerzać oba razem.
3. **Tabela z paskami + skala kolorów = wizualny chaos.** Tabela maluje wiersze na zmianę, a skala kolorów maluje gradient. Efekt bywa nieczytelny. Rozważ wyłączenie pasków tabeli (`TableStyleInfo(showRowStripes=False)`), gdy używasz skal kolorów albo pasków danych — inaczej dwie warstwy walczą o to samo tło.

To wszystko sprowadza się do jednej myśli z modułu 12: **arkusz ma kilka warstw opisujących wygląd (styl komórek, styl tabeli, reguły CF), a kolejność ich nakładania trzeba rozumieć.** Nie ma tu magii — jest kolejność.

### 2.12. Skalowanie — koszt to liczba reguł × rozmiar zakresu

Powtórka i rozszerzenie z modułu 13, bo przy regułach wbudowanych koszt jest wyższy.

**Wzór na koszt: `koszt ≈ rozmiar sqref × liczba reguł na tym zakresie`.** Ten koszt płaci **Excel** przy każdym otwarciu pliku i każdej zmianie danych — openpyxl tylko zapisuje tekst.

Reguły wbudowane (skala, pasek, ikony) są **droższe** niż standardowe, bo Excel musi dla każdej komórki **wyliczyć pozycję ikony/paska/gradientu**, a nie tylko nałożyć kolor.

**Konkretny, zły przykład:**

```python
# Nie rob tego na duzym pliku!
ws.conditional_formatting.add("A:A", ColorScaleRule(...))      # cala kolumna = 1 048 576 komorek
ws.conditional_formatting.add("A:A", DataBarRule(...))
ws.conditional_formatting.add("A:A", IconSetRule(...))
ws.conditional_formatting.add("A:A", FormulaRule(...))
# 4 194 304 ocen przy kazdym otwarciu pliku
```

Objaw: użytkownik otwiera raport i czeka. Przewija i „przycina się". Gdy zmieni wartość, Excel na chwilę zamiera.

**Wzorce ograniczające:**

1. **Zakres dokładnie na danych.** `sqref` wyliczaj z `ws.max_row` / realnego rozmiaru danych (moduł 07), a nie z „całej kolumny".
2. **Jedna reguła na kolumnę.** Nie mnóż wariantów na tym samym zakresie.
3. **Skala kolorów zamiast wielu reguł progowych.** „> 100", „> 200", „> 300" jako trzy reguły = trzy oceny. Jedna skala kolorów = jedna ocena i lepszy wizualnie efekt.
4. **Paski/ikony zamiast tysiąca ręcznych kolorów.** To właśnie po to są wbudowane.
5. **Nie formatuj warunkowo danych, których nikt nie ogląda.** 50 000 wierszy w ukrytym arkuszu nie potrzebuje skali kolorów.

**Reguła kciuka:** jeśli `sqref` przekracza realny zakres danych więcej niż dwukrotnie — to sygnał ostrzegawczy. Zobaczysz, jak to wykrywać automatycznie w module 13 (audytor) i rozszerzysz ten pomysł w wyzwaniu tutaj.

### 2.13. Co openpyxl gubi przy zapisie

Trzecia kategoria z modułu 03 w praktyce — i powtórka z modułu 13, z rozszerzeniem o wbudowane.

**Co przechodzi bez zmian:**

- Reguły `cellIs`, `expression`, `top10`, `aboveAverage`, `duplicateValues`, `containsText`, `timePeriod` itd. — tworzone i wczytywane/zapisywane.
- Wbudowane skale kolorów, paski danych i zestawy ikon **w podstawowej specyfikacji** — tworzone i odtwarzane.
- Style różnicowe (`dxf`) w `styles.xml`.

**Co ginie lub jest upraszczane:**

- **Reguły w formacie rozszerzeń (`extLst`)** — niewidoczne, giną.
- **Nowsze warianty pasków danych** (kolor obramowania, kierunek, ustawienia dla wartości ujemnych, gradient) — upraszczane, bo openpyxl obsługuje tylko oryginalną specyfikację.
- **Nowsze zestawy ikon z własnymi progami** w formacie x14 — giną.
- **Mieszane kombinacje** (ikona + wartość w formacie x14) — mogą ginąć.

**Objaw w praktyce:** wczytujesz plik z bogatym formatowaniem z Excela, zapisujesz openpyxl, otwierasz — pasek stracił obramowanie, ikona zmieniła wygląd, a jedna reguła zniknęła bez ostrzeżenia.

**Wniosek (ten sam co w modułach 13 i 18):** **bogatego pliku z Excela nie przepuszczaj bezmyślnie przez „wczytaj-zapisz".** Albo generuj go od zera (wiesz, co w środku), albo odtwórz formatowanie od nowa, albo sprawdź inwentarz reguł przed i po zapisie. Weryfikacja wzrokowa w Excelu nie jest lenistwem — jest częścią procesu.

## 3. Przykłady krok po kroku

### Przykład 1 — skale kolorów 2- i 3-kolorowe (🟢)

Pierwszy przykład pokazuje kompletny cykl i **dowodzi, że wbudowana reguła nie ma `dxf`**.

```python
"""Skale kolorow 2- i 3-kolorowe + dowod, ze wbudowana regula nie ma dxf.

Uruchom:  python examples/14_colorscale.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import ColorScaleRule

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "14_colorscale.xlsx"

DANE = [12, 45, 78, 23, 91, 5, 67, 38, 84, 50]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Wyniki"

    ws.append(["Lp.", "Wynik (2 kolory)", "Ocena (3 kolory)"])
    for indeks, wynik in enumerate(DANE, start=1):
        ws.append([indeks, wynik, wynik])

    # --- skala 2-kolorowa: min -> max ---------------------------------
    # start_color dotyczy NAJMNIEJSZEJ wartosci, end_color NAJWIEKSZEJ
    ws.conditional_formatting.add(
        "B2:B11",
        ColorScaleRule(
            start_type="min", start_color="FF63BE7B",   # zielony (male)
            end_type="max",   end_color="FFF8696B",     # czerwony (duze)
        ),
    )

    # --- skala 3-kolorowa: min -> 50. percentyl -> max ----------------
    ws.conditional_formatting.add(
        "C2:C11",
        ColorScaleRule(
            start_type="min",      start_color="FFF8696B",
            mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
            end_type="max",        end_color="FF63BE7B",
        ),
    )

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Wyniki"]
        print("=" * 84)
        print("DOWOD: WBUDOWANA REGULA NIE MA dxf")
        print("=" * 84)
        for cf in ws.conditional_formatting:
            for regula in cf.rules:
                print(f"  zakres: {cf.sqref}")
                print(f"    type    = {regula.type!r}")       # 'colorScale'
                print(f"    dxf     = {regula.dxf!r}")        # None! (wbudowana)
                print(f"    priority= {regula.priority}")

                # Zamiast dxf - wbudowany obiekt colorScale:
                cs = regula.colorScale
                if cs is not None:
                    kolory = [getattr(c.rgb, "rgb", c.rgb) if hasattr(c, "rgb") else c
                              for c in (cs.color or [])]
                    print(f"    kolory  = {kolory}")
                    progi = [(fo.type, fo.val) for fo in (cs.cfvo or [])]
                    print(f"    progi   = {progi}")
                print()
    finally:
        wb.close()

    print("=" * 84)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 84)
    print("  1. Kolumna B: plynny gradient od zielonego (5) do czerwonego (91).")
    print("  2. Kolumna C: gradient trojstopniowy - czerwony (dol), zolty (srodek),")
    print("     zielony (gora).")
    print("  3. Formatowanie warunkowe -> Zarzadzaj regulami: dwie reguly TYPU")
    print("     'Color Scale', bez stylu (bo to reguły WBUDOWANE).")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Po dodaniu dwóch reguł `ws.conditional_formatting` ma dwie grupy zakresów (`B2:B11` i `C2:C11`), każda z jedną regułą typu `colorScale`. Każda reguła niesie obiekt `ColorScale` z listą progów (`cfvo`) i listą kolorów (`color`). **`dxf` jest `None`** — i to jest poprawne.

**Co trafi do pliku.** Do `sheet1.xml`:

```xml
<conditionalFormatting sqref="B2:B11">
  <cfRule type="colorScale" priority="1">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF63BE7B"/>
      <color rgb="FFF8696B"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>
```

Zauważ: **`<colorScale>` siedzi wprost w regule**, a nie w `styles.xml` — bo to reguła wbudowana. Żadnego `dxfId`. To widać gołym okiem.

**Trzy rzeczy do przemyślenia:**

1. **Kolejność kolorów.** `start_color` = najmniejsza wartość, `end_color` = największa. Zamień je, a mapa cieplna się odwróci. To najczęstsza pomyłka przy skalach.
2. **`mid_type="percentile"`** sprawia, że żółty środek pada na **medianę**, nie na środek zakresu. W naszych danych mediana to 50 (średnia z 45 i 50... a dokładnie wartość środkowa 50). Gdybyś użył `mid_type="percent", mid_value=50`, środek wypadłby dokładnie w połowie drogi między 5 a 91 — czyli na 48. Efekt podobny, ale dla danych z outlierami różnica byłaby dramatyczna (2.3).
3. **`regula.dxf is None`** — jeśli w module 13 przyzwyczaiłeś się czytać kolor z `dxf`, to tutaj ta droga jest ślepa. **Wbudowane reguły czytaj przez `colorScale`/`dataBar`/`iconSet`.**

### Przykład 2 — paski danych z progiem percentylowym i wyróżnienie minimum (🟡)

To ćwiczenie z briefu, doprowadzone do końca. Pokazuje: `DataBarRule`, różnicę `percent` vs `percentile`, oraz **konflikt priorytetów** między regułą wbudowaną a standardową.

```python
"""Paski danych (nasycenie na 90. percentylu) + wyroznienie minimum.

Kluczowe pojecia:
  - percentile vs percent (2.3)
  - priorytet: regula wyroznienia przed paskiem (2.9)

Uruchom:  python examples/14_databar.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "14_databar.xlsx"

# Dane z OUTLIEREM: prezes (200) rozciaga skale, gdyby uzyc min/max
DANE = [
    ("Dział A", 12_000),
    ("Dział B", 8_500),
    ("Dział C", 15_200),
    ("Dział D", 6_800),
    ("Dział E", 11_400),
    ("Dział F", 9_100),
    ("Dział G", 200_000),   # outlier!
    ("Dział H", 7_300),
    ("Dział I", 13_900),
    ("Dział J", 5_400),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaz"

    ws.append(["Dział", "Sprzedaż"])
    for dzial, kwota in DANE:
        ws.append([dzial, kwota])
        ws.cell(row=ws.max_row, column=2).number_format = '#,##0 "zł"'

    ZAKRES = "B2:B11"

    # --- REGULA 1 (dodana pierwsza -> priorytet 1): MINIMUM -----------
    # Wyroznienie najmniejszej wartosci solidnym tlem.
    # UWAGA: dolny priorytet (nizszy numer), zeby kolor byl na wierzchu (2.9)
    czerwone = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    ws.conditional_formatting.add(
        ZAKRES,
        FormulaRule(formula=["B2=MIN($B$2:$B$11)"], fill=czerwone, stopIfTrue=True),
    )

    # --- REGULA 2 (dodana druga -> priorytet 2): PASEK DANYCH ---------
    # start: zero; end: 90. percentyl -> pasek pelny na progu 90. percentyla.
    # DZIEKI percentile outlier (200 tys.) NIE spłaszcza pozostalych paskow.
    ws.conditional_formatting.add(
        ZAKRES,
        DataBarRule(
            start_type="num", start_value=0,
            end_type="percentile", end_value=90,
            color="FF638EC6",
            showValue=True,
        ),
    )

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Sprzedaz"]
        print("=" * 88)
        print("INWENTARZ REGUL (w kolejnosci priorytetow)")
        print("=" * 88)
        for cf in ws.conditional_formatting:
            print(f"  zakres: {cf.sqref}")
            for regula in sorted(cf.rules, key=lambda r: r.priority):
                dodatki = ""
                if regula.dataBar is not None:
                    progi = [(fo.type, fo.val) for fo in regula.dataBar.cfvo]
                    dodatki = f" pasek: progi={progi} kolor={regula.dataBar.color}"
                if regula.dxf is not None and regula.dxf.fill is not None:
                    dodatki = f" tlo={regula.dxf.fill.start_color.rgb}"
                print(f"    priorytet {regula.priority}: {regula.type}{dodatki}")
    finally:
        wb.close()

    print()
    print("=" * 88)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 88)
    print("  1. Dzial J (5 400) ma CZERWONE tlo - to regula wyroznienia minimum.")
    print("  2. Pozostale dzialy maja niebieskie PASKI o dlugosci proporcjonalnej")
    print("     do wartosci - ALE pasek Dzialu G (200 000) NIE jest 30x dluzszy.")
    print("     Bo prog konczy sie na 90. percentylu - outlier nie spłaszcza reszty.")
    print()
    print("  EKSPERYMENT: zmien end_type na 'max' i zapisz ponownie.")
    print("     Zobaczysz, ze wszystkie paski oprocz G zrobia sie malutkie.")
    print("     To jest wlasnie roznica percent vs percentile w praktyce (2.3).")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **`percentile=90` chroni przed outlierami.** Gdybyś użył `end_type="max"`, jedna pensja 200 000 „rozciągnęłaby" skalę i paski wszystkich pozostałych działów skurczyłyby się do pasków-paznokci. To jest różnica `percent`/`percentile` w najczystszej, praktycznej postaci — i to jest powód, dla którego warto ją rozumieć.
2. **Kolejność priorytetów ma tu znaczenie.** Wyróżnienie minimum (reguła 1) leży na wierzchu, więc dziale J widzisz solidny czerwony, a nie niebieski pasek. Tylko dlatego, że dodałem je jako pierwsze. Spróbuj zamienić i sprawdź — to najlepszy sposób, by zapamiętać 2.9.
3. **`DataBarRule` nie przyjmuje `dxf`.** Kolor paska podajesz w parametrze `color`, nie przez styl. Gdybyś próbował dodać `fill=`, dostałbyś `TypeError`.

**Refleksja nad kosztem.** Ta reguła działa na 10 komórkach — koszt zerowy. Gdybyś zastosował ją na `B2:B1000000`, Excel oceniałby pasek dla miliona komórek przy każdym otwarciu. Zakres ma znaczenie (2.12).

### Przykład 3 — zestawy ikon i pułapka „szarych ikon"

```python
"""Zestawy ikon: sygnalizacja swietlna + pulapka showValue/reverse.

Uruchom:  python examples/14_iconset.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import IconSetRule

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "14_iconset.xlsx"

DANE = [("Cel A", 92), ("Cel B", 61), ("Cel C", 34), ("Cel D", 78),
        ("Cel E", 15), ("Cel F", 55), ("Cel G", 88), ("Cel H", 42)]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cele"

    ws.append(["Cel", "Realizacja", "Wersja z liczbą"])
    for nazwa, wartosc in DANE:
        ws.append([nazwa, wartosc, wartosc])

    # --- 3 sygnalizacje (z ramką): progi na 33% i 67% -----------------
    # WAZNE: 3 progi dla 3 ikon. type='percent' -> progi w procentach.
    # showValue=None -> jak w Excelu domyslnie (pokaż liczbe)
    ws.conditional_formatting.add(
        "B2:B9",
        IconSetRule(
            icon_style="3TrafficLights1",
            type="percent",
            values=[0, 33, 67],
            showValue=True,
        ),
    )

    # --- te same dane z showValue=False -> TYLKO ikona -----------------
    ws.conditional_formatting.add(
        "C2:C9",
        IconSetRule(
            icon_style="3TrafficLights1",
            type="percent",
            values=[0, 33, 67],
            showValue=False,      # liczba zniknie, zostanie sama ikona
        ),
    )

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Cele"]
        for cf in ws.conditional_formatting:
            for regula in cf.rules:
                print(f"  zakres: {cf.sqref}")
                print(f"    type       = {regula.type!r}")         # 'iconSet'
                print(f"    dxf        = {regula.dxf!r}")          # None
                icon = regula.iconSet
                print(f"    zestaw     = {icon.iconSet!r}")        # '3TrafficLights1'
                print(f"    showValue  = {icon.showValue!r}")
                print(f"    reverse    = {icon.reverse!r}")
                print(f"    progi      = {[(fo.type, fo.val) for fo in icon.cfvo]}")
                print()
    finally:
        wb.close()

    print("=" * 84)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 84)
    print("  Kolumna B (showValue=True):  zielone/czerwone swiatla + LICZBA obok.")
    print("  Kolumna C (showValue=False): TYLKO swiatla, BEZ liczby.")
    print()
    print("  PULAPKA SZARYCH IKON:")
    print("    Gdybyś uzyl zestawu '3ArrowsGray' z showValue=False, dostaniesz")
    print("    same szare strzalki bez liczby i bez tla - komorka wyglada na")
    print("    wylaczona. Dlatego: NIE lacz '...Gray' z showValue=False.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **`type="percent"` a `values=[0, 33, 67]`.** To mówi Excelowi: „progi są w procentach zakresu". Pierwszy próg `0` oznacza, że **wszystkie** wartości dostaną jakąś ikonę. Gdybyś dał `values=[33, 67, 100]`, wartości poniżej 33% zostałyby **bez ikony** — komórka pusta.
2. **`showValue=False` usuwa liczbę.** W kolumnie C widzisz tylko światła. Czasem to dobrze (czysty „dashboard"), czasem źle (użytkownik nie zna wartości). **Decyzja projektowa, nie techniczna.**
3. **Zestawy `...Gray` są niebezpieczne w połączeniu z `showValue=False`.** Ta kombinacja daje wrażenie „wyłączonych komórek". Jeśli musisz pokazać różnice kolorem, użyj zestawów kolorowych (`3TrafficLights1` zamiast `3TrafficLightsGray`... a dokładniej: zestawów bez `Gray` w nazwie).

**Uwaga o liczbie progów.** Spróbuj tu dać `values=[0, 50]` dla `3TrafficLights1` — Excel może zaproponować naprawę pliku, bo 3 ikony potrzebują 3 progów. To ta pułapka z 2.6 w praktyce.

### Przykład 4 — katalog reguł bez fabryki + inwentarz (wzorzec diagnostyczny)

Ten przykład domyka dwie rzeczy: (a) pokazuje, jak budować `top10`, `aboveAverage`, `duplicateValues`, `containsText`, `timePeriod` **bazowym `Rule`**, oraz (b) daje narzędzie **inwentarza** reguł, gotowe do użycia na dowolnym pliku.

```python
"""Katalog regul bez fabryki (Rule recznie) + inwentarz regul.

Uruchom:  python examples/14_rule_katalog.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import Rule
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "14_rule_katalog.xlsx"


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"

    ws.append(["Zadanie", "Ocena", "Status", "Termin (tekst)"])
    dane = [
        ("Audyt", 95, "pilne", "2026-01-04"),
        ("Raport", 72, "normalne", "2026-02-11"),
        ("Spotkanie", 45, "pilne", "2026-03-20"),
        ("Przegląd", 88, "normalne", "2026-04-15"),
        ("Szkolenie", 61, "normalne", "2026-05-09"),
        ("Inwentaryzacja", 33, "pilne", "2026-06-30"),
    ]
    for wiersz in dane:
        ws.append(wiersz)

    # --- wspolny dxf ------------------------------------------------
    czerwone = DifferentialStyle(
        font=Font(bold=True, color="9C0006"),
        fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    )
    zielone = DifferentialStyle(
        fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
    )

    # --- TOP 3 (typ 'top10') ----------------------------------------
    ws.conditional_formatting.add(
        "B2:B7",
        Rule(type="top10", rank=3, percent=False, bottom=False, dxf=zielone),
    )

    # --- PONIZEJ SREDNIEJ (typ 'aboveAverage' z aboveAverage=False) ---
    ws.conditional_formatting.add(
        "B2:B7",
        Rule(type="aboveAverage", aboveAverage=False, dxf=czerwone),
    )

    # --- DUPLIKATY w kolumnie Status (typ 'duplicateValues') ---------
    ws.conditional_formatting.add(
        "C2:C7",
        Rule(type="duplicateValues", dxf=czerwone),
    )

    # --- ZAWIERA TEKST (typ 'containsText' + formula) ---------------
    regula_tekst = Rule(
        type="containsText", operator="containsText", text="pilne", dxf=czerwone,
    )
    regula_tekst.formula = ['NOT(ISERROR(SEARCH("pilne",C2)))']
    ws.conditional_formatting.add("C2:C7", regula_tekst)

    # --- OKRES CZASU (typ 'timePeriod') -----------------------------
    ws.conditional_formatting.add(
        "D2:D7",
        Rule(type="timePeriod", timePeriod="thisMonth", dxf=zielone),
    )

    wb.save(cel)
    wb.close()


def kolor_z_dxf(dxf) -> str | None:
    """Bezpiecznie wyciaga kolor tla ze stylu roznicowego."""
    if dxf is None or dxf.fill is None:
        return None
    return getattr(dxf.fill.start_color, "rgb", None)


def inwentarz(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        print("=" * 108)
        print(f"INWENTARZ REGUL: {cel.name}")
        print("=" * 108)
        naglowek = (f"{'arkusz':<8} {'zakres':<9} {'typ':<17} {'prio':<5} "
                    f"{'kolor':<10} szczegoly")
        print(naglowek)
        print("-" * 108)
        for ws in wb.worksheets:
            for cf in ws.conditional_formatting:
                for regula in sorted(cf.rules, key=lambda r: r.priority):
                    szczegoly = ""
                    if regula.type == "top10":
                        szczegoly = f"rank={regula.rank} percent={regula.percent} bottom={regula.bottom}"
                    elif regula.type == "aboveAverage":
                        szczegoly = f"aboveAverage={regula.aboveAverage}"
                    elif regula.type == "containsText":
                        szczegoly = f"text={regula.text!r} formula={regula.formula}"
                    elif regula.type == "timePeriod":
                        szczegoly = f"timePeriod={regula.timePeriod!r}"
                    elif regula.type == "colorScale":
                        szczegoly = "WBUDOWANA (brak dxf)"
                    print(f"{ws.title:<8} {str(cf.sqref):<9} {regula.type:<17} "
                          f"{regula.priority:<5} {str(kolor_z_dxf(regula.dxf)):<10} {szczegoly}")
    finally:
        wb.close()


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    inwentarz(CEL)
    print()
    print("Zauważ dwie reguly na zakresie C2:C7 (duplikaty + containsText):")
    print("  obie trafily do JEDNEJ grupy zakresu (moduł 13, 2.11).")
    print("  Ich priorytety: duplikaty=3, containsText=4 (kolejnosc dodawania).")


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **Pięć typów reguł bez ani jednej fabryki.** `top10`, `aboveAverage`, `duplicateValues`, `containsText`, `timePeriod` — wszystkie zbudowane bazowym `Rule`. To jest wiedza, którą w wielu kursach zobaczysz w wersji „import `Top10Rule`" — i która nie zadziała. **Tutaj działa.**
2. **`aboveAverage=False` zamiast osobnego typu.** Zauważ, jak `aboveAverage=False` realizuje „poniżej średniej". Nie ma i nie potrzebuje być osobnej klasy — to jeden typ z flagą.
3. **Inwentarz to Twoje narzędzie diagnostyczne.** Wypisuje typ, priorytet, kolor (z `dxf` albo `None` dla wbudowanych) i szczegóły specyficzne dla typu. Uruchom go na dowolnym pliku od klienta — zobaczysz, co openpyxl widzi (2.10). Jeśli inwentarz jest krótszy niż to, co widzisz w Excelu, masz reguły rozszerzeń, które giną.

## 4. Anatomia API

| Metoda / klasa / atrybut | Co robi | Parametry / wartości | Uwagi |
|---|---|---|---|
| `FormatObject(type, val, gte)` | Pojedynczy próg (`cfvo`) | `type`: `min`/`max`/`num`/`percent`/`percentile`/`formula`; `val`: liczba lub adres | `val` liczbowy dla `num`/`percent`/`percentile`; `None` dla `min`/`max` |
| `ColorScaleRule(...)` | Skala kolorów (2 lub 3 kolory) | `start_type`, `start_value`, `start_color`, `mid_*`, `end_*` | **Funkcja fabryczna.** Wymaga `start_type` i `end_type`; brak `dxf`/`fill`/`stopIfTrue` |
| `DataBarRule(...)` | Pasek danych | `start_type`, `start_value`, `end_type`, `end_value`, `color`, `showValue`, `minLength`, `maxLength` | **Fabryka.** Tylko 2 progi; brak `dxf`/`stopIfTrue` |
| `IconSetRule(...)` | Zestaw ikon | `icon_style`, `type`, `values`, `showValue`, `percent`, `reverse` | **Fabryka.** `icon_style` = nazwa ikon, `type` = typ progów; `len(values)` = liczba ikon |
| `Rule(type, ...)` | Reguła dowolnego typu | `type` **wymagany**; `rank`, `percent`, `bottom`, `aboveAverage`, `stdDev`, `equalAverage`, `operator`, `text`, `timePeriod`, `formula`, `dxf`, `priority`, `stopIfTrue`, ... | **Jedyna droga dla `top10`/`aboveAverage`/`duplicateValues`/`containsText`/`timePeriod`** |
| `Rule.priority` | Priorytet (mniejszy = ważniejszy) | `int` | Globalny dla arkusza; ustawiaj po utworzeniu dla fabryk |
| `Rule.dxf` | Styl różnicowy | `DifferentialStyle` | Dla wbudowanych zawsze `None`; po wczytaniu rozwiązywane z `dxfId` |
| `Rule.colorScale` | Obiekt skali kolorów | `ColorScale` | Tylko dla `type="colorScale"` |
| `Rule.dataBar` | Obiekt paska | `DataBar` | Tylko dla `type="dataBar"` |
| `Rule.iconSet` | Obiekt ikon | `IconSet` | Tylko dla `type="iconSet"` |
| `DifferentialStyle(...)` | Styl różnicowy | `font`, `numFmt`, `fill`, `alignment`, `border`, `protection` | **Sześć składników** (jak styl komórki); **brak `name`** |
| `ColorScale.color` | Lista kolorów skali | `list[Color]` | Do odczytu kolorów skali |
| `ColorScale.cfvo` | Lista progów skali | `list[FormatObject]` | `.type`, `.val` |
| `DataBar.color` | Kolor paska | `Color` | Czytany/ustawiany jako napis albo `Color` |
| `DataBar.cfvo` | Progi paska | `list[FormatObject]` | Zwykle 2 elementy |
| `DataBar.minLength` / `.maxLength` | Zakres długości paska | `int` (0–100, procent) | Nie wszystkie wersje Excela honorują identycznie |
| `IconSet.iconSet` | Nazwa zestawu ikon | np. `"3TrafficLights1"` | Z listy 17 zestawów |
| `IconSet.cfvo` | Progi ikon | `list[FormatObject]` | `len` = liczba ikon |
| `IconSet.showValue` / `.reverse` | Flagi ikon | `bool` | `showValue=False` usuwa liczbę |
| `DifferentialStyleList.append(dxf)` | Dodaje styl różnicowy | — | **Deduplikuje** identyczne style (współdzielenie jak w module 09) |
| `CellRange.isdisjoint(other)` | Czy zakresy rozłączne | — | Z modułu 12/13 — do wykrywania nakładania |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — trójkolorowa skala na kolumnie wyników.**

Napisz `examples/14_cw1.py`, który tworzy `output/14_cw1.xlsx`:

1. Arkusz `Oceny` z nagłówkiem `Student`, `Punkty` oraz 12 wierszami danych (punkty 0–100, `seed` ustalony).
2. **Trójkolorową skalę** na kolumnie `Punkty`, zakres dokładnie `B2:B13`:
   - start `min` → czerwony (`FFF8696B`),
   - środek `percent` na `50` → żółty (`FFFFEB84`),
   - koniec `max` → zielony (`FF63BE7B`).
3. Funkcję `weryfikuj()`, która wczytuje plik i wypisuje dla reguły: `type`, `dxf` (oczekujesz `None`), kolory i progi z `regula.colorScale`.

W pliku `output/14_cw1_wnioski.md` odpowiedz:

1. **Dlaczego `regula.dxf` jest `None`, a `regula.colorScale` nie jest puste?** Odnieś się do różnicy między regułą wbudowaną a standardową (2.1).
2. **Co dokładnie znaczą progi `min`, `percent=50`, `max`?** Policz, jaka wartość „wypada" na środek, gdyby dane miały min=10, max=510 (2.3).
3. **Jaki byłby efekt `mid_type="percentile", mid_value=50`?** Czym różni się od `percent=50` (2.3)?
4. **Co się stanie, gdy zamienisz `start_color` z `end_color`?** Dlaczego łatwo o tę pomyłkę (2.4)?
5. **Dlaczego w komórce nagłówka `B1` nie ma koloru?** (moduł 13 — zakres reguły)

### 🟡 Warsztat

**Zadanie 2 — paski danych z progiem 80. percentyla i wyróżnieniem minimum.**

To zadanie z briefu. Napisz `examples/14_cw2.py`:

1. Arkusz `Wyniki` z kolumną `Wartość` (15–20 wierszy, w tym **jeden duży outlier** — np. 10× większy od reszty).
2. **Pasek danych** na kolumnie `Wartość`:
   - `start_type="num", start_value=0`,
   - `end_type="percentile", end_value=80` (nasycenie na 80. percentylu),
   - kolor `FF638EC6`, `showValue=True`.
3. **Wyróżnienie minimum** — regułę (np. `FormulaRule`), która maluje na czerwono komórkę z najmniejszą wartością. **Dodaj ją jako PIERWSZĄ**, żeby miała priorytet 1.
4. Funkcję `weryfikuj()`, która wypisuje priorytety obu reguł i ich typy.

W pliku `output/14_cw2_wnioski.md`:

1. **Dlaczego przy outlierze `percentile=80` jest lepsze od `max`?** Zbadaj eksperymentalnie: zmień `end_type` na `max`, zapisz i porównaj długości pasków.
2. **Dlaczego wyróżnienie minimum dodałeś jako pierwsze?** Co się stanie, gdy dodasz je jako drugie (z niższym priorytetem) (2.9)?
3. **Dlaczego `DataBarRule` nie przyjmuje `dxf`?** Odnieś się do 2.1.
4. **Jak sprawdziłeś, że na komórce z minimum działa WYRÓŻNIENIE, a nie pasek?** Opisz obserwację w Excelu.
5. **Czy zmiana `B2=MIN($B$2:$B$16)` na `$B2=MIN($B$2:$B$16)` coś zmienia?** Uzasadnij w kontekście modułu 13 (2.7 tamtego modułu — kalka).

### 🔴 Wyzwanie

**Zadanie 3 — generator reguł z deklaratywnej konfiguracji (zapowiedź wzorca Specification).**

To zadanie z briefu i zarazem bezpośrednia zapowiedź wzorca **Specification** (moduł 25). Zbudujesz mechanizm, który z listy słowników robi reguły.

```python
"""Generator regul formatowania warunkowego z konfiguracji.

Szkic API:
    KONFIG = [
        {"zakres": "B2:B100", "typ": "colorscale",
         "start_type": "min", "start_color": "FFF8696B",
         "end_type": "max", "end_color": "FF63BE7B"},
        {"zakres": "C2:C100", "typ": "iconset",
         "icon_style": "5Arrows", "type": "percent",
         "values": [0, 20, 40, 60, 80], "showValue": True},
        {"zakres": "D2:D100", "typ": "databar",
         "start_type": "num", "start_value": 0,
         "end_type": "percentile", "end_value": 90, "color": "FF638EC6"},
        {"zakres": "E2:E100", "typ": "top10",
         "rank": 5, "percent": True, "kolor_tla": "FFC7CE"},
        {"zakres": "F2:F100", "typ": "formula",
         "formula": ["$F2>TODAY()"], "kolor_tla": "FFFF00", "stopIfTrue": True},
    ]

    generator = GeneratorRegul()
    generator.zastosuj(ws, KONFIG)
"""
```

**Wymagania:**

1. **Obsłuż co najmniej pięć typów** z konfiguracji: `colorscale`, `iconset`, `databar`, `top10`, `formula`.
2. **Kolejność w konfiguracji = kolejność priorytetów.** Reguła zdefiniowana wcześniej na tym samym zakresie ma niższy priorytet. To musi być jawnie udokumentowane i przetestowane.
3. **Walidacja konfiguracji.** Odrzuć wpisy z nieznanym `typ`, brakującym `zakres` albo — dla `iconset` — gdy `len(values)` nie zgadza się z liczbą ikon zestawu (mapa z 2.6). Rzuć czytelny wyjątek `NieprawidlowaKonfiguracja`.
4. **Budowa `dxf` z koloru.** Dla typów standardowych ze słownika `kolor_tla` zbuduj `DifferentialStyle` z jednolitym tłem solid.
5. **Test równoważności.** Napisz funkcję, która po zapisaniu i wczytaniu sprawdza, że **liczba reguł i ich typy** odpowiadają konfiguracji (policz reguły przez zagnieżdżoną iterację — moduł 13, 2.11!).

W pliku `output/14_cw3_wnioski.md`:

1. **Jaka korzyść płynie z opisania reguł w konfiguracji, a nie w kodzie?** Podaj dwa konkretne scenariusze, w których to wygrywa (np. nowy raport bez zmiany kodu).
2. **Gdzie jest granica?** Kiedy konfiguracja staje się „nowym językiem programowania" i przestaje się opłacać? (Zapowiedź 2.12 modułu 26.)
3. **Jak zapewniłeś, że kolejność priorytetów jest deterministyczna?** Co by się stało, gdybyś mieszał `add()` z ręcznym ustawianiem `priority` (moduł 13, 2.5)?
4. **To realizuje pewien wzorzec.** Który wzorzec z modułu 25 opisuje budowanie obiektów z deklaratywnej specyfikacji? Uzasadnij w dwóch zdaniach.
5. **Jak walidujesz liczbę progów dla zestawu ikon?** Skąd wziąłeś mapę „zestaw → liczba ikon"? (2.6)

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Generator regul z konfiguracji - kluczowe fragmenty."""

from __future__ import annotations

from pathlib import Path

from openpyxl.formatting.rule import (
    ColorScaleRule, DataBarRule, IconSetRule, FormulaRule, Rule,
)
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle


class NieprawidlowaKonfiguracja(Exception):
    """Konfiguracja jest niekompletna albo sprzeczna."""


# Mapa: zestaw ikon -> liczba wymaganych progow (2.6)
LICZBA_IKON = {
    "3Arrows": 3, "3ArrowsGray": 3, "3Flags": 3,
    "3TrafficLights1": 3, "3TrafficLights2": 3,
    "3Signs": 3, "3Symbols": 3, "3Symbols2": 3,
    "4Arrows": 4, "4ArrowsGray": 4, "4RedToBlack": 4,
    "4Rating": 4, "4TrafficLights": 4,
    "5Arrows": 5, "5ArrowsGray": 5, "5Rating": 5, "5Quarters": 5,
}

TYPY_STANDARDOWE = {"top10", "formula", "cellis", "duplicatevalues", "aboveaverage"}


class GeneratorRegul:
    """Tlumaczy deklaratywna konfiguracje na reguly openpyxl."""

    def __init__(self) -> None:
        self._bufory_dxf: dict[str, DifferentialStyle] = {}

    def _dxf_z_koloru(self, kolor: str) -> DifferentialStyle:
        """Buduje (i buforuje) styl roznicowy z jednolitym tlem solid."""
        if kolor not in self._bufory_dxf:
            self._bufory_dxf[kolor] = DifferentialStyle(
                fill=PatternFill(start_color=kolor, end_color=kolor, fill_type="solid"),
            )
        return self._bufory_dxf[kolor]

    def _zbuduj(self, wpis: dict):
        """Buduje obiekt Rule z jednego wpisu konfiguracji."""
        typ = wpis.get("typ")
        if typ is None:
            raise NieprawidlowaKonfiguracja("wpis bez klucza 'typ'")

        if typ == "colorscale":
            return ColorScaleRule(
                start_type=wpis["start_type"], start_color=wpis["start_color"],
                mid_type=wpis.get("mid_type"), mid_value=wpis.get("mid_value"),
                mid_color=wpis.get("mid_color"),
                end_type=wpis["end_type"], end_color=wpis["end_color"],
            )
        if typ == "databar":
            return DataBarRule(
                start_type=wpis["start_type"], start_value=wpis["start_value"],
                end_type=wpis["end_type"], end_value=wpis["end_value"],
                color=wpis.get("color", "FF638EC6"),
                showValue=wpis.get("showValue", True),
            )
        if typ == "iconset":
            zestaw = wpis["icon_style"]
            wartosci = wpis.get("values", [])
            if zestaw not in LICZBA_IKON:
                raise NieprawidlowaKonfiguracja(f"nieznany zestaw ikon: {zestaw!r}")
            if len(wartosci) != LICZBA_IKON[zestaw]:
                raise NieprawidlowaKonfiguracja(
                    f"zestaw {zestaw!r} wymaga {LICZBA_IKON[zestaw]} progow, "
                    f"a podano {len(wartosci)}",
                )
            return IconSetRule(
                icon_style=zestaw, type=wpis.get("type", "percent"),
                values=wartosci, showValue=wpis.get("showValue", True),
                reverse=wpis.get("reverse"),
            )
        if typ == "formula":
            return FormulaRule(
                formula=wpis["formula"],
                fill=self._dxf_z_koloru(wpis["kolor_tla"]).fill,
                stopIfTrue=wpis.get("stopIfTrue"),
            )
        if typ == "top10":
            return Rule(
                type="top10", rank=wpis["rank"],
                percent=wpis.get("percent", False),
                bottom=wpis.get("bottom", False),
                dxf=self._dxf_z_koloru(wpis["kolor_tla"]),
            )
        raise NieprawidlowaKonfiguracja(f"nieznany typ reguly: {typ!r}")

    def zastosuj(self, ws, konfig: list[dict]) -> None:
        """Stosuje cala konfiguracje do arkusza.

        Kolejnosc wpisow = kolejnosc priorytetow (pierwszy wpis = priorytet 1).
        """
        for indeks, wpis in enumerate(konfig, start=1):
            zakres = wpis.get("zakres")
            if not zakres:
                raise NieprawidlowaKonfiguracja(f"wpis {indeks} bez 'zakres'")
            regula = self._zbuduj(wpis)
            ws.conditional_formatting.add(zakres, regula)
```

**Kluczowe decyzje i uzasadnienia:**

- **Bufor `dxf` (`_bufory_dxf`).** Identyczny kolor tła jest używany wielokrotnie — buforujemy `DifferentialStyle`, żeby nie tworzyć wielu obiektów. To wzorzec **Flyweight / Registry** (moduł 23) i dokładnie ta optymalizacja, o której mówiłem w module 09. Zauważ, że openpyxl i tak deduplikowałby identyczne dxf-y na zapisie (`DifferentialStyleList.append`), ale bufor oszczędza alokacje po stronie Pythona.

- **Walidacja liczby progów dla ikon.** Mapa `LICZBA_IKON` pochodzi wprost z 2.6. Bez niej użytkownik konfiguracji mógłby podać 2 progi dla 3 ikon i dostać „uszkodzony" plik w Excelu. **Dobra konfiguracja wyłapuje błędy wcześnie** — to sedno podejścia „fail fast".

- **Kolejność deterministyczna.** Ponieważ `add()` nadaje priorytety w kolejności dodawania (moduł 13, 2.5), iteracja po liście konfiguracji **zapewnia**, że „wpis wyżej w pliku konfiguracji = wyższy priorytet". To musi być udokumentowane, bo to kontrakt tej abstrakcji. **Nigdy nie mieszaj automatycznych priorytetów z ręcznym `priority` w tym samym mechanizmie** — inaczej tracisz determinizm.

- **`FormulaRule(fill=...)` a `DifferentialStyle`.** `FormulaRule` przyjmuje `fill` bezpośrednio, a buduje `dxf` wewnętrznie. Ale w konfiguracji chcę też używać `Rule(type="top10", dxf=...)` — tam muszę przekazać cały `DifferentialStyle`. Dlatego w `_dxf_z_koloru` buduję `DifferentialStyle`, a do `FormulaRule` przekazuję z niego tylko `.fill`. To pokazuje, że **obie drogi są równoważne** — fabryka opakowuje to, co w `Rule` podajesz wprost.

- **Czego generator NIE robi.** Nie próbuje „zgadnąć" intencji. Nie waliduje, czy formuła jest sensowna (bo openpyxl nie ma silnika oceny — moduł 06), nie sprawdza, czy progi są rosnące, nie wykrywa nakładania zakresów. **Świadomie ogranicza się do tłumaczenia konfiguracji na obiekty i walidacji strukturalnej.** Rozszerzenie o wykrywanie nakładania (`CellRange.isdisjoint`, moduł 12/13) zostawiam jako naturalny next step — i to jest dobry przykład, że nawet „prosty" generator ma swoją granicę odpowiedzialności.

- **To jest zapowiedź wzorca Specification.** Konfiguracja opisuje **co** (jakie reguły), a `GeneratorRegul` wie **jak** je zbudować. W module 25 zobaczysz ten sam pomysł w wersji obiektowej: `Specyfikacja` z metodą `jest_spelniona(wiersz)` i kompozycją `and_/or_/not_`. Ten sam duch — **oddziel regułę od jej wykonania** — tu w wersji deklaratywnej, tam w wersji obiektowej.

</details>

## 6. Typowe błędy i pułapki

**1. „`ImportError: cannot import name 'Top10Rule'`" (objaw) → **openpyxl 3.1.x NIE ma klas `Top10Rule`, `AboveAverageRule`, `DuplicateValuesRule`, `TextRule`, `DateRule`, `TimePeriodRule`.** Istnieje tylko pięć fabryk: `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule`, `IconSetRule` (2.1) (przyczyna) → **użyj bazowego `Rule` z odpowiednim `type`**: `Rule(type="top10", rank=5, dxf=...)`, `Rule(type="aboveAverage", aboveAverage=False, dxf=...)`, `Rule(type="containsText", operator="containsText", text=..., formula=[...], dxf=...)`. Pełny katalog w 2.7 (naprawa).**

**2. „`TypeError: ColorScaleRule() got an unexpected keyword argument 'fill'`" (objaw) → **reguły wbudowane (`colorScale`, `dataBar`, `iconSet`) nie przyjmują `font`/`fill`/`border`/`dxf`.** Nie mają stylu różnicowego, bo kolor jest częścią samej wizualizacji (2.1) (przyczyna) → **kolor podajesz wprost**: `ColorScaleRule(..., start_color=..., end_color=...)`, `DataBarRule(..., color=...)`. **To samo dotyczy `stopIfTrue` — reguły wbudowane go nie przyjmują** (naprawa).**

**3. „Pomyliłem `percent` z `percentile` i progi wypadły w złym miejscu" (objaw) → `percent` to **procent zakresu (od min do max)**, a `percentile` to **rank ing** — pozycja w posortowanych danych. Dla danych z outlierem różnica jest dramatyczna (2.3) (przyczyna) → **użyj `percentile` w danych z outlierami** (np. `end_type="percentile", end_value=90`, żeby pasek nie spłaszczał się przez jedną wielką wartość), a `percent` gdy chcesz progu względem szerokości przedziału (naprawa).**

**4. „`ValidationError` albo Excel proponuje naprawę pliku przy zestawie ikon" (objaw) → **liczba progów w `values` nie zgadza się z liczbą ikon zestawu.** `3TrafficLights1` potrzebuje 3 progów, `5Arrows` — 5 (2.6) (przyczyna) → **zapewnij `len(values) == liczba ikon`**. Mapa „zestaw → liczba ikon" jest w 2.6. Najczęstszy błąd: podanie 2 progów dla 3 ikon (naprawa).**

**5. „Komórki wyglądają na wyłączone — same szare strzałki bez tła i bez liczby" (objaw) → **połączyłeś zestaw `...Gray` (szare ikony) z `showValue=False` (bez liczby).** Ikona jest szara, liczby nie ma, tła zestaw ikon nie maluje — komórka wygląda martwo (2.6) (przyczyna) → **nie łącz zestawów `...Gray` z `showValue=False`.** Albo użyj kolorowego zestawu, albo pokaż liczbę (`showValue=True`), albo dodaj drugą regułę malującą tło (naprawa).**

**6. „Wyróżnienie skrajnych nie działa — wszystko jest w gradiencie" (objaw) → **reguła wyróżnienia ma wyższy numer priorytetu niż skala kolorów,** więc gradient leży na wierzchu i przesłania solidny kolor (2.9) (przyczyna) → **dodaj regułę wyróżnienia PIERWSZĄ** (dostanie priorytet 1), a skalę kolorów później. Albo ustaw `priority` ręcznie po utworzeniu. To samo dotyczy dowolnego konfliktu „szczegółowa vs ogólna reguła" (naprawa).**

**7. „`KeyError` / `ValidationError` przy `DifferentialStyle(name='...')`" (objaw) → **`DifferentialStyle` NIE ma parametru `name`.** Ma sześć składników: `font`, `numFmt`, `fill`, `alignment`, `border`, `protection` (2.8). Mylisz go z `NamedStyle` z modułu 09 (przyczyna) → **usuń `name`** i podaj tylko potrzebne składniki: `DifferentialStyle(font=..., fill=..., border=...)`. Jeśli chcesz użyć `NamedStyle` w regule — **nie da się**; zbuduj osobny `DifferentialStyle` z tych samych składników (naprawa).**

**8. „Chcę użyć mojego `NamedStyle` z modułu 09 w regule i dostaję błąd" (objaw) → **reguła przyjmuje tylko `DifferentialStyle`, a `NamedStyle` to inny byt** (2.8). Są przechowywane w różnych miejscach pliku (przyczyna) → **przekonwertuj ręcznie**: zbuduj `DifferentialStyle` z tych samych `Font`/`PatternFill`/`Border`, które miałeś w stylu nazwanym. Nie istnieje automatyczna konwersja (naprawa).**

**9. „Reguły z pliku utworzonego w Excelu nie zgadzają się z tym, co widzę — część zniknęła" (objaw) → **reguły w formacie rozszerzeń (`extLst`) są niewidoczne dla openpyxl i giną przy zapisie.** Dotyczy to nowszych wariantów pasków danych, zestawów ikon i kombinacji (2.10, 2.13) (przyczyna) → (a) **wypisz inwentarz reguł** (Przykład 4) przed i po i porównaj liczby, (b) **nie przepuszczaj bogatego pliku z Excela przez „wczytaj-zapisz"** — generuj nowy albo odtwórz formatowanie od zera, (c) zajrzyj do `sheetN.xml` (moduł 02), żeby zobaczyć, co tam naprawdę jest (naprawa).**

**10. „`regula.dxf.fill` sypie `AttributeError: 'NoneType'`" (objaw) → **dla reguł wbudowanych `regula.dxf` jest `None`** — bo `colorScale`/`dataBar`/`iconSet` nie używają stylu różnicowego (2.1). Próbujesz czytać `dxf` na regule wbudowanej (przyczyna) → **sprawdź `regula.type`**: jeśli to `colorScale`/`dataBar`/`iconSet`, czytaj `regula.colorScale`/`regula.dataBar`/`regula.iconSet`. Jeśli standardowa — czytaj `regula.dxf` (naprawa).**

**11. „Plik jest wolny przy otwieraniu, choć reguł jest tylko kilka" (objaw) → **zakres reguł jest ogromny.** `sqref="A:A"` to 1 048 576 komórek; przy regule wbudowanej Excel musi dla każdej wyliczyć gradient/ikonę/pasek (2.12) (przyczyna) → **ogranicz `sqref` do realnego zakresu danych** (wylicz z `ws.max_row`, moduł 07). To ta sama zasada co `print_area` (moduł 11) i `ref` tabeli (moduł 12). Pamiętaj: reguły wbudowane **kosztują więcej** niż standardowe (naprawa).**

**12. „`DataBarRule` z obsługą wartości ujemnych nie działa tak, jak chcę" (objaw) → **openpyxl obsługuje tylko oryginalną specyfikację pasków; nowsze warianty (obramowanie, kierunek, kolor osi, ustawienia dla wartości ujemnych) wymagają rozszerzeń, których nie ma** (2.5) (przyczyna) → **zaakceptuj domyślne zachowanie** (Excel narysuje pasek od osi, ale bez kontroli koloru) albo rozważ inny generator (`xlsxwriter` obsługuje nowsze paski). Nazwij to ograniczenie świadomie w dokumentacji raportu (naprawa).**

**13. „Zapisałem plik z `top10` i Excel pokazuje regułę jako „uszkodzoną" / dziwną" (objaw) → **brak wymaganych atrybutów dla typu reguły.** Dla `top10` potrzebne `rank`; dla `aboveAverage` potrzebne `aboveAverage`; dla `containsText` — `text` i zwykle `formula` (2.7) (przyczyna) → **sprawdź w 2.7, których atrybutów wymaga dany typ**, i ustaw je wszystkie. openpyxl **nie waliduje spójności** (moduł 03: to samo co w formułach), więc zapisze niekompletną regułę bez błędu (naprawa).**

**14. „Nowe wiersze w tabeli nie dostają formatowania warunkowego" (objaw) → **`sqref` reguły i `ref` tabeli się rozjechały** — dopisałeś wiersze, rozszerzyłeś tabelę, ale nie regułę (2.11) (przyczyna) → **rozszerzaj `sqref` i `ref` razem** przy każdym dopisaniu danych. To ta sama dyscyplina co w module 12 przy rozszerzaniu tabel. Dobra praktyka: licz `sqref` z realnego rozmiaru danych, nigdy nie „zapamiętuj" go na sztywno (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Dwa światy reguł.** *Standardowe* (cellIs, expression, top10, aboveAverage, duplicateValues, containsText, timePeriod...) **malują** komórkę przez `dxf`. *Wbudowane* (colorScale, dataBar, iconSet) **rysują w komórce mikro-wykres** i **nie mają `dxf`**. Dlatego wbudowane nie przyjmują `font`/`fill`/`border`/`stopIfTrue`, a po wczytaniu czytasz je przez `regula.colorScale`/`dataBar`/`iconSet`, nie przez `regula.dxf`.

2. **Pięć fabryk, osiemnaście typów.** openpyxl 3.1.x ma dokładnie pięć funkcji fabrycznych: `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule`, `IconSetRule`. Wszystko inne — `top10`, `aboveAverage`, `duplicateValues`, `containsText`, `timePeriod`, `containsBlanks`, `containsErrors` — budujesz bazowym `Rule` z odpowiednim `type` i atrybutami. **Nie istnieje `Top10Rule` ani `DateRule`.**

3. **Progi to serce wizualizacji — i `percent` ≠ `percentile`.** Sześć typów: `min`, `max`, `num`, `percent`, `percentile`, `formula`. `percent` = procent **zakresu** (szerokość pudełka); `percentile` = **ranking** (kolejność w kolejce). Przy outlierach `percentile` ratuje czytelność pasków. Liczba progów musi się zgadzać z liczbą kolorów (skala) i ikon (zestaw).

4. **Nakładanie rozstrzyga priorytet — mniejszy numer wygrywa.** Warstwy farby: reguła o niższym numerze priorytetu leży na wierzchu. **Reguła szczegółowa (wyróżnienie skrajnych) przed ogólną (skala kolorów).** `add()` nadaje priorytety w kolejności dodawania; fabryki nie przyjmują `priority`. Reguły wbudowane nie mają `stopIfTrue` — projektuj zakresy i kolejność świadomie.

5. **openpyxl modeluje podstawową specyfikację; rozszerzenia giną.** Klasyczne reguły i standardowe warianty wbudowane przechodzą. Ale reguły w `extLst`, nowsze paski danych i bogate kombinacje ikon — **giną lub są upraszczane przy zapisie**. Bogatego pliku z Excela nie przepuszczaj bezmyślnie przez „wczytaj-zapisz"; generuj nowy, odtwarzaj formatowanie, i **zawsze sprawdź inwentarz reguł** przed i po (i najlepiej okiem w Excelu).

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORT - PIEC FABRYK (i nic wiecej!)
# ==================================================================
from openpyxl.formatting.rule import (
    CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule,
    Rule, FormatObject,
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles import Font, PatternFill, Border, Side

# ⚠️ NIE MA: Top10Rule, AboveAverageRule, DuplicateValuesRule, TextRule, DateRule

# ==================================================================
# 2. SKALA KOLOROW (ColorScaleRule) - 2 lub 3 kolory
# ==================================================================
# KLUCZ: start_color = NAJMNIEJSZA wartosc, end_color = NAJWIEKSZA
ws.conditional_formatting.add(
    "B2:B100",
    ColorScaleRule(start_type="min", start_color="FF63BE7B",
                   end_type="max", end_color="FFF8696B"),
)
# Skala 3-kolorowa: dodaj mid_type, mid_value, mid_color
ws.conditional_formatting.add(
    "C2:C100",
    ColorScaleRule(start_type="min",      start_color="FFF8696B",
                   mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
                   end_type="max",        end_color="FF63BE7B"),
)
# ⚠️ Wymaga start_type i end_type. Brak fill/dxf/stopIfTrue!
# ⚠️ Kolory: 'RRGGBB' albo 'AARRGGBB' (najbezpieczniej 8 znaków)

# ==================================================================
# 3. PASEK DANYCH (DataBarRule) - tylko 2 progi
# ==================================================================
ws.conditional_formatting.add(
    "B2:B100",
    DataBarRule(start_type="num", start_value=0,
                end_type="percentile", end_value=90,   # nasycenie na percentylu
                color="FF638EC6", showValue=True),
)
# ⚠️ Brak dxf, brak stopIfTrue. minLength/maxLength w procentach (0-100).
# ⚠️ showValue: True/False (NIE string "None" - to blad w dokumentacji!)
# ⚠️ Ograniczenie: tylko oryginalna specyfikacja (brak kontroli nad wartosciami
#    ujemnymi, gradientem, obramowaniem).

# ==================================================================
# 4. ZESTAW IKON (IconSetRule) - len(values) == liczba ikon
# ==================================================================
ws.conditional_formatting.add(
    "C2:C100",
    IconSetRule(icon_style="3TrafficLights1",   # IMPORTANT: to NAZWA IKON
                type="percent",                  # to TYP PROGOW (nie nazwa!)
                values=[0, 33, 67],              # 3 progi dla 3 ikon
                showValue=True, reverse=None),
)
# 17 zestawow: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2,
#   3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4RedToBlack, 4Rating,
#   4TrafficLights, 5Arrows, 5ArrowsGray, 5Rating, 5Quarters
# ⚠️ NIE lacz zestawow '...Gray' z showValue=False (wygladaja na wylaczone)
# ⚠️ Ikony NIE maluja tla - chcesz tlo? dodaj druga regule z fill

# ==================================================================
# 5. REGULY BEZ FABRYKI - bazowy Rule + type
# ==================================================================
dxf = DifferentialStyle(
    font=Font(bold=True, color="9C0006"),
    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
)

# TOP N (top10)
ws.conditional_formatting.add("B2:B100",
    Rule(type="top10", rank=5, percent=False, bottom=False, dxf=dxf))

# PONIZEJ/POWYZEJ SREDNIEJ (aboveAverage - JEDEN typ, nie dwa!)
ws.conditional_formatting.add("B2:B100",
    Rule(type="aboveAverage", aboveAverage=False, dxf=dxf))   # False = ponizej

# DUPLIKATY / UNIKATY
ws.conditional_formatting.add("A2:A100", Rule(type="duplicateValues", dxf=dxf))
ws.conditional_formatting.add("A2:A100", Rule(type="uniqueValues", dxf=dxf))

# ZAWIERA TEKST (containsText) - potrzebna formula
r = Rule(type="containsText", operator="containsText", text="pilne", dxf=dxf)
r.formula = ['NOT(ISERROR(SEARCH("pilne",A2)))']
ws.conditional_formatting.add("A2:F100", r)

# OKRES CZASU (timePeriod): today, yesterday, tomorrow, last7Days, thisMonth,
#   lastMonth, nextMonth, thisWeek, lastWeek, nextWeek (NIE ma thisYear!)
ws.conditional_formatting.add("C2:C100",
    Rule(type="timePeriod", timePeriod="thisMonth", dxf=dxf))

# PUSTE / BLEDY
r = Rule(type="containsBlanks", dxf=dxf); r.formula = ['LEN(TRIM(A2))=0']
r = Rule(type="containsErrors", dxf=dxf); r.formula = ['ISERROR(A2)']

# ==================================================================
# 6. PROGI RECZNIE (FormatObject) - gdy fabryka nie wystarcza
# ==================================================================
from openpyxl.formatting.rule import ColorScale, FormatObject
from openpyxl.styles import Color
progi = [FormatObject("min"), FormatObject("num", 50), FormatObject("max")]
kolory = [Color("FFF8696B"), Color("FFFFEB84"), Color("FF63BE7B")]
cs = ColorScale(cfvo=progi, color=kolory)
rule = Rule(type="colorScale", colorScale=cs)
# ⚠️ FormatObject: typy 'min','max','num','percent','percentile','formula'
# ⚠️ val: liczba dla num/percent/percentile; napis (adres) dla formula; None dla min/max

# ==================================================================
# 7. PRIORYTETY - kolejnosc = wazne (2.9)
# ==================================================================
# Regula z nizszym numerem priorytetu lezy NA WIERZCHU (wygrywa konflikt).
# Kolejnosc dodawania = kolejnosc priorytetow.
# SCENARIUSZ: wyroznienie skrajnych + skala kolorow:
ws.conditional_formatting.add("B2:B100",    # <- PIERWSZA = priorytet 1 (na wierzchu)
    FormulaRule(formula=["$B2<60"], fill=czerwone, stopIfTrue=True))
ws.conditional_formatting.add("B2:B100",    # <- DRUGA = priorytet 2 (pod spodem)
    ColorScaleRule(start_type="min", start_color="FFF8696B",
                   end_type="max", end_color="FF63BE7B"))

# ==================================================================
# 8. ODCZYT REGUL PO WCZYTANIU (2.10)
# ==================================================================
for cf in ws.conditional_formatting:
    for regula in cf.rules:
        print(regula.type, regula.priority)
        if regula.dxf is not None:                 # standardowe: styl roznicowy
            print("  kolor tla:", regula.dxf.fill.start_color.rgb)
        if regula.colorScale is not None:          # WBUDOWANE: czytaj tu!
            print("  kolory:", [c.rgb for c in regula.colorScale.color])
        if regula.dataBar is not None:
            print("  pasek:", regula.dataBar.color)
        if regula.iconSet is not None:
            print("  ikony:", regula.iconSet.iconSet)

# ==================================================================
# 9. DIAGNOSTYKA
# ==================================================================
# Nurek: rozpakuj .xlsx (modul 02) -> xl/worksheets/sheetN.xml
#   <conditionalFormatting sqref="..."> ... <colorScale>/<dataBar>/<iconSet>
#   albo <cfRule dxfId="..."> + xl/styles.xml <dxfs>
# ⚠️ regula.dxf NIE dziala dla reguł wbudowanych (jest None)
# ⚠️ reguły w extLst sa niewidoczne i giną przy zapisie
# ⚠️ liczba regul = sum(len(cf.rules) for cf in ws.conditional_formatting)
#    (len(ws.conditional_formatting) liczy GRUPY ZAKRESOW - modul 13)
```

## 9. Co dalej

Zamknąłeś świat formatowania warunkowego. Umięsz już **wszystko**, co openpyxl potrafi w tym zakresie: reguły standardowe (moduł 13), wbudowane wizualizacje (ten moduł) i reguły bez fabryki budowane bazowym `Rule`. Trzy rzeczy, które zabierasz:

- **Mapę „pięć fabryk, osiemnaście typów".** Wiesz, że `CellIsRule`, `FormulaRule`, `ColorScaleRule`, `DataBarRule` i `IconSetRule` to funkcje, a nie klasy, i że resztę budujesz przez `Rule(type=...)`. Nie dasz się już nabrać na nieistniejące API z tutoriali.
- **Rozróżnienie „malowanie vs mikro-wykres".** Reguły standardowe mają `dxf` i malują; wbudowane mają własne dane (progi, kolory, ikony) i rysują. Po wczytaniu wiesz, gdzie szukać koloru w każdym z dwóch przypadków.
- **Dyscyplinę priorytetów.** Scenariusz „skala kolorów + wyróżnienie skrajnych" pokazał, że kolejność dodawania jest częścią kontraktu, a nie szczegółem. To ta sama lekcja co warstwy w module 12 — rzeczy w arkuszu nakładają się na siebie i **kolejność jest częścią projektu**.

**Co teraz zrobisz w praktyce — jedno ćwiczenie obserwacyjne.** Otwórz `output/14_databar.xlsx` w Excelu:

1. Wejdź w **Formatowanie warunkowe → Zarządzaj regułami** i obejrzyj dwie reguły na zakresie `B2:B11`. Zobaczysz, że wyróżnienie minimum ma priorytet wyżej niż pasek danych. **Zmień kolejność w Excelu** (przyciski „W górę"/„W dół") i zapisz. Potem **wczytaj plik openpyxl** i wypisz priorytety (sekcja 8 ściągawki). Zobaczysz, że priorytety zmieniły się dokładnie tak, jak je przestawiłeś — to dowód, że priorytet jest danymi w pliku, a nie „ustawieniem" gdzieś z boku.
2. **Dodaj ręcznie w Excelu** zestaw ikon na kolumnie `A2:A11` (dowolny zestaw). Zapisz, wczytaj openpyxl i wypisz `regula.iconSet.iconSet` oraz `regula.iconSet.cfvo`. Zobaczysz, jak Excel zapisał progi — i **porównaj z tym, co zapisałby openpyxl**. To bezpośrednie przygotowanie do modułu 18: plik, który przeszedł przez ręce człowieka, ma inną strukturę niż ten wygenerowany.

W **module 15** wejdziemy w **wykresy** — czyli w kolejny świat obiektów, które openpyxl modeluje, ale już **nie tak kompletnie**. Zobaczysz, że wykres to opis („weź te liczby i narysuj tak"), a nie obrazek — i zmierzysz się z tym, co openpyxl odtwarza, a co upraszcza. Ale zanim tam wejdziemy, w module 15 zobaczysz również, że analogie z tego modułu — warstwy, progi, priorytety — wrócą w zupełnie nowej formie: przy osiach, seriach i dwóch skalach na jednym wykresie.