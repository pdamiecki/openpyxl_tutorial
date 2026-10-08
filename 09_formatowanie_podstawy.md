Zanim napiszę moduł, zweryfikuję trzy rzeczy, które są najczęstszym źródłem nieporozumień w materiałach o stylach openpyxl: dokładne zachowanie `StyleProxy` (czy ma metodę `copy`), reguły kolorów (kanał alfa) i rejestrację `NamedStyle`.

Dwa kluczowe fakty potwierdzone (niemutowalność i **deprecacja** `StyleProxy.copy()`). Sprawdzę jeszcze dwie rzeczy, na których łatwo się pomylić w materiale dydaktycznym: walidację kolorów i wzorzec „dodawania" stylów.

# Moduł 09 — Formatowanie: model stylu i pierwsze pieczątki

> **Część:** II — Formatowanie · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–08

## 0. W tym module nauczysz się

- **Zrozumiesz, że styl komórki to nie „ustawienia na komórce", a komplet sześciu osobnych, współdzielonych obiektów**: `Font`, `PatternFill`, `Border`, `Alignment`, `Protection` i `number_format`. To rozróżnienie wyjaśnia połowę błędów, które popełniają początkujący — i całą resztę problemów z formatowaniem w dużych raportach.
- **Zrozumiesz i udowodnisz eksperymentalnie, dlaczego `cell.font.bold = True` nie działa** — i dlaczego openpyxl celowo rzuca w tym miejscu wyjątek. To ta sama własność (niemutowalność), która sprawia, że jedna zmiana stylu nie psuje dwustu komórek.
- **Poznasz trzy poprawne sposoby zmiany jednej cechy stylu** (`copy()` ze standardowej biblioteki, operator `+` z gałęzi 3.1, pełne nowe obiekty) i będziesz wiedział, że czwarty popularny sposób — `cell.font.copy(bold=True)` — nadal działa, ale **jest oznaczony jako przestarzały**.
- **Opanujesz parametry wszystkich pięciu klas stylu komórki** na poziomie, który wystarczy do zbudowania profesjonalnego nagłówka, obwódki tabeli, zebry, zawijanego tekstu i formularza do wypełniania.
- **Zrozumiesz kolor tak, jak rozumie go plik `.xlsx`**: zapis aRGB, rola kanału alfa, dlaczego `#FF0000` z krzyżykiem wybucha, a `"FF0000"` jest całkowicie poprawny — i dlaczego przy wypełnieniu jednolitym kolor idzie do `fgColor`, mimo że intuicja podpowiada `bgColor`.
- **Umiesz tworzyć `NamedStyle`**, stosować go przez `cell.style = "Nazwa"` i rozumiesz jego fundamentalnie inną naturę niż styl komórki: **styl nazwany jest mutowalny, ale komórka, która go dostała, jest już od niego odłączona**.
- **Zbudujesz pierwszy prawdziwy rejestr stylów** (`STYLE_REGISTRY`) — wzorzec, który za kilkanaście modułów stanie się pełnym wzorcem konstrukcyjnym (moduł 23) i podstawą spójności każdego generowanego raportu.
- **Zobaczysz, co dokładnie trafia do `xl/styles.xml`** — indeksy `cellXfs`, deduplikację identycznych stylów i dwa zarezerwowane wypełnienia, o których istnieniu nie wiedzą nawet ludzie z kilkuletnim stażem.

## 1. Intuicja i analogia

### 1.1. Analogia główna: pieczątki w szufladzie, nie pola w formularzu

W module 03 powiedzieliśmy jedno zdanie o stylu: **„styl to pieczątka, którą przybijasz do kratki"**. Teraz rozłóżmy ją na części, bo cały ten moduł z niej wynika.

Wyobraź sobie biuro urzędu. Do przyjmowania dokumentów używa się **pieczątek**, a nie ołówka.

- Jest **szuflada z pieczątkami**. W niej leżą pieczątki: „WPŁYNĘŁO", „PILNE", „ZAŁĄCZNIK", data, podpis.
- Każda **pieczątka jest gotowym, wyprodukowanym przedmiotem**. Wytwarza się ją raz, w zakładzie grawerskim. Ma ustalony kształt, kolor tuszu, rozmiar, grubość ramki.
- Na dokumencie **odbijasz pieczątkę** w konkretnym miejscu. Dokument nie „ma pieczątki" — ma **odcisk**, a odcisk powstał z konkretnego egzemplarza pieczątki.
- Żeby zmienić wygląd odcisku na **jednym** dokumencie, nie próbujesz poprawiać tuszu na już odbitym śladzie. **Bierzesz inną pieczątkę i odbijasz ją jeszcze raz, obok.** Albo, jeśli wszystkim dokumentom brakuje czegoś, czego brakuje też na pieczątce — **zamawiasz nowy egzemplarz pieczątki**.

I teraz najważniejsze, bo to jest sedno tego modułu:

**Jedna pieczątka służy wielu dokumentom.** Załóżmy, że masz w szufladzie jedną pieczątkę „WPŁYNĘŁO" i odbiłeś ją na trzystu dokumentach. Jeśli **zmienisz na tej pieczątce jedną literę**, zmieni się ona na wszystkich trzystu odciskach — i nikt tego nie kontroluje. Bo odcisk nie jest kopią treści pieczątki, jest **odciskiem konkretnego egzemplarza**.

Dokładnie tak działają style w openpyxl:

- `Font`, `PatternFill`, `Border`, `Alignment`, `Protection` to **pieczątki**. Tworzy się je raz, a potem przybija do wielu komórek.
- `cell.font = Font(bold=True)` to **odciśnięcie pieczątki** — komórka zapamiętuje nie „jest pogrubiona", a „używa tego egzemplarza pieczątki".
- Dlatego openpyxl **uniemożliwia zmianę pieczątki przez odcisk**. Gdyby `cell.font.bold = True` działało, zmieniłoby wygląd wszystkich trzystu komórek, które dostały ten sam obiekt `Font`. To jest dokładnie ten efekt uboczny, o którym mówi dokumentacja openpyxl: *„Style są współdzielone między obiektami i po przypisaniu nie mogą być zmieniane. Zapobiega to niepożądanym efektom ubocznym, takim jak zmiana stylu wielu komórek, gdy zmienia się tylko jedną."*

To zdanie z dokumentacji warto zapamiętać dosłownie, bo ono **wyjaśnia sens tego ograniczenia**. To nie jest niedoróbka ani utrudnienie dla utrudnienia. To świadoma decyzja projektowa, która chroni Cię przed cichym zniszczeniem raportu.

### 1.2. Analogia do niemutowalności: etykieta wydrukowana, nie formularz do wypełnienia

Drugie ujęcie tej samej rzeczy, bo warto mieć dwa obrazy — jeden pomaga zrozumieć *dlaczego tak jest*, drugi *jak z tym żyć*.

Wyobraź sobie dwie etykiety cenowe w sklepie:

- **Etykieta wydrukowana** na drukarce etykiet: „19,99 zł", czcionka bezszeryfowa, ramka. Tego napisu nie poprawisz. Możesz najwyżej **wydrukować nową etykietę** i nakleić ją na towar.
- **Formularz do wypełnienia** długopisem: rubryki „cena: ____", „jednostka: ____". Wpisujesz, ścierasz, wpisujesz ponownie.

Style openpyxl to **etykiety wydrukowane**. Nie ma w nich rubryk do wypełnienia. Każdy atrybut jest nadrukowany i do zmiany wymaga nowego egzemplarza.

Praktyczna konsekwencja, którą trzeba zapamiętać w kościach:

```python
# ❌ TAK NIE ROBIMY - to nie zadziala
ws["A1"].font.bold = True
```

Ten kod podnosi wyjątek: `AttributeError`, z komunikatem w stylu *„Style objects are immutable and cannot be changed. Reassign the style with a copy"*. W polskim tłumaczeniu: **„Obiekty stylu są niemutowalne i nie można ich zmieniać. Przypisz styl ponownie, używając kopii."**

Zwróć uwagę: **komunikat nie tylko mówi, że nie wolno — mówi, co zrobić.** openpyxl prowadzi Cię za rękę. To rzadka i bardzo dobra praktyka biblioteczna, więc warto przeczytać ten komunikat ze zrozumieniem, a nie z irytacją.

### 1.3. Analogia do `copy()`: kserokopia etykiety plus dopisek długopisem

Skoro nie można zmienić pieczątki, a chcemy tylko **jedną** komórkę zmienić nieznacznie — jak to zrobić?

Zdjęcie z życia: masz firmową etykietę z adresem wydrukowanym. Chcesz jedną wersję zmodyfikować — dodać „/ magazyn 2". Nie zamawiasz nowej matrycy. **Robilisz kserokopię etykiety, dopisujesz długopisem brakujący fragment na kserokopii i naklejasz kserokopię.** Oryginalna matryca pozostaje nietknięta, więc pozostałe etykiety się nie zmieniają.

Dokładnie to robisz kodem:

```python
from copy import copy

nowy_font = copy(ws["A1"].font)     # 1. kserokopia etykiety
nowy_font.bold = True               # 2. dopisek na kserokopii
ws["A1"].font = nowy_font           # 3. naklejenie kserokopii
```

Trzy linie. Każda robi jedną rzecz. I **każda jest potrzebna** — pominięcie którejkolwiek nie daje efektu albo daje efekt inny, niż chciałeś. W sekcji 2.3 rozłożymy je na czynniki pierwsze.

### 1.4. Analogia do `NamedStyle`: styl akapitu w edytorze tekstu

Tu jest miejsce na drugie, całkowicie odmienne narzędzie. Wróćmy do biura, ale przenieśmy się do pokoju, w którym ktoś pisze dokumenty.

W edytorze tekstu (Word, LibreOffice Writer) masz dwie rzeczy:

- **Formatowanie bezpośrednie** — zaznaczasz akapit i klikasz „pogrubienie". Zmienia się tylko ten akapit. Jeśli potem szef powie „wszystkie nagłówki mają być granatowe", musisz przejść po wszystkich.
- **Style akapitów** — „Nagłówek 1", „Cytat", „Podpis". Nadajesz akapitowi etykietę, a wygląd siedzi w **definicji stylu**. I teraz: **zmieniasz definicję „Nagłówek 1" raz — i zmienia się wygląd wszystkich nagłówków w dokumencie.**

`NamedStyle` w openpyxl to dokładnie **styl akapitu**.

| | Styl komórki (`Font`, `Fill`, …) | Styl nazwany (`NamedStyle`) |
|---|---|---|
| Czym jest | Pojedyncza pieczątka | Zestaw pieczątek z nazwą |
| Czy mutowalny | **Nie** | **Tak** (to jest jego główna zaleta) |
| Zasięg nazwy | Brak nazwy | Nazwa w obrębie **skoroszytu** |
| Jak przypisać | `cell.font = Font(...)` | `cell.style = "Naglowek"` |
| Jak zmienić wygląd | Nowa pieczątka i przypisanie | Zmiana obiektu `NamedStyle` |
| Czy zmiana propaguje się na komórki | Nie dotyczy | **Nie!** (patrz sekcja 2.11) |

Zwróć uwagę na ostatni wiersz — to pułapka numer jeden wokół `NamedStyle` i wyjaśnimy ją dokładnie w teorii. Na razie zapamiętaj obraz: **styl nazwany to szablon, ale komórka, która go dostała, dostaje kopię odcisku, a nie stałe połączenie z szablonem.**

### 1.5. Analogia do `fgColor` i `bgColor`: farba i papier

Zostańmy w zakładzie poligraficznym. Żeby coś wydrukować, potrzebujesz dwóch rzeczy:

- **Farby** — tym malujesz.
- **Papieru** — po tym malujesz.

W wzorach wypełnienia Excela jest dokładnie tak samo, i to jest źródło najczęstszej pomyłki przy `PatternFill`:

- `fgColor` to **kolor farby** (ang. *foreground*, pierwszy plan).
- `bgColor` to **kolor papieru** (ang. *background*, tło).

Intuicja podpowiada: „no dobrze, więc tło komórki to `bgColor`". **Nie.** I tu jest cała pułapka: przy wypełnieniu jednolitym (`fill_type="solid"`) Excel **maluje całą powierzchnię farbą**. Farba to `fgColor`. Papieru nie widać — jest zakryty. **Więc kolor wypełnienia jednolitego ustawiasz w `fgColor`, a `bgColor` nie ma wtedy znaczenia.**

Możesz to zapamiętać tak: *„Solid to pomalowanie całej kartki farbą. Farby szukasz w `fgColor`. Papier (`bgColor`) zostaje pod spodem i go nie widać."*

Przy wzorach (kratka, paski, szachownica) widać oba — farba rysuje wzór, papier prześwieca między jego elementami. Dlatego dopiero tam `bgColor` zaczyna mieć sens.

### 1.6. Analogia do kanału alfa: farba kryjąca i folia przezroczysta

I jeszcze jedna, krótka, potrzebna do zrozumienia kolorów.

Kolor w pliku Excel zapisuje się jako **osiem znaków szesnastkowych**, a nie sześć:

```
FF 00 00 00
│  └──┴──┘
│     └── kolor: czerwony? zielony? niebieski (RGB)
└──────── przezroczystość (alfa)
```

- **Część RGB** (6 znaków) — to kolor farby: `FF0000` to czerwień.
- **Kanał alfa** (2 dodatkowe znaki) — to „jak bardzo kryjąca jest farba". W teorii `00` znaczy w pełni przezroczysta, `FF` w pełni kryjąca.

I teraz rzecz, o której musisz wiedzieć, bo prawie każdy tutorial mówi o tym źle: **openpyxl w ogóle nie przejmuje się kanałem alfa przy stylach komórek.** Dokumentacja mówi to wprost: *„Wartość alfa odnosi się w teorii do przezroczystości koloru, ale nie ma to znaczenia dla stylów komórek. Do każdej prostej wartości RGB domyślnie dodawane jest `00`."*

Czyli dwie rzeczy na raz:

1. Możesz podać **6 znaków** (`"FF0000"`) i openpyxl **sam dopisze `00`** z przodu → `"00FF0000"`. To nie jest błąd. To **udokumentowane, obsługiwane zachowanie**.
2. Ten `00` **nie oznacza, że Twój czerwony napis będzie przezroczysty.** Excel ignoruje alfę w stylach komórek. Napis będzie czerwony i kryjący.

Wrócimy do tego w sekcji 2.6, gdzie pokażę dokładny wzorzec, jakiego używa openpyxl do walidacji koloru — i wyjaśnię, **co naprawdę** powoduje błąd `Colors must be aRGB hex values`. Bo uwaga: to nie brak kanału alfa.

## 2. Teoria

### 2.1. Styl komórki to sześć pieczątek naraz

Zacznijmy od pełnej definicji, bo większość kursów ją skraca i potem wychodzą kwiatki.

**Styl komórki składa się z sześciu niezależnych elementów:**

| Element | Klasa | Za co odpowiada | Domyślna wartość |
|---|---|---|---|
| Czcionka | `Font` | Krój, rozmiar, pogrubienie, kursywa, podkreślenie, przekreślenie, kolor, indeks górny/dolny | `Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none', strike=False, color='FF000000')` |
| Wypełnienie | `PatternFill` | Kolor tła / wzór | `PatternFill(fill_type=None, start_color='FFFFFFFF', end_color='FF000000')` |
| Obramowanie | `Border` (z `Side`) | Linie: lewa, prawa, górna, dolna, przekątne | wszystkie `Side(border_style=None, color='FF000000')` |
| Wyrównanie | `Alignment` | Poziome, pionowe, zawijanie, obrót, wcięcie, zmniejszanie do rozmiaru | `Alignment(horizontal='general', vertical='bottom', text_rotation=0, wrap_text=False, shrink_to_fit=False, indent=0)` |
| Ochrona | `Protection` | Czy komórka jest zablokowana, czy formuła jest ukryta | `Protection(locked=True, hidden=False)` |
| Format liczby | `str` (kod formatu) | Jak liczba jest wyświetlana | `'General'` |

*Wartości domyślne w tabeli pochodzą wprost z dokumentacji openpyxl 3.1 — to nie moja interpretacja, to dokładnie te obiekty, które openpyxl utworzy, jeśli nic nie ustawisz.*

I teraz kluczowa obserwacja: **szósty element, `number_format`, nie jest obiektem klasy stylu — jest zwykłym napisem.** Ale należy do tego samego kompletu: siedzi na komórce obok pozostałych pięciu i razem z nimi tworzy jeden „zestaw formatowania". Dlatego w pliku `.xlsx` wszystkie sześć trafiają do **jednego rekordu** w tablicy stylów. To wyjaśnienie przyda się w sekcji 2.4.

**Trzy rozłączne światy, w których żyją te obiekty:**

Zanim pójdziemy dalej, musisz rozróżnić trzy sytuacje, które na pierwszy rzut oka wyglądają tak samo:

```python
from openpyxl.styles import Font

# (A) OBIEKT STYLU - tworzysz pieczatke w Pythonie. W tym momencie nie jest
#     jeszcze nigdzie przypisany i NIE TRAFI do pliku.
font_a = Font(bold=True)

# (B) PRZYPISANIE DO KOMORKI - odcisk pieczatki. Od tego momentu komorka
#     "uzywa" tego egzemplarza. To wciaz tylko stan w pamieci.
ws["A1"].font = font_a

# (C) ZAPIS SKOROSZYTU - dopiero teraz zestaw stylow trafia do pliku .xlsx
wb.save("output/raport.xlsx")
```

**Co się dzieje w pamięci.** `Font(bold=True)` tworzy obiekt Pythona. `ws["A1"].font = font_a` zapisuje w komórce *odwołanie* do tego obiektu (a dokładniej: openpyxl wpisuje odpowiedni zestaw identyfikatorów w wewnętrznej strukturze komórki). `wb.save` przegląda wszystkie komórki i buduje globalne tablice stylów.

**Co trafi do pliku.** `font_a` trafi do pliku **tylko wtedy, gdy został przypisany do co najmniej jednej komórki** (albo użyty w stylu nazwanym). Sam obiekt utworzony w Pythonie i nigdzie nieprzypisany jest dla pliku nieistniejący. To pierwsza pułapka, na którą natkniesz się w praktyce: „utworzyłem styl, ale go nie ma w pliku" — bo go nigdzie nie przypisałeś.

### 2.2. Niemutowalność — dlaczego i jak to działa

Spróbujmy zrobić rzecz, którą intuicja podpowiada jako pierwszą:

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws["A1"] = "Nagłówek"

ws["A1"].font.bold = True       # <- tu openpyxl sie zatrzyma
```

Wynik: `AttributeError` z komunikatem *„Style objects are immutable and cannot be changed. Reassign the style with a copy"* (dokładne brzmienie sprawdź w swojej wersji — w 3.1.x, jak widać w kodzie źródłowym, dwa zdania są sklejone bez spacji, co samo w sobie jest ciekawostką).

**Dlaczego ten wyjątek istnieje?** Bo `ws["A1"].font` **nie zwraca obiektu `Font`**. Zwraca obiekt **`StyleProxy`** — pełnomocnika, który *zachowuje się* jak `Font`, ale blokuje zapis.

**Analogia do pełnomocnika:** masz w firmie pieczątkę, która jest własnością wspólną — leży w ogólnodostępnej szufladzie i używa jej osiem działów. Kiedy chcesz „poprawić jedną literę", nie dostajesz pieczątki do ręki. Dostajesz **przezroczystą kopię pod szkłem, z napisem „do wglądu"** i instrukcją: „jeśli chcesz inną treść, zamów własny egzemplarz". Możesz się jej przyjrzeć, możesz ją porównać, możesz odczytać każdy detal — ale nie możesz jej zmienić.

To dokładnie `StyleProxy`:

- **Odczyt** — delegowany do prawdziwego obiektu (`proxy.bold` → `True`/`False`/`None`),
- **Zapis** — zablokowany, z wyjątkiem jednego wewnętrznego atrybutu, którym proxy trzyma referencję do celu,
- **Kopiowanie** — `copy(proxy)` zwraca **prawdziwy, zwykły obiekt `Font`**, którym możesz już swobodnie operować,
- **Dodawanie** — `proxy + Font(...)` zwraca nowy, połączony obiekt (o tym w 2.3.2),
- **Porównywanie** — `proxy == inny_font` działa, bo porównanie jest delegowane.

**Co daje niemutowalność w praktyce?** Trzy konkretne korzyści:

1. **Bezpieczeństwo współdzielenia.** Możesz przypisać **jeden** obiekt `Font` do tysiąca komórek i nie martwić się, że któryś fragment kodu go zmieni. To sprawia, że styl jest przewidywalny.
2. **Mniejszy plik.** Identyczne style są deduplikowane (sekcja 2.4) — a deduplikacja działa tylko wtedy, gdy „identyczny" znaczy naprawdę identyczny i nie zmieni się w połowie pisania.
3. **Testowalność.** Funkcja, która zwraca `Font`, zwraca wartość, którą można porównać asercją — a nie obiekt, który ktoś w międzyczasie zmodyfikował.

**A ograniczenie?** Jedno, ale za to poważne: nie można zmienić *jednej* cechy istniejącego stylu bez zbudowania nowego obiektu. I tu uwaga — zbudowanie **nowego obiektu od zera** jest bardzo częstym, poważnym błędem:

```python
# ❌ CICHA UTRATA DANYCH - najgroźniejszy blad w tym module
ws["A1"].font = Font(size=14, color="FF0000")   # ustawiamy rozmiar i kolor
ws["A1"].font = Font(bold=True)                 # ...i GUBIMY rozmiar i kolor!

print(ws["A1"].font.size)      # None
print(ws["A1"].font.color)     # None (Color() bez wartosci)
```

Ten kod **nie podnosi żadnego błędu**. Po prostu w drugiej linii tworzysz **zupełnie nowy `Font`**, w którym `size` i `color` są `None` — i przypisujesz go zamiast poprzedniego. Rozmiar 14 i czerwień przepadają. W Excelu zobaczysz czarny tekst Calibri 11.

To jest dokładnie ten typ błędu, o którym mówimy w całym kursie: **cichy i przekonujący**. Nie zobaczysz go ani w logu, ani w konsoli. Zobaczysz go dopiero, gdy klient powie „czemu to jest inne niż w zeszłym miesiącu".

### 2.3. Trzy poprawne sposoby zmiany jednej cechy stylu

#### 2.3.1. Sposób pierwszy (kanoniczny): `copy()` ze standardowej biblioteki

To jest **metoda oficjalnie zalecana przez dokumentację openpyxl** i ta, której powinieneś używać domyślnie.

```python
"""Zmiana jednej cechy stylu - metoda kanoniczna."""

from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook()
ws = wb.active
ws["A1"] = "Nagłówek"

# Ustawiamy PELNY zestaw cech, ktory ma obowiazywac
ws["A1"].font = Font(name="Calibri", size=14, color="FF2F5597")

# Teraz chcemy TYLKO dodac pogrubienie, zachowujac reszte.
nowy = copy(ws["A1"].font)      # kopia zwraca ZWYKLY Font, nie proxy
print(type(nowy).__name__)      # Font
nowy.bold = True                # mozna go swobodnie zmieniac
ws["A1"].font = nowy            # i przypisac z powrotem

print(ws["A1"].font.size)       # 14 - zachowane
print(ws["A1"].font.color.rgb)  # 00FF2F5597 - zachowane
print(ws["A1"].font.bold)       # True - dodane
```

**Co się dzieje w pamięci.** `copy(ws["A1"].font)` tworzy **nowy** obiekt `Font`, w którym wszystkie atrybuty są skopiowane z oryginału. Ten nowy obiekt nie jest niczym związany ze starym — możesz go zmieniać do woli. Przypisanie `ws["A1"].font = nowy` podmienia odwołanie w komórce.

**Co trafi do pliku.** Trzy style mogą trafić do `styles.xml`: domyślny `Font`, „granatowy Calibri 14" i „granatowy Calibri 14 pogrubiony". Jeżeli ten drugi nie jest nigdzie używany (bo go nadpisaliśmy), openpyxl go nie zapisze.

**Dlaczego to działa, a `copy` na `StyleProxy` nie zwraca proxy?** Bo `StyleProxy` ma zdefiniowaną metodę `__copy__`, która zwraca **kopię obiektu docelowego**, a nie kopię pełnomocnika. To eleganckie rozwiązanie: `copy()` „zdejmuje szkło" i daje Ci prawdziwe narzędzie.

#### 2.3.2. Sposób drugi (nowoczesny): operator `+`

W gałęzi 3.1 openpyxl dodał możliwość **łączenia stylów operatorem `+`**. Wygląda to tak:

```python
from openpyxl.styles import Font

# scalamy istniejacy styl z dodatkowa cecha
ws["A1"].font = ws["A1"].font + Font(bold=True)
```

**Jak to czytać:** „weź styl, który komórka ma teraz, dołóż do niego pogrubienie i przypisz wynik". To odpowiednik `copy()` + zmiana + przypisanie w jednej linii.

Warto wiedzieć, że ta forma **nie jest przypadkiem** — pojawia się wprost w komunikacie deprecacyjnym openpyxl, obok `copy(obj)`.

⚠️ **Uczciwa uwaga o wersjach.** Operator `+` na obiektach stylów oraz metoda `StyleProxy.copy(**kwargs)` to elementy gałęzi 3.1. Metoda `.copy()` **nadal działa, ale jest oznaczona jako przestarzała** — jeśli jej użyjesz, możesz zobaczyć `DeprecationWarning`. Jeśli pracujesz na openpyxl 3.0.x, operator `+` może nie być dostępny — wtedy użyj `copy()`. **Zawsze sprawdź w dokumentacji swojej wersji**, jeśli kod ma działać w kilku środowiskach.

#### 2.3.3. Sposób trzeci: pełny nowy obiekt (najlepszy w rejestrach)

W kodzie produkcyjnym najczęściej **nie chcesz** modyfikować istniejącego stylu. Chcesz powiedzieć: „ten nagłówek ma wyglądać tak" — i mieć to w jednym miejscu. Wtedy `copy()` nie jest potrzebne, bo nie ma czego kopiować:

```python
from openpyxl.styles import Alignment, Font, PatternFill

STYL_NAGLOWKA = {
    "font": Font(name="Calibri", size=11, bold=True, color="FFFFFFFF"),
    "fill": PatternFill(fill_type="solid", fgColor="FF2F5597"),
    "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
}

for komorka in ws["A1:F1"][0]:
    komorka.font = STYL_NAGLOWKA["font"]
    komorka.fill = STYL_NAGLOWKA["fill"]
    komorka.alignment = STYL_NAGLOWKA["alignment"]
```

Zwróć uwagę: **trzy obiekty utworzone raz, użyte w sześciu komórkach.** To jest ów „rejestr", do którego wrócimy w 2.12.

#### 2.3.4. Tabela decyzyjna: czego użyć

| Sytuacja | Zalecane podejście |
|---|---|
| Buduję styl od zera dla nowej komórki | Pełny nowy obiekt (2.3.3) |
| Chcę dodać jedną cechę do stylu, który **już mam na komórce** | `copy()` (2.3.1) |
| Ta sama operacja co wyżej, pracuję na 3.1+ i chcę krócej | operator `+` (2.3.2) |
| Chcę ten sam wygląd w 500 komórkach | Jeden obiekt + pętla (rejestr) |
| Chcę zmieniać wygląd „globalnie" po fakcie | `NamedStyle` (2.11) |
| Czytam styl z istniejącego pliku i chcę go zachować, zmieniając drobiazg | `copy()` |
| Czytam styl i chcę go zrozumieć/diagnozować | Odczyt przez proxy (bez zapisu) |

### 2.4. Co trafi do pliku: `styles.xml`, indeksy i deduplikacja

To sekcja łącząca moduł 02 (anatomia `.xlsx`) z bieżącym materiałem. Bez niej formatowanie jest magią; z nią — inżynierią.

**Krok 1: komórka w arkuszu nie ma stylu. Ma numer stylu.**

Plik `xl/worksheets/sheet1.xml` wygląda tak:

```xml
<row r="1" spans="1:3">
  <c r="A1" s="1" t="inlineStr"><is><t>Miesiąc</t></is></c>
  <c r="B1" s="1" t="inlineStr"><is><t>Sprzedaż</t></is></c>
  <c r="C1" s="2"><v>120000</v></c>
</row>
```

Atrybut `s="1"` przy komórce `A1` to **indeks do tablicy stylów**. Komórka `A1` nie wie, że jest pogrubiona, granatowa i wyśrodkowana. Wie tylko: „użyj stylu numer 1".

**Analogia:** to jak numer miejsca w teatrze. Bilet mówi „rząd 5, miejsce 12" — a nie „krzesło tapicerowane na czerwono, z podłokietnikiem po prawej". Wygląd krzesła siedzi w opisie sali, nie na bilecie.

**Krok 2: opis stylów siedzi w `xl/styles.xml`.**

Ten plik ma kilka sekcji, które odpowiadają naszym „pieczątkom":

```xml
<styleSheet>
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00 &quot;zł&quot;"/>
  </numFmts>

  <fonts count="2">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>

  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/></patternFill></fill>
  </fills>

  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>

  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>
```

Cztery rzeczy warte uwagi:

1. **`cellXfs` to „kasetka na pieczątki".** Każdy `<xf>` mówi: „weź czcionkę numer `fontId`, wypełnienie numer `fillId`, obramowanie numer `borderId`, format liczby numer `numFmtId`". Atrybut `s="1"` z komórki wskazuje **wiersz tej tabeli**. To dlatego mówimy, że styl komórki to sześć elementów naraz — po prostu w pliku są one spięte jednym `<xf>`.

2. **Dwa pierwsze wypełnienia są zarezerwowane.** Widzisz `<fill><patternFill patternType="none"/></fill>` i `<fill><patternFill patternType="gray125"/></fill>`? To nie Twoje wypełnienia. Format `.xlsx` **wymaga**, żeby wypełnienie numer 0 to było „brak wypełnienia", a numer 1 — wzór `gray125`. Twoje pierwsze własne wypełnienie zawsze dostanie numer 2 i wyżej. Jeśli kiedyś zajrzysz do pliku i zdziwisz się, dlaczego jest tam „wypełnienie, którego nie tworzyłem" — to jest właśnie ono.

3. **Własne formaty liczb dostają identyfikatory od 164 w górę.** Numery poniżej 164 są zastrzeżone dla formatów wbudowanych Excela. Dlatego w kodzie formatu `#,##0.00 "zł"` dostaje `numFmtId="164"` — pierwszy wolny numer. To Zapowiedź modułu 10.

4. **Fonty i wypełnienia są wspólne dla całego skoroszytu.** Jeśli w arkuszu `Dane` i w arkuszu `Podsumowanie` użyjesz identycznego nagłówka, w `styles.xml` będzie **jeden** wpis w `<fonts>` i **jeden** w `<fills>`. Deduplikacja dzieje się globalnie.

**Krok 3: deduplikacja i jej konsekwencje.**

openpyxl przy zapisie trzyma listy stylów w strukturach indeksowanych, które **sprawdzają, czy identyczny obiekt już istnieje**. Jeśli tak — zwracają jego indeks, zamiast dodawać nowy.

**Co to znaczy dla Ciebie w praktyce:**

| Twój kod | Stan w `styles.xml` | Koszt |
|---|---|---|
| Jeden `Font(bold=True)` przypisany do 1000 komórek | **1** wpis w `<fonts>` | Zero — idealnie |
| 1000 osobnych `Font(bold=True)`, po jednym na komórkę | Wciąż **1** wpis (openpyxl scali identyczne) | 1000 niepotrzebnych obiektów w Pythonie: czas i pamięć |
| 1000 × `Font(bold=True, size=11+0.0001*i)` | **1000** wpisów | Plik puchnie, Excel wolniej otwiera |

Pierwszy wiersz to wzorzec, do którego dążymy. Drugi wiersz to marnotrawstwo **po stronie Pythona** — plik nie rośnie, ale pętla alokuje tysiące obiektów niepotrzebnie. Trzeci wiersz to realny problem: plik rośnie i każda wersja Excela go odczuwa.

Do wydajności wrócimy w 2.13, ale zapamiętaj już teraz: **liczba unikalnych stylów w pliku to koszt, który ponosi każdy, kto otworzy raport.**

### 2.5. `Font` — krój, rozmiar i forma

Pełna sygnatura z dokumentacji 3.1 (celowo podaję obie nazwy parametrów, bo obie są przyjmowane):

```python
Font(name=None, sz=None, b=None, i=None, charset=None,
     u=None, strike=None, color=None, scheme=None, family=None,
     size=None, bold=None, italic=None, strikethrough=None, underline=None,
     vertAlign=None, outline=None, shadow=None, condense=None, extend=None)
```

| Parametr (długa nazwa) | Alias | Co robi | Praktyczne wartości |
|---|---|---|---|
| `name` | — | Nazwa kroju czcionki | `"Calibri"`, `"Arial"`, `"Times New Roman"` |
| `size` | `sz` | Rozmiar w punktach | `11`, `14`, `9.5` |
| `bold` | `b` | Pogrubienie | `True` / `False` |
| `italic` | `i` | Kursywa | `True` / `False` |
| `underline` | `u` | Podkreślenie | `'none'`, `'single'`, `'double'`, `'singleAccounting'`, `'doubleAccounting'` |
| `strike` | `strikethrough` | Przekreślenie | `True` / `False` |
| `color` | — | Kolor tekstu | `"FF0000"`, `Color(rgb="...")`, `Color(theme=..., tint=...)` |
| `vertAlign` | — | Indeks górny/dolny | `'superscript'`, `'subscript'` |
| `scheme` | — | Odwołanie do czcionki motywu | `'major'`, `'minor'` |
| `family` | — | Rodzina czcionek (liczba) | rzadko potrzebne |
| `charset`, `outline`, `shadow`, `condense`, `extend` | — | Parametry niszowe/legacy | praktycznie nieużywane |

**Zalecenie stylistyczne:** używaj **długich nazw** (`bold`, `size`, `underline`). Aliasy (`b`, `sz`, `i`, `u`) istnieją głównie po to, żeby openpyxl mógł czytać pliki, w których Excel zapisał skróty. Korzystanie z `Font(b=True)` w kodzie produkcyjnym utrudnia wyszukiwanie (wpiszesz `b=` i znajdziesz pół kodu) i wprowadza w błąd osobę czytającą.

#### 2.5.1. Pułapka `name` i `scheme` — czyja czcionka wygrywa?

To subtelna rzecz, która występuje w praktyce rzadziej, ale bywa myląca.

- **`name`** to jawne wskazanie konkretnego kroju: „ma być Arial, koniec dyskusji".
- **`scheme`** to odwołanie do **motywu**: `'major'` (czcionka nagłówków) albo `'minor'` (czcionka tekstu podstawowego). Motyw żyje w skoroszycie i można go zmienić — wtedy zmieniają się wszystkie czcionki, które się do niego odwołują.

Problem: **jeśli ustawisz oba naraz, Excel może uznać motyw za ważniejszy** i zignorować Twoją nazwę. Efekt: ustawiasz `Font(name="Arial")`, a w Excelu widzisz Calibri, bo motyw mówi `minor = Calibri`.

Praktyczna rada: **jeśli ustawiasz konkretny krój, wyczyść `scheme`**:

```python
from openpyxl.styles import Font

# Jawne "Arial", bez odwolania do motywu
font = Font(name="Arial", scheme=None)
```

A jeśli styl kopiujesz z istniejącej komórki (np. z pliku klienta) i chcesz zmienić krój — pamiętaj, że prawdopodobnie kopiujesz razem z `scheme='minor'`:

```python
from copy import copy

nowy = copy(ws["A1"].font)
nowy.name = "Arial"
nowy.scheme = None      # <- to czesto decyduje o powodzeniu
ws["A1"].font = nowy
```

Dokładne zachowanie zależy od wersji Excela i zawartości motywu pliku — **sprawdź w swoim środowisku**, czy zmiana kroju faktycznie się uwidacznia. To jedno z tych miejsc, gdzie „działa u mnie" bywa prawdą lokalną.

### 2.6. `Color` — jak naprawdę działają kolory

#### 2.6.1. Walidacja koloru — dokładny wzorzec

openpyxl waliduje kolory wyrażeniem regularnym. W gałęzi 3.1 wygląda ono tak:

```python
aRGB_REGEX = re.compile("^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$")
```

Rozłóżmy to na części:

- `[A-Fa-f0-9]{8}` — **osiem** znaków szesnastkowych (zapis aRGB), **albo**
- `[A-Fa-f0-9]{6}` — **sześć** znaków szesnastkowych (zapis RGB),
- `^...$` — **cały napis** musi pasować, nie fragment.

Dodatkowo openpyxl dopisuje alfę, gdy jej zabraknie — dokładnie tak, jak opisuje to dokumentacja:

```python
>>> from openpyxl.styles import Font
>>> font = Font(color="00FF00")
>>> font.color.rgb
'0000FF00'
```

Zwróć uwagę na wynik: podałeś 6 znaków `00FF00`, a dostałeś 8 znaków `0000FF00` — openpyxl **dołożył `00` z przodu**. To nie jest błąd ani artefakt. To udokumentowane zachowanie.

#### 2.6.2. Mit, który trzeba obalić: `Color(rgb="FF0000")` **nie** jest błędem

W wielu tutorialach — i w pierwszych szkicach tego kursu — powtarza się twierdzenie: *„`Color(rgb="FF0000")` bez kanału alfa to błąd"*. **To nieprawda w openpyxl 3.1.x.** Pokażmy to kodem i sprawdźmy, co naprawdę wybucha:

```python
"""Co naprawde jest bledem w kolorach - eksperyment."""

from openpyxl.styles import Color, Font, PatternFill

# --- PRZYPADKI POPRAWNE --------------------------------------------------
poprawne = [
    ("6 znakow (RGB)",            "FF0000"),
    ("8 znakow (aRGB)",           "FFFF0000"),
    ("6 znakow, zielony",         "00FF00"),
    ("male litery tez sa OK",     "ff0000"),
]
for opis, wartosc in poprawne:
    kolor = Color(rgb=wartosc)
    print(f"OK    {opis:<24} {wartosc!r:12} -> rgb={kolor.rgb!r}")

# --- PRZYPADKI BLEDNE ----------------------------------------------------
bledne = [
    ("krzyzyk na poczatku",       "#FF0000"),
    ("skrot 3-znakowy",           "F00"),
    ("nazwa slowna",              "red"),
    ("5 znakow",                  "FF000"),
    ("9 znakow",                  "FFFF00000"),
    ("puste",                     ""),
]
for opis, wartosc in bledne:
    try:
        Color(rgb=wartosc)
        print(f"BRAK BLEDU  {opis:<22} {wartosc!r}")
    except ValueError as blad:
        print(f"ValueError  {opis:<22} {wartosc!r:12} -> {blad}")

# --- TO SAMO PRZEZ FONT --------------------------------------------------
# Font(color=...) idzie ta sama droga, wiec zachowuje sie identycznie.
try:
    Font(color="#FF0000")
except ValueError as blad:
    print()
    print(f"Font(color='#FF0000') -> ValueError: {blad}")

# A to dziala i jest w porzadku:
f = Font(color="FF0000")
print(f"Font(color='FF0000') -> rgb={f.color.rgb!r} (openpyxl dopisal alfe '00')")
```

**Wydruk (w uproszczeniu):**

```
OK    6 znakow (RGB)            'FF0000'     -> rgb='00FF0000'
OK    8 znakow (aRGB)           'FFFF0000'   -> rgb='FFFF0000'
OK    6 znakow, zielony         '00FF00'     -> rgb='0000FF00'
OK    male litery tez sa OK     'ff0000'     -> rgb='00ff0000'
ValueError  krzyzyk na poczatku   '#FF0000'    -> Colors must be aRGB hex values
ValueError  skrot 3-znakowy       'F00'        -> Colors must be aRGB hex values
...
```

**Co się dzieje w pamięci.** Deskryptor `RGB` przyjmuje napis i sprawdza go wzorcem. Jeśli napis ma 6 znaków — dopisuje `"00"` z przodu. Jeśli nie pasuje do wzorca — rzuca `ValueError` z komunikatem **`Colors must be aRGB hex values`**. To dokładny, prawdziwy komunikat z openpyxl.

**Co trafi do pliku.** Ośmioznakowy napis aRGB, np. `FF0000` → w pliku `00FF0000`. Excel odczyta kolor z części RGB i **zignoruje alfę** w stylach komórek.

**Wnioski praktyczne:**

| Wpisujesz | Efekt |
|---|---|
| `"FF0000"` | ✅ działa, staje się `"00FF0000"` |
| `"FFFF0000"` | ✅ działa (jawna alfa `FF`) |
| `"#FF0000"` | ❌ `ValueError` — **usuń krzyżyk** |
| `"F00"` | ❌ `ValueError` — używaj pełnych 6 znaków |
| `"red"` | ❌ `ValueError` — nazwy słowne to nie ten mechanizm |

To bardzo powszechna pułapka, bo kolory w HTML/CSS zapisuje się **z krzyżykiem** i często w skrócie 3-znakowym. Jeśli kopiujesz kolor z projektu graficznego, zdejmij `#` i rozwiń skrót (`#F00` → `FF0000`).

#### 2.6.3. Trzy sposoby definiowania koloru

Dokumentacja openpyxl wyróżnia trzy sposoby, ale rekomenduje jeden:

```python
from openpyxl.styles import Color

# 1. aRGB - ZALECANE, dziala niezaleznie od motywu i wersji pliku
kolor = Color(rgb="FF2F5597")

# 2. Indeksowany - legacy, zalezy od palety w skoroszycie/aplikacji
kolor = Color(indexed=32)

# 3. Motyw + tint - zalezy od motywu w skoroszycie
kolor = Color(theme=6, tint=0.5)
```

Dlaczego **aRGB jest zalecane**: bo jest samowystarczalne. Kolor `themed` będzie wyglądał inaczej w każdym pliku o innym motywie, a `indexed` — inaczej w zależności od palety aplikacji. **Kolor aRGB to jedyny, który jest jednoznaczny.**

Dla porządku: istnieją też gotowe stałe kolorów w `openpyxl.styles.colors` (`BLACK`, `WHITE`, `RED`, `GREEN`, `BLUE`, `YELLOW`, `MAGENTA`, `CYAN`). Możesz ich użyć jako skrótu, ale w kodzie produkcyjnym jawne wartości są czytelniejsze.

### 2.7. `PatternFill` — wypełnienie i odwrócona intuicja

Synygnatura z obsługą aliasów:

```python
PatternFill(patternType=None, fgColor=Color(), bgColor=Color(),
            fill_type=None, start_color=None, end_color=None)
```

| Parametr | Alias | Znaczenie |
|---|---|---|
| `fill_type` / `patternType` | — | Typ wypełnienia: `None`, `"solid"`, `"darkGray"`, `"mediumGray"`, `"lightGray"`, `"gray125"`, `"darkHorizontal"`, `"darkVertical"`, `"darkDown"`, `"darkUp"`, `"darkGrid"`, `"darkTrellis"`, `"lightHorizontal"`, `"lightVertical"`, `"lightDown"`, `"lightUp"`, `"lightGrid"`, `"lightTrellis"` |
| `fgColor` / `start_color` | ← | **Kolor farby** — dla `solid` to kolor wypełnienia |
| `bgColor` / `end_color` | ← | **Kolor papieru** — widoczny przy wzorach |

**Trzy fakty, które trzeba zapamiętać:**

**Fakt 1: `fill_type=None` oznacza brak wypełnienia.** Jeśli napiszesz `PatternFill(fgColor="FFFF0000")` i zapomnisz o `fill_type`, komórka będzie miała wypełnienie **niewidoczne**. To klasyczna „literówka polegająca na opuszczeniu parametru". Zawsze podawaj `fill_type="solid"`.

Uwaga na subtelność z dokumentacji: `PatternFill("solid", fgColor="DDDDDD")` — pierwszy argument **pozycyjny** to `patternType`. To działa, ale jest mniej czytelne. W kodzie produkcyjnym zawsze pisz `fill_type="solid"` po nazwie.

**Fakt 2: dla `solid` kolor idzie do `fgColor`.** To już wyjaśniliśmy w analogii farby i papieru. Powtórzę, bo to naprawdę najczęstszy błąd:

```python
from openpyxl.styles import PatternFill

# ✅ tak - widoczne czerwone tlo
ws["A1"].fill = PatternFill(fill_type="solid", fgColor="FFFF0000")

# ❌ tak - tlo pozostanie biale (a w niektorych odbiorcach: czarne)
ws["A1"].fill = PatternFill(fill_type="solid", bgColor="FFFF0000")
```

Dlaczego wersja z `bgColor` „w niektórych odbiorcach" daje czarne tło? Bo `bgColor` dla `solid` bywa różnie interpretowany przez różne narzędzia czytające plik. Excel najczęściej ignoruje `bgColor` przy `solid` — ale LibreOffice albo inny konwerter może próbować użyć `end_color`. **Konsekwencja: plik wygląda inaczej w Excelu i inaczej w innych programach.** To bardzo dobry powód, żeby nie polegać na `bgColor` przy `solid`.

**Fakt 3: `bgColor` ma sens przy wzorach.** Wzór `gray125` wypełni komórkę drobnym deseniem — farbą rysowanym na papierze. Jeśli chcesz mieć wzór na żółtym papierze z czerwonymi kropkami:

```python
ws["A1"].fill = PatternFill(
    fill_type="gray125",
    fgColor="FFFF0000",     # czerwona farba - kropki
    bgColor="FFFFFF00",     # żółty papier - tło między kropkami
)
```

**Wzorzec zebry** — praktyczne zastosowanie, które warto znać w formie „dwa wypełnienia na cały arkusz":

```python
from openpyxl.styles import PatternFill

ZEBRA_PARZYSTE = PatternFill(fill_type="solid", fgColor="FFF2F2F2")
ZEBRA_NIEPARZYSTE = PatternFill(fill_type=None)

# Dwa obiekty na caly arkusz - nie tysiac.
for indeks, wiersz in enumerate(ws.iter_rows(min_row=2, max_col=6)):
    fill = ZEBRA_PARZYSTE if indeks % 2 == 0 else ZEBRA_NIEPARZYSTE
    for komorka in wiersz:
        komorka.fill = fill
```

To jeszcze nie jest „formatowanie warunkowe" (to moduły 13–14) — to zwykłe naprzemienne tło, liczone raz przy generowaniu.

### 2.8. `Border` i `Side` — obramowania

`Border` składa się z obiektów `Side` — po jednym na każdą krawędź:

```python
Border(left=None, right=None, top=None, bottom=None,
       diagonal=None, diagonal_direction=0,
       outline=None, vertical=None, horizontal=None)
```

`Side` opisuje **jedną** krawędź:

```python
Side(style=None, color=None, border_style=None)
```

**Dostępne style linii** (dokładnie te wartości przyjmuje `style`):

| Wartość | Wygląd |
|---|---|
| `None` | Brak linii |
| `"hair"` | Najcieńsza, włoskowata |
| `"thin"` | Cienka (standardowa) |
| `"medium"` | Średnia |
| `"thick"` | Gruba |
| `"dashed"` | Kreskowana |
| `"dotted"` | Kropkowana |
| `"double"` | Podwójna |
| `"dashDot"` | Kreska–kropka |
| `"dashDotDot"` | Kreska–kropka–kropka |
| `"mediumDashed"` | Średnia kreskowana |
| `"mediumDashDot"` | Średnia kreska–kropka |
| `"mediumDashDotDot"` | Średnia kreska–kropka–kropka |
| `"slantDashDot"` | Ukośna kreska–kropka |

**Co się dzieje w pamięci.** `Border(top=Side(...), bottom=Side(...))` to jeden obiekt z referencjami do obiektów `Side`. Przy zapisie wszystkie kombinacje linii są deduplikowane — jeśli 500 komórek ma identyczną obwódkę, w `styles.xml` jest jeden wpis w `<borders>`.

#### 2.8.1. Obwódka tabeli — cztery boki naraz

Najczęstszy scenariusz: tabela ma mieć ramkę wokół całości. To znaczy, że **każda komórka z brzegu** musi dostać odpowiednie boki. Buduje się to raz, jako zestaw czterech obiektów, i stosuje warunkowo:

```python
"""Obwodka zakresu - klasyczny problem i klasyczne rozwiazanie."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.utils.cell import range_boundaries

OUTPUT = Path(__file__).resolve().parent.parent / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

# Cztery "pieczatki" dla czterech krawedzi - tworzone RAZ
CIENKA = Side(style="thin", color="FF808080")
GRUBA = Side(style="medium", color="FF2F5597")
BEZ = Side(style=None)

BOK_LEWY = Border(left=GRUBA, right=CIENKA, top=CIENKA, bottom=CIENKA)
BOK_PRAWY = Border(left=CIENKA, right=GRUBA, top=CIENKA, bottom=CIENKA)
BOK_GORNY = Border(left=CIENKA, right=CIENKA, top=GRUBA, bottom=CIENKA)
BOK_DOLNY = Border(left=CIENKA, right=CIENKA, top=CIENKA, bottom=GRUBA)
NAROZNIK_LG = Border(left=GRUBA, right=CIENKA, top=GRUBA, bottom=CIENKA)
# ... i tak dalej dla czterech naroznikow
WEWNETRZNA = Border(left=CIENKA, right=CIENKA, top=CIENKA, bottom=CIENKA)


def obramuj_zakres(ws, zakres: str) -> None:
    """Obramowuje zakres: gruba ramka zewnetrzna + cienkie linie wewnatrz.

    Dla kazdej komorki decydujemy osobno o kazdym z czterech bokow -
    dokladnie tak, jak robi to czlowiek w Excelu.
    """
    min_col, min_row, max_col, max_row = range_boundaries(zakres)

    for wiersz in range(min_row, max_row + 1):
        for kolumna in range(min_col, max_col + 1):
            lewy = GRUBA if kolumna == min_col else CIENKA
            prawy = GRUBA if kolumna == max_col else CIENKA
            gorny = GRUBA if wiersz == min_row else CIENKA
            dolny = GRUBA if wiersz == max_row else CIENKA
            ws.cell(row=wiersz, column=kolumna).border = Border(
                left=lewy, right=prawy, top=gorny, bottom=dolny
            )


wb = Workbook()
ws = wb.active
ws.title = "Dane"
ws.append(["Miesiąc", "Sprzedaż", "Koszt"])
ws.append(["Styczeń", 120_000, 78_000])
ws.append(["Luty", 98_500, 71_200])
ws.append(["Marzec", 143_700, 88_400])

obramuj_zakres(ws, "A1:C4")
wb.save(OUTPUT / "09_obramowanie.xlsx")
print(f"Zapisano: {OUTPUT / '09_obramowanie.xlsx'}")
```

**Co się dzieje w pamięci.** Tworzymy 16 (w praktyce mniej dzięki deduplikacji) kombinacji `Border` i przypisujemy je do komórek. **Uwaga na wydajność:** tu w pętli tworzymy nowe obiekty `Border` dla każdej komórki. Przy 10 000 komórek to 10 000 niepotrzebnych alokacji. Poprawna wersja tworzy **cztery do dziewięciu** kombinacji z góry, a w pętli tylko wybiera którąś:

```python
def obramuj_zakres_szybko(ws, zakres: str) -> None:
    """To samo, ale z gotowa tablica kombinacji - bez alokacji w petli."""
    min_col, min_row, max_col, max_row = range_boundaries(zakres)

    # Tablica: [lewy, prawy, gorny, dolny] -> gotowy Border. Budujemy leniwie.
    cache: dict[tuple[Side, Side, Side, Side], Border] = {}

    def border_dla(lewy, prawy, gorny, dolny) -> Border:
        klucz = (lewy, prawy, gorny, dolny)
        if klucz not in cache:
            cache[klucz] = Border(left=lewy, right=prawy, top=gorny, bottom=dolny)
        return cache[klucz]

    for wiersz in range(min_row, max_row + 1):
        for kolumna in range(min_col, max_col + 1):
            ws.cell(row=wiersz, column=kolumna).border = border_dla(
                GRUBA if kolumna == min_col else CIENKA,
                GRUBA if kolumna == max_col else CIENKA,
                GRUBA if wiersz == min_row else CIENKA,
                GRUBA if wiersz == max_row else CIENKA,
            )
```

**Co trafi do pliku.** W `<borders>` pojawią się wpisy dla każdej unikalnej kombinacji czterech `Side`. Przy tej metodzie będzie ich dokładnie tyle, ile jest kombinacji na obwodzie i wewnątrz — czyli maksymalnie 16, niezależnie od tego, czy tabela ma 4 wiersze, czy 4000.

**Praktyczna uwaga:** `vertical` i `horizontal` w `Border` **nie służą** do rysowania linii wewnątrz komórki obok lewej/prawej. Te pola są używane w stylach różnicowych (`dxf`, moduły 13–14) i stylach tabel. Dla zwykłej komórki potrzebujesz wyłącznie `left`, `right`, `top`, `bottom`. **Sprawdź w dokumentacji swojej wersji**, jeśli chcesz ich użyć w innym kontekście.

**Uwaga o aliasach `start`/`end`:** w plikach zapisanych przez nowsze wersje Excela krawędzie bywają zapisane jako `start`/`end` zamiast `left`/`right`. openpyxl obsługuje to przy odczycie — nie zdziw się, jeśli zobaczysz te nazwy w diagnostyce. W kodzie pisz `left`/`right`.

#### 2.8.2. Obramowanie komórek scalonych — sprawdź okiem

Dokumentacja openpyxl podaje, że *„komórka scalona zachowuje się podobnie jak inne komórki — jej wartość i format są zdefiniowane w komórce w lewym górnym rogu; żeby zmienić obramowanie całej scalonej komórki, zmień obramowanie jej komórki lewej górnej"*.

W praktyce jest to jednak obszar, w którym różne wersje Excela i różne konwertery zachowują się różnie. Dlatego **w tym jednym miejscu nie ufam dokumentacji na słowo** i daję Ci regułę:

> **Jeśli obramowujesz scalony zakres, zastosuj obramowanie do top-left cell — i jeśli w Excelu ramka nie wygląda dobrze, zastosuj obramowanie do *wszystkich* komórek w zakresie scalonym.** Sprawdź wizualnie w Excelu (i jeśli to możliwe w LibreOffice), zanim uznasz to za gotowe.

To jest dokładnie ta kategoria „openpyxl potrafi tylko częściowo" z modułu 03: mechanizm istnieje, ale wynik zależy od tego, kto odczytuje plik.

### 2.9. `Alignment` — wyrównanie, zawijanie i pułapka wysokości wiersza

```python
Alignment(horizontal=None, vertical=None, text_rotation=0,
          wrap_text=False, shrink_to_fit=False, indent=0, justifyLastLine=False,
          readingOrder=0, relativeIndent=None)
```

| Parametr | Wartości | Praktyczne znaczenie |
|---|---|---|
| `horizontal` | `'general'`, `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'centerContinuous'`, `'distributed'` | Poziome wyrównanie |
| `vertical` | `'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'` | Pionowe wyrównanie |
| `wrap_text` | `True`/`False` | Zawijanie tekstu w komórce |
| `text_rotation` | liczba (zakres sprawdź w swojej wersji) | Obrót tekstu |
| `indent` | liczba całkowita | Wcięcie od krawędzi |
| `shrink_to_fit` | `True`/`False` | Zmniejszanie czcionki, żeby tekst się zmieścił |
| `justifyLastLine` | `True`/`False` | Kolumna wyjustowana także w ostatniej linii |

#### 2.9.1. Pułapka wysokości wiersza — najważniejsza rzecz w tej sekcji

**openpyxl nigdy nie liczy wysokości wiersza.** Nie mierzy tekstu, nie wie, ile linii zajmie po zawinięciu, nie potrafi ustawić `height`. To nie jest jego zadanie i tego nie robi.

Co się dzieje dalej zależy od Ciebie:

| Co zrobiłeś | Co zrobi Excel przy otwarciu |
|---|---|
| Nie ustawiłeś `row_dimensions[n].height` | **Dopasuje wysokość automatycznie** do zawiniętego tekstu |
| Ustawiłeś `height = 15` | **Uszanuje Twoją wartość** i **tekst zostanie przycięty**, mimo `wrap_text=True` |

To już zupełnie odwraca intuicję. Bo brzmi jak: „jeśli chcesz ładnie, ustaw wysokość". A jest odwrotnie: **jeśli chcesz, żeby Excel sam dopasował wysokość do zawiniętego tekstu, nie ustawiaj wysokości w ogóle.**

Dlaczego tak jest: wysokość wiersza to własność *widoku*, którą Excel traktuje leniwie. Gdy nie ma jej zapisanej, Excel liczy ją przy renderowaniu. Gdy jest zapisana, Excel uznaje ją za decyzję użytkownika i jej nie nadpisuje.

**Praktyczna reguła:**

```python
# ✅ Zawijanie tekstu bez ustawiania wysokosci - Excel dopasuje sam
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")

# ❌ Zawijanie + sztywna wysokosc = przyciety tekst
ws["A1"].alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[1].height = 15          # tekst sie utnie

# ⚠️ Jesli MUSISZ podac wysokosc, musisz ją oszacowac samodzielnie
ws.row_dimensions[1].height = 45          # ~3 linie x 15 punktow
```

#### 2.9.2. `wrap_text` i `shrink_to_fit` — wybierz jedno

W interfejsie Excela te dwie opcje są **wzajemnie wykluczające** — zaznaczasz jedną albo drugą. Jeśli w pliku XML znajdą się ustawione obie, Excel może zgłosić problem albo wybrać jedną według własnego uznania.

Dlatego traktuj je jak przełącznik:

```python
# Albo zawijamy...
ws["A1"].alignment = Alignment(wrap_text=True)
# ...albo zmniejszamy czcionke. Nie oba naraz.
ws["A1"].alignment = Alignment(shrink_to_fit=True)
```

#### 2.9.3. Obrót tekstu

Obrót przydaje się w nagłówkach kolumn z długimi nazwami: tekst pod kątem 45° pozwala zmieścić „Liczba zamówień w kwartale" w wąskiej kolumnie.

```python
# 45 stopni w gore
ws["A1"].alignment = Alignment(text_rotation=45, vertical="bottom")
```

Zakres dopuszczalnych wartości i sposób zapisu wartości pionowej (Excel używa do tego wartości `255`) różnią się między wersjami. **Sprawdź w dokumentacji swojej wersji**, jeśli potrzebujesz tekstu pionowego — sprawdź też, czy openpyxl przyjmie wartość i jak ją zapisze.

### 2.10. `Protection` — przygotowanie do modułu 17

Szósty element stylu, o którym najczęściej się zapomina: **ochrona to część stylu komórki**. Domyślnie każda komórka ma `Protection(locked=True, hidden=False)`.

| Parametr | Znaczenie |
|---|---|
| `locked` | Czy komórka jest zablokowana do edycji |
| `hidden` | Czy formuła ma być ukryta w pasku formuły |

I teraz rzecz, która zaskakuje: **`locked=True` sam z siebie nic nie robi.** Zablokowanie komórki jest respektowane tylko wtedy, gdy **arkusz jest chroniony**. Analogia: masz zamek w drzwiach i klucz w kieszeni — dopóki nie przekręcisz klucza, drzwi stoją otwarte.

```python
from openpyxl import Workbook
from openpyxl.styles import Protection

wb = Workbook()
ws = wb.active

# Domyslnie: wszystko zablokowane. Ale arkusz nie jest chroniony,
# wiec uzytkownik i tak moze edytowac wszystko.
print(ws["A1"].protection.locked)      # True

# Klasyczny formularz: wiekszosc zablokowana, kilka pol do wypelnienia.
for wiersz in range(2, 11):
    ws.cell(row=wiersz, column=2).protection = Protection(locked=False)

# Dopiero TERAZ blokada dziala (pelne omowienie: modul 17).
ws.protection.sheet = True
```

**Uwaga na kolizję nazw.** Istnieją dwie klasy o podobnej nazwie:

- `openpyxl.styles.Protection` — **styl komórki** (czy ta jedna komórka jest zablokowana),
- `openpyxl.worksheet.protection.SheetProtection` — **ochrona arkusza** (`ws.protection`, czy arkusz jako całość jest chroniony).

Nie są tym samym i nie da się ich zamienić. Jeśli kiedykolwiek dostaniesz dziwny błąd przy `import`, to prawdopodobnie pomyliłeś te dwie.

Pełne omówienie — hasła, uprawnienia, `veryHidden`, słabość tej ochrony — jest w module 17. Tutaj zapamiętaj tylko tyle: **`locked=False` to część zestawu stylu komórki, a nie osobny mechanizm.**

### 2.11. `NamedStyle` — styl z nazwą, czyli szablon

Przejdźmy do rzeczy, dla której ten moduł istnieje.

#### 2.11.1. Tworzenie i rejestracja

```python
"""NamedStyle - pelny cykl zycia."""

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side

OUTPUT = Path(__file__).resolve().parent.parent / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "KPI"

# --- 1. Tworzymy styl nazwany -------------------------------------------
kpi = NamedStyle(name="KPI")
kpi.font = Font(name="Calibri", size=20, bold=True, color="FFFFFFFF")
kpi.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
kpi.border = Border(
    left=Side(style="thin", color="FF1F3864"),
    right=Side(style="thin", color="FF1F3864"),
    top=Side(style="thin", color="FF1F3864"),
    bottom=Side(style="thin", color="FF1F3864"),
)
kpi.alignment = Alignment(horizontal="center", vertical="center")
kpi.number_format = "#,##0"

print(f"Przed rejestracja: {'KPI' in wb._named_styles}")   # -> False

# --- 2. Rejestrujemy w skoroszycie --------------------------------------
wb.add_named_style(kpi)

print(f"Po rejestracji:    {'KPI' in wb._named_styles}")   # -> True

# --- 3. Stosujemy PO NAZWIE --------------------------------------------
ws["A1"] = "Przychód"
ws["A1"].style = "KPI"          # <- referencja po nazwie

# Uwaga: styl mozna tez przypisac obiektem - openpyxl zarejestruje go
# automatycznie przy pierwszym przypisaniu:
ws["B1"] = "Marża"
ws["B1"].style = kpi

wb.save(OUTPUT / "09_named_style.xlsx")
```

**Trzy drogi do rejestracji** — i wszystkie są poprawne:

1. **Jawna:** `wb.add_named_style(styl)` — najbezpieczniejsza, bo widzisz w kodzie, że styl istnieje.
2. **Przy pierwszym przypisaniu obiektu:** `ws["A1"].style = kpi` — dokumentacja mówi wprost: *„style nazwane zostaną zarejestrowane automatycznie przy pierwszym przypisaniu ich do komórki"*.
3. **Po nazwie:** `ws["A1"].style = "KPI"` — działa tylko wtedy, gdy styl **już** jest zarejestrowany.

**Reguła praktyczna:** rejestruj jawnie, `add_named_style`, i przypisuj po nazwie. Dlaczego? Bo wtedy **cały plik odwołuje się do stylu po nazwie**, a nie przez obiekt. To znaczy, że możesz zmienić definicję stylu w jednym miejscu i — o ile zrobisz to *przed* przypisaniem — wpłynie to na wszystko.

#### 2.11.2. ZALETA: styl nazwany jest mutowalny

To odróżnia `NamedStyle` od stylu komórki. Dokumentacja mówi: *„W przeciwieństwie do stylów komórek, style nazwane są mutowalne."*

```python
kpi = NamedStyle(name="KPI")
kpi.font = Font(size=20, bold=True)

kpi.font.size = 24              # <- TO DZIALA! NamedStyle nie jest proxy.
```

Możesz więc zdefiniować styl, a potem go **dopracowywać** — i to jest bardzo wygodne, gdy styl opisuje piętnaście właściwości, a Ty chcesz zmienić jedną.

#### 2.11.3. PUŁAPKA: komórka nie śledzi późniejszych zmian stylu

I teraz to jedno zdanie z dokumentacji, które musisz wytatuować sobie na pamięci:

> **„NB. raz przypisany styl nazwany do komórki nie podlega dalszym zmianom tego stylu — późniejsze modyfikacje stylu nie wpłyną na komórkę."**

Innymi słowy: **`NamedStyle` to nie „łącze na żywo". To moment skopiowania.**

```python
kpi = NamedStyle(name="KPI")
kpi.font = Font(size=20, bold=True, color="FFFF0000")
wb.add_named_style(kpi)

ws["A1"].style = "KPI"          # komorka A1 zapamietuje ODcisk sprzed zmiany

kpi.font = Font(size=30, italic=True)   # zmieniamy definicje stylu
wb.save("output/x.xlsx")
```

**Co zobaczysz w Excelu?** Komórka `A1` będzie miała **rozmiar 20 i pogrubienie**, a nie 30 i kursywę. Mimo że styl „KPI" definiuje teraz coś innego.

**Dlaczego?** Bo `cell.style = "KPI"` w momencie wykonania **odczytuje aktualne wartości stylu i wpisuje je na komórkę**. To jest ten sam mechanizm, o którym mówiliśmy przy `copy_worksheet` w module 08: operacja przypisania jest kopiowaniem wartości, nie tworzeniem stałego powiązania.

**Konsekwencja praktyczna, bardzo ważna:** kolejność jest nieprzypadkowa.

```python
# ✅ DOBRZE: najpierw dopracuj styl, potem przypisuj
kpi = NamedStyle(name="KPI")
kpi.font = Font(size=20, bold=True)
kpi.number_format = "#,##0"
wb.add_named_style(kpi)
# ...dopracowanie zakonczone...
ws["A1"].style = "KPI"          # teraz przypisujemy gotowy styl
ws["B1"].style = "KPI"

# ❌ ZLE: przypisujesz, potem zmieniasz
ws["A1"].style = "KPI"
kpi.font = Font(size=30)        # A1 tego NIE zobaczy
```

To jest dokładnie ten powód, dla którego w 2.3.3 powiedziałem: buduj obiekty stylu raz **przed** użyciem, a nie modyfikuj w trakcie.

#### 2.11.4. `NamedStyle` a skoroszyty — zasięg nazwy

Nazwa stylu żyje **w obrębie jednego skoroszytu**. To ma dwie konsekwencje:

1. **Dwa skoroszyty mogą mieć styl o tej samej nazwie i różnej definicji.** `raport_q1.xlsx` może mieć „Naglowek" granatowy, a `raport_q2.xlsx` — zielony. Excel nie zgłosi konfliktu, bo to dwa różne pliki.
2. **Rejestrując styl w skoroszycie, w którym nazwa już istnieje, dostaniesz błąd.** openpyxl sprawdza unikalność nazw w `add_named_style` i podnosi `ValueError` — **sprawdź dokładny komunikat i typ w swojej wersji**.

I dlatego powstaje **praktyczna pułapka projektu**, o której warto wiedzieć zawczasu: jeśli wczytujesz szablon i chcesz dodać własny styl o nazwie, która w szablonie już istnieje (np. `Naglowek`), Twój kod wybuchnie. Rozwiązanie:

```python
def zarejestruj_bezpiecznie(wb, styl: NamedStyle) -> str:
    """Rejestruje styl, a jesli nazwa jest zajeta - dokleja sufiks.

    Zwraca nazwe, pod ktora styl faktycznie zostal zarejestrowany.
    """
    if styl.name not in wb._named_styles:
        wb.add_named_style(styl)
        return styl.name

    licznik = 2
    while f"{styl.name}_{licznik}" in wb._named_styles:
        licznik += 1
    styl.name = f"{styl.name}_{licznik}"
    wb.add_named_style(styl)
    return styl.name
```

⚠️ **Uwaga:** `wb._named_styles` to atrybut **prywatny** — używam go tutaj wyłącznie do sprawdzenia, jakie nazwy są zajęte, bo openpyxl nie udostępnia w tym miejscu wygodnego publicznego API w postaci listy nazw. **Sprawdź w dokumentacji swojej wersji**, czy istnieje publiczny odpowiednik (`wb.named_styles` istnieje w nowszych wersjach — sprawdź, czy jest dostępny u Ciebie). To dokładnie ta sama granica, o której mówiliśmy przy `ws._images` w module 08.

#### 2.11.5. Style wbudowane — uwaga na lokalizację nazw

Excel ma wbudowane style nazwane. Dokumentacja openpyxl ostrzega: *„Niestety nazwy tych stylów są przechowywane w ich zlokalizowanych formach. openpyxl rozpozna tylko nazwy angielskie i tylko dokładnie tak, jak tu zapisano."*

Kilka z nich:

- `'Normal'` — odpowiednik braku stylu,
- `'Comma'`, `'Comma [0]'`, `'Currency'`, `'Currency [0]'`, `'Percent'` — formaty liczb,
- `'Title'`, `'Headline 1'` … `'Headline 4'`, `'Hyperlink'`, `'Followed Hyperlink'` — style tekstowe,
- `'Good'`, `'Bad'`, `'Neutral'`, `'Input'`, `'Output'`, `'Check Cell'` — porównania,
- `'Accent1'` … `'Accent6'` i warianty `'20 % - Accent1'`, `'40 % - Accent1'`, `'60 % - Accent1'` — wyróżnienia,
- `'Pandas'`.

**Pułapka wielojęzyczna:** w polskim Excelu te style mają polskie nazwy („Dobry", „Zły", „Neutralny"). Ale openpyxl rozpoznaje **tylko angielskie**. Jeśli więc generujesz plik i chcesz użyć wbudowanego stylu, pisz `ws["A1"].style = "Good"`, a nie `"Dobry"`. I nie zdziw się, że w polskim Excelu ten styl wyświetli się pod nazwą lokalną — to ta sama definicja, inna etykieta.

**Praktyczna rada:** wbudowanych stylów używaj **rzadko**. Są zależne od wersji i lokalizacji, a ich wygląd może się zmienić przy aktualizacji Excela. Własny `NamedStyle` daje Ci pełną kontrolę (i pełną odpowiedzialność za spójność, co zrobimy w 2.12).

### 2.12. Wzorzec: rejestr stylów

Zbierzmy wszystko w jedną konstrukcję, która jest **pierwszym prawdziwym wzorcem architektonicznym w tym kursie** — i pomostem do modułu 23.

**Problem:** w kodzie na 900 linii napiszesz `"2F5597"` czterdzieści razy. W jednym miejscu jako `Font`, w innym jako `PatternFill`, w trzecim z literówką `"2F5957"`. Po pół roku nikt nie wie, który odcień granatu jest „tym właściwym".

**Rozwiązanie:** jedno miejsce, w którym **definiuje się wygląd**, i cały kod, który tylko się do niego odwołuje.

```python
"""Rejestr stylow - jedno zrodlo prawdy o wygladzie raportu."""

from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ----------------------------------------------------------------------
# 1. PALETA - jedyne miejsce, gdzie pojawiaja sie literaly kolorow
# ----------------------------------------------------------------------
PALETA = {
    "granat":       "2F5597",
    "granat_ciemny": "1F3864",
    "zielen":       "548235",
    "czerwien":     "C00000",
    "szary_linia":  "BFBFBF",
    "szary_tlo":    "F2F2F2",
    "biel":         "FFFFFF",
    "czern":        "000000",
}

# ----------------------------------------------------------------------
# 2. LINIE I OBRAMOWANIA - tez tworzone raz
# ----------------------------------------------------------------------
LINIA_CIENKA = Side(style="thin", color=f"FF{PALETA['szary_linia']}", )
LINIA_GRUBA = Side(style="medium", color=f"FF{PALETA['granat']}")
RAMKA = Border(left=LINIA_GRUBA, right=LINIA_GRUBA, top=LINIA_GRUBA, bottom=LINIA_GRUBA)
SIATKA = Border(left=LINIA_CIENKA, right=LINIA_CIENKA, top=LINIA_CIENKA, bottom=LINIA_CIENKA)

# ----------------------------------------------------------------------
# 3. REJESTR STYLOW - nazwa -> obiekt stylu
# ----------------------------------------------------------------------
STYLE = {
    "naglowek": {
        "font": Font(name="Calibri", size=11, bold=True, color="FFFFFFFF"),
        "fill": PatternFill(fill_type="solid", fgColor=f"FF{PALETA['granat']}"),
        "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "border": SIATKA,
    },
    "naglowek_zielony": {
        "font": Font(name="Calibri", size=11, bold=True, color="FFFFFFFF"),
        "fill": PatternFill(fill_type="solid", fgColor=f"FF{PALETA['zielen']}"),
        "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "border": SIATKA,
    },
    "tytul": {
        "font": Font(name="Calibri", size=16, bold=True, color=f"FF{PALETA['granat_ciemny']}"),
        "alignment": Alignment(horizontal="left", vertical="center"),
    },
    "waluta": {
        "number_format": '#,##0.00 "zł"',
        "alignment": Alignment(horizontal="right"),
    },
    "procent": {
        "number_format": "0.0%",
        "alignment": Alignment(horizontal="right"),
    },
    "data": {
        "number_format": "yyyy-mm-dd",
        "alignment": Alignment(horizontal="center"),
    },
    "zebra": {
        "fill": PatternFill(fill_type="solid", fgColor=f"FF{PALETA['szary_tlo']}"),
    },
    "podsumowanie": {
        "font": Font(bold=True),
        "border": Border(top=LINIA_CIENKA, bottom=LINIA_GRUBA),
    },
}


# ----------------------------------------------------------------------
# 4. APLIKATOR - jedno miejsce, ktore wie, jak styl nalozic na komorke
# ----------------------------------------------------------------------
def zastosuj_styl(komorka, nazwa: str) -> None:
    """Nakłada styl z rejestru na komorke.

    Zna wszystkie szesc elementow stylu - i gdyby w przyszlosci doszedl
    siódmy, zmienia sie tylko ta funkcja.
    """
    if nazwa not in STYLE:
        raise KeyError(f"nie ma stylu {nazwa!r}; dostepne: {sorted(STYLE)}")

    for wlasciwosc, wartosc in STYLE[nazwa].items():
        setattr(komorka, wlasciwosc, wartosc)
```

**Co tu się dzieje i dlaczego to jest ważne:**

1. **`PALETA` jest jedynym miejscem z literałami kolorów.** Zmiana firmowego granatu = jedna linia w jednym pliku. To wzorzec **konfiguracja zamiast kodu** — pełna wersja w module 26.
2. **`STYLE` to słownik, a nie tysiąc linii `setattr`.** Dodanie nowego stylu to dodanie wpisu, a nie modyfikacja funkcji. To zasada **otwarte–zamknięte** (OCP) w wersji Excel-owej — zapowiedź modułu 22.
3. **`zastosuj_styl` jest jednym punktem wejścia.** Jeśli w przyszłości dojdzie siódmy element stylu, zmieniasz **jedną pętlę**, a nie sto miejsc w kodzie.
4. **`f"FF{PALETA['granat']}"`** — dodajemy alfę raz, w jednym miejscu, zamiast wpisywać `"FF2F5597"` ręcznie w każdym Stringu. Mniej okazji do literówki.

**Użycie:**

```python
for komorka in ws["A1:F1"][0]:
    zastosuj_styl(komorka, "naglowek")

for wiersz in ws["A2:F100"]:
    for komorka in wiersz:
        zastosuj_styl(komorka, "waluta")
```

**Co się dzieje w pamięci.** Wszystkie obiekty stylu powstają **przy imporcie modułu** — raz, przy starcie programu. Pętla tylko je przypisuje. Nawet przy 100 000 komórek w Pythonie istnieje dokładnie jeden `Font` nagłówka i jeden `PatternFill` nagłówka.

**Co trafi do pliku.** W `styles.xml` będzie **kilka** wpisów w `<fonts>` i `<fills>` — dokładnie tyle, ile unikalnych kombinacji użyto — niezależnie od tego, ile komórek sformatujesz.

**To jest wzorzec, który stanie się kanwą modułu 23.** Tam nazwiemy go **rejestrem** (Registry) w kontekście „współdzielenia ciężkich obiektów" i **fabryką** (Factory) w kontekście „skąd się biorą gotowe zestawy". Na razie wystarczy, że masz działającą, profesjonalną jego wersję.

### 2.13. Wydajność: tworzenie tysięcy identycznych obiektów

Zajmijmy się teraz kosztem, który realnie odczujesz.

**Trzy poziomy kosztu:**

| Poziom | Co się dzieje | Kiedy boli |
|---|---|---|
| **Python** | Alokacja obiektu `Font`/`Fill` per komórka | Zawsze — 100 000 komórek = 100 000 alokacji |
| **Pamięć procesu** | Każdy obiekt zajmuje kilkaset bajtów | Przy > 1 mln komórek |
| **Plik i Excel** | Liczba unikalnych wpisów w `styles.xml` | Gdy style różnią się na tyle, że deduplikacja nie pomaga |

**Poziom 1 i 2 leczy się trywialnie:** tworzysz obiekty stylu **raz**, poza pętlą. To wzorzec z 2.12.

**Poziom 3 to realne ryzyko,** bo deduplikacja pomaga tylko przy **identycznych** stylach. Jeśli w pętli zrobisz coś takiego:

```python
# ❌ RECEPT NA KATASTROFE: 5000 unikalnych stylow w jednym pliku
for indeks, wiersz in enumerate(ws.iter_rows(min_row=2, max_row=5000)):
    komorka = wiersz[3]
    komorka.font = Font(size=11 + indeks * 0.001)   # kazda inna!
```

...to `styles.xml` urośnie o 5000 wpisów, a Excel będzie odczuwał to przy każdym otwarciu (i przy każdej operacji na stylach). **Liczba unikalnych stylów w pliku to dług, który spłaca każdy odbiorca raportu.**

**Test, który warto zapisać sobie w notatkach:** zapisz plik, rozpakuj go, policz wpisy w `<fonts>` i `<fills>`. Jeśli liczba jest w tysiącach, a komórek w dziesiątkach tysięcy — coś jest nie tak z Twoim podejściem do stylów.

```python
"""Diagnostyka: ile unikalnych stylow naprawde jest w pliku."""

import zipfile
from pathlib import Path
from xml.etree import ElementTree

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def policz_styles(sciezka: Path) -> dict[str, int]:
    """Zwraca liczbe unikalnych wpisow w kluczowych sekcjach styles.xml."""
    with zipfile.ZipFile(sciezka) as archiwum:
        xml = archiwum.read("xl/styles.xml")
    korzen = ElementTree.fromstring(xml)

    wynik = {}
    for sekcja in ("fonts", "fills", "borders", "cellXfs", "numFmts"):
        element = korzen.find(f"m:{sekcja}", NS)
        wynik[sekcja] = int(element.get("count", len(element))) if element is not None else 0
    return wynik


if __name__ == "__main__":
    for plik in sorted(Path("output").glob("09_*.xlsx")):
        print(f"{plik.name:<32} {policz_styles(plik)}")
```

Uruchomienie tego na własnych plikach jest pouczające — zwłaszcza na tych, w których **nie** przejmowałeś się rejestrem stylów.

**Dwa dodatkowe, „za darmo" przyspieszenia:**

1. **Przypisuj cały styl naraz, nie cecha po cesze.** `komorka.font = f; komorka.fill = p; komorka.alignment = a` to trzy operacje; openpyxl przy zapisie i tak zbuduje z nich jeden `<xf>`. Nie da się tego „zoptymalizować" w inny sposób niż konsekwencja — ale warto wiedzieć, że nie ma tu ukrytego narzutu.
2. **Nie formatuj komórek, które nie istnieją.** Ustawienie `column_dimensions['A'].font` **nie uformatuje** istniejących komórek (patrz 2.14). Więc „formatowanie całej kolumny" w openpyxl znaczy zawsze: pętla po komórkach.

### 2.14. Dwa ograniczenia, o których trzeba wiedzieć od razu

#### 2.14.1. Formatowanie całej kolumny/wiersza nie działa na istniejące komórki

To ograniczenie **samego formatu pliku**, nie openpyxl. Dokumentacja mówi:

> „Style mogą być również stosowane do kolumn i wierszy, ale **dotyczy to wyłącznie komórek utworzonych (w Excelu) po zamknięciu pliku**. Jeśli chcesz zastosować styl do całych wierszy i kolumn, musisz zastosować styl do każdej komórki indywidualnie. To ograniczenie formatu pliku."

Co to znaczy w praktyce:

```python
# ❌ NIE zadziala na istniejace komorki
ws.column_dimensions["A"].font = Font(bold=True)

# ✅ Zadziala - petla po komorkach
for wiersz in range(1, ws.max_row + 1):
    ws.cell(row=wiersz, column=1).font = FONT_POGRUBIONY
```

Kod z komentarzem „❌" jest **poprawny składniowo**, nie zgłosi błędu, a nawet coś zapisze do pliku. Ale Twoje istniejące komórki w kolumnie A pozostaną niepogrubione. **W Excelu wygląda to tak, jakby nic się nie stało** — bo obiekt kolumny zapamiętał domyślny styl dla *przyszłych* komórek.

To jest doskonały przykład kategorii „openpyxl potrafi tylko częściowo" z modułu 03: API istnieje, wywołanie przechodzi, ale efekt jest inny niż intuicja.

#### 2.14.2. Scalone komórki dziedziczą format z lewego górnego rogu

Wspominaliśmy o tym przy obramowaniach. Rozszerzmy: **wartość i format scalonego zakresu są zdefiniowane w komórce lewego górnego rogu.**

```python
ws.merge_cells("B2:F4")
top_left = ws["B2"]
top_left.alignment = Alignment(horizontal="center", vertical="center")
top_left.fill = PatternFill(fill_type="solid", fgColor="FFDDDDDD")
# Komorki C2, D2, ... sa czescia scalenia, ale nie ustawiaj im stylu -
# liczy sie tylko B2.
```

Dokumentacja podkreśla, że formatowanie dla scalonego zakresu jest generowane **na potrzeby zapisu** — czyli openpyxl wygeneruje odpowiednie wpisy w pliku na podstawie komórki lewej górnej. Dlatego przy scaleniach nie powielaj stylów na wszystkich komórkach zakresu.

## 3. Przykłady krok po kroku

### Przykład 1 — dowód eksperymentalny: dlaczego `cell.font.bold = True` nie działa

Zaczynamy od dowodu, bo wiedza „nie działa" bez zrozumienia „dlaczego" jest bezużyteczna.

```python
"""Dowod eksperymentalny: niemutowalnosc stylow w openpyxl.

Uruchom:  python examples/09_dowod_niemutowalnosci.py
"""

from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Nagłówek"

    print("=" * 78)
    print("EKSPERYMENT 1: mutacja obiektu stylu 'w miejscu'")
    print("=" * 78)

    ws["A1"].font = Font(size=14, color="FF2F5597")
    print(f"  typ zwracany przez cell.font: {type(ws['A1'].font).__name__}")
    print(f"  czy to Font?                 : {isinstance(ws['A1'].font, Font)}")

    try:
        ws["A1"].font.bold = True
        print("  => UDALO SIE (nieoczekiwane!)")
    except AttributeError as blad:
        print(f"  => AttributeError: {blad}")

    print(f"  czy bold ustawil sie?        : {ws['A1'].font.bold!r}")

    print()
    print("=" * 78)
    print("EKSPERYMENT 2: wspoldzielenie - dlaczego blokada ma sens")
    print("=" * 78)

    wspolny = Font(size=14, color="FF2F5597", bold=False)
    for kolumna in ("A", "B", "C"):
        ws[f"{kolumna}2"] = f"komórka {kolumna}2"
        ws[f"{kolumna}2"].font = wspolny

    print(f"  czy wszystkie 3 komorki maja TEN SAM obiekt Font?")
    a2, b2, c2 = ws["A2"].font, ws["B2"].font, ws["C2"].font
    print(f"    A2 == B2 == C2 (porownanie wartosci): {a2 == b2 == c2}")
    print(f"    A2 is B2 (ta sama tozsamosc):         {a2 is b2}")

    print()
    print("  Gdyby mutacja byla dozwolona, ponizsze wywolanie zmieniloby")
    print("  wyglad WSZYSTKICH TRZECH komorek - i nikt by tego nie zauwazyl:")
    print()
    print("      ws['A2'].font.bold = True   # <- na szczescie niemozliwe")
    print()
    print("  To jest wlasnie ten 'niepozadany efekt uboczny', przed ktorym")
    print("  chroni openpyxl. Koszt: musisz zrobic kopie.")

    print()
    print("=" * 78)
    print("EKSPERYMENT 3: trzy poprawne naprawy")
    print("=" * 78)

    # --- Naprawa A: copy() ze standardowej biblioteki --------------------
    ws["D2"] = "copy()"
    ws["D2"].font = Font(size=14, color="FF2F5597")
    nowy = copy(ws["D2"].font)
    nowy.bold = True
    ws["D2"].font = nowy
    print(f"  copy()            -> typ kopii: {type(nowy).__name__}, "
          f"bold={ws['D2'].font.bold}, size={ws['D2'].font.size} (zachowany)")

    # --- Naprawa B: operator + (galaz 3.1) ------------------------------
    ws["E2"] = "operator +"
    ws["E2"].font = Font(size=14, color="FF2F5597")
    try:
        ws["E2"].font = ws["E2"].font + Font(bold=True)
        print(f"  operator +        -> bold={ws['E2'].font.bold}, "
              f"size={ws['E2'].font.size} (zachowany)")
    except TypeError as blad:
        print(f"  operator +        -> niedostepny w tej wersji: {blad}")

    # --- Naprawa C: nowy, pelny obiekt ----------------------------------
    ws["F2"] = "nowy obiekt"
    ws["F2"].font = Font(size=14, color="FF2F5597", bold=True)
    print(f"  nowy obiekt       -> bold={ws['F2'].font.bold}, "
          f"size={ws['F2'].font.size}")

    print()
    print("=" * 78)
    print("EKSPERYMENT 4: CICHA UTRATA DANYCH - najgrozniejszy blad")
    print("=" * 78)

    ws["A5"] = "ostrzeżenie"
    ws["A5"].font = Font(size=18, italic=True, color="FFC00000")
    print("  ustawiono: size=18, italic=True, color=FFC00000")

    ws["A5"].font = Font(bold=True)      # <- nowy Font, wszystko inne None
    f = ws["A5"].font
    print(f"  po przypisaniu Font(bold=True):")
    print(f"    bold   = {f.bold!r}")
    print(f"    size   = {f.size!r}      <- przepadlo!")
    print(f"    italic = {f.italic!r}      <- przepadlo!")
    print(f"    color  = {f.color}        <- przepadlo!")
    print()
    print("  Ten kod NIE podnosi zadnego wyjatku. Po prostu gubi dane.")

    print()
    print("=" * 78)
    print("WNIOSEK")
    print("=" * 78)
    print("  1. cell.font zwraca StyleProxy - pelnomocnika tylko do odczytu.")
    print("  2. Blokada zapisu chroni przed globalnym efektem ubocznym.")
    print("  3. Chcesz zmienic jedna ceche -> copy(), operator + albo nowy obiekt.")
    print("  4. Nowy obiekt 'na czysto' ZASTEPUJE wszystko - to cicha utrata danych.")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `ws["A1"].font` tworzy **za każdym razem nowy obiekt `StyleProxy`** opakowujący ten sam obiekt `Font`. Zwróć uwagę: `a2 is b2` zwróci `False`, mimo że to ten sam styl docelowy — bo to dwa różne pełnomocnicy. Ale `a2 == b2` zwróci `True`, bo porównanie jest delegowane do obiektu docelowego. To ważne przy testach: **porównuj style operatorem `==`, nigdy `is`.**

**Co trafi do pliku.** Z czterech eksperymentów do pliku trafiłyby: jeden styl nagłówka z Eksperymentu 1, jeden wspólny styl z trzech komórek, trzy style z naprawami i jeden „zubożony" styl z Eksperymentu 4. Ten ostatni jest dowodem, że plik będzie **poprawny formalnie, a błędny merytorycznie**.

**Czego ten przykład uczy najbardziej:** że **niemutowalność stylów jest kontraktem, a nie niedogodnością** — i że jedyne realne ryzyko nie leży w blokadzie (`AttributeError` Cię ochroni), ale w *cichej* zamianie stylu na nowy, ubogi obiekt. Tam nie ma żadnego wyjątku, który by Cię ostrzegł.

### Przykład 2 — pełny arkusz raportu: sześć elementów stylu w akcji

Teraz zbudujemy arkusz, w którym wykorzystamy **wszystkie sześć elementów** stylu — i zobaczymy, jak wygląda profesjonalny wynik.

```python
"""Pelny arkusz raportu - wszystkie szesc elementow stylu.

Uruchom:  python examples/09_pelny_styl.py
"""

from __future__ import annotations

from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "09_pelny_styl.xlsx"

# ---------------------------------------------------------------- paleta
GRANAT = "FF2F5597"
GRANAT_CIEMNY = "FF1F3864"
SZARY_TLO = "FFF2F2F2"
SZARY_LINIA = "FFBFBFBF"
CZERWIEN = "FFC00000"
BIEL = "FFFFFFFF"

# ------------------------------------------------- obiekty wspoldzielone
LINIA = Side(style="thin", color=SZARY_LINIA)
LINIA_GORA = Side(style="medium", color=GRANAT)

FONT_TYTUL = Font(name="Calibri", size=16, bold=True, color=GRANAT_CIEMNY)
FONT_NAGLOWEK = Font(name="Calibri", size=11, bold=True, color=BIEL)
FONT_ALARM = Font(name="Calibri", size=11, bold=True, color=CZERWIEN)

FILL_NAGLOWEK = PatternFill(fill_type="solid", fgColor=GRANAT)
FILL_ZEBRA = PatternFill(fill_type="solid", fgColor=SZARY_TLO)
FILL_BRAK = PatternFill(fill_type=None)

ALIGN_TYTUL = Alignment(horizontal="left", vertical="center")
ALIGN_NAGLOWEK = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LICZBA = Alignment(horizontal="right", vertical="center")
ALIGN_TEKST = Alignment(horizontal="left", vertical="center")

RAMKA_NAGLOWKA = Border(left=LINIA, right=LINIA, top=LINIA_GORA, bottom=LINIA)
RAMKA_DANYCH = Border(left=LINIA, right=LINIA, top=LINIA, bottom=LINIA)

DANE = [
    ("Styczeń", 120_000, 78_000, 0.350),
    ("Luty", 98_500, 71_200, 0.277),
    ("Marzec", 143_700, 88_400, 0.385),
    ("Kwiecień", 131_200, 82_900, 0.368),
    ("Maj", 88_000, 74_500, 0.153),      # <- niska marza
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Raport"

    # --- wiersz 1: tytul (scalony, cztery kolumny) ----------------------
    ws.merge_cells("A1:D1")
    tytul = ws["A1"]
    tytul.value = "Sprzedaż i marża — I kwartał 2026"
    tytul.font = FONT_TYTUL
    tytul.alignment = ALIGN_TYTUL
    ws.row_dimensions[1].height = 28      # <- sztywna wysokosc, ale bez zawijania

    # --- wiersz 2: naglowki ---------------------------------------------
    naglowki = ["Miesiąc", "Sprzedaż", "Koszt", "Marża"]
    for indeks, tekst in enumerate(naglowki, start=1):
        komorka = ws.cell(row=2, column=indeks, value=tekst)
        komorka.font = FONT_NAGLOWEK
        komorka.fill = FILL_NAGLOWEK
        komorka.alignment = ALIGN_NAGLOWEK
        komorka.border = RAMKA_NAGLOWKA
    ws.row_dimensions[2].height = 22

    # --- wiersze 3..7: dane ---------------------------------------------
    for przesuniecie, (miesiac, sprzedaz, koszt, marza) in enumerate(DANE):
        wiersz = 3 + przesuniecie
        zebra = FILL_ZEBRA if przesuniecie % 2 == 0 else FILL_BRAK

        for kolumna, wartosc in enumerate((miesiac, sprzedaz, koszt, marza), start=1):
            komorka = ws.cell(row=wiersz, column=kolumna, value=wartosc)
            komorka.border = RAMKA_DANYCH
            komorka.fill = zebra

            if kolumna == 1:
                komorka.alignment = ALIGN_TEKST
            elif kolumna == 4:
                komorka.alignment = ALIGN_LICZBA
                komorka.number_format = "0.0%"          # 6. element stylu!
                if wartosc < 0.20:
                    komorka.font = FONT_ALARM         # reczne "alarmowanie" - modul 13 zrobi to regulami
            else:
                komorka.alignment = ALIGN_LICZBA
                komorka.number_format = '#,##0 "zł"'

    # --- wiersz 9: podsumowanie -----------------------------------------
    ws["A9"] = "Razem"
    ws["A9"].font = Font(bold=True)
    ws["B9"] = f"=SUM(B3:B{2 + len(DANE)})"
    ws["B9"].font = Font(bold=True)
    ws["B9"].number_format = '#,##0 "zł"'
    ws["B9"].alignment = ALIGN_LICZBA
    ws["B9"].border = Border(top=LINIA)

    # --- kolumny: szerokosci --------------------------------------------
    for kolumna, szerokosc in zip("ABCD", (16, 14, 14, 10)):
        ws.column_dimensions[kolumna].width = szerokosc

    # --- pole formularzowe: odblokowane do wpisywania -------------------
    ws["F2"] = "Komentarz analityka:"
    ws["F2"].font = Font(italic=True)
    ws["F3"].protection = Protection(locked=False)     # <- przygotowanie do modulu 17
    ws["F3"].border = Border(bottom=LINIA)
    ws["F3"].alignment = Alignment(vertical="center")

    ws.freeze_panes = "A3"

    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    """Weryfikacja: czy styl naprawde jest w pliku (a nie tylko w pamieci)."""
    wb = load_workbook(cel)
    try:
        ws = wb["Raport"]
        print("=" * 86)
        print("WERYFIKACJA PO ZAPISIE I ODCZYCIE")
        print("=" * 86)

        pola = [
            ("A1  font.bold",           ws["A1"].font.bold),
            ("A1  font.size",           ws["A1"].font.size),
            ("A2  font.color.rgb",      ws["A2"].font.color.rgb),
            ("A2  fill.fgColor.rgb",    ws["A2"].fill.fgColor.rgb),
            ("A2  alignment.wrap_text", ws["A2"].alignment.wrap_text),
            ("A2  border.top.style",    ws["A2"].border.top.style),
            ("D3  number_format",       ws["D3"].number_format),
            ("B3  number_format",       ws["B3"].number_format),
            ("A3  border.left.style",   ws["A3"].border.left.style),
            ("A3  fill.fgColor.rgb",    ws["A3"].fill.fgColor.rgb),
            ("A4  fill.fgColor.rgb",    ws["A4"].fill.fgColor.rgb),
            ("F3  protection.locked",   ws["F3"].protection.locked),
            ("D7  font.color.rgb",      ws["D7"].font.color.rgb),
        ]
        for opis, wartosc in pola:
            print(f"  {opis:<28} = {wartosc!r}")

        print()
        print("  Uwaga: przez zapis i odczyt styl przetrwal w calosci.")
        print("  Jedyny wyjatek: F3.protection.locked=False to zapowiedz ochrony")
        print("  arkusza - sam z siebie nie robi nic, dopoki nie wlaczysz")
        print("  ws.protection.sheet (modul 17).")
    finally:
        wb.close()


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Kluczowe fragmenty do przemyślenia:**

1. **Wszystkie obiekty stylu są zdefiniowane na poziomie modułu** — przed funkcją, nie w pętli. To jest wzorzec z 2.12 i 2.13 w najprostszej formie.
2. **Zebra używa dwóch gotowych wypełnień**, a nie tworzy nowego per wiersz. `FILL_BRAK` to `PatternFill(fill_type=None)` — jawne „brak wypełnienia", a nie `None` w komórce. To subtelność: przypisanie `PatternFill(fill_type=None)` **czyści** wypełnienie, co jest przydatne, jeśli komórka mogła je wcześniej mieć.
3. **`number_format` jest ustawiany na komórkach danych** — i to jest to, o czym mówi moduł 10: liczba `0.35` z formatem `0.0%` wyświetli się jako `35,0%`, a wartość w pliku pozostanie `0.35`.
4. **`ws.row_dimensions[1].height = 28` działa**, bo tytuł nie ma `wrap_text=True`. Gdyby miał, ta wysokość mogłaby przyciąć dłuższy tekst.
5. **Verification po zapisie i odczycie** jest kluczowa i wraca w całym kursie: styl, którego nie sprawdziłeś po odczycie, jest stylem, o którym nic nie wiesz.

**Co trafi do pliku.** W `xl/styles.xml` znajdziesz prawdopodobnie: 4 wpisy w `<fonts>` (domyślny, tytuł, nagłówek, alarm + pogrubienie podsumowania = 5), 3–4 w `<fills>` (w tym dwa zarezerwowane), 3 wpisy w `<borders>`, oraz kilka własnych `<numFmt>` z identyfikatorami od 164. **To jest cały koszt stylowania tego arkusza — kilkanaście wpisów, nie tysiące.**

### Przykład 3 — rejestr w wersji z `NamedStyle`: porównanie dwóch podejść

Zbudujemy teraz wersję profesjonalną z `NamedStyle` i porównamy ją z podejściem bezpośrednim.

```python
"""NamedStyle vs style komorek - porownanie i pulapka braku propagacji.

Uruchom:  python examples/09_named_style_vs_cell.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "09_named_style.xlsx"


def zarejestruj_styl(wb, nazwa: str):
    """Tworzy i rejestruje styl nazwany w skoroszycie."""
    if nazwa in {s for s in wb.named_styles}:      # API publiczne (3.1+)
        return wb[nazwa] if hasattr(wb, "__getitem__") else None
    styl = NamedStyle(name=nazwa)
    styl.font = Font(name="Calibri", size=20, bold=True, color="FFFFFFFF")
    styl.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
    styl.alignment = Alignment(horizontal="center", vertical="center")
    styl.number_format = "#,##0"
    wb.add_named_style(styl)
    return styl


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "KPI"

    print("=" * 84)
    print("1. REJESTRACJA STYLU NAZWANEGO")
    print("=" * 84)
    kpi = zarejestruj_styl(wb, "KPI")
    print(f"  styl 'KPI' zarejestrowany")
    print(f"  dostepne style nazwane: {sorted(wb.named_styles)}")

    # --- przypisujemy do trzech komorek ---------------------------------
    ws["A1"] = "Przychód"
    ws["B1"] = "Marża"
    ws["C1"] = "Sztuki"
    for kolumna in ("A", "B", "C"):
        ws[f"{kolumna}1"].style = "KPI"

    print()
    print("=" * 84)
    print("2. PULAPKA: ZMIANA STYLU NIE PROPAGUJE SIE NA KOMORKI")
    print("=" * 84)
    print("  Zmieniamy teraz definicje stylu 'KPI' na rozmiar 30, kursywa.")
    kpi.font = Font(name="Calibri", size=30, italic=True, color="FFC00000")

    ws2 = wb.create_sheet("Po zmianie")
    ws2["A1"] = "Nowa komórka po zmianie stylu"
    ws2["A1"].style = "KPI"          # <- ta DOSTANIE nowa definicje

    print(f"  A1 (przypisana PRZED zmiana)  font.size = {ws['A1'].font.size}")
    print(f"  A1 (przypisana PRZED zmiana)  font.italic = {ws['A1'].font.italic}")
    print(f"  A1 (przypisana PO zmianie)    font.size = {ws2['A1'].font.size}")
    print(f"  A1 (przypisana PO zmianie)    font.italic = {ws2['A1'].font.italic}")
    print()
    print("  Wniosek: komorka zapamietuje ODcisk z momentu przypisania.")
    print("  Zmiana definicji stylu dotyczy tylko komorek przypisanych PO niej.")

    print()
    print("=" * 84)
    print("3. KONTROLA PO ZAPISIE I ODCZYCIE")
    print("=" * 84)
    wb.save(CEL)
    wb.close()

    wb2 = load_workbook(CEL)
    try:
        ws_a = wb2["KPI"]
        ws_b = wb2["Po zmianie"]
        print(f"  po odczycie: 'KPI'!A1 font.size   = {ws_a['A1'].font.size}")
        print(f"  po odczycie: 'KPI'!A1 fill.fgColor = {ws_a['A1'].fill.fgColor.rgb}")
        print(f"  po odczycie: 'Po zmianie'!A1 size = {ws_b['A1'].font.size}")
        print(f"  dostepne style nazwane w pliku    = {sorted(wb2.named_styles)}")
    finally:
        wb2.close()

    print()
    print("=" * 84)
    print("4. POROWNANIE: STYL KOMORKI vs STYL NAZWANY")
    print("=" * 84)
    tabela = [
        ("Mutowalny?",                "NIE (StyleProxy)",        "TAK"),
        ("Ma nazwe?",                 "NIE",                     "TAK, w skoroszycie"),
        ("Jak stosowac",              "cell.font = Font(...)",   "cell.style = 'Nazwa'"),
        ("Zmiana definicji dziala?",  "N/D",                     "tylko na NOWE przypisania"),
        ("Kiedy uzywac",              "styl jednej komorki",     "powtarzalne wzorce"),
        ("Ryzyko",                    "cicha utrata przy nadpisaniu", "brak propagacji"),
    ]
    print(f"  {'cecha':<30} {'styl komorki':<26} {'styl nazwany'}")
    print(f"  {'-' * 30} {'-' * 26} {'-' * 24}")
    for cecha, a, b in tabela:
        print(f"  {cecha:<30} {a:<26} {b}")

    print()
    print("  Rekomendacja: STYL NAZWANY dla powtarzalnych wzorcow raportowych")
    print("  (naglowek, KPI, waluta), STYL KOMORKI dla pojedynczych przypadkow.")


if __name__ == "__main__":
    main()
```

⚠️ **Ważna uwaga o `wb.named_styles`.** To API pojawiło się w gałęzi 3.1 i jest publicznym odpowiednikiem prywatnego `wb._named_styles`. **Sprawdź w dokumentacji swojej wersji**, czy `wb.named_styles` istnieje i co dokładnie zwraca (u mnie zwraca listę nazw). Jeśli nie — użyj `wb._named_styles` w kodzie diagnostycznym, ale wiedz, że to API prywatne.

**Czego ten przykład uczy najbardziej:** że `NamedStyle` **nie jest** dynamicznym powiązaniem, tylko sposobem na **powtarzalne przypisanie tego samego zestawu wartości**. Zaleta polega na tym, że *definiujesz go raz*, a nie na tym, że *zmienia się globalnie*. To dwie różne obietnice i tylko jedna z nich jest prawdziwa.

### Przykład 4 — `NamedStyle` „KPI" na 1000 komórkach z pomiarem czasu

Zadanie z briefu tego modułu — zrobione porządnie, z porównaniem trzech strategii.

```python
"""NamedStyle na 1000 komorek - pomiar czasu trzech strategii.

Uruchom:  python examples/09_kpi_1000.py
"""

from __future__ import annotations

import time
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

LICZBA_KOMOREK = 1000
KOLUMNY = 10
WIERSZE = 100


def zmierz(funkcja, *argumenty):
    """Uruchamia funkcje i zwraca (wynik, czas w sekundach)."""
    start = time.perf_counter()
    wynik = funkcja(*argumenty)
    return wynik, time.perf_counter() - start


# ----------------------------------------------------------------------
# STRATEGIA A: nowy NamedStyle dla KAZDEJ komorki (najgorsza mozliwa)
# ----------------------------------------------------------------------
def strategia_a_nowy_styl_za_kazdym_razem(cel: Path) -> dict:
    wb = Workbook()
    ws = wb.active
    ws.title = "KPI"

    for indeks in range(LICZBA_KOMOREK):
        nazwa = f"KPI_{indeks}"
        styl = NamedStyle(name=nazwa)
        styl.font = Font(size=20, bold=True, color="FFFFFFFF")
        styl.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
        styl.alignment = Alignment(horizontal="center", vertical="center")
        styl.number_format = "#,##0"
        wb.add_named_style(styl)

        wiersz = indeks // KOLUMNY + 1
        kolumna = indeks % KOLUMNY + 1
        komorka = ws.cell(row=wiersz, column=kolumna, value=indeks)
        komorka.style = nazwa

    sciezka = OUTPUT / f"09_kpi_{cel}.xlsx"
    wb.save(sciezka)
    wb.close()
    return {"plik": sciezka.name, "stylow": LICZBA_KOMOREK}


# ----------------------------------------------------------------------
# STRATEGIA B: NOWY OBIEKT Font/Fill per komorka (bez NamedStyle)
# ----------------------------------------------------------------------
def strategia_b_nowy_obiekt_za_kazdym_razem(cel: Path) -> dict:
    wb = Workbook()
    ws = wb.active
    ws.title = "KPI"

    for indeks in range(LICZBA_KOMOREK):
        wiersz = indeks // KOLUMNY + 1
        kolumna = indeks % KOLUMNY + 1
        komorka = ws.cell(row=wiersz, column=kolumna, value=indeks)
        komorka.font = Font(size=20, bold=True, color="FFFFFFFF")
        komorka.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
        komorka.alignment = Alignment(horizontal="center", vertical="center")
        komorka.number_format = "#,##0"

    sciezka = OUTPUT / f"09_kpi_{cel}.xlsx"
    wb.save(sciezka)
    wb.close()
    return {"plik": sciezka.name, "stylow": 1}


# ----------------------------------------------------------------------
# STRATEGIA C: JEDEN wspolny obiekt stylu na wszystkie komorki (wzorcowa)
# ----------------------------------------------------------------------
def strategia_c_jeden_wspolny_styl(cel: Path) -> dict:
    wb = Workbook()
    ws = wb.active
    ws.title = "KPI"

    # Pieczatki tworzone RAZ - przed petla.
    font = Font(size=20, bold=True, color="FFFFFFFF")
    fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
    alignment = Alignment(horizontal="center", vertical="center")

    for indeks in range(LICZBA_KOMOREK):
        wiersz = indeks // KOLUMNY + 1
        kolumna = indeks % KOLUMNY + 1
        komorka = ws.cell(row=wiersz, column=kolumna, value=indeks)
        komorka.font = font
        komorka.fill = fill
        komorka.alignment = alignment
        komorka.number_format = "#,##0"

    sciezka = OUTPUT / f"09_kpi_{cel}.xlsx"
    wb.save(sciezka)
    wb.close()
    return {"plik": sciezka.name, "stylow": 1}


def main() -> None:
    print("=" * 88)
    print(f"POMIAR CZASU: {LICZBA_KOMOREK} komorek, trzy strategie")
    print("=" * 88)

    wyniki = []

    _, czas_a = zmierz(strategia_a_nowy_styl_za_kazdym_razem, "a")
    wyniki.append(("A: NamedStyle per komorka", czas_a))

    _, czas_b = zmierz(strategia_b_nowy_obiekt_za_kazdym_razem, "b")
    wyniki.append(("B: nowy Font/Fill per komorka", czas_b))

    _, czas_c = zmierz(strategia_c_jeden_wspolny_styl, "c")
    wyniki.append(("C: jeden wspolny styl (wzorcowa)", czas_c))

    najszybszy = min(czas for _, czas in wyniki)
    for opis, czas in wyniki:
        wzgledem = czas / najszybszy
        print(f"  {opis:<36} {czas:6.3f} s   ({wzgledem:5.2f}x)")

    print()
    print("=" * 88)
    print("WNIOSKI")
    print("=" * 88)
    print("  1. Strategia A (NamedStyle per komorka) jest najgorsza z dwoch")
    print("     powodow: tworzy 1000 obiektow ORAZ 1000 nazw stylow w pliku.")
    print("  2. Strategia B tworzy 1000 obiektow w Pythonie, ale plik ma ich 1")
    print("     - bo openpyxl deduplikuje identyczne style przy zapisie.")
    print("  3. Strategia C tworzy 4 obiekty w Pythonie i 1 zestaw w pliku.")
    print()
    print("  Wniosek praktyczny: tworz obiekty stylu RAZ, poza petla.")
    print("  Roznica w czasie rosnie liniowo z liczba komorek.")
    print()
    print("  Zajrzyj tez do plikow w output/ - rozpakuj je jako ZIP")
    print("  i porownaj liczbe wpisow w <fonts> i <fills> w xl/styles.xml.")


if __name__ == "__main__":
    main()
```

**Czego ten przykład uczy najbardziej:** że różnica między „dobrym a złym" podejściem do stylów nie jest estetyczna. **Jest mierzalna.** I rośnie liniowo z liczbą komórek.

A co ważniejsze: strategia A nie tylko jest wolniejsza — ona **produkuje zły plik**. `styles.xml` z tysiącem nazw stylów jest wolniejszy do otwarcia w Excelu i trudniejszy w utrzymaniu, bo nikt nie zrozumie, po co tam są `KPI_0`, `KPI_1`… `KPI_999`.

**Uwaga o powtarzalności pomiarów:** czasy mogą się różnić między maszynami i wersjami Pythona. Ważny jest **stosunek** czasów, nie wartości bezwzględne. Uruchom to u siebie (i najlepiej kilka razy — pierwszy przebieg może być wolniejszy z powodu ładowania modułów).

### Przykład 5 — co dokładnie trafiło do pliku: diagnostyka `styles.xml`

Domykamy moduł weryfikacją na poziomie pakietu. To ćwiczenie dla Ciebie, a poniżej gotowy skrypt.

```python
"""Diagnostyka stylow na poziomie pakietu .xlsx.

Uruchom:  python examples/09_diagnostyka_styles.py
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def policz_sekcje(sciezka: Path) -> None:
    """Wypisuje liczbe wpisow w kluczowych sekcjach styles.xml."""
    with zipfile.ZipFile(sciezka) as archiwum:
        try:
            xml = archiwum.read("xl/styles.xml")
        except KeyError:
            print(f"  {sciezka.name}: brak pliku xl/styles.xml")
            return
        czesci = set(archiwum.namelist())

    korzen = ElementTree.fromstring(xml)

    print(f"  {sciezka.name}")
    for sekcja in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs"):
        element = korzen.find(f"m:{sekcja}", NS)
        if element is None:
            print(f"    {sekcja:<14} brak sekcji")
            continue
        zadeklarowane = element.get("count", "?")
        faktycznie = len(list(element))
        zgodne = "OK" if str(faktycznie) == zadeklarowane else "NIEZGODNE"
        print(f"    {sekcja:<14} zadeklarowane={zadeklarowane:<5} "
              f"faktycznie={faktycznie:<5} {zgodne}")

    # Wlasne formaty liczb (numFmtId >= 164) - zapowiedz modulu 10
    num_fmts = korzen.find("m:numFmts", NS)
    if num_fmts is not None:
        wlasne = [
            f.get("numFmtId") for f in num_fmts if int(f.get("numFmtId", "0")) >= 164
        ]
        if wlasne:
            print(f"    wlasne numFmtId (>=164): {wlasne}")

    # Media i rysunki - dla kontekstu z modulu 08
    ma_media = any(n.startswith("xl/media/") for n in czesci)
    ma_drawings = any(n.startswith("xl/drawings/") for n in czesci)
    print(f"    xl/media: {ma_media}, xl/drawings: {ma_drawings}")


def main() -> None:
    print("=" * 88)
    print("DIAGNOSTYKA STYLOW W PLIKACH Z KATALOGU output/")
    print("=" * 88)
    print()

    pliki = sorted(OUTPUT.glob("09_*.xlsx"))
    if not pliki:
        print("  Brak plikow. Uruchom najpierw inne przyklady z tego modulu.")
        return

    for plik in pliki:
        policz_sekcje(plik)
        print()

    print("=" * 88)
    print("JAK TO CZYTAC")
    print("=" * 88)
    print("  * cellXfs    = liczba unikalnych ZESTAWOW stylu (to widzi komorka)")
    print("  * fonts/fills/borders = liczba unikalnych POJEDYNCZYCH pieczatek")
    print("  * pierwsze dwa wpisy w <fills> sa zarezerwowane przez format:")
    print("      0 = brak wypelnienia, 1 = wzor gray125")
    print("  * numFmtId >= 164 => wlasny format liczby (modul 10)")


if __name__ == "__main__":
    main()
```

**Czego ten przykład uczy najbardziej:** że **liczby w pliku są sprawdzalne**. Nie musisz zgadywać, czy Twoje zarządzanie stylami jest efektywne — możesz je policzyć, rozpakowując plik i czytając trzy atrybuty. To ta sama umiejętność, którą w module 18 wykorzystamy do wykrywania utraty danych.

## 4. Anatomia API

| Klasa / metoda / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Font(...)` | Obiekt czcionki | `name`, `size`/`sz`, `bold`/`b`, `italic`/`i`, `underline`/`u`, `strike`/`strikethrough`, `color`, `vertAlign`, `scheme`, `family`, `charset`, `outline`, `shadow`, `condense`, `extend` | Używaj długich nazw; `scheme` może wygrać z `name` |
| `PatternFill(...)` | Wypełnienie | `fill_type`/`patternType`, `fgColor`/`start_color`, `bgColor`/`end_color` | **`solid` używa `fgColor`**; brak `fill_type` = brak wypełnienia |
| `Border(...)` | Obramowanie komórki | `left`, `right`, `top`, `bottom`, `diagonal`, `diagonal_direction`, `outline`, `vertical`, `horizontal` | Dla komórek praktycznie tylko cztery pierwsze |
| `Side(...)` | Jedna krawędź | `style`/`border_style`, `color` | `style=None` = brak linii |
| `Alignment(...)` | Wyrównanie | `horizontal`, `vertical`, `wrap_text`, `text_rotation`, `indent`, `shrink_to_fit`, `justifyLastLine`, `readingOrder` | `wrap_text` i `shrink_to_fit` wykluczają się |
| `Protection(...)` | Ochrona **komórki** | `locked`, `hidden` | Działa tylko przy włączonej ochronie arkusza (moduł 17) |
| `Color(...)` | Kolor | `rgb` (`"FF0000"` lub `"FFFF0000"`), `indexed`, `theme`, `tint`, `auto` | 6 znaków → openpyxl dopisze `"00"`; `"#FF0000"` → `ValueError` |
| `NamedStyle(...)` | Styl nazwany | `name`, a potem właściwości `font`, `fill`, `border`, `alignment`, `protection`, `number_format` | **Mutowalny**; komórka nie śledzi późniejszych zmian |
| `wb.add_named_style(styl)` | Rejestruje styl nazwany | obiekt `NamedStyle` | Nazwa musi być unikalna w skoroszycie |
| `wb.named_styles` | Nazwy zarejestrowanych stylów | — | API 3.1+; sprawdź swoją wersję |
| `wb._named_styles` | To samo, **API prywatne** | — | Tylko do diagnostyki |
| `cell.font` | Zwraca `StyleProxy` czcionki | — | Odczyt tak, zapis → `AttributeError` |
| `cell.fill`, `cell.border`, `cell.alignment`, `cell.protection` | Jak wyżej | — | Wszystkie zwracają `StyleProxy` |
| `cell.number_format` | Format liczby (napisy) | `str` | **Nie jest** obiektem stylu; szczegóły w module 10 |
| `cell.style` | Styl komórki — nazwa lub obiekt | `str` (nazwa `NamedStyle`) albo `NamedStyle` | Przypisanie zapisuje **wartości**, nie powiązanie |
| `copy(cell.font)` | Kopia jako zwykły obiekt | — | **Metoda kanoniczna** |
| `cell.font + Font(...)` | Scalanie stylów operatorem `+` | `Font`/`Fill`/… | Gałąź 3.1+; sprawdź w swojej wersji |
| `StyleProxy.copy(**kw)` | Kopia z nadpisaniem atrybutów | dowolne `kw` | ⚠️ **Przestarzałe** w 3.1.x — użyj `copy()` |
| `ws.merge_cells(zakres)` | Scalanie | `zakres: str` lub `(min_row, min_col, max_row, max_col)` | Format bierze się z lewej górnej komórki |
| `ws.column_dimensions["A"].font` | Styl kolumny | obiekt stylu | **Nie działa na istniejące komórki** — ograniczenie formatu |
| `ws.row_dimensions[1].height` | Wysokość wiersza | liczba (punkty) | Ustawiona → tekst zawijany może zostać przycięty |
| `ws.freeze_panes` | Zamrożenie okien | `"A3"` itp. | Pełne omówienie w module 11 |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — kolorowy, pogrubiony i obramowany nagłówek.**

Napisz `examples/09_cw1.py`, który:

1. Tworzy skoroszyt z arkuszem `Raport`.
2. Wpisuje w wierszu 1 nagłówki: `Miesiąc`, `Sprzedaż`, `Koszt`, `Marża`.
3. Nadaje tym czterem komórkom **jeden, wspólny** styl złożony z:
   - czcionki: `Calibri`, rozmiar `11`, pogrubiona, kolor biały,
   - wypełnienia jednolitego w kolorze granatowym (`2F5597`),
   - wyrównania: poziomo i pionowo `center`, z zawijaniem tekstu,
   - obramowania: cienka linia w kolorze `BFBFBF` z wszystkich czterech stron.
4. Wypełnia wiersze 2–5 danymi: cztery miesiące, sprzedaż, koszt, marża (marża jako ułamek, np. `0.35`).
5. Formatuje kolumnę `D` jako procent (`"0.0%"`), a kolumny `B` i `C` jako walutę (`'#,##0 "zł"'`).
6. Ustawia szerokości kolumn: `A` = 16, `B` = 14, `C` = 14, `D` = 10.
7. Zapisuje plik jako `output/09_cw1.xlsx`, a następnie **wczytuje go ponownie** i wypisuje dla `A1` i `D2`: `font.bold`, `font.color.rgb`, `fill.fgColor.rgb`, `alignment.wrap_text`, `border.bottom.style`, `number_format`.

W pliku `output/09_cw1_wnioski.md` odpowiedz:

- **Ile obiektów `Font` utworzyłeś w kodzie?** Ile komórek je dostało? Ile wpisów spodziewasz się zobaczyć w `<fonts>` w `styles.xml`? **Sprawdź** to skryptem diagnostycznym z Przykładu 5 i porównaj z przewidywaniem.
- Dlaczego w punkcie 3 użyłeś **jednego** obiektu stylu, a nie tworzyłeś go czterokrotnie w pętli? Odpowiedz jednym zdaniem, ale takim, które uwzględnia **oba** koszty (Python i plik).
- Co by się stało, gdybyś opuścił argument `fill_type="solid"` w `PatternFill`? Sprawdź eksperymentalnie i opisz wynik.

**Zadanie 2 — dwie reprodukcje błędów i ich naprawa.**

Napisz `examples/09_cw2.py`, który zawiera **cztery** niezależne, ponumerowane sekcje:

1. **Sekcja 1 — cicha utrata danych.** Ustaw na komórce `A1` czcionkę o rozmiarze 18, kursywę i kolor `C00000`. Następnie — dokładnie jak w Eksperymencie 4 — przypisz `Font(bold=True)`. Wypisz, które cechy przetrwały, a które przepadły. **Napisz funkcję `dodaj_pogrubienie(komorka)`, która robi to poprawnie**, zachowując wszystkie pozostałe cechy.

2. **Sekcja 2 — kolory.** Sprawdź w `try/except` **sześć** wartości: `"FF0000"`, `"FFFF0000"`, `"#FF0000"`, `"F00"`, `"red"`, `"ff0000"`. Dla każdej wypisz, czy przeszła, a jeśli tak — jaki `rgb` ma wynikowy obiekt `Color`. **Bez zaglądania do sekcji 2.6.2** — najpierw zgadnij, potem sprawdź.

3. **Sekcja 3 — `fgColor` vs `bgColor`.** Utwórz trzy komórki:
   - `A1`: `PatternFill(fill_type="solid", fgColor="FFFF0000")`,
   - `A2`: `PatternFill(fill_type="solid", bgColor="FFFF0000")`,
   - `A3`: `PatternFill(fill_type="solid", fgColor="FFFF0000", bgColor="FF00FF00")`.
   Zapisz plik, otwórz go w Excelu (i jeśli masz — w LibreOffice) i opisz, co widzisz. Zapisz spostrzeżenia w pliku `output/09_cw2_kolory.md`.

4. **Sekcja 4 — styl kolumny.** Ustaw `ws.column_dimensions["B"].font = Font(bold=True, size=20)` i zapisz plik. **Następnie otwórz plik w Excelu i opisz, co widzisz w kolumnie B.** Potem napisz pętlę, która nadaje ten sam styl komórkom `B1:B10`, i porównaj efekty. Zapisz wniosek w `output/09_cw2_kolory.md`.

### 🟡 Warsztat

**Zadanie 3 — pełny rejestr stylów w praktyce.**

Rozbuduj rejestr z sekcji 2.12 do wersji produkcyjnej. Napisz `examples/09_cw3.py` z plikiem `styles.py` (albo z rejestrem w tym samym pliku — jak wolisz), który zawiera:

1. **Paletę** `PALETA` z co najmniej sześcioma kolorami (granat, granat ciemny, zieleń, czerwień, szary linia, szary tło).
2. **Rejestr** `STYLE` z co najmniej ośmioma wpisami: `naglowek`, `naglowek_zielony`, `tytul`, `waluta`, `procent`, `data`, `zebra`, `podsumowanie`, `uwaga` (czerwony pogrubiony), `pole_formularza`.
3. **Funkcję** `zastosuj_styl(komorka, nazwa)`, która:
   - rzuca `KeyError` z listą dostępnych stylów, gdy nazwa nie istnieje,
   - obsługuje wszystkie sześć elementów stylu,
   - **nie tworzy nowych obiektów** — tylko przypisuje te z rejestru.
4. **Funkcję** `zastosuj_styl_do_zakresu(ws, zakres, nazwa)`, która używa `range_boundaries` z `openpyxl.utils.cell` i stosuje styl do wszystkich komórek zakresu.
5. **Funkcję** `zastosuj_zebre(ws, zakres)`, która naprzemiennie nakłada `zebra` i czyści wypełnienie.
6. **Skrypt demonstracyjny**, który buduje raport:
   - tytuł w `A1:F1` (scalony, styl `tytul`),
   - nagłówki w wierszu 3 (`A3:F3`, styl `naglowek`),
   - 20 wierszy danych z zebrą,
   - kolumnę z kwotami (`waluta`), kolumnę z marżami (`procent`), kolumnę z datami (`data`),
   - dwa wiersze z `uwaga` w komórkach, w których marża < 0,2,
   - wiersz podsumowania ze stylem `podsumowanie`.

Musisz samodzielnie wybrać układ kolumn i typy danych — ćwiczenie polega także na zaprojektowaniu tabeli, nie tylko na pisaniu kodu.

W pliku `output/09_cw3_wnioski.md` odpowiedz:

1. **Policz, ile obiektów stylu powstaje w Twoim kodzie** przy starcie programu (nie w pętli). Ile ich użyłeś do wygenerowania 20 × 6 = 120 komórek danych? Uruchom diagnostykę z Przykładu 5 i porównaj `fonts`/`fills`/`borders` z liczbą komórek.
2. **Co musiałbyś zmienić, gdyby szef powiedział „od teraz granat ma być inny"?** Wypisz dokładnie, ile linii w ilu plikach musisz edytować. To jest miara jakości Twojego rejestru.
3. **Dlaczego `zastosuj_styl` przyjmuje nazwę jako napis, a nie `**kwargs`?** Jakie są zalety i wady każdego z tych projektów? Sformułuj odpowiedź w kontekście tego, że w module 23 ten wzorzec zamieni się w klasę.
4. **Co by się stało, gdybyś w rejestrze ustawił raz `Alignment(horizontal="center")`, a potem w innym wpisie ją zmodyfikował „w miejscu"?** Sprawdź eksperymentalnie, co się dzieje z komórkami, którym już przypisano ten styl — i wyciągnij wniosek o tym, dlaczego wszystkie obiekty w rejestrze muszą być niezmieniane po utworzeniu.

### 🔴 Wyzwanie

**Zadanie 4 — biblioteka stylów z `NamedStyle`, cache'em i diagnostyką.**

Napisz `examples/09_cw4.py` implementujący klasę `StyleLibrary`:

```python
class StyleLibrary:
    """Biblioteka stylow raportowych: rejestruje NamedStyle, stosuje, diagnozuje."""

    def __init__(self, wb, paleta: dict[str, str] | None = None):
        ...

    def zarejestruj_wszystkie(self) -> list[str]:
        """Rejestruje wszystkie style w skoroszycie. Zwraca liste nazw."""
        ...

    def zastosuj(self, komorka, nazwa: str) -> None:
        """Nakłada styl nazwany na komorke."""
        ...

    def zastosuj_do_zakresu(self, ws, zakres: str, nazwa: str) -> int:
        """Nakłada styl na zakres. Zwraca liczbe sformatowanych komorek."""
        ...

    def diagnoza(self) -> dict:
        """Zwraca raport: nazwy stylow, ich definicje, liczbe uzyc."""
        ...
```

**Wymagania funkcjonalne:**

- **Paleta jako parametr konstruktora.** Jeśli nie podano, użyj domyślnej. Klasa **nie może** mieć literałów kolorów w ciele metod — wszystkie kolory pochodzą z palety.
- **Style rejestrowane przez `NamedStyle`**, nie przez bezpośrednie przypisywanie obiektów. Nazwy muszą być **unikalne w skoroszycie** — jeśli styl o danej nazwie już istnieje (np. wczytany z szablonu), klasa musi to obsłużyć **bez wyjątku**: albo zwrócić istniejący, albo zarejestrować pod zmodyfikowaną nazwą i **zapisać to w raporcie**.
- **`zarejestruj_wszystkie` jest idempotentne** — wywołanie dwa razy nie tworzy duplikatów i nie rzuca błędu.
- **`zastosuj` używa `cell.style = nazwa`** (przypisanie po nazwie, nie przez obiekt). Jeśli nazwa nie istnieje w bibliotece, rzuca `KeyError` z listą dostępnych stylów.
- **`zastosuj_do_zakresu` używa `range_boundaries`** i zwraca liczbę komórek, którym nadano styl (musi być zgodna z liczbą komórek w zakresie).
- **`diagnoza`** zwraca słownik z polami: `style` (lista nazw), `definicje` (słownik nazwa → opis elementów), `liczba_komorek` (ile komórek w każdym arkuszu ma ustawiony `style`), `ostrzezenia` (lista napisów).

**Dodatkowo — funkcja pomiarowa:**

```python
def zmierz_strategie(liczba_komorek: int = 5000) -> dict:
    """Porownuje trzy strategie formatowania: NamedStyle per komorka,
    nowy obiekt per komorka, wspolny obiekt. Zwraca czasy i liczby stylow."""
```

Funkcja musi:
- dla każdej strategii zmierzyć czas `perf_counter`,
- zapisać każdy plik w `output/`,
- **policzyć faktyczną liczbę wpisów** w `<fonts>`, `<fills>`, `<cellXfs>` w zapisanym pliku (użyj `zipfile` i `ElementTree`, jak w Przykładzie 5),
- zwrócić słownik z porównaniem.

**Pięć testów:**

1. **Idempotencja:** `zarejestruj_wszystkie()` dwa razy → `wb.named_styles` ma tę samą zawartość, brak wyjątku.
2. **Unikalność:** ręcznie zarejestruj `NamedStyle(name="naglowek")` w skoroszycie, potem wywołaj `zarejestruj_wszystkie()` → brak wyjątku, `ostrzezenia` zawiera informację.
3. **Zakres:** `zastosuj_do_zakresu(ws, "A1:E10", "zebra")` zwraca `50`.
4. **`KeyError`:** `zastosuj(ws["A1"], "nie_ma_takiego")` → `KeyError` z listą dostępnych nazw.
5. **Diagnostyka:** `diagnoza()["liczba_komorek"]` zgadza się z ręcznie policzoną liczbą komórek ze stylem.

Na końcu wypisz tabelę wyników `test N: OK / BŁĄD` i podsumowanie `wynik: X/5`. Zapisz raport do `output/09_cw4_raport.txt`.

W pliku `output/09_cw4_wnioski.md` odpowiedz:

1. **Ile unikalnych stylów pojawiło się w pliku przy każdej strategii?** Podaj liczby z `zmierz_strategie` i wyjaśnij, dlaczego liczby dla strategii „nowy obiekt per komórka" i „wspólny obiekt" są inne, mimo że obie tworzą **jeden wpis** w `styles.xml`.
2. **Dlaczego `StyleLibrary` używa `NamedStyle`, a nie zwykłych obiektów w słowniku?** Wymień dwie zalety i jedną realną wadę. Odwołaj się do pułapki z 2.11.3.
3. **Co zrobi Twoja klasa, jeśli wczytasz szablon, w którym istnieje już styl o nazwie `naglowek`, ale wygląda zupełnie inaczej?** Opisz decyzję projektową i jej konsekwencje. Czy chcesz nadpisać szablon, czy zachować jego styl? **Uzasadnij** — i zwróć uwagę, że przy nadpisywaniu zmieniasz wygląd komórek, które już istnieją w pliku.
4. **Gdybyś miał dodać do biblioteki jedenasty styl o nazwie `podpis`, ile miejsc w kodzie musiałbyś zmienić?** Wypisz je. To jest miara otwartości Twojej klasy na rozbudowę.

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""StyleLibrary - biblioteka stylow raportowych na NamedStyle."""

from __future__ import annotations

import zipfile
from pathlib import Path
from time import perf_counter
from xml.etree import ElementTree

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils.cell import range_boundaries

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

PALETA_DOMYSLNA = {
    "granat": "2F5597",
    "granat_ciemny": "1F3864",
    "zielen": "548235",
    "czerwien": "C00000",
    "szary_linia": "BFBFBF",
    "szary_tlo": "F2F2F2",
    "biel": "FFFFFF",
}


class StyleLibrary:
    """Rejestruje i stosuje style nazwane w jednym skoroszycie."""

    def __init__(self, wb, paleta: dict[str, str] | None = None):
        self._wb = wb
        self._paleta = dict(paleta or PALETA_DOMYSLNA)
        # nazwa stylu -> faktyczna nazwa w skoroszycie (moze byc ze sufiksem)
        self._mapa: dict[str, str] = {}
        self._ostrzezenia: list[str] = []
        self._zarejestrowane = False

    # ------------------------------------------------------------------
    # BUDOWA DEFINICJI STYLOW - wszystko z palety, zero literalow w metodach
    # ------------------------------------------------------------------
    def _definicje(self) -> dict[str, dict]:
        p = self._paleta
        f = lambda k: f"FF{p[k]}"          # noqa: E731 - krotki helper lokalny

        linia_szara = Side(style="thin", color=f("szary_linia"))
        siatka = Border(left=linia_szara, right=linia_szara,
                        top=linia_szara, bottom=linia_szara)

        return {
            "naglowek": {
                "font": Font(name="Calibri", size=11, bold=True, color=f("biel")),
                "fill": PatternFill(fill_type="solid", fgColor=f("granat")),
                "alignment": Alignment(horizontal="center", vertical="center",
                                       wrap_text=True),
                "border": siatka,
            },
            "naglowek_zielony": {
                "font": Font(name="Calibri", size=11, bold=True, color=f("biel")),
                "fill": PatternFill(fill_type="solid", fgColor=f("zielen")),
                "alignment": Alignment(horizontal="center", vertical="center",
                                       wrap_text=True),
                "border": siatka,
            },
            "tytul": {
                "font": Font(name="Calibri", size=16, bold=True,
                             color=f("granat_ciemny")),
                "alignment": Alignment(horizontal="left", vertical="center"),
            },
            "waluta": {
                "number_format": '#,##0.00 "zł"',
                "alignment": Alignment(horizontal="right"),
            },
            "procent": {
                "number_format": "0.0%",
                "alignment": Alignment(horizontal="right"),
            },
            "data": {
                "number_format": "yyyy-mm-dd",
                "alignment": Alignment(horizontal="center"),
            },
            "zebra": {
                "fill": PatternFill(fill_type="solid", fgColor=f("szary_tlo")),
            },
            "podsumowanie": {
                "font": Font(bold=True),
                "border": Border(top=linia_szara,
                                 bottom=Side(style="medium", color=f("granat"))),
            },
            "uwaga": {
                "font": Font(bold=True, color=f("czerwien")),
            },
            "pole_formularza": {
                "border": Border(bottom=linia_szara),
                "alignment": Alignment(vertical="center"),
            },
        }

    # ------------------------------------------------------------------
    # REJESTRACJA - idempotentna i odporna na kolizje nazw
    # ------------------------------------------------------------------
    def zarejestruj_wszystkie(self) -> list[str]:
        if self._zarejestrowane:
            return list(self._mapa.values())

        istniejace = set(self._nazwy_w_skoroszycie())

        for nazwa, definicja in self._definicje().items():
            nazwa_faktyczna = nazwa

            if nazwa_faktyczna in istniejace:
                # Kolizja: nie nadpisujemy. Szukamy wolnej nazwy ze sufiksem.
                licznik = 2
                while f"{nazwa}_{licznik}" in istniejace:
                    licznik += 1
                nazwa_faktyczna = f"{nazwa}_{licznik}"
                self._ostrzezenia.append(
                    f"styl {nazwa!r} juz istnial w skoroszycie - "
                    f"zarejestrowano pod nazwa {nazwa_faktyczna!r}"
                )

            styl = NamedStyle(name=nazwa_faktyczna)
            for wlasciwosc, wartosc in definicja.items():
                setattr(styl, wlasciwosc, wartosc)

            self._wb.add_named_style(styl)
            istniejace.add(nazwa_faktyczna)
            self._mapa[nazwa] = nazwa_faktyczna

        self._zarejestrowane = True
        return list(self._mapa.values())

    def _nazwy_w_skoroszycie(self) -> list[str]:
        """Nazwy stylow w skoroszycie. Preferujemy API publiczne."""
        try:
            return list(self._wb.named_styles)       # 3.1+
        except AttributeError:
            return list(self._wb._named_styles)      # API prywatne - fallback

    # ------------------------------------------------------------------
    # STOSOWANIE
    # ------------------------------------------------------------------
    def _rozwiaz_nazwe(self, nazwa: str) -> str:
        if not self._zarejestrowane:
            self.zarejestruj_wszystkie()
        if nazwa not in self._mapa:
            dostepne = sorted(self._mapa)
            raise KeyError(f"nie ma stylu {nazwa!r}; dostepne: {dostepne}")
        return self._mapa[nazwa]

    def zastosuj(self, komorka, nazwa: str) -> None:
        komorka.style = self._rozwiaz_nazwe(nazwa)

    def zastosuj_do_zakresu(self, ws, zakres: str, nazwa: str) -> int:
        nazwa_faktyczna = self._rozwiaz_nazwe(nazwa)
        min_col, min_row, max_col, max_row = range_boundaries(zakres)

        licznik = 0
        for wiersz in range(min_row, max_row + 1):
            for kolumna in range(min_col, max_col + 1):
                ws.cell(row=wiersz, column=kolumna).style = nazwa_faktyczna
                licznik += 1
        return licznik

    # ------------------------------------------------------------------
    # DIAGNOSTYKA
    # ------------------------------------------------------------------
    def diagnoza(self) -> dict:
        if not self._zarejestrowane:
            self.zarejestruj_wszystkie()

        definicje = {}
        for oryginalna, faktyczna in self._mapa.items():
            czesci = []
            for wlasciwosc, wartosc in self._definicje()[oryginalna].items():
                czesci.append(f"{wlasciwosc}={type(wartosc).__name__}")
            definicje[faktyczna] = ", ".join(czesci) or "(pusty)"

        liczba_komorek: dict[str, int] = {}
        for ws in self._wb.worksheets:
            licznik = 0
            for wiersz in ws.iter_rows():
                for komorka in wiersz:
                    if komorka.style is not None and komorka.style != "Normal":
                        licznik += 1
            liczba_komorek[ws.title] = licznik

        return {
            "style": sorted(self._mapa.values()),
            "definicje": definicje,
            "liczba_komorek": liczba_komorek,
            "ostrzezenia": list(self._ostrzezenia),
        }
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **Prywatna mapa `self._mapa` zamiast nadpisywania nazw.** To najważniejszy wybór w całej klasie. Nazwy w `_definicje()` są **kluczami logicznymi** („nagłówek"), a `_mapa` tłumaczy je na nazwy faktyczne w skoroszycie („naglowek_2", jeśli „naglowek" już istniał). Dzięki temu kod **nigdy** nie polega na tym, że nazwa logiczna istnieje w skoroszycie — a jednocześnie nie nadpisuje stylów, które ktoś wcześniej stworzył. Ostrzeżenie ląduje w `_ostrzezenia`, więc użytkownik biblioteki **dowiaduje się** o kolizji.

- **Idempotencja przez flagę `_zarejestrowane`.** `zarejestruj_wszystkie()` można wywołać w dowolnym momencie — także wielokrotnie — i nie utworzy duplikatów. To wymóg, który w realnym kodzie ratuje sytuacje, gdy biblioteka jest używana przez wiele funkcji i nie wiadomo, która wywoła się pierwsza.

- **Wszystkie kolory pochodzą z palety, przez lokalny helper `f(k)`.** W ciele `_definicje()` nie ma ani jednego literału koloru. Zmiana palety to zmiana jednego słownika — i to jest cała wartość tego projektu. Zwróć uwagę, że helper `f` dodaje alfę `FF`, więc paleta zawiera czyste 6-znakowe kolory: **jedno miejsce, jedna konwencja**.

- **`zastosuj` używa `cell.style = nazwa` (po nazwie).** To konsekwencja wyboru `NamedStyle`: przypisanie po nazwie jest bardziej niezawodne niż przekazywanie obiektu, bo nie tworzy ryzyka, że ktoś w międzyczasie zmodyfikuje obiekt stylu i wpłynie na komórki przypisane później.

- **`zastosuj_do_zakresu` używa `range_boundaries` z `openpyxl.utils.cell`.** To ten sam narzędziowy wzorzec, który poznałeś w module 07: **parsowanie zakresu robi biblioteka, nie Twój `split(":")`**. Zwraca liczbę sformatowanych komórek, żeby można to było zweryfikować testem.

- **Diagnostyka iteruje po komórkach i sprawdza `komorka.style`.** To jest drogie (O(liczba komórek)) i dlatego **jest osobną metodą**, a nie czymś, co dzieje się przy każdym zapisie. Rozdzielenie „szybkiej ścieżki produkcyjnej" i „wolnej ścieżki diagnostycznej" to wzorzec, który w module 20 stanie się podstawą testów regresji.

- **`_nazwy_w_skoroszycie` z fallbackiem.** Preferuje publiczne `wb.named_styles`, ale w razie jego braku schodzi na `wb._named_styles`. To świadome, udokumentowane w kodzie użycie API prywatnego — dopuszczalne w **jednym miejscu**, z komentarzem, tak aby wymiana po zmianie API była jednym `if`-em. Zasada: prywatne API wolno użyć raz, w izolowanej funkcji, nigdy w logice rozproszonej po kodzie.

**Funkcja `zmierz_strategie` — szkic:**

```python
def _policz_sekcje(sciezka: Path) -> dict[str, int]:
    """Liczy wpisy w kluczowych sekcjach styles.xml."""
    with zipfile.ZipFile(sciezka) as archiwum:
        xml = archiwum.read("xl/styles.xml")
    korzen = ElementTree.fromstring(xml)
    wynik = {}
    for sekcja in ("fonts", "fills", "borders", "cellXfs"):
        element = korzen.find(f"m:{sekcja}", NS)
        wynik[sekcja] = len(list(element)) if element is not None else 0
    return wynik


def zmierz_strategie(liczba_komorek: int = 5000, output: Path | None = None) -> dict:
    """Porownuje trzy strategie formatowania."""
    output = output or Path("output")
    output.mkdir(parents=True, exist_ok=True)
    kolumn = 10
    rezultaty = {}

    # --- A: NamedStyle per komorka --------------------------------------
    wb = Workbook()
    ws = wb.active
    start = perf_counter()
    for indeks in range(liczba_komorek):
        nazwa = f"S_{indeks}"
        styl = NamedStyle(name=nazwa)
        styl.font = Font(size=20, bold=True, color="FFFFFFFF")
        styl.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
        wb.add_named_style(styl)
        ws.cell(row=indeks // kolumn + 1, column=indeks % kolumn + 1).style = nazwa
    czas_a = perf_counter() - start
    plik_a = output / "09_cw4_a.xlsx"
    wb.save(plik_a)
    wb.close()

    # --- B: nowy obiekt per komorka -------------------------------------
    wb = Workbook()
    ws = wb.active
    start = perf_counter()
    for indeks in range(liczba_komorek):
        c = ws.cell(row=indeks // kolumn + 1, column=indeks % kolumn + 1)
        c.font = Font(size=20, bold=True, color="FFFFFFFF")
        c.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
    czas_b = perf_counter() - start
    plik_b = output / "09_cw4_b.xlsx"
    wb.save(plik_b)
    wb.close()

    # --- C: jeden wspolny obiekt ----------------------------------------
    wb = Workbook()
    ws = wb.active
    font = Font(size=20, bold=True, color="FFFFFFFF")
    fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
    start = perf_counter()
    for indeks in range(liczba_komorek):
        c = ws.cell(row=indeks // kolumn + 1, column=indeks % kolumn + 1)
        c.font = font
        c.fill = fill
    czas_c = perf_counter() - start
    plik_c = output / "09_cw4_c.xlsx"
    wb.save(plik_c)
    wb.close()

    return {
        "liczba_komorek": liczba_komorek,
        "A_named_style_per_komorka": {"czas": czas_a, "styles": _policz_sekcje(plik_a)},
        "B_nowy_obiekt_per_komorka": {"czas": czas_b, "styles": _policz_sekcje(plik_b)},
        "C_wspolny_obiekt":          {"czas": czas_c, "styles": _policz_sekcje(plik_c)},
    }
```

**Czego szukać w wynikach:** strategia A będzie miała w `fill`/`fonts` **niewiele** wpisów (bo definicje są identyczne — deduplikacja działa!), ale w `cellStyles` (nazwy stylów) **tysiące**. Strategie B i C będą identyczne pod względem pliku (bo deduplikacja sprowadza je do jednego wpisu) — **a mimo to B będzie wolniejsza.** To jest namacalny dowód na to, że koszt złego podejścia do stylów ma **dwie** niezależne składowe: koszt Pythona (widoczny w czasie) i koszt pliku (widoczny w liczbie wpisów). Dobry kod minimalizuje oba.

**Co jeszcze warto dopisać w wersji produkcyjnej:** metodę `usun_nieuzywane_style()` (choć openpyxl i tak nie zapisze nieprzypisanych) oraz test porównujący `diagnoza()["liczba_komorek"]` z liczbą komórek ze stylem policzoną niezależnie — bo diagnostyka, która sama siebie weryfikuje, jest warta dwa razy więcej.

</details>

## 6. Typowe błędy i pułapki

**1. „`ws["A1"].font.bold = True` wywala `AttributeError`" (objaw) → `cell.font` zwraca `StyleProxy` — pełnomocnika tylko do odczytu — a nie obiekt `Font`. openpyxl blokuje zapis, żeby jedna zmiana nie wpłynęła na wszystkie komórki dzielące ten sam styl (przyczyna) → użyj jednej z trzech poprawnych form: `copy()` ze standardowej biblioteki (`nowy = copy(cell.font); nowy.bold = True; cell.font = nowy`), operatora `+` (`cell.font = cell.font + Font(bold=True)`) albo nowego pełnego obiektu. Zignoruj formę `cell.font.copy(bold=True)` — działa w 3.1.x, ale jest **przestarzała** i w przyszłych wersjach może zostać usunięta (naprawa).**

**2. „Ustawiłem rozmiar i kolor, potem dodałem pogrubienie — i rozmiar z kolorem przepadły. Żadnego błędu nie było" (objaw) → druga linia tworzy **zupełnie nowy** `Font`, w którym nieustawione cechy mają domyślne `None`. Przypisanie nowego obiektu **zastępuje** poprzedni, a nie uzupełnia go. To cicha utrata danych, bo openpyxl nie ma jak odróżnić „chcę wyczyścić rozmiar" od „zapomniałem o rozmiarze" (przyczyna) → **stała zasada: jeśli chcesz zmienić jedną cechę, zawsze zaczynaj od `copy()` istniejącego stylu.** Nowy obiekt od zera stosuj tylko wtedy, gdy faktycznie definiujesz cały wygląd od podstaw (np. w rejestrze stylów). W kodzie produkcyjnym warto rozważyć funkcję pomocniczą `dodaj_ceche(komorka, "font", bold=True)`, która robi `copy()` wewnątrz — wtedy pomyłka jest niemożliwa (naprawa).**

**3. „Kolor `#FF0000` wywala `ValueError: Colors must be aRGB hex values`, a przecież to poprawny czerwony" (objaw) → openpyxl waliduje kolor wzorcem `^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$` — dopuszcza **wyłącznie** 6 lub 8 znaków szesnastkowych. Krzyżyk, skrót 3-znakowy, nazwy słowne (`"red"`) i cyfry spoza zakresu szesnastkowego nie pasują (przyczyna) → usuwaj `#` przy kopiowaniu koloru z HTML/CSS i rozwijaj skróty: `#F00` → `"FF0000"`. **Kanał alfa nie jest wymagany** — 6 znaków jest w pełni poprawne i openpyxl dopisze `"00"` sam. Zapamiętaj: **to nie brak alfy powoduje błąd, a znaki spoza zakresu `[0-9A-Fa-f]` lub zła długość** (naprawa).**

**4. „Ustawiłem `PatternFill(bgColor="FFFF0000")` i tło jest białe" (objaw) → dla wypełnienia jednolitego (`fill_type="solid"`) Excel maluje całą powierzchnię **farbą**, czyli `fgColor`. `bgColor` to kolor papieru, który przy wypełnieniu jednolitym jest całkowicie zakryty. Niektóre inne programy czytające plik mogą interpretować `bgColor` inaczej (np. rysować czarne tło), więc plik może wyglądać różnie w Excelu i w LibreOffice (przyczyna) → **przy `solid` zawsze ustawiaj kolor w `fgColor`.** `bgColor` ma sens tylko przy wzorach (`fill_type="gray125"` itd.), gdzie widać oba kolory. Jeśli chcesz mieć pewność, że wynik wygląda identycznie wszędzie, ustawiaj `fill_type="solid"` i tylko `fgColor` (naprawa).**

**5. „Ustawiłem `ws.column_dimensions["B"].font = Font(bold=True)`, a kolumna B jest niepogrubiona" (objaw) → to **ograniczenie formatu pliku, nie openpyxl**. Dokumentacja mówi wprost: style zastosowane do kolumn i wierszy dotyczą **tylko komórek utworzonych w Excelu po zamknięciu pliku**. Istniejące komórki nie dostają stylu (przyczyna) → jeśli chcesz sformatować zawartość kolumny, **musisz przejść po komórkach**: `for row in range(1, ws.max_row + 1): ws.cell(row=row, column=2).font = FONT_POGRUBIONY`. Kod z `column_dimensions` nie podniesie błędu i coś nawet zapisze do pliku — dlatego ten błąd jest tak trudny do wykrycia (naprawa).**

**6. „Zmieniłem definicję `NamedStyle` i w pliku nic się nie zmieniło" (objaw) → dokumentacja ostrzega: *„raz przypisany styl nazwany do komórki nie podlega dalszym zmianom tego stylu"*. `cell.style = "Nazwa"` **kopiuje wartości** w momencie przypisania, a nie tworzy stałego powiązania (przyczyna) → **kolejność ma znaczenie: najpierw dopracuj definicję stylu, potem przypisuj go do komórek.** Jeśli komórki są już przypisane, musisz albo przypisać stil ponownie (`ws["A1"].style = "KPI"` — to zadziała, bo odczyta aktualną definicję), albo — jeśli komórek jest dużo — napisać pętlę, która je wszystkie zaktualizuje. Wzorcowo: rejestruj `NamedStyle` raz, na starcie programu, przed jakimkolwiek przypisaniem (naprawa).**

**7. „Zapisałem plik z 5000 komórmi i ma 40 MB" / „Excel otwiera raport pół minuty" (objaw) → w pętli tworzysz **nowy obiekt stylu dla każdej komórki**, a jeśli różnią się one choć minimalnie (np. `Font(size=11 + indeks*0.001)` albo `Color(tint=indeks/1000)`), to deduplikacja nie działa i `styles.xml` rośnie o tysiące wpisów. Liczba unikalnych stylów w pliku jest kosztem ponoszonym przez każdego odbiorcę raportu (przyczyna) → **twórz obiekty stylu raz, poza pętlą** i przypisuj ten sam egzemplarz do wszystkich komórek. Sprawdź wynik skryptem diagnostycznym z Przykładu 5: jeśli `<fonts>` ma tysiące wpisów przy dziesiątkach tysięcy komórek, masz do naprawy pętlę. Pamiętaj też, że problem ma **dwie** składowe: koszt Pythona (widoczny w czasie wykonania) i koszt pliku (widoczny w liczbie wpisów) — dobra wersja minimalizuje oba (naprawa).**

**8. „Ustawiłem `wrap_text=True` i `row_dimensions[1].height = 15`, a tekst jest ucięty" (objaw) → **openpyxl nigdy nie liczy wysokości wiersza.** Gdy wysokość nie jest ustawiona, Excel dopasowuje ją automatycznie do zawiniętego tekstu. Gdy jest ustawiona, Excel uznaje ją za decyzję użytkownika i **nie nadpisuje** — tekst wychodzi poza komórkę i zostaje przycięty (przyczyna) → **jeśli chcesz, żeby Excel sam dopasował wysokość do zawiniętego tekstu, nie ustawiaj wysokości w ogóle.** Jeśli wysokość jest konieczna (np. dla jednolitego układu wizualnego), oszacuj ją samodzielnie: około 15 punktów na linię tekstu, plus zapas. To jedno z tych miejsc, gdzie „nic nie robić" jest lepszą decyzją niż „zrobić coś" (naprawa).**

**9. „Otwarłem plik i Excel mówi, że wystąpił problem z zawartością, mimo że skrypt nie zgłosił błędu" (objaw) → najczęstsze przyczyny w kontekście tego modułu: (a) ustawienie obu `wrap_text=True` i `shrink_to_fit=True` na tej samej komórce (Excel traktuje te opcje jako wzajemnie wykluczające), (b) nietypowe wartości `text_rotation` poza zakresem obsługiwanym przez Twoją wersję openpyxl, (c) `Diagonal` w `Border` z nietypowymi kombinacjami (przyczyna) → wybieraj jedną z opcji `wrap_text`/`shrink_to_fit` (nigdy obu), ustawiaj `text_rotation` w bezpiecznym zakresie, a nietypowe opcje obramowania — pozostaw domyślnym. **Zawsze otwórz wygenerowany plik w Excelu przed wysłaniem go komukolwiek.** To reguła operacyjna, która wróci w module 18 i 20 (naprawa).**

**10. „W rejestrze stylów zmieniłem `Alignment` „w miejscu" i wszystkie komórki się rozjechały" (objaw) → obiekty stylu **nie są chronione przed mutacją, dopóki nie są przypisane do komórki**. `Alignment(horizontal="center")` to zwykły obiekt Pythona — możesz go zmienić: `a.horizontal = "right"`. A jeśli ten sam obiekt był już przypisany do 500 komórek, to przy zapisie **wszystkie 500 dostanie nową wartość** (przyczyna) → **traktuj każdy obiekt w rejestrze jako niezmienny po utworzeniu.** Jeśli chcesz wariant stylu, utwórz **nowy** obiekt. Zasada praktyczna: obiekty stylu definiuj na poziomie modułu, w literałach, i nigdy nie przypisuj do nich zmiennych, które potem modyfikujesz. Jeśli musisz coś policzyć (np. kolor zależny od progu), zrób to przed utworzeniem obiektu — i najlepiej zamknij wynik w słowniku o skończonej liczbie wariantów (naprawa).**

**11. „Wczytuję szablon i `wb.add_named_style(...)` wywala błąd" (objaw) → `NamedStyle` żyje w skoroszycie i nazwy muszą być unikalne. Jeśli szablon już ma styl o tej samej nazwie (np. `Normal`, `Title`, `Naglowek` z szablonu klienta), openpyxl zgłosi `ValueError` — dokładny typ i komunikat sprawdź w swojej wersji (przyczyna) → przed rejestracją sprawdź, jakie nazwy są zajęte (`wb.named_styles` w 3.1, `wb._named_styles` jako fallback), i zdecyduj **świadomie**: albo użyć istniejącego stylu, albo zarejestrować pod zmodyfikowaną nazwą (i ostrzec użytkownika), albo — jeśli naprawdę chcesz zmienić wygląd — zbudować raport od nowa. **Nadpisywanie stylu z szablonu jest niebezpieczne:** zmieniasz wygląd komórek, które już istnieją w pliku (naprawa).**

**12. „Podałem `ws["A1"].style = "Naglowek"`, ale dostałem błąd, że styl nie istnieje" (objaw) → przypisanie **po nazwie** wymaga, żeby styl był **już zarejestrowany** w skoroszycie. Świeży `Workbook()` **nie ma** Twoich stylów ani (w wielu przypadkach) wbudowanych stylów Excela — ma puste `_named_styles`. Dodatkowo wbudowane style Excela rozpoznawane są **tylko pod angielskimi nazwami** (dokumentacja: *„openpyxl rozpozna tylko nazwy angielskie i tylko dokładnie tak, jak tu zapisano"*) (przyczyna) → albo zarejestruj styl jawnie (`wb.add_named_style(styl)`), albo przypisz obiektem (`ws["A1"].style = styl` — openpyxl zarejestruje go automatycznie przy pierwszym przypisaniu). Dla wbudowanych stylów używaj nazw angielskich (`"Title"`, nie `"Tytuł"`). W praktyce: **miej jeden rejestr, który rejestruje wszystkie style na starcie** — to dokładnie `StyleLibrary` z zadania 4 (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Styl komórki to sześć współdzielonych „pieczątek", nie sześć pól w formularzu.** `Font`, `PatternFill`, `Border`, `Alignment` i `Protection` to obiekty tworzone raz i przybijane do wielu komórek; szóstym elementem jest `number_format`, czyli zwykły napis. Komórka w pliku nie ma stylu — ma **numer stylu**, który wskazuje wiersz w tablicy `cellXfs` w `xl/styles.xml`, a tam spięte są identyfikatory wszystkich sześciu elementów. Ta jedna obserwacja wyjaśnia zarówno deduplikację stylów, jak i całą pułapkę z niemutowalnością.

2. **`cell.font` zwraca pełnomocnika tylko do odczytu, a to jest zabezpieczenie, nie przeszkoda.** `StyleProxy` blokuje zapis komunikatem *„Style objects are immutable and cannot be changed. Reassign the style with a copy"*, bo gdyby zapis był dozwolony, zmiana jednego odcisku zmieniłaby wygląd wszystkich komórek dzielących ten sam egzemplarz stylu. **Twoja praca z formatowaniem sprowadza się do trzech ruchów:** (a) skopiuj styl (`copy()`), (b) zmień go, (c) przypisz z powrotem. Pamiętaj, że `a1.font == b2.font` porównuje wartości, ale `a1.font is b2.font` już nie — dlatego **w testach porównuj `==`, nigdy `is`**.

3. **Największe realne ryzyko to nie blokada, a cicha zamiana stylu na nowy, ubogi obiekt.** `ws["A1"].font = Font(bold=True)` **nie** dodaje pogrubienia do istniejącej czcionki — **zastępuje ją** czcionką, w której rozmiar, kolor i kursywa są domyślne (`None`). Żaden wyjątek tego nie zgłosi, a w Excelu zobaczysz po prostu inny, „zubożony" tekst. Dlatego reguła brzmi: **chcesz zmienić jedną cechę → zaczynaj od `copy()`. Chcesz zdefiniować cały wygląd → buduj nowy obiekt, ale wtedy ustaw wszystkie istotne cechy świadomie.**

4. **`NamedStyle` to szablon, ale nie łącze na żywo, i to jest źródło największego nieporozumienia wokół niego.** Styl nazwany **jest mutowalny** (w przeciwieństwie do stylu komórki), ale komórka, której go przypisano, **nie śledzi późniejszych zmian definicji** — dokumentacja ostrzega o tym wprost. Kolejność jest więc kluczowa: **najpierw dopracuj definicję stylu, potem przypisuj go do komórek.** Styl nazwany kupujesz nie za „globalną zmianę wyglądu", a za **powtarzalne przypisanie tego samego, sprawdzonego zestawu wartości** — plus nazwę, która komunikuje intencję (`"KPI"` mówi więcej niż `Font(size=20, bold=True, color="FFFFFFFF")`).

5. **Rejestr stylów to pierwszy wzorzec architektoniczny w tym kursie — i działa tylko wtedy, gdy obiekty w nim są niezmieniane po utworzeniu.** Jedno miejsce z literałami kolorów (`PALETA`), jedno miejsce z definicjami stylów (`STYLE`), jedno miejsce, które wie, jak styl nałożyć (`zastosuj_styl`). Dzięki temu zmiana firmowego koloru to jedna linia, a nie czterdzieści wystąpień na ośmiu plikach. **Ale uwaga:** dopóki obiekt nie jest przypisany do komórki, jest zwykłym obiektem Pythona, który można zmienić „w miejscu" — i wtedy przy zapisie zmienią się wszystkie komórki, które go dostały. Twórz obiekty stylu raz, **poza pętlą**, i **następnie ich nie dotykaj**. Sprawdzalna miara jakości: liczba wpisów w `<fonts>` i `<fills>` w `xl/styles.xml` powinna być mała i **niezależna od liczby komórek**.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY - najczestszy zestaw
# ==================================================================
from copy import copy                                        # kopia stylu
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Alignment, Border, Color, Font, NamedStyle, PatternFill, Protection, Side,
)
from openpyxl.utils.cell import range_boundaries             # parsowanie zakresow

# ==================================================================
# 2. NIEMUTOWALNOSC - trzy poprawne sposoby zmiany jednej cechy
# ==================================================================
# (A) copy() - METODA KANONICZNA, zalecana przez dokumentacje
nowy = copy(ws["A1"].font)
nowy.bold = True
ws["A1"].font = nowy

# (B) operator + (galaz 3.1) - sprawdz, czy dostepny w Twojej wersji
ws["A1"].font = ws["A1"].font + Font(bold=True)

# (C) pelny nowy obiekt - tylko gdy definiujesz CALY wyglad
ws["A1"].font = Font(name="Calibri", size=12, bold=True, color="FF2F5597")

# ❌ NIE DZIALA:        ws["A1"].font.bold = True
#    -> AttributeError: "Style objects are immutable and cannot be changed."
# ❌ CICHA UTRATA:      ws["A1"].font = Font(bold=True)
#    -> gubi size, color, italic, underline... bez zadnego bledu
# ⚠️ PRZESTARZALE:      ws["A1"].font.copy(bold=True)
#    -> dziala w 3.1.x, ale emituje DeprecationWarning

# ==================================================================
# 3. SZESC ELEMENTOW STYLU
# ==================================================================
ws["A1"].font          = Font(name="Calibri", size=11, bold=True,
                              italic=False, underline="none", strike=False,
                              color="FFFFFFFF", vertAlign=None, scheme=None)
ws["A1"].fill          = PatternFill(fill_type="solid", fgColor="FF2F5597")
ws["A1"].border        = Border(left=Side(style="thin", color="FFBFBFBF"),
                                right=Side(style="thin", color="FFBFBFBF"),
                                top=Side(style="thin", color="FFBFBFBF"),
                                bottom=Side(style="thin", color="FFBFBFBF"))
ws["A1"].alignment     = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True, text_rotation=0, indent=0)
ws["A1"].protection    = Protection(locked=False, hidden=False)
ws["A1"].number_format = '#,##0.00 "zł"'          # to NIE jest obiektem stylu

# ==================================================================
# 4. KOLORY - trzy fakty
# ==================================================================
from openpyxl.styles import Font, Color

Font(color="FF0000")          # ✅ OK -> openpyxl dopisze "00" -> "00FF0000"
Font(color="FFFF0000")        # ✅ OK -> jawna alfa
Font(color="#FF0000")         # ❌ ValueError: Colors must be aRGB hex values
Font(color="F00")             # ❌ ValueError - skrot 3-znakowy nie jest wspierany
Font(color="red")             # ❌ ValueError - nazwy slowne to inny mechanizm

Color(rgb="2F5597")           # ✅ 6 znakow (aRGB_REGEX dopuszcza 6 lub 8)
Color(theme=6, tint=0.5)      # zalezy od motywu - NIE zalecane
Color(indexed=32)             # legacy - NIE zalecane
# ZASADA: uzywaj aRGB ("RRGGBB" albo "AARRGGBB") - jedyny jednoznaczny zapis

# ==================================================================
# 5. WYPELNIENIE - pulapka fgColor / bgColor
# ==================================================================
# Dla "solid" kolor idzie do fgColor (farby), bgColor jest zakryty.
PatternFill(fill_type="solid", fgColor="FFFF0000")                  # ✅ widoczne
PatternFill(fill_type="solid", bgColor="FFFF0000")                  # ❌ biale tlo
PatternFill(fgColor="FFFF0000")                                     # ❌ brak fill_type!
PatternFill(fill_type="gray125", fgColor="FFFF0000",                # ✅ wzor: farba + papier
            bgColor="FFFFFF00")

# ==================================================================
# 6. NAGLOWEK TABELI - gotowy zestaw
# ==================================================================
LINIA = Side(style="thin", color="FFBFBFBF")
FONT_NAGLOWKA = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")   # biala
FILL_NAGLOWKA = PatternFill(fill_type="solid", fgColor="FF2F5597")           # granat
ALIGN_NAGLOWKA = Alignment(horizontal="center", vertical="center", wrap_text=True)
BORDER_NAGLOWKA = Border(left=LINIA, right=LINIA, top=LINIA, bottom=LINIA)

for komorka in ws["A1:F1"][0]:
    komorka.font = FONT_NAGLOWKA
    komorka.fill = FILL_NAGLOWKA
    komorka.alignment = ALIGN_NAGLOWKA
    komorka.border = BORDER_NAGLOWKA

# ==================================================================
# 7. OBWODKA ZAKRESU (gruba ramka + cienka siatka wewnatrz)
# ==================================================================
from openpyxl.styles import Border, Side
from openpyxl.utils.cell import range_boundaries

CIENKA = Side(style="thin", color="FF808080")
GRUBA = Side(style="medium", color="FF2F5597")

_cache: dict = {}
def _border(l, p, g, d) -> Border:
    klucz = (l, p, g, d)
    if klucz not in _cache:
        _cache[klucz] = Border(left=l, right=p, top=g, bottom=d)
    return _cache[klucz]

def obramuj(ws, zakres: str) -> None:
    min_c, min_r, max_c, max_r = range_boundaries(zakres)
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            ws.cell(row=r, column=c).border = _border(
                GRUBA if c == min_c else CIENKA,
                GRUBA if c == max_c else CIENKA,
                GRUBA if r == min_r else CIENKA,
                GRUBA if r == max_r else CIENKA,
            )

# ==================================================================
# 8. ZEBRA - dwa wypelnienia na caly arkusz
# ==================================================================
FILL_PARZYSTE = PatternFill(fill_type="solid", fgColor="FFF2F2F2")
FILL_NIEPARZYSTE = PatternFill(fill_type=None)

for i, wiersz in enumerate(ws.iter_rows(min_row=2, max_col=6)):
    fill = FILL_PARZYSTE if i % 2 == 0 else FILL_NIEPARZYSTE
    for komorka in wiersz:
        komorka.fill = fill

# ==================================================================
# 9. NAMEDSTYLE - pelny cykl
# ==================================================================
kpi = NamedStyle(name="KPI")
kpi.font = Font(size=20, bold=True, color="FFFFFFFF")
kpi.fill = PatternFill(fill_type="solid", fgColor="FF2F5597")
kpi.alignment = Alignment(horizontal="center", vertical="center")
kpi.number_format = "#,##0"

wb.add_named_style(kpi)        # rejestracja (jawne)
ws["A1"].style = "KPI"         # przypisanie PO NAZWIE
# (przypisanie obiektem zarejestruje styl automatycznie:)
ws["B1"].style = kpi

# ⚠️ PULAPKA: komorka nie sledzi POZNIEJSZYCH zmian definicji stylu!
kpi.font = Font(size=30)       # A1 i B1 tego NIE zobacza

# Kolejnosc: NAJPIERW dopracuj styl, POTEM przypisuj.
# Sprawdzenie zajetych nazw (API publiczne w 3.1+):
dostepne = set(wb.named_styles)          # fallback: set(wb._named_styles)

# Wbudowane style - tylko angielskie nazwy!
ws["A1"].style = "Good"        # ✅
# ws["A1"].style = "Dobry"     # ❌ polska nazwa nie jest rozpoznawana

# ==================================================================
# 10. WZORZEC: REJESTR STYLOW (pomost do modulu 23 i 26)
# ==================================================================
PALETA = {"granat": "2F5597", "biel": "FFFFFF", "szary": "F2F2F2"}

LINIA = Side(style="thin", color="FFBFBFBF")
STYLE = {
    "naglowek": {
        "font": Font(bold=True, color=f"FF{PALETA['biel']}"),
        "fill": PatternFill(fill_type="solid", fgColor=f"FF{PALETA['granat']}"),
        "alignment": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "border": Border(left=LINIA, right=LINIA, top=LINIA, bottom=LINIA),
    },
    "waluta": {
        "number_format": '#,##0.00 "zł"',
        "alignment": Alignment(horizontal="right"),
    },
    "zebra": {
        "fill": PatternFill(fill_type="solid", fgColor=f"FF{PALETA['szary']}"),
    },
}

def zastosuj_styl(komorka, nazwa: str) -> None:
    if nazwa not in STYLE:
        raise KeyError(f"nie ma stylu {nazwa!r}; dostepne: {sorted(STYLE)}")
    for wlasciwosc, wartosc in STYLE[nazwa].items():
        setattr(komorka, wlasciwosc, wartosc)

# ==================================================================
# 11. WYDAJNOSC - reguly
# ==================================================================
# ✅ tworz obiekty stylu RAZ, poza petla
FONT = Font(bold=True)
for komorka in ws["A1:A1000"][0]:
    komorka.font = FONT

# ❌ nie tworz obiektu per komorka
# for komorka in ...: komorka.font = Font(bold=True)

# ✅ style niezmienne - nie modyfikuj ich po utworzeniu (mutacja "w miejscu"
#    zmieni WSZYSTKIE komorki, ktore je dostaly - na etapie zapisu)

# ==================================================================
# 12. CZEGO NIE DA SIE ZROBIC (ograniczenia)
# ==================================================================
# ❌ Formatowanie istniejacych komorek przez kolumne/wiersz:
#      ws.column_dimensions["A"].font = ...   <- nie zadziala na istniejace!
#    Ograniczenie FORMATU PLIKU. Trzeba petla po komorkach.
#
# ❌ Wysokosc wiersza liczona automatycznie przez openpyxl:
#      openpyxl NIE liczy wysokosci. Nie ustawiaj height -> Excel dopasuje.
#      Ustaw height -> tekst zawijany moze zostac przyciety.
#
# ❌ wrap_text RAZEM z shrink_to_fit: Excel traktuje je jako wzajemnie
#    wykluczajace. Wybierz jedno.
#
# ❌ Zmiana jednej cechy bez znajomosci pozostalych:
#      cell.font = Font(bold=True)  <- gubi wszystko inne, CICHO.

# ==================================================================
# 13. DIAGNOSTYKA - ile stylow naprawde jest w pliku
# ==================================================================
import zipfile
from xml.etree import ElementTree

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def policz_styles(sciezka) -> dict[str, int]:
    with zipfile.ZipFile(sciezka) as z:
        korzen = ElementTree.fromstring(z.read("xl/styles.xml"))
    return {
        sekcja: len(list(korzen.find(f"m:{sekcja}", NS)))
        for sekcja in ("numFmts", "fonts", "fills", "borders", "cellXfs")
        if korzen.find(f"m:{sekcja}", NS) is not None
    }
# Przykladowy wynik: {'fonts': 3, 'fills': 4, 'borders': 2, 'cellXfs': 5}
# Pierwsze DWA wpisy w <fills> sa zarezerwowane (none i gray125).
# numFmtId >= 164 => wlasny format liczby (modul 10).
```

## 9. Co dalej

Masz za sobą moduł, który w tym kursie zbiera najwięcej żniwa. Zrozumienie modelu stylu to granica, po której formatowanie przestaje być zgadywaniem („chyba trzeba zrobić `Font` i przypisać") i staje się inżynierią („tworzę obiekt raz, przypisuję świadomie, wiem, co pójdzie do `<fonts>`"). Trzy rzeczy, które zabierasz ze sobą dalej:

- **`copy()` jako odruch.** Od teraz, gdy zobaczysz w kodzie `cell.font = Font(...)`, pierwszą rzeczą, którą sprawdzisz, będzie: *„czy to na pewno definicja całego stylu od zera, czy powinienem był skopiować?"*. To pytanie uratuje Ci dziesiątki godzin debugowania „czemu ten raport wygląda inaczej".
- **Rejestr jako domyślny sposób pracy.** W modułach 13–16 (formatowanie warunkowe, wykresy) i 22–27 (wzorce, architektura) rejestr będzie wracał raz po raz — raz jako `STYLE`, raz jako `PALETA`, raz jako konfiguracja. **To nie jest ozdoba, to jest warunek spójności raportu.**
- **Rozróżnienie „potrafi / potrafi częściowo / gubi".** W tym module zobaczyłeś je trzy razy: formatowanie kolumn (gubi — nie działa na istniejące komórki), scalone komórki (częściowo — zależy od czytnika) i `NamedStyle` (potrafi, ale nie tak, jak podpowiada intuicja). Ta sama trójka wróci w każdym kolejnym module.

**W module 10** zajmiemy się **szóstym elementem stylu** — czyli `number_format`. To pozornie najprostsza rzecz w całym module 09 („zwykły napis"), a w praktyce najbardziej zdradliwa. Zobaczysz:

- dlaczego `0.35` z formatem `"0.0%"` to `35,0%`, ale `35` z tym samym formatem to `3500,0%` — i jak ta pomyłka wygląda w raporcie finansowym,
- jak zbudowany jest kod formatu (cztery sekcje, znaki `#`, `0`, `_`, `*`, `@`, przecinek jako separator tysięcy **i** dzielnik przez tysiąc jednocześnie),
- dlaczego polska nazwa miesiąca w formacie `[$-415]dd mmmm yyyy` zależy od **ustawień systemu odbiorcy**, a nie od Twojego kodu,
- jak zapisać czas trwania przekraczający 24 godziny (`[h]:mm:ss`) — bo to najczęściej pytanie w raportach o nadgodziny,
- dlaczego `Decimal` i pieniądze to temat na osobny akapit (i jak to obejść),
- oraz skąd w tym wszystkim biorą się identyfikatory `numFmtId >= 164`, które już widziałeś w diagnostyce `styles.xml`.

**Zanim przejdziesz dalej, wykonaj dwa przygotowania:**

1. **Uruchom `examples/09_diagnostyka_styles.py`** na katalogu `output/` — po tym module powinno tam być kilka plików z Przykładów 1–4. Zapisz sobie liczby z `<fonts>`, `<fills>` i `<cellXfs>`. W module 10 zobaczysz, że **`<cellXfs>` rośnie, gdy dodajesz tylko formaty liczb** — bo format liczby jest częścią zestawu stylu. To połączy Ci dwa moduły w jeden obraz.

2. **Zadaj sobie jedno pytanie o własny raport:** *„ile mam unikalnych formatów liczb?"* Nie odpowiadaj zgadywaniem — otwórz plik jako ZIP, znajdź `xl/styles.xml`, policz wpisy w `<numFmts>` i zapisz liczbę. Jeśli wpisów jest dziesięć i wszystkie są potrzebne — świetnie. Jeśli pięć i trzy z nich to wariacje tego samego („zł" raz na początku, raz na końcu, raz z kropką, raz z przecinkiem), to właśnie odkryłeś temat modułu 10: **format liczby to nie kosmetyka, to kontrakt między raportem a jego odbiorcą.**