Zanim napiszę moduł, zweryfikuję kilka szczegółów API, które w materiałach o druku i układzie arkusza są najczęściej podawane błędnie: `column_dimensions.group()`, `page_breaks` vs `row_breaks`/`col_breaks`, oraz dokładne działanie `freeze_panes`.
Ważne: `page_breaks` prawdopodobnie nie istnieje (są `row_breaks`/`col_breaks`), a `freeze_panes` ma nieoczywistą semantykę. Sprawdzę jeszcze nagłówki/stopki i `pageSetUpPr`.
Potwierdzone kluczowe fakty. Sprawdzę jeszcze dokładną semantykę `freeze_panes` (to najczęściej mylone miejsce) i strukturę `print_options`/`page_margins`.
Mam potwierdzenie, że `ws.page_breaks` w briefie to pomyłka — realne API to `ws.row_breaks` i `ws.col_breaks` z klasami `Break`/`RowBreak`/`ColBreak`. Sprawdzę jeszcze `zoomedScale`/`zoomScale` i `fitToHeight` w 3.1.x.
Sprawdzę jeszcze, czy `ws.print_area` przyjmuje listę i jak działa `sheet_properties.pageSetUpPr`.
Ostatnia weryfikacja: dokładna semantyka `freeze_panes` (reguła „zamraża wszystko powyżej i na lewo") oraz domyślne wartości `fitToWidth`/`fitToHeight`.

Potwierdzone. Teraz zweryfikuję drukowanie: czy `ws.page_breaks` istnieje, jak działają `print_title_rows`, `print_options` i nagłówki/stopki.

Ostatnia weryfikacja — najważniejsza dla pułapki z briefu: czy `ws.sheet_properties.pageSetUpPr.fitToPage` działa od razu, czy wymaga wcześniejszej inicjalizacji obiektu `PageSetupProperties`.

# Moduł 11 — Wygląd arkusza: od danych do dokumentu gotowego do druku

> **Część:** II — Formatowanie · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–10

## 0. W tym module nauczysz się

- **Zrozumiesz różnicę między arkuszem jako magazynem danych a arkuszem jako dokumentem.** Do tej pory traktowaliśmy arkusz jak tabelę: komórka, wartość, format komórki. Teraz spojrzysz na niego jak na **kartkę papieru**, która ma marginesy, kolumny o konkretnej szerokości, nagłówek powtarzany na każdej stronie i obszar przeznaczony do druku. To dwa różne spojrzenia na ten sam plik i openpyxl obsługuje oba.
- **Opanujesz wymiary arkusza:** szerokości kolumn, wysokości wierszy, grupowanie kolumn w „szuflady" (`outline`) i ukrywanie tego, czego odbiorca nie powinien widzieć (kolumny pomocnicze, wiersze robocze). Dowiesz się, dlaczego szerokość kolumny podaje się w **znakach**, a nie w pikselach, i dlaczego to nie jest niedbalstwo twórców Excela.
- **Zrozumiesz, czym są linie siatki i dlaczego nie są obramowaniem.** Zobaczysz, dlaczego tabela raportowa powinna mieć **własne** obramowanie i co się dzieje, gdy polegasz na szarej siatce — zwłaszcza przy druku.
- **Poznasz scalanie komórek wraz z jego cichą pułapką:** scalanie **usuwa wartości** z komórek innych niż lewa-górna, a próba wpisania czegokolwiek do scalonej komórki kończy się `AttributeError`. To jeden z najczęstszych sposobów na bezpowrotną utratę danych w raporcie.
- **Zrozumiesz dokładną zasadę `freeze_panes`** — jedną regułę, która tłumaczy wszystkie cztery warianty zamrażania, łącznie z tym, dlaczego `"A1"` **nie zamraża niczego** (i jest sposobem na usunięcie zamrożenia).
- **Odróżnisz autofiltr od tabeli** (moduł 12) i będziesz wiedział, kiedy użyć którego. Nauczysz się dodawać filtry i sortowanie programowo.
- **Nauczysz się przygotować arkusz do druku i eksportu do PDF** — orientacja, rozmiar papieru, marginesy, obszar wydruku, **powtarzanie nagłówka na każdej stronie** (`print_title_rows`), dopasowanie do szerokości strony (`fitToWidth` + `fitToHeight` + `fitToPage`) oraz nagłówki i stopki z numeracją stron.
- **Zobaczysz, jak działa podział strony** i poznasz prawdziwe API (`ws.row_breaks` / `ws.col_breaks` z obiektami `Break`) — bo `ws.page_breaks`, które pojawia się w wielu poradnikach, **nie istnieje**.
- **Zbudujesz gotowy „layout raportu"** — kompletny zestaw ustawień, który możesz skopiować do każdego przyszłego projektu.

## 1. Intuicja i analogia

### 1.1. Analogia główna: dane to treść, arkusz to skład tekstu

Wyobraź sobie redakcję gazety. Pracują tam **dwie różne osoby** i nigdy nie powinny się mieszać:

**Reporter** zbiera fakty. Pisze: „W marcu sprzedaż wyniosła 143 700 zł, wzrost o 12%". Reporter **nie interesuje się** tym, jaką czcionką to zostanie wydrukowane, czy tekst zajmie jedną kolumnę, czy dwie, ani czy nagłówek znajdzie się na górze strony 3. Reporter dostarcza **treść**.

**Typograf** bierze tekst i układa go na stronie. Ustala szerokość łamów, wielkość nagłówków, marginesy, to, co się powtarza na każdej stronie (np. numer gazety i data). Typograf **nie zmienia ani jednego faktu** — tylko decyduje, jak fakty zostaną pokazane.

W module 10 (i w całym module 09) zajmowaliśmy się **reporterem dla jednej komórki** — jak wygląda jedna liczba. Teraz wchodzimy w rolę typografa i układamy **cały arkusz**.

I tu jest kluczowe zdanie, które ma dokładnie tę samą strukturę co „wartość to nie reprezentacja" z modułu 10:

> **Dane w arkuszu i jego wygląd to dwa niezależne światy. Możesz zmienić cały układ — szerokości, scalenia, druk — i nie dotknąć ani jednej wartości. I odwrotnie: możesz zmienić wszystkie dane, a układ zostanie ten sam.**

To ma praktyczną konsekwencję, którą docenisz przy pierwszym większym projekcie: **układ można w całości opisać w jednym miejscu, w jednej funkcji `zastosuj_layout(ws)`, i wywołać ją na każdym arkuszu każdego raportu.** W module 26 zrobimy z tego pełnoprawną warstwę prezentacji.

### 1.2. Analogia: linie siatki to rusztowanie, nie ściana

Otwórz nowy arkusz Excela. Widzisz szarą siatkę. Ta siatka **wygląda jak tabela**, ale nią nie jest.

Pomyśl o budowie domu. Na placu budowy stoi **rusztowanie** — siatka rur i pomostów, która pozwala pracować na wysokości. Nikt jednak nie myli rusztowania ze **ścianą**. Rusztowanie jest tymczasowe, nie jest częścią budynku i zniknie, gdy budynek będzie gotowy. Ściana to element konstrukcji — z cegły, otynkowany, widoczny w projekcie.

**Linie siatki Excela to rusztowanie.** Pomagają nawigować w trakcie pracy, ale:

- nie są częścią Twojego dokumentu,
- domyślnie **nie drukują się** (i bardzo dobrze),
- nie mają nic wspólnego z obramowaniem komórki.

**Obramowanie komórki (`Border` z modułu 09) to ściana.** To element, który świadomie zaprojektowałeś, który jest częścią arkusza i który przejdzie na wydruk.

Konsekwencja, którą zapamiętaj na zawsze:

> **Raport, który „wygląda jak tabela" tylko dzięki liniom siatki, nie ma żadnych ramek.** Wyłącz siatkę (`ws.sheet_view.showGridLines = False`), wydrukuj go albo wyślij komuś, kto otworzy go na telefonie — i tabela się rozsypie.

W praktyce stosuje się to tak: **wyłącz siatkę i nadaj własne obramowanie**. Dwie linie kodu, a dokument zmienia się z „arkusza z liczbami" w „raport".

### 1.3. Analogia: scalanie to sklejanie kartek

To najważniejsza analogia ostrzegawcza tego modułu.

Masz przed sobą **cztery kartki papieru**, na każdej coś napisane:

```
[A1: Raport]  [B1: kwartalny]  [C1: 2026]  [D1: Q1]
```

Chcesz mieć **jeden długi nagłówek** rozciągnięty nad całą tabelą, więc bierzesz kartki i **sklejasz je w jedną taśmą**. Powstaje jedna szeroka kartka.

I teraz kluczowe pytanie: **co się stało z treścią z kartek B1, C1 i D1?**

Odpowiedź: **została zakryta.** Klej i taśma zaszły na tekst. Widzisz tylko to, co było na pierwszej kartce — „Raport" — bo tylko ona jest widoczna jako lewy-górny fragment sklejonej kartki. Resztę napisów trzeba było **najpierw przepisać na kartkę A1**, zanim cokolwiek skleiłeś.

Dokładnie tak działa `merge_cells`:

- **Lewa-górna komórka zachowuje wartość.**
- **Wszystkie pozostałe komórki w scalonym zakresie tracą wartość** — openpyxl zamienia je na obiekty `MergedCell`, które są **tylko do odczytu** i mają `value = None`.
- **Próba wpisania wartości do scalonej komórki niebędącej lewą-górną kończy się `AttributeError`.**

I jeszcze jedna konsekwencja, o której się zapomina: **scalenie jest nie do odwrócenia w sensie danych**. Możesz wywołać `unmerge_cells`, ale odzyskasz tylko podział komórek — **wartości już nie wrócą**, bo openpyxl ich nie zapamiętał. Wyczyścił je bezpowrotnie.

**Zasada praktyczna:** scalaj **na końcu**, po wpisaniu wszystkich danych. Albo jeszcze lepiej: scalaj tylko to, co naprawdę musi być scalone (nagłówek nad tabelą, tytuł raportu) i nigdy nie scalaj komórek, w których są dane.

### 1.4. Analogia: `freeze_panes` to przyklejony nagłówek na szybie

Wyobraź sobie, że patrzysz na długą tabelę przez **okno**. Tabela jest długa i szeroka, więc musisz „przewijać" widok — przesuwać go jak mapę pod szybą.

Teraz przyklejasz do szyby **kawałek folii z nagłówkami**: nazwy kolumn u góry i etykiety wierszy z lewej. Gdy przesuwasz mapę, przyklejony fragment **zostaje na miejscu**, a reszta się przewija.

Ale jest jeden szczegół, który sprawia, że `freeze_panes` myli wszystkich. **Nie mówisz Excelowi „zamroź wiersz 1". Mówisz mu „tu zaczyna się obszar, który się przewija".** Wszystko **powyżej i na lewo** od tego punktu zostaje zamrożone. To jedna reguła i ona tłumaczy całą tabelę:

| Ustawienie | Co zamraża | Efekt |
|---|---|---|
| `ws.freeze_panes = "A2"` | wiersz 1 | Nagłówek kolumn zostaje, resztę przewijasz w dół |
| `ws.freeze_panes = "B1"` | kolumna A | Etykiety wierszy zostają, resztę przewijasz w prawo |
| `ws.freeze_panes = "B2"` | wiersz 1 **i** kolumna A | Oba naraz — klasyczny układ raportu |
| `ws.freeze_panes = "C3"` | wiersze 1–2 **i** kolumny A–B | Dwa wiersze nagłówka i dwie kolumny etykiet |
| `ws.freeze_panes = "A1"` | **nic** | To sposób na **usunięcie** zamrożenia |
| `ws.freeze_panes = None` | **nic** | Też usuwa zamrożenie |

**Reguła kciuka:** cyfra w adresie jest o **jeden większa** niż numer ostatniego zamrożonego wiersza, a litera o **jeden dalej** niż ostatnia zamrożona kolumna.

I uwaga na pułapkę, która spotyka 90% początkujących: **`"A1"` nie zamraża wiersza 1, tylko nie zamraża niczego.** Jeśli ustawiłeś `"A1"` i dziwisz się, że nagłówek ucieka przy przewijaniu — to dlatego.

### 1.5. Analogia: `fitToPage`, `fitToWidth`, `fitToHeight` to trzy przełączniki jednego urządzenia

Znasz to z kserokopiarki. Masz tam pokrętło skali — możesz ustawić 80%, 120%, cokolwiek. I masz osobny przycisk: **„zmniejsz, żeby zmieścić na jednej stronie"**.

Możesz przekręcić pokrętło skali na 80% — i nic się nie stanie, jeśli nie wybierzesz trybu „zmniejsz, żeby zmieścić". Pokrętło działa **tylko w tym trybie**.

Tak samo jest w Excelu:

- `ws.page_setup.fitToWidth = 1` — to jest **ustawienie**: „jedna strona szerokości".
- `ws.page_setup.fitToHeight = 0` — to jest **ustawienie**: „tyle stron wysokości, ile wyjdzie" (zero znaczy „bez limitu", nie „zero stron"!).
- `ws.sheet_properties.pageSetUpPr.fitToPage = True` — **i to jest włączenie trybu.**

**Bez trzeciej linii pierwsze dwie są ignorowane.** Excel zapisze wartości do pliku, ale przy druku zachowa się tak, jakby ich nie było, i tabela rozjedzie się na cztery strony w poziomie. To najczęstsza „niewytłumaczalna" pułapka w eksporcie do PDF — i dlatego poświęcam jej w tym module osobny blok.

### 1.6. Analogia: nagłówki i stopki to osobny język, nie tekst

W module 10 mówiliśmy, że kod formatu `#,##0.00 "zł"` to nie zwykły tekst, tylko przepis. W nagłówkach i stopkach druku jest dokładnie tak samo.

Gdy wpiszesz do stopki `Strona &P z &N`, Excel **nie wyświetli** tego napisu dosłownie. Zobaczy, że `&P` to **polecenie** („wstaw numer strony"), a `&N` też polecenie („wstaw liczbę stron"). Wynik na wydruku będzie brzmiał `Strona 3 z 12`.

**Analogia:** to jak **pola w formularzu** albo **kody pocztowe w maszynie do adresowania**. Napis „&P" nie znaczy „ampersand i litera P" — znaczy „wstaw tutaj numer strony". Ten sam mechanizm, co `&[Page]` w nowszych wersjach Excela.

I tu jest pułapka, która kosztuje ludzi godziny: **jeśli naprawdę chcesz ampersand w stopce** (np. „Sprzedaż & Marketing"), musisz napisać **`&&`**. Samotny `&` zostanie potraktowany jako początek polecenia i albo zniknie, albo wygeneruje dziwny kod.

Analogia: `&&` to **escape** — ten sam mechanizm, co `\\` w kodach formatów z modułu 10 albo `""` w CSV.

## 2. Teoria

### 2.1. Wymiary: szerokości kolumn i wysokości wierszy

Zacznijmy od elementu, który wizualnie robi najwięcej dla czytelności raportu.

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Wymiary"

# Szerokosc kolumny - w "znakach", nie w pikselach!
ws.column_dimensions["A"].width = 28
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 14

# Wysokosc wiersza - w PUNKTACH (1 punkt = 1/72 cala)
ws.row_dimensions[1].height = 32
```

**Co się dzieje w pamięci.** Indeksowanie `ws.column_dimensions["A"]` nie zwraca zwykłego elementu słownika — tworzy (albo zwraca istniejący) obiekt `ColumnDimension` z `reference="A"`. Ten obiekt ma atrybuty `width`, `hidden`, `outlineLevel`, `collapsed`, `bestFit`. Przypisanie `width = 28` ustawia atrybut na tym obiekcie. Analogicznie `ws.row_dimensions[1]` daje `RowDimension` z atrybutami `height`, `hidden`, `outlineLevel`, `collapsed`.

**Co trafi do pliku.** W `xl/worksheets/sheet1.xml` na początku dokumentu pojawi się sekcja `<cols>`:

```xml
<cols>
  <col min="1" max="1" width="28" customWidth="1"/>
  <col min="2" max="3" width="14" customWidth="1"/>
</cols>
```

Zwróć uwagę na dwie rzeczy:

1. **`min="2" max="3"`** — openpyxl **grupuje sąsiednie kolumny o tej samej szerokości** w jeden wpis `<col>`. To nie Ty decydujesz o tym grupowaniu; dzieje się automatycznie przy serializacji. Jeśli ustawisz kolumny B i C na 14, w pliku będzie **jeden** wpis obejmujący obie.
2. **`customWidth="1"`** — to flaga mówiąca Excelowi „ta szerokość jest ustawiona świadomie, nie jest domyślna". Bez niej Excel mógłby uznać szerokość za domyślną. openpyxl ustawia ją sam przy przypisaniu `width`.

**Dlaczego szerokość jest w „znakach", a nie w pikselach.** To historyczne dziedzictwo po arkuszach kalkulacyjnych z lat 80., gdzie podstawową jednostką była szerokość znaku w domyślnej czcionce. Dziś Excel interpretuje `width` jako **przybliżoną liczbę znaków, które zmieszczą się w komórce** przy domyślnej czcionce arkusza. To znaczy, że rzeczywista szerokość w pikselach zależy od rozdzielczości ekranu, DPI i wybranej czcionki.

**Konsekwencja praktyczna, którą warto zapamiętać:** **nie da się precyzyjnie dopasować szerokości kolumny do obrazka ani do konkretnej liczby pikseli.** Jeśli potrzebujesz dokładnego wyrównania (np. pod obrazek z logo — moduł 16), rób to empirycznie: ustaw, zapisz, otwórz w Excelu, sprawdź, popraw. Traktuj `width` jako **jednostkę względną**, nie miarę absolutną.

Sensowne wartości startowe dla typowego raportu:

| Zawartość kolumny | Sugerowana szerokość |
|---|---|
| Etykieta / nazwa | 25–35 |
| Kwota | 12–16 |
| Procent | 8–10 |
| Data `yyyy-mm-dd` | 12–14 |
| Data z nazwą miesiąca | 18–22 |
| Tekst opisowy (długi) | 40–50, z `wrap_text` (moduł 09) |
| Wąska flaga / TAK-NIE | 6–8 |

**Wysokość wiersza** działa podobnie, ale jednostką są **punkty typograficzne** (1 pt = 1/72 cala ≈ 0,35 mm). Domyślna wysokość to około 15 punktów.

⚠️ **Ważna pułapka z wysokością przy `wrap_text`.** W module 09 wspominaliśmy, że `Alignment(wrap_text=True)` zawija tekst w komórce. Ale openpyxl **nie ma modelu renderowania tekstu** — nie wie, ile miejsca zajmie zawinięty tekst, bo to zależy od czcionki i szerokości kolumny. Więc:

- **Jeśli nie ustawisz wysokości wiersza**, Excel przy otwarciu **zwykle sam** dopasuje wysokość do zawiniętego tekstu (bo w pliku nie ma jawnej wysokości, jest `customHeight="0"`).
- **Jeśli ustawisz wysokość jawnie** (`ws.row_dimensions[3].height = 30`), Excel uzna ją za wiążącą i **może obciąć tekst**, który się nie zmieści.

**Rekomendacja:** przy kolumnach z zawijanym tekstem **nie ustawiaj wysokości** — pozwól Excelowi dopasować. Wyjątek: wiersze nagłówka, gdzie chcesz jednolity wygląd i wiesz, ile linii tekstu będzie.

### 2.2. Grupowanie i ukrywanie: szuflady w szafce

W module 08 mówiliśmy o ukrywaniu całych **arkuszy**. Teraz ukrywamy **kolumny i wiersze** — i to jest narzędzie, które czyni raport czytelnym bez usuwania danych.

**Analogia:** wyobraź sobie szafkę z dokumentami. Dokumenty, których aktualnie nie potrzebujesz, wkładasz do **szuflady i zamykasz**. Nie wyrzucasz ich — są pod ręką, gdy zajdzie potrzeba.

```python
# ==================================================================
# UKRYWANIE - najprostsza forma
# ==================================================================
ws.column_dimensions["F"].hidden = True     # kolumna F ukryta
ws.row_dimensions[12].hidden = True          # wiersz 12 ukryty

# ==================================================================
# GRUPOWANIE (outline) - z przyciskiem "+" nad kolumnami
# ==================================================================
# group(start, end, outline_level=1, hidden=False)
ws.column_dimensions.group("B", "D", hidden=True)     # kolumny B-D w grupie, zwinięte
ws.row_dimensions.group(1, 10, hidden=True)           # wiersze 1-10 w grupie, zwinięte

# Można tez grupowac pojedyncza kolumne - podajac tylko start
ws.column_dimensions.group("H", hidden=True)
```

**Czym różni się `hidden` od `group(hidden=True)`?** Ukrycie jest **ciche** — użytkownik nie widzi, że coś jest schowane, chyba że zauważy przeskok w literach kolumn. Grupowanie jest **widoczne**: pojawia się przycisk `+`/`−` nad (lub obok) kolumnami, którym użytkownik może rozwinąć grupę.

**To różnica psychologiczna, nie techniczna** — ale w raportach ma ogromne znaczenie:

| Mechanizm | Widoczność | Kiedy używać |
|---|---|---|
| `hidden = True` | Ukryte bez śladu | Kolumny czysto techniczne: klucze, identyfikatory, pola pomocnicze |
| `group(..., hidden=True)` | Widoczny przycisk rozwijania | Dane, które mogą być potrzebne: szczegóły, kolumny dodatkowe |

**Ustawienia outline na poziomie arkusza.** W sekcji properties są cztery atrybuty sterujące wyglądem grupowania:

```python
# Gdzie umieszczony jest przycisk podsumowania
ws.sheet_properties.outlinePr.summaryBelow = True    # przycisk POD grupa wierszy
ws.sheet_properties.outlinePr.summaryRight = True    # przycisk PO grupie kolumn

# Czy stosowac style do zwinietych grup
ws.sheet_properties.outlinePr.applyStyles = True

# Czy pokazywac symbole outline (+/-)
ws.sheet_properties.outlinePr.showOutlineSymbols = True
```

W przeciwieństwie do `pageSetUpPr` (o którym za chwilę), `outlinePr` **jest inicjalizowany domyślnie** — więc możesz modyfikować jego atrybuty bezpośrednio, bez tworzenia obiektu. To jedna z tych niespójności w API openpyxl, które trzeba po prostu znać.

**Wielopoziomowe grupowanie.** `group()` ma parametr `outline_level` (domyślnie 1). Możesz budować hierarchię — grupa w grupie:

```python
# Poziom 1: kolumny szczegolowe
ws.column_dimensions.group("C", "E", outline_level=1, hidden=True)
# Poziom 2: nadrzedna grupa obejmujaca poziom 1
ws.column_dimensions.group("C", "H", outline_level=2, hidden=True)
```

⚠️ **Ostrzeżenie:** wielopoziomowe grupowanie działa, ale jest **rzadko potrzebne** i trudne do przewidzenia wizualnie. Zbuduj prostą wersję, zapisz, otwórz w Excelu i **sprawdź, czy wygląda tak, jak myślałeś** — nie zakładaj. To dokładnie ta kategoria z modułu 03: „openpyxl potrafi częściowo", więc weryfikacja wzrokowa jest obowiązkowa.

**Grupy a wiersze nagłówka.** Jeśli grupujesz wiersze 1–10, a wiersz 1 to nagłówek tabeli, ukryjesz nagłówek razem z danymi. Sprawdź, czy `summaryBelow`/`summaryRight` dają układ, którego chcesz.

### 2.3. Linie siatki i obramowanie — dwa różne światy

Wracamy do analogii rusztowania. Technicznie:

```python
ws.sheet_view.showGridLines = False    # wylacz szara siatke
```

**Co trafi do pliku.** W `xl/worksheets/sheet1.xml` w sekcji `<sheetViews>` pojawi się `<sheetView showGridLines="0" .../>`. Domyślnie atrybut jest pomijany (bo domyślnie `1`).

**I teraz najważniejsze.** Wyłączenie siatki **nie daje żadnego obramowania**. Uzyskujesz zupełnie „goły" arkusz — białe tło, czarny tekst, zero podziałów. To może być efekt, który chcesz (minimalistyczny dashboard), ale częściej chcesz **własną tabelę**.

Obramowanie robisz z modułu 09:

```python
from openpyxl.styles import Border, Side

cienka = Side(style="thin", color="FF999999")
obramowanie = Border(left=cienka, right=cienka, top=cienka, bottom=cienka)

for wiersz in ws.iter_rows(min_row=1, max_row=20, min_col=1, max_col=8):
    for komorka in wiersz:
        komorka.border = obramowanie
```

⚠️ **Uwaga na niemutowalność z modułu 09.** `Border` jest niemutowalny. Jeśli chcesz obramowanie tylko z jednym bokiem, musisz **zbudować nowy obiekt** z tym bokiem:

```python
# ZLE - to nie zadziala (modul 09!)
# komorka.border.bottom = Side(style="thin")

# DOBRZE - tworzymy NOWY obiekt Border z jednym bokiem
komorka.border = Border(bottom=Side(style="thin"))
```

I ostatnia rzecz — **linie siatki a druk.** Domyślnie linie siatki **nie drukują się**, nawet gdy są widoczne na ekranie. Możesz to zmienić jawnie:

```python
ws.print_options.gridLines = True     # drukuj linie siatki
```

Ale **nie rób tego w raportach**. Drukowanie siatki Excela daje szare, nierówne linie, które nie dochodzą do krawędzi tabeli i wyglądają niechlujnie. **Własne obramowanie jest jedynym profesjonalnym rozwiązaniem.**

### 2.4. Scalanie i rozscalanie komórek

Pełna teoria, bo to najbardziej niebezpieczna operacja tego modułu.

```python
# ==================================================================
# SCALANIE
# ==================================================================
ws.merge_cells("A1:D1")            # scal zakres (notacja zakresu)
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)   # to samo

# ==================================================================
# ROZSCALANIE
# ==================================================================
ws.unmerge_cells("A1:D1")

# ==================================================================
# INSPEKCJA
# ==================================================================
print(ws.merged_cells.ranges)      # lista scalonych zakresow, np. ['A1:D1', 'A3:A5']
```

#### 2.4.1. Co się dzieje z wartościami — dokładnie

To jest sedno. Załóżmy, że masz cztery wartości:

```python
ws["A1"] = "Raport"
ws["B1"] = "kwartalny"
ws["C1"] = "2026"
ws["D1"] = "Q1"
```

Teraz scalamy:

```python
ws.merge_cells("A1:D1")
```

**Co się dzieje w pamięci (od razu, nie przy zapisie):**

- `ws["A1"].value` → nadal `"Raport"`.
- `ws["B1"]`, `ws["C1"]`, `ws["D1"]` → to teraz **obiekty `MergedCell`**, których `value` to `None`. Wartości `"kwartalny"`, `"2026"`, `"Q1"` zostały **usunięte**.

Openpyxl **nie przechowuje** informacji o tym, co tam było. Nie ma kosza, nie ma `undo`. To jest **odwracalne tylko przez cofnięcie w kodzie**.

**Co trafi do pliku.** Do `xl/worksheets/sheet1.xml` trafi wpis:

```xml
<mergeCells count="1">
  <mergeCell ref="A1:D1"/>
</mergeCells>
```

Oraz — co ważne — komórki `B1`, `C1`, `D1` w pliku **w ogóle nie powstaną** jako pełne komórki albo powstaną jako puste. To już nie jest kwestia „wartość jest ukryta" — to kwestia „wartości nie ma w pliku".

#### 2.4.2. Odczyt: tylko lewa-górna ma wartość

Gdy wczytasz plik ze scalonym zakresem:

```python
wb = load_workbook("scalony.xlsx")
ws = wb.active

print(ws["A1"].value)      # "Raport"
print(ws["B1"].value)      # None
print(type(ws["B1"]))      # <class 'openpyxl.cell.cell.MergedCell'>
```

**To klasyczna pomyłka diagnostyczna.** Iterujesz po wierszu, sprawdzasz `if cell.value is not None` i **nie widzisz** tego, co w Excelu jest wyświetlone. Bo Excel **rysuje** „Raport" rozciągnięty na cztery kolumny, ale w danych jest tylko w jednej.

**Konsekwencja praktyczna:** jeśli piszesz kod czytający cudzy plik i chcesz znać „wartość efektywną" scalonej komórki, musisz **sam** sprawdzić, w którym scalonym zakresie się znajdujesz, i cofnąć się do lewej-górnej komórki:

```python
def wartosc_efektywna(ws, adres: str):
    """Zwraca wartosc komorki, rozumiejac scalone zakresy.

    Jesli komorka jest czescia scalonego zakresu, zwraca wartosc
    z komorki lewej-gornej tego zakresu.
    """
    komorka = ws[adres]
    if komorka.value is not None:
        return komorka.value

    # Sprawdz, czy komorka nalezy do jakiegos scalonego zakresu
    for zakres in ws.merged_cells.ranges:
        if adres in zakres:
            # zakres.top i zakres.left to wspolrzedne lewej-gornej komorki
            return ws.cell(row=zakres.min_row, column=zakres.min_col).value
    return None
```

To zapowiedź modułu 18 (modyfikacja istniejących plików) — tam będziesz czytać pliki, które tworzył człowiek w Excelu, a tam scalenia są nagminne.

#### 2.4.3. Zapis do scalonej komórki kończy się błędem

```python
ws.merge_cells("A5:C5")
ws["A5"] = "Naglowek"      # OK - lewa-gorna
ws["B5"] = "Cos"           # AttributeError!
# AttributeError: 'MergedCell' object attribute 'value' is read-only
```

**To dobra wiadomość.** Openpyxl **krzyczy**, zamiast cicho ignorować Twoją próbę. Wyjątek jest natychmiastowy i jednoznaczny — od razu wiesz, że trafiłeś w scalony obszar.

Ale uwaga: to nie znaczy, że jesteś bezpieczny. **Nie krzyczy przy odczycie** — `ws["B5"].value` zwróci `None` bez ostrzeżenia. Więc:

- **Zapis do scalonej komórki** → `AttributeError` (dobrze).
- **Odczyt ze scalonej komórki** → `None` (cicho, ryzyko błędu logicznego).

#### 2.4.4. Bezpieczny wzorzec scalania

```python
def scal_z_zachowaniem_tresci(ws, zakres: str, tekst: str) -> None:
    """Scala zakres i wpisuje tekst do lewej-gornej komorki.

    Robi to w bezpiecznej kolejnosci: najpierw scala (co czysci
    pozostale komorki), potem wpisuje wartosc do lewej-gornej.
    """
    from openpyxl.utils.cell import range_boundaries

    min_col, min_row, _, _ = range_boundaries(zakres)
    ws.merge_cells(zakres)
    ws.cell(row=min_row, column=min_col).value = tekst


# Uzycie:
scal_z_zachowaniem_tresci(ws, "A1:D1", "Raport kwartalny Q1 2026")
```

**Dlaczego ta kolejność?** Bo openpyxl ma funkcję `_clean_merge_range`, która **przy scalaniu czyści wartości** z komórek innych niż lewa-górna. Jeśli więc najpierw wpiszesz tekst, a potem scalisz, a Twój tekst trafił do `B1` — zostanie skasowany. Jeśli najpierw scalisz, a potem wpiszesz do `A1` — wszystko jest na miejscu i działa.

⚠️ **Ale uwaga na niuans:** jeśli scalasz zakres, w którym **już są dane** w komórkach innych niż lewa-górna, **te dane zostaną bezpowrotnie utracone** — niezależnie od kolejności. Ten wzorzec chroni Cię tylko przed utratą tekstu, który sam wpisujesz. Nie chroni przed utratą danych wcześniej wypełnionych.

### 2.5. `freeze_panes` — jedna reguła dla wszystkich przypadków

Powtórzmy regułę z analogii 1.4, teraz w pełni formalnie.

**`ws.freeze_panes` to adres komórki, od której zaczyna się obszar przewijalny.**

- Wszystko **powyżej** tego adresu → zamrożone wiersze.
- Wszystko **na lewo** od tego adresu → zamrożone kolumny.
- Sama komórka-znacznik → przewijalna (należy do obszaru ruchomego).

```python
ws.freeze_panes = "A2"     # zamrozone: wiersz 1
ws.freeze_panes = "B1"     # zamrozone: kolumna A
ws.freeze_panes = "B2"     # zamrozone: wiersz 1 + kolumna A
ws.freeze_panes = "C3"     # zamrozone: wiersze 1-2 + kolumny A-B
ws.freeze_panes = "A1"     # zamrozone: NIC (sposob na usuniecie zamrozenia)
ws.freeze_panes = None     # zamrozone: NIC
```

**Co się dzieje w pamięci.** `freeze_panes` to deskryptor na `Worksheet`, który przekazuje wartość do `ws.sheet_view.pane`. Openpyxl parsuje adres na współrzędne i konfiguruje pane (podział okna).

**Co trafi do pliku.** W `<sheetViews>`:

```xml
<sheetView tabSelected="1" workbookViewId="0">
  <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
  <selection pane="bottomLeft"/>
</sheetView>
```

Zwróć uwagę na `state="frozen"` — to odróżnia **zamrożenie** od **podziału okna** (`state="split"`), o którym za chwilę.

**Trzy fakty praktyczne:**

1. **`freeze_panes` jest ustawieniem arkusza, nie komórki.** Ustaw raz i działa cały czas. Iteruj po `wb.worksheets`, jeśli chcesz na wszystkich arkuszach.
2. **Zamrożenie nic nie kosztuje.** To metadata arkusza, nie dane. Arkusz z milionem wierszy zamraża nagłówek dokładnie tak samo tanio jak arkusz z dziesięcioma. **Nie ma żadnego powodu, żeby pomijać zamrożenie w dużych raportach** — a to najczęstszy brakujący element w automatycznie generowanych plikach.
3. **Jeśli wczytujesz i modyfikujesz istniejący plik, zamrożenie zwykle przechodzi bez zmian** (moduł 18). Ale jeśli **nadpiszesz** `ws.sheet_view` albo zbudujesz arkusz od nowa, musisz je ustawić ponownie.

#### 2.5.1. Podział okna — kuzyn zamrożenia

Openpyxl potrafi też ustawić **podział okna** (`state="split"`) — okno dzieli się linią, którą użytkownik może **przeciągać**, i obie części przewijają się **niezależnie**. To inna funkcja niż zamrożenie (gdzie części są zsynchronizowane).

⚠️ API do podziału okna jest mniej oczywiste i rzadko potrzebne. **Jeśli naprawdę go potrzebujesz — sprawdź w dokumentacji swojej wersji openpyxl**, jak ustawia się `ws.sheet_view.pane` z `state="split"`. W praktyce raportowej `freeze_panes` rozwiązuje 99% potrzeb.

### 2.6. Autofiltr — sitko nałożone na tabelę

**Analogia:** wyobraź sobie, że masz stos dokumentów i **sitko**, które nakładasz na wierzch. Możesz przez nie przesiewać — pokazać tylko dokumenty z jednego działu. Dokumenty pozostają na stosie, ale **widzisz** tylko wybrane.

Autofiltr to dokładnie to: **nakładka na zakres komórek**, która pozwala użytkownikowi filtrować i sortować, nie zmieniając danych.

```python
# ==================================================================
# PODSTAWY
# ==================================================================
ws.auto_filter.ref = "A1:H100"          # zakres, na ktorym dziala filtr

# ==================================================================
# FILTROWANIE PROGRAMOWE - ustawienie filtra przy zapisie
# ==================================================================
ws.auto_filter.add_filter_column(
    3,                                   # indeks kolumny (0-based, wzgledem zakresu!)
    ['Kowalska Anna', 'Nowak Piotr'],    # wartosci do pokazania
    blank=False,                         # czy pokazywac puste
)

# ==================================================================
# SORTOWANIE PROGRAMOWE
# ==================================================================
ws.auto_filter.add_sort_condition(
    "D2:D100",                           # zakres sortowania
    descending=True,                     # malejaco
)
```

⚠️ **Pułapka indeksowania.** `add_filter_column` przyjmuje **indeks liczony od zera, względem lewej kolumny zakresu `ref`**. Czyli jeśli `ref = "A1:H100"`, to indeks `0` = kolumna A, indeks `3` = kolumna D. **To inna konwencja niż w całym pozostałym API openpyxl**, gdzie kolumny liczy się od 1. Musisz o tym pamiętać — i to jest miejsce, w którym łatwo się pomylić o jeden.

⚠️ **Uwaga o wartości filtra.** Wartości w `add_filter_column` muszą **dokładnie odpowiadać** temu, co jest w komórkach — łącznie z typem. Jeśli w kolumnie są liczby `2026`, a Ty podasz `"2026"` (napis), filtr może nie zadziałać zgodnie z oczekiwaniem. To pokrewieństwo problemu `1` vs `"1"` z modułu 05.

#### 2.6.1. Autofiltr kontra tabela (zapowiedź modułu 12)

To jest pytanie, które zadaje sobie każdy, kto zaczyna automatyzację Excela. Odpowiedź w jednym zdaniu, a szczegóły w module 12:

| Cecha | Autofiltr (`ws.auto_filter`) | Tabela (`Table`, moduł 12) |
|---|---|---|
| Co to jest | Nakładka na **zakres** komórek | **Obiekt** arkusza z nazwą i tożsamością |
| Nazwa | Brak | `displayName`, np. `SprzedazTabela` |
| Structured references | Nie | Tak: `=SUM(SprzedazTabela[Kwota])` |
| Automatyczne rozszerzanie | Nie — dopisany wiersz jest **poza** zakresem | Tak — tabela rozszerza się przy dopisywaniu |
| Style | Trzeba nadać ręcznie | Gotowe style (`TableStyleInfo`) |
| Wiersz podsumowania | Nie | Tak (choć obsługa w openpyxl jest ograniczona) |
| Kiedy używać | Prosty, jednorazowy filtr na zakresie | Gdy dane to **stały, rosnący zbiór** i mają na nich stać formuły |

**Analogia:** autofiltr to **sitko nałożone na stos kartek**. Tabela to **segregator z etykietami i przegrodami** — ma nazwę, ma tożsamość, „sam" się rozszerza, gdy dokładasz kartki, i możesz się do niego odwoływać po nazwie.

**Reguła kciuka:** jeśli dane są statyczne i tylko je prezentujesz — autofiltr wystarczy. Jeśli dane rosną i inne arkusze się do nich odwołują — użyj tabeli.

### 2.7. Ustawienia widoku: siatka, formuły, zoom, aktywna zakładka

Poza zamrożeniem, `ws.sheet_view` pozwala ustawić kilka rzeczy, które decydują, co użytkownik zobaczy **w chwili otwarcia pliku**:

```python
ws.sheet_view.showGridLines = False       # wylacz szara siatke
ws.sheet_view.showFormulas = True         # pokaz FORMULY zamiast wynikow
ws.sheet_view.showZeros = False           # pokaz zera jako puste komorki
ws.sheet_view.zoomScale = 85              # zoom 85% (zakres 10-400)
ws.sheet_view.zoomScaleNormal = 100       # zoom po powrocie z widoku strony
ws.sheet_view.tabSelected = True          # czy ta zakladka jest aktywna
ws.sheet_view.showRowColHeaders = False   # ukryj naglowki A/B/1/2 (rzadko!)
```

**Kilka uwag praktycznych:**

- **`showGridLines = False` stosuj niemal zawsze** w raportach. Wyjątek: arkusze robocze, gdzie siatka pomaga Tobie nawigować.
- **`showFormulas = True` to świetne narzędzie diagnostyczne** — użytkownik otwiera plik i widzi `=SUM(A1:A10)` zamiast liczby. **Nie zostawiaj tego włączonego w raporcie dla klienta**, ale użyj przy własnym debugowaniu (moduł 06).
- **`showZeros = False`** sprawia, że komórki z wartością `0` wyglądają jak puste. To **wizualnie** przydatne w raportach o rzadkich danych, ale **niebezpieczne** — odbiorca nie odróżni „zero" od „brak danych". W module 10 mówiliśmy o trzeciej sekcji formatu (`;—;`) jako lepszym rozwiązaniu: pokazuje myślnik dla zera, a puste komórki pozostawia puste. **Używaj formatu, nie `showZeros`.**

**Ustawienie aktywnej zakładki i komórki.** Dwa niezależne mechanizmy:

```python
# Ktory arkusz jest AKTYWNY przy otwarciu
wb.active = 1                     # indeks arkusza (0-based!)

# Ktora KOMORKA jest zaznaczona przy otwarciu
ws.sheet_view.selection[0].activeCell = "B5"
ws.sheet_view.selection[0].sqref = "B5"
```

⚠️ Wybór aktywnej komórki to szczegół, który rzadko jest potrzebny, a API bywa kapryśne. **Sprawdź w dokumentacji swojej wersji**, jeśli chcesz go użyć. W praktyce wystarczy ustawić `wb.active`, żeby użytkownik otwierał plik na arkuszu podsumowania, a nie na surowych danych.

### 2.8. Drukowanie i eksport do PDF — serce modułu

To najobszerniejsza część i ta, która w praktyce daje najwięcej. Zasada nadrzędna:

> **Wszystkie ustawienia druku są przechowywane w pliku `.xlsx`.** Możesz je przygotować na serwerze bez Excela (openpyxl!), a Excel albo LibreOffice użyje ich przy druku albo eksporcie do PDF. **Drukowanie i PDF to jedne z niewielu funkcji, gdzie openpyxl daje pełną kontrolę.**

#### 2.8.1. Orientacja i rozmiar papieru

```python
# Orientacja - dwie stale/stringi
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE    # poziomo
# ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT   # pionowo (domyslna)

# Rozmiar papieru - stala numeryczna
ws.page_setup.paperSize = ws.PAPERSIZE_A4               # A4
# ws.page_setup.paperSize = ws.PAPERSIZE_A3             # A3
# ws.page_setup.paperSize = ws.PAPERSIZE_A5             # A5
# ws.page_setup.paperSize = ws.PAPERSIZE_LETTER         # Letter (US)
```

**Co się dzieje w pamięci.** `ws.ORIENTATION_LANDSCAPE` i `ws.PAPERSIZE_A4` to **stałe zdefiniowane na klasie `Worksheet`**. Możesz ich używać przez `ws.`, ale są to po prostu wartości liczbowe — `paperSize` to indeks w tabeli rozmiarów OOXML (A4 = 9). Możesz też wpisać liczbę wprost, ale **używaj stałych**, bo są czytelne.

**Kiedy poziomo:** gdy tabela ma więcej niż ~6–8 kolumn albo gdy kolumny są szerokie (opisy, daty z nazwą miesiąca). **Kiedy pionowo:** raporty jednostronicowe, dokumenty tekstowe, zestawienia z wąskimi kolumnami.

#### 2.8.2. Marginesy

```python
# Wartosci w calach (inches). 1 cal = 2.54 cm.
ws.page_margins.left = 0.5
ws.page_margins.right = 0.5
ws.page_margins.top = 0.75
ws.page_margins.bottom = 0.75
ws.page_margins.header = 0.3           # odleglosc naglowka od gornej krawedzi
ws.page_margins.footer = 0.3           # odleglosc stopki od dolnej krawedzi
```

⚠️ **Trzy rzeczy, o których warto wiedzieć:**

1. **Jednostką są cale**, nie centymetry ani punkty. Jeśli myślisz w centymetrach: 1 cal ≈ 2,54 cm, więc 0,5 cala ≈ 1,27 cm.
2. **Domyślne wartości** to (w przybliżeniu) 0,7 cala na boki i 0,75 cala na górę i dół. **Sprawdź w swojej wersji openpyxl**, jakie są dokładne wartości domyślne — nie polegaj na moim przybliżeniu.
3. **Nagłówek i stopka muszą się zmieścić w marginesach!** Jeśli `header` (odległość nagłówka od górnej krawędzi) jest większa lub równa `top` (górny margines treści), nagłówek **wejdzie na dane**. **Reguła: `page_margins.header < page_margins.top`.**

Ostatnia uwaga z punktu 3 to typowa „niewyjaśnialna" usterka — nagłówek nachodzi na pierwszy wiersz tabeli, a nikt nie rozumie dlaczego. To jest przyczyna.

#### 2.8.3. Obszar wydruku: kadr kamery

**Analogia:** `print_area` to **kadr kamery**, który mówi, co wchodzi w zdjęcie. Wszystko poza kadrem nie zostanie sfotografowane — nawet jeśli istnieje.

```python
# Jeden zakres
ws.print_area = "A1:H50"

# Wiele zakresow - mozna podac liste (drukowane jako oddzielne strony)
ws.print_area = ["A1:H50", "J1:N50"]
```

**Dlaczego to jest ważne dla raportów generowanych programowo.** Bardzo często masz w arkuszu **kolumny robocze** — klucze, identyfikatory, pola pomocnicze do obliczeń, techniczne flagi. Są potrzebne w danych, ale **nie powinny być widoczne na wydruku**.

Masz trzy sposoby, żeby je ukryć:

| Metoda | Efekt na ekranie | Efekt na wydruku |
|---|---|---|
| `hidden = True` | Kolumna ukryta | Nie drukowana |
| `print_area` bez tej kolumny | Kolumna widoczna | Nie drukowana |
| `group(hidden=True)` | Kolumna zwinięta, przycisk `+` | Nie drukowana |

**`print_area` jest najbezpieczniejszy**, gdy chcesz zachować kolumny widoczne na ekranie (bo użytkownik ich potrzebuje), ale nie drukować ich. **`hidden` jest lepszy**, gdy kolumny są czysto techniczne i nikt ich nie potrzebuje.

⚠️ **Bardzo ważna uwaga o `print_area` i pustych stronach.** Jeśli `print_area` obejmuje zakres **większy niż dane**, do PDF trafią **puste strony**. Dlatego:

```python
# ZLE - recznie wpisany, za duzy zakres
ws.print_area = "A1:H1000"          # jesli danych jest 200 wierszy - puste strony!

# DOBRZE - wyliczony z danych
from openpyxl.utils import get_column_letter
ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
```

Ale pamiętaj z modułu 07: **`ws.max_row` kłamie, gdy ktoś sformatował komórki poza danymi**. Jeśli więc formatujesz kolumnę do wiersza 1000 „na zapas", `max_row` zwróci 1000 i dostaniesz puste strony. **To nie jest problem z `print_area` — to problem z formatowaniem „na zapas".** Wniosek: nie formatuj komórek, których nie zamierzasz wypełnić (to też temat wydajności z modułu 19).

#### 2.8.4. Powtarzanie nagłówka na każdej stronie — funkcja, o której wszyscy zapominają

To **jedna linia kodu, która robi więcej dla czytelności wielostronicowego raportu niż wszystko inne.**

**Analogia:** pomyśl o słowniku. Na górze każdej strony masz powtórzone litery alfabetu — dzięki temu wiesz, gdzie jesteś, nie wracając do początku. **Bez tego, strona 7 twojego raportu to gołe liczby bez kolumn.**

```python
ws.print_title_rows = "1:1"        # powtarzaj wiersz 1 na kazdej stronie
# ws.print_title_rows = "1:2"      # powtarzaj wiersze 1-2

ws.print_title_cols = "A:B"        # powtarzaj kolumny A-B na kazdej stronie
# ws.print_title_cols = "A:A"      # powtarzaj tylko kolumne A
```

⚠️ **Format jest kluczowy i jest to najczęstsza pomyłka:**

```python
ws.print_title_rows = "1:1"     # ✅ DOBRZE - format "od:do"
ws.print_title_rows = "1"       # ❌ ZLE - Excel moze nie zrozumiec
ws.print_title_rows = 1         # ❌ ZLE - to liczba, nie zakres
```

**Musi to być napis w formacie `"start:koniec"`** — tak jak zapisujesz zakres w Excelu. `"1:1"`, nie `"1"`. To dokładnie ta sama rodzina problemów co `print_title_cols = "A:B"` (nie `"A"`).

**A teraz rzecz, która łączy się z 2.8.5:** `print_title_cols` ma sens **tylko wtedy**, gdy tabela dzieli się na kilka stron **w poziomie**. Jeśli włączysz dopasowanie do jednej strony szerokości (`fitToWidth = 1`), tabela **nigdy** nie podzieli się w poziomie — więc `print_title_cols` jest zbędne. Nie zaszkodzi, ale też nic nie da.

#### 2.8.5. Dopasowanie do stron: trzy ustawienia, które muszą działać razem

Wracamy do analogii kserokopiarki z 1.5, tym razem z kodem.

```python
from openpyxl.worksheet.properties import PageSetupProperties

# USTAWIENIE 1: ile stron szerokosci
ws.page_setup.fitToWidth = 1        # wszystko na jednej stronie szerokosci

# USTAWIENIE 2: ile stron wysokosci (0 = bez limitu!)
ws.page_setup.fitToHeight = 0       # tyle stron w dol, ile trzeba

# USTAWIENIE 3: WŁĄCZ TRYB dopasowania - bez tego dwa powyzsze sa ignorowane!
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
```

**Trzy scenariusze i co ustawić:**

| Chcesz… | `fitToWidth` | `fitToHeight` | `fitToPage` |
|---|---|---|---|
| Wszystko na **jednej stronie** (dashboard, KPI) | `1` | `1` | `True` |
| Jedna strona **szerokości**, dowolna wysokość (długi raport) | `1` | `0` | `True` |
| Wszystko na jednej stronie **wysokości** (rzadko) | `0` | `1` | `True` |
| Bez dopasowania (użyj `scale`) | — | — | `False` / nie ustawiaj |

**`fitToHeight = 0` znaczy „bez limitu", nie „zero stron".** To druga najczęstsza pomyłka. Gdybyś chciał ograniczyć do dwóch stron wysokości, ustawiłbyś `fitToHeight = 2`.

**O subtelności z `PageSetupProperties` — mówię wprost, bo dokumentacja bywa tu myląca.** Dokumentacja openpyxl zawiera notkę, że „page setup properties nie są inicjalizowane domyślnie i trzeba najpierw utworzyć obiekt `PageSetupProperties`". Jednak **kod źródłowy `WorksheetProperties.__init__` w 3.1.x inicjalizuje ten obiekt**:

```python
if pageSetUpPr is None:
    pageSetUpPr = PageSetupProperties()
self.pageSetUpPr = pageSetUpPr
```

Co to znaczy w praktyce? **Obie formy działają w 3.1.x:**

```python
# Forma A: jawna, samodokumentujaca sie (POLECANA)
from openpyxl.worksheet.properties import PageSetupProperties
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# Forma B: bezposrednie przypisanie (dziala, bo obiekt jest juz utworzony)
ws.sheet_properties.pageSetUpPr.fitToPage = True
```

**Polecam formę A**, bo: (a) jest jednoznaczna i widoczna w kodzie, (b) nie zależy od tego, czy wewnętrzna inicjalizacja się zmieni, (c) w starszych wersjach openpyxl forma B **rzucała `AttributeError`** — a jeśli Twój kod ma działać też na starszych instalacjach, forma A jest bezpieczna. Ale **wiedz, że forma B w 3.1.5 działa** — dokumentacja jest tu po prostu nieaktualna.

#### 2.8.6. Skala: alternatywa dla dopasowania

Zamiast „dopasuj do stron" możesz ustawić **stałą skalę** — dokładnie jak na kserokopiarce:

```python
ws.page_setup.scale = 80        # 80% rozmiaru (zakres 10-400)
```

⚠️ **Skala i dopasowanie wykluczają się.** Jeśli `fitToPage = True`, Excel **ignoruje** `scale`. Jeśli `fitToPage = False` (lub nieustawione), Excel używa `scale`.

**Kiedy co stosować:**

- **`fitTo*`** — gdy chcesz mieć pewność, że tabela zmieści się na określonej liczbie stron, niezależnie od jej rozmiaru. To typowy wybór dla raportów generowanych automatycznie, bo **nie wiesz z góry, ile będzie danych**.
- **`scale`** — gdy chcesz zachować **jednolity wygląd** na wszystkich stronach i nie przeszkadza Ci, że tabela rozjedzie się na kilka stron. To typowy wybór dla dokumentów o stałej strukturze.

#### 2.8.7. Nagłówki i stopki: osobny język

Wracamy do analogii pól w formularzu z 1.6.

```python
# Trzy sekcje kazdego naglowka/stopki: left, center, right
ws.oddHeader.center.text = "Raport sprzedażowy 2026"
ws.oddHeader.center.size = 12
ws.oddHeader.center.font = "Calibri,Bold"
ws.oddHeader.center.color = "FF333333"

# Stopka z numeracja stron
ws.oddFooter.left.text = "&D"                    # data
ws.oddFooter.center.text = "Strona &P z &N"      # numer strony / liczba stron
ws.oddFooter.right.text = "&A"                   # nazwa arkusza
```

**Trzy pary nagłówek/stopka:** `oddHeader`/`oddFooter` (strony nieparzyste), `evenHeader`/`evenFooter` (parzyste), `firstHeader`/`firstFooter` (pierwsza strona). Każda ma trzy sekcje: `.left`, `.center`, `.right`, a każda sekcja ma `.text`, `.size`, `.font`, `.color`.

⚠️ **Uwaga: `evenHeader`/`evenFooter` zadziałają tylko wtedy, gdy włączysz drukowanie różnych stron parzystych i nieparzystych.** W Excelu to opcja „Different odd & even pages" w ustawieniach strony. Openpyxl **zapisuje** treść, ale **włączenie tej opcji** to osobne ustawienie — **sprawdź w dokumentacji swojej wersji**, czy jest dostępne przez `ws.print_options` albo pokrewne. W praktyce najczęściej używa się `oddHeader`/`oddFooter`.

**Pełna lista kodów:** ⚠️ Uczciwe ostrzeżenie — **dokładna lista obsługiwanych kodów zależy od wersji Excela i od tego, czy plik jest otwierany w Excelu, LibreOffice czy Google Sheets.** openpyxl przekazuje tekst bez walidacji. Poniżej najbezpieczniejszy, powszechnie działający rdzeń:

| Kod | Znaczenie |
|---|---|
| `&P` | Numer bieżącej strony |
| `&N` | Łączna liczba stron |
| `&D` | Bieżąca data |
| `&T` | Bieżąca godzina |
| `&F` | Nazwa pliku |
| `&A` | Nazwa arkusza |
| `&Z` | Ścieżka do pliku |
| `&&` | **Dosłowny ampersand** |
| `&L` / `&C` / `&R` | Przejdź do sekcji lewej / środkowej / prawej |
| `&"czcionka,styl"` | Ustaw czcionkę, np. `&"Arial,Bold"` |
| `&nn` | Rozmiar czcionki w punktach, np. `&14` |

**Dwie rzeczy, o których warto pamiętać:**

1. **`&[Page]` i `&P` to oba poprawne sposoby zapisu numeru strony.** Dokumentacja openpyxl używa w przykładzie `&[Page]`. Excel zapisuje w XML klasyczną formę `&P`. **Sprawdź w swoim Excelu, która forma wyświetla się poprawnie** — większość wersji rozumie obie.
2. **`&` musi być escape'owany jako `&&`**, jeśli chcesz go zobaczyć dosłownie. Stopka `Firma & Syn` wyświetli się dziwnie albo zgubi znak. Poprawnie: `Firma && Syn`.

**Nagłówki i stopki są widoczne tylko w podglądzie wydruku** — w normalnym widoku arkusza ich nie zobaczysz. To frustrujące, ale tak działa Excel: nagłówek to element **druku**, nie ekranu.

#### 2.8.8. Podział strony: prawdziwe API

⚠️ **Uwaga na błąd powszechny w poradnikach.** W wielu miejscach w internecie zobaczysz `ws.page_breaks` jako sposób na ustawienie podziału strony. **Taki atrybut w openpyxl nie istnieje.** Prawdziwe API to **dwa oddzielne obiekty**:

```python
from openpyxl.worksheet.pagebreak import Break

# Podzial WIERSZY (nowa strona po wierszu N)
ws.row_breaks.append(Break(id=20))       # po wierszu 20 zaczyna sie nowa strona

# Podzial KOLUMN (nowa strona po kolumnie N)
ws.col_breaks.append(Break(id=5))        # po kolumnie 5 zaczyna sie nowa strona
```

**Kluczowa semantyka: `Break(id=n)` wstawia podział PO wierszu/kolumnie `n`.** To znaczy, że wiersz `n` jest **ostatnim** na bieżącej stronie, a `n+1` zaczyna nową. To jest miejsce, w którym łatwo pomylić się o jeden — i skutek jest taki, że **każda sekcja zaczyna się z jednym osieroconym wierszem poprzedniej grupy**.

Praktyczny wzorzec — podział przy zmianie wartości w kolumnie (raport grupowy):

```python
from openpyxl.worksheet.pagebreak import Break


def podziel_przy_zmianie(ws, kolumna: int, pierwszy_wiersz: int = 2) -> int:
    """Wstawia podzial strony przy kazdej zmianie wartosci w kolumnie.

    Zwraca liczbe wstawionych podzialow.
    """
    poprzednia = None
    wstawione = 0

    for wiersz in range(pierwszy_wiersz, ws.max_row + 1):
        wartosc = ws.cell(row=wiersz, column=kolumna).value
        if wartosc is None:
            continue
        if poprzednia is not None and wartosc != poprzednia:
            # Break(id=wiersz-1) -> nowa strona ZACZYNA sie od 'wiersz'
            ws.row_breaks.append(Break(id=wiersz - 1))
            wstawione += 1
        poprzednia = wartosc

    return wstawione
```

**Trzy ograniczenia, które trzeba znać:**

1. **Excel ma limit ręcznych podziałów strony na arkusz: 1026.** Powyżej tego limitu kolejne podziały są **cicho ignorowane**. Przy raporcie z tysiącami grup trzeba to zabezpieczyć:

```python
liczba_grup = 5000
if liczba_grup > 1000:
    print(f"{liczba_grup} grup - za duzo na podzialy; pomijam")
else:
    podziel_przy_zmianie(ws, kolumna=1)
```

2. **`Break(id=n)` nie przesuwa danych.** Wstawiasz znacznik „tu kończy się strona" — nie dzielisz arkusza na części. Wszystkie dane zostają w jednym arkuszu, tylko druk je rozdziela.
3. **Podział strony działa razem z `fitTo*`, ale może dać nieoczekiwane wyniki.** Jeśli `fitToWidth = 1, fitToHeight = 0`, a Ty wstawisz podziały, Excel uszanuje podziały, ale **nie dopasuje skali do nich** — może wyjść brzydko. **Zawsze otwórz podgląd wydruku i sprawdź.**

### 2.9. Wyśrodkowanie wydruku, siatka na wydruku

Trzy drobne, ale przydatne ustawienia:

```python
ws.print_options.horizontalCentered = True    # wysrodkuj tabele w poziomie
ws.print_options.verticalCentered = True      # wysrodkuj tabele w pionie
ws.print_options.gridLines = True             # DRUKUJ linie siatki (rzadko!)
```

**Wyśrodkowanie jest bardzo przydatne w raportach jednostronicowych** (dashboard, karta KPI) — tabela na środku strony wygląda znacznie lepiej niż przyklejona do lewego marginesu.

**`gridLines = True` — powtórzę ostrzeżenie z 2.3: prawie nigdy tego nie chcesz.** Wyjątek: arkusze robocze dla zespołu wewnętrznego, gdzie celem jest czytelność techniczna, nie estetyka.

### 2.10. Wszystko razem: gotowy „layout raportu"

To jest funkcja, którą powinieneś mieć w każdym projekcie. Zbiera wszystkie ustawienia z tego modułu w jedno miejsce — i jest realizacją analogii „typograf" z 1.1.

```python
"""Warstwa prezentacji: jeden layout dla kazdego arkusza raportu."""

from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.utils import get_column_letter


def zastosuj_layout_raportu(
    ws,
    tytul_raportu: str = "",
    wiersze_naglowka: str = "1:1",
    szerokosci: dict[str, int] | None = None,
) -> None:
    """Nakłada na arkusz kompletny, gotowy do druku layout raportu.

    Parametry:
        ws: arkusz openpyxl
        tytul_raportu: tekst do naglowka wydruku
        wiersze_naglowka: zakres wierszy powtarzanych na kazdej stronie
        szerokosci: mapowanie litera kolumny -> szerokosc
    """
    # --- 1. WIDOK EKRANU -------------------------------------------
    ws.sheet_view.showGridLines = False      # rusztowanie wylaczone (1.2)
    ws.sheet_view.zoomScale = 100            # nie odziedziczaj cudzego zoomu
    ws.freeze_panes = "A2"                   # naglowek kolumn przyklejony

    # --- 2. WYMIARY -------------------------------------------------
    if szerokosci:
        for litera, szerokosc in szerokosci.items():
            ws.column_dimensions[litera].width = szerokosc

    # --- 3. STRONA --------------------------------------------------
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    # Marginesy (w calach). Naglowek MUSI byc mniejszy niz gorny margines.
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.7
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    # --- 4. DOPASOWANIE --------------------------------------------
    # Jedna strona szerokosci, dowolna liczba stron w dol.
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    # TA LINIA JEST OBOWIAZKOWA - bez niej dwie powyzsze sa ignorowane!
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    # --- 5. NAGLOWEK I STOPKA --------------------------------------
    if tytul_raportu:
        ws.oddHeader.center.text = tytul_raportu
        ws.oddHeader.center.size = 12
        ws.oddHeader.center.font = "Calibri,Bold"

    ws.oddFooter.left.text = "&D"
    ws.oddFooter.center.text = "Strona &P z &N"
    ws.oddFooter.right.text = "&A"

    # --- 6. POWTARZANIE NAGLOWKA -----------------------------------
    ws.print_title_rows = wiersze_naglowka


def ustaw_obszar_wydruku(ws, margines_wierszy: int = 0, margines_kolumn: int = 0) -> str:
    """Wyznacza obszar wydruku na podstawie realnych danych i zwraca go.

    UWAGA: korzysta z max_row/max_column, ktore bywaja zawyzone przez
    formatowanie "na zapas" (modul 07). W razie problemow uzyj wlasnej
    funkcji skanujacej.
    """
    ostatni_wiersz = ws.max_row - margines_wierszy
    ostatnia_kolumna = get_column_letter(ws.max_column - margines_kolumn)
    obszar = f"A1:{ostatnia_kolumna}{ostatni_wiersz}"
    ws.print_area = obszar
    return obszar
```

**Co się dzieje w pamięci.** Ta funkcja **nie tworzy żadnych danych** — tylko ustawia atrybuty na obiekcie arkusza. Wywołanie jej na arkuszu z 10 000 wierszy jest tak samo tanie jak na arkuszu z jednym.

**Co trafi do pliku.** Wszystkie ustawienia trafią do `xl/worksheets/sheetN.xml` w sekcjach `<sheetViews>` (widok, zamrożenie), `<cols>` (szerokości), `<pageMargins>`, `<pageSetup>`, `<headerFooter>`, `<printOptions>`. **Nic z tego nie jest w komórkach** — to metadane arkusza.

**Dlaczego to jest ważne dla architektury (zapowiedź modułu 26):** ta funkcja to prototyp **warstwy prezentacji**. Oddziela „jak wygląda dokument" od „co jest w dokumencie". W module 26 stanie się metodą strategii albo elementem konfiguracji z YAML.

## 3. Przykłady krok po kroku

### Przykład 1 — arkusz gotowy do druku: A4 poziomo z powtarzanym nagłówkiem

Ćwiczenie 🟢 z briefu, wykonane porządnie i z weryfikacją.

```python
"""Arkusz gotowy do druku: A4 poziomo + powtarzany naglowek.

Uruchom:  python examples/11_gotowy_do_druku.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "11_gotowy_do_druku.xlsx"

# --- dane testowe: 12 miesiecy x 6 kolumn ---------------------------
NAGLOWKI = ["Miesiąc", "Sprzedaż PLN", "Koszt PLN", "Marża PLN", "Marża %", "Sztuki"]

DANE = [
    ("Styczeń",   120_000, 78_000, 42_000, 0.35, 1_240),
    ("Luty",       98_500, 70_920, 27_580, 0.28,   980),
    ("Marzec",    143_700, 122_145, 21_555, 0.15, 1_510),
    ("Kwiecień",  131_200, 76_096, 55_104, 0.42, 1_330),
    ("Maj",       156_800, 94_080, 62_720, 0.40, 1_620),
    ("Czerwiec",  148_300, 91_946, 56_354, 0.38, 1_480),
    ("Lipiec",    112_400, 78_680, 33_720, 0.30, 1_090),
    ("Sierpień",  104_900, 76_577, 28_323, 0.27,   990),
    ("Wrzesień",  167_500, 105_525, 61_975, 0.37, 1_740),
    ("Październik", 178_200, 108_702, 69_498, 0.39, 1_860),
    ("Listopad",  189_600, 117_552, 72_048, 0.38, 1_950),
    ("Grudzień",  204_100, 130_624, 73_476, 0.36, 2_120),
]

# Style (z modulu 09) i formaty (z modulu 10) w rejestrach - jedno zrodlo prawdy
STYL_NAGLOWKA = {
    "font": Font(bold=True, color="FFFFFFFF", size=11),
    "fill": PatternFill(fill_type="solid", fgColor="FF2F5496"),
    "align": Alignment(horizontal="center", vertical="center", wrap_text=True),
}

FORMATY = {
    "kwota":   '#,##0 "zł"',
    "procent": "0.0%",
    "liczba":  "#,##0",
}

CIENKA = Side(style="thin", color="FFB4C6E7")
OBRAMOWANIE = Border(left=CIENKA, right=CIENKA, top=CIENKA, bottom=CIENKA)


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaż 2026"

    # --- naglowek tabeli -------------------------------------------
    for indeks, tekst in enumerate(NAGLOWKI, start=1):
        komorka = ws.cell(row=1, column=indeks, value=tekst)
        komorka.font = STYL_NAGLOWKA["font"]
        komorka.fill = STYL_NAGLOWKA["fill"]
        komorka.alignment = STYL_NAGLOWKA["align"]
        komorka.border = OBRAMOWANIE

    # --- dane ------------------------------------------------------
    for przesuniecie, wiersz_danych in enumerate(DANE):
        wiersz = 2 + przesuniecie
        for kolumna, wartosc in enumerate(wiersz_danych, start=1):
            komorka = ws.cell(row=wiersz, column=kolumna, value=wartosc)
            komorka.border = OBRAMOWANIE

    # --- formaty ---------------------------------------------------
    ostatni = 1 + len(DANE)
    for wiersz in range(2, ostatni + 1):
        ws.cell(row=wiersz, column=2).number_format = FORMATY["kwota"]
        ws.cell(row=wiersz, column=3).number_format = FORMATY["kwota"]
        ws.cell(row=wiersz, column=4).number_format = FORMATY["kwota"]
        ws.cell(row=wiersz, column=5).number_format = FORMATY["procent"]
        ws.cell(row=wiersz, column=6).number_format = FORMATY["liczba"]

    # --- wiersz sumy (formuly - Excel policzy przy otwarciu) --------
    wiersz_sumy = ostatni + 1
    ws.cell(row=wiersz_sumy, column=1, value="RAZEM")
    for kolumna in (2, 3, 4, 6):
        litera = get_column_letter(kolumna)
        komorka = ws.cell(
            row=wiersz_sumy,
            column=kolumna,
            value=f"=SUM({litera}2:{litera}{ostatni})",
        )

    # --- WYGLAD ARKUSZA (modul 11!) --------------------------------
    # 1. Rusztowanie wylaczone - mamy wlasne obramowanie.
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100

    # 2. Szerokosci kolumn - w "znakach".
    for litera, szerokosc in zip("ABCDEF", (16, 15, 15, 15, 12, 11)):
        ws.column_dimensions[litera].width = szerokosc
    ws.row_dimensions[1].height = 30        # naglowek wyzszy na 2 linie

    # 3. Naglowek przyklejony przy przewijaniu.
    ws.freeze_panes = "B2"                  # wiersz 1 + kolumna A

    # 4. Autofiltr na calym zakresie danych (z naglowkiem).
    ostatnia_litera = get_column_letter(len(NAGLOWKI))
    ws.auto_filter.ref = f"A1:{ostatnia_litera}{wiersz_sumy}"

    # 5. Strona: A4 poziomo.
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    # 6. Marginesy (w calach). header < top - inaczej naglowek wchodzi na dane.
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.7
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3

    # 7. Dopasowanie: jedna strona szerokosci, dowolna liczba w dol.
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    # 8. Naglowek i stopka.
    ws.oddHeader.center.text = "Raport sprzedaży 2026"
    ws.oddHeader.center.size = 12
    ws.oddHeader.center.font = "Calibri,Bold"
    ws.oddFooter.left.text = "&D"
    ws.oddFooter.center.text = "Strona &P z &N"
    ws.oddFooter.right.text = "&A"

    # 9. POWTARZANIE NAGLOWKA - jedno z najwazniejszych ustawien!
    ws.print_title_rows = "1:1"

    # 10. Obszar wydruku (wyliczony, nie na sztywno).
    ws.print_area = f"A1:{ostatnia_litera}{wiersz_sumy}"

    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    """Weryfikacja: co naprawde zapisalo sie do pliku."""
    wb = load_workbook(cel)
    try:
        ws = wb["Sprzedaż 2026"]
        print("=" * 88)
        print("WERYFIKACJA USTAWIEN PO ZAPISIE I ODCZYCIE")
        print("=" * 88)
        print(f"  siatka widoczna:        {ws.sheet_view.showGridLines}")
        print(f"  zamrozenie:             {ws.freeze_panes}")
        print(f"  autofiltr:              {ws.auto_filter.ref}")
        print(f"  orientacja:             {ws.page_setup.orientation}")
        print(f"  rozmiar papieru:        {ws.page_setup.paperSize} (A4 = 9)")
        print(f"  fitToWidth/fitToHeight: {ws.page_setup.fitToWidth}/{ws.page_setup.fitToHeight}")
        print(f"  fitToPage (WŁĄCZNIK!):  {ws.sheet_properties.pageSetUpPr.fitToPage}")
        print(f"  powtarzane wiersze:     {ws.print_title_rows!r}")
        print(f"  obszar wydruku:         {ws.print_area!r}")
        print(f"  szerokosc kolumny A:    {ws.column_dimensions['A'].width}")
        print(f"  naglowek stopki:        {ws.oddFooter.center.text!r}")
    finally:
        wb.close()

    print()
    print("=" * 88)
    print("CO SPRAWDZIC W EXCELU (Ctrl+P)")
    print("=" * 88)
    print("  1. Naglowek tabeli POJAWIA SIE na kazdej stronie wydruku.")
    print("     (gdyby go nie bylo -> print_title_rows zle ustawione)")
    print("  2. Tabela NIE dzieli sie w poziomie na kilka stron.")
    print("     (gdyby sie dzielila -> brak fitToPage=True)")
    print("  3. W stopce widac 'Strona 1 z 1' - nie literalne '&P z &N'.")
    print("  4. Siatka Excela zniknela - sa tylko wlasne obramowania tabeli.")
    print("  5. Przy przewijaniu w dol wiersz 1 zostaje, kolumna A zostaje.")
    print()
    print("  Eksport do PDF: Ctrl+P -> wybierz drukarke 'Microsoft Print to PDF'")
    print("  albo w LibreOffice: plik -> eksportuj jako PDF.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Kluczowe punkty do przemyślenia:**

1. **Punkt 7 to cała istota pułapki z 1.5.** Trzy linie: `fitToWidth`, `fitToHeight`, i `PageSetupProperties(fitToPage=True)`. Bez trzeciej dwie pierwsze są bezużyteczne.
2. **Punkt 9 (`print_title_rows = "1:1"`)** — to jedna linia, która zmienia wielostronicowy raport z „nieczytelnego" w „normalny". Zwróć uwagę na format `"1:1"`.
3. **Punkt 10 (`print_area` liczony, nie sztywny)** — `f"A1:{ostatnia_litera}{wiersz_sumy}"` zamiast `"A1:F100"`. Wyliczony obszar nigdy nie da pustych stron.
4. **Punkt 3 (`freeze_panes = "B2"`)** — wiersz 1 i kolumna A. Przy 12 wierszach na ekranie zamrożenie może wydać się zbędne, ale przy raporcie sprzedaży ma sens — a nawyk zostanie.
5. **Punkt 1 (`showGridLines = False`)** — bez tego obramowanie wygląda dziwnie, bo dochodzą do niego szare linie siatki.

**Co trafi do pliku.** Wszystkie dziesięć ustawień trafi do `xl/worksheets/sheet1.xml`, ale **do różnych sekcji**:

- `<sheetViews>` — `showGridLines="0"`, `<pane>`, `zoomScale`
- `<cols>` — szerokości kolumn
- `<autoFilter>` — zakres filtra
- `<pageMargins>` — marginesy
- `<pageSetup>` — orientacja, rozmiar papieru, `fitToWidth`, `fitToHeight`
- `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` — **włącznik trybu dopasowania**
- `<headerFooter>` — nagłówek i stopka
- `<printOptions>` — (brak, bo nic nie ustawiliśmy)
- `definedName` w `xl/workbook.xml` — `print_title_rows` i `print_area` są zapisywane jako **definiowane nazwy** (`_xlnm.Print_Titles`, `_xlnm.Print_Area`)! To zaskakujący szczegół — omówiony w pułapkach.

### Przykład 2 — scalanie nagłówka i dowód utraty danych

Ćwiczenie 🟡 z briefu, z pełną weryfikacją.

```python
"""Scalanie naglowka: co sie dzieje z wartosciami w pozostalych komorkach.

Uruchom:  python examples/11_scalanie.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import range_boundaries

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "11_scalanie.xlsx"


def pokaz_stan(ws, opis: str) -> None:
    """Wypisuje zawartosc A1:D1 oraz typy komorek."""
    print(f"  --- {opis} ---")
    for adres in ("A1", "B1", "C1", "D1"):
        komorka = ws[adres]
        typ = type(komorka).__name__
        print(f"    {adres}: value={komorka.value!r:<18} typ={typ}")
    print()


def main() -> None:
    print("=" * 92)
    print("EKSPERYMENT: SCALANIE KOMOREK I UTRATA DANYCH")
    print("=" * 92)
    print()

    wb = Workbook()
    ws = wb.active
    ws.title = "Scalanie"

    # --- KROK 1: wpisujemy cztery wartosci w cztery komorki ----------
    ws["A1"] = "Raport"
    ws["B1"] = "kwartalny"
    ws["C1"] = "2026"
    ws["D1"] = "Q1"

    pokaz_stan(ws, "KROK 1: cztery komorki, cztery wartosci (PRZED scaleniem)")

    # --- KROK 2: scalamy ---------------------------------------------
    ws.merge_cells("A1:D1")

    pokaz_stan(ws, "KROK 2: po scaleniu A1:D1 - TRZY WARTOSCI ZNIKNELY")

    print("  >>> WNIOSEK: scalanie usunelo 'kwartalny', '2026' i 'Q1'.")
    print("  >>> Nie ma kosza, nie ma undo. Dane sa bezpowrotnie utracone.")
    print("  >>> Komorki B1-D1 to teraz obiekty MergedCell z value=None.")
    print()

    # --- KROK 3: proba zapisu do scalonej komorki --------------------
    try:
        ws["B1"] = "proba zapisu"
        print("  (nieoczekiwanie) zapis sie udal")
    except AttributeError as blad:
        print(f"  >>> ZAPIS do scalonej komorki B1 -> wyjatek:")
        print(f"      {type(blad).__name__}: {blad}")
    print()

    # --- KROK 4: zapis i ponowny odczyt ------------------------------
    wb.save(CEL)
    wb.close()

    wb2 = load_workbook(CEL)
    try:
        ws2 = wb2["Scalanie"]
        pokaz_stan(ws2, "KROK 4: po zapisie i PONOWNYM odczycie")

        print(f"  scalone zakresy w pliku: {list(ws2.merged_cells.ranges)}")
    finally:
        wb2.close()

    # --- KROK 5: poprawny wzorzec -----------------------------------
    print()
    print("=" * 92)
    print("JAK ZROBIC TO POPRAWNIE")
    print("=" * 92)
    print()

    wb3 = Workbook()
    ws3 = wb3.active
    ws3.title = "Poprawnie"

    # Chcesz naglowek "Raport kwartalny 2026 Q1" przez cztery komorki?
    # NIE wpisuj go w czesciach. Zbuduj CALY tekst i wpisz do lewej-gornej.
    caly_tekst = "Raport kwartalny 2026 Q1"
    scal_z_tekstem(ws3, "A1:D1", caly_tekst)

    # Albo jeszcze lepiej: w ogole nie scalaj, tylko wlacz wrap_text
    # i pozwol tekstowi "przelewac sie" wizualnie - dane zostaja nietkniete.
    ws3["A5"] = "Raport kwartalny 2026 Q1"      # tekst w jednej komorce
    ws3["A5"].alignment = ws3["A5"].alignment.copy(horizontal="left")

    pokaz_stan(ws3, "KROK 5: poprawny wzorzec (caly tekst w A1)")

    wb3.save(OUTPUT / "11_scalanie_poprawnie.xlsx")
    wb3.close()

    print("  Zapisano: output/11_scalanie_poprawnie.xlsx")
    print()
    print("  Zasady bezpiecznego scalania:")
    print("    1. Zbuduj CALY tekst w Pythonie, wpisz do lewej-gornej komorki.")
    print("    2. Scal NA KONCU, po wpisaniu wszystkich danych.")
    print("    3. NIGDY nie scalaj komorek, ktore zawieraja dane do obliczen.")
    print("    4. Rozwaz, czy scalanie jest w ogole potrzebne.")


def scal_z_tekstem(ws, zakres: str, tekst: str) -> None:
    """Scala zakres i wpisuje tekst do lewej-gornej komorki.

    Kolejnosc ma znaczenie: najpierw scal (co wyczysci pozostale komorki),
    potem wpisz wartosc do lewej-gornej.
    """
    min_col, min_row, _, _ = range_boundaries(zakres)
    ws.merge_cells(zakres)
    ws.cell(row=min_row, column=min_col).value = tekst


if __name__ == "__main__":
    main()
```

**Co pokaże ten program:**

```
KROK 1: A1='Raport', B1='kwartalny', C1='2026', D1='Q1'   <- wszystko na miejscu
KROK 2: A1='Raport', B1=None, C1=None, D1=None            <- TRZY WARTOŚCI STRACONE
KROK 3: AttributeError: 'MergedCell' object attribute 'value' is read-only
KROK 4: to samo po zapisie i odczycie - utrata jest trwała
```

**Czego uczy najbardziej:** że scalanie nie jest „operacją kosmetyczną". To operacja **destrukcyjna**, która przy okazji zmienia typ komórek. I że **odczyt ze scalonej komórki jest cichy** (`None`), a **zapis głośny** (`AttributeError`). Ta asymetria jest źródłem błędów logicznych: kod czyta `None`, myśli „puste", i nie widzi, że w Excelu coś tam jest wyświetlone.

### Przykład 3 — grupowanie, ukrywanie i „szuflady"

```python
"""Grupowanie kolumn i ukrywanie danych technicznych.

Uruchom:  python examples/11_grupowanie.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "11_grupowanie.xlsx"

# Kolumny: A=Klient, B-S=sty/lut/mar..., T=Suma, U=VAT (techniczne)
MIESIACE = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze",
            "Lip", "Sie", "Wrz", "Paz", "Lis", "Gru"]

DANE = [
    ("Alfa Sp. z o.o.",   [12, 14, 11, 15, 18, 16, 13, 12, 19, 21, 22, 24]),
    ("Beta S.A.",         [8, 9, 10, 9, 7, 11, 12, 10, 13, 14, 12, 15]),
    ("Gamma sp.k.",       [20, 22, 19, 24, 26, 23, 21, 20, 27, 29, 31, 33]),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Grupowanie"

    # --- naglowek ---------------------------------------------------
    ws["A1"] = "Klient"
    for indeks, miesiac in enumerate(MIESIACE):
        ws.cell(row=1, column=2 + indeks, value=miesiac)
    kolumna_sumy = 2 + len(MIESIACE)          # T
    kolumna_vat = kolumna_sumy + 1            # U

    ws.cell(row=1, column=kolumna_sumy, value="Suma")
    ws.cell(row=1, column=kolumna_vat, value="VAT 23% (techn.)")

    # --- dane -------------------------------------------------------
    for przesuniecie, (klient, wartosci) in enumerate(DANE):
        wiersz = 2 + przesuniecie
        ws.cell(row=wiersz, column=1, value=klient)
        for indeks, wartosc in enumerate(wartosci):
            ws.cell(row=wiersz, column=2 + indeks, value=wartosc)
        litera_sumy = get_column_letter(kolumna_sumy)
        ws.cell(
            row=wiersz,
            column=kolumna_sumy,
            value=f"=SUM(B{wiersz}:{get_column_letter(1 + len(MIESIACE))}{wiersz})",
        )
        ws.cell(
            row=wiersz,
            column=kolumna_vat,
            value=f"=ROUND({litera_sumy}{wiersz}*0.23,2)",
        )

    # --- WYGLAD -----------------------------------------------------
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B2"                     # wiersz 1 + kolumna A

    ws.column_dimensions["A"].width = 22
    for indeks in range(len(MIESIACE)):
        ws.column_dimensions[get_column_letter(2 + indeks)].width = 7
    ws.column_dimensions[get_column_letter(kolumna_sumy)].width = 10
    ws.column_dimensions[get_column_letter(kolumna_vat)].width = 16

    # --- SZUFLADA 1: grupa miesiecy (zwinięta, z przyciskiem +) -----
    # group(start, end, outline_level=1, hidden=False)
    ws.column_dimensions.group(
        "B", get_column_letter(1 + len(MIESIACE)),
        outline_level=1,
        hidden=True,
    )

    # --- SZUFLADA 2: kolumna techniczna W UKRYCIU (bez sladu) ------
    ws.column_dimensions[get_column_letter(kolumna_vat)].hidden = True

    # Gdzie ma byc przycisk podsumowania grupy
    ws.sheet_properties.outlinePr.summaryRight = True      # przycisk po prawej

    # Obszar wydruku - bez technicznej kolumny
    ws.print_area = f"A1:{get_column_letter(kolumna_sumy)}{ws.max_row}"
    ws.print_title_rows = "1:1"

    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Grupowanie"]
        print("=" * 88)
        print("WERYFIKACJA GRUPOWANIA I UKRYWANIA")
        print("=" * 88)
        for litera in ("A", "B", "F", "S", "T", "U"):
            wymiar = ws.column_dimensions[litera]
            print(
                f"  kolumna {litera}: width={wymiar.width!r:<6} "
                f"hidden={wymiar.hidden!r:<6} "
                f"outlineLevel={wymiar.outlineLevel!r:<6} "
                f"collapsed={wymiar.collapsed!r}"
            )
        print()
        print("  Wnioski:")
        print("    B-F, S: outlineLevel=1, hidden=True  -> grupa zwinięta z przyciskiem")
        print("    U:      hidden=True bez outlineLevel -> ukryta bez śladu")
        print("    T:      widoczna (suma)")
        print()
        print(f"  obszar wydruku: {ws.print_area!r} - bez kolumny U (VAT techniczny)")
    finally:
        wb.close()

    print()
    print("=" * 88)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 88)
    print("  1. Nad kolumnami B-luty widać przycisk + (rozwiń grupę).")
    print("  2. Kolumna U (VAT) jest ukryta - między T a V jest przeskok.")
    print("  3. Przy rozwinieciu grupy wiersze sie nie rozjezdzaja.")
    print("  4. Podglad wydruku (Ctrl+P) NIE pokazuje kolumny U.")
    print()
    print("  Różnica psychologiczna:")
    print("    grupa (widoczny +)  -> uzytkownik WIE, ze cos jest schowane")
    print("    hidden (bez sladu)  -> uzytkownik NIE WIE, ze cos jest schowane")
    print("  Wybieraj swiadomie!")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Czego uczy najbardziej:** że **ukrywanie i grupowanie to dwie różne intencje.** Grupa mówi użytkownikowi „tu są szczegóły, rozwiń, jeśli chcesz". Ukrycie mówi „nie interesuj się". W raporcie dla klienta ukryjesz kolumny techniczne, a pogrupujesz szczegóły miesięczne. W raporcie wewnętrznym możesz pogrupować jedno i drugie.

I ważna rzecz techniczna: **`column_dimensions[litera].hidden` nie czyści `outlineLevel`.** W weryfikacji zobaczysz, że grupa i ukrycie to niezależne atrybuty tego samego obiektu. Możesz mieć kolumnę, która jest jednocześnie w grupie i ukryta.

### Przykład 4 — dashboard jednostronicowy

Ćwiczenie 🔴 z briefu.

```python
"""Dashboard jednostronicowy: freeze panes + ukryte kolumny pomocnicze.

Uruchom:  python examples/11_dashboard.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "11_dashboard.xlsx"

# Dane zrodlowe (docelowo z arkusza "Dane" - tu uproszczone)
REGIONY = [
    ("Mazowieckie",   1_240_000, 0.31, 4_180),
    ("Śląskie",         980_500, 0.28, 3_320),
    ("Wielkopolskie",   870_300, 0.35, 2_910),
    ("Małopolskie",     760_900, 0.33, 2_540),
    ("Dolnośląskie",    690_200, 0.29, 2_280),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    # ==================================================================
    # SEKCJA 1: TYTUL (scalony naglowek - ale budujemy CALY tekst!)
    # ==================================================================
    ws.merge_cells("A1:J1")
    tytul = ws["A1"]
    tytul.value = "SPRZEDAŻ WEDŁUG REGIONÓW — I PÓŁROCZE 2026"
    tytul.font = Font(bold=True, size=16, color="FF1F3864")
    tytul.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Podtytul
    ws.merge_cells("A2:J2")
    podtytul = ws["A2"]
    podtytul.value = "Wartość sprzedaży, marża i liczba transakcji"
    podtytul.font = Font(size=11, italic=True, color="FF595959")
    podtytul.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # ==================================================================
    # SEKCJA 2: TABELA GLOWNA (naglowek w wierszu 4)
    # ==================================================================
    naglowki = ["Region", "Sprzedaż", "Marża %", "Transakcje", "Śr. koszyk"]
    for indeks, tekst in enumerate(naglowki, start=1):
        komorka = ws.cell(row=4, column=indeks, value=tekst)
        komorka.font = Font(bold=True, color="FFFFFFFF", size=11)
        komorka.fill = PatternFill(fill_type="solid", fgColor="FF2F5496")
        komorka.alignment = Alignment(horizontal="center", vertical="center")
        komorka.border = Border(
            left=Side(style="thin", color="FF8EA9DB"),
            right=Side(style="thin", color="FF8EA9DB"),
            top=Side(style="thin", color="FF8EA9DB"),
            bottom=Side(style="thin", color="FF8EA9DB"),
        )
    ws.row_dimensions[4].height = 24

    for przesuniecie, (region, sprzedaz, marza, transakcje) in enumerate(REGIONY):
        wiersz = 5 + przesuniecie
        ws.cell(row=wiersz, column=1, value=region)
        ws.cell(row=wiersz, column=2, value=sprzedaz).number_format = '#,##0 "zł"'
        ws.cell(row=wiersz, column=3, value=marza).number_format = "0.0%"
        ws.cell(row=wiersz, column=4, value=transakcje).number_format = "#,##0"

        # ==============================================================
        # KOLUMNA POMOCNICZA (F) - sredni koszyk. Obliczana, ale widoczna.
        # ==============================================================
        k = ws.cell(row=wiersz, column=5, value=f"=B{wiersz}/D{wiersz}")
        k.number_format = '#,##0.00 "zł"'

    wiersz_sumy = 5 + len(REGIONY)
    ws.cell(row=wiersz_sumy, column=1, value="RAZEM").font = Font(bold=True)
    for kolumna in (2, 4):
        litera = get_column_letter(kolumna)
        komorka = ws.cell(
            row=wiersz_sumy, column=kolumna,
            value=f"=SUM({litera}5:{litera}{wiersz_sumy - 1})",
        )
        komorka.font = Font(bold=True)
    ws.cell(row=wiersz_sumy, column=2).number_format = '#,##0 "zł"'
    ws.cell(row=wiersz_sumy, column=4).number_format = "#,##0"
    ws.cell(row=wiersz_sumy, column=3).value = f"=AVERAGE(C5:C{wiersz_sumy - 1})"
    ws.cell(row=wiersz_sumy, column=3).number_format = "0.0%"
    ws.cell(row=wiersz_sumy, column=3).font = Font(bold=True)

    # ==================================================================
    # SEKCJA 3: KOLUMNY TECHNICZNE (ukryte)
    # ==================================================================
    # G: ranking (obliczany), H: klucz sortowania, I: % udzialu
    for przesuniecie in range(len(REGIONY)):
        wiersz = 5 + przesuniecie
        ws.cell(row=wiersz, column=7, value=f"=RANK(B{wiersz},$B$5:$B$9)")
        ws.cell(row=wiersz, column=8, value=f"=B{wiersz}")
        ws.cell(
            row=wiersz, column=9,
            value=f"=B{wiersz}/$B${wiersz_sumy}",
        ).number_format = "0.0%"

    ws.column_dimensions["G"].hidden = True
    ws.column_dimensions["H"].hidden = True
    ws.column_dimensions["I"].hidden = True

    # ==================================================================
    # WYGLAD I WYMIARY
    # ==================================================================
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 10

    # Zamrozenie: wiersz 4 (naglowek tabeli) - ale NIE wiersze 1-3!
    # Dlatego NIE uzywamy "A5" (zamrozilby tytul razem z naglowkiem).
    # Chcemy zamrozic naglowek tabeli, ale tytul moze uciekac -
    # albo zamrozic WSZYSTKO do wiersza 4 wlacznie: A5.
    ws.freeze_panes = "B5"      # wiersze 1-4 + kolumna A

    # ==================================================================
    # DRUK: wszystko na JEDNEJ stronie
    # ==================================================================
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2

    # JEDNA strona: i szerokosc, i wysokosc!
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    ws.oddFooter.center.text = "Dashboard — strona &P z &N"
    ws.oddFooter.center.size = 9

    # Wysrodkowanie na stronie
    ws.print_options.horizontalCentered = True

    ws.print_area = "A1:F10"    # bez ukrytych kolumn G-I

    wb.save(cel)
    wb.close()


def sprawdz(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Dashboard"]
        print("=" * 88)
        print("WERYFIKACJA DASHBOARDU")
        print("=" * 88)
        print(f"  zamrozenie:            {ws.freeze_panes}")
        print(f"  fitToWidth/Height:     {ws.page_setup.fitToWidth}/{ws.page_setup.fitToHeight}")
        print(f"  fitToPage:             {ws.sheet_properties.pageSetUpPr.fitToPage}")
        print(f"  wysrodkowanie:         {ws.print_options.horizontalCentered}")
        print(f"  obszar wydruku:        {ws.print_area!r}")
        print(f"  scalone zakresy:       {[str(z) for z in ws.merged_cells.ranges]}")
        print()
        print("  Kolumny techniczne:")
        for litera in ("G", "H", "I"):
            print(f"    {litera}: hidden={ws.column_dimensions[litera].hidden}")
    finally:
        wb.close()

    print()
    print("=" * 88)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 88)
    print("  1. Tytul 'SPRZEDAŻ WEDŁUG REGIONÓW' rozciaga sie nad tabela,")
    print("     ale w pliku wartosc siedzi TYLKO w A1 (scalenie!).")
    print("  2. Przy przewijaniu w dol: tytul, podtytul i naglowek tabeli")
    print("     zostaja (zamrozony A5 = wiersze 1-4).")
    print("  3. Przy przewijaniu w prawo: kolumna A (Region) zostaje.")
    print("  4. Kolumny G, H, I sa ukryte.")
    print("  5. Podglad wydruku: WSZYSTKO na jednej stronie A4 poziomo.")
    print("  6. Stopka: 'Dashboard — strona 1 z 1'.")
    print()
    print("  Uwaga na freeze_panes:")
    print("    'A5'  -> zamraza wiersze 1-4 (tytul + podtytul + naglowek)")
    print("    'B5'  -> zamraza wiersze 1-4 ORAZ kolumne A")
    print("    'A1'  -> NIE zamraza nic (czesty blad!)")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    sprawdz(CEL)


if __name__ == "__main__":
    main()
```

**Czego uczy najbardziej:** że **`freeze_panes` w dashboardzie wymaga decyzji**, której w zwykłej tabeli nie ma. W tabeli chcesz zamrozić wiersz 1. W dashboardzie masz tytuł (wiersz 1), podtytuł (wiersz 2), pusty wiersz (wiersz 3) i nagłówek tabeli (wiersz 4). **Czy chcesz, żeby tytuł uciekał przy przewijaniu, czy został?**

- `ws.freeze_panes = "A5"` — zostają wiersze 1–4 (cały nagłówek z tytułem).
- `ws.freeze_panes = "B5"` — jw. plus kolumna A.

I druga lekcja: **dopasowanie do jednej strony to `fitToWidth = 1` ORAZ `fitToHeight = 1`** — dla dashboardu chcesz mieć pewność, że zmieści się na jednej kartce. W długim raporcie ustawiłbyś `fitToHeight = 0`.

## 4. Anatomia API

| Metoda / atrybut | Co robi | Parametry / wartości | Uwagi |
|---|---|---|---|
| `ws.column_dimensions["B"].width` | Szerokość kolumny | liczba **w znakach** | Nie piksele! Bez `customWidth` Excel uzna za domyślną |
| `ws.column_dimensions["B"].hidden` | Ukrycie kolumny | `True` / `False` | Ukrycie ciche — bez przycisku rozwijania |
| `ws.column_dimensions.group(start, end=None, outline_level=1, hidden=False)` | Grupowanie kolumn | `start`/`end` jako litery (`"B"`, `"D"`) | Tworzy **widoczny** przycisk `+` |
| `ws.row_dimensions[1].height` | Wysokość wiersza | liczba **w punktach** (1 pt = 1/72 cala) | Jawna wysokość może obciąć `wrap_text` |
| `ws.row_dimensions[1].hidden` | Ukrycie wiersza | `True` / `False` | |
| `ws.row_dimensions.group(start, end=None, outline_level=1, hidden=False)` | Grupowanie wierszy | numery wierszy | |
| `ws.sheet_properties.outlinePr.summaryBelow/summaryRight` | Położenie przycisku grupy | `bool` | `outlinePr` **jest** inicjalizowany domyślnie |
| `ws.sheet_view.showGridLines` | Szara siatka | `bool` (domyślnie `True`) | **Nie jest** obramowaniem! Wyłączaj w raportach |
| `ws.sheet_view.zoomScale` | Zoom arkusza | `10`–`400` (procent) | Ustaw, żeby nie dziedziczyć cudzego zoomu |
| `ws.sheet_view.zoomScaleNormal` | Zoom „normalny" | `10`–`400` | |
| `ws.sheet_view.showFormulas` | Pokaż formuły | `bool` | Świetne do diagnostyki, nie dla klienta |
| `ws.sheet_view.showZeros` | Pokaż zera | `bool` | Lepiej użyć formatu `;—;` (moduł 10) |
| `ws.sheet_view.tabSelected` | Czy zakładka aktywna | `bool` | |
| `ws.freeze_panes` | Zamrożenie okien | Adres, np. `"B2"` | **Zamraża wszystko powyżej i na lewo!** `"A1"` = nic |
| `ws.merge_cells(zakres)` | Scalanie komórek | `"A1:D1"` lub `start_row=..., ...` | **Usuwa wartości** z komórek ≠ lewa-górna |
| `ws.unmerge_cells(zakres)` | Rozscalanie | jak wyżej | **Nie przywraca** wartości! |
| `ws.merged_cells.ranges` | Lista scalonych zakresów | — | Iterowalna lista obiektów `CellRange` |
| `ws.auto_filter.ref` | Zakres autofiltra | `"A1:H100"` | Bez niego filtry nie zadziałają |
| `ws.auto_filter.add_filter_column(index, values, blank=False)` | Filtr programowy | **`index` liczony od 0!** | Inna konwencja niż reszta API |
| `ws.auto_filter.add_sort_condition(zakres, descending=False)` | Sortowanie | `"D2:D100"` | |
| `ws.page_setup.orientation` | Orientacja | `ws.ORIENTATION_LANDSCAPE` / `_PORTRAIT` | Stałe na klasie `Worksheet` |
| `ws.page_setup.paperSize` | Rozmiar papieru | `ws.PAPERSIZE_A4` (=9), `_A3`, `_A5`, `_LETTER` | Wewnętrznie liczba indeksu OOXML |
| `ws.page_setup.fitToWidth` | Strony szerokości | `int` (`1` = jedna) | **Działa tylko z `fitToPage=True`** |
| `ws.page_setup.fitToHeight` | Strony wysokości | `int` (**`0` = bez limitu**) | Nie „zero stron"! |
| `ws.page_setup.scale` | Stała skala | `10`–`400` | **Ignorowane** gdy `fitToPage=True` |
| `ws.sheet_properties.pageSetUpPr` | **Włącznik dopasowania** | `PageSetupProperties(fitToPage=True)` | **Kluczowa linia!** |
| `ws.page_margins.left/right/top/bottom/header/footer` | Marginesy | liczby **w calach** | `header` musi być `< top` |
| `ws.print_area` | Obszar wydruku | `"A1:H50"` albo lista zakresów | Za duży zakres → **puste strony** |
| `ws.print_title_rows` | Powtarzane wiersze | **`"1:1"`** (format `start:koniec`) | Zapisane jako `_xlnm.Print_Titles` |
| `ws.print_title_cols` | Powtarzane kolumny | `"A:B"` | Zbędne przy `fitToWidth = 1` |
| `ws.print_options.horizontalCentered` / `verticalCentered` | Wyśrodkowanie na stronie | `bool` | Przydatne w dashboardach |
| `ws.print_options.gridLines` | Drukuj siatkę | `bool` | Prawie nigdy nie chcesz `True` |
| `ws.oddHeader.left/center/right` | Nagłówek wydruku | `.text`, `.size`, `.font`, `.color` | Widoczny **tylko** w podglądzie wydruku |
| `ws.oddFooter.left/center/right` | Stopka wydruku | jak wyżej | `&P`, `&N`, `&D`, `&A`, `&&` |
| `ws.evenHeader` / `ws.firstHeader` | Nagłówki parzyste / pierwszej strony | jak `oddHeader` | Wymaga włączonej opcji w Excelu |
| `ws.row_breaks.append(Break(id=n))` | Podział strony po wierszu `n` | `Break` z `openpyxl.worksheet.pagebreak` | ⚠️ **`ws.page_breaks` NIE ISTNIEJE** |
| `ws.col_breaks.append(Break(id=n))` | Podział strony po kolumnie `n` | jak wyżej | Limit ~1026 podziałów na arkusz |
| `wb.active = index` | Aktywny arkusz przy otwarciu | indeks **od 0** | |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — arkusz gotowy do druku na A4 poziomo z powtarzanym nagłówkiem.**

Napisz `examples/11_cw1.py`, który tworzy `output/11_cw1.xlsx`:

1. Arkusz o nazwie `Raport` z tabelą: 3 kolumny (Nazwa, Data, Kwota), 25 wierszy danych wygenerowanych w pętli.
2. Nagłówek w wierszu 1 z nazwami kolumn, pogrubiony i z wypełnieniem.
3. Wszystkie komórki (nagłówek + dane) z obramowaniem `thin` (moduł 09).
4. Kolumna `Data` z formatem `[$-415]dd.mm.yyyy`, kolumna `Kwota` z formatem `#,##0.00 "zł"` (moduł 10).
5. Wyłączone linie siatki.
6. **Orientacja pozioma, papier A4.**
7. **`print_title_rows = "1:1"`** — powtarzanie nagłówka na każdej stronie.
8. Marginesy: 0,4 cala na boki, 0,7 cala góra/dół, `header` 0,3 cala.
9. Stopka: po lewej data, po środku `Strona &P z &N`.
10. Zamrożenie `"A2"`.

Następnie dodaj drugą funkcję, która **wczytuje plik ponownie** i wypisuje:

- `ws.page_setup.orientation`,
- `ws.page_setup.paperSize`,
- `ws.print_title_rows`,
- `ws.oddFooter.center.text`,
- `ws.freeze_panes`,
- `ws.sheet_view.showGridLines`.

W pliku `output/11_cw1_wnioski.md` odpowiedz:

1. **Co zobaczysz w stopce na wydruku?** Napisz dosłownie, jak to będzie wyglądać dla strony 1. Wyjaśnij, dlaczego `&P` nie wyświetla się jako „&P".
2. **Otwórz plik w Excelu i wejdź w podgląd wydruku (`Ctrl+P`).** Czy nagłówek tabeli powtarza się na każdej stronie? Jaki byłby efekt, gdybyś wpisał `print_title_rows = "1"` zamiast `"1:1"`? Sprawdź eksperymentalnie.
3. **Dlaczego nagłówek i stopka nie są widoczne w normalnym widoku arkusza?** Kiedy je widzisz?
4. **Czy linie siatki Excela wydrukowałyby się bez `showGridLines = False`?** Sprawdź w podglądzie — i wyjaśnij różnicę między „siatka widoczna" a „siatka drukowana".

**Zadanie 2 — scalanie nagłówka: co się dzieje z wartościami.**

Napisz `examples/11_cw2.py`, który:

1. Tworzy arkusz i wpisuje cztery wartości w `A1`, `B1`, `C1`, `D1`.
2. Wypisuje ich wartości i **typy** komórek.
3. Scala zakres `A1:D1`.
4. Wypisuje wartości i typy ponownie.
5. Próbuje **wpisać** wartość do `B1` i obsługuje `AttributeError`, wypisując treść błędu.
6. Zapisuje plik, wczytuje go ponownie i wypisuje wartości z `A1:D1` po odczycie.

W pliku `output/11_cw2_wnioski.md`:

1. **Które wartości zostały utracone i w którym momencie** — przy scalaniu czy przy zapisie? Uzasadnij na podstawie obserwacji, nie teorii.
2. **Jaki typ mają komórki `B1:D1` po scaleniu?** Dlaczego to nie jest zwykły `Cell`?
3. **Dlaczego odczyt ze scalonej komórki zwraca `None`, a zapis rzuca wyjątek?** Jaką ma to konsekwencję dla kodu, który **czyta** cudze pliki? Opisz scenariusz, w którym ten `None` prowadzi do błędnej decyzji.
4. **Jak zrobić nagłówek „Raport kwartalny 2026 Q1" nad czterema kolumnami, nie tracąc danych?** Podaj dwa sposoby (jeden ze scaleniem, jeden bez) i oceń, który jest lepszy w raporcie automatycznym.
5. **Czy `unmerge_cells` przywraca utracone wartości?** Sprawdź eksperymentalnie i opisz wynik.

### 🟡 Warsztat

**Zadanie 3 — pełny layout raportu z konfiguracją.**

Rozbuduj funkcję `zastosuj_layout_raportu` z sekcji 2.10 do wersji konfigurowalnej.

**Część A.** Napisz `examples/11_cw3.py` z:

1. Klasą `LayoutRaportu` (albo `@dataclass`) trzymającą konfigurację:
   - `orientacja` (`"pozioma"` / `"pionowa"` → mapowane na stałe Excela),
   - `rozmiar_papieru` (`"A4"` / `"A3"` / `"LETTER"` → mapowane na stałe),
   - `tryb_dopasowania` (`"jedna_strona"` / `"szerokosc"` / `"brak"` → mapowane na `fitToWidth`/`fitToHeight`/`fitToPage`),
   - `tytul_naglowka`, `wiersze_naglowka`, `szerokosci_kolumn`, `wysrodkuj`.
2. Metodą `zastosuj(ws)` nakładającą konfigurację na arkusz.
3. Funkcją `zbuduj_raport(tytul, dane, layout)` tworzącą arkusz i nakładającą layout.
4. Skrypt demonstracyjny generujący **trzy** pliki o różnych layoutach:
   - `output/11_cw3_dashboard.xlsx` — jedna strona, poziomo, wyśrodkowane,
   - `output/11_cw3_dlugi.xlsx` — szerokość dopasowana, wysokość bez limitu,
   - `output/11_cw3_pionowo.xlsx` — pionowo, bez dopasowania (`scale = 85`).

**Część B — weryfikacja.** Dodaj funkcję `weryfikuj_layout(sciezka)`, która wczytuje plik i sprawdza, czy **wszystkie** ustawienia z konfiguracji faktycznie się zapisały. Wypisuje tabelę `ustawienie | oczekiwane | w pliku | OK/BŁĄD` i podsumowanie `wynik: X/Y`.

W pliku `output/11_cw3_wnioski.md`:

1. **Dlaczego warto mapować `"pozioma"` → `ws.ORIENTATION_LANDSCAPE`, zamiast wpisywać stałą wprost w kodzie?** Podaj dwa argumenty (jeden o zmianie w przyszłości, jeden o czytelności).
2. **Czym różni się `"szerokosc"` od `"jedna_strona"` w `trybie_dopasowania`?** Podaj konkretne wartości `fitToWidth`, `fitToHeight`, `fitToPage` i scenariusz użycia dla każdego.
3. **Dlaczego dla trybu `"brak"` trzeba ustawić `fitToPage = False`** (a nie tylko pominąć tę linię)? Co by się stało, gdybyś wczytał istniejący plik, który miał `fitToPage = True`, i chciał wyłączyć dopasowanie?
4. **W `weryfikuj_layout` porównujesz wartości z pliku z konfiguracją.** Które porównania wymagały konwersji (np. `"A4"` → `9`)? Dlaczego to pokazuje, że **konfiguracja w formie czytelnej dla człowieka** wymaga warstwy tłumaczącej? To zapowiedź modułu 26 — zapisz tę myśl.
5. **Dodaj do konfiguracji pole `kolumny_ukryte: list[str]` i przetestuj.** Ile miejsc w kodzie musiałeś zmienić? Czy `weryfikuj_layout` sam wykrył nowe ustawienie, czy trzeba było go dopisać?

### 🔴 Wyzwanie

**Zadanie 4 — dashboard jednostronicowy z kolumnami pomocniczymi i audytem druku.**

Napisz `examples/11_cw4.py` implementujący **generator dashboardu** i **audyt gotowości do druku**.

```python
"""Generator dashboardu + audyt ustawien druku."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class SekcjaDashboardu:
    """Opis jednej sekcji dashboardu (tabela z naglowkiem i danymi)."""
    tytul: str
    naglowki: list[str]
    wiersze: list[list]
    formaty: list[str | None] = field(default_factory=list)   # per kolumna
    wiersz_sumy: bool = False


def zbuduj_dashboard(
    sciezka: Path,
    tytul: str,
    podtytul: str,
    sekcje: list[SekcjaDashboardu],
    kolumny_pomocnicze: dict[str, int] | None = None,
) -> None:
    """Buduje jednostronicowy dashboard z sekcjami i ukrytymi kolumnami."""
    ...


def audyt_druku(sciezka: Path) -> dict:
    """Sprawdza, czy plik jest gotowy do druku/PDF - zwraca raport.

    Kontroluje:
      - czy fitToPage jest wlaczone, jesli ustawiono fitToWidth/fitToHeight,
      - czy header < top (inaczej naglowek wchodzi na dane),
      - czy print_area nie jest wiekszy niz realne dane (puste strony!),
      - czy print_title_rows ma poprawny format ('1:1', nie '1'),
      - czy freeze_panes nie jest ustawione na 'A1' (nie zamraza nic),
      - czy nie ma komorek sformatowanych poza obszarem druku,
      - ile stron wyjdzie przy druku (szacunek z wymiarow i ustawien).
    """
    ...


def raport_audytu_tekstowy(raport: dict) -> str:
    """Formatuje raport audytu jako czytelny tekst."""
    ...
```

**Wymagania funkcjonalne:**

1. **`zbuduj_dashboard`** musi:
   - zbudować tytuł (scalony **z pełnym tekstem w lewej-górnej komórce**!),
   - umieścić sekcje jedna pod drugą, z odstępami,
   - nadać formaty z `formaty` per kolumna (odwołując się do modułu 10),
   - dodać obramowania i style nagłówków (moduł 09),
   - ukryć kolumny pomocnicze,
   - zapisać **wszystko na jednej stronie A4 poziomo**,
   - ustawić `freeze_panes` na tyle, żeby nagłówek tabeli głównej został, ale **nie** uciekał tytuł,
   - ustawić stopkę z numeracją stron i **wyśrodkowanie** na stronie.

2. **`audyt_druku`** musi **realnie wykrywać** każdą z siedmiu wymienionych anomalii. To znaczy:
   - dla pierwszej: porównaj `page_setup.fitToWidth`/`fitToHeight` z `pageSetUpPr.fitToPage` — jeśli któreś z fitTo* jest niezerowe (albo nie-None), a `fitToPage` nie jest `True` → **anomalia**;
   - dla drugiej: porównaj `page_margins.header` z `page_margins.top`;
   - dla trzeciej: porównaj `max_row` z wierszem, na którym kończy się `print_area`;
   - dla czwartej: sprawdź, czy `print_title_rows` pasuje do wzorca `^\d+:\d+$` (użyj `re`);
   - dla piątej: sprawdź, czy `freeze_panes` **nie** jest `"A1"` i nie jest `None`, gdy arkusz ma > 20 wierszy;
   - dla szóstej: znajdź najdalszą komórkę z formatowaniem (nasłuchując `_cells` albo iterując do `max_row`/`max_column`) i sprawdź, czy jest w `print_area`;
   - dla siódmej: **oszacuj liczbę stron** — jeśli `fitToPage` i `fitToHeight = 1`, to jedna; jeśli `fitToHeight = 0`, oszacuj z `max_row` podzielonego przez przybliżoną liczbę wierszy na stronę; opisz ten szacunek jako **przybliżony i oznaczony**.

3. **`raport_audytu_tekstowy`** tworzy czytelny raport:

```
=== AUDYT DRUKU: 11_cw4.xlsx ===
Arkusz: Dashboard

USTAWIENIA:
  orientacja:              landscape (poziomo)
  papier:                  9 (A4)
  fitToWidth/fitToHeight:  1/1
  fitToPage (WŁĄCZNIK):    True
  margines header/top:     0.2 / 0.5
  print_title_rows:        '1:1'
  freeze_panes:            'B5'
  print_area:              'A1:F10'

ANOMALIE (0):
  (brak)

SZACOWANA LICZBA STRON: 1
```

4. **Sześć testów:**
   - `zbuduj_dashboard` tworzy plik z poprawną liczbą arkuszy i sekcji,
   - tytuł jest scalony i wartość jest **tylko** w lewej-górnej komórce,
   - kolumny pomocnicze są ukryte,
   - `audyt_druku` na poprawnym pliku zwraca **pustą** listę anomalii,
   - `audyt_druku` na pliku z `fitToWidth = 1` ale **bez** `fitToPage` wykrywa tę konkretną anomalię,
   - `audyt_druku` na pliku z `print_title_rows = "1"` wykrywa błędny format.

   Wypisz tabelę `test N: OK / BŁĄD` i `wynik: X/6`. Zapisz raport do `output/11_cw4_raport.txt`.

W pliku `output/11_cw4_wnioski.md`:

1. **Jak oszacowałeś liczbę stron?** Opisz metodę i jej **ograniczenia**. Dlaczego openpyxl **nie może** policzyć stron dokładnie? (Podpowiedź: openpyxl nie renderuje — nie wie, jakiej szerokości wyjdzie kolumna po zmianie czcionki.)
2. **Dlaczego anomalia „`fitToWidth` bez `fitToPage`" jest wykrywalna programowo** — i dlaczego to pokazuje, że warte jest pisanie takich audytów? Kiedy w praktyce byś je uruchamiał?
3. **Test szósty wymaga znalezienia „najdalszej sformatowanej komórki".** Jak to zrobiłeś? Dlaczego `ws.max_row` i `ws.max_column` **nie wystarczą** i co to mówi o modelu wymiarów z modułu 07?
4. **Gdybyś chciał rozszerzyć audyt o wykrywanie, czy nagłówek wydruku nie nachodzi na dane** — jak byś to zrobił **bez** otwierania Excela? Jakie informacje masz, a jakich brakuje? To pytanie o **granicę możliwości openpyxl** — nazwij ją wprost.
5. **Twoje `zbuduj_dashboard` używa `formaty: list[str | None]` — czyli kodów formatów jako napisów w konfiguracji.** W module 10 zalecaliśmy **rejestr formatów**. Jak pogodzić te dwa podejścia? Zaproponuj rozwiązanie (podpowiedź: konfiguracja trzyma **nazwy** z rejestru, nie kody) i uzasadnij, dlaczego to lepsze.

<details>
<summary><strong>Szkic rozwiązania zadania 4 — kluczowe fragmenty i uzasadnienia decyzji</strong></summary>

```python
"""Audyt gotowosci do druku - kluczowe fragmenty."""

from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


WZORZEC_TITLES = re.compile(r"^\d+:\d+$")


def _najdalsza_komorka(ws) -> tuple[int, int]:
    """Znajduje najdalszy wiersz i kolumne z REALNA zawartoscia lub formatem.

    Uwaga: ws.max_row/max_column sa wyliczane przez openpyxl z wymiarow
    arkusza i bywaja zawyzone przez formatowanie "na zapas". Tu idziemy
    po rzeczywistych komorkach w _cells (implementacja wewnetrzna!).
    """
    max_wiersz = 0
    max_kolumna = 0
    # _cells to wewnetrzna struktura - uzywamy swiadomie i ostroznie.
    # W razie zmian w bibliotece: fallback na max_row/max_column.
    for (wiersz, kolumna), komorka in getattr(ws, "_cells", {}).items():
        if komorka.value is not None or komorka.has_style:
            max_wiersz = max(max_wiersz, wiersz)
            max_kolumna = max(max_kolumna, kolumna)
    return max_wiersz, max_kolumna


def audyt_druku(sciezka: Path) -> dict:
    """Sprawdza gotowosc pliku do druku."""
    wb = load_workbook(sciezka)
    raporty = []

    try:
        for ws in wb.worksheets:
            anomalie: list[str] = []
            ustawienia: dict[str, object] = {}

            # --- zbierz ustawienia ---
            ustawienia["orientacja"] = ws.page_setup.orientation
            ustawienia["papier"] = ws.page_setup.paperSize
            ustawienia["fitToWidth"] = ws.page_setup.fitToWidth
            ustawienia["fitToHeight"] = ws.page_setup.fitToHeight
            page_setup_pr = ws.sheet_properties.pageSetUpPr
            fit_to_page = (
                page_setup_pr.fitToPage if page_setup_pr is not None else None
            )
            ustawienia["fitToPage"] = fit_to_page
            ustawienia["margines_header"] = ws.page_margins.header
            ustawienia["margines_top"] = ws.page_margins.top
            ustawienia["print_title_rows"] = ws.print_title_rows
            ustawienia["freeze_panes"] = ws.freeze_panes
            ustawienia["print_area"] = ws.print_area

            # --- ANOMALIA 1: fitTo* bez fitToPage ---------------------
            zada_fit = any(
                wartosc not in (None, 0)
                for wartosc in (ws.page_setup.fitToWidth, ws.page_setup.fitToHeight)
            )
            if zada_fit and fit_to_page is not True:
                anomalie.append(
                    "fitToWidth/fitToHeight ustawione, ale fitToPage nie jest True"
                    " - Excel ZIGNORUJE dopasowanie (tabela rozjedzie sie na strony)"
                )

            # --- ANOMALIA 2: header >= top -----------------------------
            if (
                ws.page_margins.header is not None
                and ws.page_margins.top is not None
                and ws.page_margins.header >= ws.page_margins.top
            ):
                anomalie.append(
                    f"margines header ({ws.page_margins.header}) >= top "
                    f"({ws.page_margins.top}) - naglowek moze wejsc na dane"
                )

            # --- ANOMALIA 3: print_area wiekszy niz dane ---------------
            najdalszy_wiersz, najdalsza_kolumna = _najdalsza_komorka(ws)
            if ws.print_area:
                obszary = (
                    ws.print_area if isinstance(ws.print_area, list) else [ws.print_area]
                )
                for obszar in obszary:
                    try:
                        _, min_row, _, max_row = range_boundaries(obszar)
                    except (ValueError, TypeError):
                        anomalie.append(f"print_area {obszar!r} nie jest poprawnym zakresem")
                        continue
                    if max_row > najdalszy_wiersz > 0:
                        anomalie.append(
                            f"print_area siega wiersza {max_row}, a dane koncza sie"
                            f" na {najdalszy_wiersz} - do PDF trafia"
                            f" {max_row - najdalszy_wiersz} PUSTYCH wierszy"
                        )

            # --- ANOMALIA 4: zly format print_title_rows ----------------
            if ws.print_title_rows and not WZORZEC_TITLES.match(str(ws.print_title_rows)):
                anomalie.append(
                    f"print_title_rows={ws.print_title_rows!r} - zly format."
                    f" Powinno byc '1:1' (zakres), nie '1'"
                )

            # --- ANOMALIA 5: freeze_panes 'A1' lub brak -----------------
            if najdalszy_wiersz > 20:
                if ws.freeze_panes is None:
                    anomalie.append(
                        f"arkusz ma {najdalszy_wiersz} wierszy i BRAK zamrozenia"
                        f" - naglowek bedzie uciekal przy przewijaniu"
                    )
                elif str(ws.freeze_panes).upper() == "A1":
                    anomalie.append(
                        "freeze_panes='A1' NIE zamraza niczego - to sposob na"
                        " usuniecie zamrozenia, nie na zamrozenie naglowka"
                    )

            # --- ANOMALIA 6: komorki sformatowane poza obszarem druku ---
            if ws.print_area:
                obszary = (
                    ws.print_area if isinstance(ws.print_area, list) else [ws.print_area]
                )
                _, min_row, min_col, max_row, max_col = None, 10**9, 10**9, 0, 0
                for obszar in obszary:
                    a, b, c, d = range_boundaries(obszar)
                    min_col, min_row = min(min_col, a), min(min_row, b)
                    max_col, max_row = max(max_col, c), max(max_row, d)
                if najdalsza_kolumna > max_col or najdalszy_wiersz > max_row:
                    anomalie.append(
                        f"dane lub formatowanie siega poza obszar wydruku"
                        f" (dane: {najdalszy_wiersz}x{najdalsza_kolumna},"
                        f" obszar: {max_row}x{max_col}) - czesc arkusza"
                        f" NIE zostanie wydrukowana"
                    )

            # --- INFORMACJA 7: szacunek liczby stron -------------------
            if fit_to_page and ws.page_setup.fitToHeight == 1:
                strony = 1
                sposob = "fitToHeight=1 - jedna strona (dokladnie)"
            elif fit_to_page and ws.page_setup.fitToHeight == 0:
                # Grube przyblizenie: ~45 wierszy na strone A4 poziomo
                strony = max(1, -(-najdalszy_wiersz // 45))
                sposob = f"szacunek ({najdalszy_wiersz} wierszy / ~45 na strone)"
            else:
                strony = None
                sposob = "nie da sie oszacowac bez dopasowania"

            raporty.append(
                {
                    "arkusz": ws.title,
                    "ustawienia": ustawienia,
                    "anomalie": anomalie,
                    "szacowana_liczba_stron": strony,
                    "sposob_szacunku": sposob,
                }
            )
    finally:
        wb.close()

    return {"plik": str(sciezka), "arkusze": raporty}
```

**Kluczowe decyzje projektowe i uzasadnienia:**

- **Anomalia „`fitTo*` bez `fitToPage`" jest wykrywalna w 100% programowo i to jest najcenniejszy element tego audytu.** Powód: to jest błąd, który **nie daje żadnego objawu w pliku**. Plik się zapisuje, otwiera, wygląda poprawnie. Błąd objawia się **dopiero przy druku albo przy konwersji do PDF** — i to zwykle u klienta, nie u Ciebie. Napisanie trzech linii sprawdzających oszczędza całą klasę „nie wiem, czemu u mnie działa". To jest dokładnie ta filozofia, którą w module 18 zastosujemy do audytu utraconych funkcji pliku.

- **`_najdalsza_komorka` używa `ws._cells` — czyli struktury wewnętrznej.** To **świadome** naruszenie zasady „nie dotykaj podkreślonych atrybutów". Uzasadnienie: `ws.max_row` jest wyliczany z wymiarów arkusza, które bywają zawyżone (moduł 07), więc **nie da się** odróżnić „sformatowana pusta komórka" od „komórka z danymi" inaczej. Mitygacja: funkcja ma fallback na `max_row`/`max_column` i jest **jednym, izolowanym miejscem** — jeśli openpyxl zmieni strukturę wewnętrzną, poprawiasz jedną funkcję. To wzorzec: **naruszenie enkapsulacji jest akceptowalne, jeśli jest jedno, udokumentowane i otoczone fallbackiem**. Zapowiedź modułów 24 (Adapter — izolowanie wiedzy o obcym API) i 26.

- **Szacowanie liczby stron to jawna aproksymacja i jest tak oznaczona.** Openpyxl **nie renderuje arkuszy** — nie zna szerokości kolumn w pikselach (bo zależą od czcionki), ani wysokości wierszy po zawinięciu tekstu. Więc **nie da się policzyć stron dokładnie** bez uruchomienia silnika Excela albo LibreOffice. Oszacowanie 45 wierszy na stronę A4 poziomo jest grube i **nazwane grubym**. To kluczowa lekcja tego modułu: **wiedzieć, gdzie openpyxl mówi prawdę, a gdzie tylko przybliżenie** — i zawsze nazywać przybliżenie przybliżeniem. Laboratorium tego rozróżnienia to moduł 18.

- **Dlaczego `_najdalsza_komorka` sprawdza `value is not None OR has_style`.** Bo komórka sformatowana, ale pusta, **też zajmuje miejsce w pliku** i **też wydłuża wydruk**. To dokładnie ta „formatowanie na zapas" z pułapki `max_row`. Sprawdzanie obu warunków łapie oba przypadki: prawdziwe dane i sformatowane pustki.

- **Szósty test (błędny `print_title_rows`) używa regexa, nie `int()`.** Bo `print_title_rows` to napis, który może być `"1:1"`, `"1"`, `"A:B"`, `1` (liczba), `None`. Regex `^\d+:\d+$` jest **jednoznaczny** i natychmiast odrzuca wszystko, co nie jest zakresem wierszy. Próba `int(...)` rzuciłaby wyjątek na `"1:1"` — czyli **dokładnie na poprawnym** formacie. To pułapka walidacji: **narzędzie walidacji musi być bardziej tolerancyjne na poprawne dane niż Twoja intuicja**.

**Czego szukać w wynikach:** uruchom `audyt_druku` na plikach z **Przykładu 1** i **Przykładu 4**. Oba powinny zwrócić **pustą listę anomalii** — to jest test „czy mój audyt nie daje fałszywych alarmów". Jeśli któryś zgłasza anomalię, sprawdź, czy to prawdziwy problem, czy nadgorliwość w Twojej implementacji. Następnie **celowo zepsuj** jeden plik (np. usuń `fitToPage`, wpisz `print_title_rows = "1"`) i sprawdź, czy audyt to złapie. **Dopiero to jest prawdziwy test — audyt, który nie potrafi wykryć znanego błędu, jest bezwartościowy.**

**Co dopisać w wersji produkcyjnej:** `audyt_druku` na **wczytanym pliku** (nie ścieżce), żeby można było sprawdzać skoroszyt **przed** zapisem — wtedy wykrywasz problem, zanim plik trafi do klienta. To jest jeden z tych przypadków, gdzie audyt lepiej mieć wbudowany w pipeline niż uruchamiany ręcznie (moduł 20 — testy).

</details>

## 6. Typowe błędy i pułapki

**1. „Ustawiłem `fitToWidth = 1`, a tabela nadal rozjeżdża się na cztery strony w poziomie" (objaw) → brak **włącznika trybu dopasowania**. `fitToWidth` i `fitToHeight` to tylko **wartości liczbowe** — Excel zapisze je do pliku, ale zignoruje, dopóki nie ma włączonego `fitToPage`. To dokładnie jak przekręcenie pokrętła skali na kserokopiarki bez wybrania trybu „zmniejsz, żeby zmieścić" (analogia 1.5) (przyczyna) → **dodaj trzecią linię**: `from openpyxl.worksheet.properties import PageSetupProperties; ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)`. Weryfikacja: wczytaj plik i sprawdź `ws.sheet_properties.pageSetUpPr.fitToPage is True`. To najczęstsza i najbardziej „niewytłumaczalna" pułapka tego modułu — bo **błąd nie daje żadnego objawu w pliku**, tylko przy druku. Dlatego warto ją wykrywać audytem (zadanie 4) (naprawa).**

**2. „Wpisałem `print_title_rows = "1"` i nagłówek nie powtarza się na kolejnych stronach" (objaw) → **zły format wartości**. `print_title_rows` wymaga **zakresu w formacie `"start:koniec"`**, dokładnie tak, jak zapisujesz zakres w Excelu. Wartość `"1"` albo liczba `1` to nie zakres (przyczyna) → używaj `"1:1"` dla jednego wiersza, `"1:2"` dla dwóch. Analogicznie `print_title_cols = "A:B"`, nie `"A"`. **Uwaga na dodatkową subtelność:** te właściwości są zapisywane w pliku jako **definiowane nazwy** (`_xlnm.Print_Titles`), a nie jako atrybuty arkusza — więc jeśli wczytasz plik i sprawdzisz `ws.print_title_rows`, openpyxl **odtworzy** wartość z definiowanej nazwy. Działa, ale pokazuje, że to mechanizm oparty na czymś innym niż zwykłe ustawienia (naprawa).**

**3. „Scaliłem nagłówek i zniknęły mi dane z pozostałych komórek" (objaw) → **scalanie jest operacją destrukcyjną**. `merge_cells` usuwa wartości ze wszystkich komórek poza lewą-górną i zamienia je na obiekty `MergedCell` (tylko do odczytu). Nie ma kosza, nie ma `undo`, `unmerge_cells` **nie przywraca** wartości (analogia 1.3) (przyczyna) → **buduj cały tekst w Pythonie i wpisuj go do lewej-górnej komórki**, scalaj **na końcu**, po wpisaniu wszystkich danych. Nigdy nie scalaj komórek, które zawierają dane do obliczeń. Jeśli wczytujesz **cudzy** plik ze scaleniami, pamiętaj, że odczyt ze scalonej komórki zwraca `None` — napisz funkcję `wartosc_efektywna`, która cofa się do lewej-górnej. Sprawdź też, czy scalenie jest w ogóle potrzebne: często `wrap_text` i szeroka kolumna dają lepszy efekt bez utraty danych (naprawa).**

**4. „Obramowanie tabeli wygląda dobrze, ale na wydruku brakuje krawędzi / tabela jest obcięta" (objaw) → dwie możliwe przyczyny, często występujące razem. Po pierwsze: **polegasz na liniach siatki**, a one domyślnie **nie drukują się** — więc na ekranie widzisz tabelę, a na papierze dostajesz gołe liczby. Po drugie: **obramowanie nadałeś tylko komórkom z danymi**, a Excel przy druku rysuje obramowanie komórki na jej krawędziach — jeśli sąsiednia komórka (poza zakresem danych) nie ma obramowania, prawa krawędź ostatniej kolumny może się nie wydrukować albo wyjść niepełna. Trzecia możliwość: `print_area` jest **węższy** niż tabela i obcina skrajne kolumny (przyczyna) → (a) wyłącz siatkę i nadaj **własne obramowanie** wszystkim komórkom tabeli, (b) po nadaniu obramowania sprawdź, czy prawa i dolna krawędź tabeli mają komplet boków — dla obwódki potrzebujesz **czterech** boków na skrajnych komórkach, (c) upewnij się, że `print_area` obejmuje **całą** tabelę, włącznie z ostatnią kolumną i ostatnim wierszem. Sprawdź to w podglądzie wydruku — nie na ekranie (naprawa).**

**5. „Ustawiłem `ws.freeze_panes = "A1"` i nagłówek ucieka przy przewijaniu" (objaw) → **`"A1"` nie zamraża niczego**. Reguła jest jedna: `freeze_panes` to adres **pierwszej komórki obszaru przewijalnego**, więc zamrożone jest wszystko **powyżej i na lewo**. `"A1"` leży na samym początku arkusza, nie ma nic powyżej ani na lewo — więc nie ma czego zamrozić. To jest zresztą **sposób na usunięcie zamrożenia** z szablonu (analogia 1.4) (przyczyna) → **żeby zamrozić wiersz 1, użyj `"A2"`** (cyfra o jeden większa niż ostatni zamrożony wiersz). Żeby zamrozić wiersz 1 i kolumnę A — `"B2"`. Żeby zamrozić dwa wiersze nagłówka — `"A3"`. Weryfikacja: sprawdź, czy adres nie jest `"A1"`; w kodzie produkcyjnym rozważ asercję, że `ws.freeze_panes.upper() != "A1"` (naprawa).**

**6. „Podgląd wydruku pokazuje trzy puste strony na końcu" (objaw) → **`print_area` obejmuje zakres większy niż dane**. Najczęstsza przyczyna: `print_area = "A1:H1000"` wpisane „na sztywno", gdy danych jest 200 wierszy. Druga, podstępniejsza przyczyna: obszar został **wyliczony** z `ws.max_row`, ale `max_row` jest zawyżony przez formatowanie „na zapas" (moduł 07) — sformatowałeś kolumnę do wiersza 1000, więc `max_row` zwróciło 1000 (przyczyna) → (a) wyliczaj `print_area` z **realnego** zakresu danych, nie z `max_row`, jeśli masz wątpliwości co do formatowania, (b) **nie formatuj komórek, których nie zamierzasz wypełnić** — to nie tylko problem druku, ale i wydajności (moduł 19), (c) napisz własną funkcję skanującą ostatni wiersz z danymi (moduł 07 ma szkic). Sprawdzenie: policz wiersze w `print_area` i porównaj z liczbą wierszy, które faktycznie mają wartości (naprawa).**

**7. „Nagłówek wydruku nachodzi na pierwszy wiersz tabeli" (objaw) → **`page_margins.header >= page_margins.top`**. Margines `header` to odległość nagłówka od **górnej krawędzi strony**, a `top` to odległość **treści** od górnej krawędzi. Jeśli nagłówek jest odsunięty głębiej niż początek treści — nachodzą na siebie (przyczyna) → **zachowaj regułę `page_margins.header < page_margins.top`**. Sensowne wartości: `header = 0.2`–`0.3`, `top = 0.5`–`0.8`. To dokładnie ta klasa błędu, którą łatwo wykryć programowo — i jedno z siedmiu sprawdzeń w audycie z zadania 4 (naprawa).**

**8. „Chciałem wpisać do stopki `Sprzedaż & Marketing` i wyświetla się dziwnie / znak ginie" (objaw) → **`&` jest znakiem specjalnym w języku nagłówków i stopek**. Rozpoczyna **polecenie** (`&P` = numer strony, `&N` = liczba stron, `&D` = data). Samotny ampersand zostanie potraktowany jako początek polecenia — a że po nim następuje spacja, Excel albo zgubi znak, albo wyświetli coś nieoczekiwanego (analogia 1.6) (przyczyna) → **escape'uj ampersand przez podwojenie**: `&&`. Poprawnie: `"Sprzedaż && Marketing"`. To ta sama rodzina mechanizmów co `\\` w kodach formatów (moduł 10) i `""` w CSV (naprawa).**

**9. „Nagłówek i stopka nie wyświetlają się — nie widzę ich w arkuszu" (objaw) → **nagłówki i stopki są elementem DRUKU, nie ekranu**. W normalnym widoku arkusza Excel ich nie pokazuje. Nie ma tu błędu — tak działa Excel. Nagłówki i stopki widzisz wyłącznie w **podglądzie wydruku** (`Ctrl+P`) albo w widoku „Page Layout" (mapa strony) (przyczyna) → **weryfikuj nagłówki i stopki w podglądzie wydruku albo eksportując do PDF**. To także **wskazówka procesowa**: jeśli w Twoim pipeline generujesz PDF, sprawdź ten PDF — bo tam nagłówek będzie widoczny, a w pliku `.xlsx` nie. Uwaga dodatkowa: `evenHeader`/`evenFooter` zadziałają tylko po włączeniu w Excelu drukowania różnych stron parzystych i nieparzystych — **sprawdź w dokumentacji swojej wersji**, jak to ustawić programowo (naprawa).**

**10. „Ustawiłem `scale = 85`, ale tabela rozjeżdża się na strony / skala jest ignorowana" (objaw) → **`scale` i `fitTo*` wykluczają się**. Jeśli w pliku (albo w szablonie, albo w ustawieniach odziedziczonych) jest `fitToPage = True`, Excel **całkowicie ignoruje** `scale`. I odwrotnie: jeśli używasz `fitToWidth`, to `scale` nie ma znaczenia (przyczyna) → **wybierz jedno podejście.** Jeśli chcesz stałą skalę — ustaw `fitToPage = False` (jawnie!) i `page_setup.scale = 85`. Jeśli chcesz dopasowanie — ustaw `fitToWidth`/`fitToHeight` **i** `fitToPage = True`, i nie dotykaj `scale`. **Szczególna uwaga przy wczytywaniu cudzych plików:** odziedziczony `fitToPage = True` może sprawić, że Twoje ustawienia skali nie zadziałają, mimo że ich nie dotykałeś — to powód, dla którego warto ustawiać `fitToPage` **jawnie**, także na `False` (naprawa, zapowiedź modułu 18).**

**11. „`ws.page_breaks` nie istnieje / rzuca `AttributeError`" (objaw) → **takie API w openpyxl nie istnieje.** Prawdopodobnie skopiowałeś kod z poradnika albo z odpowiedzi, której autor nie sprawdził. Obecność atrybutu o podobnej nazwie jest naturalna przy zgadywaniu — ale openpyxl używa **dwóch oddzielnych obiektów** dla podziałów (przyczyna) → używaj `ws.row_breaks.append(Break(id=n))` i `ws.col_breaks.append(Break(id=n))`, importując `from openpyxl.worksheet.pagebreak import Break`. **Zapamiętaj semantykę: `Break(id=n)` wstawia podział PO wierszu/kolumnie `n`** — więc nowa strona **zaczyna się** od `n+1`. Pomyłka o jeden daje efekt „każda sekcja zaczyna się z osieroconym wierszem". Pamiętaj też o limicie ~1026 ręcznych podziałów na arkusz — powyżej kolejne są **cicho ignorowane** (naprawa).**

**12. „Dodałem autofiltr, ale nie działa / pokazuje dziwne wyniki" (objaw) → dwie typowe przyczyny. Pierwsza: **`add_filter_column` przyjmuje indeks kolumny liczony od zera, względem lewej kolumny zakresu `ref`** — czyli dla `ref = "A1:H100"` indeks `0` to A, `3` to D. To **inna konwencja niż w całym pozostałym API openpyxl**, gdzie kolumny liczy się od 1. Druga: **wartości filtra muszą dokładnie odpowiadać danym, łącznie z typem** — jeśli w komórkach są liczby, a podajesz napisy, filtr może nie zadziałać (przyczyna) → (a) przelicz indeks z konwencji „od 1" na „od 0" w jednym, izolowanym miejscu — i napisz to w komentarzu, (b) sprawdź typy wartości w kolumnie przed użyciem w filtrze (`isinstance(cell.value, (int, float))`), (c) pamiętaj, że **autofiltr to tylko nakładka na zakres** — nie rozszerza się przy dopisywaniu danych. Jeśli potrzebujesz automatycznego rozszerzania, użyj **tabeli** (moduł 12) (naprawa).**

**13. „Szerokość kolumny ustawiona na 20 wygląda zupełnie inaczej na innym komputerze / obrazek się nie wyrównuje" (objaw) → **`width` jest wyrażone w „znakach", nie w pikselach** i zależy od domyślnej czcionki arkusza, DPI ekranu i rozdzielczości. Nie ma sposobu, żeby precyzyjnie przeliczyć `width` na piksele (przyczyna) → **traktuj `width` jako jednostkę względną.** Jeśli potrzebujesz wyrównania obrazka do komórki (moduł 16), rób to empirycznie: ustaw, zapisz, otwórz w Excelu, zmierz, popraw. Sensowne wartości startowe: etykiety 25–35, kwoty 12–16, procenty 8–10, daty ISO 12–14. Jeśli chcesz zmienić szerokość domyślną całego arkusza, użyj `ws.sheet_format.defaultColWidth` — ale **sprawdź w dokumentacji swojej wersji**, bo zachowanie przy mieszaniu z `column_dimensions` bywa nieoczywiste (naprawa).**

**14. „Wysokość wiersza ustawiona jawnie obcina mi zawinięty tekst" (objaw) → **openpyxl nie ma modelu renderowania tekstu.** Nie wie, ile miejsca zajmie tekst po zawinięciu — to zależy od czcionki i szerokości kolumny. Jeśli **nie** ustawisz wysokości, Excel przy otwarciu **sam** dopasuje wysokość do zawiniętego tekstu. Jeśli **ustawisz** ją jawnie, Excel uzna ją za wiążącą i przy zbyt małej wartości **obetnie** tekst (przyczyna) → przy kolumnach z `wrap_text` (moduł 09) **nie ustawiaj wysokości wiersza** — pozwól Excelowi dopasować. Wyjątek: wiersz nagłówka, gdzie wiesz z góry, ile linii tekstu będzie i chcesz jednolity wygląd. Sprawdzenie: otwórz plik, zaznacz wiersz i wybierz „Autodopasuj wysokość wiersza" — jeśli wysokość się zmieni, Twoja wartość była zbyt mała (naprawa).**

**15. „Dodałem ustawienia druku do arkusza, ale przy wczytaniu pliku przez inny program nic nie działa" (objaw) → **większość tych ustawień jest specyficzna dla Excela i bywa ignorowana przez inne programy.** Dotyczy to zwłaszcza: `fitToPage` (LibreOffice zwykle obsługuje, ale bywają różnice), nietypowych kodów nagłówków i stopek, oraz grupowania/outline przy konwersji. Kod `&[Page]` (używany w dokumentacji openpyxl) i `&P` (klasyczny) są obsługiwane różnie w różnych programach (przyczyna) → (a) **jeśli Twoim celem jest PDF, generuj PDF w środowisku docelowym** (Excel albo LibreOffice) i **sprawdź wynik**, (b) używaj **najprostszych, klasycznych kodów** (`&P`, `&N`, `&D`), (c) testuj u odbiorcy — nie u siebie. To jest ta sama zasada co w module 10 z LCID: **plik `.xlsx` opisuje intencję, a wynik zależy od tego, kto go renderuje.** To zapowiedź modułu 18 i 26: rozróżnienie między „co zapisałem" a „co zobaczy odbiorca" (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Arkusz ma dwie warstwy: dane i układ — i są całkowicie rozłączne.** Do tej pory zmieniałeś zawartość komórek. Teraz zmieniasz **metadane arkusza**: szerokości kolumn, zamrożenia, ustawienia druku, scalenia, filtry, grupowanie. Te warstwy nie mieszają się: możesz przepisać cały układ, nie dotykając ani jednej wartości, i odwrotnie. **Praktyczna konsekwencja: cały układ można zebrać w jednej funkcji `zastosuj_layout(ws)`** i wywołać na każdym arkuszu każdego raportu — i to jest prototyp warstwy prezentacji z modułu 26.

2. **Linie siatki to rusztowanie, obramowanie to ściana — i nigdy ich nie mieszaj.** Szara siatka pomaga na ekranie, nie jest częścią dokumentu i domyślnie **nie drukuje się**. Raport, który „wygląda jak tabela" tylko dzięki siatce, nie ma żadnych ramek. **Reguła: `showGridLines = False` i własne `Border` na wszystkich komórkach tabeli.** Bez tego dwie linie kodu, dokument wygląda jak arkusz roboczy, a nie raport.

3. **Scalanie komórek to operacja destrukcyjna z cichą asymetrią.** Lewa-górna komórka zachowuje wartość, pozostałe tracą ją **bezpowrotnie** i stają się obiektami `MergedCell` (tylko do odczytu). **Zapis** do scalonej komórki rzuca `AttributeError` (głośno — dobrze), ale **odczyt** zwraca `None` (cicho — niebezpiecznie). `unmerge_cells` **nie przywraca** wartości. **Reguła: buduj cały tekst w Pythonie, wpisuj do lewej-górnej, scalaj na końcu.** I naucz się, że `None` przy odczycie nie znaczy „puste" — może znaczyć „scalone".

4. **Druk to jedyne miejsce, gdzie openpyxl daje niemal pełną kontrolę — ale wymaga trzech ustawień razem.** `fitToWidth` i `fitToHeight` to tylko **wartości**; włącznikiem trybu jest `pageSetUpPr = PageSetupProperties(fitToPage=True)`. Bez niego Excel cicho ignoruje dopasowanie, a błąd objawia się **dopiero przy druku albo w PDF** — czyli u odbiorcy, nie u Ciebie. Do tego dochodzą: `print_area` (za duży = puste strony), `print_title_rows` (format `"1:1"`, nie `"1"` — jedna linia, która ratuje czytelność wielostronicowego raportu) i marginesy (`header < top`). **Reguła: przygotuj ustawienia, wyeksportuj PDF i sprawdź go** — bo tylko PDF pokazuje prawdę o wydruku.

5. **Ale granice tej kontroli trzeba znać i nazywać.** Openpyxl **nie renderuje** — nie wie, ile pikseli ma kolumna `width = 20` (bo zależy od czcionki), nie policzy dokładnie stron, nie dopasuje wysokości wiersza do zawijanego tekstu, a jego nagłówki i stopki mogą zachowywać się różnie w Excelu, LibreOffice i podglądach. To ta sama kategoria co „openpyxl potrafi częściowo" z modułu 03. **Konsekwencja praktyczna: tam, gdzie openpyxl daje tylko przybliżenie, nazywaj je przybliżeniem w kodzie i w dokumentacji, i zawsze weryfikuj efekt w programie docelowym.** Umiejętność odróżnienia „zapisane" od „zobaczone przez odbiorcę" jest podstawą modułów 18 i 26.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. WYMIARY
# ==================================================================
ws.column_dimensions["A"].width = 28        # szerokosc w ZNAKACH (nie px!)
ws.column_dimensions["F"].hidden = True     # ukrycie kolumny
ws.row_dimensions[1].height = 30            # wysokosc w PUNKTACH (1pt=1/72")
ws.row_dimensions[12].hidden = True

# Grupowanie (widoczny przycisk +/-)
ws.column_dimensions.group("B", "D", outline_level=1, hidden=True)
ws.row_dimensions.group(1, 10, hidden=True)
ws.sheet_properties.outlinePr.summaryBelow = True
ws.sheet_properties.outlinePr.summaryRight = True

# ⚠️ hidden = ciche ukrycie; group = widoczny przycisk. Wybieraj swiadomie.
# ⚠️ Ustawiona JAWNIE wysokosc moze obciac wrap_text. Nie ustawiaj przy wrap.

# ==================================================================
# 2. WIDOK EKRANU
# ==================================================================
ws.sheet_view.showGridLines = False      # rusztowanie wylaczone (modul 11!)
ws.sheet_view.zoomScale = 100            # zoom 10-400
ws.sheet_view.showFormulas = True        # pokaz formuly (diagnostyka!)
ws.sheet_view.showZeros = False          # zera jak puste (lepiej format ;-;)
ws.sheet_view.tabSelected = True
wb.active = 1                            # indeks od 0 -> ktory arkusz otwarty

# ==================================================================
# 3. ZAMROZENIE - REGULA: "powyzej i na lewo"
# ==================================================================
ws.freeze_panes = "A2"     # wiersz 1
ws.freeze_panes = "B1"     # kolumna A
ws.freeze_panes = "B2"     # wiersz 1 + kolumna A
ws.freeze_panes = "A3"     # wiersze 1-2
ws.freeze_panes = "C3"     # wiersze 1-2 + kolumny A-B
ws.freeze_panes = "A1"     # ⚠️ NIC! (sposob na USUNIECIE zamrozenia)
ws.freeze_panes = None     # tez nic

# ==================================================================
# 4. SCALANIE - UWAGA, NISZCZY DANE!
# ==================================================================
ws.merge_cells("A1:D1")
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
ws.unmerge_cells("A1:D1")            # ⚠️ NIE przywraca wartosci!
print(ws.merged_cells.ranges)        # lista scalonych zakresow

# ⚠️ SCALANIE: lewa-gorna zachowuje wartosc, reszta -> MergedCell(value=None)
# ⚠️ Zapis do scalonej komorki (nie lewej-gornej) -> AttributeError
# ⚠️ Odczyt ze scalonej komorki -> None (CICHO!)

# BEZPIECZNY WZORZEC:
from openpyxl.utils.cell import range_boundaries

def scal_z_tekstem(ws, zakres: str, tekst: str) -> None:
    min_col, min_row, _, _ = range_boundaries(zakres)
    ws.merge_cells(zakres)                          # najpierw scal...
    ws.cell(row=min_row, column=min_col).value = tekst   # ...potem wpisz

# ==================================================================
# 5. AUTOFILTR (rozni sie od TABELI - modul 12!)
# ==================================================================
ws.auto_filter.ref = "A1:H100"
ws.auto_filter.add_filter_column(3, ["Alfa", "Beta"], blank=False)
#                                  ^ indeks OD ZERA! (reszta API od 1)
ws.auto_filter.add_sort_condition("D2:D100", descending=True)

# ==================================================================
# 6. DRUK - STRONA
# ==================================================================
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE    # poziomo
ws.page_setup.paperSize = ws.PAPERSIZE_A4               # A4 (=9)

# Marginesy w CALACH. header MUSI byc < top!
ws.page_margins.left = 0.4
ws.page_margins.right = 0.4
ws.page_margins.top = 0.7
ws.page_margins.bottom = 0.7
ws.page_margins.header = 0.3
ws.page_margins.footer = 0.3

# ==================================================================
# 7. DRUK - DOPASOWANIE (TRZY LINIE RAZEM!)
# ==================================================================
from openpyxl.worksheet.properties import PageSetupProperties

ws.page_setup.fitToWidth = 1        # jedna strona szerokosci
ws.page_setup.fitToHeight = 0       # 0 = BEZ LIMITU (nie "zero stron")
ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
# ^^^^ TA LINIA JEST OBOWIAZKOWA - bez niej dwie powyzsze SA IGNOROWANE!

# JEDNA STRONA (dashboard):        fitToWidth=1, fitToHeight=1, fitToPage=True
# DLUGI RAPORT:                    fitToWidth=1, fitToHeight=0, fitToPage=True
# STALA SKALA (bez dopasowania):   fitToPage=False, scale=85
# ⚠️ scale i fitTo* WYKLUCZAJA SIE. Przy fitToPage=True Excel ignoruje scale.

# ==================================================================
# 8. DRUK - OBSZAR I POWTARZANIE NAGLOWKA
# ==================================================================
ws.print_area = "A1:H50"
ws.print_area = ["A1:H50", "J1:N50"]        # wiele zakresow

ws.print_title_rows = "1:1"     # ⚠️ FORMAT "start:koniec", nie "1"!
ws.print_title_cols = "A:B"     # ⚠️ "A:B", nie "A"

# Wyliczaj obszar, nie wpisuj na sztywno (puste strony!):
from openpyxl.utils import get_column_letter
ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

# ==================================================================
# 9. DRUK - NAGLOWEK, STOPKA, OPCJE
# ==================================================================
ws.oddHeader.center.text = "Raport sprzedaży 2026"
ws.oddHeader.center.size = 12
ws.oddHeader.center.font = "Calibri,Bold"
ws.oddHeader.center.color = "FF333333"

ws.oddFooter.left.text = "&D"                    # data
ws.oddFooter.center.text = "Strona &P z &N"      # strona / liczba stron
ws.oddFooter.right.text = "&A"                   # nazwa arkusza

# KODY: &P strona | &N liczba stron | &D data | &T godzina
#       &F plik | &A arkusz | &Z sciezka | && doslowny &
# ⚠️ "Sprzedaz && Marketing" - ampersand ESCAPE'owany przez podwojenie
# ⚠️ Naglowki widoczne TYLKO w podgladzie wydruku (Ctrl+P), nie w arkuszu
# ⚠️ evenHeader/firstHeader wymagaja wlaczonej opcji w Excelu

ws.print_options.horizontalCentered = True   # wysrodkuj w poziomie
ws.print_options.verticalCentered = True     # wysrodkuj w pionie
ws.print_options.gridLines = True            # ⚠️ drukuj siatke - RZADKO!

# ==================================================================
# 10. PODZIAL STRONY (⚠️ ws.page_breaks NIE ISTNIEJE!)
# ==================================================================
from openpyxl.worksheet.pagebreak import Break

ws.row_breaks.append(Break(id=20))   # nowa strona PO wierszu 20
ws.col_breaks.append(Break(id=5))    # nowa strona PO kolumnie 5
# ⚠️ Break(id=n) = podzial PO n; nowa strona ZACZYNA sie od n+1
# ⚠️ Limit ~1026 recznych podzialow na arkusz - dalsze CICHO ignorowane

# ==================================================================
# 11. GOTOWY LAYOUT RAPORTU - SKOPIUJ DO PROJEKTU
# ==================================================================
from openpyxl.worksheet.properties import PageSetupProperties

def zastosuj_layout_raportu(ws, tytul="", tytul_rows="1:1", szerokosci=None):
    # Widok
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.freeze_panes = "A2"
    # Wymiary
    if szerokosci:
        for litera, szer in szerokosci.items():
            ws.column_dimensions[litera].width = szer
    # Strona
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.7
    ws.page_margins.header = ws.page_margins.footer = 0.3
    # Dopasowanie (TRZY linie!)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    # Naglowek i stopka
    if tytul:
        ws.oddHeader.center.text = tytul
        ws.oddHeader.center.font = "Calibri,Bold"
    ws.oddFooter.left.text = "&D"
    ws.oddFooter.center.text = "Strona &P z &N"
    ws.oddFooter.right.text = "&A"
    # Powtarzanie naglowka
    ws.print_title_rows = tytul_rows

# ==================================================================
# 12. CO OTWORZYC W EXCELU, ZEBY SPRAWDZIC
# ==================================================================
# Ctrl+P (podglad wydruku)  -> naglowek/stopka, liczba stron, dopasowanie
# Ctrl+End                  -> gdzie Excel mysli, ze koncza sie dane
# Widok -> Podzial stron    -> gdzie Excel lamie strony
# Ctrl+Shift+8 (nawigacja)  -> pokazac ukryte kolumny? NIE: Zaznacz -> Widok
# Karta "Widok" -> Freeze    -> ktore okna sa zamrozone
```

## 9. Co dalej

Zamknąłeś drugą połowę części o formatowaniu. Masz teraz pełny obraz tego, co openpyxl potrafi zrobić z arkuszem od strony wizualnej: od pojedynczej komórki (moduły 09–10), przez wymiary i układ (ten moduł), aż po gotowy dokument do druku. Trzy rzeczy, które zabierasz ze sobą:

- **Nawyk wyłączania siatki i nadawania własnego obramowania.** To dwie linie kodu, które w każdym nowym raporcie zmieniają go z arkusza roboczego w dokument. Jeśli zapamiętasz z tego modułu tylko jedno — niech to będzie ta para: `showGridLines = False` + `Border` na wszystkich komórkach tabeli.
- **`print_title_rows = "1:1"` jako odruch przy każdym raporcie wielostronicowym.** To jedna linia kodu, która ratuje czytelność. I jej brak to najczęstszy brak w automatycznie generowanych plikach.
- **Świadomość asymetrii scalania.** Zapis rzuca wyjątek (dobrze), odczyt zwraca `None` (podstępnie). Jeśli kiedykolwiek będziesz czytać cudze pliki — a w module 18 będziesz — ta asymetria jest źródłem błędów, których nikt nie zauważy, bo plik „wygląda dobrze".

**W module 12** zajmiemy się **tabelami** — obiektami, które są czymś więcej niż zakresem z formatowaniem. Zobaczysz:

- **czym różni się `Table` od `ws.auto_filter`** (i dlaczego to nie to samo, choć oba potrafią filtrować),
- jak zbudować tabelę z nazwą, stylem i **structured references** (`=SUM(SprzedazTabela[Kwota])`),
- **dlaczego tabele mają tożsamość** — i co to znaczy dla formuł, które się do nich odwołują,
- jak tabela **rozszerza się sama**, gdy użytkownik dopisuje wiersz (czego autofiltr nie robi),
- **czym są `TableStyleInfo`** i jak wybrać styl (i dlaczego nazwy stylów wyglądają jak `TableStyleMedium9`),
- jakie są **realne ograniczenia** obsługi tabel w openpyxl: brak API „wstaw wiersz do tabeli", problemy z nakładającymi się zakresami, brak tabel w `read_only`/`write_only`, oraz historyczny błąd „Table filters are always overridden" (naprawiony w 3.1.0),
- **kiedy tabela przeszkadza** — bo nie każde dane powinny być tabelą.

**Zanim przejdziesz dalej, wykonaj jedno ćwiczenie obserwacyjne:**

1. **Uruchom `examples/11_gotowy_do_druku.py` i `examples/11_dashboard.py`**, a potem **wyeksportuj oba do PDF** — w Excelu (`Ctrl+P` → „Microsoft Print to PDF") albo w LibreOffice (`Plik` → `Eksportuj jako PDF`). Otwórz oba PDF-y i zadaj sobie trzy pytania:
   - **Czy nagłówek tabeli powtarza się na każdej stronie?** (Jeśli raport ma jedną stronę, dodaj 200 wierszy danych i sprawdź ponownie — to jedyny sposób, żeby zobaczyć działanie `print_title_rows`.)
   - **Czy jest widoczna różnica między „wygaszoną" stopką `&P z &N` a literałami?** Jeśli w PDF widzisz `&P`, to znaczy, że program, którym konwertowałeś, **nie zrozumiał kodu** — i to jest ta granica z punktu 5 podsumowania: **plik opisuje intencję, a renderer decyduje o wyniku**.
   - **Czy jakiś element jest obcięty?** Ostatnia kolumna, nagłówek nachodzący na dane, brakująca prawa krawędź obramowania. Wszystko, co zauważysz w PDF, a nie widzisz w `.xlsx` — to dokładnie ta klasa problemów, dla której napisałeś audyt w zadaniu 4.

2. **Wróć do swojego starego raportu** (tego samego, który analizowałeś po module 10) i sprawdź trzy rzeczy:
   - **Czy ma zamrożony nagłówek?** `ws.freeze_panes` — jeśli `None`, to najprawdopodobniej brakuje go tam, gdzie najbardziej by pomógł.
   - **Czy tabela ma własne obramowanie, czy polega na siatce?** Wyłącz siatkę w czytniku albo wyeksportuj do PDF — jeśli tabela się rozsypuje, masz odpowiedź.
   - **Czy `print_area` jest wyliczony, czy wpisany na sztywno?** Jeśli znalazłeś tam liczbę w rodzaju `1000`, to właśnie znalazłeś kandydata na puste strony w PDF — i konkretny powód, żeby zapisać listę kontrolną z tego modułu na stałe w swoim projekcie.