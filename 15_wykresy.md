Zanim napiszę moduł, zweryfikuję kluczowe sygnatury i ograniczenia API wykresów w openpyxl 3.1.x — to obszar, w którym brief zawiera kilka niepewnych pozycji (`height`/`width`, `dLblPos`, zachowanie wykresów w `write_only`).

Kluczowe potwierdzenia uzyskane. Teraz pobiorę strony dotyczące drugiej osi, wykresów kombinowanych, pie/bubble/radar/surface oraz graficznych właściwości — to obszary, na których brief opiera się najbardziej.

Teraz zweryfikuję sygnatury `Series`, osi, etykiet danych i właściwości graficznych, plus zachowanie wykresów w trybach wydajnościowych.

`Workbook(write_only=True)` z wykresami **działa** (oficjalny przykład bar.html) — to ważna korekta. Weryfikuję jeszcze tryb `read_only`, atrybuty `height`/`width` oraz sygnatury osi i etykiet.

Mam potwierdzone sygnatury osi i etykiet. Ostatnia partia weryfikacji: atrybuty bazowe wykresu (`height`/`width`/`legend`), `Reference`, układ legendy oraz zachowanie w trybie `read_only`.

Ostatnia weryfikacja: lista znanych błędów wykresów w changelogu (kluczowa dla sekcji „co openpyxl gubi/upraszcza") oraz sygnatura `Marker`.

Dwie ostatnie weryfikacje źródłowe: pełny moduł `axis.py` (czy istnieje alias `NumberAxis`) oraz klasa bazowa wykresu (`height`, `width`, `legend`, `dataLabels`).

Ostatnia weryfikacja: gdzie dokładnie żyją etykiety danych (poziom wykresu vs serii) i atrybuty wykresów słupkowych.

Zweryfikowałem API wykresów bezpośrednio w źródłach openpyxl 3.1.x. Cztery ustalenia **prostują brief** i zostaną w module nazwane wprost:

- **`NumberAxis` nie istnieje w 3.1.x** — klasa nazywa się `NumericAxis` (pełny `axis.py` nie definiuje żadnego aliasu). Istnieją: `NumericAxis`, `TextAxis`, `DateAxis`, `SeriesAxis`.
- **Wykresy w trybie `write_only` DZIAŁAJĄ** (oficjalny przykład `bar.html` używa `Workbook(write_only=True)` + `ws.add_chart`). Niedostępne są wyłącznie w `read_only` — potwierdzone w źródle `reader/excel.py`: gałąź read-only robi `continue` *przed* przetwarzaniem rysunków.
- **`DifferentialStyle`… nie, tu:** `chart.legend = None` usuwa legendę (dokumentowany sposób); `legend.delete` nie występuje w API. Legenda jest domyślnie obiektem `Legend()` z `position` ∈ {`r`,`l`,`t`,`b`,`tr`}.
- **`set_categories()` iteruje po `self.ser`** — wywołane **przed** `add_data()` nie ustawi niczego i **nie zgłosi błędu**.

Dodatkowo potwierdzone: `BarChart(barDir="col", grouping="clustered", gapWidth=150)`, domyślne `axId` (TextAxis 10, NumericAxis 100, DateAxis 500, SeriesAxis 1000), `NumericAxis` ma `majorGridlines = ChartLines()` **domyślnie włączone**, `Marker.symbol` ∈ {`circle`,`dash`,`diamond`,`dot`,`plus`,`square`,`star`,`triangle`,`x`,`auto`,`picture`}, `dLblPos` ∈ {`l`,`t`,`bestFit`,`r`,`b`,`outEnd`,`inEnd`,`inBase`,`ctr`}, `chart.width`/`chart.height` domyślnie 15 × 7,5 cm.

---

# Moduł 15 — Wykresy: opis, nie obrazek

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–14

## 0. W tym module nauczysz się

- **Zrozumiesz kluczową ideę:** wykres w Excelu to **nie obrazek, a opis**. To zbiór instrukcji („weź te komórki, narysuj je w ten sposób"), który Excel wykonuje od nowa przy każdym otwarciu pliku. Zobaczysz, dlaczego to rozróżnienie zmienia sposób myślenia o całym API.
- **Poznasz trzy filary każdego wykresu** i zrozumiesz, dlaczego są właśnie takie: `Reference` (adres zakresu komórek), `Series` (seria danych), `Chart` (typ wykresu i jego wygląd). Dowiesz się, że openpyxl **nie kopiuje danych do wykresu** — wykres trzyma **adres**, dokładnie jak formuła (moduł 06).
- **Opanujesz pełny warsztat:** `chart.add_data()`, `chart.set_categories()`, `ws.add_chart()`, oś X i Y, tytuły, legenda, etykiety danych, kolory serii, znaczniki, siatka, skalowanie osi, rozmiar i pozycja.
- **Nauczysz się budować wykres z dwiema osiami** — klasyczny scenariusz „sprzedaż (słupki) + marża % (linia)" — i zrozumiesz mechanikę `axId`/`crosses`, która za tym stoi.
- **Zobaczysz tabelę wszystkich typów wykresów** obsługiwanych przez openpyxl z gotowymi snippetami: słupkowy, liniowy, kołowy, pierścieniowy, punktowy, obszarowy, radarowy, bąbelkowy, giełdowy, powierzchniowy.
- **Dowie się wprost, czego openpyxl NIE potrafi:** nowych typów wykresów z rodziny „chartex" (mapa, treemap, sunburst, histogram, Pareto, waterfall, funnel, pudełko-wąsy), **sparkline'ów**, wykresów przestawnych. Zobaczysz też, **co openpyxl gubi przy zapisie** i jak to sprawdzić.
- **Zrozumiesz ograniczenia trybów pracy:** wykresy **nie są dostępne** w `read_only`, ale **są dostępne** w `write_only` — i dlaczego tak jest.

## 1. Intuicja i analogia

### 1.1. Wykres to przepis kulinarny, nie potrawa

To najważniejsze zdanie w całym module. Przeczytaj je jeszcze raz.

Wyobraź sobie **przepis w książce kucharskiej**. Przepis nie jest jedzeniem. To lista instrukcji: „weź 200 g mąki, 2 jajka, wymieszaj, piecz 20 minut". Jeśli zmienisz w spiżarni mąkę na razową, a potem ugotujesz z tego przepisu — danie wyjdzie inne. Przepis się nie zmienił, zmieniły się **składniki**, do których przepis się odwołuje.

Wykres w Excelu działa dokładnie tak samo:

- **Przepis** = definicja wykresu („weź zakres `B2:B13`, narysuj jako słupki, tytuł: Sprzedaż 2026"). W pliku `.xlsx` to osobny plik XML: `xl/charts/chart1.xml`.
- **Składniki** = komórki arkusza, do których wykres się odwołuje.
- **Danie** = obraz, który widzisz po otwarciu pliku. Excel przygotowuje go **od nowa za każdym razem**.

Z tego wynikają trzy konsekwencje, które będą wracać w całym module:

**Konsekwencja 1 — zmiana danych w komórce zmienia wykres.** Nie musisz nic robić z wykresem. Zmieniasz `B3` z `135` na `500`, zapisujesz — słupek rośnie. Bo wykres wskazuje na `B3`, a nie na „liczbę 135". To bardzo praktyczne: nie musisz przeliczać niczego ręcznie.

**Konsekwencja 2 — usunięcie komórek psuje wykres.** Wyczyść zakres, usuń arkusz, przesuń dane w inne miejsce — wykres pokaże pustkę albo błąd. Bo przepis wskazuje na adres, którego już nie ma. Analogia: przepis mówi „weź mąkę z górnej półki", a ty przestawiłeś mąkę — sięgniesz w pustą przestrzeń.

**Konsekwencja 3 — openpyxl nie „rysuje" niczego i nie musi.** To jest ważne dla zrozumienia ograniczeń. Pytałeś kiedyś, dlaczego openpyxl nie musi mieć Excela zainstalowanego? Bo nie liczy formuł (moduł 06) i **nie renderuje wykresów**. Zapisuje przepis. Excel (albo LibreOffice, albo cokolwiek otworzy plik) gotuje danie. openpyxl jest **skrybą piszącym przepisy**.

### 1.2. `Reference` to palec wskazujący, a nie koszyk z zakupami

Skoro wykres trzyma adres, musimy umieć ten adres zapisać. Do tego służy `Reference`.

Wyobraź sobie, że stoisz w magazynie i **wskazujesz palcem półkę**: „od tego pudełka tutaj, do tamtego tam". Nie przenosisz pudełek. Nie robisz zdjęcia. **Wskazujesz.**

`Reference(ws, min_col=2, min_row=1, max_col=2, max_row=13)` to właśnie taki gest wskazania: „w arkuszu `ws`, od kolumny 2, wiersza 1, do kolumny 2, wiersza 13". W pliku zapisze się to jako tekst `'Sprzedaz'!$B$1:$B$13`.

Porównaj to z tym, co robiliśmy w poprzednich modułach:

| Sposób | Co trzyma | Kiedy zmienię komórkę | Czy jest kopią danych? |
|---|---|---|---|
| `ws["A1"] = dane` | **wartość** | komórka zmienia się | tak, to zapis |
| `ws["C1"] = "=SUM(A1:A5)"` | **adres** (formuła) | Excel przelicza przy otwarciu | nie, to wskazanie |
| `Reference(ws, ...)` + wykres | **adres** | Excel przerysowuje przy otwarciu | nie, to wskazanie |

Trzecia linia jest dokładnie tą samą kategorią co druga. **Wykres i formuła to ten sam gatunek:** oba są przepisami, oba wskazują adresy, oba są wykonywane przez Excela, a nie przez openpyxl. Twoja wiedza z modułu 06 przenosi się tu w całości.

### 1.3. `add_chart` to pinezka w tablicy korkowej

Gdzie „mieszka" wykres? To pytanie zaskakuje początkujących, bo intuicja podpowiada „w komórce". Nie — **wykres nie jest w komórce**.

Wyobraź sobie **tablicę korkową z przypiętą kartką**. Kartka wisi **nad** tablicą, a pinezka trzyma ją w jednym punkcie. Jeśli tablica ma narysowaną siatkę (komórki), kartka zasłania część kratek, ale sama nie jest żadną z nich.

Tak właśnie działa wykres:

- **Kartka** = obiekt wykresu.
- **Pinezka** = *anchor*, czyli komórka, do której przyczepiona jest **lewy górny róg** wykresu. W kodzie: `ws.add_chart(chart, "E15")`.
- **Siatka** = komórki arkusza. Wykres „pływa" nad nimi.

Konsekwencje praktyczne, które zbijają ludzi z tropu:

1. **Wykres zasłania dane.** Jeśli masz dane w `A1:D20` i wstawisz wykres w `E15`, nic nie zasłoni. Ale gdy wpiszesz `ws.add_chart(chart, "B5")` — wykres **przykryje** twoje dane. W Excelu nadal tam są, ale ich nie widzisz. To najczęstsza „pomyłka estetyczna" w raportach.
2. **Sortowanie i filtrowanie nie przenoszą wykresu.** Wykres jest przyczepiony do komórki-kotwicy, a nie do danych. (`Reference` dynamicznie się przesuwa w Excelu przy wstawianiu wierszy, ale sam wykres zostaje tam, gdzie jest pinezka.)
3. **Scalanie komórek psuje układ.** Scalone komórki zmieniają geometrię siatki, a wykres liczy swoją pozycję w jednostkach komórek.
4. **Rozmiar jest podany w centymetrach**, a nie w komórkach: domyślnie $15 \times 7{,}5$ cm, co dokumentacja opisuje jako „w przybliżeniu 5 kolumn na 14 wierszy". Zmieniasz przez `chart.width` i `chart.height`.

### 1.4. Druga oś to druga linijka na tej samej kartce

Wyobraź sobie rysunek przekroju drogi. Obok narysowano **dwie linijki**:

- jedna w **centymetrach** (grubość warstw asfaltu: od 0 do 20),
- druga w **metrach** (szerokość jezdni: od 0 do 12).

Ten sam rysunek, dwie różne skale. Bez drugiej linijki narysowanie obu rzeczy na jednej skali byłoby bezużyteczne — jezdnia o szerokości 12 m zamieniłaby grubość asfaltu 5 cm w niewidoczną kreskę.

**Druga oś** w Excelu robi dokładnie to. Przykład z życia: „sprzedaż w tysiącach złotych" i „marża w procentach". Sprzedaż to liczby rzędu 100–250, marża to 0,18–0,25. Na jednej skali procenty byłyby płaską linią przy dnie wykresu.

Kluczowe zdanie do zapamiętania: **w openpyxl druga oś to nie „ustawienie wykresu", a drugi wykres.** Tworzysz osobny `LineChart` dla marży, nadajesz mu **nowy identyfikator osi** (`axId`), a potem **sklejasz oba wykresy** w jeden (`chart1 += chart2`). Odziedziczyliśmy to wprost ze specyfikacji OOXML — i w 2.8 zobaczysz, jaką pułapkę z tego powodu nosi ze sobą parametr `crosses`.

### 1.5. Wykres to obywatel drugiej kategorii

Wróćmy na chwilę do trzech kategorii z modułu 03 — „potrafi", „potrafi częściowo", „gubi przy zapisie". Wykresy są **najlepszym przykładem kategorii drugiej** w całym kursie.

openpyxl **potrafi** stworzyć wykres od zera, ze wszystkimi typami klasycznymi. **Potrafi częściowo** odtworzyć wykres wczytany z pliku Excela — bo model OOXML opisuje setki drobnych opcji, z których openpyxl implementuje większość, ale nie wszystkie. I **gubi** całe rodziny wykresów, których nie zna (sparkline'y, chartex).

To dlatego moduł 18 (modyfikacja istniejących plików) będzie miał osobny akapit o wykresach. Zapamiętaj zasadę: **wykres wygenerowany przez openpyxl od zera jest w pełni kontrolowany; wykres z Excela przepuszczony przez „wczytaj–zapisz" to loteria, którą trzeba sprawdzić okiem.**

## 2. Teoria

### 2.1. Gdzie w pliku mieszka wykres (powrót do modułu 02)

Zanim wejdziemy w API, zajrzyjmy do środka pliku. To rozwieje wątpliwości raz na zawsze.

Wygeneruj skoroszyt z jednym wykresem (zrobisz to w Przykładzie 1), zmień rozszerzenie na `.zip`, rozpakuj i popatrz:

```text
raport.xlsx (rozpakowany)
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── xl/
│   ├── workbook.xml
│   ├── styles.xml
│   ├── charts/
│   │   ├── chart1.xml          ← PRZEPIS: definicja wykresu
│   │   └── _rels/
│   │       └── chart1.xml.rels ← powiązania wykresu (np. do arkusza danych)
│   ├── drawings/
│   │   ├── drawing1.xml        ← POZYCJA: gdzie wykres leży (anchor!)
│   │   └── _rels/
│   │       └── drawing1.xml.rels ← powiązanie rysunku z plikiem wykresu
│   └── worksheets/
│       ├── sheet1.xml
│       └── _rels/
│           └── sheet1.xml.rels  ← powiązanie arkusza z rysunkiem
```

Popatrz, ile ogniw pośredniczy. To wyjaśnia dwie rzeczy, o których mówiłem w 1.3:

- **Wykres ma osobny plik** (`xl/charts/chart1.xml`) — bo jest „kartką", a nie komórką.
- **Wykres nie jest częścią `sheet1.xml`.** Arkusz jedynie **odwołuje się do rysunku** przez plik relacji. Dlatego wykres może „przepłynąć" nad siatką, nie zaburzając samych komórek.

W `chart1.xml` zobaczysz mniej więcej taką strukturę (uproszczoną):

```xml
<c:chartSpace xmlns:c="...">
  <c:chart>
    <c:title>...</c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>'Sprzedaz'!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:numRef><c:f>'Sprzedaz'!$A$2:$A$13</c:f></c:numRef></c:cat>
          <c:val><c:numRef><c:f>'Sprzedaz'!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="10"/>
        <c:axId val="100"/>
      </c:barChart>
      <c:catAx>...</c:catAx>
      <c:valAx>...</c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

Zwróć uwagę na trzy rzeczy w tym XML-u, bo to sedno modułu:

1. **`<c:f>` to formuła-adres**, dokładnie jak w komórce: `'Sprzedaz'!$B$2:$B$13`. Wykres nie ma kopii danych — ma adres. Dokładnie o tym mówiłem w 1.2.
2. **`<c:idx>` i `<c:order>`** — numer serii. To one decydują o kolejności serii w legendzie i o kolorach. Wspomnę o nich w 2.13, bo są źródłem znanych problemów.
3. **`<c:axId val="10">` i `val="100"`** — identyfikatory osi. Zapamiętaj liczby `10` i `100`; za chwilę (2.8) okażą się kluczowe dla drugiej osi.

### 2.2. `Reference` — bo adres trzeba jakoś zapisać

`Reference` to obiekt opisujący zakres komórek. Pełna sygnatura ze źródła 3.1.x:

```python
Reference(worksheet=None, min_col=None, min_row=None,
          max_col=None, max_row=None, range_string=None)
```

Możesz go zbudować na dwa sposoby:

```python
# Sposób 1: wspolrzedne (najczestszy)
Reference(ws, min_col=2, min_row=1, max_col=2, max_row=13)  # B1:B13

# Sposób 2: gotowy napis zakresu (gdy masz tekst)
Reference(range_string="'Sprzedaz'!$B$1:$B$13")
```

**Filozofia `Reference`:** to *leniwe* wskazanie. Obiekt nie czyta żadnej komórki. Nie wiesz, ile w nim jest danych. Jeśli wskazujesz `$B$1:$B$1000` w pustym arkuszu — to nie błąd. Wykres powstanie, będzie w nim 999 pustych punktów, a Excel pokaże pusty wykres. **openpyxl nie waliduje, czy zakres ma sens** — to ta sama kategoria co w module 06 („skryba, nie matematyk").

**Praktyczna konsekwencja — licz zakres z danych, nie na sztywno.** Dokładnie jak w module 07 (i w module 13 przy `sqref`, i w module 12 przy `ref` tabeli):

```python
ostatni_wiersz = ws.max_row           # uwaga na pulapke z modulu 07!
dane = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=ostatni_wiersz)
```

**Sprostowanie do briefu:** parametr `min_col`, `max_col` itd. to **liczby** (indeksy kolumn, licząc od 1). W notacji A1 `B` to kolumna 2. Jeśli masz literę, użyj `column_index_from_string("B")` (moduł 07).

### 2.3. `add_data` i `titles_from_data` — z czego zrobić serię

Mając zakres, dodajesz go do wykresu:

```python
chart.add_data(data, titles_from_data=True, from_rows=False)
```

Trzy rzeczy do zrozumienia:

**1. Kierunek (`from_rows`).** Domyślnie `from_rows=False`, co znaczy: **każda kolumna zakresu to jedna seria**. Wykres „Sprzedaż / Koszty / Zysk" to trzy kolumny → trzy serie. Jeśli masz dane ułożone odwrotnie (każdy wiersz to seria), użyj `from_rows=True`. To ta sama dychotomia co `dataframe_to_rows` w module 07.

**2. `titles_from_data=True`.** Pierwszy wiersz (albo kolumna) zakresu **nie jest danymi** — służy jako **etykieta serii**. Analogia: etykieta na pudełku — nie jest towarem w środku, tylko opisem. W pliku trafi to do `<c:tx><c:strRef><c:f>'Sprzedaz'!$B$1</c:f></c:strRef></c:tx>` — czyli nazwa serii to **adres komórki z nagłówkiem**, nie skopiowany tekst.

**3. `add_data` można wywołać wielokrotnie** — każdy zakres doda kolejne serie. To jest podstawa wykresów kombinowanych i drugiej osi.

**Pułapka, która ma własną sekcję:** jeśli `titles_from_data=True`, to zakres danych „zużywa" pierwszy wiersz na tytuł i **wartości zaczyna od drugiego wiersza** — openpyxl robi to automatycznie. Ale **zakres kategorii musisz podać samodzielnie i bez nagłówka**. To najczęstszy błąd początkujących i wrócę do niego w 2.4 oraz w Pułapkach.

### 2.4. `set_categories` — i cicha pułapka kolejności

Kategorie to etykiety osi X („Sty", „Lut", „Mar"...). Ustawiasz je tak:

```python
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))
```

Zwróć uwagę: `min_row=2`, **bez nagłówka**. Bo kategorie to nazwy miesięcy, a nie nagłówek „Miesiąc". Jeśli podasz `min_row=1`, na wykresie pojawi się dodatkowa kategoria „Miesiąc" i wszystko przesunie się o jedno miejsce.

> **⚠️ Cicha pułapka — kolejność wywołań.** W źródle openpyxl metoda `set_categories` iteruje po **już istniejących seriach** i przypisuje im kategorie. Jeśli wywołasz ją **przed** `add_data`, lista serii jest pusta, pętla wykona się zero razy, a openpyxl **nie zgłosi żadnego błędu**. Otworzysz plik i zobaczysz wykres bez etykiet na osi X — i będziesz szukać przyczyny w dobrym miejscu, czyli nigdzie.
>
> **Reguła:** najpierw `add_data`, potem `set_categories`. Zawsze.

To dobra ilustracja zasady z modułu 03: openpyxl nie waliduje spójności twoich działań. Skrypt może się „udać" i wyprodukować plik, który jest w połowie zrobiony.

### 2.5. Anatomia wykresu — co ma każdy wykres

Każdy obiekt wykresu w openpyxl dziedziczy z `ChartBase` / `_ChartBase` i ma kilka wspólnych elementów. To twoja mapa na resztę modułu:

| Element | Atrybut | Co to jest |
|---|---|---|
| **Tytuł** | `chart.title` | Napis nad wykresem. Przyjmuje `str` (openpyxl opakuje w `Title`) |
| **Obszar wykresu** | `chart.plotArea` | Wewnętrzny obiekt; tworzony automatycznie |
| **Osie** | `chart.x_axis`, `chart.y_axis` (i `chart.z_axis` w 3D) | Osobne obiekty osi |
| **Serie** | `chart.series` | Lista obiektów `Series` (tworzona przez `add_data`) |
| **Legenda** | `chart.legend` | Domyślnie `Legend()`; `None` usuwa legendę |
| **Etykiety danych** | `chart.dataLabels` (alias `chart.dLbls`) | `DataLabelList` — pokazywanie liczb na słupkach |
| **Styl** | `chart.style` | Liczba 1–48, presety kolorów Excela |
| **Tło i obramowanie** | `chart.graphical_properties` | `GraphicalProperties` (sekcja 2.10) |
| **Rozmiar** | `chart.width`, `chart.height` | W centymetrach (domyślnie $15 \times 7{,}5$) |
| **Kotwica** | `chart.anchor` | Komórka lub obiekt anchora |

Kilka z tych atrybutów ma **aliasy**. To nie przypadek — openpyxl celowo ukrywa pod czytelniejszą nazwą surowe nazwy OOXML (np. `spPr`, `dLbls`, `barDir`):

```python
chart.dataLabels  is  chart.dLbls          # alias
chart.type        is  chart.barDir         # alias (tylko wykresy słupkowe)
chart.graphical_properties is chart.graphicalProperties ... # nie laczyc z aliasami serii
```

W praktyce: **używaj czytelnej wersji** (`chart.dataLabels`, `chart.type`) — i wiedz, że istnieje druga, bo zobaczysz ją w XML-u i w internecie.

### 2.6. Typy wykresów — katalog z gotowymi snippetami

openpyxl implementuje **klasyczne wykresy OOXML**. Pełna lista dostępnych klas (importowane z `openpyxl.chart`):

| Klasa | Do czego | Kluczowe ustawienia | Kiedy używać |
|---|---|---|---|
| `BarChart` | słupki pionowe/poziome | `type="col"/"bar"`, `grouping`, `gapWidth`, `overlap` | porównanie wartości między kategoriami |
| `BarChart3D` | jak wyżej, w 3D | `shape`, `gapDepth`, `z_axis` | rzadko; gorzej czytelne |
| `LineChart` | linie | `grouping`, znaczniki w `Series.marker` | trendy w czasie |
| `LineChart3D` | linie 3D | — | rzadko |
| `AreaChart` | obszar pod linią | `grouping="standard"/"stacked"/"percentStacked"` | kumulacja w czasie |
| `AreaChart3D` | obszar 3D | — | rzadko |
| `PieChart` | kołowy | **tylko jedna seria**, `firstSliceAng` | udział w całości |
| `ProjectedPieChart` | kołowy z wyciągniętymi plasterkami | `type="pie"/"bar"`, `splitType="percent"/"val"/"pos"` | gdy jest dużo drobnych kategorii |
| `PieChart3D`, `DoughnutChart` | kołowy 3D / pierścieniowy | `firstSliceAng`, `holeSize` | estetyka, „gauge" |
| `ScatterChart` | punktowy (X vs Y) | `Series(y, x)` | korelacja dwóch zmiennych |
| `BubbleChart` | bąbelkowy | `Series(y, x, zvalues=...)` | trzy wymiary danych |
| `RadarChart` | radarowy / „pajęczy" | `type="standard"/"marker"/"filled"` | profile wielu cech |
| `StockChart` | giełdowy | `hiLowLines`, `upDownBars` | dane OHLC; też temperatury |
| `SurfaceChart`, `SurfaceChart3D` | powierzchniowy | `wireframe` | funkcje dwóch zmiennych |

Szybkie snippety (wszystkie zweryfikowane wobec dokumentacji 3.1.x):

```python
from openpyxl.chart import (
    BarChart, LineChart, AreaChart, PieChart, DoughnutChart,
    ScatterChart, BubbleChart, RadarChart, StockChart, SurfaceChart,
    Reference, Series,
)

# --- SLUPKOWY (pionowy) ---
c = BarChart()
c.type = "col"; c.grouping = "clustered"; c.gapWidth = 150
c.add_data(Reference(ws, min_col=2, min_row=1, max_row=13), titles_from_data=True)
c.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))

# --- SLUPKOWY WARSTWOWY (stacked) - UWAGA: overlap = 100 ---
c = BarChart()
c.type = "col"; c.grouping = "stacked"; c.overlap = 100

# --- LINIOWY ---
c = LineChart()
c.grouping = "standard"
c.add_data(Reference(ws, min_col=2, min_row=1, max_row=13), titles_from_data=True)
c.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))

# --- OBSZAROWY ---
c = AreaChart(); c.grouping = "standard"

# --- KOLOWY (JEDNA seria!) ---
c = PieChart()
c.add_data(Reference(ws, min_col=2, min_row=1, max_row=6), titles_from_data=True)
c.set_categories(Reference(ws, min_col=1, min_row=2, max_row=6))

# --- PIERSCIENIOWY ---
c = DoughnutChart(firstSliceAng=270, holeSize=50)

# --- PUNKTOWY (inaczej niz reszta! Serie budujesz recznie) ---
c = ScatterChart()
xvals = Reference(ws, min_col=1, min_row=2, max_row=10)
yvals = Reference(ws, min_col=2, min_row=1, max_row=10)   # z naglowkiem
c.series.append(Series(yvals, xvals, title_from_data=True))

# --- BABELKOWY ---
c = BubbleChart()
c.series.append(Series(values=yvals, xvalues=xvals, zvalues=zvals, title="2026"))

# --- RADAROWY ---
c = RadarChart(); c.type = "marker"

# --- POWIERZCHNIOWY ---
c = SurfaceChart(); c.wireframe = False

ws.add_chart(c, "E2")
```

**Dwie rzeczy, które odróżniają `ScatterChart` od reszty.** Po pierwsze, w wykresach punktowych **nie używa się `add_data`** — serie budujesz ręcznie przez `Series(yvalues, xvalues)`. Po drugie, istnieje atrybut `scatterStyle`, ale dokumentacja ostrzega wprost: *„the specification says that there are the following types of scatter charts: 'line', 'lineMarker', 'marker', 'smooth', 'smoothMarker'. However, at least in Microsoft Excel, this is just a shortcut for other settings that otherwise have no effect."* Czyli: **`scatterStyle` nie działa w Excelu** — styl każdej serii ustawiaj ręcznie (`series.marker`, `series.graphicalProperties.line`).

### 2.7. Osie — co ma która i co możesz ustawić

Osie to najbardziej niedoceniana część wykresu. Zajmijmy się nimi porządnie.

**Rodzaje osi.** Klasa każdej osi zależy od typu wykresu (nie tworzysz ich ręcznie — powstają same):

| Klasa | `tagname` | Kiedy | Domyślne `axId` |
|---|---|---|---|
| `TextAxis` | `catAx` | oś kategorii (np. miesiące) | 10 |
| `NumericAxis` | `valAx` | oś wartości (liczby) | 100 |
| `DateAxis` | `dateAx` | oś czasu (daty z osią „datową") | 500 |
| `SeriesAxis` | `serAx` | trzecia oś w wykresach 3D | 1000 |

> **⚠️ Sprostowanie do briefu.** W openpyxl 3.1.x **nie ma klasy `NumberAxis`**. Klasa nosi nazwę **`NumericAxis`** (i nie ma aliasu — sprawdziłem pełny kod `openpyxl/chart/axis.py`). Jeśli gdzieś zobaczysz `from openpyxl.chart import NumberAxis`, to kod ze starej wersji albo z tutoriala, który jej nie sprawdził. Poprawnie: `from openpyxl.chart.axis import NumericAxis`.

**Wspólne atrybuty obu głównych osi** (`_BaseAxis`) — pełna lista ze źródła:

| Atrybut | Wartości | Znaczenie |
|---|---|---|
| `title` | `str` albo `Title` | Tytuł osi. `chart.x_axis.title = "Miesiąc"` |
| `numFmt` (alias `number_format`) | kod formatu (moduł 10) | Format etykiet osi, np. `'#,##0'` |
| `scaling` | obiekt `Scaling` | Skala: `min`, `max`, `logBase`, `orientation` |
| `majorGridlines` | `ChartLines()` albo `None` | Siatka główna. **Domyślnie WŁĄCZONA na osi wartości!** |
| `minorGridlines` | `ChartLines()` albo `None` | Siatka pomocnicza |
| `delete` | `True`/`False` | **Ukrycie całej osi** |
| `crosses` | `'autoZero'`, `'max'`, `'min'` | Gdzie oś przecina drugą (sekcja 2.8) |
| `axId` | `int` | Identyfikator osi (sekcja 2.8) |
| `majorTickMark`, `minorTickMark` | `'cross'`, `'in'`, `'out'` | „Kreseczki" na osi |
| `tickLblPos` | `'high'`, `'low'`, `'nextTo'` | Pozycja etykiet |

**Kluczowa niespodzianka: siatka jest domyślnie włączona.** Spójrz na kod źródłowy `NumericAxis`:

```python
def __init__(self, crossBetween=None, majorUnit=None, minorUnit=None,
             dispUnits=None, extLst=None, **kw):
    ...
    kw.setdefault('majorGridlines', ChartLines())   # ← domyślnie tworzy siatkę!
    kw.setdefault('axId', 100)
    kw.setdefault('crossAx', 10)
```

Czyli gdy piszesz `chart = BarChart()`, nowa oś Y **już ma** siatkę. Żeby ją usunąć, przypisujesz `None`:

```python
chart.y_axis.majorGridlines = None   # wylacz siatke osi Y
```

To ważne przy wykresach kombinowanych (dwie osie = dwie siatki = wizualny bałagan). Wrócę do tego w 2.9.

**Skalowanie.** Obiekt `Scaling` (sygnatura: `Scaling(logBase=None, orientation='minMax', max=None, min=None)`):

```python
chart.y_axis.scaling.min = 0        # wymus dolny prog osi
chart.y_axis.scaling.max = 300      # wymus gorny prog
chart.y_axis.scaling.logBase = 10   # skala logarytmiczna (potegi 10)
chart.x_axis.scaling.orientation = "maxMin"  # ODWROC os (kierunek malejacy)
```

Uwaga na dwie rzeczy: `orientation` przyjmuje dokładnie `'minMax'` (normalna) albo `'maxMin'` (odwrócona). A `logBase` na osi X **odrzuca wartości ujemne** (logarytm z liczby ujemnej nie istnieje). Dokumentacja dodaje cenną uwagę wydajnościową: *„For large datasets, rendering of scatter plots will be much faster when using subsets of the data rather than axis limits."* — czyli na dużych zbiorach lepiej **przyciąć dane przed wykresem**, niż ukrywać je ustawieniem osi. Excel i LibreOffice wciąż muszą przetworzyć wszystkie punkty.

**Ukrycie osi.** Czasem chcesz wykres „bez osi" (np. minimalistyczny dashboard):

```python
chart.x_axis.delete = True
chart.y_axis.delete = True
```

**Format etykiet osi.** Tu wraca cały moduł 10:

```python
chart.y_axis.numFmt = '#,##0 "zł"'
chart.y_axis.numFmt = '0.0%'
```

**Oś dat.** Jeśli masz kolumnę z prawdziwymi datami (a nie tekstem — moduł 05!), możesz chcieć `DateAxis`:

```python
from openpyxl.chart.axis import DateAxis
chart.x_axis = DateAxis()
```

Wtedy kategorie to prawdziwa os czasu — Excel nie rysuje odstępów proporcjonalnie do odstępu czasu sam z siebie, ale z `DateAxis` możesz ustawić `majorTimeUnit` (`days`/`months`/`years`) itd.

### 2.8. Druga oś — mechanika, którą trzeba zrozumieć raz

Wróćmy do obrazu dwóch linijek (1.4). Jak to zrobić w openpyxl?

**Krok po kroku, na przykładzie sprzedaży (słupki) i marży (linia):**

1. **Zbuduj pierwszy wykres** (słupki ze sprzedażą) — normalnie, jak w 2.3–2.4.
2. **Zbuduj DRUGI wykres** (linia z marżą) — jako osobny obiekt `LineChart`.
3. **Nadaj drugiemu wykresowi nowy `axId` dla osi Y.** Domyślnie oś wartości ma `axId = 100`. Musisz nadać inny, żeby Excel wiedział, że to **inna** oś:
   ```python
   c2.y_axis.axId = 200
   ```
4. **Ustaw `crosses = "max"` na osi, która ma się pokazać po prawej stronie.** W dokumentacji openpyxl ustawia się to na osi Y **pierwszego** wykresu:
   ```python
   c1.y_axis.crosses = "max"
   ```
5. **Sklej wykresy:**
   ```python
   c1 += c2
   ```
6. **Dodaj do arkusza** — `ws.add_chart(c1, "E2")`. Dodajesz `c1`, czyli ten po lewej stronie operatora `+=`.

**Co się dzieje w pliku?** openpyxl zapisuje plotArea z **dwoma** wykresami (`barChart` i `lineChart`) i **czterema** osiami (dwie wspólne osie X + dwie osie Y o różnych `axId`). Dokładnie tak Excel to modeluje.

**⚠️ Uczciwa uwaga o `crosses`.** Dokumentacja openpyxl ustawia `c1.y_axis.crosses = "max"` i komentuje to jako *„Display y-axis of the second chart on the right by setting it to cross the x-axis at its maximum"*. W wielu przykładach w internecie zobaczysz to samo ustawienie na `c2.y_axis`. **To jest obszar, w którym trzeba patrzeć na efekt w Excelu, a nie ufać kodowi.** Trzymaj się wzorca z dokumentacji, obejrzyj plik, a jeśli oś jest po złej stronie — przenieś `crosses = "max"` na drugą oś. To dokładnie ta sama lekcja co w module 14 (priorytety reguł) i w module 06 (formuły): **openpyxl nie gwarantuje, że zrobił dobrze — gwarantuje, że zapisał to, co kazałeś.**

**Wykresy kombinowane bez drugiej osi.** Jeśli chcesz tylko **jeden** wykres z seriami różnych typów (np. słupki i linia na **jednej** skali — rzadko sensowne, ale możliwe), mechanizm jest ten sam: `chart1 += chart2`, tylko **bez** zmiany `axId`.

**Kto dziedziczy tytuł?** Przy `c1 += c2` tytuł, legenda itd. pochodzą z `c1`. `c2` dostarcza tylko swoją definicję plotArea (serie + osie). Dlatego styl i tytuł ustawiasz na `c1`.

### 2.9. Legenda i etykiety danych

**Legenda.**

```python
chart.legend.position = "b"     # dol, dozwolone: r, l, t, b, tr
chart.legend = None             # brak legendy
```

Domyślnie legenda **jest** tworzona (`BarChart.__init__` ustawia `self.legend = Legend()`), a jej pozycja to `r` (prawa). Przy jednej serii legenda zwykle jest zbędna — i wtedy `chart.legend = None` to najprostszy sposób, żeby raport był czytelniejszy.

> **Sprostowanie do briefu.** `chart.legend.delete = True` **nie jest** udokumentowaną drogą w 3.1.x. Usuwanie legendy robi się przez `chart.legend = None` (tak robią wszystkie oficjalne przykłady openpyxl). Możesz też sterować **ręcznie** układem legendy przez `chart.legend.layout = Layout(manualLayout=ManualLayout(...))` — ale to rzadko potrzebne.

**Etykiety danych.** Chcesz pokazać liczby na słupkach? Użyj `DataLabelList`:

```python
from openpyxl.chart.label import DataLabelList

chart.dataLabels = DataLabelList()
chart.dataLabels.showVal = True        # pokaz wartosci
chart.dataLabels.numFmt = '#,##0'      # format wartosci na etykiecie
chart.dataLabels.dLblPos = "outEnd"    # gdzie umiescic (l/t/r/b/ctr/outEnd/inEnd/inBase/bestFit)
```

Pełna lista przełączników `DataLabelList` (wszystkie `bool`, wszystkie ze źródła 3.1.x):

| Atrybut | Pokazuje |
|---|---|
| `showVal` | wartość liczbową |
| `showCatName` | nazwę kategorii (np. „Sty") |
| `showSerName` | nazwę serii |
| `showPercent` | udział procentowy (kluczowe dla wykresów kołowych!) |
| `showLegendKey` | znacznik legendy |
| `showBubbleSize` | rozmiar bąbelka (wykresy bąbelkowe) |
| `showLeaderLines` | linie wiodące (kołowe/projected) |
| `numFmt` | format liczby (moduł 10!) |
| `dLblPos` | pozycja: `'l'`, `'t'`, `'r'`, `'b'`, `'ctr'`, `'outEnd'`, `'inEnd'`, `'inBase'`, `'bestFit'` |
| `separator` | separator między elementami etykiety |

**Ważna uwaga o `dLblPos`.** Nie każdy typ wykresu honoruje każdą pozycję. `'outEnd'` ma sens dla słupków; dla kołowego Excel zwykle użyje `'bestFit'`; dla liniowego pozycje są ograniczone. Zasada jak wyżej: **ustaw, obejrzyj, popraw.**

**Etykiety na poziomie serii.** Możesz też dać etykiety tylko jednej serii:

```python
chart.series[0].dLbls = DataLabelList()
chart.series[0].dLbls.showVal = True
```

Nazwy: `Series.dLbls` to nazwa techniczna, `Series.labels` to alias. Oba działają, ale w dokumentacji i przykładach używa się `dLbls` (patrz przykład „Reusing XML" w dokumentacji openpyxl).

### 2.10. Stylowanie: kolory serii, znaczniki, tło

Tu wchodzimy do **bardzo abstrakcyjnego** kawałka API. Dokumentacja sama ostrzega:

> *„Many advanced options require using the Graphical Properties of OOXML. This is a much more abstract API than the chart API itself and may require considerable studying of the OOXML specification to get right. It is often unavoidable to look at the XML source of some charts you've made."*

Posłuchaj tej rady. Nie ucz się tego na pamięć — naucz się **wzorca**: kolor serii siedzi w `graphicalProperties`, a nie w jakimś intuicyjnym `color`.

**Kolor słupka/linii:**

```python
from openpyxl.chart.shapes import GraphicalProperties

chart.series[0].graphicalProperties = GraphicalProperties(solidFill="5B5CF0")
```

**Dlaczego `graphicalProperties`, a nie `color`?** Bo w OOXML wygląd każdego elementu rysunku (słupek, linia, znacznik, tło wykresu, obramowanie) opisuje ten sam typ elementu: `spPr` (ang. *shape properties*). openpyxl nadaje mu czytelniejszy alias `graphicalProperties`. Nie ma osobnego `color` dla serii, bo kolor to tylko jedna z wielu właściwości kształtu (obok obrysu, wypełnienia gradientem, przezroczystości, cienia...).

**Rodzina właściwości w `GraphicalProperties`:**

| Właściwość | Przykład | Uwagi |
|---|---|---|
| `solidFill` | `GraphicalProperties(solidFill="FF0000")` | jednolity kolor; kanał alfa opcjonalny |
| `noFill` | `gp.noFill = True` | brak wypełnienia (przezroczyste) |
| `line` (alias `ln`) | `gp.line.solidFill = "0070C0"` | obramowanie/kreska |
| `gp.line.noFill` | `= True` | usuwa obramowanie |
| `gp.line.width` | `= 40000` | grubość w EMU (1 pt ≈ 12700) |
| `gp.line.prstDash` | `= "solid"`, `"dash"` | styl kreski |
| `gradFill` | `GradientFillProperties(...)` | wypełnienie gradientem |

**Tło wykresu przezroczyste** (dokumentowany przykład):

```python
from openpyxl.chart.shapes import GraphicalProperties
chart.graphical_properties = GraphicalProperties()
chart.graphical_properties.noFill = True
```

**Usunięcie obramowania wykresu:**

```python
chart.graphical_properties.line.noFill = True
chart.graphical_properties.line.prstDash = None
```

**Znaczniki (markery) na wykresach liniowych.** Domyślnie w `LineChart` `series.marker` jest `None` — czyli linia bez kropek. Żeby je dodać:

```python
from openpyxl.chart.marker import Marker

series = chart.series[0]
series.marker = Marker(symbol="circle", size=7)
```

Dozwolone `symbol` (dokładna lista ze źródła 3.1.x): `circle`, `dash`, `diamond`, `dot`, `plus`, `square`, `star`, `triangle`, `x`, `auto`, `picture`.

> Uwaga praktyczna: na liście **nie ma** wartości `'none'`. Jeśli chcesz pozbyć się znaczników, po prostu **nie ustawiaj** `marker` (zostaw `None`). Jeśli chcesz ukryć **linię**, a zostawić kropki (jak w wykresach giełdowych), użyj `series.graphicalProperties.line.noFill = True` — to wzorzec z oficjalnego przykładu wykresu giełdowego.

**Kolor pojedynczego punktu** (np. jeden wyspecjalizowany słupek, albo plasterek koła):

```python
from openpyxl.chart.series import DataPoint

punkt = DataPoint(idx=0)
punkt.graphicalProperties.solidFill = "FF0000"
chart.series[0].data_points = [punkt]
```

`data_points` to alias `Series.dPt`. `idx` liczy się **od zera** (pierwszy punkt = 0). Ten mechanizm jest podstawą wykresu „gauge" z dokumentacji i wyciąganego plasterka koła (`DataPoint(idx=0, explosion=20)`).

**Styl presety.** Najprostszy sposób na „ładny wykres bez pracy":

```python
chart.style = 10      # liczba 1-48
```

To paleta kolorów z Excela. Jeśli nie masz wymagań brandingowych, `chart.style` daje przyzwoity efekt w jednej linii. Gdy masz wymagania — użyj `graphicalProperties`.

### 2.11. Rozmiar, pozycja i kotwica

```python
chart.width = 18        # cm
chart.height = 9        # cm
ws.add_chart(chart, "E2")   # pinezka w lewym gornym rogu = E2
```

Możesz też ustawić kotwicę **przez obiekt**, jeśli potrzebujesz precyzji (albo wykresu zakotwiczonego do dwóch komórek, żeby rósł razem z nimi):

```python
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import cm_to_EMU

wysokosc = XDRPositiveSize2D(cx=cm_to_EMU(18), cy=cm_to_EMU(9))
marker = AnchorMarker(col=4, row=1)     # 0-based!
chart.anchor = OneCellAnchor(_from=marker, ext=wysokosc)
ws.add_chart(chart)
```

> **Uwaga:** to jest zaawansowany kawałek API i **indeksowanie w `AnchorMarker` liczy się od zera** (`col=4` to kolumna E!). To pułapka odwrotna do całej reszty openpyxl (moduł 03: „notatnik — indeksowanie od 1"). Jeśli nie potrzebujesz precyzji co do piksela — **używaj prostego `ws.add_chart(chart, "E2")`** i ustaw `width`/`height`. To pokrywa 99% potrzeb.

**Planowanie rozmieszczenia (bardzo praktyczne).** Przy jednym arkuszu z danymi i wykresami trzymaj się prostej konwencji:

```python
# dane w A1:D100
# wykres 1 w F2 (18 x 9 cm)
# wykres 2 w F20
# wykres 3 w F38
```

Zasada: **wykresy po prawej stronie danych, nigdy na nich.** Odstępy co ~18 wierszy dla wykresu 9 cm wysokiego (wiersz ≈ 0,5 cm przy domyślnej wysokości).

### 2.12. Trzy kategorie prawdy o wykresach w openpyxl

Czas na uczciwe rozliczenie. To najważniejsza tabela w module.

**Potrafi (w pełni kontrolowane przy generowaniu od zera):**

- Wszystkie klasyczne typy wykresów: słupkowe, liniowe, obszarowe, kołowe, pierścieniowe, punktowe, bąbelkowe, radarowe, giełdowe, powierzchniowe (2D i 3D).
- Serie, tytuły serii z komórek, kategorie, etykiety danych, legenda, wszystkie rodzaje osi.
- Skalowanie osi (min/max/logarytmiczna/odwrócona), formaty liczb na osiach (moduł 10).
- Kolory serii i punktów, znaczniki, gradienty, przezroczystość, obramowania.
- Wykresy kombinowane i wykresy z drugą osią.
- Wykresy na osobnym arkuszu wykresów (`Chartsheet`, moduł 08).
- Wykresy w trybie `write_only` (potwierdzone oficjalnym przykładem).

**Potrafi tylko częściowo:**

- **Odtworzenie wykresu wczytanego z Excela.** openpyxl wczyta wykres i zapisze go z powrotem, ale modele OOXML są ogromne — część formatowania może zostać uproszczona. Traktuj „przeszło" jako „przeszło, ale sprawdź okiem".
- **Paski błędów, linie trendu.** Istnieją (`ErrorBars`, `Trendline`), ale mają swoje słabości. W changelogu 3.1.3 znajdziesz np. `#2106 Setting Trendline.name attribute raises exception when saving` — błąd naprawiony, ale pokazujący, że te zakątki API są rzadziej używane i testowane.
- **Wykresy giełdowe z liniami hi-lo.** Dokumentacja wprost mówi: *„Due to a bug in Excel high-low lines will only be shown if at least one of the data series has some dummy values"*. Trzeba **ręcznie wstrzyknąć atrapę pamięci podręcznej**:

  ```python
  from openpyxl.chart.data_source import NumData, NumVal
  pts = [NumVal(idx=i) for i in range(len(data) - 1)]
  cache = NumData(pt=pts)
  c1.series[-1].val.numRef.numCache = cache
  ```
  
  To jest wzorcowy przykład kategorii „częściowo": funkcja działa, ale wymaga hacku, którego **nie da się wymyślić** bez przeczytania dokumentacji.
- **Odczyt niektórych typów wykresów punktowych.** `#1360 Can't read some ScatterCharts if the x-axis is not numerical` — historyczny błąd, naprawiony, ale pokazuje, że wykres rozproszenia bywa problematyczny.

**Gubi przy zapisie (lub nie zna wcale):**

- **Sparkline'y** — mini-wykresy w komórce (`x14` extensions). openpyxl **nie ma dla nich API**. Przy wczytaniu pliku ostrzega, że je usunie, i przy zapisie **gubi je bezpowrotnie**. To udokumentowane i niezmienne ograniczenie.
- **Nowe typy wykresów z rodziny „chartex"** (Excel 2016+): mapa (Filled Map), treemap, sunburst, histogram, Pareto, waterfall (kaskadowy), funnel (lejek), pudełko-wąsy. openpyxl implementuje **tylko klasyczny `chartSpace`** — te typy żyją w innym namespace i openpyxl ich nie modeluje.
- **Wykresy przestawne (pivot charts).** openpyxl nie tworzy wykresów opartych na tabelach przestawnych. Co więcej, związek wykresu z tabelą przestawną bywał gubiony (`#1232 Link to pivot table lost from charts` — historyczny błąd, naprawiany).
- **Niektóre elementy dekoracyjne i nowe warianty formatowania** — np. niestandardowe wypełnienia, nowe typy pasków (moduł 14), efekty 3D.
- **Problemy z kolejnością serii w wykresach kombinowanych.** W kodzie źródłowym `Series.to_tree(tagname=None, idx=None)` znajdziesz komentarz: *„The index can need rebasing"*. Serie mają `idx` i `order`, które muszą być spójne i unikalne. Przy skomplikowanych kombinacjach (serie z kilku wykresów, wykresy odczytane z pliku i połączone) może się zdarzyć, że Excel zgłosi „uszkodzony plik" albo pokaże serie w złej kolejności. **Jeśli tak się stanie: zbuduj wykres od zera, w jednej kolejności, zamiast sklejać wczytane.**

**Czego nie ma i nie będzie:**

- Silnika renderowania (nie zobaczysz obrazka wykresu bez otwarcia pliku w Excelu/LibreOffice — moduł 00),
- Wykresów osadzonych jako **obrazy** (do tego jest `matplotlib` + `add_image`, moduł 16),
- Zaliczenia formuł w wykresie (moduł 06 — wykres pokaże pustkę, jeśli w komórkach są formuły bez pamięci podręcznej!).

> **⚠️ Najważniejsza konsekwencja praktyczna tego ostatniego punktu.** Jeśli zapisujesz w komórkach **formuły** (moduł 06) i budujesz na nich wykres, a plik nigdy nie był otwarty w Excelu, wykres nie ma żadnych danych do pokazania — bo w pliku nie ma **wartości z pamięci podręcznej**. Excel po otwarciu przeliczy formuły i wykres się pojawi, ale w pliku „w locie" dane są puste. **Zasada: jeśli budujesz wykres z danych generowanych przez openpyxl, wpisuj wartości, a nie formuły.** To kolejny argument za podziałem z modułu 06 („Excel to warstwa prezentacji, nie silnik biznesowy").

### 2.13. Wykresy a tryby pracy i arkusze wykresów

**Tryb normalny** — wszystko działa. To standard.

**Tryb `read_only` — wykresy są NIEDOSTĘPNE.** Potwierdzone w źródle (`openpyxl/reader/excel.py`): dla arkusza czytanego leniwie parser robi `continue` **przed** blokiem przetwarzającym rysunki, więc ani wykresy, ani obrazy nie trafiają do modelu. Czytając plik w `read_only`, zobaczysz dane, ale nie zobaczysz wykresów. To nie błąd — to konsekwencja kompromisu (pamięć vs kompletność, moduł 03).

**Tryb `write_only` — wykresy DZIAŁAJĄ.** To zaskakuje ludzi, bo brzmi jak sprzeczność: „skoro nie można wracać do komórek, jak zbudować wykres?". Odpowiedź: wykres **nie potrzebuje** wracania do komórek — potrzebuje **adresu**. Oficjalny przykład dokumentacji (`bar.html`) buduje wykres właśnie w `Workbook(write_only=True)`:

```python
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

wb = Workbook(write_only=True)
ws = wb.create_sheet()

for row in rows:
    ws.append(row)

chart = BarChart()
chart.type = "col"
data = Reference(ws, min_col=2, min_row=1, max_row=7, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=7)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "A10")       # ← dziala w write_only!

wb.save("bar.xlsx")
```

**Dlaczego to działa?** Bo `add_chart` dopisuje jedynie **definicję** wykresu (adresy komórek), a nie odczytuje ich wartości. A że w trybie strumieniowym komórki już zostały zapisane na dysk (w kolejności), adresy są poprawne. Dokumentacja `optimized.html` przypomina jednak: *„Everything that appears in the file before the actual cell data must be created before cells are added"* — czyli **wykres dodaj PO danych**, nie przed. To nie problem, bo w `write_only` i tak nie mógłbyś nic dopisać do już zapisanych komórek.

**Arkusz wykresów (`Chartsheet`).** Zamiast wstawiać wykres „nad" danymi, możesz poświęcić mu cały arkusz:

```python
cs = wb.create_chartsheet(title="Wykres sprzedazy")
cs.add_chart(chart)
```

To ma sens w raportach dla zarządu (jeden duży wykres na ekran) i było tematem modułu 08. Rzadko używane w automatyzacji.

**Odczyt wykresów z istniejącego pliku.** Nie ma publicznego API do listowania wykresów — używa się atrybutu prywatnego `ws._charts`:

```python
for chart in ws._charts:          # UWAGA: atrybut prywatny (podkreslnik)
    print(type(chart).__name__, len(chart.series))
```

Podkreślnik oznacza „API wewnętrzne, może się zmienić". Ale do diagnostyki to niezastąpione narzędzie — pokaże ci, ile wykresów openpyxl faktycznie widzi. Jeśli w Excelu są trzy, a tu jedno — masz wykresy, których openpyxl nie modeluje.

### 2.14. Wzorzec: jedna funkcja, jeden wykres

Brief wspomina o tym na końcu, ale to najważniejsza rzecz **projektowa** w module — zapowiedź modułu 24 (Facade).

Popatrz na kod z 2.6. Widzisz problem? Budowa każdego wykresu to 6–10 linii rozrzuconych po skrypcie: `Reference`, `add_data`, `set_categories`, tytuły, osie, `add_chart`. Przy pięciu wykresach w raporcie robi się 50 linii, w których **nie widać, co jest wspólne** (styl, rozmiar, brak legendy), a co specyficzne (zakres, typ).

Rozwiązanie: **jedna funkcja = jeden wykres**. Funkcja przyjmuje opis („jaki typ, jaki zakres, jaki tytuł, gdzie postawić") i zwraca gotowy wykres dodany do arkusza:

```python
def dodaj_wykres_sprzedazy(ws, zakres_danych, zakres_kategorii, tytul, anchor):
    """Buduje i dodaje wykres slupkowy sprzedazy. Jeden wykres = jedna funkcja."""
    chart = BarChart()
    chart.type = "col"
    chart.title = tytul
    chart.style = 10
    chart.legend = None
    chart.width = 18
    chart.height = 9
    chart.y_axis.title = "Sprzedaż (tys. zł)"
    chart.x_axis.title = "Miesiąc"

    chart.add_data(zakres_danych, titles_from_data=True)   # NAJPIERW dane
    chart.set_categories(zakres_kategorii)                 # POTEM kategorie

    ws.add_chart(chart, anchor)
    return chart
```

Zysk jest potrójny:

1. **Wspólny styl w jednym miejscu.** Zmiana rozmiaru wszystkich wykresów to edycja jednej liczby. To ta sama motywacja co rejestr stylów w module 09 i generatory reguł w module 14.
2. **Testowalność.** Możesz wywołać funkcję i sprawdzić, czy wykres powstał i ile ma serii (moduł 20).
3. **Czytelny kod główny.** `dodaj_wykres_sprzedazy(ws, dane, kategorie, "Sprzedaż 2026", "F2")` — mówisz **co**, nie **jak**.

W module 24 zobaczysz, jak z tej funkcji wyrasta **Facade** — klasa `RaportExcel`, której publiczne metody nie przyjmują żadnego typu openpyxl w sygnaturach. Ale już tutaj warto zacząć od tej prostej dyscypliny: **wykres nie powstaje „w locie", on ma swoją funkcję.**

## 3. Przykłady krok po kroku

### Przykład 1 — pierwszy wykres słupkowy i odczyt z powrotem (🟢)

Pierwszy przykład pokazuje **kompletny cykl**: utwórz dane → zbuduj wykres → zapisz → wczytaj i sprawdź, co openpyxl widzi.

```python
"""Modul 15 - pierwszy wykres slupkowy + odczyt z powrotem.

Uruchom:  python examples/15_bar_podstawy.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "15_bar_podstawy.xlsx"

MIESIACE = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze",
            "Lip", "Sie", "Wrz", "Paz", "Lis", "Gru"]
SPRZEDAZ = [120, 135, 98, 150, 172, 160, 143, 158, 181, 190, 210, 240]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaz"

    # --- dane: naglowek + 12 wierszy -> wiersze 1..13 ----------------
    ws.append(["Miesiąc", "Sprzedaż (tys. zł)"])
    for miesiac, wartosc in zip(MIESIACE, SPRZEDAZ):
        ws.append([miesiac, wartosc])

    # --- wykres ------------------------------------------------------
    chart = BarChart()
    chart.type = "col"                       # "col" = slupki pionowe
    chart.grouping = "clustered"             # domyslne, ale jawnie
    chart.gapWidth = 80                      # odstepy miedzy slupkami (domyslnie 150)
    chart.title = "Sprzedaż 2026"
    chart.x_axis.title = "Miesiąc"
    chart.y_axis.title = "Sprzedaż (tys. zł)"
    chart.style = 10                         # preset kolorow z Excela

    # data: KOLUMNA B, z nagłówkiem (wiersz 1) -> tytul serii
    data = Reference(ws, min_col=2, min_row=1, max_row=13)
    # kategorie: kolumna A, BEZ nagłówka (wiersz 2) -> to sa miesiace
    kategorie = Reference(ws, min_col=1, min_row=2, max_row=13)

    chart.add_data(data, titles_from_data=True)     # 1) NAJPIERW dane
    chart.set_categories(kategorie)                 # 2) POTEM kategorie

    # etykiety danych: pokaz wartosc nad slupkiem
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.numFmt = "#,##0"

    # jedna seria -> legenda zbedna
    chart.legend = None

    # rozmiar (cm) i pozycja (pinezka w lewym gornym rogu)
    chart.width = 18
    chart.height = 9
    ws.add_chart(chart, "D2")

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    """Wczytuje plik i sprawdza, CO NAPRAWDE trafilo do modelu."""
    wb = load_workbook(cel)
    try:
        ws = wb["Sprzedaz"]

        # UWAGA: nie ma publicznego API do listowania wykresow -> _charts
        print("=" * 78)
        print(f"Wykresow w arkuszu '{ws.title}': {len(ws._charts)}")
        print("=" * 78)

        for i, chart in enumerate(ws._charts):
            print(f"[{i}] typ obiektu : {type(chart).__name__}")
            print(f"    tytul        : {chart.title}")
            print(f"    styl         : {chart.style}")
            print(f"    szerokosc    : {chart.width} x {chart.height} (cm)")
            print(f"    liczba serii : {len(chart.series)}")

            for j, seria in enumerate(chart.series):
                # val -> wartosci (adres!), cat -> kategorie (adres!)
                adres_val = seria.val.numRef.f
                adres_cat = seria.cat.numRef.f if seria.cat is not None else None
                tytul_serii = seria.tx.strRef.f if seria.tx is not None else None
                print(f"    seria {j}:")
                print(f"        tytul  (adres): {tytul_serii}")
                print(f"        wartosci(adres): {adres_val}")
                print(f"        kategorie(adres): {adres_cat}")
            print()
    finally:
        wb.close()

    print("=" * 78)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 78)
    print("  1. Kliknij na wykres -> pasek formuly pokaze zakresy komorek.")
    print("     To dowod, ze wykres trzyma ADRES, a nie kopie danych!")
    print("  2. Zmien wartosc w B3 (Sty) z 120 na 500, zapisz.")
    print("     Slupek natychmiast urosnie - bo wykres wskazuje na B3.")
    print("  3. Kliknij na slupek i nacisnij Delete -> seria zniknie.")
    print("     Zapisz i wczytaj openpyxl: seria zniknela z modelu (nie ma jej w XML).")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Po `chart.add_data(data, ...)` openpyxl tworzy jeden obiekt `Series`, w którym `val` (wartości) wskazuje `Reference` na kolumnę B, a `tx` (tytuł serii) — na komórkę `B1`. Po `set_categories` do tej serii dopisuje się `cat` wskazujący kolumnę A (od wiersza 2). **W pamięci nie ma ani jednej liczby ze sprzedaży** — są tylko adresy.

**Co trafi do pliku.** Nowa część `xl/charts/chart1.xml` z definicją wykresu, nowa część `xl/drawings/drawing1.xml` z kotwicą `D2`, oraz **dwie nowe relacje** wiążące arkusz z rysunkiem i rysunek z wykresem (`_rels`). W definicji wykresu zobaczysz dokładnie to, o czym mówiłem w 2.1:

```xml
<c:ser>
  <c:idx val="0"/>
  <c:order val="0"/>
  <c:tx><c:strRef><c:f>'Sprzedaz'!$B$1</c:f></c:strRef></c:tx>
  <c:cat><c:numRef><c:f>'Sprzedaz'!$A$2:$A$13</c:f></c:numRef></c:cat>
  <c:val><c:numRef><c:f>'Sprzedaz'!$B$2:$B$13</c:f></c:numRef></c:val>
</c:ser>
<c:axId val="10"/>   <!-- os kategorii (TextAxis) -->
<c:axId val="100"/>  <!-- os wartosci (NumericAxis) -->
```

**Trzy rzeczy do przemyślenia:**

1. **`<c:f>` to adres, nie wartość.** Jeśli do tej pory myślałeś o wykresach jak o obrazkach, ten jeden wiersz XML zmienia wszystko. Wykres to **wskazanie palcem** — jak mówiłem w 1.2.
2. **W `chart.dataLabels.numFmt` działa cały moduł 10.** `#,##0` sprawi, że Excel pokaże `120`, a nie `120.0`. Format liczb jest wszędzie taki sam: komórki, osie, etykiety danych.
3. **`chart.gapWidth = 80`** — domyślna wartość to `150`, czyli szerokie odstępy. Zmniejszenie daje „gęstszy" wykres, w którym słupki dominują. Drobiazg, który bardzo zmienia wrażenie estetyczne.

### Przykład 2 — dwie osie: sprzedaż (słupki) i marża (linia) (🟡)

To zadanie z briefu, doprowadzone do końca. Zobaczysz pełną mechanikę z 2.8.

```python
"""Modul 15 - wykres z DWIEMA osiami: sprzedaz (slupki) + marza % (linia).

Kluczowe elementy:
  - c2.y_axis.axId = 200    -> osobny identyfikator osi wartosci
  - c1.y_axis.crosses = "max" -> gdzie os ma przeciac os X
  - c1 += c2                -> sklejenie dwoch wykresow w jeden
  - reczne wylaczenie siatki na jednej z os (dwie siatki = chaos)

Uruchom:  python examples/15_dwie_osie.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "15_dwie_osie.xlsx"

MIESIACE = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze",
            "Lip", "Sie", "Wrz", "Paz", "Lis", "Gru"]
# (sprzedaz w tys. zl, marza jako ulamek)
DANE = [
    (120, 0.18), (135, 0.21), (98, 0.09), (150, 0.24),
    (172, 0.27), (160, 0.22), (143, 0.19), (158, 0.25),
    (181, 0.29), (190, 0.31), (210, 0.28), (240, 0.34),
]


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaz"

    ws.append(["Miesiąc", "Sprzedaż (tys. zł)", "Marża"])
    for miesiac, (sprzedaz, marza) in zip(MIESIACE, DANE):
        ws.append([miesiac, sprzedaz, marza])

    # kolumna C jako procent (modul 10!) - w komorkach i na etykietach
    for wiersz in range(2, 14):
        ws.cell(row=wiersz, column=3).number_format = "0.0%"

    # =============== WYKRES 1: SLUPKI (sprzedaz) ======================
    c1 = BarChart()
    c1.type = "col"
    c1.title = "Sprzedaż i marża 2026"
    c1.style = 10
    c1.x_axis.title = "Miesiąc"
    c1.y_axis.title = "Sprzedaż (tys. zł)"

    # NumericAxis ma siatke DOMYSLNIE WLACZONA. Dwie osie = dwie siatki
    # = wizualny chaos. Wylaczamy siatke na osi sprzedazy.
    c1.y_axis.majorGridlines = None

    dane_sprzedaz = Reference(ws, min_col=2, min_row=1, max_row=13)
    kategorie = Reference(ws, min_col=1, min_row=2, max_row=13)
    c1.add_data(dane_sprzedaz, titles_from_data=True)
    c1.set_categories(kategorie)

    # =============== WYKRES 2: LINIA (marza %) ========================
    c2 = LineChart()
    c2.y_axis.axId = 200               # NOWA os wartosci (domyslna to 100)
    c2.y_axis.title = "Marża (%)"
    c2.y_axis.numFmt = "0%"            # format etykiet osi (modul 10)

    dane_marza = Reference(ws, min_col=3, min_row=1, max_row=13)
    c2.add_data(dane_marza, titles_from_data=True)
    c2.set_categories(kategorie)
    c2.y_axis.majorGridlines = None    # nie chcemy drugiej siatki

    # Os ma przeciac os X na maksimum -> pokaze sie po prawej stronie.
    # UWAGA: dokumentacja ustawia to na y_axis PIERWSZEGO wykresu.
    c1.y_axis.crosses = "max"

    # =============== SKLEJENIE ========================================
    c1 += c2                           # c1 ma tytul/styl, c2 ma swoja os
    c1.width = 22
    c1.height = 10
    ws.add_chart(c1, "E2")

    # --- DRUGI wykres: pokazujemy, ze dwie osie mozna takze ZEPSUC ----
    c3 = BarChart()
    c3.type = "col"
    c3.title = "ŹLE: bez drugiej osi (marża płaska przy dnie)"
    c3.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=13),
                titles_from_data=True)
    c3.set_categories(kategorie)
    c3.width = 22
    c3.height = 10
    ws.add_chart(c3, "E24")

    wb.save(cel)
    wb.close()


def weryfikuj(cel: Path) -> None:
    wb = load_workbook(cel)
    try:
        ws = wb["Sprzedaz"]
        print("=" * 84)
        print(f"Wykresow: {len(ws._charts)}")
        print("=" * 84)
        for i, chart in enumerate(ws._charts):
            print(f"[{i}] {type(chart).__name__}, seria={len(chart.series)}, "
                  f"tytul={chart.title}")
            # osie: NumericAxis ma atrybut 'crosses', TextAxis tez
            os_y = getattr(chart, "y_axis", None)
            if os_y is not None:
                print(f"     os Y: axId={os_y.axId}, crosses={os_y.crosses}, "
                      f"numFmt={os_y.numFmt}")
            # drugi wykres w kombinacji siedzi w chart._charts (plotArea)
            dodatkowe = getattr(chart, "_charts", None)
            if dodatkowe:
                for k, sub in enumerate(dodatkowe):
                    print(f"     + sub-wykres {k}: {type(sub).__name__}, "
                          f"seria={len(sub.series)}, "
                          f"axId_y={getattr(sub.y_axis, 'axId', '?')}")
    finally:
        wb.close()

    print()
    print("=" * 84)
    print("CO SPRAWDZIC W EXCELU")
    print("=" * 84)
    print("  1. Pierwszy wykres: slupki = sprzedaz (os LEWA, tys. zl),")
    print("     linia = marza (os PRAWA, procenty). Dwie rozne skale!")
    print("  2. Kliknij linie -> PPM -> Formatuj serie danych -> os: WTÓRNA.")
    print("     To potwierdza, ze Excel widzi dwie osobne osie wartosci.")
    print("  3. Drugi wykres (zly) - marza to plaska linia przy dnie,")
    print("     bo 0,18-0,34 przy sprzedazy 98-240 to wizualnie zero.")
    print("     TO JEST WLASNIE POWOD, dla ktorego istnieje druga os.")
    print("  4. EKSPERYMENT: usun ustawienie c1.y_axis.crosses = 'max'")
    print("     i zapisz ponownie - zobacz, po ktorej stronie pojawi sie os.")


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}\n")
    weryfikuj(CEL)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `c2` to pełnoprawny wykres liniowy, ale z **innym** `axId` na osi wartości (200 zamiast 100). Po `c1 += c2` openpyxl tworzy nowy plotArea zawierający **oba** wykresy i **cztery** osie: dwie wspólne osie X (`axId=10`) i dwie osie Y (`axId=100` i `axId=200`).

**Co trafi do pliku.** `xl/charts/chart1.xml` z `<c:plotArea>` zawierającym `<c:barChart>` **i** `<c:lineChart>` oraz czterema elementami `<c:axId>`/`<c:catAx>`/`<c:valAx>`. Powiązania osi wyraża atrybut `crossAx`: oś Y z `axId=200` deklaruje, że przecina oś X (`axId=10`).

**Trzy rzeczy do przemyślenia:**

1. **Druga oś to drugi wykres.** To najczęściej myląca rzecz w tym module. Nie ma „metody `add_second_axis()`". Jest drugi obiekt wykresu, nowy `axId`, i `+=`. Gdy to zrozumiesz, `crosses` przestanie być magią.
2. **Siatka: dwie osie = dwie siatki.** Wykres z dwiema włączonymi siatkami wygląda jak padający deszcz. Dlatego w kodzie `majorGridlines = None` na obu osiach wartości. To bezpośrednia konsekwencja odkrycia z 2.7 (siatka jest domyślnie włączona).
3. **Drugi wykres w tym przykładzie jest celowo „zły".** Zobacz, jak marża wygląda na jednej skali ze sprzedażą: płaska kreska przy dnie. **Zawsze obejrzyj oba warianty** — to najlepszy sposób, żeby zrozumieć, po co druga oś istnieje, i żeby umieć uzasadnić jej użycie w raporcie.

### Przykład 3 — katalog typów wykresów (🟡)

Ten przykład to wzornik: sześć typów wykresów na sześciu arkuszach. Uruchom raz, obejrzyj wszystkie, wróć do niego jak do ściągawki — zamiast szukać po internecie.

```python
"""Modul 15 - katalog typow wykresow. Kazdy typ na osobnym arkuszu.

Uruchom:  python examples/15_katalog_typow.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    BarChart,
    BubbleChart,
    DoughnutChart,
    LineChart,
    PieChart,
    RadarChart,
    Reference,
    ScatterChart,
    Series,
)

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "15_katalog_typow.xlsx"

MIESIACE = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze"]
BATCH1 = [40, 40, 50, 30, 25, 50]
BATCH2 = [30, 25, 30, 10, 5, 10]


def arkusz_danych(wb: Workbook, tytul: str):
    """Tworzy arkusz z danymi 'Miesiac/Batch 1/Batch 2'."""
    ws = wb.create_sheet(title=tytul)
    ws.append(["Miesiąc", "Batch 1", "Batch 2"])
    for m, b1, b2 in zip(MIESIACE, BATCH1, BATCH2):
        ws.append([m, b1, b2])
    return ws


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)   # usuwamy domyslny arkusz

    # ---------- 1. SLUPKOWY ------------------------------------------
    ws = arkusz_danych(wb, "Slupkowy")
    chart = BarChart()
    chart.type = "col"
    chart.title = "Słupkowy (kolumny)"
    chart.y_axis.title = "Wartość"
    chart.x_axis.title = "Miesiąc"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 2. SLUPKOWY WARSTWOWY --------------------------------
    ws = arkusz_danych(wb, "Warstwowy")
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100        # WYMAGANE dla stacked (dokumentacja!)
    chart.title = "Warstwowy (stacked)"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 3. LINIOWY ------------------------------------------
    ws = arkusz_danych(wb, "Liniowy")
    chart = LineChart()
    chart.title = "Liniowy"
    chart.y_axis.title = "Wartość"
    chart.x_axis.title = "Miesiąc"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 4. PIERSCIENIOWY ------------------------------------
    ws = wb.create_sheet(title="Pierscien")
    ws.append(["Kategoria", "Udział"])
    for kat, wart in [("Kawa", 40), ("Herbata", 25), ("Woda", 20), ("Sok", 15)]:
        ws.append([kat, wart])
    chart = DoughnutChart(firstSliceAng=270, holeSize=50)
    chart.title = "Pierścieniowy"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=5),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "E2")

    # ---------- 5. PUNKTOWY (scatter) -------------------------------
    # UWAGA: tu NIE uzywamy add_data - serie budujemy recznie!
    ws = wb.create_sheet(title="Punktowy")
    ws.append(["Rozmiar", "Wynik"])
    for x, y in [(2, 40), (3, 40), (4, 50), (5, 30), (6, 25), (7, 20)]:
        ws.append([x, y])
    chart = ScatterChart()
    chart.title = "Punktowy (scatter)"
    chart.scatterStyle = "marker"
    chart.x_axis.title = "Rozmiar"
    chart.y_axis.title = "Wynik"
    xvals = Reference(ws, min_col=1, min_row=2, max_row=7)
    yvals = Reference(ws, min_col=2, min_row=1, max_row=7)   # z naglowkiem!
    chart.series.append(Series(yvals, xvals, title_from_data=True))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 6. BABELKOWY ----------------------------------------
    ws = wb.create_sheet(title="Babelkowy")
    ws.append(["X", "Y", "Rozmiar"])
    for x, y, z in [(14, 12200, 15), (20, 60000, 33), (18, 24400, 10),
                    (22, 32000, 42), (12, 8200, 18)]:
        ws.append([x, y, z])
    chart = BubbleChart()
    chart.style = 18
    chart.title = "Bąbelkowy"
    xvals = Reference(ws, min_col=1, min_row=2, max_row=6)
    yvals = Reference(ws, min_col=2, min_row=2, max_row=6)
    zvals = Reference(ws, min_col=3, min_row=2, max_row=6)
    chart.series.append(
        Series(values=yvals, xvalues=xvals, zvalues=zvals, title="2026")
    )
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 7. RADAROWY -----------------------------------------
    ws = arkusz_danych(wb, "Radarowy")
    chart = RadarChart()
    chart.type = "marker"
    chart.title = "Radarowy"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 8. OBSZAROWY ----------------------------------------
    ws = arkusz_danych(wb, "Obszarowy")
    chart = AreaChart()
    chart.title = "Obszarowy"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "F2")

    # ---------- 9. KOLOWY -------------------------------------------
    ws = wb.create_sheet(title="Kolowy")
    ws.append(["Produkt", "Sprzedaż"])
    for prod, wart in [("Jabłka", 50), ("Wiśnie", 30), ("Dynia", 10), ("Czeko", 40)]:
        ws.append([prod, wart])
    chart = PieChart()
    chart.title = "Kołowy (JEDNA seria!)"
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=5),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))
    # etykiety: kategoria + procent (klasyka dla kolowego)
    from openpyxl.chart.label import DataLabelList
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showCatName = True
    chart.dataLabels.showPercent = True
    chart.width = 16
    chart.height = 8
    ws.add_chart(chart, "E2")

    wb.save(cel)
    wb.close()


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}")
    print("Otwórz plik i obejrzyj wszystkie arkusze.")
    print("Zauważ: warstwowy wymaga overlap=100; kolowy ma JEDNA serie;")
    print("punktowy i babelkowy buduja serie recznie (bez add_data).")


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **Każdy typ ma swoje „widzimisię".** Stacked wymaga `overlap = 100` (inaczej Excel rysuje je obok siebie). Kołowy wymaga **jednej** serii. Scatter i bubble budują serie ręcznie, **bez** `add_data`. To nie są wyjątki od reguły — to cechy typów. **Kataloguj je, zamiast za każdym razem odkrywać od nowa.**
2. **`ScatterChart` a `LineChart` — to różne narzędzia.** Scatter rysuje **Y przeciwko X** (dwie zmienne numeryczne). Line rysuje **Y w kolejności kategorii**. Jeśli użyjesz `LineChart` dla danych „temperatura od czasu", gdzie pomiary są nieregularne, Excel potraktuje odstępy jako równe — i zafałszuje obraz. To pułapka o charakterze **merytorycznym**, nie technicznym.
3. **Katalog jest twoim narzędziem pracy.** W praktycznym raporcie użyjesz 3–4 typów. Ale gdy pojawi się „wykres bąbelkowy dla klienta", nie będziesz szukać po internecie — otworzysz ten plik i skopiujesz sprawdzony fragment.

### Przykład 4 — stylowanie: kolory, znaczniki, osie, tło (🟡)

Ten przykład pokazuje wszystkie techniki z 2.10 i 2.7 w jednym miejscu.

```python
"""Modul 15 - stylowanie wykresow: kolory serii, znaczniki, osie, tlo.

Uruchom:  python examples/15_stylowanie.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)
CEL = OUTPUT / "15_stylowanie.xlsx"

MIESIACE = ["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze"]
ALFA = [120, 135, 98, 150, 172, 160]
BETA = [80, 95, 130, 110, 140, 155]

# Paleta marki - JEDNO miejsce na kolory (wzorzec z modulu 09!)
PALETA = {
    "alfa": "5B5CF0",
    "beta": "F0705B",
    "siatka": "D9D9D9",
}


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    ws.append(["Miesiąc", "Alfa", "Beta"])
    for m, a, b in zip(MIESIACE, ALFA, BETA):
        ws.append([m, a, b])

    # =================================================================
    # WYKRES 1: SLUPKI z kolorami serii, etykietami i skalowana osia
    # =================================================================
    c = BarChart()
    c.type = "col"
    c.title = "Porównanie Alfa vs Beta"
    c.style = None                      # nie chcemy presetu - mamy wlasne kolory
    c.x_axis.title = "Miesiąc"
    c.y_axis.title = "Sprzedaż (tys. zł)"

    c.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
               titles_from_data=True)
    c.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))

    # kolory serii - tu, a nie w "color" (sekcja 2.10)
    c.series[0].graphicalProperties = GraphicalProperties(solidFill=PALETA["alfa"])
    c.series[1].graphicalProperties = GraphicalProperties(solidFill=PALETA["beta"])

    # etykiety danych tylko nad slupkami, bez legendy
    c.dataLabels = DataLabelList()
    c.dataLabels.showVal = True
    c.dataLabels.numFmt = "#,##0"
    c.legend.position = "b"

    # skala osi: wymuszamy zakres, zeby oba wykresy byly porownywalne
    c.y_axis.scaling.min = 0
    c.y_axis.scaling.max = 200
    c.y_axis.numFmt = '#,##0'

    c.width = 18
    c.height = 9
    ws.add_chart(c, "F2")

    # =================================================================
    # WYKRES 2: LINIOWY ze znacznikami, gruba kreska, przezroczyste tlo
    # =================================================================
    c2 = LineChart()
    c2.title = "Trend Alfa i Beta"
    c2.x_axis.title = "Miesiąc"
    c2.y_axis.title = "Sprzedaż (tys. zł)"

    c2.add_data(Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7),
                titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))

    # znaczniki: LineChart ma je DOMYSLNIE jako None -> trzeba utworzyc
    for seria, kolor in zip(c2.series, (PALETA["alfa"], PALETA["beta"])):
        seria.marker = Marker(symbol="circle", size=8)
        seria.marker.graphicalProperties = GraphicalProperties(solidFill=kolor)
        seria.graphicalProperties = GraphicalProperties()
        seria.graphicalProperties.line.solidFill = kolor
        seria.graphicalProperties.line.width = 30000   # EMU (~2.4 pt)

    # przezroczyste tlo wykresu + brak obramowania
    c2.graphical_properties = GraphicalProperties()
    c2.graphical_properties.noFill = True
    c2.graphical_properties.line.noFill = True
    c2.graphical_properties.line.prstDash = None

    # siatka pomocnicza tylko na osi Y, w kolorze marki
    c2.y_axis.majorGridlines = None    # najpierw usuwamy domyslna
    c2.x_axis.delete = False           # os X zostaje widoczna

    c2.width = 18
    c2.height = 9
    ws.add_chart(c2, "F22")

    # =================================================================
    # WYKRES 3: UKRYTE OSIE (minimalistyczny "sparkline w duzym formacie")
    # =================================================================
    c3 = BarChart()
    c3.type = "col"
    c3.title = "Bez osi (minimalistycznie)"
    c3.add_data(Reference(ws, min_col=2, min_row=1, max_row=7),
                titles_from_data=True)
    c3.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7))
    c3.legend = None
    c3.x_axis.delete = True            # ukryj os X
    c3.y_axis.delete = True            # ukryj os Y
    c3.y_axis.majorGridlines = None
    c3.series[0].graphicalProperties = GraphicalProperties(solidFill=PALETA["alfa"])
    c3.dataLabels = DataLabelList()
    c3.dataLabels.showVal = True
    c3.width = 18
    c3.height = 9
    ws.add_chart(c3, "F42")

    wb.save(cel)
    wb.close()


def main() -> None:
    zbuduj(CEL)
    print(f"Zapisano: {CEL}")
    print()
    print("CO SPRAWDZIC W EXCELU")
    print("  Wykres 1: dwa kolory serii (fioletowy/pomaranczowy), wartosci nad")
    print("            slupkami, legenda na dole, os Y od 0 do 200.")
    print("  Wykres 2: linie z kolowymi znacznikami, grube kreski,")
    print("            PRZEZROCZYSTE tlo (widac linie siatki arkusza pod spodem).")
    print("  Wykres 3: brak osi i legendy - same slupki z wartosciami.")
    print()
    print("EKSPERYMENT: usun linie 'c.style = None' i ustaw c.style = 10.")
    print("  Zobaczysz, ze styl presetu NADPISUJE kolory serii w niektorych")
    print("  przypadkach. Dlatego przy wlasnej palecie ustawiamy style=None.")


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **Kolor serii siedzi w `graphicalProperties`.** Nie ma i nie będzie `series.color`. Musisz zapamiętać to jedno słowo — inaczej będziesz szukać po internecie za każdym razem. (W XML-u to `<c:spPr><a:solidFill>`, czyli te same „właściwości kształtu", które opisują każdy element rysunku.)
2. **`LineChart` nie ma domyślnych znaczników.** `series.marker` jest `None` i trzeba go **utworzyć** (`Marker(...)`), a nie modyfikować. To kontrast wobec np. wykresów giełdowych, gdzie fabryka serii sama tworzy `marker`. Dlatego w kodzie jest `Marker(symbol="circle", size=8)`, a nie `series.marker.symbol = "circle"`.
3. **Paleta w jednym miejscu.** `PALETA` to dokładnie ten sam wzorzec co rejestr stylów z modułu 09 i `FORMAT` z modułu 10. Kolory marki w jednym słowniku, używane wszędzie: w `graphicalProperties`, w stylach komórek (moduł 09), w formatowaniu warunkowym (moduł 14). **Spójność raportu bierze się z jednego źródła kolorów, a nie z dyscypliny przy trzydziestym wykresie.**

### Przykład 5 — odczyt wykresów z pliku i diagnostyka (🔴)

Ostatni przykład to narzędzie, do którego będziesz wracał w module 18. Pokazuje, jak **zbadać** plik, który dostałeś od kogoś, i jak wykryć, czego openpyxl nie widzi.

```python
"""Modul 15 - inwentarz wykresow + diagnostyka rozbieznosci.

Cel: narzedzie do URUCHOMIENIA NA CUDZYM PLIKU.
     Porownaj to, co wypisze, z tym, co widzisz w Excelu.

Uruchom:  python examples/15_inwentarz.py [sciezka.xlsx]
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
DOMYSLNY = ROOT / "output" / "15_dwie_osie.xlsx"


def adres(ref) -> str:
    """Bezpiecznie wyciaga napis formuly z obiektu Reference (lub None)."""
    if ref is None:
        return "-"
    try:
        return ref.numRef.f
    except AttributeError:
        pass
    try:
        return ref.strRef.f
    except AttributeError:
        pass
    return "?"


def inwentarz_wykresow(sciezka: Path) -> None:
    wb = load_workbook(sciezka)          # UWAGA: NIE read_only (sekcja 2.13!)
    try:
        print("=" * 100)
        print(f"INWENTARZ WYKRESOW: {sciezka.name}")
        print("=" * 100)

        for ws in wb.worksheets:
            wykresy = list(ws._charts)   # atrybut PRYWATNY - do diagnostyki OK
            if not wykresy:
                continue

            print(f"\nARKUSZ '{ws.title}': {len(wykresy)} wykres(ow)")
            for i, chart in enumerate(wykresy):
                print(f"  [{i}] {type(chart).__name__}")
                print(f"      tytul      : {chart.title}")
                print(f"      styl       : {chart.style}")
                print(f"      rozmiar    : {chart.width} x {chart.height} cm")
                print(f"      legenda    : "
                      f"{'brak' if chart.legend is None else chart.legend.position}")
                print(f"      serie      : {len(chart.series)}")

                for j, s in enumerate(chart.series):
                    print(f"        [{j}] tytul={adres(s.tx)}")
                    print(f"            wartosci={adres(s.val)}")
                    print(f"            kategorie={adres(s.cat)}")
                    etykiety = "tak" if s.dLbls is not None else "nie"
                    print(f"            etykiety danych: {etykiety}")

                # osie
                for nazwa in ("x_axis", "y_axis", "z_axis"):
                    os_ = getattr(chart, nazwa, None)
                    if os_ is None:
                        continue
                    print(f"      {nazwa}: {type(os_).__name__} "
                          f"axId={os_.axId} delete={os_.delete} "
                          f"gridlines={'tak' if os_.majorGridlines else 'nie'}")

                # pod-wykresy (wykresy kombinowane / druga os)
                pod = getattr(chart, "_charts", None) or []
                for k, sub in enumerate(pod):
                    print(f"      + pod-wykres [{k}]: {type(sub).__name__}, "
                          f"seria={len(sub.series)}, "
                          f"axId_y={getattr(sub.y_axis, 'axId', '?')}")

        # --- wykrywanie NIEKOMPLETNOSCI -------------------------------
        print()
        print("=" * 100)
        print("CHECKLISTA ROZBIEZNOSCI (porownaj z tym, co widzisz w Excelu)")
        print("=" * 100)
        print("  [ ] Czy liczba wykresow sie zgadza?")
        print("      Jesli NIE -> openpyxl nie zna czesci wykresow.")
        print("      Najczestsze przyczyny: sparkline'y (x14), nowe typy")
        print("      chartex (mapa, treemap, histogram, waterfall, funnel),")
        print("      wykresy przestawne (pivot charts).")
        print("  [ ] Czy typy sie zgadzaja? (BarChart/LineChart/PieChart...)")
        print("      Jesli NIE -> Excel uzywa wariantu, ktorego openpyxl")
        print("      nie modeluje w pelni.")
        print("  [ ] Czy liczba serii sie zgadza?")
        print("      Jesli NIE -> niektore serie maja nieznana strukture.")
        print("  [ ] Czy adresy wskazuja na WLAŚCIWE arkusze?")
        print("      Jesli adres zawiera '#REF!' -> powiazanie jest zerwane.")
        print("  [ ] Zapisz plik ponownie i porownaj inwentarze PRZED i PO.")
        print("      To najtanszy test utraty danych (modul 18).")
    finally:
        wb.close()


def main() -> None:
    sciezka = Path(sys.argv[1]) if len(sys.argv) > 1 else DOMYSLNY
    if not sciezka.exists():
        print(f"Brak pliku: {sciezka}")
        print("Uruchom najpierw: python examples/15_dwie_osie.py")
        return
    inwentarz_wykresow(sciezka)


if __name__ == "__main__":
    main()
```

**Trzy rzeczy do przemyślenia:**

1. **`load_workbook` bez `read_only`.** To kluczowe. Przy `read_only=True` wykresy **nie są wczytywane** (2.13) — inwentarz byłby pusty i doszedłbyś do błędnego wniosku „plik nie ma wykresów". Pamiętaj o tym przy każdym narzędziu diagnostycznym. (Cena: cały plik w pamięci, moduł 03.)
2. **`ws._charts` to atrybut prywatny.** Podkreślnik oznacza „API wewnętrzne, może się zmienić". Używaj go do **diagnostyki**, nigdy do logiki produkcyjnej. W kodzie raportu wykresy dodawaj świadomie i trzymaj do nich referencję (wzorzec z 2.14).
3. **Checklista rozbieżności to esencja modułu 18.** Zapis → wczytanie → porównanie inwentarza **przed i po** to najtańszy możliwy test utraty danych. Zobaczysz go tam rozwinięty. Ten przykład to zapowiedź.

## 4. Anatomia API

| Klasa / metoda / atrybut | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Reference(ws, ...)` | Adres zakresu komórek | `worksheet`, `min_col`, `min_row`, `max_col`, `max_row`, `range_string` | **Nie czyta danych** — to wskazanie. Indeksy **od 1** |
| `chart.add_data(data, titles_from_data=False, from_rows=False)` | Dodaje serie z zakresu | `titles_from_data` — pierwszy wiersz/kolumna to tytuły; `from_rows` — każdy wiersz to seria | Wywołuj **przed** `set_categories` |
| `chart.set_categories(labels)` | Ustawia etykiety osi X | `Reference` albo napis zakresu | **Cicho nic nie robi, gdy nie ma serii!** Kategorie bez nagłówka |
| `ws.add_chart(chart, anchor)` | Dodaje wykres do arkusza | `anchor` — komórka lub obiekt anchora | Kotwica = lewy górny róg. Wykres „pływa" nad siatką |
| `chart.title` | Tytuł wykresu | `str` albo `Title` | `str` jest automatycznie opakowywany |
| `chart.style` | Preset kolorów Excela | liczba 1–48 | Nadpisuje własne kolory serii w części przypadków |
| `chart.width` / `chart.height` | Rozmiar wykresu | liczby (cm) | Domyślnie $15 \times 7{,}5$ cm |
| `chart.legend` | Legenda | `Legend()` albo `None` | Domyślnie `Legend()` z pozycją `r` |
| `chart.legend.position` | Pozycja legendy | `'r'`, `'l'`, `'t'`, `'b'`, `'tr'` | `chart.legend = None` usuwa legendę |
| `chart.dataLabels` (alias `dLbls`) | Etykiety danych wykresu | `DataLabelList` | Patrz `DataLabelList` niżej |
| `chart.series` | Lista serii | lista `Series` | Tworzona przez `add_data` |
| `chart.graphical_properties` | Tło/obramowanie wykresu | `GraphicalProperties` | `noFill = True` = przezroczyste |
| `chart.anchor` | Kotwica wykresu | komórka lub `AnchorBase` | Można ustawić ręcznie (zaawansowane) |
| `BarChart(...)` | Wykres słupkowy | `barDir="col"/"bar"`, `grouping`, `gapWidth=150`, `overlap` | `type` to alias `barDir` |
| `BarChart.grouping` | Grupowanie | `'clustered'`, `'standard'`, `'stacked'`, `'percentStacked'` | **`stacked` wymaga `overlap = 100`** |
| `BarChart.overlap` | Nachodzenie słupków | `0`–`100` | Ustaw `100` dla warstwowych |
| `BarChart.gapWidth` | Odstępy między grupami | liczba (domyślnie 150) | 80–100 = „gęstszy" wykres |
| `BarChart3D(...)` | Słupkowy 3D | `shape`, `gapDepth`, `z_axis` | `shape` ∈ `'cone'`,`'coneToMax'`,`'box'`,`'cylinder'`,`'pyramid'`,`'pyramidToMax'` |
| `LineChart(...)` | Wykres liniowy | `grouping` | Znaczniki **tylko** przez `Series.marker` |
| `AreaChart(...)` | Wykres obszarowy | `grouping` | Domyślne `'standard'` |
| `PieChart(...)` | Wykres kołowy | `firstSliceAng` | **Tylko JEDNA seria!** |
| `DoughnutChart(...)` | Pierścieniowy | `firstSliceAng`, `holeSize` | `holeSize=50` = klasyczny pierścień |
| `ScatterChart(...)` | Punktowy | `scatterStyle` | Serie **ręcznie** przez `Series(y, x)`; `scatterStyle` nie działa w Excelu |
| `BubbleChart(...)` | Bąbelkowy | — | `Series(values, xvalues, zvalues, title)` |
| `RadarChart(...)` | Radarowy | `type` ∈ `'standard'`,`'marker'`,`'filled'` | — |
| `StockChart(...)` | Giełdowy | `hiLowLines`, `upDownBars` | `hiLowLines` wymaga **hacku z `numCache`** (2.12) |
| `SurfaceChart(...)` | Powierzchniowy | `wireframe` | Rzadko używany |
| `Series(values, xvalues=..., zvalues=..., title=..., title_from_data=...)` | Pojedyncza seria | `values`, `xvalues`, `zvalues`, `title`, `title_from_data` | Aliasy: `graphicalProperties`↔`spPr`, `data_points`↔`dPt`, `dLbls`↔`labels` |
| `Series.graphicalProperties` | Wygląd serii | `GraphicalProperties` | **Kolor słupka/linii jest TU** |
| `Series.marker` | Znacznik punktu | `Marker` albo `None` | `None` domyślnie w `LineChart` |
| `Series.dLbls` | Etykiety serii | `DataLabelList` | Alias `labels` |
| `Series.data_points` | Pojedyncze punkty | lista `DataPoint` | Alias `dPt`. **`idx` od zera** |
| `Marker(symbol=..., size=...)` | Znacznik | `symbol` ∈ `circle`,`dash`,`diamond`,`dot`,`plus`,`square`,`star`,`triangle`,`x`,`auto`,`picture`; `size` 2–72 | Brak wartości `'none'` |
| `DataPoint(idx=..., explosion=...)` | Pojedynczy punkt | `idx` (od 0), `explosion`, `graphicalProperties` | Podstawa wykresu „gauge" i wyciąganych plasterków |
| `DataLabelList(...)` | Etykiety danych | `showVal`, `showCatName`, `showSerName`, `showPercent`, `showLegendKey`, `showBubbleSize`, `showLeaderLines`, `numFmt`, `dLblPos`, `separator` | `dLblPos` ∈ `l`,`t`,`r`,`b`,`ctr`,`outEnd`,`inEnd`,`inBase`,`bestFit` |
| `GraphicalProperties(...)` | Właściwości kształtu | `solidFill`, `noFill`, `line`, `gradFill` | `.line.noFill`, `.line.width` (EMU), `.line.prstDash` |
| `_BaseAxis.title` | Tytuł osi | `str` albo `Title` | `chart.y_axis.title = "Wartość"` |
| `_BaseAxis.numFmt` (alias `number_format`) | Format etykiet osi | kod formatu (moduł 10) | `'0%'`, `'#,##0 "zł"'` |
| `_BaseAxis.scaling` | Skala osi | `Scaling` | `.min`, `.max`, `.logBase`, `.orientation` |
| `_BaseAxis.majorGridlines` | Siatka | `ChartLines()` albo `None` | **Domyślnie włączona** na `NumericAxis` |
| `_BaseAxis.delete` | Ukrycie osi | `True`/`False` | `chart.x_axis.delete = True` |
| `_BaseAxis.crosses` | Punkt przecięcia | `'autoZero'`, `'max'`, `'min'` | Kluczowe dla drugiej osi (2.8) |
| `_BaseAxis.axId` | Identyfikator osi | `int` | Domyślnie: cat 10, val 100, date 500, ser 1000. Druga oś = **200** |
| `Scaling(logBase, orientation, max, min)` | Skala | `orientation` ∈ `'minMax'`,`'maxMin'` | `logBase` na osi X odrzuca wartości ujemne |
| `ChartLines(spPr=None)` | Siatka | — | `chart.y_axis.majorGridlines = ChartLines()` |
| `chart1 += chart2` | Sklejenie wykresów | — | Podstawa drugiej osi i wykresów kombinowanych |
| `wb.create_chartsheet(title)` | Arkusz wykresów | `title` | `cs.add_chart(chart)` |
| `ws._charts` | Lista wykresów arkusza | — | **Prywatne**; do diagnostyki |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — wykres słupkowy sprzedaży po miesiącach.**

Napisz `examples/15_cw1.py`, który tworzy `output/15_cw1.xlsx`:

1. Arkusz `Sprzedaz` z nagłówkiem `Miesiąc`, `Sprzedaż` i 12 wierszami danych (twoje własne liczby, ale muszą być **zróżnicowane** — nie wszystkie podobne).
2. Wykres słupkowy (`type="col"`) z:
   - tytułem `"Sprzedaż 2026"`,
   - tytułem osi X: `"Miesiąc"`, osi Y: `"Sprzedaż (tys. zł)"`,
   - **kategoriami z kolumny A, ale BEZ nagłówka**,
   - tytułem serii z komórki `B1` (`titles_from_data=True`),
   - etykietami danych pokazującymi wartości (`showVal=True`, `numFmt='#,##0'`),
   - **bez legendy** (`chart.legend = None`),
   - rozmiarem $18 \times 9$ cm, zakotwiczonym w `D2`.
3. Funkcję `weryfikuj()`, która wczytuje plik i wypisuje: liczbę wykresów, typ obiektu, liczbę serii oraz **adresy** wartości, kategorii i tytułu serii (użyj kodu z Przykładu 1).

W pliku `output/15_cw1_wnioski.md` odpowiedz:

1. **Dlaczego kategorie podajemy z `min_row=2`, a nie `min_row=1`?** Co by się stało, gdybyś podał `min_row=1`?
2. **Co dokładnie zawiera `seria.val.numRef.f`?** Dlaczego to adres, a nie liczby (1.2)?
3. **Co się stanie, gdy po zapisaniu zmienisz `B3` w Excelu z `120` na `500`?** Dlaczego (1.1)?
4. **Dlaczego `ws._charts` ma podkreślnik?** Co to znaczy dla twojego kodu produkcyjnego?
5. **Co się stanie, gdy wywołasz `set_categories` PRZED `add_data`?** Sprawdź eksperymentalnie i opisz wynik (2.4).

### 🟡 Warsztat

**Zadanie 2 — wykres z dwiema osiami (sprzedaż + marża %).**

To zadanie z briefu. Napisz `examples/15_cw2.py`:

1. Arkusz `Dane` z kolumnami: `Miesiąc`, `Sprzedaż`, `Marża`.
2. Kolumnę `Marża` sformatuj jako procent (`number_format = '0.0%'` — moduł 10) i wpisz **współczynniki** (np. `0.18`), nie liczby całkowite.
3. **Wykres z dwiema osiami:**
   - słupki = sprzedaż (lewa oś, `"Sprzedaż (tys. zł)"`),
   - linia = marża (prawa oś, `"Marża (%)"`, `numFmt='0%'`),
   - drugie wykres z `y_axis.axId = 200`,
   - **wyłącz siatkę** na obu osiach wartości,
   - `c1.y_axis.crosses = "max"`,
   - sklej przez `c1 += c2`.
4. **Wykres kontrolny** obok — te same dane **na jednej osi** (bez drugiej osi). To ma być celowo brzydkie.
5. Funkcję `weryfikuj()`, która wczytuje plik i wypisuje `axId` i `crosses` osi obu wykresów.

W pliku `output/15_cw2_wnioski.md`:

1. **Dlaczego marża na jednej osi ze sprzedażą jest wizualnie bezużyteczna?** Policz rząd wielkości obu zbiorów i uzasadnij.
2. **Do czego służy `y_axis.axId = 200`?** Co by się stało, gdyby oba wykresy miały `axId = 100` (2.8)?
3. **Dlaczego wyłączyłeś siatkę na obu osiach?** Obejrzyj wersję z dwiema siatkami i opisz efekt (2.7).
4. **Czy `c1.y_axis.crosses = "max"` umieściło oś po właściwej stronie?** Jeśli nie — przenieś to ustawienie na `c2.y_axis` i opisz różnicę. **To pytanie celowo nie ma jednej odpowiedzi** — chodzi o to, byś zauważył, że trzeba weryfikować okiem (2.8).
5. **Co byś zrobił, gdyby marża miała wartości ujemne?** Czy skala `logBase` byłaby dobrym pomysłem? Uzasadnij (2.7 — logarytm odrzuca wartości ujemne).

### 🔴 Wyzwanie

**Zadanie 3 — dashboard z trzech wykresów z różnych arkuszy, jednolity styl.**

To zadanie z briefu i zarazem zapowiedź wzorca **Facade** (moduł 24). Zbudujesz coś, co ma **wspólny styl** i **oddzielone dane od prezentacji**.

```python
"""Dashboard: 3 wykresy z roznych arkuszy, jednolity styl.

Wymagania:
  1. Arkusze danych: "Sprzedaz", "Regiony", "Trend" - kazdy z wlasnymi danymi.
  2. Arkusz "Dashboard" - puste miejsce na trzy wykresy.
  3. Wykresy na Dashboardzie odwołują się do ZAKRESÓW z innych arkuszy.
  4. Wszystkie trzy mają: ten sam rozmiar, ta sama paletę, brak legendy,
     ten sam format liczb na osi Y, wylaczona siatke pomocnicza.
  5. JEDNA funkcja fabryczna dla wszystkich wykresów (wzorzec z 2.14).

Szkic API:
    STYL = {
        "szerokosc": 16, "wysokosc": 8, "legenda": None,
        "paleta": ["5B5CF0", "F0705B", "3BA55C"],
        "numfmt_osi": "#,##0", "siatka": False,
    }

    def wykres_slupkowy(ws, zakres, kategorie, tytul, anchor, styl, kolory):
        ...

    def wykres_liniowy(ws, zakres, kategorie, tytul, anchor, styl, kolory):
        ...

    def wykres_kolowy(ws, zakres, kategorie, tytul, anchor, styl, kolory):
        ...
"""
```

**Wymagania:**

1. **Trzy arkusze danych** (`Sprzedaz`, `Regiony`, `Trend`) — każdy z nagłówkiem i co najmniej 8 wierszami. Dane mają sens (sprzedaż po miesiącach, udziały regionów, trend czegoś).
2. **Arkusz `Dashboard`** — bez danych, tylko trzy wykresy w pionie (`F2`, `F22`, `F42`).
3. **Wykresy odwołują się do innych arkuszy.** To kluczowy punkt: `Reference(ws_sprzedaz, ...)` i `ws_dashboard.add_chart(chart, ...)`. Sprawdź, że adresy w inwentarzu (Przykład 5) zawierają **nazwy właściwych arkuszy**.
4. **Jednolity styl** — jeden słownik konfiguracji, z którego korzystają wszystkie trzy wykresy. Zmiana `wysokosc` w słowniku zmienia wszystkie trzy.
5. **Funkcja fabryczna** — nie buduj wykresów w `main()`. Każdy typ ma swoją funkcję przyjmującą opis.
6. **Sprawdzenie spójności** — funkcja `sprawdz_styl()`, która wczytuje plik i weryfikuje, że wszystkie wykresy mają ten sam `numFmt` osi Y i ten sam rozmiar.

W pliku `output/15_cw3_wnioski.md`:

1. **Jak uniknąłeś powtórzenia stylu w trzech miejscach?** Opisz mechanizm (słownik + funkcje fabryczne).
2. **Czym różni się `Reference(ws_innego_arkusza, ...)` od `Reference(ws_dashboard, ...)`?** Co się dzieje z nazwą arkusza w adresie (podpowiedź: spójrz na `seria.val.numRef.f`)?
3. **Co się stanie, gdy zmienisz nazwę arkusza `Sprzedaz` na `Sprzedaż 2026` (ze spacją)?** Sprawdź, czy wykres nadal działa. Czemu zawdzięczasz wynik? (podpowiedź: cudzysłowy w adresie)
4. **Dlaczego funkcje fabryczne mają `styl` jako parametr, a nie mają go wpisanego w środku?** Odnieś się do modułu 22 (wzorce) i zasady „konfiguracja zamiast kodu".
5. **Jak sprawdziłbyś automatycznie, że wszystkie trzy wykresy mają ten sam styl?** Opisz, jak to zrobić na podstawie `ws._charts` (Przykład 5).
6. **To realizuje pewien wzorzec.** Który wzorzec z modułu 24 opisuje obiekt ukrywający wiele operacji na wykresie? Uzasadnij w dwóch zdaniach.

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Dashboard z trzema wykresami - kluczowe fragmenty."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties

# --- JEDNO zrodlo prawdy o stylu (wzorzec z modulu 09 i 2.14) --------
STYL = {
    "szerokosc": 16,
    "wysokosc": 8,
    "paleta": ["5B5CF0", "F0705B", "3BA55C"],
    "numfmt_osi": "#,##0",
    "siatka": False,
    "etykiety": True,
}


def _nadaf_styl(chart, tytul: str, osie: tuple[str, str], styl: dict) -> None:
    """Wspolny szkielet stylu - jedno miejsce dla wszystkich wykresow."""
    chart.title = tytul
    chart.x_axis.title, chart.y_axis.title = osie
    chart.width = styl["szerokosc"]
    chart.height = styl["wysokosc"]
    chart.legend = None
    chart.y_axis.numFmt = styl["numfmt_osi"]
    if not styl["siatka"]:
        chart.y_axis.majorGridlines = None
    if styl["etykiety"]:
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.numFmt = styl["numfmt_osi"]


def wykres_slupkowy(ws_zrodlo, ws_cel, zakres, kategorie, tytul, anchor, styl, tytul_osi_y):
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    _nadaf_styl(chart, tytul, ("Kategoria", tytul_osi_y), styl)

    chart.add_data(zakres, titles_from_data=True)   # 1) dane
    chart.set_categories(kategorie)                 # 2) kategorie

    chart.series[0].graphicalProperties = GraphicalProperties(
        solidFill=styl["paleta"][0])
    ws_cel.add_chart(chart, anchor)
    return chart


def wykres_liniowy(ws_zrodlo, ws_cel, zakres, kategorie, tytul, anchor, styl, tytul_osi_y):
    chart = LineChart()
    chart.series.clear() if False else None    # nic nie robimy - add_data doda serie
    _nadaf_styl(chart, tytul, ("Kategoria", tytul_osi_y), styl)
    chart.add_data(zakres, titles_from_data=True)
    chart.set_categories(kategorie)
    for seria, kolor in zip(chart.series, styl["paleta"]):
        seria.graphicalProperties = GraphicalProperties()
        seria.graphicalProperties.line.solidFill = kolor
        seria.graphicalProperties.line.width = 30000
    ws_cel.add_chart(chart, anchor)
    return chart


def wykres_kolowy(ws_zrodlo, ws_cel, zakres, kategorie, tytul, anchor, styl):
    chart = PieChart()
    chart.title = tytul
    chart.width = styl["szerokosc"]
    chart.height = styl["wysokosc"]
    chart.add_data(zakres, titles_from_data=True)
    chart.set_categories(kategorie)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True        # dla kolowego to procent
    chart.dataLabels.showCatName = True
    ws_cel.add_chart(chart, anchor)
    return chart


def zbuduj(cel: Path) -> None:
    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"

    # --- arkusz 1: sprzedaz po miesiacach ---------------------------
    ws1 = wb.create_sheet("Sprzedaz")
    ws1.append(["Miesiąc", "Sprzedaż"])
    for m, v in zip(["Sty", "Lut", "Mar", "Kwi", "Maj", "Cze",
                     "Lip", "Sie", "Wrz", "Paz"],
                    [120, 135, 98, 150, 172, 160, 143, 158, 181, 190]):
        ws1.append([m, v])

    # --- arkusz 2: regiony ------------------------------------------
    ws2 = wb.create_sheet("Regiony")
    ws2.append(["Region", "Udział"])
    for r, v in [("Północ", 32), ("Południe", 28), ("Wschód", 22), ("Zachód", 18)]:
        ws2.append([r, v])

    # --- arkusz 3: trend --------------------------------------------
    ws3 = wb.create_sheet("Trend")
    ws3.append(["Tydzień", "Odwiedziny"])
    for t, v in zip([f"T{i}" for i in range(1, 11)],
                    [800, 850, 820, 900, 980, 1020, 1100, 1050, 1200, 1350]):
        ws3.append([t, v])

    # --- trzy wykresy na Dashboardzie, dane Z INNYCH arkuszy --------
    wykres_slupkowy(
        ws1, dash,
        Reference(ws1, min_col=2, min_row=1, max_row=11),
        Reference(ws1, min_col=1, min_row=2, max_row=11),
        "Sprzedaż po miesiącach", "A2", STYL, "Sprzedaż",
    )
    wykres_liniowy(
        ws3, dash,
        Reference(ws3, min_col=2, min_row=1, max_row=11),
        Reference(ws3, min_col=1, min_row=2, max_row=11),
        "Trend odwiedzin", "A22", STYL, "Odwiedziny",
    )
    wykres_kolowy(
        ws2, dash,
        Reference(ws2, min_col=2, min_row=1, max_row=5),
        Reference(ws2, min_col=1, min_row=2, max_row=5),
        "Udział regionów", "A42", STYL,
    )

    wb.save(cel)
    wb.close()
```

**Kluczowe decyzje i uzasadnienia:**

- **`_nadaf_styl` jako wspólny szkielet.** Trzy funkcje fabryczne mają **różne** typy wykresu, ale **ten sam** styl. Wspólna funkcja pomocnicza eliminuje duplikację dokładnie tak, jak `BazowyRaport` z modułu 25 (Template Method) wyeliminuje duplikację szkieletu raportu. To ten sam pomysł w mniejszej skali.

- **`STYL` jako parametr, a nie stała globalna wewnątrz funkcji.** Funkcja fabryczna przyjmuje `styl` z zewnątrz — dzięki temu możesz wygenerować **drugi dashboard z inną paletą**, nie kopiując kodu. To dokładnie zasada „konfiguracja zamiast kodu" z 2.14 i zapowiedź deklaratywnych raportów z modułu 26. (I dokładnie ten sam mechanizm co `KONFIG` w module 14.)

- **`Reference(ws_zrodlo, ...)` + `ws_cel.add_chart(...)`.** To rozdzielenie jest istotne: **dane są w innym arkuszu niż wykres**. openpyxl zapisze w adresie nazwę arkusza źródłowego (`'Sprzedaz'!$B$2:$B$11`) i doda relację z wykresu do arkusza (`xl/charts/_rels/chart1.xml.rels`). Dlatego w XML-u są osobne pliki relacji dla wykresów — to nie ozdoba, a konieczność techniczna.

- **Nazwy arkuszy ze spacjami i polskimi znakami.** `Reference` sam zadba o cudzysłowy wokół nazwy (`'Sprzedaż 2026'!$B$2`). Historycznie openpyxl miał z tym problemy (`#1190 Cannot create charts for worksheets with quotes in the title`, naprawione), dlatego **nazwy arkuszy z apostrofem** (`Jan's dane`) to obszar, który warto przetestować u siebie. Reszta — spacje, polskie znaki, myślniki — działa.

- **Dlaczego nie `add_data` z dwóch arkuszy naraz?** Nie da się — `add_data` przyjmuje jeden `Reference`, więc seria należy do jednego arkusza. Możesz natomiast dodać **dwie serie z różnych arkuszy** wywołując `add_data` dwa razy. To rzadkie, ale możliwe — i wtedy w wykresie są adresy z dwóch różnych arkuszy.

- **Sprawdzenie spójności stylu.** `sprawdz_styl()` czyta `ws._charts` z arkusza `Dashboard` (Przykład 5) i porównuje `chart.width`, `chart.height` i `chart.y_axis.numFmt` dla wszystkich trzech wykresów. To jest dokładnie ten rodzaj testu, który w module 20 zbudujesz na `BytesIO` — bez zapisywania na dysk.

- **To jest zalążek Facade.** Zauważ, że funkcje fabryczne przyjmują **typy openpyxl** w sygnaturach (`Reference`, `GraphicalProperties`). To jeszcze nie fasada w rozumieniu modułu 24 — to *warstwa pośrednia*. W module 24 zobaczysz, jak zmienić sygnatury tak, żeby **żaden typ openpyxl nie przeciekał na zewnątrz** (przyjmowanie `list[Zamowienie]` zamiast `Reference`). Wtedy będzie można przetestować logikę raportu **bez openpyxl w ogóle** — i to jest cały sens fasady.

</details>

## 6. Typowe błędy i pułapki

**1. „Wykres ma puste kategorie na osi X — nie ma nazw miesięcy" (objaw) → **`set_categories` zostało wywołane PRZED `add_data`.** Metoda iteruje po istniejących seriach; gdy lista jest pusta, pętla wykonuje się zero razy i **openpyxl nie zgłasza błędu** (2.4) (przyczyna) → **Zawsze `add_data` najpierw, `set_categories` potem.** Dodaj do weryfikacji sprawdzenie `seria.cat is not None`. To najczęstszy błąd całego modułu i dlatego jest w ścieżce weryfikacji w Przykładzie 1 (naprawa).**

**2. „Na wykresie jest dodatkowa kategoria «Miesiąc» i wszystko przesunięte o jedno miejsce" (objaw) → **zakres kategorii zawiera wiersz nagłówka** (`min_row=1` zamiast `min_row=2`). Dane mają 12 wartości, ale kategorii jest 13 (2.3, 2.4) (przyczyna) → **zakres kategorii zaczynaj od wiersza danych, nie od nagłówka.** Reguła: jeśli `titles_from_data=True` dla danych, to kategorie **bez** nagłówka (naprawa).**

**3. „Sąsiad w firmie mówi, że nie ma klasy `NumberAxis` — a mój tutorial jej używa" (objaw) → **brief i wiele tutoriali podają `NumberAxis`, ale w openpyxl 3.1.x klasa nazywa się `NumericAxis` i nie ma aliasu** (2.7, potwierdzone w pełnym źródle `axis.py`) (przyczyna) → **importuj `from openpyxl.chart.axis import NumericAxis`** — albo, co lepsze, **nie importuj osi wcale**: oś powstaje sama (`BarChart` tworzy `TextAxis` + `NumericAxis`), a ty tylko **modyfikujesz** istniejące `chart.x_axis`/`chart.y_axis`. `DateAxis` importuj tylko wtedy, gdy naprawdę chcesz oś czasu (naprawa).**

**4. „Wykres jest pokryty «deszczem» z linii siatki — są dwie siatki" (objaw) → **każda `NumericAxis` ma `majorGridlines` **domyślnie włączone** (potwierdzone w źródle: `kw.setdefault('majorGridlines', ChartLines())`).** Wykres z dwiema osiami wartości = dwie siatki (2.7) (przyczyna) → **wyłącz siatkę jawnie na osi, której nie potrzebujesz: `chart.y_axis.majorGridlines = None`.** Przy wykresie kombinowanym musisz to zrobić dla obu osi wartości (naprawa).**

**5. „Wykres zasłania moje dane" (objaw) → **anchor wskazuje na komórki, w których są dane.** Wykres „pływa" nad siatką i zasłania to, co pod nim (1.3) (przyczyna) → **planuj układ: dane po lewej, wykresy po prawej.** Ustal konwencję (np. wykresy od kolumny F, w odstępach 20 wierszy) i trzymaj się jej. Jeśli wykresów jest dużo — rozważ osobny `Chartsheet` (2.13) lub osobny arkusz (naprawa).**

**6. „Excel pokazuje «uszkodzony plik» przy tak dużym `IconSet`… nie, przy wykresie z dwiema osiami" (objaw) → **problem z kolejnością serii w wykresie kombinowanym.** Serie mają atrybuty `idx` i `order`, które muszą być spójne. Źródło openpyxl ma komentarz `„The index can need rebasing"` (2.12) (przyczyna) → **Zbuduj wykres kombinowany od zera, w jednej kolejności, zamiast łączyć wykresy wczytane z pliku.** Jeśli nadal jest problem: uprość do jednego typu wykresu i sprawdź, czy plik się otwiera. Wykresy kombinowane to obszar, w którym trzeba testować empirycznie (naprawa).**

**7. „Kolory serii, które ustawiłem, nie działają — wykres ma inne kolory" (objaw) → **ustawiłeś `chart.style` (preset Excela), który nadpisuje własne `graphicalProperties`.** Presety i własne kolory walczą o to samo (2.10) (przyczyna) → **wybierz jedno: albo `chart.style`, albo własne `GraphicalProperties` na seriach.** Przy palecie brandowej ustaw `chart.style = None` i kolory serii jawnie (naprawa).**

**8. „Znaczniki na wykresie liniowym nie chcą się ustawić — `series.marker.symbol` rzuca `AttributeError`" (objaw) → **w `LineChart` `series.marker` jest domyślnie `None`.** Nie można modyfikować atrybutu nieistniejącego obiektu (2.10) (przyczyna) → **utwórz marker: `series.marker = Marker(symbol="circle", size=8)`.** Uwaga: w wykresach giełdowych `marker` bywa tworzony przez fabrykę serii — tam `series.marker.symbol` może działać (naprawa).**

**9. „Etykiety danych pokazują wartości, ale z dziwnym formatowaniem" (objaw) → **brak `numFmt` na `DataLabelList`.** Etykiety dziedziczą domyślny format ogólny Excela (2.9) (przyczyna) → **ustaw `chart.dataLabels.numFmt`** tak samo jak `number_format` w komórkach (moduł 10). Format liczb jest jeden dla całego pliku — komórki, osie, etykiety (naprawa).**

**10. „Wykres pokazuje pustkę, choć w komórkach są liczby" (objaw) → **w komórkach są formuły bez pamięci podręcznej wartości.** openpyxl nie ma silnika obliczeń (moduł 06), a wykres czyta **adresy**, nie wyniki. Plik nigdy nie był otwarty w Excelu, więc wartości nie istnieją w pliku (2.12) (przyczyna) → **jeśli budujesz wykres z danych generowanych przez openpyxl, wpisuj WARTOŚCI, nie formuły.** Albo otwórz i zapisz plik w Excelu przed dalszym przetwarzaniem (naprawa).**

**11. „Nie mogę znaleźć wykresów w wczytanym pliku — lista jest pusta" (objaw) → **użyłeś `load_workbook(..., read_only=True)`.** W trybie `read_only` wykresy i obrazy **nie są wczytywane** (parser robi `continue` przed przetwarzaniem rysunków) (2.13) (przyczyna) → **do odczytu wykresów użyj zwykłego `load_workbook`.** Zaakceptuj koszt pamięci (moduł 03). W `write_only` (generowanie) wykresy **działają** — to nie jest symetryczne (naprawa).**

**12. „Wykres kołowy wygląda dziwnie — Excel pokazuje coś nie takiego" (objaw) → **do `PieChart` dodano więcej niż jedną serię.** Wykresy kołowe przyjmują **jedną** serię danych (dokumentacja: *„Pie charts can only take a single series of data"*) (2.6) (przyczyna) → **podaj tylko jedną kolumnę danych**: `Reference(ws, min_col=2, min_row=1, max_row=N)`. Jeśli masz trzy kolumny i chcesz trzy wykresy kołowe — to trzy wykresy, nie jeden (naprawa).**

**13. „Wykres punktowy pokazuje dziwne wartości albo jest pusty" (objaw) → **użyłeś `add_data` dla `ScatterChart`.** Wykres rozproszenia buduje serie **ręcznie** przez `Series(yvalues, xvalues)` (2.6). Dodatkowo, jeszcze w 3.1.0 był błąd `#1360 Can't read some ScatterCharts if the x-axis is not numerical`, więc **oś X musi być numeryczna** (przyczyna) → **dla scatter: `series = Series(yvals, xvals, title_from_data=True); chart.series.append(series)`.** I upewnij się, że kolumna X zawiera **liczby**, nie tekst (moduł 05) (naprawa).**

**14. „Ustawienie `scatterStyle` nic nie zmienia" (objaw) → **dokumentacja openpyxl ostrzega wprost: w Excelu `scatterStyle` to tylko skrót do innych ustawień, które inaczej nie mają efektu** (2.6) (przyczyna) → **styl każdej serii ustawiaj ręcznie** — `series.marker`, `series.graphicalProperties.line`. To ta sama rodzina co „magia, która nic nie robi" z module 14 (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Wykres to przepis, nie potrawa.** W pliku `.xlsx` wykres to osobny XML (`xl/charts/chartN.xml`) zawierający **adresy komórek** (`'Sprzedaz'!$B$2:$B$13`), a nie skopiowane wartości. Excel wykonuje ten przepis od nowa przy każdym otwarciu. Dlatego zmiana komórki zmienia wykres, a usunięcie komórki go psuje. openpyxl nie renderuje niczego — jest skrybą piszącym przepisy.

2. **`Reference` to palec wskazujący, `add_data` to budowa serii, `set_categories` to etykiety osi X.** Trzy filary, w tej kolejności. `add_data` **przed** `set_categories` — inaczej kategorie nie zostaną ustawione, a openpyxl **nie zgłosi błędu**. Zakres danych może zawierać nagłówek (`titles_from_data=True`), zakres kategorii — **nigdy**.

3. **Oś to obiekt z własnymi właściwościami.** `chart.x_axis` i `chart.y_axis` są pełnoprawnymi obiektami (`TextAxis`, `NumericAxis`, `DateAxis`). Mają tytuł, format liczb (`numFmt`, moduł 10), skalowanie (`scaling.min/max/logBase/orientation`), możliwość ukrycia (`delete = True`) i — uwaga — **domyślnie włączoną siatkę** (`majorGridlines`). Nazwa klasy to **`NumericAxis`**, nie `NumberAxis`.

4. **Druga oś to drugi wykres.** Nie ma metody `add_second_axis()`. Tworzysz osobny wykres (`LineChart`), nadajesz mu **nowy `axId`** na osi wartości (200), ustawiasz `crosses = "max"` na osi, która ma się pokazać po prawej, i sklejasz przez `chart1 += chart2`. W pliku powstaje plotArea z dwoma wykresami i czterema osiami.

5. **openpyxl implementuje klasyczny OOXML — i tylko tyle.** Potrafi wszystkie klasyczne typy (słupkowe, liniowe, kołowe, punktowe, obszarowe, radarowe, bąbelkowe, giełdowe, powierzchniowe, 2D i 3D), wykresy kombinowane, drugą oś i wykresy w `write_only`. **Nie potrafi:** sparkline'ów, nowych typów „chartex" (mapa, treemap, sunburst, histogram, Pareto, waterfall, funnel, pudełko-wąsy), wykresów przestawnych. **Gubi** je przy zapisie. **Nie wczytuje** wykresów w trybie `read_only`. I wreszcie: **nie liczy formuł** — wykres zbudowany na formułach bez pamięci podręcznej pokaże pustkę.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORT - TYPY WYKRESOW (wszystkie z openpyxl.chart)
# ==================================================================
from openpyxl.chart import (
    BarChart, BarChart3D,
    LineChart, LineChart3D,
    AreaChart, AreaChart3D,
    PieChart, PieChart3D, ProjectedPieChart, DoughnutChart,
    ScatterChart, BubbleChart, RadarChart, StockChart,
    SurfaceChart, SurfaceChart3D,
    Reference, Series,
)
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.axis import ChartLines, NumericAxis, TextAxis, DateAxis
# ⚠️ NIE MA klasy NumberAxis! Jest NumericAxis.

# ==================================================================
# 2. NAJPROSTSZY WYKRES - PIEC LINII
# ==================================================================
chart = BarChart()                     # domyslnie type="col", grouping="clustered"
chart.type = "col"                     # "col" = pionowe, "bar" = poziome
chart.title = "Sprzedaż 2026"
chart.style = 10                       # preset 1-48 (ALBO wlasne kolory, nie oba)
chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=13),
               titles_from_data=True)   # 1) DANE
chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))  # 2) KATEGORIE (bez naglowka!)
ws.add_chart(chart, "D2")              # pinezka: lewy gorny rog w D2

# ==================================================================
# 3. TYPY - SZYBKIE SNIPPETY
# ==================================================================
c = BarChart(); c.type = "col"; c.grouping = "clustered"; c.gapWidth = 80
c = BarChart(); c.type = "col"; c.grouping = "stacked"; c.overlap = 100   # stacked!
c = LineChart()                        # znaczniki TYLKO przez Series.marker
c = AreaChart(); c.grouping = "standard"
c = PieChart()                         # ⚠️ TYLKO JEDNA SERIA!
c = DoughnutChart(firstSliceAng=270, holeSize=50)
c = RadarChart(); c.type = "marker"
c = SurfaceChart(); c.wireframe = False

# PUNKTOWY - bez add_data, serie recznie!
c = ScatterChart()
c.series.append(Series(yvals, xvals, title_from_data=True))
# BABELKOWY
c = BubbleChart()
c.series.append(Series(values=yvals, xvalues=xvals, zvalues=zvals, title="2026"))

# ==================================================================
# 4. OSIE
# ==================================================================
chart.x_axis.title = "Miesiąc"
chart.y_axis.title = "Sprzedaż (tys. zł)"
chart.y_axis.numFmt = '#,##0 "zł"'          # format etykiet (modul 10)
chart.y_axis.scaling.min = 0                # wymus zakres
chart.y_axis.scaling.max = 300
chart.y_axis.scaling.logBase = 10           # skala log (uwaga: odrzuca ujemne na X)
chart.x_axis.scaling.orientation = "maxMin" # odwroc os
chart.y_axis.majorGridlines = None          # ⚠️ siatka jest DOMYSLNIE WLACZONA!
chart.y_axis.majorGridlines = ChartLines()  # wlacz z powrotem
chart.x_axis.delete = True                  # ukryj os
chart.y_axis.delete = True
# os czasu (rzadko):
c.x_axis = DateAxis()

# ==================================================================
# 5. DRUGA OS - SLUPKI (sprzedaz) + LINIA (marza %)
# ==================================================================
c1 = BarChart(); c1.type = "col"; c1.title = "Sprzedaż i marża"
c1.y_axis.title = "Sprzedaż"; c1.y_axis.majorGridlines = None
c1.add_data(Reference(ws, min_col=2, min_row=1, max_row=13), titles_from_data=True)
c1.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))

c2 = LineChart()
c2.y_axis.axId = 200                # NOWY identyfikator osi wartosci (domyslnie 100)
c2.y_axis.title = "Marża (%)"; c2.y_axis.numFmt = "0%"
c2.y_axis.majorGridlines = None
c2.add_data(Reference(ws, min_col=3, min_row=1, max_row=13), titles_from_data=True)
c2.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))

c1.y_axis.crosses = "max"           # gdzie os ma przeciac os X (patrz dokumentacja!)
c1 += c2                            # KOMBINACJA
ws.add_chart(c1, "E2")

# ==================================================================
# 6. LEGENDA I ETYKIETY DANYCH
# ==================================================================
chart.legend.position = "b"         # r (domyslna), l, t, b, tr
chart.legend = None                 # brak legendy
chart.dataLabels = DataLabelList()  # alias: chart.dLbls
chart.dataLabels.showVal = True
chart.dataLabels.showCatName = True
chart.dataLabels.showPercent = True       # dla kolowego!
chart.dataLabels.numFmt = '#,##0'
chart.dataLabels.dLblPos = "outEnd"       # l,t,r,b,ctr,outEnd,inEnd,inBase,bestFit
# tylko jedna seria:
chart.series[0].dLbls = DataLabelList()
chart.series[0].dLbls.showVal = True

# ==================================================================
# 7. STYLOWANIE (GraphicalProperties)
# ==================================================================
from openpyxl.chart.shapes import GraphicalProperties
chart.series[0].graphicalProperties = GraphicalProperties(solidFill="5B5CF0")
# kolor pojedynczego punktu (idx OD ZERA!):
p = DataPoint(idx=0); p.graphicalProperties.solidFill = "FF0000"
chart.series[0].data_points = [p]
# znaczniki na linii:
chart.series[0].marker = Marker(symbol="circle", size=8)
# symbol: circle, dash, diamond, dot, plus, square, star, triangle, x, auto, picture
# brak 'none'! Aby usunac znaczniki - zostaw marker=None
# gruba kreska:
chart.series[0].graphicalProperties = GraphicalProperties()
chart.series[0].graphicalProperties.line.solidFill = "5B5CF0"
chart.series[0].graphicalProperties.line.width = 30000   # EMU
# przezroczyste tlo wykresu + brak obramowania:
chart.graphical_properties = GraphicalProperties()
chart.graphical_properties.noFill = True
chart.graphical_properties.line.noFill = True
# ukryj linie, zostaw markery (wykresy gieldowe):
chart.series[0].graphicalProperties.line.noFill = True

# ==================================================================
# 8. ROZMIAR, POZYCJA, ARKUSZ WYKRESOW
# ==================================================================
chart.width = 18            # cm (domyslnie 15)
chart.height = 9            # cm (domyslnie 7.5)
ws.add_chart(chart, "E2")   # anchor = lewy gorny rog

cs = wb.create_chartsheet(title="Duzy wykres")
cs.add_chart(chart)

# ==================================================================
# 9. ODCZYT I DIAGNOSTYKA
# ==================================================================
wb = load_workbook(path)          # ⚠️ NIE read_only! (wykresy niedostepne)
for ws in wb.worksheets:
    for chart in ws._charts:      # ⚠️ atrybut PRYWATNY - tylko diagnostyka
        print(type(chart).__name__, chart.title, len(chart.series))
        for s in chart.series:
            print("  val:", s.val.numRef.f)          # adres wartosci
            print("  cat:", s.cat.numRef.f if s.cat else "-")
            print("  tx :", s.tx.strRef.f if s.tx else "-")
        print("  osY :", chart.y_axis.axId, chart.y_axis.crosses)

# ==================================================================
# 10. CZEGO NIE MA (trzy kategorie z modulu 03)
# ==================================================================
# SPARKLINE'Y          -> brak API, x14 ext, GUBIONE przy zapisie
# CHARTEX (Excel 2016+) -> mapa, treemap, sunburst, histogram, Pareto,
#                          waterfall, funnel, pudelko-wasy: NIE OBSLUGIWANE
# WYKRESY PRZESTAWNE    -> nie tworzone
# READ_ONLY             -> wykresy NIEDOSTEPNE (write_only: DZIALAJA!)
# SILNIK OBLICZEN       -> brak (modul 06): wykres na formulach bez cache = pustka
```

## 9. Co dalej

Wykresy to trzeci „świat obiektów" w kursie — po danych i formatowaniu. I pierwszy, w którym tak wyraźnie zobaczyłeś **trzy kategorie prawdy z modułu 03** w akcji: „potrafi" (klasyczne typy), „potrafi częściowo" (odczyt wykresu z Excela, wykresy kombinowane, linie hi-lo) i „gubi" (sparkline'y, chartex, wykresy przestawne). Zapamiętaj tę strukturę — to nie jest lista ciekawostek, a **mapa ryzyka**, która wróci w module 18 z pełną siłą.

Trzy rzeczy, które zabierasz do dalszej pracy:

- **Model „adres, nie kopia".** `Reference` wskazuje komórki, wykres trzyma wskazanie. To ta sama intuicja co formuła (moduł 06) i `print_area` (moduł 11) i `ref` tabeli (moduł 12) i `sqref` reguły warunkowej (moduł 13). **Cały Excel jest zbudowany na adresach i wskazaniach.** Gdy to zrozumiesz, przestaniesz szukać „magicznych" zachowań.
- **Dyscyplinę „jedna funkcja, jeden wykres".** Wzorzec z 2.14 to zalążek **Facade** (moduł 24). Zaczynasz od porządku, nie od abstrakcji — ale ten porządek zaprocentuje, gdy raportów będzie dziesięć i trzeba będzie zmienić wspólny styl.
- **Nawyk weryfikacji okiem.** Wykresy są pełne miejsc (crosses, dLblPos, indeksowanie kotwicy anchora), w których openpyxl **wykona to, co kazałeś**, ale efekt zależy od Excela. To ta sama lekcja co priorytety reguł w module 14 — **kod zapisuje instrukcję, a nie gwarantuje wyglądu.**

**Zadanie do zrobienia teraz, zanim ruszysz dalej.** Otwórz `output/15_dwie_osie.xlsx` w Excelu i wykonaj jedno ćwiczenie z odczytem:

1. Z menu **Projektowanie wykresu → Zmień typ wykresu** zmień linię marży na słupki **na osi wtórnej**. Zapisz plik.
2. Wczytaj go openpyxl (Przykład 5) i wypisz inwentarz.
3. Sprawdź, czy openpyxl nadal widzi **dwa** pod-wykresy. Jeśli tak — zmiana typu wykresu w Excelu jest dla openpyxl widoczna w pełni. **To dobra wiadomość i warto ją znać.**

To jest sprawdzenie granicy między „potrafi" i „potrafi częściowo" w praktyce — dokładnie ta umiejętność, którą rozwiniesz w module 18.

W **module 16** wyjdziemy poza dane i wykresy, do rzeczy, które **nie są komórkami i nie są wykresami**: **obrazy, hiperlinki i komentarze.** Zobaczysz, że obrazy mają pozycję (anchor — ten sam mechanizm co kotwica wykresu, tylko piksele zamiast centymetrów), że **nie są kopiowane** razem z arkuszem, i że ich wstawienie wymaga dodatkowej biblioteki (`pillow`). Będzie też o tym, jak wygenerować wykres w `matplotlib` i wstawić go do arkusza jako **obraz** — czyli obejście wszystkich ograniczeń „chartex" z tego modułu. I o tekstach alternatywnych, które w instytucjach publicznych nie są uprzejmością, a wymogiem.