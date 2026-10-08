Zweryfikuję fakty techniczne dla tego kluczowego modułu: pełną sygnaturę `load_workbook`, dokładne brzmienie ostrzeżeń, zachowanie `keep_links`, `rich_text`, wsparcie kontekst managera oraz problem reindeksacji serii wykresów.

Ważne odkrycie: pojawił się konflikt źródeł co do kontekst managera. Zweryfikuję to dokładnie, bo brief wymaga ostrożności w tym punkcie, oraz sprawdzę ostrzeżenia o „shapes" i zachowanie `keep_links`.

Zweryfikuję bezpośrednio źródło `Workbook` w 3.1.5 (sprawa kontekst managera jest kluczowa) oraz changelog.

Kluczowe ustalenie: `Workbook` w 3.1.x **nie implementuje** `__enter__`/`__exit__` — brief sugerował sprawdzenie i to potwierdzam. Zweryfikuję jeszcze ostrzeżenia o „shapes"/rozszerzeniach oraz problem reindeksacji serii wykresów.

# Moduł 18 — Modyfikacja istniejących plików bez utraty zawartości

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–17

## 0. W tym module nauczysz się

- **Zrozumiesz, dlaczego `wb.save()` to przepisanie dokumentu od nowa**, a nie edycja w miejscu — i jakie są z tego konsekwencje dla cudzego, „bogatego" pliku.
- **Poznasz pełną sygnaturę `load_workbook()`** oraz wszystkie sześć parametrów wraz z tym, **co każdy z nich zepsuje**, jeśli użyjesz go bezmyślnie.
- **Będziesz mieć kompletny inwentarz trzech kategorii**: co przechodzi bez zmian, co przechodzi częściowo i wymaga oględzin, a co **ginie bezpowrotnie**.
- **Nauczysz się wykrywać utratę danych programowo** — przez nasłuch ostrzeżeń (`warnings`) i inspekcję archiwum ZIP tak, jak robi to openpyxl od środka.
- **Opanujesz wzorzec „scan-before-save"** — inwentarz pliku wejściowego przed zapisem i porównanie po zapisie jako najtańszy test regresji.
- **Zbudujesz `bezpieczny_zapis()`** — zapis atomowy (`os.replace`), backup z timestampem, plik tymczasowy w tym samym katalogu, log audytowy i zachowanie uprawnień.
- **Poznasz cztery wzorce modyfikacji istniejącego pliku** oraz zasadę idempotencji, czyli „uruchom dwa razy, efekt jeden".
- **Zapamiętasz regułę kciuka**, która uchroni Cię przed najdroższym błędem tego kursu — i którą powtórzymy tu trzy razy.

## 1. Intuicja i analogia

### 1.1. Fotokopia i przepisanie na czysto

W module 03 poznałeś fundament, do którego wracamy teraz z pełną siłą:

> **`load_workbook()` składa cały model w pamięci. `save()` serializuje cały model od nowa. openpyxl nie edytuje pliku w miejscu.**

Wyobraź sobie, że masz **odręczny rękopis** — z przypisami na marginesach, przyklejonymi karteczkami, zakładkami, rysunkami na skrawkach i pieczątkami notarialnymi. Ktoś prosi Cię, żebyś „dopisał jedno zdanie".

Masz dwie możliwości:

- **Edytor z funkcją dopisania.** Otwierasz rękopis i dopisujesz zdanie. Wszystko inne zostaje nietknięte — karteczki, pieczątki, rysunki. To jest **edycja w miejscu** i tego **openpyxl nie potrafi**.
- **Skryba, który przepisuje całość na czysto.** Skryba czyta oryginał i **przepisuje wszystko, co rozpozna**. Potem Ty dopisujesz zdanie do **jego kopii**, a on ją oprawia w nowy zeszyt. Rękopis nie jest tknięty — powstaje **nowy** dokument.

openpyxl to ten drugi skryba. I teraz najważniejsze pytanie całego modułu:

> **Co zrobi skryba z rzeczą, której nie rozpoznaje?**

Zrobi dokładnie to, co zrobiłby każdy uczciwy człowiek: **przepisze to, co umie przeczytać, a to, czego nie umie — pominie.** Nie zgubi tego złośliwie. Nie zniszczy. Po prostu **nie ma jak tego przenieść**, bo nie zna tego języka. A ponieważ oprawia **nowy** zeszyt, a nie dopisuje do starego, brak tej rzeczy w nowym zeszycie jest **nieodwracalny**.

To jedno zdanie wyjaśnia całą resztę modułu. Kształty, pola tekstowe, sparkline'y, slicery, makra, fragmenty rozszerzeń — wszystko to są „karteczki i pieczątki w języku, którego skryba nie zna". Dokumentacja openpyxl mówi to wprost:

> **Warning** — *„openpyxl does currently not read all possible items in an Excel file so shapes will be lost from existing files if they are opened and saved with the same name."*

I jeszcze jedno, równie ważne:

> **Warning** — *„This operation will overwrite existing files without warning."*

Czyli gdy powiesz skrybie „zapisz jako `raport.xlsx`", a `raport.xlsx` już istnieje — **skryba go wyrzuci i położy na jego miejsce swój nowy zeszyt**. Bez pytania, bez kopii, bez cofania.

### 1.2. Przeprowadzka — trzy kategorie rzeczy

Skryba jest dobrym obrazem przepisywania, ale do inwentarza lepsza jest **przeprowadzka**.

Wprowadzasz się do nowego mieszkania. Pakujesz wszystko, co rozpoznajesz:

- **Fortepian.** Wiesz, czym jest, wiesz, jak go przenieść. **Bierzesz cały.** → *wartości, formuły, daty, style, formaty liczb.*
- **Szafa wbudowana w ścianę.** Wiesz, że tam jest, ale **nie umiesz jej zdemontować bez zniszczenia ściany**. Zostawiasz — i wpisujesz w protokole „była szafa". → *wykresy: openpyxl je odczyta i odtworzy z własnego modelu, ale nietypowe formatowanie może zostać uproszczone.*
- **Żyrandol, o którym ktoś powiedział „zostaje".** Nie przenosisz. **Nie ma go w nowym mieszkaniu wcale.** → *kształty, pola tekstowe, sparkline'y, slicery, makra, Power Query.*
- **Sejf, do którego nie masz klucza.** Wiesz, że istnieje, ale otwiera go tylko specjalista. → *pliki `.xlsm`: potrzebujesz `keep_vba=True`, a i wtedy makra są tylko przenoszone, a nie edytowalne.*

I jeszcze dwie kategorie, które trzeba dodać do tego obrazu:

- **Rzeczy, które zgubisz po cichu**, bo masz je w rękach, ale nie zauważysz ubytku: *komentarze w scalonych komórkach, fragmenty formatowania rich text (bez `rich_text=True`).*
- **Rzeczy, które przeniesiesz w złej kolejności** — *kolejność serii w wykresie kombinowanym. Przeniesiesz je wszystkie, ale w takiej kolejności, że Excel odmówi otwarcia pliku i zażąda „naprawy".*

### 1.3. Reguła kciuka — i dlaczego jest w tym module trzy razy

Z tych dwóch obrazów wynika jedna praktyczna reguła, którą trzeba **powtórzyć trzy razy**, bo jest najważniejszą rzeczą w module:

> **Jeśli plik był tworzony ręcznie przez człowieka w Excelu i jest bogaty wizualnie → nie przepuszczaj go przez openpyxl.**

Bogaty wizualnie znaczy: ma logo, pola tekstowe, strzałki i strzałki-bloki, sparkline'y, slicery, wykresy z nietypowym formatowaniem, makra, tabele przestawne z pełnym modelem danych, chronione zakresy. To są pliki, o które **ktoś dbał** — i to są pliki, o które najbardziej boli, gdy je zepsujesz.

**Nie chodzi o to, żeby openpyxl nie używać.** Chodzi o to, żeby używać go **na właściwej warstwie**:

1. **Wczytaj dane ze źródła i wygeneruj NOWY plik.** Excel to warstwa prezentacji, nie magazyn prawdy (moduł 26).
2. **Użyj szablonu jako wzorca i zapisuj do INNEGO pliku.** Szablon zostaje nietknięty i służy wielokrotnie.
3. **Oddaj zapis Excelowi.** Na Windows/macOS przez `xlwings` lub COM. Wtedy to prawdziwy Excel przepisuje plik i zachowuje wszystko, bo to jego własny format.
4. **Zapisuj zmiany w ODDZIELNYM pliku i scalaj świadomie.** Wyniki obok oryginału, nigdy „na miejscu".

Zapamiętaj też wersję skróconą, do powtarzania kolegom z zespołu:

> **Nigdy nie nadpisuj pliku wejściowego.**

### 1.4. Pisanie na brudno — i dlaczego zapis ma być atomowy

Skoro zapis polega na „wyrzuceniu starego pliku i położeniu nowego", pojawia się problem: **co, jeśli w połowie zapisu coś pójdzie nie tak?**

Wyobraź sobie, że **poprawiasz umowę zamazanym długopisem** i po piątym zdaniu skończył Ci się tusz. Dokument jest teraz **w połowie stary, w połowie nowy** — i bezużyteczny. Nikt nie potrafi powiedzieć, która wersja jest wiążąca.

Właściwy sposób: **piszesz nową wersję na czysto na osobnej kartce**, a dopiero gdy jest gotowa i sprawdzona — **podmieniasz kartkę w segregatorze**. Podmiana trwa jedną sekundę i nigdy nie ma stanu pośredniego.

Tym właśnie jest **zapis atomowy**: plik tymczasowy w tym samym katalogu, a potem `os.replace(tmp, cel)` — **jedna operacja systemowa**, która podmienia plik „wszystko albo nic". Nigdy nie zostawiasz użytkownika z połowicznie zapisanym raportem.

To jest też odpowiedź na pytanie, które usłyszysz od klienta: *„A jak mi się komputer zawiesi w trakcie?"* Odpowiedź: **albo masz stary plik w całości, albo nowy w całości**. Nigdy hybrydę.

### 1.5. Idempotencja — przycisk windy

**Idempotencja** to słowo, które brzmi groźnie, a znaczy jedną rzecz:

> **Wykonanie tej samej operacji dwa razy daje ten sam efekt, co wykonanie jej raz.**

Analogia: **przycisk wezwania windy**. Naciśniesz go raz — winda przyjedzie. Naciśniesz pięć razy — winda przyjedzie **raz**. Nie przyjadą pięć wind.

W automatyzacji Excela to nie jest teoria, tylko **bezpieczeństwo pracy**. Wyobraź sobie skrypt, który dopisuje dzisiejszą sprzedaż do arkusza. Ktoś z zespołu uruchamia go dwa razy, bo „nie widział, że już poszedł". Jeśli skrypt nie jest idempotentny, **sprzedaż dnia jest w raporcie podwójnie** — a nikt tego nie zauważy, bo suma wygląda „jakoś tak dużo, ale w tym miesiącu są promocje".

Idempotencja to zabezpieczenie, które wygląda na biurokrację, dopóki nie uratuje Ci raportu. W praktyce realizuje się ją **znacznikiem**: arkusz `_meta`, właściwości dokumentu (`wb.properties`), numer uruchomienia, hash danych wejściowych. Skrypt przed dopisaniem sprawdza, czy już to zrobił — i jeśli tak, **mówi o tym i wychodzi**.

### 1.6. Trzy kategorie prawdy dla tego modułu

Zgodnie z obietnicą z modułu 03 — cały moduł jest wokół tej tabeli:

| Kategoria | Co to znaczy | Przykłady | Twoja reakcja |
|---|---|---|---|
| **openpyxl potrafi** | przechodzi bez zmian, możesz na tym polegać | wartości, formuły, daty, style, formaty liczb, tabele, filtry, walidacja standardowa, komentarze, hiperlinki, nazwy zdefiniowane, ustawienia druku, freeze panes, właściwości dokumentu | nic — działa |
| **openpyxl potrafi częściowo** | przechodzi, ale **odbudowane z modelu** — sprawdź okiem | wykresy (formatowanie, kolejność serii), część formatowania warunkowego z Excela, rich text (tylko z `rich_text=True`) | **otwórz plik w Excelu i porównaj z oryginałem** |
| **openpyxl gubi przy zapisie** | **znika bezpowrotnie** z pliku wynikowego | kształty, pola tekstowe, formanty, sparkline'y, slicery, timeline, Power Pivot/Power Query, makra bez `keep_vba`, nieznane rozszerzenia `extLst` | **nie przepuszczaj takiego pliku** — użyj jednej z czterech alternatyw |

I jeszcze **czwarta kategoria, ukryta między wierszami**: rzeczy, o których **openpyxl nie mówi ani słowem**, bo nie ma ostrzeżenia. O jednym z nich — zapisie po `data_only=True` — powiemy wprost w 2.2. O innych — w 2.4.

## 2. Teoria

### 2.1. Co się dzieje przy zapisie — mechanicznie

Warto raz zobaczyć, co openpyxl robi w `save()`, żeby przestać się dziwić utracie danych. `Workbook.save()` przekazuje sterowanie do `save_workbook()`, który:

1. Tworzy **nowy** archiwum ZIP od zera.
2. Uruchamia na modelu **zestaw writerów**, każdy dla swojej części: `workbook.xml`, `styles.xml`, po jednym `worksheetN.xml` na arkusz, `sharedStrings.xml`, `theme1.xml`, części rysunków, wykresów, tabel, komentarzy.
3. Każdy writer **serializuje obiekt ze swojego modelu** — czyli pisze to, co jest **w modelu**, a nie to, co było w pliku źródłowym.
4. Dopisuje części „pamięciowe”, jeśli model je trzyma: `vba_archive` (makra) przy `keep_vba=True`, `xl/media/*` (obrazy), `printerSettings`.
5. Zamyka archiwum.

Wniosek, który trzeba wyciągnąć **raz i na zawsze**:

> **Plik wynikowy jest funkcją modelu openpyxl, nie pliku wejściowego.**

Jeżeli coś nie ma reprezentacji w modelu, to nie ma jak trafić do kroku 3 ani 4. To nie jest błąd do naprawienia — to **konsekwencja architektury**. Rozmowa społeczności openpyxl ujmuje to trafnie:

> *„OpenPyXL uses the second approach. Thus the resulting Office Open XML Excel-file only contains parts, the OpenPyXL API provides methods for. That's why the warnings."*

Warto też wiedzieć, że openpyxl **nie egzekwuje rozszerzenia pliku**. W źródle `Workbook` jest właściwość `mime_type` z komentarzem:

> *„Excel requires the file extension to match but openpyxl does not enforce this."*

Czyli openpyxl zapisze treść skoroszytu pod dowolną nazwą. Excel — nie. To źródło dwóch klasycznych pułapek: `.xlsm` zapisany bez makr i `.xlsx` z makrami, o których powiemy w sekcji 6.

### 2.2. `load_workbook()` w pełni — sześć flag i ich ciemne strony

Pełna sygnatura z 3.1.5:

```python
# openpyxl/reader/excel.py
def load_workbook(filename, read_only=False, keep_vba=KEEP_VBA,
                  data_only=False, keep_links=True, rich_text=False):
    """Open the given filename and return the workbook

    :param filename: the path to open or a file-like object
    :type filename: string or a file-like object open in binary mode c.f., :class:`zipfile.ZipFile`
    :param read_only: optimised for reading, content cannot be edited
    :type read_only: bool
    :param keep_vba: preserve vba content (this does NOT mean you can use it)
    :type keep_vba: bool
    :param data_only: controls whether cells with formulae have either the
                      formula (default) or the value stored the last time
                      Excel read the sheet
    :type data_only: bool
    :param keep_links: whether links to external workbooks should be preserved.
                       The default is True
    :type keep_links: bool
    :param rich_text: if set to True openpyxl will preserve any rich text
                      formatting in cells. The default is False
    :type rich_text: bool
    """
    reader = ExcelReader(filename, read_only, keep_vba,
                         data_only, keep_links, rich_text)
    reader.read()
    return reader.wb
```

I teraz tabela, która jest **esencją tego modułu** — bo każda z tych flag z osobna potrafi zniszczyć dane:

| Flaga | Co robi | Kiedy używać | **Co zepsuje, jeśli użyjesz bezmyślnie** |
|---|---|---|---|
| `filename` | ścieżka `str`/`Path` **albo** obiekt plikopodobny binarny (`BytesIO`) | zawsze | otwarcie tekstowe albo `StringIO` → `BadZipFile`; ścieżka bez katalogu → brak pliku |
| `read_only=False` | pełny model w pamięci, zapis możliwy | **domyślnie**, i zawsze, gdy chcesz zapisywać | — |
| `read_only=True` | tryb leniwy, „przez szczelinę", minimalna pamięć | tylko odczyt ogromnych plików | **brak wykresów, obrazów, walidacji, formatowania warunkowego, komentarzy i ochrony w modelu**; `save()` → `TypeError("Workbook is read-only")` |
| `keep_vba=False` | **nie** przenosi `vbaProject.bin` | — | **przy pliku `.xlsm` makra znikają**; Excel odmawia otwarcia albo zgłasza uszkodzenie. **Zawsze `True` dla `.xlsm`/`.xltm`** |
| `data_only=False` | formuły zostają formułami | **domyślnie**, i zawsze, gdy chcesz zapisywać | — |
| `data_only=True` | formuła → **ostatnia wartość zapisana przez Excela** (albo `None`, jeśli cache nie ma) | tylko odczyt wyników, gdy nie zapisujesz | **`save()` po `data_only=True` ZAMIENIA FORMUŁY NA WARTOŚCI.** Model nie trzyma już formuł, więc writer nie ma czego zapisać. To najczęstsza katastrofa tego modułu |
| `keep_links=True` | zachowuje odwołania zewnętrzne **i ich cache** | pliki z odwołaniami zewnętrznymi | — |
| `keep_links=False` | ucina odwołania zewnętrzne | gdy cache jest ogromny i chcesz szybciej | **użytkownik zobaczy `#REF!`** — formuły odwołujące się do innych skoroszytów nie mają już do czego się odwołać |
| `rich_text=False` | tekst z komórki czytany jako **zwykły napis** | gdy nie dbasz o formatowanie fragmentów | **ginie formatowanie rich text** (pogrubienia, kursywy części napisu), a przy zapisie fragmenty zostają spłaszczone |
| `rich_text=True` | zachowuje formatowanie fragmentów (`CellRichText`) | pliki z tekstem częściowo formatowanym | nowość 3.1 — **sprawdź w swojej wersji**, bo to stosunkowo nowa ścieżka kodu |

Zatrzymajmy się na dwie rzeczy, które ludzie mylą najczęściej.

**`keep_links` w źródle.** W `WorkbookParser.__init__` jest `keep_links=True` domyślnie, a w `parse()`:

```python
# external links contain cached worksheets and can be very big
if not self.keep_links:
    package.externalReferences = []
```

Komentarz w kodzie mówi wprost, po co ta flaga istnieje: **odwołania zewnętrzne zawierają w cache kopie tamtych skoroszytów i mogą być bardzo duże.** Czyli `keep_links=False` to optymalizacja pamięci, a **nie** „nie używaj odwołań zewnętrznych". Jeśli wyłączysz je w pliku, który takich odwołań potrzebuje, to **fizycznie usuniesz części `xl/externalLinks/*`** — i formuły odnoszące się do innych plików zostaną bez celu.

**`rich_text`.** To nowość z 3.1 (PR409). Wcześniej openpyxl czytał tekst komórki jako jeden napis i **całkowicie gubił** informację, że część napisu była pogrubiona. Teraz możesz poprosić o zachowanie tego. Uczciwe zastrzeżenie: to stosunkowo nowa funkcjonalność, więc **sprawdź na własnym pliku**, czy zachowanie jest zgodne z oczekiwaniem.

**Jeszcze jedno — o kontekst managerze, bo to częsta nieprawda w sieci.** W briefie tego modułu zasugerowano, że nowsze wersje mają wsparcie dla `with load_workbook(...) as wb:`. Sprawdziłem: **w 3.1.x go NIE MA.** W źródle klasy `Workbook` nie ma ani `__enter__`, ani `__exit__`. Jedyne, co jest:

```python
def close(self):
    """
    Close workbook file if open. Only affects read-only and write-only modes.
    """
    if hasattr(self, '_archive'):
        self._archive.close()
```

Konsekwencja: `with load_workbook(...) as wb:` → `AttributeError: __enter__` (w nowszych Pythonach komunikat bywa inny, ale efekt ten sam). W tym kursie używamy więc **`try/finally`** — i to jest poprawne rozwiązanie, bo działa w każdej wersji:

```python
wb = load_workbook(path, read_only=True)
try:
    ...
finally:
    wb.close()          # istotne TYLKO w read_only / write_only
```

I jeszcze doprecyzowanie, które niejednego zdziwi: **`wb.close()` w trybie normalnym nic nie robi pożytecznego.** Docstring mówi „Only affects read-only and write-only modes". W trybie normalnym model siedzi w pamięci, a plik jest już zamknięty po wczytaniu. Zamykanie go to dobry nawyk (i wymóg w `read_only`), ale **nie myl tego z „zwolnieniem pliku na dysku"** — o blokadach pliku powiemy w sekcji 6.

### 2.3. Inwentarz: co przechodzi bez zmian

Poniżej lista „pewniaków". Możesz na nich polegać — z jednym zastrzeżeniem, które powtórzę: „pewniak" znaczy *„openpyxl to modeluje i zapisuje z modelu"*, a nie *„wygląda identycznie jak w oryginale"*.

**Wartości, formuły, typy, daty.** Tak. Formuła w `sheet1.xml` ma zapisany literał, więc przepisanie jest wierne. Daty przechodzą jako liczby z odpowiednim formatem (moduł 05).

**Style komórek, formaty liczb, style nazwane.** Tak — wszystkie cztery tabele `styles.xml` (`fonts`, `fills`, `borders`, `cellXfs`) plus `numFmts`. W 3.1 naprawiono kilka błędów wokół stylów nazwanych (#1786 „NamedStyles share attributes — mutables gotcha", #1457 „Improved handling of duplicate named styles"), więc **pracuj na 3.1.4+**.

**Tabele (`ListObject`).** Tak. Naprawiono też znany błąd „Table filters are always overridden" (#1156, 3.1.0).

**Filtry i `AutoFilter`.** Tak.

**Walidacja danych — tylko standardowa.** Walidacje zapisane w zwykłym miejscu przechodzą. Te w bloku rozszerzeń `x14` **giną** (moduł 17, sekcja 2.12).

**Formatowanie warunkowe — standardowe tak, część z Excela nie.** Reguły standardowe przechodzą. Ale w 3.1 naprawiano wcześniejsze problemy („ConditionalFormatting lost when pivot table updated", #1858), więc **przy plikach z tabelami przestawnymi i formatowaniem warunkowym zachowaj czujność**.

**Komentarze i hiperlinki.** Tak. Uwaga: komentarz w **scalonej komórce** jest usuwany z ostrzeżeniem (moduł 16). I uwaga druga, praktyczna: w historii był błąd „Hyperlinks duplicated on multiple saves" (#1496, naprawiony w 3.0.5) — czyli jeśli spotkasz w projekcie stare skrypty robiące „save, save, save" na tym samym obiekcie, to jest to dokładnie ten problem.

**Obrazy.** Tak — są kopiowane do `xl/media/` i przenoszone. Ale **nie są kopiowane przez `copy_worksheet`** (moduł 08).

**Nazwy zdefiniowane.** Tak, choć API zmieniło się w 3.1: stary `Workbook.get_named_range` / `add_named_range` / `remove_named_range` został **usunięty**, a weszło `DefinedName` + `wb.defined_names.add(...)` (zmiana #1864 „Better handling of defined names"). Kod z tutoriali sprzed 3.1 tu nie zadziała.

**Ustawienia wydruku, widoki, freeze panes, ukryte wiersze/kolumny, ochrona arkusza i skoroszytu.** Tak (ochrona czytana od dawna; naprawiano „Worksheet protection missing in existing files", #1414).

**Właściwości dokumentu — w tym własne.** Tak; własne właściwości to nowość 3.1 (PR407).

**Wykresy — tu zaczyna się kategoria „częściowo".** openpyxl **odczytuje** wykresy (wsparcie czytania jest od 2.5) i **zapisuje je z powrotem**, ale z **własnego modelu**. Model pokrywa standardowe typy i większość podstawowego formatowania; nietypowe rzeczy mogą zostać uproszczone albo zmienione. W 3.1 zrobiono sporo w tym kierunku („Mapped chartspace graphical properties to charts for advanced formatting", #1666 „Improved normalisation of chart series"), ale to nadal **rekonstrukcja**, nie przeniesienie. Traktuj „przechodzi" jako **„przechodzi — sprawdź okiem"**.

### 2.4. Inwentarz: co ginie lub jest zagrożone

Tu jest druga połowa modułu — lista rzeczy, których openpyxl **nie ma jak** przenieść.

**Kształty, pola tekstowe, strzałki, formanty.** Dokumentacja ostrzega wprost o „shapes". Wszystko, co w Excelu nazywa się „Insert → Shapes / Text Box", nie ma reprezentacji w modelu openpyxl. Jeżeli Twój szablon ma ładne strzałki łączące KPI, **po przepuszczeniu przez openpyxl ich nie będzie**.

**Sparkline'y.** Wpisane w rozszerzenie `x14`. openpyxl ostrzega przy wczytaniu:

```text
UserWarning: Sparkline Group extension is not supported and will be removed
```

i **gubi je przy zapisie**. Nie ma parametru, który to zmieni — openpyxl nie ma klas sparkline'ów w ogóle:

```python
>>> import pkgutil, openpyxl
>>> [m.name for m in pkgutil.walk_packages(openpyxl.__path__, prefix="openpyxl.")
...  if "spark" in m.name]
[]
```

**Slicery i timeline'y.** Wpisane w rozszerzenia `x14`. Ostrzeżenie:

```text
UserWarning: Slicer List extension is not supported and will be removed
```

**Model danych: Power Pivot, Power Query, tabele przestawne z modelem.** Części `xl/pivotCache/`, `xl/pivotTables/`, definicje połączeń i zapytań — poza modelem albo w modelu częściowym.

**Makra VBA.** Tylko z `keep_vba=True`. I nawet wtedy: **są przenoszone, ale nieedytowalne**. Docstring jest tu zaskakująco szczery: `"preserve vba content (this does NOT mean you can use it)"`.

**Nieznane rozszerzenia `extLst`.** To najszersza i najciekawsza kategoria. openpyxl ma w readerze arkuszów funkcję `parse_extensions`, która jest źródłem wszystkich ostrzeżeń tego modułu:

```python
# openpyxl/worksheet/_reader.py
def parse_extensions(self, element):
    extLst = ExtensionList.from_tree(element)
    for e in extLst.ext:
        ext_type = EXT_TYPES.get(e.uri.upper(), "Unknown")
        msg = "{0} extension is not supported and will be removed".format(ext_type)
        warn(msg)
```

Czyli: **każdy** blok `<extLst>` w arkuszu, którego openpyxl nie zna, jest **wypisywany jako ostrzeżenie i wyrzucany**. Lista znanych nazw (`EXT_TYPES`) obejmuje między innymi: *Conditional Formatting*, *Data Validation*, *Sparkline Group*, *Slicer List*, *Protected Range*, *Ignored Error*, *Web Extension*, *Timeline Ref*. Wszystko spoza tej listy trafi do komunikatu jako **„Unknown extension is not supported and will be removed"** — i to jest jedno z najważniejszych ostrzeżeń, na jakie możesz trafić, bo znaczy „była tu rzecz, o której nie wiem nic".

**Kolejność serii w wykresach kombinowanych — przypadek szczególny.** To jedyna pozycja na tej liście, która **przenosi się w całości, ale psuje plik**. W OOXML każda seria wykresu ma elementy `<c:idx>` i `<c:order>`. Excel numeruje serie globalnie w obrębie wykresu — także wtedy, gdy wykres jest kombinacją dwóch typów (np. słupki + linia). openpyxl normalizuje serie **per typ**, więc po odczytaniu i zapisie kombinacji serie drugiego typu mogą zaczynać się znowu od zera. Excel wykrywa duplikaty i zgłasza uszkodzenie pliku:

```text
We found a problem with some content in 'raport.xlsx'. Do you want us to
try to recover as much as we can? If you trust the source of this workbook,
click Yes.
```

openpyxl ma na to lekarstwo: **prywatną metodę `_reindex()`**, która ustawia `order` kolejno od zera:

```python
# openpyxl/chart/tests/test_chart.py — zachowanie _reindex()
orders = [40, 20, 34, 11]
for o, s in zip(orders, chart.series):
    s.order = o
# ... po wywołaniu chart._reindex():
# reordered == [0, 1, 2, 3]
```

To jest **private API** (podkreślnik!), więc stosuj je świadomie i po weryfikacji w swojej wersji. Ale sama wiedza — *„kombinacje wykresów to najczęstszy powód komunikatu o uszkodzonym pliku po round-tripie"* — jest bezcenna przy diagnozie.

### 2.5. Reguła kciuka i cztery alternatywy — wariant operacyjny

Wróćmy do reguły z 1.3, tym razem z konkretnym testem decyzyjnym. Zanim przepuścisz cudzy plik przez openpyxl, odpowiedz na trzy pytania:

1. **Czy plik ma `xl/media/`, `xl/charts/`, `xl/drawings/vmlDrawing*.vml`, `xl/vbaProject.bin`, `customXml/`, albo bloki `extLst` w arkuszach?** Jeśli tak — **ryzyko jest realne**, a nie teoretyczne.
2. **Czy w pliku są rzeczy, których nie umiesz wymienić z pamięci?** Jeśli nie wiesz, co jest w środku, **nie masz prawa go nadpisać**.
3. **Czy wynik ma zastąpić oryginał, czy być nowym plikiem?** Jeśli ma zastąpić — **musisz** mieć kopię i porównanie przed/po.

Jeżeli którykolwiek sygnał jest niepokojący — wybierz jedną z **czterech alternatyw** z 1.3. W praktyce najczęściej wybierana jest druga i to ona będzie tematem modułu 23 (wzorzec Prototype):

> **Szablon nigdy nie jest zapisywany. Szablon jest czytany, wypełniany i zapisywany pod nową nazwą.**

### 2.6. Wzorzec „scan-before-save" — zajrzyj do skrytki przed zapisem

Skoro utrata danych dzieje się **przy zapisie**, to momentem, w którym trzeba reagować, jest **przed** zapisem. I kluczowe: **nie musisz zgadywać, co jest w pliku** — możesz to **policzyć**.

Plik `.xlsx` to **ZIP** (moduł 02). Inspekcja to kilka linii:

```python
import zipfile
with zipfile.ZipFile("raport.xlsx") as z:
    for name in z.namelist():
        print(name)
```

Zobaczysz dokładnie te części, o których mówi inwentarz z 2.4: `xl/charts/chart1.xml`, `xl/media/image1.png`, `xl/drawings/vmlDrawing1.vml`, `xl/vbaProject.bin`, `customXml/item1.xml`, `xl/pivotCache/pivotCacheDefinition1.xml`, `xl/printerSettings/printerSettings1.bin`.

Wzorzec **scan-before-save** w pełnej postaci:

1. **Zeskanuj plik wejściowy** — wypisz części i wykryj rozszerzenia `extLst`.
2. **Zestaw z inwentarzem ryzyk** — każda część z sekcji 2.4 to **czerwona flaga**.
3. **Zapisz** (najlepiej do nowego pliku).
4. **Zeskanuj plik wyjściowy** tą samą funkcją.
5. **Porównaj** — różnica to Twoja utrata danych, wypisana w liczbach.
6. **Otwórz plik w Excelu i obejrzyj go okiem.** To nie jest opcja. To jest **obowiązek**. I mówimy to w tym module **trzy razy**: tutaj, w 2.10 i w sekcji 7.

Krok 4–5 jest ważniejszy, niż wygląda, bo **łapie rzeczy, o których openpyxl nie ostrzega**. Wykresy: nie ma ostrzeżenia o zmianach formatowania — jest tylko inny `chart1.xml`. Komentarze w scalonych komórkach: ostrzeżenie bywa jedno i łatwo je przeoczyć. Pivot cache: ostrzeżenie nie pojawia się wcale. **Porównanie liczbowe dwóch inwentarzy to jedyne narzędzie, które to złapie.**

### 2.7. Atomowy i bezpieczny zapis — wzorzec produkcyjny

Zbierzmy wszystkie zasady w jedną sekwencję, którą potem zaimplementujesz w `bezpieczny_zapis()`.

**Zasada 1 — nigdy nie nadpisuj wejścia.** Nowa nazwa: `raport_v2.xlsx`, `raport_2026-10-08.xlsx`, albo przynajmniej kopia z timestampem obok.

**Zasada 2 — plik tymczasowy w tym samym katalogu.** Dlaczego **w tym samym**? Bo `os.replace` jest atomowy **tylko w obrębie tego samego systemu plików**. Próba atomowej podmiany między dyskami albo udziałami sieciowymi kończy się kopiowaniem i utratą atomowości. Twój plik tymczasowy musi leżeć obok celu.

```python
import os
import tempfile
from pathlib import Path

def _plik_tymczasowy(katalog: Path, przyrostek: str) -> Path:
    """Tworzy nazwę pliku tymczasowego w danym katalogu (bez otwierania go)."""
    fd, name = tempfile.mkstemp(dir=katalog, prefix=".tmp-raport-",
                               suffix=przyrostek)
    os.close(fd)              # zamykamy deskryptor - otworzy go openpyxl
    return Path(name)
```

Zwróć uwagę na dwie rzeczy. Pierwsza: `tempfile.mkstemp` **tworzy** plik i zwraca deskryptor — trzeba go zamknąć, żeby openpyxl mógł pisać. Druga: **losowa nazwa** pliku tymczasowego to nie ozdoba. Dopóki plik leży obok, musi mieć nazwę, której nikt nie pomyli z prawdziwym raportem, i której nie da się wcześniej „przewidzieć".

**Zasada 3 — podmiana atomowa.** `os.replace(src, dst)`. Nie `os.rename` (na Windows różnica jest istotna: `rename` zawodzi, gdy cel istnieje, `replace` podmienia). Nie `shutil.move` (nieatomowy przy przekraczaniu systemów plików).

**Zasada 4 — backup z timestampem.** `shutil.copy2` a nie `shutil.copy`, bo `copy2` **zachowuje metadane** — czas modyfikacji i uprawnienia. Dla pliku raportu czas modyfikacji bywa istotny (audyt, „kiedy powstała ta wersja").

**Zasada 5 — zachowanie uprawnień.** `os.replace` wpisze na miejsce celu plik z **uprawnieniami pliku tymczasowego**, nie starymi uprawnieniami celu. Jeśli raport miał ograniczone prawa dostępu, musisz je przenieść **ręcznie**:

```python
stare = os.stat(cel)                 # jeśli cel istnieje
os.chmod(tmp, stare.st_mode)         # przed podmianą
```

**Zasada 6 — sprzątanie w każdym scenariuszu.** Jeśli `wb.save(tmp)` rzuci wyjątek, plik tymczasowy musi zostać **usunięty**, a wyjątek **propagowany dalej**. Nigdy nie połykaj wyjątku po cichu — użytkownik musi wiedzieć, że raport nie powstał.

**Zasada 7 — zamknięcie.** `wb.close()` w `finally`. W trybie normalnym to formalność, ale w `read_only` to **wymóg** — bez tego zostawiasz otwarte uchwyty na ZIP-a. Historia błędów w tym obszarze jest bogata: #2122 „File handlers not always released in read-only mode" i #2149 „Workbook files not properly closed on Python ≥ 3.11.8 and Windows" — oba naprawione w 3.1.3, oba dotyczyły właśnie tego, o czym mówimy.

**Zasada 8 — log audytowy.** Jeden wiersz JSON na operację: kiedy, skąd, dokąd, ile arkuszy, ile wierszy, jakie ostrzeżenia. To nie „dodatek dla dorosłych" — to jedyny sposób, żeby po miesiącu odpowiedzieć na pytanie „dlaczego raport z 12 października ma o 40 wierszy mniej".

### 2.8. Cztery wzorce modyfikacji istniejącego pliku

To wzorce, które w module 26 nazwiesz „strategiami dostępu do danych". Każdy ma swoje zastosowanie i swoje ryzyko.

**1. Read-modify-write.** Wczytaj → zmień → zapisz. Najprostszy, najczęstszy i **naj bardziej ryzykowny**.

```python
wb = load_workbook("plik.xlsx")
wb["Dane"]["A1"] = "Nowe"
wb.save("plik.xlsx")          # ⚠️ nadpisanie oryginału!
```

Kiedy wolno: gdy wiesz, że plik jest **„czysty technicznie"** — powstał z Twojego własnego generatora albo jest prostą tabelą danych bez warstwy wizualnej. Gdy **nie wolno**: dla każdego pliku „od człowieka".

**2. Fill a template.** Szablon + nowy plik wynikowy. **Zalecany domyślnie.**

```python
wb = load_workbook("templates/raport_wzor.xlsx")
wb["Dane"]["A1"] = "Nowe"
wb.save(f"output/raport_{data}.xlsx")     # szablon nietknięty
```

Zysk: szablon jest **jednym źródłem decyzji** o wyglądzie (jak `NamedStyle` z modułu 09), a każdy wynik jest osobny i możliwy do porównania. Ograniczenie: szablon z sparkline'ami i makrami nadal traci (moduł 23 pokaże, jak to obejść).

**3. Append-only log.** Dopisywanie do arkusza logu — nigdy nie modyfikujesz istniejących wierszy, tylko dodajesz na końcu. Świetne dla audytu i dla danych przyrostowych. Wymaga idempotencji (2.9), bo „dopisz" bez sprawdzenia oznacza „dopisz podwójnie".

**4. Sidecar.** Wyniki dopisywane do **nowego arkusza w kopii** pliku.

```python
wb = load_workbook("raport_klienta.xlsx")
ws = wb.create_sheet("Wyniki_2026-10-08")     # nowy arkusz, dane nietknięte
...
wb.save("raport_klienta_wyniki.xlsx")         # nowy plik
```

Zaleta: oryginalne arkusze **nie są w ogóle dotykane**, więc nawet gdyby coś w nich zginęło przy przepisywaniu, Twoja praca jest oddzielona. Wada: nadal przepisujesz cały plik, więc **utrata obejmuje oryginał w pliku wynikowym** — sidecar chroni Twoje dane, nie cudzą warstwę wizualną.

### 2.9. Idempotencja — jak ją wprowadzić

Trzy praktyczne mechanizmy, od najprostszego:

**A. Znacznik w komórce/arkuszu `_meta`.** Przed dopisaniem sprawdzasz:

```python
def juz_przetworzony(ws, identyfikator: str) -> bool:
    """Czy dany identyfikator (np. hash danych) jest już w arkuszu meta?"""
    for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
        if row[0] == identyfikator:
            return True
    return False
```

**B. Właściwości dokumentu.** `wb.properties.description` albo własna właściwość: *„wygenerowano z <hash danych wejściowych> dnia <data>"*. Zaleta: informacja jedzie z plikiem i jest widoczna dla użytkownika bez otwierania arkusza meta (moduł 04).

**C. Nazwany numer uruchomienia / wersji.** Arkusz `_meta` z kolumnami: `run_id`, `timestamp`, `hash_wejscia`, `wersja_kodu`.

I najważniejsza uwaga o samym identyfikatorze: **musi pochodzić z danych, nie z zegara**. Jeśli jako identyfikator użyjesz `datetime.now()`, to każdy kolejny run będzie „nowy" i idempotencja przestanie działać. Hash danych wejściowych (`hashlib.sha256`) jest właściwym wyborem:

```python
import hashlib

def odcisk_danych(dane: str) -> str:
    return hashlib.sha256(dane.encode("utf-8")).hexdigest()[:16]
```

### 2.10. Test regresji formatowania

Skoro utrata danych jest cicha, potrzebujesz **miary**. Nie „otworzyłem i wydaje się ok", a **liczby**. Wzorzec:

1. **`podsumowanie_skoroszytu(path)`** — funkcja zwracająca słownik: liczba arkuszy, liczba komórek z wartościami, liczba formuł, liczba wpisów w `styles.xml`, liczba reguł formatowania warunkowego, liczba tabel, liczba nazw zdefiniowanych, **lista części ZIP** (`xl/media/*`, `xl/charts/*`, `xl/vbaProject.bin`, `customXml/*`), **lista rozszerzeń `extLst`**.
2. **Uruchom na oryginale**, zapisz w JSON.
3. **Uruchom na wyniku**, zapisz w JSON.
4. **Porównaj** — różnice wypisz jako raport.

To jest **ten sam pomysł co nasłuch ostrzeżeń** (2.6), tylko mocniejszy: ostrzeżenia mówią Ci o tym, o czym openpyxl **wie**. Porównanie inwentarzy mówi Ci także o tym, **o czym nie wie**.

Ten test ma jeszcze jedną zaletę, o której powiemy w module 20: **jest testem regresji.** Raz nagrany „odcisk" dobrego pliku działa jak alarm — gdy zmiana w kodzie zacznie gubić część pliku, dowiesz się o tym **przed** wysłaniem raportu do klienta.

### 2.11. Weryfikacja „na oko" — powiedziane wprost

**Trzeci raz** i tym razem jako reguła: **każdy plik, który przeszedł przez openpyxl po wczytaniu istniejącego, MUSI zostać otwarty w Excelu i porównany z oryginałem — zanim trafi do kogokolwiek.**

Nie chodzi o to, że automatyzacja jest zawodna. Chodzi o to, że **automatyzacja potrafi być cicho zawodna**, a Ty jesteś ostatnią osobą, która może to złapać przed klientem. Benchmarki openpyxl pokazują to bezlitośnie: dla tego samego pliku (walidacja z `x14` + sparkline, jedna zmiana w `A1`, zapis) openpyxl 3.1.5 zapisał plik **bez ani walidacji, ani sparkline'a** — z dwoma ostrzeżeniami, które łatwo przeoczyć.

## 3. Przykłady krok po kroku

### Przykład 1 — „read-modify-write" z pełnym porównaniem (🟢)

Najpierw najprostsza rzecz, ale zrobiona **poprawnie**, czyli z kopiami i porównaniem.

```python
"""Modul 18 - read-modify-write z porownaniem przed/po.

Uruchom:  python examples/18_read_modify_write.py
"""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
DATA.mkdir(parents=True, exist_ok=True)
OUTPUT.mkdir(parents=True, exist_ok=True)

ZRODLO = DATA / "18_dane_wejsciowe.xlsx"


def przygotuj_zrodlo() -> None:
    """Tworzy plik wejsciowy, jesli go nie ma. Czysciutky, bez warstwy wizualnej."""
    if ZRODLO.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaż"
    ws.append(["Data", "Region", "Kwota"])
    ws.append(["2026-09-01", "Północ", 1200.50])
    ws.append(["2026-09-02", "Południe", 980.00])
    ws.append(["2026-09-03", "Północ", 1540.75])
    wb.save(ZRODLO)
    wb.close()
    print(f"Utworzono plik testowy: {ZRODLO}")


def main() -> None:
    przygotuj_zrodlo()

    znacznik = datetime.now().strftime("%Y%m%d-%H%M%S")
    kopia_oryginalu = OUTPUT / f"18_oryginal_kopia_{znacznik}.xlsx"
    wynik = OUTPUT / f"18_wynik_{znacznik}.xlsx"

    # KROK 1: kopia oryginalu (shutil.copy2 zachowuje metadane)
    shutil.copy2(ZRODLO, kopia_oryginalu)
    print(f"1. Kopia oryginalu: {kopia_oryginalu.name}")

    # KROK 2: wczytaj - TRYBY DOMYSLNE (nie read_only, nie data_only!)
    wb = load_workbook(ZRODLO)          # formuly i styl zostają
    try:
        ws = wb["Sprzedaż"]
        print(f"2. Wczytano. Arkusze: {wb.sheetnames}")
        print(f"   max_row={ws.max_row}  max_column={ws.max_column}")

        # KROK 3: modyfikacja - dopisz kolumne z marżą (formuła!)
        naglowek_col = ws.max_column + 1
        ws.cell(row=1, column=naglowek_col, value="Marża 20%")
        for row in range(2, ws.max_row + 1):
            kwota = ws.cell(row=row, column=3).coordinate
            ws.cell(row=row, column=naglowek_col,
                    value=f"=ROUND({kwota}*0.2,2)")
        print(f"3. Dopisano kolumnę {ws.cell(row=1, column=naglowek_col).column_letter}")

        # KROK 4: zapis do NOWEGO pliku - oryginał nietknięty
        wb.save(wynik)
        print(f"4. Zapisano wynik: {wynik.name}")
    finally:
        wb.close()          # dobry nawyk; w trybie normalnym to formalność

    # KROK 5: porownanie - czytamy oba pliki i wypisujemy roznice
    porownaj(kopia_oryginalu, wynik)

    print()
    print("=" * 76)
    print("CO SPRAWDZIĆ TERAZ W EXCELU (weryfikacja 'na oko' - obowiązkowa!)")
    print("=" * 76)
    print(f"  Otwórz {kopia_oryginalu.name} oraz {wynik.name}")
    print("  Oczekiwane różnice: nowa kolumna D z formułami.")
    print("  Oczekiwane BRAKI różnic: identyczne nagłówki, identyczne dane")
    print("  w kolumnach A-C, identyczne szerokości kolumn.")
    print("  Jeśli w kolumnie D widzisz 240.1 zamiast formuły =ROUND(...),")
    print("  to znaczy, że wczytałeś plik z data_only=True. Błąd krytyczny!")


def porownaj(a: Path, b: Path) -> None:
    """Wypisuje roznice w modelu miedzy dwoma plikami."""
    wa = load_workbook(a, read_only=True)
    wb_ = load_workbook(b, read_only=True)
    try:
        print()
        print("=" * 76)
        print(f"PORÓWNANIE MODELU: {a.name}  ->  {b.name}")
        print("=" * 76)
        for nazwa_arkusza in wa.sheetnames:
            if nazwa_arkusza not in wb_.sheetnames:
                print(f"  ⚠️  arkusz zniknął: {nazwa_arkusza}")
                continue
            s1, s2 = wa[nazwa_arkusza], wb_[nazwa_arkusza]
            print(f"  arkusz {nazwa_arkusza!r}:")
            print(f"    wymiary: {s1.calculate_dimension()} -> "
                  f"{s2.calculate_dimension()}")
            for row in s1.iter_rows(values_only=True):
                print(f"    ORG  {row}")
                break
            for row in s2.iter_rows(values_only=True):
                print(f"    WYN  {row}")
                break
        nowe = set(wb_.sheetnames) - set(wa.sheetnames)
        if nowe:
            print(f"  nowe arkusze: {sorted(nowe)}")
    finally:
        wa.close()
        wb_.close()


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `load_workbook(ZRODLO)` bez flag wczytuje **pełny model**: wartości, formuły, style każdej komórki, szerokości kolumn, tabele. Wywołanie `ws.cell(...)` tworzy **nową komórkę** — także poza pierwotnym zasięgiem danych. Nic z tego nie dotyka dysku.

**Co trafi do pliku.** Dopiero `wb.save(wynik)` **buduje nowe archiwum** i zapisuje w nim wszystkie części: `xl/workbook.xml`, `xl/styles.xml`, `xl/worksheets/sheet1.xml`, `xl/sharedStrings.xml`. Kolumna D z formułami trafi jako tekst do `sheet1.xml` — **wynik zostanie obliczony przez Excela przy otwarciu** (moduł 06).

**Trzy rzeczy do przemyślenia:**

1. **Kopia powstała przez `shutil.copy2`, nie przez `load_workbook` + `save`.** To nie kosmetyka: gdybyś „zrobił kopię" przez openpyxl, skopiowałbyś **model**, a więc już okrojoną wersję. Kopia ma być **bajt w bajt**, i to zapewnia tylko `shutil.copy2`.
2. **Zapisaliśmy do nowego pliku.** Wariant `wb.save(ZRODLO)` byłby technicznie możliwy (openpyxl nadpisałby plik) — i mimo że tu nie byłoby straty (nasz plik jest prosty), **zasada jest jedna: nigdy nie nadpisuj wejścia**.
3. **Porównanie w `porownaj()` działa na `read_only=True` — i tu jest to poprawne.** Bo porównujemy **dane**, nie warstwę wizualną. Ale w `podsumowanie_skoroszytu` z 2.10, gdzie liczymy style i reguły, **`read_only` jest zabronione** — bo gubi walidacje, formatowanie warunkowe i obrazy z modelu.

### Przykład 2 — inwentarz ZIP: co w pliku jest, czego openpyxl nie zna (🟡)

To jest implementacja **scan-before-save** z 2.6. Uruchom ją na **każdym** pliku, który zamierzasz przepuścić przez openpyxl.

```python
"""Modul 18 - inwentarz archiwum: czesci i rozszerzenia extLst.

Dziala bez Excela. Uruchom:
    python examples/18_inwentarz_zip.py ścieżka.xlsx
albo bez argumentu - na pliku testowym z przykladu 1.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOMYSLNY = ROOT / "data" / "18_dane_wejsciowe.xlsx"

# Regex na bloki rozszerzen w XML arkusza.
RE_EXT_URI = re.compile(r'<ext\s+uri="\{?([0-9A-Fa-f\-]+)\}?"')

# Znane identyfikatory rozszerzen. Wartosci sa identyfikatorami GUID OOXML.
# Zrodlo "prawdy" w bibliotece to openpyxl.worksheet._reader.EXT_TYPES
# (API prywatne) - dlatego najpierw probujemy je zaimportowac.
KNOWN_EXT: dict[str, str] = {
    "CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF": "Data Validation",
    "05C60535-1F16-4FD2-B633-F4F36F0B64E0": "Sparkline Group",
    "A8765BA9-456A-4DAB-B4F3-ACF838C121DE": "Slicer List",
    "7E03D99C-DC04-49D9-9315-930204A7B6E9": "Timeline Ref",
    "2E9D1231-5D1C-4C77-BB3B-5DB5A2C50B3C": "Protected Range",
    "01252116-DBEB-4CFA-B13A-4E4CA9B9C7E8": "Ignored Error",
}

def _wczytaj_znane() -> dict[str, str]:
    """Uzupelnia slownik tym, co wie zainstalowana wersja openpyxl."""
    try:
        from openpyxl.worksheet._reader import EXT_TYPES   # API prywatne
    except Exception:
        return KNOWN_EXT
    wynik = dict(KNOWN_EXT)
    for uri, nazwa in EXT_TYPES.items():
        wynik[uri.upper().strip("{}")] = nazwa
    return wynik

# Czesci, ktore sa sygnalami ostrzegawczymi (inwentarz ryzyk z sekcji 2.4).
RYZYKO = {
    "xl/vbaProject.bin":            "makra VBA (potrzebne keep_vba=True)",
    "xl/charts/":                   "wykresy (odbudowywane z modelu - sprawdź okiem)",
    "xl/drawings/vmlDrawing":       "VML (komentarze, formanty, stare kształty)",
    "xl/media/":                    "obrazy (przenoszone, ale nie przez copy_worksheet)",
    "xl/pivotTables/":              "tabele przestawne",
    "xl/pivotCache/":               "cache tabel przestawnych",
    "customXml/":                   "customXml (własne dane, inteligentne tagi)",
    "xl/printerSettings/":          "ustawienia drukarki",
    "xl/externalLinks/":            "odwołania zewnętrzne (keep_links!)",
}


def inwentarz(sciezka: Path) -> dict:
    """Zwraca slownik z inwentarzem archiwum: czesci + rozszerzenia."""
    znane = _wczytaj_znane()
    czesci: list[str] = []
    rozszerzenia: dict[str, list[str]] = {}       # arkusz -> lista typow

    with zipfile.ZipFile(sciezka) as z:
        czesci = z.namelist()
        for name in czesci:
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            xml = z.read(name).decode("utf-8", errors="replace")
            znalezione = []
            for guid in RE_EXT_URI.findall(xml):
                typ = znane.get(guid.upper(), f"NIEZNANE ({guid})")
                znalezione.append(typ)
            if znalezione:
                rozszerzenia[name] = znalezione

    return {"sciezka": sciezka, "czesci": czesci, "rozszerzenia": rozszerzenia}


def raport(sciezka: Path) -> None:
    inv = inwentarz(sciezka)
    czesci = inv["czesci"]

    print("=" * 78)
    print(f"INWENTARZ ARCHIWUM: {sciezka.name}")
    print(f"  części razem: {len(czesci)}   rozmiar: "
          f"{sciezka.stat().st_size / 1024:.1f} kB")
    print("=" * 78)

    # --- 1. wszystkie czesci, pogrupowane po katalogu ---
    print("\nCZĘŚCI (pogrupowane):")
    grupy: dict[str, list[str]] = {}
    for name in czesci:
        katalog = name.rsplit("/", 1)[0] if "/" in name else "(korzeń)"
        grupy.setdefault(katalog, []).append(name)
    for katalog in sorted(grupy):
        pliki = grupy[katalog]
        if len(pliki) <= 3:
            print(f"  {katalog}/  ->  {', '.join(p.rsplit('/', 1)[-1] for p in pliki)}")
        else:
            print(f"  {katalog}/  ->  {len(pliki)} plików "
                  f"(np. {pliki[0].rsplit('/', 1)[-1]})")

    # --- 2. czesci z listy ryzyka ---
    print("\nSIGNALY OSTRZEGAWCZE (inwentarz ryzyk):")
    znaleziono_ryzyko = False
    for prefiks, opis in RYZYKO.items():
        trafienia = [c for c in czesci if c.startswith(prefiks)]
        if trafienia:
            znaleziono_ryzyko = True
            print(f"  🔴 {prefiks:<26} ({len(trafienia)}×)  {opis}")
    if not znaleziono_ryzyko:
        print("  🟢 brak części z listy ryzyka")

    # --- 3. rozszerzenia extLst w arkuszach ---
    print("\nROZSZERZENIA extLst W ARKUSZACH (to ginie przy zapisie!):")
    if not inv["rozszerzenia"]:
        print("  🟢 brak bloków extLst")
    for arkusz, typy in inv["rozszerzenia"].items():
        print(f"  🔴 {arkusz}:")
        for typ in typy:
            print(f"       - {typ}  ->  ZOSTANIE USUNIĘTE przy save()")

    # --- 4. werdykt ---
    print()
    print("=" * 78)
    ryzykowne = any(c.startswith(p) for p in RYZYKO for c in czesci)
    if ryzykowne or inv["rozszerzenia"]:
        print("WERDYKT: ⚠️  PLIK WYSOKIEGO RYZYKA")
        print("  Nie przepuszczaj go przez openpyxl z zapisem na tym samym pliku.")
        print("  Wybierz jedną z czterech alternatyw (sekcja 2.5):")
        print("    (1) wczytaj dane ze źródła i wygeneruj nowy plik,")
        print("    (2) użyj szablonu jako wzorca i zapisz pod nową nazwą,")
        print("    (3) oddaj zapis Excelowi (xlwings/COM),")
        print("    (4) zapisz zmiany w oddzielnym pliku i scalaj świadomie.")
        print("  Jeśli MUSISZ zapisać: zrób kopię, zapisz do nowego pliku,")
        print("  uruchom ten inwentarz PONOWNIE i PORÓWNAJ wyniki.")
    else:
        print("WERDYKT: 🟢 PLIK NISKIEGO RYZYKA (na podstawie heurystyk)")
        print("  Brak części z listy ryzyk i brak bloków extLst.")
        print("  Uwaga: to nie jest gwarancja! Wykresy i nietypowe części")
        print("  mogą nie dać żadnego sygnału. Weryfikacja 'na oko' nadal")
        print("  obowiązuje.")
    print("=" * 78)


def main() -> None:
    sciezka = Path(sys.argv[1]) if len(sys.argv) > 1 else DOMYSLNY
    if not sciezka.exists():
        print(f"Brak pliku: {sciezka}")
        print("Uruchom najpierw: python examples/18_read_modify_write.py")
        return
    raport(sciezka)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Nic z Excela — tylko `zipfile`. Cała funkcja `inwentarz()` czyta archiwum i wyciąga **nazwy części** oraz — dla arkuszy — **identyfikatory rozszerzeń**. To jest ta sama informacja, którą openpyxl zdobywa w `parse_extensions`, tylko odczytana bez budowania modelu.

**Trzy rzeczy do przemyślenia:**

1. **Import `EXT_TYPES` jest prywatny.** Dlatego zabezpieczyłem go `try/except` i własnym słownikiem. To wzorzec wart zapamiętania na cały kurs: **gdy sięgasz po API prywatne, zawsze miej plan B** i sprawdź w swojej wersji.
2. **Identyfikatory w XML-u są w nawiasach klamrowych, w wielkich literach.** Dlatego `RE_EXT_URI` ma `\{?\}?` i robimy `.upper().strip("{}")`. Bez tego porównanie z `EXT_TYPES` (który kluczuje wielkimi literami) nie trafiałoby nigdy.
3. **Werdykt „🟢 niskiego ryzyka" nie jest gwarancją.** To heurystyka po nazwach części i rozszerzeniach. Ma na końcu ostrzeżenie — i to nie jest grzecznościowy dopisek, a najważniejsze zdanie w wydruku. **Weryfikacja „na oko" pozostaje obowiązkowa.**

### Przykład 3 — nasłuch ostrzeżeń i raport wpływu (🟡)

Ostrzeżenia openpyxl to **jedyny kanał, przez który biblioteka mówi Ci o utracie danych**, zanim do niej dojdzie. Ten przykład zamienia je w czytelny raport — i pokazuje, jak **nie** wyciszyć ich przypadkiem.

```python
"""Modul 18 - nasluch ostrzezen openpyxl przy wczytaniu pliku.

Uruchom:  python examples/18_ostrzezenia.py ścieżka.xlsx
"""

from __future__ import annotations

import sys
import warnings
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
DOMYSLNY = ROOT / "data" / "18_dane_wejsciowe.xlsx"

# Znane ostrzezenia i ich znaczenie dla utraty danych.
ZNACZENIE = {
    "Sparkline Group extension":
        "Sparkline'y znikną. Nie ma API, które je odtworzy.",
    "Data Validation extension":
        "Listy rozwijane z innego arkusza (x14) znikną. "
        "Obejście: nazwy zdefiniowane (moduł 17, sekcja 2.12).",
    "Slicer List extension":
        "Slicery znikną. Excel 2010+ funkcja poza modelem openpyxl.",
    "Timeline Ref extension":
        "Timeline'y znikną.",
    "Ignored Error extension":
        "Zignorowane ostrzeżenia Excela znikną (kosmetyka).",
    "Protected Range extension": "Chronione zakresy znikną.",
    "Unknown extension":
        "openpyxl NIE WIE, co to jest. To najgroźniejsze ostrzeżenie: "
        "najprawdopodobniej zniknie nieznana część pliku.",
    "shapes":
        "Kształty i/lub pola tekstowe znikną (dokumentacja openpyxl).",
}


def rzecz_znaczenie(tresc: str) -> str:
    for klucz, opis in ZNACZENIE.items():
        if klucz.lower() in tresc.lower():
            return opis
    return "Nieznane ostrzeżenie - sprawdź ręcznie, co może zginąć."


def przeanalizuj(sciezka: Path) -> int:
    """Wczytuje plik z nasłuchem ostrzeżeń. Zwraca liczbę ostrzeżeń."""
    print("=" * 82)
    print(f"ANALIZA OSTRZEŻEŃ: {sciezka.name}")
    print("=" * 82)

    with warnings.catch_warnings(record=True) as zlapane:
        # "always" jest KRYTYCZNE: domyślnie Python pokazuje każde
        # unikalne ostrzeżenie tylko RAZ na sesję. Przy pliku z pięcioma
        # arkuszami ze sparkline'ami zobaczysz tylko jedno ostrzeżenie.
        warnings.simplefilter("always")

        wb = load_workbook(sciezka)          # NIE read_only! Ta analiza ma sens
        try:                                  # tylko z pełnym parserem.
            arkusze = wb.sheetnames
            print(f"\nWczytano bez wyjątku. Arkusze: {arkusze}")

            # dodatkowa diagnostyka: co JEST w modelu
            media = charts = tables = reguły = 0
            for ws in wb.worksheets:
                tables += len(ws.tables)
                reguły += len(ws.conditional_formatting)
            print(f"\nW modelu: tabel={tables}  reguł warunkowych={reguły}")
            print("(obrazy i wykresy nie są tu liczone - siedzą w rysunkach)")
        finally:
            wb.close()

    liczba = len(zlapane)
    print(f"\nOSTRZEŻEŃ: {liczba}")
    if not liczba:
        print("  🟢 Brak ostrzeżeń. To dobra wiadomość - ale NIE gwarancja.")
        print("     Niektóre utraty (np. formatowanie wykresów) nie dają")
        print("     żadnego ostrzeżenia. Porównaj inwentarze ZIP (przykład 2).")
    else:
        for i, w in enumerate(zlapane, 1):
            print(f"\n  [{i}] {w.category.__name__}  (plik: "
                  f"{Path(w.filename).name}, linia {w.lineno})")
            print(f"      {w.message}")
            print(f"      → {rzecz_znaczenie(str(w.message))}")

    print()
    print("=" * 82)
    print("JAK NIE WYCISZYĆ TEGO PRZYPADKIEM")
    print("=" * 82)
    print("  ❌ ZLE - wyciszy WSZYSTKO, także ostrzeżenia o błędach połączeń:")
    print("       warnings.simplefilter('ignore')")
    print("  ❌ ZLE - wyciszy cały moduł openpyxl, więc i przyszłe ostrzeżenia:")
    print("       warnings.filterwarnings('ignore', module='openpyxl')")
    print()
    print("  ✅ DOBRZE - wycisz JEDNO ostrzeżenie, tylko na czas wczytania,")
    print("     i tylko wtedy, gdy plik jest tylko DO ODCZYTU:")
    print("       with warnings.catch_warnings():")
    print("           warnings.filterwarnings(")
    print("               'ignore',")
    print("               message='Sparkline Group extension',")
    print("               category=UserWarning,")
    print("               module='openpyxl',)")
    print("           wb = load_workbook(path)")
    print()
    print("     Uwaga: jeśli PO wczytaniu zapiszesz plik, sparkline'y znikną")
    print("     mimo wyciszenia. Wyciszenie ukrywa komunikat, nie efekt.")
    return liczba


def main() -> None:
    sciezka = Path(sys.argv[1]) if len(sys.argv) > 1 else DOMYSLNY
    if not sciezka.exists():
        print(f"Brak pliku: {sciezka}")
        return
    przeanalizuj(sciezka)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `warnings.catch_warnings(record=True)` **przechwytuje** ostrzeżenia zamiast wypisywać je na konsolę. `simplefilter("always")` mówi: **nie stosuj reguły „raz na sesję"**. Oba razem dają kompletny obraz tego, co openpyxl wie o pliku.

**Trzy rzeczy do przemyślenia:**

1. **`warnings.catch_warnings` jest kontekst managerem** — i ta uwaga nie jest przypadkowa. To **nie** `Workbook`; obiekt `warnings` wspiera `with`, więc **tę** konstrukcję wolno Ci stosować. Jeśli w module 2.2 zapamiętałeś „z `with` uważaj", to teraz widzisz, skąd się bierze to nieporozumienie: **`with` jest wszędzie w Pythonie; problem był konkretnie z `load_workbook`**.
2. **Liczba ostrzeżeń to metryka, nie ozdoba.** Trzy ostrzeżenia znaczą trzy różne, niezależne utraty. Zapisanie tej liczby do logu audytowego (przykład 4) daje Ci **punkt odniesienia** — gdy po zmianie w kodzie liczba wzrośnie, wiesz, że coś zepsułeś.
3. **Wyciszanie jest tu pokazane, żebyś wiedział, kiedy jest ZŁE.** Zobacz sekcję „JAK NIE WYCISZYĆ": zła praktyka to `simplefilter('ignore')` albo filtr na cały moduł `openpyxl`. **Wyciszasz konkretne ostrzeżenie, konkretną wiadomością, konkretnym modułem — i tylko wtedy, gdy plik jest wyłącznie do odczytu.** Nie ma tu litości dla „wyciszę, bo mi zaśmieca konsolę".

### Przykład 4 — `bezpieczny_zapis()` i kompletny potok modyfikacji (🔴)

Kulminacja modułu: wzorzec produkcyjny z 2.7 plus inwentarz przed i po (2.6) i log audytowy.

```python
"""Modul 18 - bezpieczny zapis: temp + os.replace + backup + log + inwentarz.

Uruchom:  python examples/18_bezpieczny_zapis.py
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

# importujemy inwentarz z przykladu 2 (ten sam katalog)
from pathlib import Path as _P
import sys
sys.path.insert(0, str(_P(__file__).resolve().parent))
from importlib import import_module
_inw = import_module("18_inwentarz_zip" if False else "modul18_inwentarz")

ROOT = _P(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
BACKUP = OUTPUT / "backup"
LOG = OUTPUT / "18_audyt.jsonl"


# ======================================================================
# 1. PODSUMOWANIE MODELU - "odcisk" do testu regresji (sekcja 2.10)
# ======================================================================
def podsumowanie_skoroszytu(sciezka: Path) -> dict[str, Any]:
    """Zwraca liczbowy odcisk skoroszytu.

    ⚠️ NIE read_only! W read_only nie ma walidacji, formatowania
    warunkowego, obrazów ani ochrony w modelu (moduł 17, sekcja 2.11).
    """
    wb = load_workbook(sciezka)
    try:
        suma_komorek = suma_formul = suma_regul = suma_tabel = 0
        suma_walidacji = suma_hidden = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        suma_komorek += 1
                        if isinstance(c.value, str) and c.value.startswith("="):
                            suma_formul += 1
                    if c.protection is not None and c.protection.hidden:
                        suma_hidden += 1
            suma_regul += len(ws.conditional_formatting)
            suma_tabel += len(ws.tables)
            suma_walidacji += len(ws.data_validations)
        return {
            "arkusze": wb.sheetnames,
            "liczba_arkuszy": len(wb.sheetnames),
            "komorki_z_wartoscia": suma_komorek,
            "formuly": suma_formul,
            "reguly_warunkowe": suma_regul,
            "tabele": suma_tabel,
            "walidacje": suma_walidacji,
            "komorki_hidden": suma_hidden,
            "nazwy_zdefiniowane": sorted(wb.defined_names.keys()),
            "wlasciwosci_tytul": wb.properties.title,
        }
    finally:
        wb.close()


def czesci_zip(sciezka: Path) -> dict[str, int]:
    """Liczy czesci archiwum wg kategorii. To wykryje utrate obrazow i wykresow."""
    import zipfile
    kategorie = {
        "media": 0, "charts": 0, "drawings": 0, "pivotTables": 0,
        "pivotCache": 0, "customXml": 0, "externalLinks": 0,
        "vba": 0, "worksheets": 0,
    }
    with zipfile.ZipFile(sciezka) as z:
        for name in z.namelist():
            if name.startswith("xl/media/"):
                kategorie["media"] += 1
            elif name.startswith("xl/charts/"):
                kategorie["charts"] += 1
            elif name.startswith("xl/drawings/"):
                kategorie["drawings"] += 1
            elif name.startswith("xl/pivotTables/"):
                kategorie["pivotTables"] += 1
            elif name.startswith("xl/pivotCache/"):
                kategorie["pivotCache"] += 1
            elif name.startswith("customXml/"):
                kategorie["customXml"] += 1
            elif name.startswith("xl/externalLinks/"):
                kategorie["externalLinks"] += 1
            elif name.endswith("vbaProject.bin"):
                kategorie["vba"] += 1
            elif name.startswith("xl/worksheets/sheet"):
                kategorie["worksheets"] += 1
    return kategorie


# ======================================================================
# 2. BEZPIECZNY ZAPIS - wzorzec z sekcji 2.7
# ======================================================================
def bezpieczny_zapis(wb, cel: Path, *,
                     backup: bool = True,
                     przyrostek: str = ".xlsx",
                     katalog_backup: Path | None = None,
                     katalog_log: Path | None = None) -> dict[str, Any]:
    """Zapisuje skoroszyt atomowo, z opcjonalnym backupem i logiem.

    Sekwencja:
      1. backup celu (z timestampem) - jesli plik istnieje
      2. plik tymczasowy W TYM SAMYM katalogu
      3. wb.save(tmp)
      4. zachowanie uprawnien starego celu
      5. os.replace(tmp, cel)   <- atomowa podmiana
      6. log audytowy
    """
    cel = Path(cel).resolve()
    cel.parent.mkdir(parents=True, exist_ok=True)
    katalog_backup = katalog_backup or BACKUP
    katalog_log = katalog_log or LOG

    znacznik = datetime.now().strftime("%Y%m%d-%H%M%S")
    rekord: dict[str, Any] = {
        "kiedy": datetime.now().isoformat(timespec="seconds"),
        "cel": str(cel),
        "istnial": cel.exists(),
        "backup": None,
        "rozmiar_przed": None,
        "rozmiar_po": None,
    }

    # 1. BACKUP z timestampem (copy2 zachowuje metadane)
    if cel.exists():
        rekord["rozmiar_przed"] = cel.stat().st_size
        if backup:
            katalog_backup.mkdir(parents=True, exist_ok=True)
            kopia = katalog_backup / f"{cel.stem}_{znacznik}{cel.suffix}"
            shutil.copy2(cel, kopia)
            rekord["backup"] = str(kopia)

    # 2. PLIK TYMCZASOWY W TYM SAMYM KATALOGU (atomycznosc os.replace!)
    fd, tmp_name = tempfile.mkstemp(dir=cel.parent, prefix=".tmp-raport-",
                                   suffix=przyrostek)
    os.close(fd)
    tmp = Path(tmp_name)

    try:
        # 3. ZAPIS
        wb.save(tmp)

        # 4. UPRAWNIENIA: os.replace nada uprawnienia PLIKU TYMCZASOWEGO.
        #    Jesli cel mial ograniczone prawa, przeniesmy je.
        if cel.exists():
            stary = os.stat(cel)
            os.chmod(tmp, stary.st_mode)

        # 5. ATOMOWA PODMIANA
        os.replace(tmp, cel)
        rekord["rozmiar_po"] = cel.stat().st_size
        rekord["ok"] = True

    except Exception as exc:                       # 6. SPRZATANIE
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
        rekord["ok"] = False
        rekord["blad"] = f"{type(exc).__name__}: {exc}"
        _dopisz_log(katalog_log, rekord)
        raise                                      # NIE polykamy wyjatku!

    _dopisz_log(katalog_log, rekord)
    return rekord


def _dopisz_log(sciezka_log: Path, rekord: dict[str, Any]) -> None:
    sciezka_log.parent.mkdir(parents=True, exist_ok=True)
    with sciezka_log.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rekord, ensure_ascii=False) + "\n")


# ======================================================================
# 3. POTOK: wczytaj -> inwentarz -> zmien -> zapisz -> inwentarz -> raport
# ======================================================================
def main() -> None:
    import modul18_inwentarz as inw

    ZRODLO = DATA / "18_dane_wejsciowe.xlsx"
    if not ZRODLO.exists():
        print("Brak pliku wejściowego. Uruchom najpierw:")
        print("  python examples/18_read_modify_write.py")
        return

    CEL = OUTPUT / "18_raport_bezpieczny.xlsx"

    print("=" * 82)
    print("POTOK BEZPIECZNEJ MODYFIKACJI")
    print("=" * 82)

    # --- A. SCAN-BEFORE-SAVE: inwentarz pliku wejsciowego ---
    print("\n[A] INWENTARZ WEJŚCIA")
    inv_przed = inw.inwentarz(ZRODLO)
    print(f"    części: {len(inv_przed['czesci'])}   "
          f"rozszerzenia: {inv_przed['rozszerzenia'] or 'brak'}")

    odcisk_przed = podsumowanie_skoroszytu(ZRODLO)
    zip_przed = czesci_zip(ZRODLO)

    # --- B. WCZYTANIE z poprawnymi flagami ---
    print("\n[B] WCZYTANIE")
    wb = load_workbook(ZRODLO)               # formuly i styl zostaja
    print(f"    arkuszу: {wb.sheetnames}")

    # --- C. MODYFIKACJA ---
    print("\n[C] MODYFIKACJA")
    ws: Worksheet = wb["Sprzedaż"]
    ws.cell(row=1, column=5, value="Marża 20%")
    for row in range(2, ws.max_row + 1):
        kot = ws.cell(row=row, column=3).coordinate
        ws.cell(row=row, column=5, value=f"=ROUND({kot}*0.2,2)")
    ws["G1"] = "Wygenerowano"
    ws["G2"] = datetime.now()
    ws["G2"].number_format = "yyyy-mm-dd hh:mm"
    print("    dodano kolumnę E (formuły) i metadane w G")

    # --- D. BEZPIECZNY ZAPIS ---
    print("\n[D] BEZPIECZNY ZAPIS")
    try:
        rekord = bezpieczny_zapis(wb, CEL)
        print(f"    zapisano: {rekord['cel']}")
        print(f"    backup:   {rekord['backup']}")
        print(f"    rozmiar:  {rekord['rozmiar_przed']} -> {rekord['rozmiar_po']} B")
    finally:
        wb.close()                     # ⚠️ закрыть zawsze

    # --- E. SCAN-AFTER-SAVE + PORÓWNANIE ---
    print("\n[E] INWENTARZ WYJŚCIA I PORÓWNANIE")
    inv_po = inw.inwentarz(CEL)
    odcisk_po = podsumowanie_skoroszytu(CEL)
    zip_po = czesci_zip(CEL)

    print("\n    MODEL (odcisk skoroszytu):")
    for klucz in odcisk_przed:
        a, b = odcisk_przed[klucz], odcisk_po[klucz]
        zmiana = "" if a == b else "   <-- ZMIANA"
        if klucz != "arkusze":
            print(f"      {klucz:<24} {a!s:<28} -> {b!s:<28}{zmiana}")

    print("\n    ARCHIWUM (części ZIP):")
    for klucz in zip_przed:
        a, b = zip_przed[klucz], zip_po[klucz]
        flaga = ""
        if b < a:
            flaga = "   🔴 UTRATA!"
        elif b > a:
            flaga = "   (nowe części)"
        print(f"      {klucz:<16} {a:>4} -> {b:>4}{flaga}")

    # --- F. WERDYKT ---
    utraty = [k for k in zip_przed if zip_po[k] < zip_przed[k]]
    rozsz_przed = inv_przed["rozszerzenia"]
    print()
    print("=" * 82)
    if utraty or rozsz_przed:
        print("WERDYKT KOŃCOWY: ⚠️  SPRAWDŹ RĘCZNIE")
        if utraty:
            print(f"  Ubytki w archiwum: {utraty}")
        if rozsz_przed:
            print("  Rozszerzenia w wejściu, których openpyxl nie modeluje:")
            for arkusz, typy in rozsz_przed.items():
                print(f"    {arkusz}: {typy}")
    else:
        print("WERDYKT KOŃCOWY: 🟢 brak sygnałów utraty w porównaniu liczbowym")
    print()
    print("  OSTATNI KROK, KTÓREGO NIE WOLNO POMINĄĆ:")
    print(f"  Otwórz {CEL.name} ORAZ {ZRODLO.name} w Excelu i porównaj okiem.")
    print("  Formuły w kolumnie E muszą się liczyć po otwarciu (openpyxl ich")
    print("  NIE oblicza - moduł 06). Nagłówek w G2 zapisany jako data musi")
    print("  być datą, nie liczbą.")
    print("=" * 82)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci i na dysku.** Kolejność jest tu istotna i warto ją prześledzić:

- `podsumowanie_skoroszytu()` wczytuje **pełny model** (dla `styles`, reguł, walidacji — bez `read_only`), liczy i **zamyka**. Model idzie do śmieci.
- `bezpieczny_zapis()` tworzy **plik tymczasowy z losową nazwą** w katalogu docelowym. Nazwa zaczyna się od kropki (`.tmp-raport-...`), więc nawet gdyby proces padł, plik jest niewidoczny w większości menedżerów plików.
- `os.replace(tmp, cel)` to **jedna operacja systemowa**: w tym momencie `cel` widzi albo starą wersję, albo nową. Nie ma stanu pośredniego.
- `podsumowanie_skoroszytu()` na wyniku daje **odcisk po**. Różnica między odciskami to Twoja utrata danych — **wypisana w liczbach**, nie „na wyczucie".

**Trzy rzeczy do przemyślenia:**

1. **`os.chmod` przed `os.replace` to nie przesada.** Bez tego plik wynikowy dostanie uprawnienia pliku tymczasowego z `mkstemp`, czyli **`0600`** (tylko właściciel). Jeśli raport leżał w katalogu współdzielonym i miał prawa grupowe, po zapisie **stracisz je** — i dopiero wtedy zauważysz, że współpracownicy nie mogą go otworzyć. Na Linuksie to realna wpadka; w `os.chmod` widzisz dokładnie dlaczego.
2. **`raise` w bloku `except` jest tu obowiązkowe.** Kuszące byłoby napisać `except Exception: pass` w imię „skrypt nie może się wywalać". Ale wtedy `bezpieczny_zapis` zwróciłby sukces, mimo że plik nie powstał. **Wyjątek musi iść dalej**, a log audytowy rejestruje porażkę. Log i wyjątek to dwa kanały tej samej informacji dla dwóch różnych odbiorców: log dla Ciebie, wyjątek dla wywołującego kod.
3. **`finally: wb.close()` jest tu naprawdę ważne**, mimo że w trybie normalnym `close()` „tylko zamyka, jeśli coś jest otwarte". Powód jest inny: jeśli `bezpieczny_zapis` rzuci wyjątek, `wb` **nadal siedzi w pamięci** z całym modelem. W długo działającym procesie (moduł 26, zadanie w kolejce) to wyciek pamięci, który objawi się dopiero po tysiącu raportów.

## 4. Anatomia API

| Klasa / funkcja / właściwość | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `load_workbook(filename, **flagi)` | Wczytuje plik i buduje pełny model | `read_only`, `keep_vba`, `data_only`, `keep_links`, `rich_text` | Patrz tabela 2.2. **Bez `with` — nie działa!** |
| `load_workbook(..., read_only=True)` | Tryb leniwy | — | Brak wykresów, obrazów, DV, CF, komentarzy, ochrony w modelu. `save()` → `TypeError` |
| `load_workbook(..., keep_vba=True)` | Przenosi `vbaProject.bin` | — | Makra **nieedytowalne**; wymagane dla `.xlsm`/`.xltm` |
| `load_workbook(..., data_only=True)` | Formuła → ostatnia wartość z cache | — | **`save()` po tym niszczy formuły!** Tylko do odczytu |
| `load_workbook(..., keep_links=False)` | Ucina odwołania zewnętrzne i ich cache | — | Potrzebne dla szybkości, ale psuje formuły z `#REF!` |
| `load_workbook(..., rich_text=True)` | Zachowuje formatowanie fragmentów tekstu | — | Nowość 3.1; sprawdź działanie na swoim pliku |
| `wb.close()` | Zamyka uchwyt archiwum | — | **Istotne tylko w `read_only`/`write_only`**; w normalnym trybie formalność |
| `wb.save(filename)` | Serializuje **cały model** do nowego ZIP-a | ścieżka lub obiekt plikopodobny binarny | Nadpisuje bez ostrzeżenia; `save()` w `read_only` → `TypeError("Workbook is read-only")` |
| `wb.save(BytesIO())` | Zapis do pamięci | — | Do testów (moduł 20) i odpowiedzi HTTP (moduł 26) |
| `wb.vba_archive` | Archiwum z makrami (jeśli `keep_vba`) | — | `None`, gdy nie żądano makr |
| `wb.mime_type` | Typ MIME wynikający z `vba_archive` i `template` | — | **openpyxl nie wymusza zgodności z rozszerzeniem** — Excel wymusza |
| `wb.properties` | `DocumentProperties`: autor, tytuł, daty | — | Ślad audytowy; zachowywane przy zapisie |
| `wb.custom_doc_props` | Własne właściwości (**nowość 3.1**) | — | Zachowywane przy zapisie |
| `wb.defined_names` | Słownik nazw zdefiniowanych | `.add(DefinedName(...))` | **API zmienione w 3.1** — stare `add_named_range` usunięte |
| `DefinedName(name, attr_text=...)` | Obiekt nazwy zdefiniowanej | `attr_text`: `'Arkusz'!$A$1:$A$10` | Zamiast `NamedRange` z openpyxl 2.x |
| `ws.protection`, `wb.security` | Ochrona arkusza i skoroszytu | — | Zachowywana przy round-tripie (moduł 17) |
| `zipfile.ZipFile(path).namelist()` | Lista części archiwum | — | **Podstawa inwentarza — scan-before-save** |
| `openpyxl.worksheet._reader.EXT_TYPES` | Mapa identyfikatorów rozszerzeń na nazwy | — | **API prywatne** — obuduj `try/except` |
| `openpyxl.worksheet._reader.parse_extensions` | Źródło wszystkich ostrzeżeń o rozszerzeniach | — | To ona wypisuje „... will be removed" |
| `chart._reindex()` | Ustawia `order` serii kolejno od 0 | — | **Prywatne**; lekarstwo na „uszkodzony plik" w kombinacjach wykresów |
| `os.replace(src, dst)` | **Atomowa** podmiana pliku | — | Atomowa w obrębie tego samego systemu plików |
| `tempfile.mkstemp(dir=..., prefix=..., suffix=...)` | Tworzy unikalny plik tymczasowy | zwraca `(fd, path)` | **Zamknij `fd`** — otworzy go openpyxl |
| `os.chmod(tmp, mode)` | Uprawnienia pliku tymczasowego | — | Bez tego `mkstemp` daje `0600` i tracisz prawa grupowe |
| `shutil.copy2(src, dst)` | Kopia **z metadanymi** | — | Do backupów — zachowuje mtime i prawa |
| `warnings.catch_warnings(record=True)` | Przechwytuje ostrzeżenia | — | **Elegancki kontekst manager — tu `with` działa** |
| `warnings.simplefilter("always")` | Pokazuje każde ostrzeżenie | — | Bez tego zobaczysz tylko pierwsze z każdego rodzaju |
| `hashlib.sha256(...).hexdigest()` | Hash danych wejściowych | — | Identyfikator do idempotencji — **z danych, nie z zegara** |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — dopisz kolumnę i udowodnij, że nic nie zginęło.**

Napisz `examples/18_cw1.py`:

1. Utwórz plik źródłowy `output/18_cw1_wejsciowy.xlsx` z arkuszem `Dane`:
   nagłówki `Produkt`, `Cena`, `Ilość` + 5 wierszy danych, nagłówki pogrubione z tłem.
2. **Zrób kopię bajtową** (`shutil.copy2`) i wypisz jej rozmiar.
3. Wczytaj plik **w trybie domyślnym** (`load_workbook` bez `read_only` i bez `data_only`).
4. Dopisz kolumnę `Wartość` z formułą `=B2*C2` w każdym wierszu.
5. Zapisz **pod nową nazwą** `output/18_cw1_wynik.xlsx`.
6. Wczytaj ponownie oba pliki i wypisz:
   - `sheetnames`, `max_row`, `max_column` każdego,
   - czy nagłówki mają `bold=True` w obu,
   - czy w kolumnie `D` wyniku jest **formuła** (tekst zaczynający się od `=`) czy liczba.
7. Wypisz **trzykrotnie** zdanie przypominające o weryfikacji „na oko".

W `output/18_cw1_wnioski.md` odpowiedz:

1. **Czy rozmiar pliku wynikowego jest większy, mniejszy, czy podobny?** Dlaczego (podpowiedź: `sharedStrings`, `styles.xml`, nowe komórki)?
2. **Co by się zmieniło, gdybyś w kroku 3 użył `data_only=True`?** Sprawdź to eksperymentalnie: zrób to samo w drugim pliku i porównaj zawartość kolumny `D` w wyniku. Wklej oba wyniki i opisz różnicę, odwołując się do tabeli z 2.2.
3. **Czy `shutil.copy2` różni się od `shutil.copy` w tym zadaniu?** Sprawdź `st_mtime` obu plików po skopiowaniu i wyjaśnij.
4. **Dlaczego nie zapisałeś wyniku pod nazwą pliku wejściowego?** Odpowiedz jednym zdaniem — tym, które powtórzyliśmy w module trzy razy.
5. **Czy `wb.close()` coś tu zmienił?** Wyjaśnij, odwołując się do docstringa metody `Workbook.close`.

### 🟡 Warsztat

**Zadanie 2 — inwentarz i wykrywanie utraty.**

Napisz `examples/18_cw2.py`, który:

1. Tworzy **bogaty** plik testowy `output/18_cw2_bogaty.xlsx`, zawierający:
   - arkusz `Dane` z tabelą danych (`Table` z modułu 12) i filtrem,
   - **wykres** słupkowy z modułu 15,
   - **obraz** PNG wygenerowany w kodzie (użyj `matplotlib` jeśli dostępny, w przeciwnym razie utwórz minimalny PNG bajtowo przez `Pillow`),
   - **komentarz** (`Comment`) w komórce,
   - **hiperlink**,
   - **walidację danych** typu `list` z listy wpisanej wprost,
   - **formatowanie warunkowe** `CellIsRule` na kolumnie,
   - **ochronę arkusza** (bez hasła, żeby nie blokować sobie testu),
   - **właściwości dokumentu** (`wb.properties.title`, `.creator`).
2. Uruchamia na tym pliku **inwentarz ZIP** (własną implementację z przykładu 2 lub zaimportowaną) i wypisuje:

   - liczbę części wg kategorii,
   - listę rozszerzeń `extLst`,
   - werdykt.
3. Wczytuje plik **z nasłuchem ostrzeżeń** i wypisuje wszystkie ostrzeżenia z ich znaczeniem.
4. Modyfikuje **jedną komórkę** w arkuszu `Dane` (np. wartości w pierwszej komórce danych).
5. Zapisuje jako `output/18_cw2_po_obrobce.xlsx`.
6. Uruchamia inwentarz i nasłuch **ponownie**, na wyniku.
7. Wypisuje **tabelę porównawczą przed/po**: kategoria → liczba przed → liczba po → werdykt (🔴 utrata / 🟢 ok).

W `output/18_cw2_wnioski.md` odpowiedz:

1. **Które elementy wytrzymały bez zmian?** Sprawdź każdy z punktu 1 osobno i uzasadnij wynik, odwołując się do inwentarza z 2.3 i 2.4.
2. **Czy obraz nadal jest w pliku?** Sprawdź nie tylko `xl/media/`, ale też **czy narysowanie nadal na niego wskazuje** (`xl/drawings/*_rels`). Co się dzieje z obrazem, gdy nie ma go w `xl/media/`, ale rysunek go wymienia?
3. **Czy wykres nadal się otwiera bez „naprawy"?** Otwórz plik w Excelu i zanotuj **dokładny komunikat** albo jego brak. Czy wykres wygląda jak w oryginale? Opisz **jedną** różnicę, jeśli jest.
4. **Czy ostrzeżenia z punktu 3 i 6 są identyczne?** Jeśli wynik ma mniej ostrzeżeń niż wejście — wyjaśnij to, odwołując się do tego, co robi `parse_extensions` (2.4).
5. **Co z ochroną arkusza i formatowaniem warunkowym?** Sprawdź przez API (`ws.protection.sheet`, `ws.conditional_formatting`). Czy to, co widzisz w API, wystarcza, by twierdzić, że użytkownik zobaczy to samo w Excelu?
6. **Jaka **jedna** pozycja z Twojego inwentarza NIE dała żadnego ostrzeżenia, a jednak się zmieniła?** To pytanie jest esencją modułu — odpowiedź uzasadnij.

### 🔴 Wyzwanie

**Zadanie 3 — `bezpieczny_zapis()` w wersji produkcyjnej.**

Napisz `examples/18_cw3.py` z **pełnym, przetestowanym** mechanizmem bezpiecznego zapisu.

**Wymagania funkcjonalne `bezpieczny_zapis(wb, cel, **opcje)`:**

1. **Kopia zapasowa** celu z timestampem w nazwie (`stem_YYYYmmdd-HHMMSS.ext`), tworzona **tylko gdy cel istnieje**, przez `shutil.copy2` (**nie** `copy` — uzasadnij w komentarzu).
2. **Plik tymczasowy w tym samym katalogu co cel** (uzasadnij w komentarzu, dlaczego nie w `/tmp`).
3. **Zapis do pliku tymczasowego** przez `wb.save(tmp)`.
4. **Zachowanie uprawnień** poprzedniego celu (`os.stat` → `os.chmod`) — z komentarzem wyjaśniającym, że `os.replace` nadałby uprawnienia pliku tymczasowego.
5. **Atomowa podmiana** przez `os.replace`.
6. **Sprzątanie w `except`** — usuń plik tymczasowy, ale **propaguj wyjątek dalej** (`raise`), nie połykaj.
7. **Log audytowy** — dopisywany **linia po linii** do pliku JSONL (`json.dumps(...) + "\n"`), zawierający: timestamp ISO, ścieżkę celu, informację czy istniał, ścieżkę backupu, rozmiar przed i po, status `ok`, a w razie błędu — klasę i treść wyjątku.
8. **Parametr `backup: bool = True`** — możliwość wyłączenia kopii.
9. **Parametr `prefix_tymczasowego`** — domyślnie `.tmp-raport-`, żeby plik tymczasowy był ukryty w menedżerach plików.
10. **Weryfikacja po zapisie** — funkcja po `os.replace` sprawdza, czy plik istnieje i czy ma niezerowy rozmiar; jeśli nie — rzuca `RuntimeError`.

**Wymagania jakościowe:**

- **Testy** (bez pytest — proste asercje w `main()`, albo z pytest jeśli masz):
  - zapis do **nieistniejącego** celu (bez backupu) → działa,
  - zapis do **istniejącego** celu → powstaje backup, cel zaktualizowany,
  - **symulowany błąd zapisu** (monkey-patch `wb.save` rzucający wyjątek) → brak pliku tymczasowego po zakończeniu, wyjątek propagowany, wpis w logu z `ok: false`,
  - **uprawnienia** — ustaw na pliku docelowym `0o640`, zapisz, sprawdź `st_mode & 0o777` po zapisie.
- **Log** — po testach wypisz jego zawartość, sformatowaną.
- **Podsumowanie** — na końcu wypisz liczbę operacji: udanych, nieudanych, kopii zapasowych.

**Wymaganie dodatkowe — idempotencja:**

Dodaj drugą funkcję `dopisz_dane_idempotentnie(sciezka, identyfikator, wiersze)`, która:

1. wczytuje skoroszyt (tryb domyślny!),
2. sprawdza, czy `identyfikator` już istnieje w arkuszu `_meta` (utwórz go, jeśli nie ma),
3. jeśli istnieje — **wypisuje komunikat i kończy bez zapisu**,
4. jeśli nie — dopisuje wiersze, dopisuje wpis do `_meta` i zapisuje przez `bezpieczny_zapis`,
5. używa jako identyfikatora **hasha danych** (`hashlib.sha256`), nie znacznika czasu.

**Weryfikacja końcowa:** uruchom `dopisz_dane_idempotentnie` **dwa razy** z tym samym identyfikatorem i pokaż, że drugie uruchomienie **nie zmienia pliku** (porównaj `st_mtime` i rozmiar przed/po).

W `output/18_cw3_wnioski.md` odpowiedz:

1. **Dlaczego `os.replace`, a nie `shutil.move` albo `os.rename`?** Trzy różnice, po jednym zdaniu każda. Kiedy `os.rename` zawodzi tam, gdzie `os.replace` działa?
2. **Co konkretnie tracisz, jeśli pominiesz `os.chmod`?** Pokaż wartość `st_mode` przed i po zapisie w swojej wersji.
3. **Dlaczego plik tymczasowy musi być w **tym samym** katalogu?** Wyjaśnij, kiedy atomowość znika i dlaczego.
4. **Dlaczego w `except` MUSI być `raise`?** Co by się stało, gdybyś napisał `return rekord` zamiast `raise`? Wskaż scenariusz, w którym to by kogoś poważnie oszukało.
5. **Dlaczego identyfikatorem idempotencji jest hash danych, a nie `datetime.now()`?** Sformułuj w jednym zdaniu, co by się stało z podwójnym uruchomieniem przy znaczniku czasu.
6. **Jak wpisałbyś do logu utratę danych?** Zaprojektuj rozszerzenie logu: obok `rozmiar_przed`/`rozmiar_po` dodaj `utraty_zip` (lista kategorii części, których ubyło) i `ostrzezenia` (lista treści ostrzeżeń). Nie musisz tego implementować — opisz **trzy linie** kodu, które by to zrobiły.

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Bezpieczny zapis - wersja produkcyjna (kluczowe fragmenty)."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import tempfile
import warnings
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


# ======================================================================
# 1. BEZPIECZNY ZAPIS
# ======================================================================
def bezpieczny_zapis(
    wb,
    cel: Path | str,
    *,
    backup: bool = True,
    prefix_tymczasowego: str = ".tmp-raport-",
    katalog_backup: Path | None = None,
    sciezka_log: Path | None = None,
) -> dict[str, Any]:
    """Zapis atomowy: temp w tym samym katalogu -> os.replace.

    Gwarancja: w dowolnym momencie 'cel' zawiera ALBO starą wersję
    w całości, ALBO nową w całości. Nigdy hybrydę.
    """
    cel = Path(cel).resolve()
    cel.parent.mkdir(parents=True, exist_ok=True)
    katalog_backup = katalog_backup or cel.parent / "backup"
    sciezka_log = sciezka_log or cel.parent / "audyt.jsonl"

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    rekord: dict[str, Any] = {
        "kiedy": datetime.now().isoformat(timespec="seconds"),
        "cel": str(cel),
        "istnial": cel.exists(),
        "backup": None,
        "rozmiar_przed": None,
        "rozmiar_po": None,
        "ok": False,
    }

    # ---- 1. BACKUP -------------------------------------------------
    if cel.exists():
        rekord["rozmiar_przed"] = cel.stat().st_size
        if backup:
            katalog_backup.mkdir(parents=True, exist_ok=True)
            kopia = katalog_backup / f"{cel.stem}_{ts}{cel.suffix}"
            # copy2, a nie copy: zachowuje mtime i uprawnienia.
            # Dla audytu "kiedy powstała ta wersja" to jest istotne.
            shutil.copy2(cel, kopia)
            rekord["backup"] = str(kopia)

    # ---- 2. PLIK TYMCZASOWY ---------------------------------------
    # W TYM SAMYM katalogu, bo os.replace jest atomowy TYLKO w obrębie
    # jednego systemu plików. Katalog /tmp to inny system plików ->
    # podmiana zamienia się w kopiowanie i traci atomowość.
    fd, tmp_name = tempfile.mkstemp(
        dir=cel.parent, prefix=prefix_tymczasowego, suffix=cel.suffix
    )
    os.close(fd)          # mkstemp TWORZY plik; openpyxl otworzy go sam
    tmp = Path(tmp_name)

    try:
        # ---- 3. ZAPIS ---------------------------------------------
        wb.save(tmp)

        # ---- 4. UPRAWNIENIA ---------------------------------------
        # os.replace nada 'cel' uprawnienia pliku TYMCZASOWEGO,
        # czyli 0o600 z mkstemp -> stracilibyśmy prawa grupowe.
        if cel.exists():
            os.chmod(tmp, os.stat(cel).st_mode)

        # ---- 5. PODMIANA ATOMOWA ----------------------------------
        os.replace(tmp, cel)

        # ---- 6. WERYFIKACJA ---------------------------------------
        if not cel.exists() or cel.stat().st_size == 0:
            raise RuntimeError(f"Po zapisie cel jest pusty/nie istnieje: {cel}")

        rekord["rozmiar_po"] = cel.stat().st_size
        rekord["ok"] = True

    except Exception as exc:
        # ---- 7. SPRZATANIE ----------------------------------------
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass          # nie nadpisuj oryginalnego wyjątku
        rekord["blad"] = f"{type(exc).__name__}: {exc}"
        _log(sciezka_log, rekord)
        # NIE polykamy: wywołujący kod MUSI wiedzieć, że raport nie powstał.
        raise

    _log(sciezka_log, rekord)
    return rekord


def _log(sciezka: Path, rekord: dict[str, Any]) -> None:
    sciezka.parent.mkdir(parents=True, exist_ok=True)
    with sciezka.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(rekord, ensure_ascii=False) + "\n")


# ======================================================================
# 2. IDEMPOTENCJA - dopisz raz, ile razy nie uruchomisz
# ======================================================================
def _odcisk(dane: str) -> str:
    """Hash danych wejsciowych - identyfikator biegu, NIE czas."""
    return hashlib.sha256(dane.encode("utf-8")).hexdigest()[:16]


def dopisz_dane_idempotentnie(
    sciezka: Path,
    identyfikator: str,
    wiersze: list[list[Any]],
    *,
    arkusz_danych: str = "Dane",
    arkusz_meta: str = "_meta",
) -> bool:
    """Dopisuje wiersze TYLKO jesli 'identyfikator' nie byl jeszcze uzyty.

    Zwraca True, jesli plik zostal zmieniony.
    """
    wb = load_workbook(sciezka)      # tryb domyslny: formuly i styl zostaja!
    try:
        if arkusz_meta not in wb.sheetnames:
            meta = wb.create_sheet(arkusz_meta)
            meta.append(["identyfikator", "kiedy", "wierszy"])
        meta = wb[arkusz_meta]

        # idempotencja: sprawdz, czy identyfikator juz jest
        for row in meta.iter_rows(min_col=1, max_col=1, values_only=True):
            if row[0] == identyfikator:
                print(f"  ⏭  Pominięto: identyfikator {identyfikator!r} "
                      f"już obecny. Plik NIE zmieniony.")
                return False

        ws = wb[arkusz_danych]
        for w in wiersze:
            ws.append(w)
        meta.append([identyfikator, datetime.now().isoformat(timespec="seconds"),
                     len(wiersze)])

        bezpieczny_zapis(wb, sciezka)
        print(f"  ✓  Dopisano {len(wiersze)} wierszy (id={identyfikator})")
        return True
    finally:
        wb.close()


# ======================================================================
# 3. TESTY (fragment)
# ======================================================================
def testy() -> None:
    from openpyxl import Workbook

    tmpdir = Path("output/cw3")
    tmpdir.mkdir(parents=True, exist_ok=True)
    cel = tmpdir / "test.xlsx"

    # (a) zapis do nieistniejacego celu - brak backupu
    wb = Workbook(); wb["Sheet"]["A1"] = 1
    rek = bezpieczny_zapis(wb, cel)
    wb.close()
    assert rek["ok"] and rek["backup"] is None, "brak backupu przy nowym pliku"
    print("  ✓ (a) nowy plik zapisany bez backupu")

    # (b) zapis do istniejacego - backup powstaje
    wb = Workbook(); wb["Sheet"]["A1"] = 2
    rek = bezpieczny_zapis(wb, cel)
    wb.close()
    assert rek["ok"] and rek["backup"] is not None, "backup nie powstał"
    assert Path(rek["backup"]).exists(), "wskazany backup nie istnieje"
    print(f"  ✓ (b) backup: {Path(rek['backup']).name}")

    # (c) symulowany blad zapisu -> brak pliku tymczasowego, wyjatek wyzej
    class PsujacyWorkbook:
        def save(self, _):
            raise IOError("symulowany blad dysku")

    przed = {p.name for p in tmpdir.iterdir()}
    try:
        bezpieczny_zapis(PsujacyWorkbook(), cel)
        raise AssertionError("wyjątek NIE został propagowany!")
    except IOError:
        print("  ✓ (c) wyjątek propagowany dalej")
    po = {p.name for p in tmpdir.iterdir()}
    resztki = [n for n in po - przed if n.startswith(".tmp-")]
    assert not resztki, f"pozostały pliki tymczasowe: {resztki}"
    print("  ✓ (c) brak pliku tymczasowego po błędzie")

    # (d) uprawnienia
    os.chmod(cel, 0o640)
    wb = Workbook(); wb["Sheet"]["A1"] = 3
    bezpieczny_zapis(wb, cel)
    wb.close()
    mode = cel.stat().st_mode & 0o777
    assert mode == 0o640, f"uprawnienia zmienione: {oct(mode)}"
    print(f"  ✓ (d) uprawnienia zachowane: {oct(mode)}")

    # (e) idempotencja
    ident = _odcisk("zestaw-1")
    zmiana1 = dopisz_dane_idempotentnie(cel, ident, [["a", 1], ["b", 2]])
    mtime1 = cel.stat().st_mtime_ns
    rozmiar1 = cel.stat().st_size
    zmiana2 = dopisz_dane_idempotentnie(cel, ident, [["a", 1], ["b", 2]])
    mtime2 = cel.stat().st_mtime_ns
    assert zmiana1 is True and zmiana2 is False, "idempotencja nie działa"
    assert mtime1 == mtime2 and rozmiar1 == cel.stat().st_size, \
        "drugi bieg zmienił plik!"
    print("  ✓ (e) idempotencja: drugi bieg nie zmienił pliku (mtime identyczny)")
```

**Kluczowe decyzje i uzasadnienia:**

- **`tempfile.mkstemp` + `os.close(fd)`.** `mkstemp` jest bezpieczniejszy od ręcznego składania nazwy, bo gwarantuje unikatowość atomowo (bez wyścigu). Deskryptor trzeba zamknąć — inaczej plik pozostaje otwarty dla zapisu i openpyxl (który otwiera go po nazwie) nie nadpisze go na Windows.

- **Weryfikacja po `os.replace`.** Jedna linia, a łapie klasę błędów, o której nikt nie myśli: `wb.save()` **przeszedł bez wyjątku**, a plik ma zero bajtów. To się zdarza, gdy skoroszyt był pusty w nietypowy sposób albo gdy zapis trafił na specyficzny problem systemu plików. Bez tej asercji użytkownik dostanie pusty plik i będzie przekonany, że „wszystko grało".

- **`except OSError: pass` przy sprzątaniu — ale `raise` poza nim.** To istotne rozróżnienie: **sprzątanie** nie może zamaskować **przyczyny** awarii. Jeśli `tmp.unlink()` zawiedzie (np. plik zniknął w międzyczasie), to nie jest problem — problemem jest to, że zapis się nie udał. Kolejność jest tu przemyślana.

- **Idempotencja przez hash danych, nie przez czas.** To jest sedno punktu 5 wniosków. Ze znacznikiem czasu **każde** uruchomienie ma „nowy" identyfikator, więc skrypt dopisuje dane ponownie. Hash danych wejściowych jest **niezmiennikiem**: te same dane → ten sam identyfikator → drugi bieg pominięty. To dokładnie „przycisk windy" z 1.5.

- **`meta.append(...)` zamiast `ws.append(...)`** — bo piszemy do arkusza `_meta`, a nie do danych. Brzmi banalnie, ale w kodzie produkcyjnym łatwo pomieszać `ws` z `meta` i dopisać identyfikator do tabeli danych. Trzymanie **jawnych nazw** (`meta`, `ws`) zamiast jednej zmiennej `sheet` to zabezpieczenie przed tym błędem.

- **Czego tu nie ma i dlaczego.** Nie ma blokady pliku (`fcntl`/`msvcrt`) — bo openpyxl nie oferuje takiej integracji, a dodawanie jej tutaj byłoby obietnicą bez pokrycia. Nie ma logowania ostrzeżeń openpyxl — bo to operacja ortogonalna (nasłuch przy **wczytywaniu**, zapis przy **zapisywaniu**); w wersji produkcyjnej połączyłbyś oba w jeden kontekst `warnings.catch_warnings` obejmujący **cały** potok. To właśnie sugeruje punkt 6 wniosków.

- **Co tu jest antywzorcem?** Dwie rzeczy, które poprawisz w module 26: (a) `bezpieczny_zapis` przyjmuje `wb` (typ openpyxl) — to przeciek abstrakcji; docelowo przyjmuje **swój** obiekt raportu i sam decyduje, jak go zapisać; (b) log jest pisany do pliku w tym samym miejscu co raporty — docelowo to **osobny port** (`Logger`), który w testach podmieniasz na listę.

</details>

## 6. Typowe błędy i pułapki

**1. „Wszystkie makra zniknęły i Excel mówi, że plik jest uszkodzony" (objaw) → **wczytano `.xlsm` bez `keep_vba=True`.** Model nie ma `vba_archive`, więc przy zapisie nie ma czego dopisać do archiwum, a plik nadal ma rozszerzenie `.xlsm` (przyczyna) → **`load_workbook(path, keep_vba=True)` dla każdego `.xlsm`/`.xltm`.** Uznaj to za regułę bez wyjątków (naprawa).**

**2. „Formuły zamieniły się w liczby, a użytkownik pyta, dlaczego nie może prześledzić wyliczeń" (objaw) → **wczytano z `data_only=True` i zapisano.** Model nie trzyma formuł, tylko wartości z cache, więc writer nie ma czego zapisać (2.2) (przyczyna) → **`data_only=True` tylko do odczytu i tylko wtedy, gdy NIE zapisujesz.** Jeśli musisz mieć jedno i drugie: **dwa osobne `load_workbook`** — jedno do odczytu wartości, drugie do zapisu formuł (naprawa).**

**3. „Po zapisie użytkownik zobaczył `#REF!` w kolumnie z odwołaniami do innego skoroszytu" (objaw) → **użyto `keep_links=False`** albo plik stracił części `xl/externalLinks/*`. Kod reader-a robi `package.externalReferences = []`, więc formuły nie mają do czego się odwołać (2.2) (przyczyna) → **Ustaw `keep_links=True` (domyślnie) dla plików z odwołaniami zewnętrznymi.** Sprawdź przed zapisem, czy w archiwum jest `xl/externalLinks/` (naprawa).**

**4. „Zniknęły wszystkie listy rozwijane" (objaw) → **walidacje były w bloku `x14`** i poszły w `parse_extensions` z ostrzeżeniem, które przeoczyłeś; dotyczy to szczególnie list wskazujących na inny arkusz (2.4, moduł 17) (przyczyna) → **Nasłuchuj ostrzeżeń** (`catch_warnings` + `simplefilter("always")`), **odetnij ryzyko na wejściu** (inwentarz ZIP) i **odbuduj reguły** albo przejdź na nazwy zdefiniowane (naprawa).**

**5. „Plik otwiera się, ale Excel pyta, czy go naprawić" (objaw) → **najprawdopodobniej wykres kombinowany** z powielonymi `idx`/`order`. openpyxl normalizuje serie per typ, więc drugi typ startuje od zera i Excel widzi duplikat (2.4) (przyczyna) → **Przejrzyj serie wykresu i użyj `chart._reindex()`** (API prywatne — sprawdź w swojej wersji). Jeśli nie pomaga: **nie przepuszczaj pliku przez openpyxl** (naprawa).**

**6. „Nasłuchałem ostrzeżeń, a widzę tylko jedno" (objaw) → **Python domyślnie pokazuje każde **unikalne** ostrzeżenie tylko raz na sesję** (przyczyna) → **`warnings.catch_warnings(record=True)` + `warnings.simplefilter("always")`.** Zbieraj do listy i raportuj wszystkie naraz (przykład 3) (naprawa).**

**7. „Wyciszyłem ostrzeżenia, bo zaśmiecały konsolę — i przegapiłem prawdziwy problem" (objaw) → **`simplefilter('ignore')` albo filtr na cały moduł `openpyxl`** zamiast filtra na konkretną wiadomość (przykład 3) (przyczyna) → **Filtruj po `message=` i `module='openpyxl'`, i tylko wtedy, gdy plik jest wyłącznie do odczytu.** Wyciszenie ukrywa komunikat, a nie efekt — jeśli potem zapiszesz plik, dane i tak zginą (naprawa).**

**8. „Zapisałem plik pod tą samą nazwą i nie zauważyłem problemu — aż ktoś poprosił o «tę starszą wersję»" (objaw) → **nadpisanie wejścia bez kopii.** openpyxl nadpisuje bez ostrzeżenia, a oryginał traci rzeczy, których nie modeluje (1.1, 2.7) (przyczyna) → **Nigdy nie nadpisuj wejścia.** Zapisuj do nowej nazwy albo rób backup `shutil.copy2` **przed** zapisem (naprawa).**

**9. „Skrypt działał, a teraz na serwerze połowa raportu jest zapisana" (objaw) → **zapis nieatomowy** — bezpośredni `wb.save(cel)` przerwany w połowie (brak miejsca, przerwane połączenie, restart) (przyczyna) → **Zapis do pliku tymczasowego w tym samym katalogu + `os.replace`.** Atomiczność działa tylko w obrębie jednego systemu plików (2.7) (naprawa).**

**10. „Po zapisie plik ma prawa `rw-------` i nikt z zespołu go nie otworzy" (objaw) → **`tempfile.mkstemp` tworzy plik z `0600`, a `os.replace` nadaje celowi uprawnienia pliku tymczasowego** (2.7) (przyczyna) → **`os.chmod(tmp, os.stat(cel).st_mode)` przed `os.replace`.** Zachowuj prawa poprzednika (naprawa).**

**11. „Skrypt uruchomiony dwa razy zdublował dane w raporcie" (objaw) → **brak idempotencji w trybie „append-only log"** (2.9) (przyczyna) → **Identyfikator z hash danych wejściowych, arkusz `_meta`, sprawdzenie przed dopisaniem.** Znacznik czasu jako identyfikator **nie działa** — każdy bieg jest „nowy" (naprawa).**

**12. „Chcę dopisać arkusz i zapisuję ten sam plik, który jest otwarty w Excelu — coś się zepsuło" (objaw) → **Excel trzyma plik zablokowany** i zapisuje swoją wersję przy zamknięciu, nadpisując Twoje zmiany; albo openpyxl zapisuje, a Excel trzyma otwarty uchwyt (przyczyna) → **Sprawdź, czy plik nie jest otwarty, zanim zapiszesz** (na Windows plik bywa zablokowany; na macOS/Linux mniej restrykcyjnie). W praktyce produkcyjnej: katalog wyjściowy **nie jest** katalogiem, z którego ludzie otwierają pliki, a nazwy zawierają timestamp (naprawa).**

**13. „Wszystkie obliczenia i wymiary komórek są inne niż w oryginale — plik wygląda brzydko" (objaw) → **wykresy są odbudowywane z modelu openpyxl**, a nietypowe formatowanie może zostać uproszczone (2.3) (przyczyna) → **To jest poprawne działanie biblioteki, nie błąd.** Traktuj wykresy jako «openpyxl potrafi częściowo» i **zawsze porównaj plik okiem.** Jeśli formatowanie jest krytyczne — użyj alternatywy z 2.5 (naprawa).**

**14. „Zniknęły wszystkie sparkline'y / slicery / timeline'y" (objaw) → **w rozszerzeniu `x14`** — openpyxl nie ma dla nich API i ostrzega przy wczytaniu (2.4) (przyczyna) → **Nie przepuszczaj takiego pliku przez openpyxl.** Jeśli musisz: odczytaj blok `x14` z XML-a i **wklej go z powrotem** do arkusza w wyniku (działa dla sparkline'ów — są samowystarczalne), albo użyj **`xlsxwriter`** do generowania od zera (naprawa).**

**15. „Zapisałem plik jako `.xlsm`, ale nie ma w nim makr; Excel zgłasza uszkodzenie" (objaw) → **openpyxl nie wymusza zgodności rozszerzenia z zawartością** (2.1), więc zapisał treść bez `vbaProject.bin` pod nazwą sugerującą makra (przyczyna) → **Dla `.xlsm`: `keep_vba=True` przy wczytaniu.** Dla plików bez makr: **używaj `.xlsx`**. Nie „przemianowuj" rozszerzeń ręcznie (naprawa).**

**16. „Kod z tutoriala wywala `AttributeError: __enter__`" (objaw) → **`Workbook` w 3.1.x NIE implementuje kontekst managera**, a w sieci krążą przykłady z `with load_workbook(...) as wb:` (2.2) (przyczyna) → **`try/finally` z `wb.close()`.** Pamiętaj, że `warnings.catch_warnings` **jest** kontekst managerem — `with` jest problemem tylko przy `Workbook` (naprawa).**

**17. „Analiza wykazała, że plik nie ma tabel ani reguł warunkowych — a w Excelu je widzę" (objaw) → **użyto `load_workbook(..., read_only=True)` w narzędziu diagnostycznym.** Reader robi wtedy `continue` przed parsowaniem arkusza, więc walidacji, CF, tabel i obrazów **nie ma w modelu** (moduł 17, 2.11) (przyczyna) → **Do diagnostyki i do inwentarza zawsze tryb normalny.** `read_only` tylko do strumieniowego czytania danych (naprawa).**

**18. „Zmieniłem jedno słowo w komórce i wszystkie komentarze w scalonych komórkach zniknęły" (objaw) → **komentarz w scalonej komórce** jest usuwany z ostrzeżeniem przy wczytaniu (moduł 16). Przy zapisie nie ma go w modelu (przyczyna) → **Wykryj to na wejściu** (nasłuch ostrzeżeń) i **przenieś komentarz na komórkę pomocniczą obok scalenia**, jeśli musi przetrwać (naprawa).**

**19. „Po migracji na 3.1 mój kod z nazwami zakresów przestał działać" (objaw) → **w 3.1 usunięto `get_named_range`, `add_named_range`, `remove_named_range`**, a weszło `DefinedName` + `wb.defined_names.add(...)` (2.3) (przyczyna) → **Przepisz na nowe API** i zaimportuj `from openpyxl.workbook.defined_name import DefinedName`. Zajrzyj do changelogu przed użyciem egzotycznych metod — kurs powtarza to od modułu 01 (naprawa).**

**20. „Mój wniosek brzmiał «nie ma ostrzeżeń, więc jest bezpiecznie»" (objaw) → **brak ostrzeżenia nie znaczy brak utraty.** Zmiana formatowania wykresu, zmiana pivot cache, uproszczenie formatowania warunkowego **nie generują żadnego komunikatu** (2.6) (przyczyna) → **Porównaj inwentarze ZIP i odciski modelu przed/po** (przykład 4), a na końcu **otwórz plik w Excelu i obejrzyj go okiem.** To jest trzeci raz, kiedy to mówimy — i nie bez powodu (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Zapis to przepisanie, nie edycja — i wszystko, czego skryba nie rozumie, nie trafia do nowego zeszytu.** `load_workbook()` buduje model w pamięci, `save()` serializuje **cały model od nowa**. Plik wynikowy jest **funkcją modelu**, nie pliku wejściowego. To nie błąd do naprawienia — to architektura. Wniosek: **utrata danych jest nieodwracalna**, a jedynym lekarstwem jest **nie przepuszczać przez openpyxl tego, czego nie chcesz stracić**.

2. **Sześć flag `load_workbook` i każda ma ciemną stronę.** `read_only` gubi wszystko poza danymi; `data_only=True` + `save()` **niszczy formuły**; `keep_vba=False` kasuje makra; `keep_links=False` daje `#REF!`; `rich_text=False` spłaszcza formatowanie fragmentów. A `with load_workbook(...)` **nie istnieje** — używaj `try/finally`. **Domyślne wartości są właściwe w 90% przypadków**; zmieniasz je tylko wtedy, gdy wiesz dlaczego.

3. **Inwentarz jest konkretny, nie „jakoś".** Przechodzi: wartości, formuły, daty, style, formaty liczb, tabele, filtry, walidacja standardowa, komentarze, hiperlinki, nazwy zdefiniowane, ustawienia druku, właściwości dokumentu. **Częściowo** (odbudowywane z modelu): wykresy — *sprawdź okiem*. **Ginie:** kształty, pola tekstowe, formanty, sparkline'y, slicery, timeline, Power Query, makra bez `keep_vba`, **każde nieznane rozszerzenie `extLst`**. I uwaga na dwie ciche utraty: komentarze w scalonych komórkach i kolejność serii w kombinacjach wykresów (→ „napraw plik").

4. **Masz trzy narzędzia, które zamieniają „mam nadzieję" na „wiem".** (a) **Inwentarz ZIP** — policz części i rozszerzenia *przed* i *po*. (b) **Nasłuch ostrzeżeń** — `catch_warnings(record=True)` + `simplefilter("always")`, bo openpyxl przez ten kanał mówi o utracie. (c) **Zapis atomowy** — plik tymczasowy w tym samym katalogu, `os.chmod`, `os.replace`, backup `copy2`, log audytowy JSONL, `raise` w `except`. Trzy narzędzia + **weryfikacja „na oko"** = kompletna ochrona.

5. **Reguła kciuka rządzi wszystkim.** *Jeśli plik był tworzony ręcznie przez człowieka i jest bogaty wizualnie — nie przepuszczaj go przez openpyxl.* Zamiast tego: **wczytaj dane ze źródła i wygeneruj nowy plik**, **użyj szablonu jako wzorca i zapisz pod nową nazwą**, **oddaj zapis Excelowi** (`xlwings`/COM), albo **zapisz zmiany w oddzielnym pliku i scalaj świadomie**. Do tego dwie zasady operacyjne: **nigdy nie nadpisuj wejścia** i **spraw, by skrypt był idempotentny** (identyfikator z hash danych, nie z zegara). Bo raport, który jest „prawie dobry", jest gorszy od raportu, który się nie wygenerował.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
import hashlib, json, os, shutil, tempfile, warnings, zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

# ==================================================================
# 2. WCZYTANIE - WŁAŚCIWE FLAGI (tabela z 2.2)
# ==================================================================
wb = load_workbook("plik.xlsx")                                  # danych + zapis
wb = load_workbook("plik.xlsm", keep_vba=True)                   # makra!
wb = load_workbook("plik.xlsx", rich_text=True)                  # rich text
wb = load_workbook("plik.xlsx")                                  # formuly (NIE data_only!)

# TYLKO do odczytu (i NIGDY potem save!):
wb_odczyt = load_workbook("plik.xlsx", data_only=True)

# Duze pliki, TYLKO odczyt (uwaga: brak DV, CF, obrazow, wykresow!):
wb_duzy = load_workbook("plik.xlsx", read_only=True)

# ⚠️ NIE MA KONTEKST MANAGERA dla Workbook w 3.1.x!
#    with load_workbook(...) as wb:  ->  AttributeError: __enter__
#    Uzywaj try/finally:
wb = load_workbook("plik.xlsx")
try:
    ...
finally:
    wb.close()          # istotne TYLKO w read_only / write_only

# ==================================================================
# 3. SKAN PRZED ZAPISEM (scan-before-save)
# ==================================================================
with zipfile.ZipFile("plik.xlsx") as z:
    czesci = z.namelist()

RYZYKO = ("xl/vbaProject.bin", "xl/charts/", "xl/drawings/vmlDrawing",
          "xl/media/", "xl/pivotTables/", "xl/pivotCache/",
          "customXml/", "xl/externalLinks/")
for c in czesci:
    if any(c.startswith(p) for p in RYZYKO):
        print("🔴", c)

# Rozszerzenia extLst (to ginie przy zapisie!):
import re
for name in czesci:
    if re.match(r"xl/worksheets/sheet\d+\.xml$", name):
        xml = zipfile.ZipFile("plik.xlsx").read(name).decode("utf-8", "replace")
        for guid in re.findall(r'<ext\s+uri="\{?([0-9A-Fa-f\-]+)\}?"', xml):
            print(f"🔴 {name}: rozszerzenie {guid} zostanie USUNIĘTE")

# Zrodlo prawdy o znanych rozszerzeniach (API prywatne - obuduj try/except):
try:
    from openpyxl.worksheet._reader import EXT_TYPES
except Exception:
    EXT_TYPES = {}

# ==================================================================
# 4. NASŁUCH OSTRZEŻEŃ (jedyny kanal informacji o utracie!)
# ==================================================================
with warnings.catch_warnings(record=True) as zlapane:
    warnings.simplefilter("always")          # bez tego tylko PIERWSZE!
    wb = load_workbook("plik.xlsx")          # NIE read_only!
    try:
        ...
    finally:
        wb.close()

for w in zlapane:
    print(w.category.__name__, w.message)
# Typowe tresci:
#   "Data Validation extension is not supported and will be removed"
#   "Sparkline Group extension is not supported and will be removed"
#   "Slicer List extension is not supported and will be removed"
#   "Unknown extension is not supported and will be removed"

# Wyciszenie JEDNEGO ostrzezenia, tylko dla odczytu:
with warnings.catch_warnings():
    warnings.filterwarnings("ignore",
                            message="Sparkline Group extension",
                            category=UserWarning,
                            module="openpyxl")
    wb = load_workbook("plik.xlsx")

# ==================================================================
# 5. BEZPIECZNY ZAPIS (wzorzec z 2.7)
# ==================================================================
def bezpieczny_zapis(wb, cel, *, backup=True):
    cel = Path(cel).resolve()
    cel.parent.mkdir(parents=True, exist_ok=True)

    # (1) BACKUP z timestampem (copy2 = zachowuje metadane)
    if cel.exists() and backup:
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        shutil.copy2(cel, cel.parent / f"{cel.stem}_{ts}{cel.suffix}")

    # (2) PLIK TYMCZASOWY W TYM SAMYM KATALOGU (atomowosc!)
    fd, name = tempfile.mkstemp(dir=cel.parent, prefix=".tmp-",
                               suffix=cel.suffix)
    os.close(fd)
    tmp = Path(name)
    try:
        wb.save(tmp)                                    # (3) ZAPIS
        if cel.exists():                                # (4) PRAWA
            os.chmod(tmp, os.stat(cel).st_mode)
        os.replace(tmp, cel)                            # (5) ATOMOWO
    except Exception:
        if tmp.exists():
            tmp.unlink()                                # (6) SPRZATANIE
        raise                                           # NIE POLYKAJ!
    finally:
        wb.close()

# ==================================================================
# 6. ODCISK SKOROSZYTU (test regresji formatowania)
# ==================================================================
def odcisk(path):
    wb = load_workbook(path)          # NIE read_only!
    try:
        kom = frm = reg = tab = wal = hid = 0
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        kom += 1
                        if isinstance(c.value, str) and c.value.startswith("="):
                            frm += 1
                    if c.protection and c.protection.hidden:
                        hid += 1
            reg += len(ws.conditional_formatting)
            tab += len(ws.tables)
            wal += len(ws.data_validations)
        return {"arkusze": len(wb.sheetnames), "komorki": kom, "formuly": frm,
                "reguly": reg, "tabele": tab, "walidacje": wal, "hidden": hid}
    finally:
        wb.close()

# ==================================================================
# 7. IDEMPOTENCJA (identyfikator z DANYCH, nie z zegara!)
# ==================================================================
def identyfikator(dane: str) -> str:
    return hashlib.sha256(dane.encode("utf-8")).hexdigest()[:16]

# w arkuszu "_meta": kolumny ["identyfikator", "kiedy", "wierszy"]
# przed dopisaniem: sprawdz, czy identyfikator juz istnieje -> jesli tak, KONIEC

# ==================================================================
# 8. WZORCE MODYFIKACJI (2.8) - wybierz swiadomie
# ==================================================================
# 1. read-modify-write  -> tylko dla plikow "czystych technicznie"
# 2. fill a template    -> ZALECANE: szablon nigdy nie jest zapisywany
# 3. append-only log    -> dopisywanie; wymaga idempotencji
# 4. sidecar            -> wyniki do nowego arkusza w kopii pliku

# ==================================================================
# 9. INNE
# ==================================================================
wb.properties.title = "Raport Q3"        # slad audytowy
wb.mime_type                             # openpyxl NIE wymusza rozszerzenia
chart._reindex()                         # PRYWATNE: naprawa idx/order w wykresach
wb.defined_names.add(DefinedName("N", attr_text="'A'!$A$1:$A$9"))  # API 3.1

# ==================================================================
# 10. TRZY ZDANIA DO ZAPAMIETANIA
# ==================================================================
# 1) Zapis to przepisanie calego modelu - co openpyxl nie rozumie, tego nie ma.
# 2) Nigdy nie nadpisuj pliku wejsciowego; zapisuj atomowo (tmp + os.replace).
# 3) Nie ma ostrzezenia != nie ma utraty. Porownaj inwentarze i sprawdz OKIEM.
```

## 9. Co dalej

Ten moduł był **sercem praktycznej wartości całego kursu** — i dlatego był też najdłuższy. Zbierzmy, co się w nim złożyło.

W **module 02** rozebrałeś `.xlsx` na części i zobaczyłeś, że skoroszyt to teczka z dokumentami. W **module 03** odkryłeś, że `save()` to przepisanie, a nie edycja. W **module 15** dowiedziałeś się, że wykres to przepis, nie potrawa. W **module 16** zobaczyłeś, że obraz to kopia bajtów, a nie adres — i że komentarz w scalonej komórce ginie. W **module 17** przeżyłeś cykl życia szablonu: Excel → blok `x14` → openpyxl → nic. **Wszystkie te cząstki składały się na jedno pytanie, które w tym module zadałeś do końca:**

> **Co dokładnie dzieje się z plikiem, który przechodzi przez openpyxl?**

I odpowiedź brzmi: **dokładnie to, co openpyxl potrafi odczytać i zapisać z powrotem.** Nic więcej. Ani jednej rzeczy więcej. Ani jednej rzeczy mniej. To jest model mentalny, do którego wracasz za każdym razem, gdy coś „nie działa" — i to jest też odpowiedź na pytanie, dlaczego ten moduł musiał stanąć **przed** wydajnością (19), testowaniem (20) i bezpieczeństwem (21), a nie po nich.

Trzy rzeczy, które zabierasz ze sobą:

- **Trzy kategorie prawdy jako nawyk, nie jako tabela.** Zanim cokolwiek zrobisz z cudzym plikiem, pytasz: *co przechodzi cało? co odbudowane? co zginie?* Ta trójka — powtarzana przy każdej operacji — to różnica między programistą, który „umie openpyxl", a inżynierem, któremu można powierzyć raport klienta.
- **Trzy narzędzia pomiaru.** `zipfile` (inwentarz części), `warnings.catch_warnings` (nasłuch), odcisk modelu (regresja). Wszystkie trzy mówią **liczbami**, a nie „wydaje się ok". Bez nich każda obietnica „nic nie zginęło" jest zgadywaniem — a przy raporcie klienta zgadywanie jest najgorszą z możliwych metod.
- **Jedna reguła kciuka, powiedziana trzy razy.** Jeśli plik jest od człowieka i bogaty wizualnie — **nie przepuszczaj go przez openpyxl**. Cztery alternatywy czekają w sekcji 2.5, a najważniejsza z nich stanie się tematem modułu 23: **szablon nigdy nie jest zapisywany**.

Jedna rzecz techniczna, którą warto zapamiętać na dłużej, bo **wróci w module 20 z pełną siłą**: `bezpieczny_zapis` i `odcisk` to **funkcje, które można przetestować bez pliku na dysku**. `wb.save(BytesIO())` i `load_workbook(BytesIO(...))` pozwalają sprawdzić cały potok w pamięci, a `BytesIO` jest jedynym „plikiem", jakiego potrzebujesz. To dlatego struktura `bezpieczny_zapis(wb, cel)` — przy całym swoim przecieku abstrakcji (przyjmuje typ openpyxl) — jest **lepsza niż** funkcja czytająca i zapisująca sama: rozdzielenie „zbuduj model" od „zapisz model" czyni obie połowy testowalnymi.

W **module 19** zmienimy skalę. Do tej pory rozmawialiśmy o plikach, które mieszczą się w pamięci w całości. Teraz zobaczysz, co się dzieje przy **100 000 wierszy**, **1 000 000 wierszy** i plikach większych niż RAM — i jak trzy tryby pracy z modułu 03 (`normalny`, `read_only`, `write_only`) przekładają się na **kolosalne** różnice w pamięci i czasie. Dowiesz się, kiedy `lxml` daje zysk, a kiedy nie, jak streamować dane z bazy prosto do arkusza, dlaczego `ProcessPoolExecutor` na **osobnych plikach** działa, a na jednym `wb` nie, i kiedy **openpyxl przestaje być właściwym narzędziem** — bo nie wozi się węgla ferrari. Zobaczysz też realne pomiary: tabelę „operacja → czas" dla 10 000, 100 000 i 1 000 000 wierszy, którą zbudujesz własnymi rękami.