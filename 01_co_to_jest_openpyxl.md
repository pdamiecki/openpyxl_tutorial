# Moduł 01 — Czym jest openpyxl i kiedy warto go użyć

> **Część:** 0 — Fundamenty i model mentalny · **Poziom:** ⭐ · **Wymaga:** modułu 00

## 0. W tym module nauczysz się

- Rozumiesz, co openpyxl **konkretnie potrafi**: jakie formaty plików czyta i zapisuje, czego potrzebuje do pracy, a czego nie potrzebuje.
- Znasz kluczowe rozróżnienie architektoniczne: openpyxl implementuje **format pliku**, a nie **program Excel** — i potrafisz wyciągnąć z tego praktyczne wnioski.
- Umiesz odróżnić **„zapisuję formułę”** od **„obliczam formułę”** i wiesz, dlaczego to dwie zupełnie różne rzeczy, które ludzie ciągle mylą.
- Rozpoznajesz trzy kategorie prawdy o API openpyxl: **potrafi** / **potrafi częściowo** / **gubi przy zapisie** — i wiesz, gdzie szukać szczegółów.
- Umiesz świadomie wybrać narzędzie: openpyxl kontra `xlsxwriter`, `pandas`, `xlwings` i `pyexcel` — na podstawie konkretnego zadania, a nie mody.
- Wiesz, co się zmieniło w wersji 3.1 (i dlaczego tutoriale z 2018 roku nie działają) oraz dlaczego 3.1.x jest punktem odniesienia w tym kursie.
- Umiesz napisać skrypt, który **przed użyciem biblioteki wykrywa ryzyko** w pliku wejściowym: plik `.xls`, makra, sparkline'y, slicery.

## 1. Intuicja i analogia

Zacznijmy od rozdzielenia dwóch rzeczy, które w potocznej rozmowie zlewają się w jedno: **program Excel** i **plik `.xlsx`**.

Program Excel to aplikacja. Ma okna, przyciski, wstążkę, silnik przeliczający formuły, kreator wykresów, edytor VBA i wątek rysujący wszystko na ekranie. Plik `.xlsx` to dokument — teczka z liczbami, tekstami, formułami i opisami wyglądu. Program to narzędzie, plik to materiał. Tak jak „program do obróbki zdjęć” i „plik `.jpg`” to nie to samo.

openpyxl pracuje z **plikiem**. Nie ma nic wspólnego z programem Excel. Nie uruchamia go, nie łączy się z nim, nie wie nawet, czy Excel jest w ogóle zainstalowany na Twoim komputerze — i w większości przypadków nie jest, bo openpyxl najczęściej działa na serwerze Linuksowym bez żadnego środowiska graficznego.

**Analogia centralna tego modułu — i całego działu o ograniczeniach:**

> openpyxl to **tłumacz i maszynista, nie samolot**.
>
> Wyobraź sobie międzynarodowe lotnisko. Samolot to Excel: tylko on faktycznie lata, czyli tylko on umie policzyć formuły, odświeżyć tabelę przestawną i wyświetlić wykres na ekranie.
>
> Ale samolot nie startuje bez dokumentów. Potrzebny jest plan lotu, karty załadunku, deklaracja pasażerów, wszystko wypełnione dokładnie według ustalonego formularza lotniczego. I tu wchodzi openpyxl: to człowiek, który **umie wypełnić każdy z tych dokumentów perfekcyjnie i według przepisów**, umie je też odczytać, poprawić i przygotować nowy zestaw.
>
> Czego ten człowiek nie zrobi? **Nie poleci.** Nie ma silników, nie ma skrzydeł i nie jest to jego zadanie. Jeśli zapytasz go „ile paliwa zużyje ten lot?”, odpowie: „nie moja działka — to policzy samolot, ja tylko zapiszę liczbę, jeśli ktoś mi ją poda”.

To jedno zdanie wyjaśnia 80% frustracji początkujących: **piszesz formułę, ale nie dostajesz wyniku**. Nie dlatego, że zrobiłeś coś źle. Dlatego, że poprosiłeś tłumacza o latanie.

Analogia pomocnicza, przydatna w drugiej części modułu — do wyboru narzędzia:

> openpyxl to **zecer w starej drukarni**. Zecer układa tekst, dobiera czcionki, wie, gdzie postawić nagłówek, potrafi dopisać stronę do istniejącej już książki i wydrukować nową wersję. Ale zecer nie jest księgarzem, który wie, co klient chce kupić, ani drukarzem szybkiej prasy, która wypuszcza tysiące stron w minutę. Do szybkiego druku używa się innej maszyny — i o tym jest cała sekcja „jak wybrać bibliotekę”.

Trzecia analogia, którą zapamiętasz na cały kurs, bo wraca w każdym module:

> openpyxl **nie edytuje pliku — przepisuje go na czysto**. Jak kancelaria, która dostaje oryginał aktu, odczytuje go i wystawia nowy dokument. Jeśli kancelaria nie umie odczytać jakiejś adnotacji (naklejka, pieczątka w nieznanym języku, odręczny dopisek na marginesie), to tej adnotacji **nie będzie w nowym dokumencie**. Nie dlatego, że ktoś ją złośliwie usunął. Dlatego, że nigdy nie została odczytana.

To jest źródło najważniejszego ostrzeżenia w całym kursie: „openpyxl gubi dane przy zapisie” nie znaczy „openpyxl jest zepsuty”. Znaczy „openpyxl nie zna tej części formatu, więc nie ma jej czym przepisać”. Szczegóły w module 18.

## 2. Teoria

### 2.1. Co openpyxl konkretnie robi

Lista jest krótka, ale precyzyjna. Nie „potężna biblioteka” — konkretne możliwości:

1. **Tworzy nowe skoroszyty od zera** i zapisuje je jako `.xlsx`.
2. **Otwiera istniejące skoroszyty** (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`), rozkłada je na obiekty Pythona i pozwala je modyfikować.
3. **Czyta i zapisuje wartości komórek**: tekst, liczby, wartości logiczne, daty, czas, czasy trwania.
4. **Zapisuje formuły** jako literał tekstowy (bez obliczania ich).
5. **Formatuje komórki**: czcionki, kolory, obramowania, wyrównanie, formaty liczb.
6. **Formatuje arkusze**: szerokości kolumn, wysokości wierszy, scalanie komórek, zamrażanie okien, ustawienia wydruku, obszary wydruku.
7. **Tworzy tabele Excela** (`Table`) ze stylami i filtrami.
8. **Tworzy formatowanie warunkowe**: reguły komórkowe, formułowe, skale kolorów, paski danych, zestawy ikon.
9. **Tworzy wykresy**: słupkowe, liniowe, kołowe, punktowe, warstwowe, bąbelkowe, radarowe i inne — z osiami, tytułami, legendą i etykietami danych.
10. **Wstawia obrazy** (wymaga `pillow`), **hiperlinki** i **komentarze**.
11. **Dodaje walidację danych** (listy rozwijane, ograniczenia liczbowe i tekstowe) oraz **ochronę** arkusza i skoroszytu.
12. **Odczyta i zapisze z powrotem** większość elementów istniejącego pliku, o ile zna ich strukturę.
13. **Działa w trybach strumieniowych** dla bardzo dużych plików: `read_only` (czytanie wiersz po wierszu) i `write_only` (pisanie wiersz po wierszu bez trzymania całości w pamięci).
14. **Działa całkowicie bez sieci i bez Excela** — na serwerze, w kontenerze Dockera, w GitHub Actions, na Raspberry Pi.

### 2.2. Czego openpyxl nie robi

Ta lista jest równie ważna, a bywa pomijana w tutorialach:

1. **Nie oblicza formuł.** Nigdy. To nie jest luka, którą ktoś kiedyś załata — to świadoma decyzja architektoniczna (wyjaśnienie w 2.3). Potrafi odczytać **wynik zapisany w pliku** przez Excela (`data_only=True`), ale sam go nie wytworzy.
2. **Nie obsługuje formatu `.xls`** (Excel 97–2003). To inny, binarny format. Ani odczyt, ani zapis.
3. **Nie obsługuje formatu `.xlsb`** (binarny skoroszyt Excela). Do odczytu `.xlsb` istnieje osobna biblioteka `pyxlsb` (tylko odczyt).
4. **Nie obsługuje formatu `.ods`** (LibreOffice). Do tego jest `odfpy` albo `pyexcel`.
5. **Nie tworzy sparkline'ów.** To funkcja Excela 2010+, zapisana w bloku rozszerzeń `x14`, którego openpyxl nie modeluje.
6. **Nie tworzy nowych tabel przestawnych** (tabelę przestawną potrafi wczytać i zachować, ale nie zbudować ani odświeżyć jej pamięci podręcznej).
7. **Nie tworzy ani nie zachowuje kształtów, pól tekstowych i formantów** (przyciski, pola wyboru, elementy ActiveX).
8. **Nie obsługuje slicerów, osi czasu ani modelu danych Power Pivot / Power Query.**
9. **Nie wykonuje makr VBA.** Może je zachować w pliku `.xlsm` przy `keep_vba=True`, ale nie uruchomi ani nie zmodyfikuje ich kodu.
10. **Nie otwiera pliku w Excelu** i nie sprawdza, czy plik jest poprawny „z punktu widzenia Excela”. Jeśli wygenerujesz coś, czego Excel nie zaakceptuje, dowiesz się o tym dopiero, gdy otworzysz plik ręcznie.
11. **Nie jest edytorem tekstu sformatowanego w pełnym znaczeniu.** Obsługuje tekst bogaty (`CellRichText`, nowość 3.1), ale nie jest to zamiennik Worda.
12. **Nie jest bazą danych.** Wczyta plik do pamięci w całości (w trybie normalnym), więc plik 2 GB nie jest dla niego dobrym zadaniem.

### 2.3. Kluczowe rozróżnienie architektoniczne: format, nie program

Dlaczego openpyxl nie liczy formuł? Bo **implementuje format pliku, a nie program**.

Format `.xlsx` (formalnie: OOXML SpreadsheetML) to specyfikacja opisująca, jak zapisać w XML liczby, teksty, formuły i opisy wyglądu. Ta specyfikacja mówi: „formuła to tekst zaczynający się od znaku `=`”. **Nie mówi nic o tym, jak tę formułę policzyć.** Semantyka formuł (co znaczy `SUM`, `XLOOKUP`, `SUMPRODUCT`, jak Excel zaokrągla, jak działa arytmetyka dat) to domena **programu**, a nie dokumentu.

Analogia do prawa: specyfikacja formatu to **kodeks przepisów o formie dokumentu** (jak wypełnić rubryki, gdzie postawić pieczątkę). Jak policzyć podatek, który wpiszesz w rubrykę — to osobne przepisy, których nie ma w instrukcji wypełniania formularza. openpyxl zna instrukcję wypełniania. Nie zna przepisów podatkowych.

Praktyczna konsekwencja, którą trzeba sobie mocno wbić do głowy:

> Biblioteka implementująca **format** daje Ci **wierność zapisu**. Biblioteka implementująca **program** dałaby Ci **wierność obliczeń**. openpyxl wybrał pierwsze. Dlatego jest szybki, lekki, działa wszędzie i nie wymaga Excela. Zapłaciłeś za to brakiem silnika obliczeń.

Trzy konkretne konsekwencje tej decyzji:

**Konsekwencja 1: brak silnika obliczeń.** Wpisujesz `"=SUM(A1:A10)"` i w pliku pojawia się formuła. Wartość pojawi się, gdy plik otworzy Excel, LibreOffice albo inne narzędzie z silnikiem. Jeśli Twój kod Pythona potrzebuje tej liczby *teraz*, musisz ją policzyć sam.

Dobra wiadomość: openpyxl domyślnie zapisuje w skoroszycie ustawienie `fullCalcOnLoad`, które mówi Excelowi „przelicz to wszystko od nowa przy otwarciu”. Użytkownik dostanie poprawne liczby — Ty, w Pythonie, nie.

**Konsekwencja 2: brak pełnego modelu formatowania wykresów.** Wykres w pliku `.xlsx` to nie obrazek. To opis: „weź zakres `B2:B13` jako wartości, zakres `A2:A13` jako kategorie, narysuj słupki, oś Y wyskalowana automatycznie”. Openpyxl ten opis rozumie i potrafi go odtworzyć. Ale formatowanie wykresu w Excelu obejmuje setki opcji (gradienty, cienie, niestandardowe etykiety, efekty 3D), z których openpyxl modeluje tylko część. Efekt: **wykres wczyta się i zapisze, ale część ozdobników może zostać uproszczona**. To właśnie kategoria „potrafi częściowo”.

**Konsekwencja 3: brak modelu elementów rozszerzeń.** Nowsze funkcje Excela (sparkline'y, część opcji pasków danych, slicery) nie są zapisywane w podstawowej strukturze arkusza, tylko w **blokach rozszerzeń** (`extLst`) — czyli w miejscu przeznaczonym na „różne rzeczy, które dopiszemy później”. openpyxl czyta te bloki jako dane, których nie rozumie, i wypisuje ostrzeżenie na standardowe wyjście błędów. Przy zapisie ich nie ma. To kategoria „gubi przy zapisie”.

### 2.4. Trzy kategorie prawdy o API — tabela na cały kurs

W tym kursie będziemy konsekwentnie mówić o każdej funkcji w jednej z trzech kategorii. To najważniejsze narzędzie pojęciowe, jakie zabierzesz z tego modułu.

| Kategoria | Znaczenie | Konsekwencja praktyczna |
|-----------|-----------|--------------------------|
| **potrafi** | openpyxl w pełni obsługuje tę funkcję: i przy tworzeniu, i przy odczycie, i przy zapisie | Możesz na tym budować logikę bez obaw |
| **potrafi częściowo** | Obsługiwany jest podstawowy model, ale szczegóły mogą zostać uproszczone lub zgubione | **Zawsze sprawdź efekt okiem** w Excelu/LibreOffice |
| **gubi przy zapisie** | Elementu nie ma w modelu openpyxl, więc przy zapisie nie istnieje | Nie przepuszczaj przez openpyxl plików, które tego używają |

Zestawienie dla najważniejszych elementów pliku Excel:

| Element pliku | Kategoria | Uwagi |
|---------------|-----------|-------|
| Wartości, teksty, liczby, wartości logiczne | **potrafi** | W pełni; uwaga na typy (moduł 05) |
| Daty, czasy, czasy trwania | **potrafi** | Z wyjątkiem stref czasowych (`tzinfo` nie jest zapisywane) |
| Formuły | **potrafi** (zapis), **nie oblicza** | Excel policzy przy otwarciu |
| Style: czcionki, kolory, obramowania | **potrafi** | Model stylu w module 09 |
| Formaty liczb | **potrafi** | Kody formatów w module 10 |
| Scalanie komórek, wymiary, freeze panes | **potrafi** | Uwaga: scalanie niszczy treść poza lewą górną komórką |
| Ustawienia wydruku, obszary wydruku, tytuły | **potrafi** | Moduł 11 |
| Tabele (`Table`), filtry | **potrafi** | Moduł 12 |
| Formatowanie warunkowe (reguły, skale, paski, ikony) | **potrafi** | Niektóre opcje pasków danych leżą w `extLst` — mogą ucierpieć |
| Komentarze, hiperlinki | **potrafi** | Moduł 16 |
| Obrazy | **potrafi** | Wymaga `pillow`; wymaga, by plik źródłowy obrazu istniał w chwili zapisu |
| Definiowane nazwy (named ranges) | **potrafi** | Nowe API w 3.1: `DefinedName`, `wb.defined_names.add(...)` |
| Wykresy | **potrafi częściowo** | Wczytywane i zapisywane, ale **odbudowywane z modelu openpyxl** — część formatowania może zniknąć |
| Właściwości dokumentu, w tym własne | **potrafi** | Własne właściwości to nowość 3.1 |
| Tabele przestawne (pivot) | **potrafi częściowo** | Odczyt i zachowanie; **brak tworzenia i odświeżania danych** |
| Makra VBA (`.xlsm`) | **potrafi częściowo** | Tylko z `keep_vba=True`; zachowane, ale niedostępne do edycji |
| **Sparkline'y** | **gubi przy zapisie** | Blok rozszerzeń `x14`; openpyxl ostrzega przy wczytaniu |
| Kształty, pola tekstowe | **gubi przy zapisie** | Dokumentacja openpyxl ostrzega o tym wprost |
| Formanty (przyciski, pola wyboru, ActiveX) | **gubi przy zapisie** | Poza modelem |
| Slicery, osie czasu (timeline) | **gubi przy zapisie** | Poza modelem |
| Power Pivot / Power Query / model danych | **gubi przy zapisie** | Poza modelem |

Nie musisz tego pamiętać teraz. Wróć do tej tabeli po module 18 — wtedy wszystko się zazębi.

### 2.5. Kiedy openpyxl jest właściwym wyborem

Poniżej realne scenariusze. Zwróć uwagę, że w każdym z nich openpyxl wygrywa **z konkretnego powodu**, a nie „bo jest dobry”.

| Scenariusz | Dlaczego openpyxl |
|------------|-------------------|
| Nocny raport generowany na serwerze Linuksowym, wysyłany mailem do 200 osób | Nie wymaga Excela, działa w kontenerze, cały proces jest deterministyczny |
| Wypełnianie firmowego szablonu z logo i pięcioma zakładkami, gdzie zmieniają się tylko dane | **Jedyna z popularnych bibliotek, która potrafi otworzyć istniejący plik i zapisać go z powrotem**, zachowując niezmienione części |
| Dopisanie kolumny z wyliczonym wskaźnikiem do istniejącego raportu | To operacja read-modify-write — domena openpyxl |
| Raport z tabelami, formatowaniem warunkowym, wykresem i komentarzami dla odbiorcy | Wszystko w jednym pliku, jedno narzędzie, bez łańcuchów zależności |
| Czyszczenie i normalizacja brudnych danych z arkusza | Pełny dostęp do komórek i typów, kontrola nad tym, co trafia do pliku |
| API, które zwraca wygenerowany plik w odpowiedzi HTTP | `wb.save(BytesIO())` — plik nigdy nie dotyka dysku |
| Zamiana setek plików `.xls` na `.xlsx` | Tylko z pomocą LibreOffice headless do samej konwersji, ale openpyxl wykonuje potem całą obróbkę |
| Wyszukanie konkretnych wartości w tysiącach skoroszytów (audyt, compliance) | Tryb `read_only` czyta duże pliki strumieniowo |

### 2.6. Kiedy openpyxl **nie** jest właściwym wyborem

Ta tabela jest ważniejsza od poprzedniej, bo chroni Cię przed tygodniami niepotrzebnej pracy.

| Twoja potrzeba | Właściwe narzędzie | Dlaczego nie openpyxl |
|----------------|--------------------|------------------------|
| Potrzebujesz **obliczonych** wartości formuł w Pythonie | Excel albo LibreOffice (przeliczenie), biblioteka `formulas` lub `pycel` (tylko proste przypadki) | openpyxl nie ma silnika obliczeń |
| Potrzebujesz **sparkline'ów** w nowym pliku | `xlsxwriter` | openpyxl nie ma API do sparkline'ów i gubi je przy zapisie |
| Potrzebujesz **największej kontroli nad formatowaniem i wykresami** przy tworzeniu pliku od zera | `xlsxwriter` | Bogatsze API formatowania i wykresów; tryb `constant_memory` dla ogromnych plików |
| Chcesz **wpisać dane do otwartego Excela**, uruchomić makro, przeliczyć skoroszyt | `xlwings` (albo `pywin32` na Windows) | openpyxl nie komunikuje się z programem Excel w ogóle |
| Chcesz **przeanalizować dane** z arkusza (filtry, grupowania, statystyki) | `pandas` (+ `polars`/`duckdb` dla bardzo dużych zbiorów) | openpyxl to narzędzie do plików, nie do analizy danych |
| Chcesz **jeden interfejs do wielu formatów** (`xls`, `xlsx`, `ods`, `csv`) | `pyexcel` jako warstwa nad innymi bibliotekami | openpyxl obsługuje wyłącznie rodzinę `.xlsx` |
| Potrzebujesz **konwersji `.xls` → `.xlsx`** | LibreOffice headless (`soffice --convert-to`) | openpyxl nie czyta `.xls` |
| Masz **plik projektowany ręcznie w Excelu**: dashboard ze sparkline'ami, slicerami, kształtami, Power Query | Zostaw go Excelowi. Skrypt może dostarczać **dane** do osobnego arkusza, ale nie zapisywać pliku projektowanego | To dokładnie ta kategoria „gubi przy zapisie” |

To ostatnie zalecenie warto rozwinąć, bo ratuje projekty. Wzorzec nazywa się „rozdzielenie danych od projektu”: plik z dashboardem czyta liczby z arkusza (albo z innego pliku), który Twój skrypt generuje w całości. Skrypt nigdy nie dotyka pliku z dashboardem. Wtedy sparkline'y, slicery i makra są bezpieczne, bo nikt ich nie przepisuje.

### 2.7. Tabela porównawcza bibliotek

Poniższa tabela jest najczęściej wracającym fragmentem tego kursu. Czytaj ją jak zestaw **zadań**, nie jak ranking — żadna kolumna nie jest „lepsza”.

| Biblioteka | Czyta istniejące | Pisze nowe | Wymaga Excela | Liczy formuły | Wykresy | Sparkline'y | Tryb strumieniowy | Typowy scenariusz |
|------------|------------------|------------|---------------|---------------|---------|-------------|-------------------|-------------------|
| **openpyxl** | tak (`.xlsx`, `.xlsm`, `.xltx`, `.xltm`) | tak | **nie** | **nie** (zapisuje; odczyt tylko z pamięci podręcznej) | tak, podstawowe, odbudowywane z modelu | **nie** | tak: `read_only`, `write_only` | modyfikacja istniejących plików, szablony, raporty na serwerze |
| **xlsxwriter** | **nie** (write-only z założenia) | tak | nie | nie (można zapisać gotowy wynik obok formuły) | tak, bogate | **tak** | tak: `constant_memory` | nowe, bogato sformatowane raporty i dashboardy od zera |
| **pandas** | tak (przez silniki: openpyxl, calamine, xlrd dla `.xls`) | tak (silniki: openpyxl, xlsxwriter) | nie | nie (czyta wartości z pamięci podręcznej) | nie | nie | tak (`chunksize` przy odczycie) | analiza danych, szybki eksport `DataFrame` |
| **xlwings** | tak (przez program Excel) | tak (przez program Excel) | **tak** (Excel na Windows/macOS) | **tak** (przelicza prawdziwy Excel) | wszystko, co potrafi Excel | tak (bo to Excel) | nie (wolno, przez COM/AppleScript) | interaktywna automatyzacja, przeliczanie, makra, funkcje UDF |
| **pyexcel** | tak (wiele formatów: `.xls`, `.xlsx`, `.ods`, `.csv`) | tak | nie | nie | nie | nie | ograniczone | jednolite API do wielu formatów; projekt rozwijany znacznie wolniej niż openpyxl |

Do tego trzy uzupełnienia, które nie zmieściły się w tabeli:

- **Do odczytu `.xlsb`** służy `pyxlsb` (tylko odczyt, bez zapisu).
- **Do formatu `.ods`** (LibreOffice) służy `odfpy` albo `pyexcel`, nie openpyxl.
- **Do obliczania formuł w Pythonie** istnieją biblioteki `formulas` i `pycel`. Działają na wycinku języka formuł Excela: proste arytmetyki, popularne funkcje, zależności między komórkami. Zawodzą na tabelach przestawnych, Power Query, odwołaniach zewnętrznych i własnych funkcjach VBA. Traktuj je jako eksperyment, nie jako zamiennik Excela.

### 2.8. Historia i wersjonowanie — dlaczego 3.1.x

openpyxl rozwija się od 2010 roku. Dla praktyki liczą się cztery daty:

| Wersja | Data | Co wnosi |
|--------|------|----------|
| 2.6.4 | 2019-09-25 | Ostatnie wydanie obsługujące Pythona 2.7 i 3.5 |
| **3.0.0** | 2019-09-25 | Tylko Python 3.6+; nowa struktura wewnętrzna |
| **3.1.0** | 2023-01-31 | Tekst bogaty (`CellRichText`), nowe API definiowanych nazw (`DefinedName`), własne właściwości dokumentu, ulepszone mapowanie właściwości graficznych wykresów, poprawki stylów nazwanych. **Usunięcia** (poniżej) |
| 3.1.4 | 2024-06-12 | **Koniec wsparcia dla Pythona 3.6 i 3.7** |
| **3.1.5** | 2024-06-28 | Najnowsze stabilne wydanie linii 3.1.x; wymaga Pythona ≥ 3.8 |

Uwaga praktyczna: istnieje wydanie `3.2.0b1` z 2022 roku, ale zostało wycofane (yanked) z PyPI. **Linia 3.1.x pozostaje punktem odniesienia** i to jej dotyczy cały kurs. Jeśli czytasz ten kurs później niż w chwili jego pisania (październik 2026), zajrzyj na stronę projektu na PyPI i sprawdź, czy nie pojawiła się nowsza linia — i przeczytaj jej changelog, zanim zastosujesz jego przykłady.

**Co dokładnie usunięto w 3.1.0** — to lista „dlaczego tutorial z 2018 roku nie działa”:

1. Z klasy `Workbook` usunięto przestarzałe metody definiowanych nazw: `get_named_range`, `add_named_range`, `remove_named_range`. Zamiast nich używa się obiektów `DefinedName` i kolekcji `wb.defined_names`.
2. Z obrazów usunięto metodę `get_emu_dimensions`.
3. Z arkuszy usunięto przestarzałe właściwości: `formula_attributes`, `page_breaks`, `show_summary_below`, `show_summary_right`, `page_size` i `orientation`. Kod kliencki ma używać odpowiednich **obiektów** — na przykład zamiast `ws.orientation` ustawia się `ws.page_setup.orientation`, a zamiast `ws.page_size` — `ws.page_setup.paperSize`. Podziały stron obsługuje `ws.row_breaks` / `ws.col_breaks` (moduł 11).
4. Metoda `Workbook.create_named_range` jest oznaczona jako przestarzała; zalecane jest bezpośrednie przypisywanie nazw do arkusza (nazwy lokalne) albo do skoroszytu (nazwy globalne).
5. Wcześniej (w linii 3.0) zniknęły również stare akcesory arkuszy typu `wb.get_sheet_by_name(...)`, `wb.get_sheet_names()` i `wb.get_sheet_by_index(...)`. Dziś: `wb["Nazwa"]`, `wb.sheetnames`, `wb.worksheets[i]`. Warto to sprawdzić w changelogu swojej wersji, bo te nazwy pojawiają się w ogromnej liczbie starych przykładów w internecie.

I jeszcze jedna rzecz, która myli:

```python
# POPRAWNIE (od wersji 3.1):
from openpyxl.utils.dataframe import dataframe_to_rows
```

Funkcja konwertująca `DataFrame` (pandas) na wiersze nadające się do arkusza mieszka w `openpyxl.utils.dataframe`. Jeśli natkniesz się na tutorial, który importuje ją z innej, „krótszej” ścieżki — ten tutorial jest przestarzały albo dotyczy innej wersji. Zawsze sprawdzaj ścieżkę importu w dokumentacji swojej wersji.

**Dlaczego wersja ma znaczenie w tym kursie:** openpyxl zmieniał się ewolucyjnie, a jego API bywa „ciche”. Metoda, która jeszcze istnieje, może być oznaczona jako przestarzała i zniknąć w kolejnym wydaniu. Dlatego w tym kursie zawsze podajemy wersję, a przy mniej oczywistych wywołaniach dodajemy komentarz „sprawdź w dokumentacji swojej wersji”.

### 2.9. Mit do obalenia: „openpyxl oblicza formuły”

Zajmijmy się tym mitem raz na zawsze, bo kosztuje ludzi tygodnie debugowania.

**Skąd się bierze?** Z trzech źródeł. Po pierwsze, openpyxl domyślnie zapisuje w skoroszycie flagę `fullCalcOnLoad` (właściwość `CalcProperties.fullCalcOnLoad`), która sprawia, że **Excel** przelicza wszystko przy otwarciu. Użytkownik otwiera plik i widzi poprawne liczby — więc wydaje się, że „openpyxl to policzył”. Po drugie, gdy openpyxl wczytuje plik, który **kiedyś był otwarty i zapisany przez Excela**, w pliku są już zapisane **wartości z pamięci podręcznej** obok formuł. Wtedy `data_only=True` zwraca liczby — i znowu wygląda to jak obliczenie. Po trzecie, dokumentacja mówi o „data-only mode”, co brzmi jak „tryb liczenia”.

**Prawda:** openpyxl nigdy nie wykonuje żadnych obliczeń formuł. Może jedynie odczytać wartość, którą ktoś inny (Excel, LibreOffice) już zapisał w pliku.

Zobaczmy to na konkretnym kodzie — to nasz pierwszy pełny przykład.

### 2.10. Jak „odświeżyć” formuły bez Excela — uczciwy przegląd

Skoro openpyxl nie liczy, a Ty potrzebujesz liczby, masz pięć dróg. Wymieniam je z realnymi ograniczeniami, bo to pytanie wraca jak bumerang.

| Droga | Jak działa | Ograniczenia |
|-------|------------|--------------|
| **Policz w Pythonie** | Zamiast `=SUM(A1:A10)` piszesz `sum(...)` i zapisujesz wynik jako liczbę | Tracisz „żywe” formuły — użytkownik nie może zmienić danych i zobaczyć aktualizacji. Czasem to zaleta, czasem wada |
| **Zostaw formuły, licz na Excela u odbiorcy** | Zapisujesz formuły; Excel przy otwarciu przelicza (bo `fullCalcOnLoad=True`) | Twoja logika Pythona nie zna wyniku — nie możesz na nim opierać dalszych decyzji |
| **LibreOffice headless** | `soffice --headless --convert-to xlsx plik.xlsx` na serwerze | LibreOffice musi być zainstalowany; przeliczanie przy konwersji **zależy od ustawień** i nie jest gwarantowane — trzeba to sprawdzić na własnym pliku. Konwersja zmienia też część formatowania |
| **`formulas` lub `pycel`** | Biblioteki Pythona parsujące i obliczające wycink języka formuł | Obsługują podzbiór: proste funkcje i zależności. Nie policzą tabel przestawnych, Power Query, odwołań zewnętrznych, funkcji VBA |
| **`xlwings` / COM** | Otwiera prawdziwy Excel (Windows/macOS), pisze dane, zapisuje przez Excela | Wymaga Excela zainstalowanego. Na Linuksie bez Excela — odpada |

**Rekomendacja kursu:** domyślnie wybieraj drogę pierwszą lub drugą, świadomie. Jeśli dana liczba jest potrzebna Twojemu kodowi — policz ją w Pythonie. Jeśli dana liczba ma być „żywa” dla użytkownika — zostaw formułę i nie oczekuj, że ją odczytasz. Mieszanie tych dwóch światów to źródło najdziwniejszych błędów. Wrócimy do tego w module 06, gdzie rozbierzemy `data_only` na części i pokażemy, jak **nie zniszczyć** pliku.

## 3. Przykłady krok po kroku

Wszystkie przykłady zapisują do `output/`, zgodnie z konwencją modułu 00.

### Przykład 1 — dowód, że openpyxl nie liczy formuł

Ten przykład jest sercem modułu. Uruchom go i **otwórz wygenerowany plik w Excelu** — zobaczysz, że tam liczba jest. Potem wróć do wyników Pythona i zobacz, że tam jej nie ma.

```python
"""Dowód, że openpyxl zapisuje formułę, ale jej nie oblicza.

Uruchom:  python examples/01_formula_dowod.py
Efekt:    output/01_formula_dowod.xlsx
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"


def glowny() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    sciezka = OUTPUT / "01_formula_dowod.xlsx"

    # --- KROK 1: tworzymy skoroszyt i wpisujemy liczby ORAZ formułę.
    wb = Workbook()
    ws = wb.active
    ws.title = "Dowod"

    ws["A1"] = 2
    ws["A2"] = 3
    # Wartość zaczynająca się od "=" jest traktowana jako FORMUŁA, nie jako tekst.
    # W pamięci to zwykły łańcuch znaków. W pliku to element <f> w XML komórki.
    ws["A3"] = "=A1+A2"

    # (Ciekawostka: openpyxl domyślnie ustawia w skoroszycie flagę
    #  fullCalcOnLoad - dlatego Excel przeliczy formuły przy otwarciu.
    #  Sprawdź nazwę atrybutu w dokumentacji swojej wersji.)
    wb.save(sciezka)
    print(f"Zapisano: {sciezka}\n")

    # --- KROK 2: odczyt DOMYŚLNY. Dostajemy literał formuły.
    wb_formula = load_workbook(sciezka)
    ws_formula = wb_formula["Dowod"]
    print("Odczyt domyślny (data_only=False):")
    print(f"  A1 = {ws_formula['A1'].value!r}")
    print(f"  A3 = {ws_formula['A3'].value!r}   <- to jest TEKST formuły")
    print(f"  typ danych komórki A3: {ws_formula['A3'].data_type!r}\n")

    # --- KROK 3: odczyt z data_only=True. Próbujemy sięgnąć do pamięci podręcznej.
    wb_values = load_workbook(sciezka, data_only=True)
    ws_values = wb_values["Dowod"]
    print("Odczyt z data_only=True:")
    print(f"  A1 = {ws_values['A1'].value!r}")
    print(f"  A3 = {ws_values['A3'].value!r}   <- None, bo pliku nie widział żaden Excel!\n")

    print("TERAZ otwórz plik w Excelu, zapisz go (Ctrl+S) i uruchom skrypt ponownie.")
    print("Zobaczysz, że data_only=True zwróci 5 - bo Excel zapisał wartość do pliku.")


if __name__ == "__main__":
    glowny()
```

Oczekiwany wynik pierwszego uruchomienia:

```text
Zapisano: /home/anna/projekty/kurs-openpyxl/output/01_formula_dowod.xlsx

Odczyt domyślny (data_only=False):
  A1 = 2
  A3 = '=A1+A2'   <- to jest TEKST formuły
  typ danych komórki A3: 'f'

Odczyt z data_only=True:
  A1 = 2
  A3 = None   <- None, bo pliku nie widział żaden Excel!

TERAZ otwórz plik w Excelu, zapisz go (Ctrl+S) i uruchom skrypt ponownie.
Zobaczysz, że data_only=True zwróci 5 - bo Excel zapisał wartość do pliku.
```

**Co się dzieje w pamięci:** `ws["A3"] = "=A1+A2"` tworzy obiekt komórki, którego `value` to łańcuch `"=A1+A2"`, a `data_type` to `'f'` (formula). Nic więcej. Nie ma tu żadnego drzewa zależności ani ewaluatora — nie ma czego ewaluować, bo openpyxl nie ma takiego komponentu.

**Co trafi do pliku:** w `xl/worksheets/sheet1.xml` komórka `A3` dostanie element `<f>A1+A2</f>` (bez wiodącego znaku `=`) i **nie dostanie elementu `<v>`** z wartością. Excel po otwarciu zauważy brak wartości i — dzięki fladze `fullCalcOnLoad` — przeliczy formułę, a przy zapisie dopisze `<v>5</v>`.

**Konsekwencja praktyczna:** to jest dokładnie ten moment, w którym musisz zdecydować, czy budujesz raport dla człowieka (formuły są w porządku i nawet pożądane), czy potok danych dla maszyny (formuły są bezużyteczne — pisz wartości).

### Przykład 2 — skaner ryzyka: co jest w tym pliku?

Ten skrypt odpowiada na pytanie zadawane zanim użyjesz openpyxl na cudzym pliku: **co w nim jest i czy openpyxl to obsłuży?** Wykorzystujemy wiedzę z modułu 00 (plik `.xlsx` to ZIP) i wglądamy do środka. To uproszczona wersja tego, co w pełnej formie zrobisz w module 18.

```python
"""Skanuje plik .xlsx i ostrzega o elementach, które openpyxl gubi.

Uruchom:  python examples/01_skaner_ryzyka.py output/01_formula_dowod.xlsx
Jeśli nie podasz pliku, skaner utworzy plik demonstracyjny samodzielnie.

To narzędzie jest HEURYSTYKĄ: wykrywa znane części pakietu po nazwach
i szuka wzorców w XML. Nie jest wyrocznią - pełny inwentarz ryzyk
poznasz w module 18.
"""

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

# Części pakietu, których openpyxl nie modeluje -> znikną przy zapisie.
CZESCI_GUBIONE = {
    "vbaProject.bin": "Makra VBA (zachowasz TYLKO z keep_vba=True)",
    "slicer": "Slicery - poza modelem openpyxl",
    "timeline": "Osie czasu - poza modelem openpyxl",
    "pivotCache": "Pamięć podręczna tabeli przestawnej (odczyt tak, odświeżenie nie)",
    "activeX": "Formanty ActiveX - poza modelem openpyxl",
    "ctrlProps": "Właściwości formantów - poza modelem openpyxl",
}

# Części, które openpyxl obsługuje - informacyjnie.
CZESCI_ZNANE = {
    "xl/worksheets/": "zawartość arkuszy",
    "xl/styles.xml": "style i formaty",
    "xl/tables/": "tabele Excela",
    "xl/charts/": "wykresy (odbudowywane z modelu - sprawdź okiem)",
    "xl/media/": "obrazy",
    "xl/sharedStrings.xml": "współdzielone łańcuchy tekstowe",
    "xl/comments": "komentarze (w starszym formacie)",
    "xl/threadedComments": "komentarze w nowym formacie",
}


def skanuj(sciezka: Path) -> None:
    print(f"Skanuję: {sciezka}\n")

    # Test deterministyczny: plik .xlsx JEST archiwum ZIP. Plik .xls nie jest.
    if not zipfile.is_zipfile(sciezka):
        print("!! To nie jest archiwum ZIP.")
        print("   Najprawdopodobniej masz plik .xls (Excel 97-2003)")
        print("   albo .xlsb - openpyxl ich nie obsługuje.")
        print("   Rozwiązanie: przekonwertuj plik (np. LibreOffice) do .xlsx.")
        return

    with zipfile.ZipFile(sciezka) as archiwum:
        czesci = archiwum.namelist()

        print(f"Pakiet zawiera {len(czesci)} części.")
        znane = [c for c in czesci if any(k in c for k in CZESCI_ZNANE)]
        print(f"Z tego {len(znane)} rozpoznaję jako części obsługiwane przez openpyxl.\n")

        # --- Kontrola 1: części o znanych nazwach, których nie ma w modelu.
        ostrzezenia: list[str] = []
        for czesc in czesci:
            for fragment, opis in CZESCI_GUBIONE.items():
                if fragment in czesc:
                    ostrzezenia.append(f"{czesc}: {opis}")

        # --- Kontrola 2: bloki rozszerzeń i sparkline'y w XML arkuszy.
        for czesc in czesci:
            if not (czesc.startswith("xl/worksheets/sheet") and czesc.endswith(".xml")):
                continue
            xml = archiwum.read(czesc).decode("utf-8", errors="replace")

            if "sparkline" in xml:
                ostrzezenia.append(
                    f"{czesc}: SPARKLINE - openpyxl ich nie tworzy i gubi przy zapisie"
                )
            if "xdr:sp" in xml or "xdr:sp" in xml:
                pass  # kształty siedzą w xl/drawings, nie tutaj - patrz niżej
            if "extLst" in xml:
                ostrzezenia.append(
                    f"{czesc}: blok rozszerzeń (extLst) - część funkcji może zniknąć"
                )

        # --- Kontrola 3: kształty, pola tekstowe, wykresy w warstwie rysunkowej.
        for czesc in czesci:
            if czesc.startswith("xl/drawings/") and czesc.endswith(".xml"):
                xml = archiwum.read(czesc).decode("utf-8", errors="replace")
                if "xdr:sp" in xml:
                    ostrzezenia.append(
                        f"{czesc}: kształt / pole tekstowe - openpyxl tego nie modeluje"
                    )

    print("=" * 68)
    if ostrzezenia:
        print(f"ZNALAZŁEM {len(ostrzezenia)} POTENCJALNYCH PROBLEMÓW:")
        print("=" * 68)
        for o in sorted(set(ostrzezenia)):
            print("  -", o)
        print("\nWniosek: NIE zapisuj tego pliku przez openpyxl bez planu.")
        print("Szczegóły i bezpieczne wzorce: moduł 18.")
    else:
        print("Nie znalazłem elementów, które openpyxl gubi przy zapisie.")
        print("Pamiętaj jednak: to heurystyka, nie gwarancja.")


def utworz_plik_demo() -> Path:
    """Plik demonstracyjny, gdy użytkownik nie podał żadnego."""
    from openpyxl import Workbook

    OUTPUT.mkdir(parents=True, exist_ok=True)
    sciezka = OUTPUT / "01_skaner_demo.xlsx"
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Plik demonstracyjny"
    wb.save(sciezka)
    return sciezka


def glowny() -> None:
    if len(sys.argv) > 1:
        sciezka = Path(sys.argv[1])
        if not sciezka.exists():
            print(f"Nie ma takiego pliku: {sciezka}")
            return
    else:
        print("Nie podano pliku - tworzę demonstracyjny.\n")
        sciezka = utworz_plik_demo()

    skanuj(sciezka)


if __name__ == "__main__":
    glowny()
```

Zwróć uwagę na fragment z komentarzem „kształty siedzą w `xl/drawings`” — zostawiłem tam celowy ślad, bo pokazuje on sposób myślenia o tym problemie: elementy, które openpyxl gubi, nie siedzą w jednym miejscu, tylko w różnych częściach pakietu. Dlatego skaner sprawdza nazwy części **i** zawartość XML.

**Co się dzieje w pamięci:** czytamy części pliku jako bajty i dekodujemy je do tekstu. Dla dużych plików warto to robić tylko dla wybranych części (tak jak tutaj).

**Co trafi do pliku:** nic — skaner jest wyłącznie diagnostyczny. Wyjątkiem jest ścieżka „brak argumentu”, która tworzy `output/01_skaner_demo.xlsx`.

**Konsekwencja praktyczna:** ten skrypt jest Twoim pierwszym narzędziem w pracy z cudzymi plikami. Uruchom go na każdym pliku, zanim pozwolisz openpyxl go zapisać. Tekst „Wniosek: NIE zapisuj tego pliku przez openpyxl bez planu” to nie ozdoba — po module 18 będziesz dokładnie wiedział, co to znaczy.

### Przykład 3 — openpyxl i pandas: kto naprawdę pisze plik

Jeśli znasz pandas, prędzej czy później zadasz pytanie: po co mi openpyxl, skoro `df.to_excel()` działa? Ten przykład pokazuje, jak te narzędzia współpracują i gdzie jest granica.

**Uwaga:** ten przykład wymaga `pandas`, którego kurs nie instalował w module 00. Zainstaluj go osobno: `python -m pip install pandas`. Jeśli nie chcesz instalować pandas, pomiń ten przykład — nie jest potrzebny do dalszej nauki.

```python
"""Pokazuje, że pandas pod spodem korzysta z openpyxl.

Uruchom:  python examples/01_pandas_wspolpraca.py
Wymaga:   python -m pip install pandas
"""

from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"


def glowny() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    sciezka = OUTPUT / "01_pandas.xlsx"

    # --- KROK 1: zwykły eksport z pandas.
    # Silnikiem zapisu jest openpyxl (engine="openpyxl" to dla .xlsx wartość domyślna).
    df = pd.DataFrame(
        {
            "Region": ["Mazowieckie", "Śląskie", "Wielkopolskie"],
            "Sprzedaz": [120_000, 98_500, 76_200],
            "Marza_proc": [0.18, 0.21, 0.15],
        }
    )
    df.to_excel(sciezka, engine="openpyxl", index=False)
    print("pandas zapisał plik. Pod spodem pracował openpyxl.\n")

    # --- KROK 2: otwieramy TEN SAM plik openpyxl-em i czytamy formaty.
    # pandas tego nie robi - dla niego formatowanie nie istnieje.
    ws = load_workbook(sciezka).active
    print("Co widzi openpyxl w pliku zapisanym przez pandas:")
    for wiersz in ws.iter_rows(min_row=1, max_row=2, values_only=False):
        for komorka in wiersz:
            print(f"  {komorka.coordinate}: {komorka.value!r:>20}  "
                  f"[typ={komorka.data_type}, format={komorka.number_format!r}]")

    # Wniosek: pandas zapisał WARTOSCI, nie sformatował liczb.
    # Kolumna C trzyma ułamki (0.18), a nie procenty (18%).
    # Dopiero openpyxl (moduły 09-10) pozwala nadać format "0.0%".

    # --- KROK 3: alternatywa dla dużych zbiorów - strumień bez pandas.
    print("\nTen sam DataFrame wprost do arkusza, wiersz po wierszu:")
    bufor = BytesIO()
    from openpyxl import Workbook

    wb = Workbook()
    ws2 = wb.active
    for wiersz in dataframe_to_rows(df, index=False, header=True):
        ws2.append(wiersz)
    wb.save(bufor)  # plik powstaje w PAMIĘCI, nie na dysku
    print(f"  rozmiar wygenerowanego pliku w pamięci: {len(bufor.getvalue())} bajtów")


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** `df.to_excel` tworzy wewnętrznie skoroszyt openpyxl, wypełnia go i zapisuje — a potem wszystko znika. Nie masz dostępu do tego obiektu. Drugie podejście (`dataframe_to_rows` + `append`) daje Ci ten dostęp, bo skoroszyt tworzysz sam.

**Co trafi do pliku:** `output/01_pandas.xlsx` z trzema kolumnami i czterema wierszami (nagłówek + dane). **Bez formatowania** — liczby są liczbami, ale nie mają formatu waluty ani procentu.

**Konsekwencja praktyczna:** pandas i openpyxl to nie konkurenci. pandas jest dobry w „co jest w danych”, openpyxl w „jak to wygląda i jak jest zapisane”. Typowy podział pracy: pandas do przetworzenia, openpyxl do prezentacji. Ważne ostrzeżenie: **pandas odczytuje wartości z pamięci podręcznej formuł** — na pliku, którego nie widział Excel, `pd.read_excel` zwróci w miejscu formuły `NaN`. Dokładnie ten sam problem, co w przykładzie 1.

### Przykład 4 — mikrobenchmark: tryby openpyxl

Ten przykład nie porównuje bibliotek (do tego potrzebowałbyś `xlsxwriter`), ale pokazuje różnicę między dwoma trybami **wewnątrz** openpyxl. Jest to zapowiedź modułu 19 i praktyczny argument za rozdzieleniem „tworzę od zera” od „modyfikuję istniejące”.

```python
"""Porównuje tryb normalny i write_only przy zapisie 20 000 wierszy.

Uruchom:  python examples/01_benchmark_openpyxl.py
Wyniki zależą od komputera - traktuj je jako rząd wielkości, nie jak wyrok.
"""

import time
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

WIERSZY = 20_000
KOLUMN = 10


def dane():
    """Generator: nie tworzy listy 200 000 wartości w pamięci."""
    for r in range(WIERSZY):
        yield [f"w{r}-k{c}" if c == 0 else r * c for c in range(KOLUMN)]


def zapis_normalny(sciezka: Path) -> float:
    wb = Workbook()
    ws = wb.active
    start = time.perf_counter()
    for r, wiersz in enumerate(dane(), start=1):
        for c, wartosc in enumerate(wiersz, start=1):
            ws.cell(row=r, column=c, value=wartosc)
    wb.save(sciezka)
    return time.perf_counter() - start


def zapis_strumieniowy(sciezka: Path) -> float:
    # write_only: arkusz nie trzyma komórek w pamięci, wypuszcza je na bieżąco.
    # Uwaga: w tym trybie nie ma wb.active - arkusz trzeba utworzyć jawnie.
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Dane")
    start = time.perf_counter()
    for wiersz in dane():
        ws.append(wiersz)
    wb.save(sciezka)
    return time.perf_counter() - start


def glowny() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    a = OUTPUT / "01_bench_normalny.xlsx"
    b = OUTPUT / "01_bench_write_only.xlsx"

    czas_a = zapis_normalny(a)
    czas_b = zapis_strumieniowy(b)

    print(f"Wierszy: {WIERSZY}, kolumn: {KOLUMN}\n")
    print(f"tryb normalny    : {czas_a:6.2f} s   ({a.stat().st_size / 1024:.0f} kB)")
    print(f"tryb write_only  : {czas_b:6.2f} s   ({b.stat().st_size / 1024:.0f} kB)")
    print(f"\ntryb write_only był {czas_a / czas_b:.1f}x szybszy w tym teście.")


if __name__ == "__main__":
    glowny()
```

**Co się dzieje w pamięci:** w trybie normalnym openpyxl trzyma w pamięci **obiekt dla każdej komórki** — to 200 000 obiektów przy tym rozmiarze. W trybie `write_only` komórki są serializowane i wypuszczane na bieżąco, a w pamięci zostaje tylko bufor.

**Co trafi do pliku:** dwa pliki z tymi samymi danymi. Pliki mogą się nieznacznie różnić rozmiarem i kolejnością części — to normalne.

**Konsekwencja praktyczna:** jeśli generujesz duży plik od zera i nie musisz niczego poprawiać po fakcie, `write_only` jest właściwym trybem. Jeśli musisz wrócić do komórki `B7` i coś zmienić, `write_only` Ci tego nie pozwoli — trzeba było użyć trybu normalnego albo trybu normalnego w połączeniu z mniejszym plikiem. Wniosek dla wyboru biblioteki: dla bardzo dużych, „jednorazowych” eksportów `xlsxwriter` (który ma podobny model strumieniowy) bywa jeszcze szybszy, ale nie potrafi niczego odczytać.

### Przykład 5 — jak rozpoznać plik, którego openpyxl nie otworzy

```python
"""Wykrywa format pliku i wyjaśnia, co dalej.

Uruchom:  python examples/01_rozpoznaj_format.py
"""

import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output"

# Format: (nazwa pliku, opis sytuacji)
PRZYKLADY = [
    ("stary_plik.xls", "Excel 97-2003 - binarny, NIE jest ZIP-em"),
    ("nowy_plik.xlsx", "Excel 2010+ - to archiwum ZIP"),
    ("binarny.xlsb", "Binarny skoroszyt - obsługuje go pyxlsb (tylko odczyt)"),
]


def przygotuj_przyklady() -> None:
    """Tworzy pliki demonstracyjne, żebyś nie musiał ich szukać."""
    DATA.mkdir(parents=True, exist_ok=True)

    # .xls: piszemy zwykły tekst - openpyxl takiego pliku nie odczyta.
    (DATA / "stary_plik.xls").write_bytes(b"to nie jest prawdziwy xls, tylko demo")

    # .xlsx: prawdziwy, poprawny skoroszyt utworzony openpyxl-em.
    from openpyxl import Workbook

    wb = Workbook()
    wb.active["A1"] = "poprawny plik"
    wb.save(DATA / "nowy_plik.xlsx")

    # .xlsb: też zwykły tekst - demo "nieobsługiwanego binarnego formatu".
    (DATA / "binarny.xlsb").write_bytes(b"demo")


def zdiagnozuj(sciezka: Path, opis: str) -> None:
    print(f"\n{sciezka.name}  ({opis})")

    # Test 1: czy to ZIP? Determinista, nie zgadywanie.
    if not zipfile.is_zipfile(sciezka):
        print("   -> to NIE jest archiwum ZIP.")
        print("      openpyxl tego nie otworzy (dotyczy .xls, .xlsb).")
        return

    # Test 2: spróbuj wczytać. Nawet jeśli to ZIP, może nie być OOXML-em.
    # W zależności od wersji openpyxl zgłasza InvalidFileException
    # albo zipfile.BadZipFile - dlatego łapiemy oba.
    try:
        wb = load_workbook(sciezka, read_only=True)
    except (InvalidFileException, zipfile.BadZipFile) as blad:
        print(f"   -> to ZIP, ale nie skoroszyt OOXML: {type(blad).__name__}")
        return

    print(f"   -> OK. arkusze: {wb.sheetnames}")
    wb.close()  # W trybie read_only zamknięcie jest OBOWIĄZKOWE.


def glowny() -> None:
    przygotuj_przyklady()
    print("Diagnoza plików z katalogu data/:")
    for nazwa, opis in PRZYKLADY:
        zdiagnozuj(DATA / nazwa, opis)


if __name__ == "__main__":
    glowny()
```

Oczekiwany wynik:

```text
Diagnoza plików z katalogu data/:

stary_plik.xls  (Excel 97-2003 - binarny, NIE jest ZIP-em)
   -> to NIE jest archiwum ZIP.
      openpyxl tego nie otworzy (dotyczy .xls, .xlsb).

nowy_plik.xlsx  (Excel 2010+ - to archiwum ZIP)
   -> OK. arkusze: ['Sheet']

binarny.xlsb  (Binarny skoroszyt - obsługuje go pyxlsb (tylko odczyt))
   -> to NIE jest archiwum ZIP.
      openpyxl tego nie otworzy (dotyczy .xls, .xlsb).
```

**Co się dzieje w pamięci:** test `zipfile.is_zipfile` czyta wyłącznie nagłówek archiwum — jest natychmiastowy. Wczytanie `read_only` tworzy lekki obiekt skoroszytu bez komórek.

**Co trafi do pliku:** tworzymy trzy pliki demonstracyjne w `data/` — jedyny wyjątek od zasady „nie piszemy do `data/`”. Są to pliki **syntetyczne, niczyje**, utworzone w tym samym skrypcie; w prawdziwej pracy nie pisałbyś do `data/`. Zauważ, że `.xls` i `.xlsb` to zwykły tekst — celowo, żeby przykład działał bez pobierania prawdziwych plików.

**Konsekwencja praktyczna:** zamiast łapać wyjątki „na ślepo”, najpierw zadaj pytanie „czy to ZIP?”. To jedna linia kodu, która natychmiast rozdziela przypadki i daje użytkownikowi zrozumiały komunikat.

## 4. Anatomia API

Tabela zbiera wszystko, co pojawiło się w tym module. To jeszcze nie jest kurs openpyxl — to kurs **diagnostyki i wyboru narzędzia**.

| Metoda / klasa | Co robi | Parametry | Uwagi |
|----------------|---------|-----------|-------|
| `openpyxl.__version__` | Wersja biblioteki | — | Oczekujemy `3.1.5` |
| `Workbook()` | Nowy, pusty skoroszyt w pamięci | `write_only=False`, `iso_dates=False` | Nic nie zapisuje na dysk |
| `Workbook(write_only=True)` | Skoroszyt w trybie strumieniowym | — | Brak `wb.active`; arkusze tworzysz przez `create_sheet()` |
| `wb.save(plik_lub_bufor)` | Serializuje model do `.xlsx` | ścieżka, `str`, `Path`, obiekt plikopodobny (`BytesIO`) | **Nadpisuje istniejący plik bez ostrzeżenia** |
| `load_workbook(sciezka)` | Wczytuje plik do modelu | `read_only`, `keep_vba`, `data_only`, `keep_links`, `rich_text` | Pełny opis flag w module 18; `data_only` w module 06 |
| `wb.sheetnames` | Lista nazw arkuszy | — | Odpowiednik usuniętego `get_sheet_names()` |
| `wb["Nazwa"]` | Arkusz po nazwie | nazwa arkusza | Odpowiednik usuniętego `get_sheet_by_name()` |
| `wb.active` | Aktywny arkusz | — | Niedostępne w `write_only` |
| `wb.close()` | Zamyka skoroszyt i zwalnia zasoby | — | **Obowiązkowe** przy `read_only` |
| `ws["A3"]` | Komórka po adresie A1 | adres | Zapis: `ws["A3"] = wartość` |
| `cell.value` | Wartość komórki | — | Formuła zwracana jako tekst `"=A1+A2"` |
| `cell.data_type` | Typ zawartości komórki | — | `'n'` liczba, `'s'` tekst, `'f'` formuła, `'d'` data, `'b'` bool, `'e'` błąd |
| `openpyxl.utils.dataframe.dataframe_to_rows` | `DataFrame` → wiersze dla arkusza | `df`, `index`, `header` | Import właśnie z tej ścieżki |
| `Workbook.calculation` | Właściwości obliczeń skoroszytu | obiekt `CalcProperties` | `fullCalcOnLoad` (domyślnie `True`) wymusza przeliczenie przy otwarciu w Excelu; sprawdź nazwę atrybutu w dokumentacji swojej wersji |
| `openpyxl.utils.exceptions.InvalidFileException` | Wyjątek „nieobsługiwany format” | — | Zgłaszany np. dla `.xls` |
| `zipfile.ZipFile(path).namelist()` | Lista części pakietu | ścieżka | Podstawa skanera ryzyka |
| `zipfile.is_zipfile(path)` | Czy plik jest archiwum ZIP | ścieżka | Szybka, deterministyczna diagnoza formatu |
| `wb.defined_names.add(...)` | Dodaje definiowaną nazwę (3.1+) | obiekt `DefinedName` | Zastępuje usunięte `add_named_range` |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1.** Uruchom przykład 1. Następnie otwórz `output/01_formula_dowod.xlsx` w Excelu (albo LibreOffice Calc), zapisz plik (`Ctrl+S`), zamknij go i uruchom skrypt jeszcze raz. Zapisz w notatce, co zwróciło `data_only=True` przed zapisem w Excelu, a co po. To doświadczenie zapamiętasz lepiej niż jakikolwiek opis.

**Zadanie 2.** Uruchom przykład 5 i przeczytaj komunikaty. Odpowiedz pisemnie na jedno pytanie: **dlaczego** `stary_plik.xls` nie jest archiwum ZIP? (Odpowiedź: format `.xls` jest binarnym formatem „BIFF” z lat 90., nie pakietem XML; dopiero Excel 2007 wprowadził `.xlsx`, który jest ZIP-em.)

**Zadanie 3.** Wypisz trzy rzeczy, które openpyxl **potrafi**, i trzy, których **nie potrafi** — bez zaglądania do modułu. Potem sprawdź się. Jeśli pomyliłeś się w kategorii „sparkline'y”, to dobrze rokuje: to najczęstsza pomyłka.

### 🟡 Warsztat

**Zadanie 4.** Uruchom przykład 2 (skaner ryzyka) na pliku `output/01_pandas.xlsx` z przykładu 3. Zapisz wynik. Następnie utwórz w Excelu (ręcznie!) prosty plik z jednym wykresem słupkowym i jednym obrazkiem, zapisz go w `data/`, zeskanuj i porównaj raporty. Zobacz różnicę w tym, jak wygląda pakiet „prosty, techniczny” i „zrobiony ręcznie przez człowieka”.

**Zadanie 5.** Uruchom mikrobenchmark z przykładu 4. Zapisz czasy. Zmień `WIERSZY` na `50_000` i uruchom ponownie. Czy przewaga `write_only` rośnie, maleje, czy zostaje taka sama? Sformułuj hipotezę, dlaczego tak jest. (Podpowiedź: myśl o tym, ile obiektów komórek trzyma w pamięci każdy tryb.)

**Zadanie 6.** Napisz funkcję `wybierz_biblioteke(opis_zadania: str) -> str`, która na podstawie słów kluczowych w opisie zwraca nazwę właściwej biblioteki — i **wypisuje uzasadnienie**. Zadanie jest celowo „naiwne”: chodzi o to, żebyś samodzielnie zapisał w kodzie reguły, które właśnie poznałeś. Reguły: „sparkline” → `xlsxwriter`; „modyfikuj istniejący” → `openpyxl`; „oblicz formuły” → `xlwings`; „analiza”/„statystyka” → `pandas`; „makro”/„uruchom Excel” → `xlwings`.

### 🔴 Wyzwanie

**Zadanie 7.** Dla każdego z trzech poniższych scenariuszy wybierz bibliotekę i uzasadnij wybór w dwóch zdaniach. Następnie napisz dla każdego krótki szkic kodu (5–15 linii), który pokazuje **pierwszą** operację wykonywaną w tym scenariuszu. Wynik zapisz w pliku `output/01_decyzje.md`.

**Scenariusz A — „Raport nocny”.** Firma ubezpieczeniowa generuje co noc raport sprzedaży dla 300 agentów. Skrypt działa na serwerze Linuksowym w kontenerze Dockera. Nie ma tam Excela i nigdy nie będzie. Raport: jeden arkusz z danymi, drugi z podsumowaniem i wykresem, trzeci z metodyką. Dostępny jest firmowy szablon z logo i ustalonymi kolorami w arkuszu nagłówka.

**Scenariusz B — „Dashboard dyrektorski”.** Istnieje plik zbudowany ręcznie przez analityka: 6 zakładek, w tym dashboard ze sparkline'ami w 80 wierszach, dwa slicery do filtrowania, pole tekstowe z komentarzem i makro odświeżające dane przez Power Query. Dyrektor chce, żeby dane liczbowe w zakładce „Dane” były aktualizowane co tydzień ze skryptu, a reszta pliku — łącznie z dashboardem — pozostała nietknięta.

**Scenariusz C — „Audyt compliance”.** Zespół prawny musi sprawdzić 12 000 plików `.xlsx` z lat 2019–2025. W każdym trzeba znaleźć arkusze, których nazwa zawiera „kalkulacja”, i odczytać trzy konkretne komórki. Pliki są w archiwum na dysku sieciowym, łącznie 40 GB. Część z nich ma rozszerzenie `.xlsx`, ale w rzeczywistości jest starymi plikami `.xls` przemianowanymi siłą.

<details>
<summary><strong>Szkic rozwiązania zadania 7</strong></summary>

**Scenariusz A — openpyxl.**

Uzasadnienie: skrypt działa na serwerze bez Excela (odpada `xlwings`), a musi **wypełnić istniejący szablon** z logo i kolorami, zachowując jego formatowanie — a to potrafi tylko openpyxl (odpada też `xlsxwriter`, który jest write-only i nie otworzy szablonu). Wykres i drugi arkusz openpyxl też obsłuży.

Szkic pierwszej operacji:

```python
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
szablon = ROOT / "templates" / "raport_sprzedazy.xlsx"
wyjscie = ROOT / "output" / "raport_2026-10-07.xlsx"   # NOWA nazwa, nigdy nadpisanie

wb = load_workbook(szablon)                 # otwieramy SZABLON, nie tworzymy od zera
ws = wb["Dane"]
ws["B5"] = 120_000                          # pierwsza komórka z danymi
wb.save(wyjscie)                            # zapisujemy do NOWEGO pliku
```

**Scenariusz B — openpyxl jako dostawca danych, dashboard zostaje w Excelu.**

Uzasadnienie: openpyxl **nie może** bezpiecznie zapisać tego pliku — gubi sparkline'y, slicery, pole tekstowe i może uszkodzić makro (makra przeżyją z `keep_vba=True`, ale i tak nie ma gwarancji dla reszty). Właściwym rozwiązaniem jest rozdzielenie: skrypt zapisuje dane do **osobnego pliku** (albo do arkusza danych w pliku, który sam tworzy), a dashboard czyta z niego swoje liczby. Jeśli naprawdę trzeba zapisać ten sam plik — jedyną bezpieczną drogą jest `xlwings` (Excel zapisuje sam) na maszynie z Excelem.

Szkic pierwszej operacji:

```python
from pathlib import Path
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
wyjscie = ROOT / "output" / "dashboard_dane.xlsx"   # plik DANYCH, nie dashboard

wb = Workbook()                  # tworzymy od zera - nie dotykamy pliku dashboardu
ws = wb.active
ws.title = "Dane"
ws.append(["Region", "Sprzedaz", "Marza_proc"])       # wiersz nagłówka
ws.append(["Mazowieckie", 120_000, 0.18])             # pierwszy wiersz danych
wb.save(wyjscie)                 # dashboard czytany z tego pliku pozostaje nietknięty
```

**Scenariusz C — openpyxl w trybie `read_only`.**

Uzasadnienie: to zadanie wyłącznie **odczytu** z tysięcy plików, więc potrzebny jest tryb strumieniowy (`read_only=True`) i obowiązkowe `wb.close()` — bez tego skrypt wycieknie pamięć przy 12 000 plików. Dodatkowo trzeba odfiltrować pliki, które nie są ZIP-em (przemianowane `.xls`), zanim openpyxl zgłosi wyjątek.

Szkic pierwszej operacji:

```python
import zipfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
archiwum = ROOT / "data"

znalezione = []
for plik in archiwum.rglob("*.xlsx"):
    if not zipfile.is_zipfile(plik):        # przemianowany .xls? pomijamy świadomie
        znalezione.append((plik, "NIE ZIP - prawdopodobnie .xls"))
        continue
    wb = load_workbook(plik, read_only=True, data_only=True)
    try:
        for nazwa in wb.sheetnames:
            if "kalkulacja" in nazwa.lower():
                ws = wb[nazwa]
                # W read_only nie ma dostępu po adresie "A1" - czytamy wierszami.
                for wiersz in ws.iter_rows(min_row=1, max_row=3, values_only=True):
                    znalezione.append((plik, nazwa, wiersz))
    finally:
        wb.close()                          # OBOWIĄZKOWO, inaczej wyciek pamięci
```

</details>

## 6. Typowe błędy i pułapki

**1. „Wpisałem formułę i odczytuję `None`” (objaw) → openpyxl nie ma silnika obliczeń, a w pliku nie ma zapisanej wartości z pamięci podręcznej (przyczyna) → albo zostaw formuły odbiorcy, albo policz wartość w Pythonie i wpisz liczbę; jeśli potrzebujesz wartości z istniejącego pliku, otwórz i zapisz go najpierw w Excelu/LibreOffice (naprawa).**

Najczęstszy błąd w całym ekosystemie openpyxl. Wartość dodana wiedzy: nie próbuj tego „obejść” kolejnymi flagami. Nie ma flagi, która to włączy — nie istnieje komponent, który by to policzył.

**2. „Wczytałem plik z `data_only=True`, coś zmieniłem i zapisałem — wszystkie formuły zamieniły się na wartości” (objaw) → w trybie `data_only` formuły nie są wczytywane do modelu, więc przy zapisie nie ma czego zapisać (przyczyna) → nigdy nie zapisuj skoroszytu wczytanego z `data_only=True`; wczytaj drugi raz bez tej flagi, jeśli chcesz modyfikować (naprawa).**

To najkosztowniejszy błąd w kursie, bo niszczy dane. Analogia: `data_only=True` to „wynotuj mi tylko wyniki, nie przepisuj samych wzorów”. Jeśli potem odtworzysz z tych notatek dokument, wzorów w nim nie będzie. Szczegóły w module 06.

**3. „`load_workbook("dane.xls")` rzuca wyjątek, choć Excel otwiera ten plik bez problemu” (objaw) → openpyxl nie obsługuje formatu `.xls` (przyczyna) → przekonwertuj plik do `.xlsx` (LibreOffice headless, Excel, `libreoffice --headless --convert-to xlsx`); do odczytu `.xls` w Pythonie służy `xlrd` (naprawa).**

Excel otwiera `.xls`, bo to jego własny, historyczny format. openpyxl implementuje **inny** format. To nie wada openpyxl, to granica zakresu.

**4. „Wczytałem `.xlsm`, zapisałem i makra zniknęły” (objaw) → brak flagi `keep_vba=True` przy wczytaniu (przyczyna) → `load_workbook(path, keep_vba=True)` i zapis pod tą samą rodziną rozszerzeń (naprawa).**

Uwaga: nawet z `keep_vba=True` openpyxl **nie zmodyfikuje** kodu makr — potrafi je tylko zachować. Jeśli makro ma się wykonać, potrzebujesz `xlwings`/COM.

**5. „Chciałem dopisać wiersz do istniejącego pliku, wybrałem `xlsxwriter` i okazało się, że nie ma metody otwarcia pliku” (objaw) → `xlsxwriter` jest biblioteką write-only z założenia (przyczyna) → do modyfikacji istniejących plików używaj openpyxl; `xlsxwriter` traktuj jako narzędzie do generowania nowych plików (naprawa).**

Ta „wada” jest w rzeczywistości źródłem jego zalet: brak modelu odczytu pozwala mu strumieniować milion wierszy w stałej pamięci i oferować najbogatsze API formatowania.

**6. „Chcę sparkline'y w nowym pliku — openpyxl nie ma takiej metody” (objaw) → sparkline'y żyją w bloku rozszerzeń `x14`, którego openpyxl nie modeluje (przyczyna) → wygeneruj plik `xlsxwriter`-em; nie przepuszczaj go potem przez openpyxl (naprawa).**

I uwaga na dezinformację w internecie: zdarzają się blogi twierdzące, że openpyxl obsługuje sparkline'y. Nie obsługuje — nie ma w nim klasy ani metody `add_sparkline`. Jeśli gdzieś to widzisz, materiał dotyczy `xlsxwriter` albo jest po prostu błędny.

**7. „Mój skrypt z 2018 roku przestał działać po aktualizacji openpyxl” (objaw) → usunięto przestarzałe API w wersji 3.1.0 (przyczyna) → zamień wywołania: `wb.get_sheet_names()` → `wb.sheetnames`, `wb.get_sheet_by_name("X")` → `wb["X"]`, `wb.add_named_range(...)` → `wb.defined_names.add(DefinedName(...))`, `ws.orientation` → `ws.page_setup.orientation`, `ws.page_size` → `ws.page_setup.paperSize` (naprawa).**

Jeśli utkniesz, zajrzyj do changelogu konkretnej wersji — tam zawsze jest sekcja „Removals”. To szybsze niż szukanie po forach.

**8. „Na serwerze nie mam Excela i `xlwings` nie działa” (objaw) → `xlwings` wymaga zainstalowanego programu Excel (Windows/macOS), bo to on wykonuje pracę (przyczyna) → jeśli naprawdę potrzebujesz przeliczania formuł na serwerze, zainstaluj LibreOffice headless albo przenieś logikę do Pythona (naprawa).**

To nie jest „problem z biblioteką” — to konsekwencja tego, co ta biblioteka robi. Zawsze zadaj sobie pytanie: „kto w moim rozwiązaniu ma być tym samym procesem, który liczy?”.

**9. „pandas wczytał mi tylko jedną zakładkę i w miejscu formuł mam `NaN`” (objaw) → `pd.read_excel` domyślnie czyta pierwszy arkusz i tylko wartości zapisane w pliku (przyczyna) → użyj `sheet_name=None` (wszystkie arkusze, słownik), a w miejscu formuł — najpierw przelicz plik w Excelu/LibreOffice (naprawa).**

Pamiętaj, że pandas i openpyxl mają **ten sam** problem z formułami, bo mają to samo źródło danych: plik, nie silnik obliczeń.

**10. „Zapisałem plik z rozszerzeniem `.xlsm`, ale bez VBA — Excel mówi, że plik jest uszkodzony” (objaw) → niezgodność zawartości i rozszerzenia (przyczyna) → używaj `.xlsx` dla plików bez makr, `.xlsm` tylko gdy plik faktycznie zawiera projekt VBA (naprawa).**

Analogia: wsadziłeś dokument bez pieczątki do koperty z napisem „dokument urzędowy”. Koperta obiecuje coś, czego w środku nie ma; odbiorca (Excel) zgłasza to jako błąd.

**11. „Przepuściłem firmowy dashboard przez openpyxl i zniknęły sparkline'y, slicery i pole tekstowe” (objaw) → to elementy spoza modelu openpyxl; przy zapisie, który jest przepisaniem, nie mają z czego powstać (przyczyna) → stosuj wzorzec „rozdzielenie danych od projektu”: skrypt dostarcza dane do osobnego pliku/arkusza, a plik projektowany pozostaje nietknięty (naprawa).**

W module 18 nauczysz się dodatkowo **wykrywać** takie pliki przed zapisem, żeby nigdy nie zadziałać na nich przez przypadek.

**12. „Skrypt działa lokalnie, ale na serwerze produkcyjnym pliki wychodzą uszkodzone przy zapisie, gdy użytkownik ma je otwarte w Excelu” (objaw) → brak koordynacji dostępu do pliku (przyczyna) → zapisuj do pliku tymczasowego i podmieniaj atomowo (`os.replace`), nie nadpisuj pliku, który ktoś trzyma otwarty; ustaw też politykę nazw z datą/godziną (naprawa).**

To pułapka produkcyjna, którą rozbierzemy na części w module 18. Zapamiętaj samą zasadę: zapisywanie „w miejsce” pliku, który może być otwarty, to hazard.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **openpyxl pracuje z plikiem, nie z programem.** Implementuje format `.xlsx`/`.xlsm` (OOXML), a nie program Excel. Stąd bierze się jego lekkość, wszechobecność — i wszystkie jego braki.

2. **Nie ma silnika obliczeń i nigdy nie będzie.** openpyxl potrafi zapisać formułę jako tekst i odczytać **wynik zapisany wcześniej przez kogoś innego** (`data_only=True`). Nie policzy niczego sam. Jeśli Twoja logika potrzebuje liczby — policz ją w Pythonie.

3. **Każdą funkcję oceniaj w trzech kategoriach: potrafi / potrafi częściowo / gubi przy zapisie.** To nie są stopnie jakości openpyxl, tylko mapa jego możliwości. „Potrafi częściowo” zawsze oznacza: sprawdź efekt okiem. „Gubi” oznacza: nie przepuszczaj przez niego takich plików.

4. **Dobór narzędzia zależy od zadania, nie od sympatii.** Czytasz i modyfikujesz istniejący plik → openpyxl. Tworzysz nowy, bogaty raport od zera, potrzebujesz sparkline'ów → `xlsxwriter`. Analizujesz dane → `pandas`. Musisz przeliczyć formuły albo uruchomić makro w prawdziwym Excelu → `xlwings`. Zmieniasz formaty plików → `pyexcel` albo LibreOffice headless.

5. **Wersja to część kontraktu.** Kurs opiera się na **openpyxl 3.1.5** (gałąź 3.1.x, wymaga Pythona ≥ 3.8). W 3.1.0 zniknęło kilka starych metod (`get_named_range`, `add_named_range`, `remove_named_range`, `ws.orientation`, `ws.page_size`, `ws.page_breaks` oraz dawne akcesory arkuszy). Zanim zaczniesz debugować kod z tutoriala, sprawdź, dla której wersji ten tutorial powstał.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. WYBÓR NARZĘDZIA - drzewko decyzyjne w komentarzu
# ==================================================================
#
# Czy plik już istnieje i mam go ZMODYFIKOWAĆ?
#   TAK  -> openpyxl (jedyna z popularnych, która to potrafi)
#   NIE  -> dalej
#
# Czy potrzebuję wartości OBLICZONYCH formuł?
#   TAK  -> xlwings (prawdziwy Excel) / LibreOffice headless / policz w Pythonie
#   NIE  -> dalej
#
# Czy potrzebuję sparkline'ów albo najbogatszego formatowania?
#   TAK  -> xlsxwriter (nowy plik od zera; nie odczytuje niczego)
#   NIE  -> dalej
#
# Czy to głównie ANALIZA danych?
#   TAK  -> pandas (+ polars/duckdb przy dużych zbiorach)
#   NIE  -> openpyxl
#
# Czy plik ma makra (.xlsm) i muszą przetrwać?
#   TAK  -> load_workbook(path, keep_vba=True) ; nadal nie zmienisz kodu makr
#   NIE  -> openpyxl bez flag
#
# Czy plik ma sparkline'y / slicery / kształty / Power Query?
#   TAK  -> NIE zapisuj go openpyxl-em. Rozdziel dane od projektu.

# ==================================================================
# 2. FORMULARZE - openpyxl zapisuje, nie liczy
# ==================================================================
from pathlib import Path
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
sciezka = ROOT / "output" / "demo.xlsx"
ROOT.joinpath("output").mkdir(parents=True, exist_ok=True)

wb = Workbook()
ws = wb.active
ws["A1"] = 2
ws["A2"] = 3
ws["A3"] = "=A1+A2"        # to TEKST formuły; openpyxl jej nie policzy
wb.save(sciezka)

print(load_workbook(sciezka)["Sheet"]["A3"].value)              # '=A1+A2'
print(load_workbook(sciezka, data_only=True)["Sheet"]["A3"].value)  # None
# 5 pojawi się dopiero wtedy, gdy plik zostanie otwarty i zapisany w Excelu.
# UWAGA: NIGDY nie zapisuj skoroszytu wczytanego z data_only=True - zniszczysz formuły.

# ==================================================================
# 3. DIAGNOSTYKA FORMATU I RYZYKA
# ==================================================================
import zipfile
from openpyxl.utils.exceptions import InvalidFileException

def czy_obslugiwany(sciezka: Path) -> bool:
    """Szybki, deterministyczny test: .xlsx JEST archiwum ZIP, .xls nie jest."""
    return zipfile.is_zipfile(sciezka)

def załaduj_bezpiecznie(sciezka: Path):
    """Wczytuje plik, tłumacząc typowe problemy na zrozumiały komunikat."""
    if not czy_obslugiwany(sciezka):
        raise ValueError(f"{sciezka.name}: to nie jest pakiet OOXML (.xls? .xlsb?)")
    try:
        return load_workbook(sciezka, read_only=True)
    except (InvalidFileException, zipfile.BadZipFile) as blad:
        raise ValueError(f"{sciezka.name}: ZIP, ale nie skoroszyt ({type(blad).__name__})")

with zipfile.ZipFile(sciezka) as archiwum:
    czesci = archiwum.namelist()          # lista części pakietu
    # Części, których openpyxl NIE modeluje (znikną przy zapisie):
    #   xl/vbaProject.bin, *slicer*, *timeline*, *pivotCache*, *activeX*
    # Szukaj też w XML arkuszy: "sparkline" oraz "extLst".

# ==================================================================
# 4. TRZY KATEGORIE - skrócona tabela
# ==================================================================
# POTRAFI:            wartości, typy, daty, style, formaty liczb, scalanie,
#                     wydruk, tabele, formatowanie warunkowe, komentarze,
#                     hiperlinki, obrazy, walidacja, ochrona, nazwy zdefiniowane
# POTRAFI CZĘŚCIOWO:  wykresy (odbudowywane z modelu), tabele przestawne
#                     (odczyt/zachowanie, bez tworzenia), makra VBA (keep_vba)
# GUBI PRZY ZAPISIE:  sparkline'y, kształty i pola tekstowe, formanty,
#                     slicery, osie czasu, Power Pivot / Power Query

# ==================================================================
# 5. WERSJA
# ==================================================================
import openpyxl
print(openpyxl.__version__)   # oczekiwane: 3.1.5  (gałąź 3.1.x, Python >= 3.8)
# Usunięte w 3.1.0: wb.get_named_range / add_named_range / remove_named_range,
#                   ws.formula_attributes / page_breaks / show_summary_below /
#                   show_summary_right / page_size / orientation,
#                   image.get_emu_dimensions
# Zamiast tego: wb.defined_names (DefinedName), ws.page_setup.orientation,
#               ws.page_setup.paperSize, ws.row_breaks / ws.col_breaks
```

## 9. Co dalej

Masz teraz narzędzie pojęciowe, które będzie Ci służyło do końca kursu: trójpodział na **potrafi / potrafi częściowo / gubi przy zapisie**. Każdy kolejny moduł będzie do niego wracał — czasem jednym zdaniem, czasem całą sekcją.

W **module 02** zajrzymy do środka pliku `.xlsx` na poważnie. Rozbierzemy pakiet OOXML na części: `xl/workbook.xml` jako spis treści, `xl/worksheets/sheet1.xml` jako kartkę, `xl/sharedStrings.xml` jako skoroszyt wspólnych etykiet, `xl/styles.xml` jako legendę wyglądu. Zrozumiesz, dlaczego komórka nie „ma” formatowania, tylko numer stylu — i dlaczego to jedno zdanie wyjaśnia połowę zachowań openpyxl. Dowiesz się też, gdzie dokładnie w pakiecie siedzą te elementy, które openpyxl gubi: sparkline'y, kształty i slicery. Bez tego modułu moduł 18 byłby dla Ciebie zbiorem niepowiązanych ostrzeżeń.

Przygotuj do modułu 02 dwa pliki: jeden utworzony przez nasz skrypt (np. `output/01_formula_dowod.xlsx`) i **jeden zrobiony ręcznie w Excelu z wykresem i obrazkiem**. Będziemy je porównywać część po części.