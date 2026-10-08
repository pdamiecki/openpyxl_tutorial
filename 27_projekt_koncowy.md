# Moduł 27 — Projekt końcowy: generator raportów sprzedażowych

> **Część:** V — Projekt i domknięcie · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–26

## 0. W tym module nauczysz się

- **Zamienisz całą wiedzę z 26 modułów w jeden działający produkt** — aplikację CLI `python -m raporty --config config.yaml`, która czyta CSV + konfigurację YAML i zapisuje wieloarkuszowy raport `.xlsx` z tabelą, wykresem dwuosiowym, paskiem danych, ochroną arkusza i logiem audytowym.
- **Naprawdę zrozumiesz specyfikację jako dokument techniczny** — jak z listy życzeń „raport ma ładnie wyglądać" zrobić dziesięć ponumerowanych wymagań funkcjonalnych (`F1`–`F10`) i zestaw wymagań niefunkcjonalnych z **liczbami**, które da się zmierzyć.
- **Przejdziesz dziesięć kamieni milowych** (M1–M10) — od pliku, który zapisuje jedną komórkę, do aplikacji z transakcyjnym zapisem i 200 000 wiersze. **Po każdym kroku masz działający kod i przechodzący test**, więc nigdy nie znajdziesz się w sytuacji „napisałem 900 linii i nic nie działa".
- **Zobaczysz dwie architektury tego samego zadania** — „raport minimalistyczny" (poziom 1 abstrakcji: jeden plik, 120 linii) i „raport produkcyjny" (poziom 3: warstwy, porty, transakcje). Zmierzysz koszt i korzyść obu i zrozumiesz, że **żadna nie jest „tą właściwą"** — właściwa zależy od tego, ile razy ten kod ma się jeszcze wykonać.
- **Poznasz dwie decyzje projektowe, których nie da się obejść w prawdziwych raportach** — (1) dlaczego przy dużych plikach raport trzeba **rozdzielić na dwa pliki** (strumieniowe dane + bogate podsumowanie), a nie próbować zmieścić wszystkiego w jednym; (2) dlaczego reguły biznesowe muszą być **zamkniętą paletą w kodzie**, a nie wyrażeniami w YAML-u.
- **Otrzymasz rubrykę oceny, warianty rozszerzenia i mapę „problem → moduł kursu"** — żebyś po ukończeniu projektu wiedział, do którego modułu wrócić, gdy pojawi się konkretna potrzeba.
- **Zamkniesz kurs listą kontrolną przed wdrożeniem** — czternastoma pytaniami, na które trzeba sobie odpowiedzieć **przed** wypuszczeniem generatora raportów do ludzi.

## 1. Intuicja i analogia

### 1.1. Dom, nie cegły

Przez dwadzieścia siedem modułów dostawałeś **cegły, cement, kielnię i poziomnicę**. Każdy moduł był o innym narzędziu: moduł 05 o typach, moduł 09 o stylach, moduł 13 o warunkach, moduł 19 o wydajności, moduł 26 o architekturze. Uczciwie powiedzmy: **to była wiedza w kawałkach**. Umiesz położyć cegłę, ale nie wiesz jeszcze, czy stawiasz ścianę nośną czy ozdobną.

Ten moduł jest o **domu**. O tym, żeby stanąć z boku, spojrzeć na cały projekt i zdecydować: ile pokoi, gdzie okna, jak grube ściany. Cegły są już znajome. Teraz chodzi o **decyzje, których nie da się wycofać później** — a właśnie te decyzje odróżniają architekta od murarza.

Analogie, których użyję dalej, są trzy i warto je zapamiętać:

**Przepis vs wykonanie.** Specyfikacja to przepis kucharski. Kod to gotowanie. Jeżeli przepis mówi „dopraw do smaku", to nie jest przepis — to wskazówka. W prawdziwym przepisie jest „½ łyżeczki soli", bo **musi** być powtarzalne. Twoje wymagania mają być jak ten drugi przepis, a nie jak „dopraw do smaku".

**Budowa domu etapami.** Nie zbudujesz domu, zaczynając od dachu. Najpierw fundamenty, potem ściany, potem dach, na końcu okna. **Po każdym etapie dom stoi** — nie jest skończony, ale stoi. To jest cała idea dziesięciu kamieni milowych: po każdym kroku **masz działającą aplikację**, tylko mniej w niej jest.

**Dwa samochody na jedną trasę.** Jeżeli musisz przewieźć 200 ton żwiru i jednocześnie mieć w kabinie klimatyzację, skórzaną tapicerkę i system audio — nie próbuj projektować jednej ciężarówki, która to wszystko ma. Węższą, szybką i tanio — do żwiru. Ozdobną, wolną, z wyposażeniem — jako samochód towarzyszący. To jest dokładnie wybór, przed którym stanie ten projekt na etapie M10, i zobaczymy go na żywym przykładzie.

### 1.2. Specyfikacja to nie biurokracja — to kontrakt

Wyobraź sobie, że zamawiasz stół. Mówisz stolarzowi: „stół, drewniany, ładny, na 8 osób, na taras". To **nie jest** specyfikacja. To wyobrażenie. Za tydzień stolarz przynosi stół z sosny, 90 × 90 cm, na 4 osoby, na wysokości 60 cm. I mówi: „ładny, drewniany". Ma rację! A Ty jesteś wściekły.

Prawdziwa specyfikacja mówi: **dąb, blat 200 × 100 cm, wysokość 78 cm, nogi 8 × 8 cm, wykończenie olejem twardym, odporny na wodę**. Teraz każdy punkt **da się zweryfikować** — a jeżeli nie da się, to nie jest specyfikacja, tylko marketing.

W oprogramowaniu różnica jest jeszcze ostrzejsza. Specyfikacja funkcjonalna ma **ponumerowane wymagania** (`F1`, `F2`, ...). Numer nie jest po to, żeby ładnie wyglądał. Jest po to, żeby w code review napisać: „ta funkcja realizuje F6, a nie realizuje F4". Bez numerów nie ma testów, nie ma odbioru i nie ma dyskusji o zmianach — jest tylko „no chyba działa".

I trzecia rzecz: specyfikacja **nie mówi jak**. „Tabela z filtrem" (`F3`) nie mówi „użyj `openpyxl.worksheet.table.Table`". To jest decyzja wykonawcy — a jak wiesz z modułu 26, **architektura to właśnie miejsce, w którym podejmuje się takie decyzje i nazywa je po imieniu**.

### 1.3. Dwie ścieżki i ich prawdziwy koszt

Zapytanie, które pada zawsze przy projektach: „Po co te warstwy? To 200 wierszy kodu na raport."

Uczciwa odpowiedź brzmi: **zależy, ile razy ten kod ma się wykonać**.

Popatrz na dwa realne scenariusze:

**Scenariusz A — raport robimy raz.**
Kasia z księgowości pisze do Ciebie: „Czy możesz wygenerować zestawienie z tego arkusza? Potrzebuję na jutro, jednorazowo." Tu poziom 3 abstrakcji jest **szkodliwy**: zamiast 30 minut, spędzasz 6 godzin na projektowaniu portów i pisaniu testów. Oddajesz to samo, tylko później. To jest nadmiarowa abstrakcja — i **jest zła**.

**Scenariusz B — ten raport robimy co miesiąc, dla 40 regionów, i zmieniał się 7 razy w ciągu roku.**
Po pierwszej zmianie kolumny trzeba wejść w kod. Po trzeciej ktoś się boi. Po piątej „właściwie to robimy to ręcznie w Excelu, bo automat się psuje". Po siódmej nikt już nie pamięta, jak to działa. **To jest dokładnie ten moment**, w którym brak architektury kosztuje sto razy więcej niż jej zbudowanie.

Ale jest jeszcze scenariusz C, o którym się zapomina.

**Scenariusz C — raport robimy co miesiąc, ale przez pół roku nie zmienił się ani razu.**
Tu również nie potrzebujesz pełnych warstw. Potrzebujesz **funkcji, konfiguracji i testu regresji**. Poziom 1.5, nie 3.

To jest cała różnica: **architektura ma zarabiać na siebie**. Płacisz za nią teraz (czas projektowania), a odbierasz w przyszłości (każda zmiana bez strachu). Jeżeli zmian jest zero, płacisz i nie odbierasz. Nasz projekt jest **duży i produkcyjny** — dlatego idziemy w poziom 3. Ale w ćwiczeniach zbudujesz jeszcze wersję minimalistyczną, żeby **poczuć różnicę w rękach**, a nie z opisu.

### 1.4. Żwir i tapicerka (przeczytaj dwa razy)

Najpierw **fakty**, bo tu ludzie się potykają:

- openpyxl w trybie **normalnym** trzyma **cały skoroszyt w pamięci**. 200 000 wierszy × 8 kolumn to 1,6 miliona obiektów `Cell`, każdy ze wskaźnikiem do stylu. Realnie: **kilka gigabajtów pamięci** i **minuty**, nie sekundy.
- openpyxl w trybie **`write_only`** strumieniuje wiersze — pamięć jest stała (rzędu megabajtów), czas liniowy. Ale **nie da się** cofnąć, nie da się dopisać nic „z boku", a część funkcji (tabele, wykresy na tym samym arkuszu, dostęp losowy) jest niedostępna albo ograniczona.
- **Wykresy, tabele, reguły warunkowe** są tanie wtedy, gdy są przypięte do **zakresu**, a nie do komórek. Reguła warunkowa na `F2:F200001` to **jeden element XML**, nie 200 000.

I teraz wniosek, do którego prowadzę od początku: **Twoje 200 000 wierszy i Twój wykres dwuosiowy nie muszą być w tym samym pliku.** To nie jest ustępstwo ani „kombinowanie" — to jest **lepsza architektura**:

| Plik | Zawartość | Tryb | Rozmiar |
|---|---|---|---|
| `dane_2026-09.xlsx` | surowe wiersze, kolumny, formaty liczbowe | `write_only` + `lxml` | 200 000 wierszy |
| `raport_2026-09.xlsx` | agregaty per region, KPI, wykres dwuosiowy, tabela z filtrem, warunki, Metodyka, ochrona | normalny | kilka tysięcy komórek |

Dwa samochody na jeden transport. Szybki i wąski do żwiru, ozdobny do prezentacji. **To jest wymaganie architektoniczne, którego nie ma w specyfikacji na początku — i które sam musisz wykryć na etapie M10.** Dokładnie tak wygląda praca architekta: wymagania są ze sobą w konflikcie, a Ty znajdujesz rozwiązanie, które nie polega na wybraniu jednego i udawaniu, że drugie nie istnieje.

### 1.5. „Po sprawdzeniu na oko" — czyli dlaczego „działa" nie wystarcza

W module 18 trzy razy powtórzyłem: **weryfikacja na oko jest obowiązkowa**. Tutaj dodaję drugą część tego zdania: **weryfikacja na oko nie wystarcza**.

Raport może wyglądać wspaniale i mieć w sobie błąd, którego nikt nie zobaczy przez rok:

- kwota zapisana jako tekst `"1 234,56"` zamiast liczby `1234.56` — sumy w Excelu pokażą zero,
- data jako `2026-10-08` wpisana jako string — filtr po zakresie dat nie działa,
- `#REF!` w miejscu, którego nikt nie ogląda,
- pasek danych, który dla wartości ujemnych skaluje się „od zera" i wygląda bezsensownie,
- wykres dwuosiowy, który po otwarciu na innym komputerze ma osie odwrócone,
- sparkline'y, które zniknęły przy modyfikacji istniejącego pliku.

Dlatego cały projekt ma **trzy poziomy sprawdzania**, i to jest jego konstrukcja, nie dodatek:

1. **Testy** (moduł 20) — szybkie, deterministyczne, sprawdzają model i wynik przez `BytesIO`.
2. **Weryfikacja struktury pliku** (moduł 18) — po zapisie otwieramy plik, liczymy arkusze, wiersze i reguły (`podsumowanie_skoroszytu`).
3. **Oko człowieka** — otwierasz plik w Excelu i patrzysz. **Zawsze.** Bo test nie sprawdzi, czy wykres wygląda sensownie, ani czy pasek danych przy wartościach ujemnych nie przypomina plamy.

Trzy poziomy, trzy różne koszty i trzy różne klasy błędów. Pominięcie któregokolwiek z nich to nie „przyspieszenie" — to **odłożenie błędu** do momentu, gdy jego wyłapanie będzie droższe.

## 2. Teoria

### 2.1. Specyfikacja biznesowa

**Kontekst.** Dział sprzedaży potrzebuje co miesiąc zestawienia sprzedaży per region. Dane są eksportowane z systemu sprzedażowego jako CSV. Obecnie zestawienie tworzy jedna osoba ręcznie w Excelu — zajmuje jej to pół dnia, zdarza się pomyłka. Ma powstać generator.

**Uruchomienie:**

```bash
python -m raporty --config config/raport_sprzedazy.yaml
```

**Dane wejściowe:**

- `data/orders.csv` — eksport sprzedaży (kolumny: `data`, `region`, `klient`, `produkt`, `ilosc`, `cena_netto`, `koszt_netto`),
- `config/raport_sprzedazy.yaml` — konfiguracja raportu.

**Dane wyjściowe:** dwa pliki w `output/` (decyzja architektoniczna z sekcji 1.4).

### 2.2. Wymagania funkcjonalne (F1–F10)

Każde wymaganie ma **numer** i **sposób weryfikacji**. Bez tego drugiego nie jest wymaganiem.

| Nr | Wymaganie | Sposób weryfikacji |
|---|---|---|
| **F1** | Skoroszyt zawiera wiele arkuszy: `Podsumowanie`, `Regiony`, `Dynamika`, `Metodyka`, `_meta`, a w osobnym pliku `Dane` | `wb.sheetnames` zawiera dokładnie oczekiwaną listę; test asercji na krotkę nazw |
| **F2** | Nagłówek raportu zawiera logo firmowe (plik `assets/logo.png`), wyrównane do lewej-górnej komórki, o wysokości 44 px | Test sprawdza obecność `xl/media/image1.png` w ZIP; weryfikacja **na oko** rozmiaru |
| **F3** | Arkusz `Regiony` zawiera tabelę Excela z filtrem i stylem `TableStyleMedium9` | `ws.tables` zawiera tabelę o `displayName="TabelaRegiony"`; `ref` obejmuje nagłówki |
| **F4** | Kwoty mają format walutowy polski: tysiące oddzielone spacją, dwie cyfry po przecinku, minus w nawiasie | `ws["D2"].number_format` równy formatowi ze słownika; weryfikacja **na oko** |
| **F5** | Arkusz `Dynamika` wyróżnia warunkowo spadki (dynamika < 0) tłem `FCE4E4` i kolorem tekstu `9C0006` | Liczba reguł w `ws.conditional_formatting`; zakres `sqref` zgodny z danymi |
| **F6** | Wykres na arkuszu `Regiony`: słupki = sprzedaż netto, linia = marża % na **drugiej osi** | Obecność `xl/charts/chart1.xml`; **weryfikacja na oko** — dwie osie, dwie serie |
| **F7** | Pasek danych (`DataBarRule`) dla kolumny dynamiki, obsługujący wartości ujemne | Reguła z `start_value < 0`; **weryfikacja na oko** — paski ujemne muszą być widoczne |
| **F8** | Arkusz `Metodyka` opisuje słownie: skąd dane, jak liczona marża, co znaczą kolory, jaki jest schemat | Test sprawdza obecność ≥ 6 wierszy tekstu; **weryfikacja na oko** — czy opis jest zrozumiały |
| **F9** | Właściwości dokumentu zawierają tytuł, autora, wersję schematu i identyfikator generowania; arkusz `_meta` zawiera `generator`, `schemat`, `hash_wsadu`, `wygenerowano` | `wb.properties.title`, `description`; odczyt komórek `_meta` po ponownym wczytaniu |
| **F10** | Arkusze `Podsumowanie`, `Regiony`, `Dynamika` są **chronione hasłem**; komórki danych są zablokowane, ale pola filtrów dostępne | `ws.protection.sheet is True`; **weryfikacja na oko** — klikanie w komórkę daje komunikat |
| **F11** | Wszystkie teksty pochodzące z CSV przechodzą sanitację: tekst zaczynający się od `=`, `+`, `-`, `@`, tabulatora lub CR jest **odrzucany** z czytelnym błędem | Testy parametryzowane po pięciu znakach; test „poprawne dane przechodzą" |
| **F12** | Walidacja danych w kolumnie regionu: lista rozwijana z listy regionów z arkusza `Regiony` | `ws.data_validations` zawiera walidację typu `list` z `formula1` wskazującym zakres |

To jest dwanaście wymagań (rozszerzyłem z dziesięciu o `F11` i `F12`, bo sanitacja i walidacja to **odrębne** wymagania i muszą mieć osobne numery — inaczej „walidacja wejścia" znaczy jednocześnie wszystko i nic).

**Zauważ coś ważnego.** Przy `F1`, `F5`, `F9`, `F11`, `F12` da się napisać test z automatu. Przy `F2`, `F4`, `F6`, `F7`, `F8`, `F10` — **nie do końca**, i te wymagania mówią wprost: „weryfikacja na oko". To nie jest słabość specyfikacji; to jest **jej uczciwość**. Udawanie, że rozmiar logo da się przetestować automatycznie, prowadzi do testu, który sprawdza, że `img.width == 180`, i do logo rozciągniętego w pionie — bo test nie widzi proporcji.

### 2.3. Wymagania niefunkcjonalne

To są wymagania, których spełnienie **mierzy się**, a nie „sprawdza".

| Nr | Wymaganie | Pomiar |
|---|---|---|
| **N1** | 200 000 wierszy przetworzonych w czasie **< 60 s** | `time.perf_counter()` wokół pełnego przebiegu; test z `pytest.mark.slow` |
| **N2** | Szczytowe zużycie pamięci **< 500 MB** | `tracemalloc.get_traced_memory()` wokół przebiegu |
| **N3** | Zapis **atomowy** z backupem; brak pliku tymczasowego po błędzie | Test: zasymulowany wyjątek → plik docelowy niezmieniony, brak `.*.tmp.xlsx` |
| **N4** | **Idempotencja**: drugie uruchomienie nie zmienia zawartości pliku | Dwa przebiegi → `podsumowanie_skoroszytu()` identyczne |
| **N5** | Dane wejściowe **nigdy** nie są nadpisywane ani modyfikowane | Hash pliku CSV przed i po przebiegu identyczny |
| **N6** | **Deterministyczność**: dwa przebiegi z tym samym wsadem → ten sam `hash_wyniku` | `a.hash_wyniku == b.hash_wyniku` |
| **N7** | Pokrycie testami **> 80 %** logiki domenowej | `coverage` z filtrem `--include="src/raporty/domain/*"` |
| **N8** | Testy domeny i aplikacji działają **bez openpyxl i bez Excela** | CI w środowisku bez `openpyxl`; `pytest tests/domain tests/application` |
| **N9** | Każdy wygenerowany plik niesie wersję generatora i hash wsadu | Odczyt `_meta` z pliku po miesiącu — odpowiedź „skąd to" |
| **N10** | Utrata funkcji pliku (sparkline, slicery) jest **wykrywana i logowana** jako `WARNING` | Test na pliku z sparkline'ami; alert w logu i w `_meta` |

**Dlaczego `N1` to 60 sekund, a nie „szybko".** Bo „szybko" nie da się zmierzyć, więc nigdy nie wiadomo, czy jesteśmy w budżecie. 60 sekund to liczba, przy której użytkownik jeszcze czeka, a nie wraca do starej metody. Liczba w wymaganiu jest **decyzją biznesową**, nie techniczną — i dlatego musi być uzgodniona z odbiorcą, a nie wymyślona przez programistę.

**Dlaczego `N6` jest osobno od `N4`.** Bo to są **dwie różne właściwości**. Idempotencja (`N4`) mówi: „drugie uruchomienie nie psuje pliku". Deterministyczność (`N6`) mówi: „pierwsze uruchomienie na tej samej maszynie i na innej daje ten sam wynik". Możesz mieć `N4` i nie mieć `N6` — jeżeli w kodzie jest `datetime.now()` w logice, drugie uruchomienie **nadpisze** plik identycznym wynikiem (idempotencja OK), ale na **innej** maszynie wynik będzie inny (determinizm zły).

### 2.4. Architektura rozwiązania

```text
raporty/                                    <- katalog projektu
├── pyproject.toml
├── README.md
├── assets/
│   └── logo.png
├── config/
│   ├── raport_sprzedazy.yaml               <- konfiguracja produkcyjna
│   └── raport_minimal.yaml                 <- konfiguracja minimalistyczna
├── data/
│   └── orders.csv                          <- WEJSCIE (tylko odczyt!)
├── output/                                 <- wszystko, co generujemy
├── src/raporty/
│   ├── __init__.py
│   ├── __main__.py                         <- CLI (python -m raporty)
│   ├── domain/                             <- ZERO importow openpyxl
│   │   ├── __init__.py
│   │   ├── modele.py                       <- Zamowienie, WierszRaportu, ModelRaportu
│   │   ├── reguly.py                       <- marza, dynamika, sanitacja
│   │   ├── paleta.py                        <- ZAMKNIETA PALETA regul
│   │   └── porty.py                         <- EksporterRaportu, ZrodloDanych, ...
│   ├── application/                        <- ZERO importow openpyxl
│   │   ├── __init__.py
│   │   ├── raport_bazowy.py                <- BazowyRaport (Template Method)
│   │   ├── raport_sprzedazy.py             <- RaportSprzedazy
│   │   ├── raport_minimal.py               <- RaportMinimal
│   │   └── komendy.py                      <- Command + StosKomend
│   └── infrastructure/                     <- TU WOLNO importowac openpyxl
│       ├── __init__.py
│       ├── excel/
│       │   ├── __init__.py
│       │   ├── styl.py                     <- RejestrStylow (nazwane style)
│       │   ├── eksporter.py                <- EksporterXlsxOpenpyxl
│       │   ├── wykresy.py                  <- wykres dwuosiowy
│       │   ├── repository.py               <- JednostkaPracy (atomowy zapis)
│       │   ├── eksporter_danych.py         <- write_only (duzy plik)
│       │   └── utrata.py                   <- skan utraty funkcji
│       ├── zrodla/
│       │   ├── __init__.py
│       │   ├── csv_zrodlo.py               <- ZrodloCsv
│       │   └── fake_zrodlo.py              <- ZrodloFake (testy)
│       ├── konfig/
│       │   ├── __init__.py
│       │   └── loader.py                   <- RaportSpec + walidacja
│       ├── fake/
│       │   └── eksporter.py                <- EksporterFake
│       └── audyt/
│           └── dziennik.py                 <- dziennik JSONL + _meta
└── tests/
    ├── conftest.py
    ├── domain/                             <- bez openpyxl
    ├── application/                        <- bez openpyxl
    └── infrastructure/excel/               <- z openpyxl
```

**Reguła zależności (z modułu 26, powtórzona tu jako twarda):**

```bash
# Musi zwrócić PUSTY wynik:
grep -rn "openpyxl" src/raporty/domain/ src/raporty/application/
# Musi zwrócić wyłącznie pliki z infrastructure/:
grep -rn "openpyxl" src/raporty/ | cut -d: -f1 | sort -u
```

**Wzorce, których używa ten projekt (i po co):**

| Wzorzec | Gdzie | Po co konkretnie |
|---|---|---|
| **Ports & Adapters** | `domain/porty.py` + `infrastructure/excel/eksporter.py` | testy bez openpyxl; wymiana na CSV/PDF bez zmian w logice |
| **Template Method** | `application/raport_bazowy.py` | dwa raporty (pełny i minimalny) dzielą szkielet `generuj()` |
| **Factory Method** | `__main__.py` (słownik `eksportery`) | wybór adaptera po nazwie z CLI |
| **Registry / Flyweight** | `infrastructure/excel/styl.py` | style tworzone raz i współdzielone przez komórki |
| **Command** | `application/komendy.py` | kolejność operacji na skoroszycie, log audytowy, testowalność |
| **Repository / Unit of Work** | `infrastructure/excel/repository.py` | jeden transakcyjny zapis, backup, deterministyczny commit |
| **Specification (paleta)** | `domain/paleta.py` | reguły wybierane z konfiguracji przez **nazwę**, nie przez wyrażenie |
| **Facade** | `EksporterXlsxOpenpyxl` | aplikacja woła `eksportuj()`, nie zna `Workbook`/`Table`/`DataBarRule` |

**Jedna decyzja, którą podkreślam szczególnie.** Wszystkie operacje na skoroszycie idą przez `StosKomend` (`infrastructure/excel/repository.py` + `application/komendy.py`), a nie przez bezpośrednie `ws["A1"] = ...` w kodzie aplikacji. Dzięki temu:

1. każda zmiana ma **opis** (idzie do logu audytowego),
2. kolejność zmian jest **jawna** i **testowalna** (można wykonać stos na `Workbook` w pamięci i sprawdzić),
3. `wb.save()` jest wywoływane **raz**, na końcu (`JednostkaPracy.zatwierdz`),
4. moduł 25 (`StosKomend` z cofaniem) jest tu **realnie użyty**, a nie tylko zademonstrowany.

### 2.5. Konfiguracja — co jest w YAML-u, a czego nie ma

Zasada z modułu 26 (sekcja 2.6): **konfiguracja opisuje CO, kod opisuje JAK**.

```yaml
# config/raport_sprzedazy.yaml
raport:
  identyfikator: sprzedaz_miesieczna
  tytul: "Raport sprzedazy miesiecznej"
  okres: "2026-09"
  schemat: v3
  wersja_generatora: "27.0"

dane:
  sciezka: data/orders.csv
  kodowanie: utf-8-sig
  separator: ","
  mapowanie:                       # nazwy kolumn w CSV -> pola modelu
    data: "Data zamowienia"
    region: "Region"
    klient: "Nabywca"
    ilosc: "Ilosc"
    cena_netto: "Cena netto"
    koszt_netto: "Koszt netto"

wyjscie:
  katalog: output
  nazwa: raport_sprzedazy
  polityka: wersjonuj              # nadpisz | wersjonuj | data
  zapisuj_dane: true               # czy tworzyc drugi plik (Dane, write_only)

kolumny_podsumowanie:
  - { nazwa: "Region",           zrodlo: region,        format: text }
  - { nazwa: "Pozycji",          zrodlo: liczba_pozycji, format: int }
  - { nazwa: "Sprzedaz netto",   zrodlo: netto_pln,     format: currency }
  - { nazwa: "Marza %",          zrodlo: marza_proc,    format: percent }

kpi:
  - { etykieta: "Sprzedaz netto",   agregat: suma_netto,    format: currency }
  - { etykieta: "Marza zbiorcza",   agregat: marza_zbiorcza, format: percent }
  - { etykieta: "Regionow",         agregat: liczba_regionow, format: int }

reguly:                             # WYBOR NAZW z zamknietej palety
  - nazwa: spadek_dynamiki
    tlo: "FCE4E4"
    kolor: "9C0006"
  - nazwa: marza_ponizej_progu
    parametr: 0.05
    tlo: "FFF2CC"

pasek_danych:
  kolumna: "Dynamika %"
  zakres: [-0.5, 0.5]

wykres:
  typ: sprzedaz_marza_region
  tytul: "Sprzedaz i marza per region"

arkusze:
  ochrona_haslo: "raport2026"
  ukryj_meta: true
```

**Czego w tym YAML-u NIE MA i dlaczego:**

- **Żadnych wyrażeń.** Nie ma `warunek: "marza < prog"`. Jest `nazwa: marza_ponizej_progu` + `parametr: 0.05`. Konfiguracja **wybiera z palety**, nie tworzy reguł.
- **Żadnych kodów formatów Excela.** Jest `format: currency`, a nie `'#,##0.00 "zl";(#,##0.00)'`. Kod formatu jest **szczegółem reprezentacji** i należy do adaptera (warstwa infrastruktury). Zmiana symbolu waluty to zmiana w infrastrukturze, nie w konfiguracji użytkownika.
- **Żadnego `if`/`when`/`condition`.** Test: `grep -n "if \|when \|condition" config/*.yaml` musi być pusty.
- **Żadnych ścieżek absolutnych.** Wszystkie ścieżki są względne wobec katalogu projektu. Wyjątek: zmienna środowiskowa `RAPORTY_KATALOG` jako opcjonalne nadpisanie dla wdrożenia.

**Paleta reguł w kodzie** (`domain/paleta.py`) — to jest kontrakt:

```python
REGULY: dict[str, tuple[Callable, tuple[str, ...]]] = {
    "spadek_dynamiki":       (spadek_dynamiki, ()),
    "marza_ponizej_progu":   (marza_ponizej_progu, ("parametr",)),
    "brak_sprzedazy":        (brak_sprzedazy, ()),
    "duzy_klient":           (duzy_klient, ("parametr",)),
}
```

Krotka przy nazwie wymienia **wymagane parametry**. Konfiguracja, która podaje `marza_ponizej_progu` bez `parametr`, zostanie **odrzucona przy wczytaniu** z komunikatem „reguła `marza_ponizej_progu` wymaga parametru `parametr`". To jest walidacja konfiguracji w jednym miejscu — w module 26 nazwaliśmy to „zamykaniem drzwi do języka w konfiguracji".

### 2.6. Etapy realizacji (kamienie milowe M1–M10)

**Zasada:** po każdym kroku **kod działa** i **testy przechodzą**. Nie przechodzisz dalej, dopóki nie masz zielonego.

| Nr | Cel | Kod, który powstaje | Test, który przechodzi | Wymaganie |
|---|---|---|---|---|
| **M1** | Uruchamialny szkielet | `__main__.py`, `Workbook()` + `ws.append` + `save` | Test: plik powstaje i ma jeden arkusz | — |
| **M2** | Model domeny + sanitacja | `modele.py`, `reguly.py`, `paleta.py` | 8 testów walidacji, sanitacji, marży | F11 |
| **M3** | Port + adapter + fake | `porty.py`, `eksporter.py`, `fake/eksporter.py` | Test kontraktu, testy domeny **bez openpyxl** | N8 |
| **M4** | Rejestr stylów + arkusze | `styl.py`, arkusze `Podsumowanie`, `Regiony`, `_meta` | Test: `wb.sheetnames ==` oczekiwana krotka | F1 |
| **M5** | Formatowanie warunkowe + pasek danych | reguły z palety, `DataBarRule` | Liczba reguł, `sqref` | F5, F7 |
| **M6** | Wykres dwuosiowy | `wykresy.py` | `xl/charts/chart1.xml` w ZIP | F6 |
| **M7** | Tabela + logo + Metodyka | `Table`, `Image`, arkusz `Metodyka` | `ws.tables`, `xl/media/` | F2, F3, F8 |
| **M8** | Ochrona + walidacja + właściwości | `protection`, `DataValidation`, `wb.properties` | `ws.protection.sheet`, `data_validations` | F9, F10, F12 |
| **M9** | Zapis transakcyjny + audyt | `repository.py`, `utrata.py`, `dziennik.py` | Test atomowości, brak `.tmp`, backup | N3, N9, N10 |
| **M10** | Duże pliki (write_only) + CI + determinizm | `eksporter_danych.py`, testy `slow`, `conftest.py` | 200k < 60 s, pamięć < 500 MB, dwa hashe zgodne | N1, N2, N4, N6, N7 |

**Dlaczego akurat taka kolejność.** Trzy rzeczy są w tej kolejności nieprzypadkowe:

1. **M2 przed M3.** Najpierw domena bez infrastruktury, potem port i adapter. Gdyby było odwrotnie, „domena" powstałaby z myślą o Excelu i **przeciekłaby** (moduł 26, pułapka 1).
2. **M3 przed M4–M8.** Dzięki temu **wszystkie testy od M4 wzwyż piszesz w dwóch wersjach**: przez `EksporterFake` (szybko, bez openpyxl) i przez prawdziwego eksportera (z openpyxl). Gdyby port powstał na końcu, cała praca M4–M8 byłaby nieprzetestowalna.
3. **M9 przed M10.** Transakcyjność najpierw na małych plikach, potem wydajność. Odwrotność prowadzi do koszmaru: optymalizujesz kod, którego jeszcze nie umiesz bezpiecznie zapisać, więc każdy błąd niszczy plik i tracisz dane testowe.

### 2.7. Kryteria oceny (rubryka)

To jest rubryka, według której ocenia się rozwiązanie. Używaj jej **na sobie**, po każdym kamieniu milowym — nie tylko na końcu.

| Kryterium | Waga | Na co patrzę (dowód) |
|---|---|---|
| **Poprawność danych** | ★★★ | Kwoty to `Decimal`→`float`, daty to `datetime`, brak tekstu tam, gdzie ma być liczba; sumy w Excelu liczą się poprawnie po otwarciu |
| **Poprawność typów** | ★★★ | Test: `isinstance(ws["D2"].value, (int, float))`, `isinstance(ws["A2"].value, datetime)`; brak `#REF!` |
| **Czytelność** | ★★☆ | Każda funkcja ≤ 40 linii; nazwy mówią **co**, nie **jak**; brak `magic numbers` (kolory tylko z rejestru) |
| **Brak przecieku abstrakcji** | ★★★ | `grep -rn "openpyxl" src/raporty/domain/ src/raporty/application/` → pusto; port nie ma typów openpyxl |
| **Wydajność** | ★★★ | 200k < 60 s, pamięć < 500 MB zmierzone (`perf_counter`, `tracemalloc`) |
| **Bezpieczeństwo** | ★★★ | Sanitacja na wejściu do domeny; brak `=`, `+`, `-`, `@` w danych przechodzących dalej; testy parametryzowane |
| **Testy** | ★★★ | > 80 % logiki domenowej; testy domeny i aplikacji **bez openpyxl**; wszystkie zielone w CI |
| **Architektura** | ★★☆ | Domena → aplikacja → infrastruktura; jeden kierunek; jeden plik kompozycji |
| **Obserwowalność** | ★★☆ | Jedno zdarzenie logu na raport z pełnym kompletem faktów; ostrzeżenia z utraty funkcji |
| **Transakcyjność** | ★★★ | Test: wyjątek w połowie → plik docelowy nietknięty, brak `.tmp`, backup z timestampem |
| **Dokumentacja** | ★☆☆ | `README.md`: jak uruchomić, jak zmienić konfigurację, co jest w `output/` |

Trzy gwiazdki = bez tego rozwiązanie **nie przechodzi odbioru**.

## 3. Przykłady krok po kroku

Przechodzimy przez wszystkie dziesięć kamieni milowych. Każdy pokazuje **najważniejszy fragment** — nie cały plik — i **test**, który na tym etapie przechodzi. Kod jest kumulatywny: to, co powstało w M1, jest rozszerzane w M2 i tak dalej.

### M1 — uruchamialny szkielet

Cel: mieć **cokolwiek**, co się uruchamia i tworzy plik. Wiele osób próbuje zacząć od pełnej wersji i po trzech dniach ma 700 linii i błąd `AttributeError` w linii 512. Nie rób tego.

```python
# src/raporty/__main__.py (wersja M1)
"""CLI generatora raportow. Na tym etapie: tylko szkielet."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook


def kompozycja(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="raporty")
    parser.add_argument("--config", required=True)
    parser.add_argument("--wyjscie", default="output/raport")
    args = parser.parse_args(argv)

    wb = Workbook()
    ws = wb.active
    ws.title = "Podsumowanie"
    ws["A1"] = "Raport sprzedazy"
    ws["A2"] = f"wygenerowano: {datetime.now(timezone.utc):%Y-%m-%d %H:%M}"

    cel = Path(args.wyjscie).with_suffix(".xlsx")
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)
    print(f"OK {cel}")
    return 0


if __name__ == "__main__":
    raise SystemExit(kompozycja(sys.argv[1:]))
```

```python
# tests/infrastructure/excel/test_m1_szkielet.py
def test_powstaje_plik(tmp_path):
    kod = kompozycja(["--config", "config/raport_sprzedazy.yaml",
                      "--wyjscie", str(tmp_path / "raport")])
    assert kod == 0
    assert (tmp_path / "raport.xlsx").exists()

    from openpyxl import load_workbook
    wb = load_workbook(tmp_path / "raport.xlsx")
    assert wb.sheetnames == ["Podsumowanie"]
```

**Co się dzieje w pamięci.** `Workbook()` tworzy skoroszyt z **jednym domyślnym arkuszem** o nazwie `Sheet` — dlatego `ws.title = "Podsumowanie"` to **zmiana nazwy**, nie utworzenie nowego arkusza. To jest źródło niespodzianki z pułapki modułu 04: „dlaczego mam cztery arkusze, jak tworzyłem trzy".

**Co trafi do pliku.** Minimum: `xl/workbook.xml` z jednym arkuszem, `xl/worksheets/sheet1.xml` z dwiema komórkami, `xl/styles.xml` z zerem niestandardowych stylów. Plik ma kilkanaście KB.

**Test na tym etapie jest banalny — i to jest sens.** Wiesz, że rura jest podłączona. Nie musisz jeszcze wiedzieć, co płynie.

### M2 — model domeny i sanitacja

Cel: **dane i reguły bez Excela**. To jest najważniejszy krok całego projektu, bo wszystko dalej zależy od tego, jak wygląda model.

```python
# src/raporty/domain/modele.py
"""Modele domenowe. ZERO openpyxl. ZERO plikow."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP


@dataclass(frozen=True)
class Zamowienie:
    """Jedna pozycja sprzedazy. Kwoty w GROSZACH (int) - unikamy float."""

    data: date
    region: str
    klient: str
    produkt: str
    ilosc: int
    cena_grosze: int        # cena netto jednostkowa w groszach
    koszt_grosze: int       # koszt netto jednostkowy w groszach

    @property
    def netto_grosze(self) -> int:
        return self.ilosc * self.cena_grosze

    @property
    def marza_grosze(self) -> int:
        return self.ilosc * (self.cena_grosze - self.koszt_grosze)


@dataclass(frozen=True)
class AgregatRegionu:
    """Podsumowanie po regionie. To trafia na arkusz 'Regiony'."""

    region: str
    liczba_pozycji: int
    netto_pln: Decimal
    marza_proc: Decimal


@dataclass(frozen=True)
class WierszDynamiki:
    """Wiersz arkusza 'Dynamika'. Zawiera biezacy i poprzedni okres."""

    region: str
    netto_poprzedni_pln: Decimal
    netto_biezacy_pln: Decimal
    dynamika_proc: Decimal


@dataclass(frozen=True)
class Kpi:
    etykieta: str
    wartosc: object
    format_hint: str          # "currency" | "percent" | "int" | "text"


@dataclass(frozen=True)
class ModelRaportu:
    """Zupelny opis raportu. Bez kolorow, bez arkuszy, bez openpyxl."""

    tytul: str
    okres: str
    schemat: str
    wersja_generatora: str
    hash_wsadu: str
    wygenerowano: str
    kpi: tuple[Kpi, ...]
    regiony: tuple[AgregatRegionu, ...]
    dynamika: tuple[WierszDynamiki, ...]
    kolumny_regiony: tuple[tuple[str, str], ...]   # (nazwa, format_hint)
    metodyka: tuple[str, ...]
    ostrzezenia: tuple[str, ...] = ()
```

**Uwaga do `Decimal` vs `float`.** Wewnątrz modelu wszystko jest `Decimal` (marże, agregaty), a `int` (grosze) tam, gdzie to możliwe. `float` pojawia się **dopiero w adapterze**, bo Excel przechowuje liczby jako double i nie da się tego obejść. Ale **konwersja jest jawna i w jednym miejscu** — a nie rozsiana po całym kodzie. To jest ta różnica między „wiem, że mam problem" a „nie wiem".

```python
# src/raporty/domain/reguly.py (fragmenty kluczowe)
ZNAKI_WYZWALAJACE = ("=", "+", "-", "@", "\t", "\r")


class BladDomeny(Exception):
    """Blad reguly biznesowej lub danych. Nie jest bledem technicznym."""


def przyjmij_tekst_zewnetrzny(wartosc: object, *, pole: str) -> str:
    """Odrzuca podejrzane teksty. NIE modyfikuje danych klienta."""
    if wartosc is None:
        return ""
    tekst = str(wartosc)
    if len(tekst) > 32767:
        raise BladDomeny(f"{pole!r}: tekst dluzszy niz limit komorki Excela")
    if tekst and tekst[0] in ZNAKI_WYZWALAJACE:
        raise BladDomeny(
            f"{pole!r}: tekst zaczyna sie od {tekst[0]!r} "
            f"(formula injection). Odrzucamy."
        )
    return "".join(ch for ch in tekst if ch == " " or ch.isprintable()).strip()


def marza_zbiorcza(zamowienia: list[Zamowienie]) -> Decimal:
    """Iloraz sum, nie srednia marz. To jest reguła biznesowa."""
    netto = sum(z.netto_grosze for z in zamowienia)
    marza = sum(z.marza_grosze for z in zamowienia)
    if netto == 0:
        return Decimal("0")
    return (Decimal(marza) / Decimal(netto)).quantize(
        Decimal("0.0001"), rounding=ROUND_HALF_UP
    )


def dynamika_proc(biezacy: Decimal, poprzedni: Decimal) -> Decimal:
    """Dynamika okres/okres. Przy zerze poprzednim zwracamy 0 (nie dzielimy)."""
    if poprzedni == 0:
        return Decimal("0")
    return ((biezacy - poprzedni) / poprzedni).quantize(
        Decimal("0.0001"), rounding=ROUND_HALF_UP
    )
```

```python
# tests/domain/test_sanitacja.py
import pytest
from src.raporty.domain.reguly import BladDomeny, przyjmij_tekst_zewnetrzny


@pytest.mark.parametrize("znak", ["=", "+", "-", "@", "\t", "\r"])
def test_odrzuca_kazdy_znak_wyzwaajacy(znak):
    with pytest.raises(BladDomeny, match="formula injection"):
        przyjmij_tekst_zewnetrzny(f"{znak}zle", pole="klient")


def test_przyjmuje_poprawny_tekst():
    assert przyjmij_tekst_zewnetrzny("  Kowalski  ", pole="klient") == "Kowalski"


def test_odrzuca_tekst_za_dlugi():
    with pytest.raises(BladDomeny, match="limit komorki"):
        przyjmij_tekst_zewnetrzny("x" * 40000, pole="klient")


def test_none_daje_pusty_string():
    assert przyjmij_tekst_zewnetrzny(None, pole="klient") == ""
```

**To jest mały plik i wielka decyzja.** Sanitacja **odrzuca**, nie modyfikuje. Jeżeli klient nazywa się `-Kowalski` (bo tak ma w systemie), to **człowiek** ma zdecydować, co z tym zrobić — a nie program cicho zmieniać dane. To jest różnica między „program pomaga" a „program ukrywa problem".

### M3 — port, adapter, fake

Cel: **odać openpyxl od logiki**. Powtarzamy konstrukcję z modułu 26, ale tym razem w kontekście realnego projektu.

```python
# src/raporty/domain/porty.py
from __future__ import annotations
from dataclasses import dataclass
from typing import Protocol
from .modele import ModelRaportu


@dataclass(frozen=True)
class WynikEksportu:
    format: str
    bajtow: int
    arkuszy: int
    wierszy: int
    sciezki: tuple[str, ...]
    hash_wyniku: str = ""
    ostrzezenia: tuple[str, ...] = ()


class EksporterRaportu(Protocol):
    """PORT. Zero typow openpyxl w sygnaturze."""

    @property
    def format(self) -> str: ...

    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu: ...
```

```python
# src/raporty/infrastructure/fake/eksporter.py
@dataclass
class EksporterFake:
    _zapisy: list[tuple[ModelRaportu, str]] = field(default_factory=list)
    _awaria: Exception | None = None

    def ustaw_awarie(self, blad: Exception) -> "EksporterFake":
        self._awaria = blad
        return self

    @property
    def format(self) -> str:
        return "fake"

    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu:
        if self._awaria is not None:
            blad, self._awaria = self._awaria, None
            raise blad
        self._zapisy.append((raport, cel))
        return WynikEksportu(
            format="fake", bajtow=0, arkuszy=1,
            wierszy=len(raport.regiony), sciezki=(),
            hash_wyniku=f"fake-{len(self._zapisy)}",
        )

    def ostatni(self) -> tuple[ModelRaportu, str]:
        if not self._zapisy:
            raise AssertionError("EksporterFake: nic nie zapisano")
        return self._zapisy[-1]
```

**Test tego etapu to kamień węgielny całego projektu** — test, który uruchomi się na maszynie **bez openpyxl**:

```python
# tests/application/test_generuj_fake.py
def test_raport_sprzedazy_generuje_model(eksporter_fake, zrodlo_fake):
    przypadek = GenerujRaportSprzedazy(zrodlo=zrodlo_fake, eksporter=eksporter_fake)
    wynik = przypadek.wykonaj(okres="2026-09", cel="output/x")
    model, cel = eksporter_fake.ostatni()
    assert model.okres == "2026-09"
    assert len(model.regiony) == 3
    assert wynik.wynik.format == "fake"
```

To jest ten moment, w którym projekt staje się **testowalny**. Wszystko od M4 wzwyż dostanie dwa testy: przez `EksporterFake` (szybki) i przez prawdziwego eksportera (rzetelny).

### M4 — rejestr stylów i arkusze

Cel: **koniec z kolorami wpisanymi w 40 miejscach**. Jeden rejestr, nazwane style (moduł 09), współdzielone obiekty.

```python
# src/raporty/infrastructure/excel/styl.py
"""RejestrStylow - jedna matryca, tysiace nadrukow.

Analogia z modulu 09: drukarnia z gotowymi matrycami. Nie rzezbimy
czcionki dla kazdego wyrazu - tworzymy styl RAZ i przypinamy go komorkom.
"""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side

# Slownik kodow formatow - JEDNO miejsce dla calej reprezentacji (modul 10).
FORMATY: dict[str, str] = {
    # Polska waluta: separator tysiecy spacją, 2 miejsca, minus w nawiasie
    "currency": '#,##0.00 "zl";(#,##0.00)',
    "percent": "0,0%",
    "int": "#,##0",
    "date": "yyyy-mm-dd",
    "text": "@",
    "dynamika": "0.0%;[Red]-0.0%",
}

# Paleta kolorow - JEDNO miejsce. Zadnego "FFC000" w kodzie adaptera!
KOLORY = {
    "granat": "1F4E78",
    "granat_jasny": "D9E2F3",
    "biel": "FFFFFF",
    "czerwone_tlo": "FCE4E4",
    "czerwony_tekst": "9C0006",
    "zielone_tlo": "E2EFDA",
    "zielony_tekst": "375623",
    "szary_tekst": "595959",
}


@dataclass(frozen=True)
class RejestrStylow:
    """Nazwane style skoroszytu. Tworzone raz, uzywane wielokrotnie."""

    tytul: NamedStyle
    naglowek: NamedStyle
    naglowek_granat: NamedStyle
    komorka_tekst: NamedStyle
    komorka_liczba: NamedStyle
    komorka_waluta: NamedStyle
    komorka_procent: NamedStyle
    kpi_etykieta: NamedStyle
    kpi_wartosc: NamedStyle
    metodyka_tekst: NamedStyle
    meta_etykieta: NamedStyle

    def dodaj_do(self, wb) -> None:
        for styl in self._wszystkie():
            wb.add_named_style(styl)

    def _wszystkie(self) -> tuple[NamedStyle, ...]:
        return (self.tytul, self.naglowek, self.naglowek_granat,
                self.komorka_tekst, self.komorka_liczba, self.komorka_waluta,
                self.komorka_procent, self.kpi_etykieta, self.kpi_wartosc,
                self.metodyka_tekst, self.meta_etykieta)


def zbuduj_rejestr() -> RejestrStylow:
    """Tworzy komplet stylow. Wywolywane RAZ na skoroszyt."""
    cienka = Side(style="thin", color=KOLORY["granat_jasny"])
    obramowanie_naglowka = Border(left=cienka, right=cienka, top=cienka, bottom=cienka)

    return RejestrStylow(
        tytul=NamedStyle(
            name="RaportTytul",
            font=Font(name="Calibri", size=18, bold=True, color=KOLORY["granat"]),
            alignment=Alignment(vertical="center"),
        ),
        naglowek=NamedStyle(
            name="RaportNaglowek",
            font=Font(bold=True, size=11),
            fill=PatternFill(fill_type="solid", fgColor=KOLORY["granat_jasny"]),
            alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
            border=obramowanie_naglowka,
        ),
        naglowek_granat=NamedStyle(
            name="RaportNaglowekGranat",
            font=Font(bold=True, size=11, color=KOLORY["biel"]),
            fill=PatternFill(fill_type="solid", fgColor=KOLORY["granat"]),
            alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
        ),
        komorka_tekst=NamedStyle(name="RaportTekst", number_format=FORMATY["text"]),
        komorka_liczba=NamedStyle(name="RaportLiczba", number_format=FORMATY["int"]),
        komorka_waluta=NamedStyle(name="RaportWaluta", number_format=FORMATY["currency"]),
        komorka_procent=NamedStyle(name="RaportProcent", number_format=FORMATY["percent"]),
        kpi_etykieta=NamedStyle(
            name="RaportKpiEtykieta",
            font=Font(bold=True, size=11, color=KOLORY["szary_tekst"]),
        ),
        kpi_wartosc=NamedStyle(
            name="RaportKpiWartosc",
            font=Font(bold=True, size=16, color=KOLORY["granat"]),
        ),
        metodyka_tekst=NamedStyle(
            name="RaportMetodyka",
            alignment=Alignment(wrap_text=True, vertical="top"),
        ),
        meta_etykieta=NamedStyle(
            name="RaportMetaEtykieta",
            font=Font(bold=True, size=10),
        ),
    )
```

I użycie — `cell.style = "RaportWaluta"`:

```python
for indeks, agregat in enumerate(raport.regiony, start=2):
    ws.cell(row=indeks, column=1, value=agregat.region).style = "RaportTekst"
    ws.cell(row=indeks, column=2, value=agregat.liczba_pozycji).style = "RaportLiczba"
    ws.cell(row=indeks, column=3, value=float(agregat.netto_pln)).style = "RaportWaluta"
    ws.cell(row=indeks, column=4, value=float(agregat.marza_proc)).style = "RaportProcent"
```

**Trzy rzeczy do przemyślenia.**

1. **`cell.style = "Nazwa"`** to jedyna forma nadawania stylu, w której nie tworzysz obiektu na komórkę. Jest szybka, czytelna i — co najważniejsze — **spójna z konfiguracją w `styl.py`**. Gdy zmieniasz format waluty, zmieniasz **jedną** linię w `FORMATY`.
2. **Nazwane style są w `wb`.** `wb.add_named_style()` musi być wywołane **przed** użyciem nazwy. Kolejność ma znaczenie — kompilator tego nie sprawdzi, więc kolejność jest w `eksporter.py` na początku `eksportuj()`.
3. **Nie tworzymy `Font` w pętli po 200 000 wierszach.** Wszystkie style są utworzone raz w `zbuduj_rejestr()`. To jest Flyweight (moduł 23) w praktyce i jednocześnie optymalizacja pamięci i czasu.

### M5 — formatowanie warunkowe i pasek danych

Cel: **wyróżnić spadki i pokazać dynamikę bez malowania komórek**.

```python
# fragment eksportera - arkusz "Dynamika"
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Font, PatternFill


def _arkusz_dynamika(self, wb, raport) -> None:
    ws = wb.create_sheet("Dynamika")
    ws.sheet_view.showGridLines = False

    naglowki = ("Region", "Poprzedni okres", "Biezacy okres", "Dynamika %")
    for kolumna, naglowek in enumerate(naglowki, start=1):
        ws.cell(row=1, column=kolumna, value=naglowek).style = "RaportNaglowekGranat"

    for indeks, wiersz in enumerate(raport.dynamika, start=2):
        ws.cell(row=indeks, column=1, value=wiersz.region).style = "RaportTekst"
        (ws.cell(row=indeks, column=2,
                 value=float(wiersz.netto_poprzedni_pln)).style) = "RaportWaluta"
        ws.cell(row=indeks, column=3,
                value=float(wiersz.netto_biezacy_pln)).style = "RaportWaluta"
        ws.cell(row=indeks, column=4,
                value=float(wiersz.dynamika_proc)).style = "RaportProcent"

    ostatni = len(raport.dynamika) + 1
    zakres_dynamiki = f"D2:D{ostatni}"

    # REGULA 1: pasek danych - "wykres zrobiony z wypelnienia" (modul 14).
    # start_value ujemny, bo dynamika moze byc ujemna (spadki)!
    pasek = DataBarRule(
        start_type="num", start_value=-0.5,
        end_type="num", end_value=0.5,
        color="638EC6", showValue=True,
    )
    ws.conditional_formatting.add(zakres_dynamiki, pasek)

    # REGULA 2: spadki - tlo i kolor tekstu. Konfiguracja przyszla z reguly
    # "spadek_dynamiki" z palety (nazwa -> kolor z YAML).
    spadek = CellIsRule(
        operator="lessThan",
        formula=["0"],
        fill=PatternFill(fill_type="solid", start_color="FCE4E4",
                         end_color="FCE4E4"),
        font=Font(color="9C0006", bold=True),
    )
    ws.conditional_formatting.add(zakres_dynamiki, spadek)

    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 24
    for litera in ("B", "C", "D"):
        ws.column_dimensions[litera].width = 18
```

**Uwaga o kolejności reguł.** `DataBarRule` i `CellIsRule` nakładają się na tym samym zakresie. Pasek danych rysuje się **za** tekstem, więc wizualnie nie kolidują — ale **kolejność priorytetów** ma znaczenie przy bardziej złożonych regułach. Reguła dodana **pierwsza** otrzymuje **niższy** numer priorytetu (mniejsza liczba = ważniejsza; moduł 13). Tutaj to nie szkodzi, ale w projekcie z zielonym tłem dla wzrostów i czerwonym dla spadków — **musisz** pamiętać o `rule.priority` i o `stopIfTrue`, jeśli reguły mają się wykluczać.

**Co trafi do pliku.** Element `<conditionalFormatting sqref="D2:D37">` z **dwoma** `<cfRule>`. Pamiętaj: w komórce **nie ma** zapisanego koloru — kolor pojawi się dopiero, gdy Excel otworzy plik. Dlatego test nie sprawdzi „czy komórka D2 jest czerwona", a tylko „czy istnieje reguła o zakresie D2:D37".

### M6 — wykres dwuosiowy

Cel: **sprzedaż jako słupki, marża jako linia na drugiej osi**. To wymaganie `F6`.

```python
# src/raporty/infrastructure/excel/wykresy.py
"""Wykres dwuosiowy. BarChart (sprzedaz) + LineChart (marza) na osi 200."""

from __future__ import annotations

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties

KOLOR_SLUPEK = "1F4E78"
KOLOR_LINII = "C00000"


def dodaj_wykres_sprzedaz_marza(ws, *, wiersze: int, anchor: str,
                                 kolumna_region: int = 1,
                                 kolumna_netto: int = 3,
                                 kolumna_marza: int = 4) -> None:
    """Buduje wykres dwuosiowy na podstawie gotowego zakresu agregatow.

    Zalozenia:
      - wiersz 1 to naglowki (titles_from_data=True),
      - kolumna 1 to kategorie (regiony),
      - wiersze 2..wiersze+1 to dane.
    """
    ostatni = wiersze + 1

    slupki = BarChart()
    slupki.type = "col"
    slupki.style = 10
    slupki.title = "Sprzedaz i marza per region"
    slupki.y_axis.title = "Sprzedaz netto (zl)"
    slupki.x_axis.title = "Region"
    slupki.y_axis.majorGridlines = ChartLines()
    slupki.y_axis.numFmt = '#,##0'
    slupki.gapWidth = 60

    dane_netto = Reference(ws, min_col=kolumna_netto, min_row=1, max_row=ostatni)
    kategorie = Reference(ws, min_col=kolumna_region, min_row=2, max_row=ostatni)
    slupki.add_data(dane_netto, titles_from_data=True)
    slupki.set_categories(kategorie)

    if slupki.series:
        slupki.series[0].graphicalProperties = GraphicalProperties(
            solidFill=KOLOR_SLUPEK
        )

    linia = LineChart()
    dane_marzy = Reference(ws, min_col=kolumna_marza, min_row=1, max_row=ostatni)
    linia.add_data(dane_marzy, titles_from_data=True)

    if linia.series:
        linia.series[0].graphicalProperties = GraphicalProperties(
            solidFill=KOLOR_LINII, ln=None
        )

    # DRUGA OS: nowy axId i przesuniecie osi sprzedazy do krawedzi.
    # To jest ten fragment, ktory wszyscy znajduja po godzinie szukania.
    linia.y_axis.axId = 200
    linia.y_axis.title = "Marza (%)"
    linia.y_axis.numFmt = "0.0%"
    linia.y_axis.majorGridlines = None

    slupki.y_axis.crosses = "max"     # os sprzedazy na krawedzi wykresu
    slupki += linia                   # zgodnie z dokumentacja 3.1.x

    slupki.height = 9.5
    slupki.width = 20.0
    ws.add_chart(slupki, anchor)
```

**Co się dzieje w pamięci.** `BarChart()` i `LineChart()` to **dwa osobne obiekty**. `slupki += linia` **skleja** serie obu do jednego wykresu — i to jest ten moment, w którym linia ląduje na osi o `axId=200`. Bez tego zabiegu obie serie są na **jednej** osi, a skala „sprzedaży" (setki tysięcy) całkowicie spłaszcza „marżę" (0,05) do linii prostej na dole. To jest najczęstszy błąd przy wykresach dwuosiowych: **zapomnisz o `axId` i masz wykres, który wygląda jak uszkodzony**.

**Dlaczego `graphicalProperties` a nie `color`.** Na serii wykresu nie ma atrybutu `color`. Kolor słupka to `series.graphicalProperties = GraphicalProperties(solidFill="...")`. Jest to niespójne z API komórek (`cell.fill.fgColor`) i jest to **klasyczna pułapka** — 30 minut szukania atrybutu, który nie istnieje.

**Co trafi do pliku.** `xl/charts/chart1.xml` z **trzema** osiami (kategoria, wartość `axId=100`, wartość `axId=200`) i dwiema seriami. Wykres „pływa" nad siatką, zakotwiczony w komórce `anchor` (np. `"F2"`).

**Test tego etapu:**

```python
def test_wykres_trafia_do_zip(wynik_xlsx):
    import zipfile
    with zipfile.ZipFile(wynik_xlsx) as archiwum:
        czesci = [n for n in archiwum.namelist() if n.startswith("xl/charts/")]
    assert len(czesci) == 1
    assert "xl/charts/chart1.xml" in czesci
```

Test sprawdza **obecność** wykresu. Nie sprawdzi, czy dwie osie działają — to jest ta część, o której mówiłem w sekcji 1.5: **weryfikacja na oko**.

### M7 — tabela, logo, Metodyka

Cel: trzy wymagania naraz — `F2` (logo), `F3` (tabela), `F8` (Metodyka).

```python
# fragment eksportera
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image as PILImage


def _wstaw_logo(self, ws, sciezka_logo: str, *, wysokosc_px: int = 44,
                komorka: str = "A1") -> tuple[int, int]:
    """Wstawia logo, ZACHOWUJAC PROPORCJE. Zwraca (szerokosc, wysokosc)."""
    with PILImage.open(sciezka_logo) as obraz:
        oryginal_w, oryginal_h = obraz.size      # (szerokosc, wysokosc)

    skala = wysokosc_px / oryginal_h
    szerokosc_px = int(oryginal_w * skala)

    obraz = XLImage(sciezka_logo)
    obraz.width = szerokosc_px
    obraz.height = wysokosc_px
    ws.add_image(obraz, komorka)
    return szerokosc_px, wysokosc_px


def _tabela_regiony(self, ws, *, wiersze: int, kolumny: int) -> None:
    ostatni = wiersze + 1
    ostatnia_litera = get_column_letter(kolumny)
    tabela = Table(displayName="TabelaRegiony", ref=f"A1:{ostatnia_litera}{ostatni}")
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True,
        showColumnStripes=False,
        showFirstColumn=False,
        showLastColumn=False,
    )
    ws.add_table(tabela)
```

**Trzy pułapki, o których trzeba wiedzieć.**

1. **Jednostki obrazu.** `obraz.width = 180` **nie znaczy** 180 pikseli CSS. Obraz w openpyxl skaluje się w oparciu o DPI metadanych obrazu (zwykle 96), więc rozmiar w komórkach Excela nie jest liniowy. Jedyna niezawodna metoda: **liczyć proporcje** (`skala = wysokość_docelowa / wysokość_oryginalna`) i ustawić oba wymiary z tej samej skali. Nigdy nie ustawiaj tylko jednego wymiaru — obraz się rozciągnie.
2. **Tabela nie może się nakładać** na inną tabelę, nie może zawierać scalonych komórek w zakresie i jej `ref` musi zawierać wiersz nagłówkowy. Błąd `ValueError: Table ... is already in use` oznacza, że masz dwie tabele o tym samym `displayName` — nazwy są **unikalne w skoroszycie**, nie w arkuszu.
3. **Logo przesuwa wiersze danych.** Jeżeli logo jest w `A1`, to nagłówki tabeli muszą być **niżej** (np. w wierszu 4). To wymaga przesunięcia zakresów, indeksów i ankora wykresu — i to jest powód, dla którego **logo wstawiamy na samym początku budowy arkusza**, a nie na końcu. Dodanie logo do gotowego arkusza „na wierzchu" zawsze kończy się nadpisaniem danych.

### M8 — ochrona, walidacja danych, właściwości

Cel: `F9`, `F10`, `F12`.

```python
# fragment eksportera - ochrona + walidacja
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation


def _ochrona(self, ws, haslo: str, *, komorki_odblokowane: list[str]) -> None:
    """Chroni arkusz. Domyslnie WSZYSTKO zablokowane, potem odblokowujemy
    wskazane komorki (formularz). Odblokowanie MUSI byc PRZED wlaczeniem
    ochrony - inaczej Excel to zignoruje."""
    for adres in komorki_odblokowane:
        ws[adres].protection = Protection(locked=False)

    ws.protection.sheet = True
    ws.protection.password = haslo
    ws.protection.enable()


def _walidacja_regionu(self, ws, *, zakres: str, nazwa_arkusza_regions: str,
                       liczba_regionow: int) -> None:
    """Lista rozwijana z zakresu innego arkusza - nie z literalnej listy!
    Powod (modul 17): literowa lista ma limit 255 znakow."""
    dv = DataValidation(
        type="list",
        formula1=f"={nazwa_arkusza_regions}!$A$2:$A${liczba_regionow + 1}",
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Nieznany region",
        error="Wybierz region z listy rozwijanej.",
        promptTitle="Region",
        prompt="Wybierz region sprzedazy z listy.",
    )
    ws.add_data_validation(dv)
    dv.add(zakres)


def _wlasciwosci(self, wb, raport) -> None:
    wb.properties.title = raport.tytul
    wb.properties.creator = "generator raportow"
    wb.properties.lastModifiedBy = "generator raportow"
    wb.properties.category = "Sprzedaz"
    wb.properties.description = (
        f"schemat={raport.schemat}; generator={raport.wersja_generatora}; "
        f"wsad={raport.hash_wsadu}; okres={raport.okres}"
    )
    wb.properties.keywords = f"schemat:{raport.schemat};okres:{raport.okres}"
    wb.properties.version = raport.wersja_generatora
```

**Kolejność w `_ochrona` jest krytyczna.** `Protection(locked=False)` musi być ustawione **przed** `ws.protection.enable()`. Jeżeli zrobisz to po, Excel zignoruje odblokowanie — bo już zapisałeś strukturę ochrony. To jest dokładnie ten rodzaj błędu, który „działa" (plik ma ochronę), ale nie robi tego, po co go zrobiłeś (nie można nic wpisać).

**Hasło w OOXML to nie szyfrowanie.** `protection.password` ustawia **hash** na ochronie arkusza — to odstrasza od przypadkowych zmian, nie chroni przed determinowanym atakiem. Moduł 17: **kłódka na szufladzie, nie sejf**. I jeszcze jedno: **openpyxl nie umie zapisać zaszyfrowanego pliku** — nie ma obsługi szyfrowania OOXML. Jeżeli dane są poufne, to nie jest miejsce na hasło; to jest miejsce na szyfrowanie pliku na poziomie transportu lub systemu.

### M9 — zapis transakcyjny, audyt, utrata funkcji

Cel: `N3`, `N9`, `N10`. Zapis przez `JednostkaPracy` (moduł 26) plus skan utraty funkcji.

```python
# src/raporty/infrastructure/excel/repository.py (fragment kluczowy)
@dataclass
class JednostkaPracy:
    _komendy: list[tuple[str, Callable]] = field(default_factory=list)

    def dodaj(self, opis: str, operacja: Callable) -> "JednostkaPracy":
        self._komendy.append((opis, operacja))
        return self

    def zatwierdz(self, *, wb, cel: Path, wersja: str) -> Path:
        cel = Path(cel)
        cel.parent.mkdir(parents=True, exist_ok=True)

        # 1. Operacje w PAMIECI. Zero dotykania dysku - pelny rollback przy bledzie.
        for opis, operacja in self._komendy:
            try:
                operacja(wb)
            except Exception as blad:
                raise BladZapisu(
                    f"Operacja {opis!r} nie udala sie w pamieci. "
                    f"Plik docelowy NIE zmieniony."
                ) from blad

        # 2. Tmp W TYM SAMYM katalogu (atomowosc os.replace).
        fd, tmp_name = tempfile.mkstemp(
            prefix=f".{cel.stem}.", suffix=".tmp.xlsx", dir=cel.parent
        )
        os.close(fd)
        tmp = Path(tmp_name)
        try:
            wb.save(tmp)

            # 3. Weryfikacja struktury (modul 18): otworz i policz.
            sprawdzony = load_workbook(tmp, read_only=True)
            try:
                if len(sprawdzony.sheetnames) != len(wb.sheetnames):
                    raise BladZapisu("Weryfikacja tmp: niezgodna liczba arkuszy")
            finally:
                sprawdzony.close()

            # 4. Skontroluj utrate funkcji PRZED podmiana.
            ostrzezenia = przeskanuj_przed_zapisem(tmp)
            if ostrzezenia:
                log.warning("utrata funkcji: %s", "; ".join(ostrzezenia))

            # 5. Backup PRZED podmiana.
            if cel.exists():
                znacznik = datetime.now().strftime("%Y%m%d-%H%M%S")
                shutil.copy2(cel, cel.with_name(f"{cel.stem}.{znacznik}.bak{cel.suffix}"))

            # 6. COMMIT.
            os.replace(tmp, cel)
            return cel
        except Exception as blad:
            tmp.unlink(missing_ok=True)          # 7. Sprzatanie ZAWSZE.
            if isinstance(blad, BladZapisu):
                raise
            raise BladZapisu(f"Zapis do {cel.name} nie udal sie: {blad}") from blad
```

```python
# src/raporty/infrastructure/excel/utrata.py
NIEZAPISYWALNE = {
    "xl/sparklineGroups": "sparkline'y (rozszerzenie x14)",
    "xl/slicers": "slicery",
    "xl/timelines": "timeline'y",
    "xl/pivotCache": "cache tabel przestawnych",
    "xl/model": "model danych (Power Pivot)",
}
CZESCIOWE = {
    "xl/drawings/": "obrazy i ksztalty",
    "xl/charts/": "wykresy (formatowanie moze byc uproszczone)",
}


def przeskanuj_przed_zapisem(sciezka: Path) -> tuple[str, ...]:
    """Tanie skanowanie ZIP: nazwy czesci, nie ich tresc."""
    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()
    ostrzezenia: list[str] = []
    for prefiks, opis in NIEZAPISYWALNE.items():
        if any(n.startswith(prefiks) for n in nazwy):
            ostrzezenia.append(f"UTRATA: {opis}")
    for prefiks, opis in CZESCIOWE.items():
        if any(n.startswith(prefiks) for n in nazwy):
            ostrzezenia.append(f"CZESCIOWO: {opis}")
    return tuple(ostrzezenia)
```

**Test atomowości — najważniejszy test w projekcie:**

```python
def test_wyjatek_w_polowie_nie_niszczy_pliku(tmp_path):
    cel = tmp_path / "raport.xlsx"
    wb = Workbook(); wb.active["A1"] = "wersja 1"
    JednostkaPracy().zatwierdz(wb=wb, cel=cel, wersja="1")
    hash_przed = _hash(cel)

    wb2 = Workbook(); wb2.active["A1"] = "wersja 2 - nie zapisze sie"
    jednostka = JednostkaPracy().dodaj(
        "psujaca operacja",
        lambda _: (_ for _ in ()).throw(RuntimeError("bum")),
    )
    with pytest.raises(BladZapisu, match="NIE zmieniony"):
        jednostka.zatwierdz(wb=wb2, cel=cel, wersja="2")

    assert _hash(cel) == hash_przed                 # plik nietkniety
    assert not list(cel.parent.glob(".*.tmp.xlsx")) # brak smieci
```

**Ten test dowodzi trzech rzeczy naraz:** brak nadpisania przy błędzie, brak pliku tymczasowego, brak „cichego" zachowania. Bez niego `repository` jest **nazwą**, a nie **gwarancją**.

### M10 — 200 000 wierszy, CI, determinizm

Cel: `N1`, `N2`, `N4`, `N6`, `N7`. Tu wchodzi decyzja architektoniczna z sekcji 1.4.

```python
# src/raporty/infrastructure/excel/eksporter_danych.py
"""Strumieniowy eksport surowych danych. Tryb write_only (modul 19).

Decyzja architektoniczna: DUZE dane ida do OSOBNEGO pliku, bo w trybie
write_only nie da sie dodac tabeli ani wykresu do tego samego arkusza.
Dwa samochody na jedna trase (sekcja 1.4).
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell

from .styl import FORMATY


def zapisz_dane_strumieniowo(
    wiersze,
    cel: Path,
    *,
    naglowki: tuple[str, ...],
    formaty: tuple[str, ...],
) -> int:
    """Zapisuje 200k+ wierszy przy STALEJ pamieci. Zwraca liczbe wierszy."""
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Dane")

    # Naglowek: w write_only tworzymy komorki przez WriteOnlyCell.
    komorki_naglowka = []
    for naglowek in naglowki:
        cell = WriteOnlyCell(ws, value=naglowek)
        cell.font = _FONT_NAGLOWKA       # utworzony RAZ, poza petla!
        cell.fill = _FILL_NAGLOWKA       # utworzony RAZ!
        komorki_naglowka.append(cell)
    ws.append(komorki_naglowka)

    licznik = 0
    for dane_wiersza in wiersze:         # generator - nie lista!
        ws.append(list(dane_wiersza))
        licznik += 1

    wb.save(cel)
    return licznik
```

**Dlaczego nie lista.** `wiersze` to **generator** (`yield`), nie `list`. Różnica jest fundamentalna: lista 200 000 tupli to kilkaset megabajtów w pamięci, a generator to **jedna** tuplesz naraz. W module 19 nazwaliśmy to „czytaniem książki strona po stronie" — tu dokładnie ten mechanizm.

**Dlaczego `_FONT_NAGLOWKA` poza pętlą.** W trybie `write_only` każda komórka jest pisana raz i zapominana, ale **obiekt stylu** nadal jest tworzony w Pythonie. Utworzenie `Font(bold=True)` w pętli 200 000 razy to 200 000 alokacji, których Excel i tak nie zapisze wielokrotnie (deduplikuje `styles.xml`). To jest marnowanie zarówno czasu, jak i pamięci.

**Polityka nazw — wymaganie zapomniane w specyfikacji:**

```python
def zbuduj_nazwe_wyjscia(katalog: Path, nazwa: str, *,
                         polityka: str = "wersjonuj",
                         znacznik_czasu: str | None = None) -> Path:
    """Trzy jawne polityki (modul 26, pulapka 5). Zadnego 'raport(1).xlsx'!"""
    if polityka == "nadpisz":
        return katalog / f"{nazwa}.xlsx"

    if polityka == "data":
        dzisiaj = znacznik_czasu or datetime.now().strftime("%Y%m%d")
        return katalog / f"{nazwa}_{dzisiaj}.xlsx"

    if polityka == "wersjonuj":
        numer = 1
        while True:
            kandydat = katalog / f"{nazwa}_v{numer}.xlsx"
            if not kandydat.exists():
                return kandydat
            numer += 1

    raise ValueError(f"Nieznana polityka nazewnictwa: {polityka!r}")
```

**Test determinizmu i wydajności:**

```python
# tests/infrastructure/excel/test_wydajnosc.py
import pytest

@pytest.mark.slow
def test_200k_wierszy_w_budzecie(tmp_path):
    """N1 + N2. Ten test trwa ~30-60 s, dlatego osobny marker."""
    import tracemalloc, time

    tracemalloc.start()
    start = time.perf_counter()

    wiersze = _generuj_200k()               # generator
    liczba = zapisz_dane_strumieniowo(
        wiersze, tmp_path / "dane.xlsx",
        naglowki=("Data", "Region", "Klient", "Produkt", "Ilosc",
                  "Sprzedaz netto", "Marza %", "Dynamika %"),
        formaty=("date", "text", "text", "text", "int",
                 "currency", "percent", "dynamika"),
    )

    czas = time.perf_counter() - start
    _, szczyt = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    assert liczba == 200_000
    assert czas < 60.0, f"przekroczono budzet: {czas:.1f}s"
    assert szczyt < 500 * 1024 * 1024, f"szczyt pamieci: {szczyt/1e6:.0f} MB"


def test_dwa_uruchomienia_ten_sam_hash(tmp_path, zrodlo_stale, eksporter_xlsx):
    """N6: deterministycznosc."""
    raport = _zbuduj_raport(zrodlo_stale)
    a = eksporter_xlsx.eksportuj(raport, str(tmp_path / "a"))
    b = eksporter_xlsx.eksportuj(raport, str(tmp_path / "b"))
    assert a.hash_wyniku == b.hash_wyniku
```

I plik CI (`.github/workflows/ci.yml`), w którym **dwa joby**:

```yaml
jobs:
  testy-bez-openpyxl:
    # N8: domena i aplikacja MUSZA dzialac bez openpyxl
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.12" }
      - run: pip install pytest pyyaml
      - run: pytest tests/domain tests/application -v

  testy-z-openpyxl:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.12" }
      - run: pip install openpyxl==3.1.5 pillow lxml pytest pyyaml pytest-cov
      - run: pytest tests/ -v --cov=src/raporty/domain --cov-fail-under=80
      - run: pytest tests/ -v -m slow          # osobno, z budzetem czasu
```

**Dwa joby to nie duplikacja, a kontrakt.** Pierwszy job **nie ma** openpyxl. Jeżeli przechodzi — domena jest naprawdę oddzielona. Jeżeli nie — do domeny coś przeciekło, i CI to złapie **automatycznie**, przy każdym commicie. To jest realizacja testu z sekcji 2.4, ale **egzekwowana maszynowo**, a nie przez pamięć programisty.

## 4. Anatomia API

### Część A — elementy własne projektu

| Element | Plik | Co robi | Sygnatura / parametry | Uwagi |
|---|---|---|---|---|
| `Zamowienie` | `domain/modele.py` | jedna pozycja sprzedaży, kwoty w groszach | `dataclass(frozen=True)` | `float` pojawia się dopiero w adapterze |
| `AgregatRegionu` | `domain/modele.py` | podsumowanie po regionie | `netto_pln: Decimal`, `marza_proc: Decimal` | buduje arkusz `Regiony` |
| `ModelRaportu` | `domain/modele.py` | kompletny opis raportu | `kpi`, `regiony`, `dynamika`, `kolumny_*`, `metodyka` | **bez kolorów i arkuszy** |
| `przyjmij_tekst_zewnetrzny` | `domain/reguly.py` | sanitacja wejścia | `(wartosc, *, pole) -> str` | odrzuca, nie modyfikuje (F11) |
| `marza_zbiorcza` | `domain/reguly.py` | marża jako iloraz sum | `(list[Zamowienie]) -> Decimal` | **nie** średnia marz |
| `dynamika_proc` | `domain/reguly.py` | dynamika okres/okres | `(biezacy, poprzedni) -> Decimal` | zero poprzednie → 0 |
| `REGULY` | `domain/paleta.py` | **zamknięta paleta** | `dict[str, (Callable, tuple[str, ...])]` | konfiguracja wybiera **nazwę** |
| `EksporterRaportu` | `domain/porty.py` | port wyjściowy | `eksportuj(raport, cel) -> WynikEksportu` | zero typów openpyxl |
| `WynikEksportu` | `domain/porty.py` | DTO wyniku | `format, bajtow, arkuszy, wierszy, sciezki, hash_wyniku, ostrzezenia` | `sciezki` jako `tuple[str, ...]` — dwa pliki! |
| `BazowyRaport` | `application/raport_bazowy.py` | szkielet Template Method | `generuj()`, haczyki `_*` | jeden szkielet, dwa raporty |
| `RaportSprzedazy` | `application/raport_sprzedazy.py` | raport produkcyjny | dziedziczy, nadpisuje haczyki | pełny zestaw arkuszy |
| `RaportMinimal` | `application/raport_minimal.py` | raport minimalistyczny | dziedziczy, nadpisuje haczyki | jeden arkusz, bez wykresów |
| `StosKomend` | `application/komendy.py` | operacje edycyjne | `wykonaj(opis, callable)`, `historia()` | log audytowy, testowalność |
| `RejestrStylow` | `infrastructure/excel/styl.py` | nazwane style | `dodaj_do(wb)` | Flyweight: style tworzone raz |
| `zapisz_dane_strumieniowo` | `infrastructure/excel/eksporter_danych.py` | 200k wierszy | `(generator, cel, *, naglowki, formaty) -> int` | `write_only` + `WriteOnlyCell` |
| `dodaj_wykres_sprzedaz_marza` | `infrastructure/excel/wykresy.py` | wykres dwuosiowy | `(ws, *, wiersze, anchor, kolumny...)` | `axId=200`, `crosses="max"` |
| `JednostkaPracy` | `infrastructure/excel/repository.py` | transakcja zapisu | `dodaj(opis, op)`, `zatwierdz(wb, cel, wersja)` | pamięć → tmp → weryfikacja → backup → `replace` |
| `przeskanuj_przed_zapisem` | `infrastructure/excel/utrata.py` | wykrywanie utraty | `(Path) -> tuple[str, ...]` | `zipfile.namelist()`, tanie |
| `zbuduj_nazwe_wyjscia` | `infrastructure/excel/repository.py` | polityka nazw | `polityka: nadpisz\|data\|wersjonuj` | koniec z `raport(1).xlsx` |
| `dziennik_zapisz` | `infrastructure/audyt/dziennik.py` | dziennik JSONL | `(wpis: dict) -> None` | append-only, odporne na awarię |
| `wczytaj_spec` | `infrastructure/konfig/loader.py` | walidacja YAML | `(Path) -> RaportSpec` | błędy z listą dostępnych wartości |

### Część B — API `openpyxl`, którego używa ten projekt

Wszystko poniżej pojawia się **wyłącznie** w `src/raporty/infrastructure/excel/`.

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | nowy skoroszyt w pamięci | — | ma domyślny arkusz `Sheet` — usuń albo zmień nazwę |
| `Workbook(write_only=True)` | skoroszyt strumieniowy | `write_only=True` | brak dostępu losowego; `ws.append()` tylko w przód |
| `wb.create_sheet("Nazwa")` | nowy arkusz | `title`, `index` | kolejność tworzenia = kolejność w pliku |
| `wb.remove(wb.active)` | usuwa domyślny `Sheet` | — | bez tego lista arkuszy jest dłuższa niż oczekiwana |
| `wb.active = 0` | wybiera aktywny arkusz | indeks | użytkownik widzi `Podsumowanie`, nie `_meta` |
| `wb.add_named_style(style)` | rejestruje nazwany styl | `NamedStyle` | **przed** pierwszym `cell.style = "Nazwa"` |
| `wb.properties.title/creator/description/keywords/version` | metadane dokumentu | `str` | `version` = wersja generatora (F9) |
| `wb.save(sciezka)` | zapis na dysk | ścieżka lub obiekt plikopodobny | **COMMIT** — po tym cofanie stosu zmienia tylko pamięć |
| `load_workbook(tmp, read_only=True)` | weryfikacja po zapisie | `read_only` | **`close()` obowiązkowe** (moduł 19) |
| `ws.title`, `ws.sheet_state` | nazwa, widoczność | `"hidden"` / `"visible"` | `_meta` jest `hidden` |
| `ws.sheet_properties.tabColor` | kolor zakładki | `"1F4E78"` | tanie, poprawia UX |
| `ws.sheet_view.showGridLines = False` | odcięcie siatki | `bool` | tabela ma własne obramowanie |
| `ws.freeze_panes = "A2"` | zamrożenie nagłówka | adres | reguła pozycji zależy od adresu (moduł 11) |
| `ws.append([...])` | dopisanie wiersza | iterowalne | `append("abc")` wpisze 3 komórki — uwaga! |
| `ws.cell(row, column, value)` | dostęp do komórki | 1-based! | indeksowanie od 1 (moduł 03) |
| `cell.style = "Nazwa"` | nadanie nazwanego stylu | `str` | szybkie, nie tworzy obiektu na komórkę |
| `cell.number_format` | format liczby | kod z `FORMATY` | nie walidowany przez openpyxl (moduł 10) |
| `cell.protection = Protection(locked=False)` | odblokowanie komórki | — | **przed** `protection.enable()` |
| `ws.protection.sheet = True` + `enable()` + `password` | ochrona arkusza | — | to **nie** szyfrowanie (moduł 17) |
| `ws.column_dimensions[litera].width` | szerokość kolumny | `float` | litera z `get_column_letter()` |
| `ws.row_dimensions[n].height` | wysokość wiersza | `float` | nagłówek 22 pkt |
| `ws.auto_filter.ref = "A1:H100"` | autofiltr na zakresie | `str` | alternatywa dla tabeli |
| `ws.print_title_rows = "1:1"` | powtarzanie nagłówka | **format `"1:1"`** | `"1"` nie zadziała (moduł 11) |
| `ws.sheet_properties.pageSetUpPr.fitToPage = True` | włącza dopasowanie | — | **bez tego `fitToWidth` ignorowane** |
| `ws.page_setup.fitToWidth/fitToHeight/orientation/paperSize` | ustawienia druku | `1`, `0`, `"landscape"`, `ws.PAPERSIZE_A4` | trzy ustawienia razem |
| `ws.oddHeader.center.text = "&P / &N"` | nagłówek wydruku | kod `&` | przepis Excela, nie zwykły tekst |
| `Table(displayName=, ref=)` | tabela Excela | nazwa unikalna, `ref` z nagłówkiem | nie nakładaj tabel, nie w scalonym zakresie |
| `TableStyleInfo(name=, showRowStripes=)` | styl tabeli | `"TableStyleMedium9"` | dodaj **przed** `add_table` |
| `ws.add_table(tabela)` | dodanie tabeli | — | `displayName` unikalne w skoroszycie |
| `ws.add_data_validation(dv)` + `dv.add(zakres)` | lista rozwijana | `type="list"`, `formula1` | lista z zakresu: `"=Regiony!$A$2:$A$20"` |
| `ws.conditional_formatting.add(zakres, reguła)` | reguła warunkowa | zakres **osobno** | zakres i formuła muszą być spójne |
| `DataBarRule(start_type, start_value, end_type, end_value, color, showValue)` | pasek danych | wartości ujemne: `start_value=-0.5` | „wykres z wypełnienia" (moduł 14) |
| `CellIsRule(operator, formula, fill, font)` | reguła „komórka vs wartość" | `operator="lessThan"` | `formula` jako lista |
| `FormulaRule(formula, fill, font)` | reguła z formuły | **bez `=`**, adresy `$D2` | podświetlenie całego wiersza |
| `ColorScaleRule(start_type, start_color, mid_*, end_*)` | skala kolorów | typy: `min`, `max`, `percent`, `num` | 3 kolory = 2 progi |
| `BarChart()`, `LineChart()` | obiekty wykresów | `type="col"`, `style=10` | `slupki += linia` skleja serie |
| `Reference(ws, min_col, min_row, max_col, max_row)` | wskazanie danych | 1-based | poza zakresem = pusty wykres |
| `chart.add_data(dane, titles_from_data=True)` | dodanie serii | — | `titles_from_data=True` bierze nagłówek |
| `chart.set_categories(ref)` | kategorie osi X | — | brak tego = oś X „1, 2, 3" |
| `linia.y_axis.axId = 200` | **druga oś** | dowolny unikalny int | bez tego jedna wspólna oś |
| `slupki.y_axis.crosses = "max"` | przesunięcie osi | `"max"` | „prawa" oś na krawędzi |
| `series.graphicalProperties = GraphicalProperties(solidFill=...)` | kolor serii | `solidFill`, `ln` | **nie ma** `series.color`! |
| `chart.y_axis.numFmt`, `majorGridlines = ChartLines()` | format osi, siatka | — | `ChartLines()` = widoczna; `None` = brak |
| `ws.add_chart(chart, "F2")` | zakotwiczenie wykresu | adres lewej górnej krawędzi | wykres „pływa" nad siatką |
| `XLImage(sciezka)` + `ws.add_image(img, "A1")` | obraz | `img.width`, `img.height` | wymaga `pillow`; **licz proporcje** |
| `WriteOnlyCell(ws, value=...)` | komórka w `write_only` | `value` | nadaj style **przed** `append` |
| `get_column_letter(n)` | numer → litera | 1-based | `4` → `"D"` |
| `zipfile.ZipFile(x).namelist()` | inwentarz części | — | tanie; skan utraty funkcji |
| `os.replace(tmp, cel)` | atomowa podmiana | — | **tylko ten sam wolumen!** |
| `tempfile.mkstemp(dir=cel.parent)` | tmp w tym samym katalogu | `dir=` **kluczowe** | bez `dir=` tracisz atomowość |
| `shutil.copy2(oryginał, backup)` | backup z metadanymi | — | **`copy2`**, nie `copy` |

## 5. Ćwiczenia

To ćwiczenie **jest całym projektem**. Nie ma osobnych „zadań" — jest specyfikacja powyżej i dziesięć kamieni milowych. Ale żebyś nie tylko zbudował jedną wersję, dostajesz trzy warianty. Porównanie ich jest **ważniejsze** niż samo napisanie kodu.

### 🟢 Rozgrzewka — „raport minimalistyczny" (poziom 1 abstrakcji)

**Cel:** poczuć, jakim kosztem kupuje się elastyczność.

Napisz plik `raport_minimal.py` — **jeden plik, bez warstw, bez portów, bez YAML-a**. Wymagania:

- czyta `data/orders.csv` (`csv.DictReader`, `encoding="utf-8-sig"`),
- liczy sumę netto i marżę per region — **w Pythonie**, na `Decimal`,
- tworzy skoroszyt z **jednym** arkuszem `Regiony`,
- stosuje format walutowy i procentowy,
- dodaje **jedną** regułę warunkową: czerwone tło dla marży < 5 %,
- zapisuje do `output/raport_minimal.xlsx`.

**Ograniczenia:** maksymalnie **120 linii** (licząc komentarze i puste linie). Bez importu z `src/raporty/`. Bez testów jednostkowych — tylko **jeden** test, który sprawdza, że plik powstał i ma 1 arkusz.

**Uruchom i zmierz:**

```bash
wc -l raport_minimal.py
python -c "import time; t=time.perf_counter(); import raport_minimal; print(time.perf_counter()-t)"
```

**Zapisz w `output/27_wariant_minimal.md`:**

1. Ile linii ma plik?
2. Ile czasu zajęło napisanie (zmierz zegarkiem)?
3. **Trzy konkretne zmiany**, które Cię zablokują przy modyfikacji tego pliku. Przykład: „zmiana formatu waluty wymaga znalezienia kodu w środku funkcji `main()`". Podaj **trzy** i napisz, jak długo byś szukał.
4. Czy ten plik **przechodzi** Twoje wymagania `N6` (determinizm) i `N3` (atomowość)? Sprawdź naprawdę: uruchom dwa razy i porównaj. Spróbuj zasymulować błąd w połowie — co się dzieje z plikiem?

### 🟡 Warsztat — „raport produkcyjny" (poziom 3 abstrakcji)

**Cel:** zbudować pełną wersję i **udowodnić**, że spełnia wymagania.

**Krok 1.** Przejdź wszystkie dziesięć kamieni milowych `M1`–`M10`. Po każdym kroku **uruchom testy** i **otwórz plik w Excelu**. Bez wyjątków.

**Krok 2.** Dla każdego wymagania funkcjonalnego `F1`–`F12` napisz test **albo** uzasadnij, dlaczego jest niezbadane automatem i wymaga „weryfikacji na oko". Wypełnij tabelę:

| Nr | Test (ścieżka + nazwa) | Albo: dlaczego „na oko" |
|---|---|---|
| F1 | `tests/infrastructure/excel/test_arkusze.py::test_arkusze_f1` | — |
| F2 | — | Rozmiar logo w komórkach Excela zależy od DPI; test sprawdziłby obecność `xl/media/image1.png`, ale nie proporcje i nie pozycję |
| ... | ... | ... |

**Krok 3.** Zbuduj zestaw danych testowych: `python scripts/gen_dane.py --wierszy 200000 > data/orders_big.csv`. Uruchom generator i zmierz `N1`, `N2`.

**Krok 4.** Zrób próbę sił: **zaburz plik CSV**. Dodaj wiersz, w którym klient nazywa się `=HYPERLINK("http://zly.example","kliknij")`. Dodaj wiersz z pustym regionem. Dodaj kolumnę zamiast oczekiwanej wartości. Za każdym razem program **musi** zareagować jawnie: albo czytelnym błędem, albo komunikatem w `ostrzezeniach` — **nigdy cichym pominięciem**.

**Krok 5.** Porównaj oba warianty w `output/27_porownanie.md`:

| Kryterium | Minimalny (1 plik) | Produkcyjny (warstwy) |
|---|---|---|
| Linii kodu | | |
| Czas napisania | | |
| Zmiana formatu waluty | sekundy? minuty? | jedna linia w `FORMATY` |
| Dodanie nowego arkusza | | |
| Testy domeny bez openpyxl | | |
| Deterministyczność (`N6`) | | |
| Atomowość (`N3`) | | |
| Drugi format wyjścia (CSV) | | jeden adapter |
| Czytelność po 6 miesiącach | | |

**Krok 6 — najważniejszy.** Napisz **trzy zdania** odpowiedzi na pytanie: **„Kiedy bym wybrał wariant minimalistyczny, a kiedy produkcyjny?"** Odpowiedź wymagająca „zawsze produkcyjny" jest **zła** — jest nieprawdziwa i będziesz się jej wstydził, gdy przyjdzie jednorazowe zadanie na 30 minut.

### 🔴 Wyzwanie — trzy rozszerzenia i audyt kodu

**Zadanie A — drugi adapter: CSV jako format wyjścia.**

Dodaj `EksporterCsv` implementujący ten sam port. Wyzwanie: **port ma `arkuszy`**. CSV nie ma arkuszy. Trzy możliwe rozwiązania:

1. zwrócić `arkuszy=1`,
2. wygenerować **wiele** plików CSV (po jednym na arkusz) i zwrócić `len(sciezki)`,
3. **zmienić port** tak, żeby nie obiecywał arkuszy.

Wybierz jedno i **uzasadnij w komentarzu**, dlaczego pozostałe dwa odrzuciłeś. Punkt nie jest za wybór — punkt jest za **uzasadnienie**. Odpowiedź „bo tak było łatwiej" też jest uzasadnieniem, o ile powiesz, co za to płacisz.

**Zadanie B — test kontraktu sparametryzowany.**

Napisz test, który uruchamia **ten sam zestaw asercji** na trzech adapterach (`EksporterXlsxOpenpyxl`, `EksporterCsv`, `EksporterFake`):

```python
@pytest.mark.parametrize("nazwa", ["xlsx", "csv", "fake"])
def test_kontrakt_portu(nazwa, model_raportu, tmp_path, eksportery):
    eksporter = eksportery[nazwa]()
    wynik = eksporter.eksportuj(model_raportu, str(tmp_path / "raport"))
    assert wynik.wierszy == len(model_raportu.regiony)
    assert wynik.bajtow >= 0
    assert wynik.format == nazwa
```

**Cel jest epistemologiczny**, nie techniczny: ten test **obnaży**, które asercje dotyczyły `xlsx`, a nie kontraktu. Jeżeli któraś asercja padnie dla CSV — to znaczy, że port wymaga czegoś, co jest prawdą tylko dla Excela, i to trzeba poprawić **w porcie**, nie w teście.

**Zadanie C — mapa „problem → moduł kursu".**

W `output/27_mapa.md` zbuduj tabelę. W lewej kolumnie **konkretne problemy z tego projektu**, w prawej moduł, do którego wracasz:

| Problem z projektu | Moduł | Co dokładnie przeczytać |
|---|---|---|
| Kwoty się rozjeżdżają o grosz | 05, 10 | `Decimal` vs `float`, reprezentacja liczby |
| Pasek danych źle wygląda dla spadków | 14 | `DataBarRule`, wartości ujemne |
| Trzeba dodać drugą linię do wykresu | 15 | `axId`, `crosses`, `+= chart` |
| Klient przysyła plik z sparkline'ami | 18 | inwentarz utraty, `scan-before-save` |
| 500k wierszy i brakuje pamięci | 19 | `write_only`, `lxml`, generator |
| CI w Dockerze bez LibreOffice | 20 | testy przez `BytesIO` |
| Chcę pokazać kwoty z symbolem `€` | 10 | kody formatów, locale |
| Użytkownik wpisuje `=` w nazwę klienta | 21 | formula injection, sanitacja |
| Zmieniła się reguła marży | 25, 27 | Specification, paleta reguł |
| Trzeba wysłać raport e-mailem | 26, 27 | adapter, kolejka, **sekcja 2.11** |
| Reguła warunkowa nie działa — kolumna bez zmian | 13 | `sqref`, adresy `$D2` vs `$D$2` |
| W tabeli nie mogę dać tego samego `displayName` | 12 | `Table`, nazwy unikalne w skoroszycie |

**Minimum 15 wierszy.** Ta tabela nie jest dla nikogo innego — jest **dla Ciebie**, żebyś za pół roku nie czytał kursu od początku, tylko otwierał **jeden** moduł.

<details>
<summary><strong>Szkic rozwiązania — struktura, sygnatury, decyzje i pełny kod trzech najważniejszych plików</strong></summary>

Poniżej nie ma całego kodu (projekt ma kilkanaście plików i wiele tysięcy linii). Jest **struktura, sygnatury, uzasadnienia decyzji** i **pełny kod trzech plików**, które decydują o wszystkim: porty, `BazowyRaport` z dwoma raportami konkretnymi, adapter openpyxl.

### Struktura i sygnatury

```python
# --- domain/porty.py ---
class ZrodloDanych(Protocol):
    def pobierz(self, okres: str) -> tuple[list[Zamowienie], str]: ...

class EksporterRaportu(Protocol):
    @property
    def format(self) -> str: ...
    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu: ...

class MagazynAudytu(Protocol):
    def zapisz_zdarzenie(self, wpis: dict) -> None: ...

# --- application/raport_bazowy.py ---
class BazowyRaport(ABC):
    def __init__(self, *, zrodlo: ZrodloDanych, eksporter: EksporterRaportu,
                 spec: RaportSpec, zegar: Callable[[], datetime] | None = None) -> None: ...

    # SZKIELET (final) - wolno go zmienic tylko tutaj:
    def generuj(self, *, okres: str, cel: str) -> WynikUzycia: ...

    # HACZYKI (abstrakcyjne) - kazdy raport wypelnia po swojemu:
    @abstractmethod
    def _zbierz(self, zamowienia: list[Zamowienie], okres: str) -> ModelRaportu: ...
    @abstractmethod
    def _arkusze(self) -> tuple[str, ...]: ...
    @abstractmethod
    def _metodyka(self) -> tuple[str, ...]: ...
    def _zapisuj_dane_strumieniowo(self) -> bool:
        return False                       # haczyk z wartoscia domyslna

# --- application/raport_sprzedazy.py ---
class RaportSprzedazy(BazowyRaport): ...
class RaportMinimal(BazowyRaport): ...

# --- infrastructure/excel/eksporter.py ---
class EksporterXlsxOpenpyxl:
    def __init__(self, *, styl: RejestrStylow | None = None,
                 haslo: str | None = None) -> None: ...
    @property
    def format(self) -> str: ...
    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu: ...
```

### Decyzja 1: dlaczego Template Method, a nie dwa osobne raporty

Mamy **dwa raporty** (`RaportSprzedazy` i `RaportMinimal`) i będą następne. Bez wzorca różnią się **trzema** metodami i powtarzają **całą** orkiestrację: pobierz → zbuduj → eksportuj → zaloguj → zapisz audyt. To około 40 linii powtórzonych **na każdy** raport. Template Method (moduł 25) wynosi tę orkiestrację **raz**, do `BazowyRaport.generuj()`, a podklasy wypełniają trzy haczyki. Konsekwencja: zmiana kolejności kroków w orkiestracji (np. dodanie walidacji przed eksportem) wymaga **jednej** zmiany, a nie „zmiany w każdym raporcie i modlitwy, że nic nie pominąłem".

### Pełny kod 1 z 3: `src/raporty/domain/porty.py`

```python
"""src/raporty/domain/porty.py - porty projektu.

Porty sa napisane w jezyku DOMENY. Zadnego Workbook, Worksheet, Cell,
BytesIO, Path z openpyxl. Test:  grep -n "openpyxl" porty.py  -> pusto.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from .modele import ModelRaportu, Zamowienie


@dataclass(frozen=True)
class WynikEksportu:
    """DTO wyniku. Nasz typ, nie biblioteczny."""

    format: str                                # "xlsx" | "csv" | "fake"
    bajtow: int
    arkuszy: int
    wierszy: int
    sciezki: tuple[str, ...]                   # moze byc wiecej niz jeden plik!
    hash_wyniku: str = ""
    ostrzezenia: tuple[str, ...] = ()


class ZrodloDanych(Protocol):
    """PORT WEJSCIOWY. Jedno miejsce, przez ktore wchodza dane do domeny."""

    def pobierz(self, okres: str) -> tuple[list[Zamowienie], str]:
        """Zwraca (lista_zamowien, hash_wsadu).

        Hash wsadu jest OBOWIAZKOWY - trafia do _meta i do logu (N9).
        """
        ...


class EksporterRaportu(Protocol):
    """PORT WYJSCIOWY. To, czego aplikacja potrzebuje od swiata zewnetrznego."""

    @property
    def format(self) -> str: ...

    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu:
        """Zapisuje model. Zwraca metryki.

        Kontrakt:
          - NIE modyfikuje obiektu `raport` (jest frozen),
          - NIE nadpisuje danych wejsciowych,
          - NIE zostawia plikow tymczasowych po bledzie,
          - przy bledzie rzuca wyjatek z jasnym komunikatem (nie cichy return).
        """
        ...


class MagazynAudytu(Protocol):
    """PORT AUDYTU. Dziennik zdarzen poza plikiem xlsx."""

    def zapisz_zdarzenie(self, wpis: dict) -> None: ...


class Zegar(Protocol):
    """PORT CZASU. Wstrzykiwany, zeby raport byl deterministyczny (N6)."""

    def teraz(self) -> "object": ...          # datetime
```

### Decyzja 2: dlaczego `WynikEksportu.sciezki` jest krotką, nie pojedynczą ścieżką

Bo nasz projekt **rzeczywiście** zapisuje dwa pliki (sekcja 1.4). Gdyby port obiecywał **jedną** ścieżkę, to M10 wymusiłby albo (a) rezygnację z drugiego pliku, albo (b) przeciek abstrakcji — aplikacja musiałaby wiedzieć, że plików jest dwa i że drugi powstaje w trybie `write_only`. **Krotka obsługuje oba przypadki**: `len(sciezki) == 1` dla małych raportów, `len(sciezki) == 2` dla dużych. Port nie musi się zmieniać w zależności od rozmiaru danych — a to znaczy, że aplikacja i testy **też** się nie muszą zmieniać.

### Pełny kod 2 z 3: `src/raporty/application/raport_bazowy.py`

```python
"""src/raporty/application/raport_bazowy.py - SZKIELET (Template Method).

Ten plik NIE importuje openpyxl. Orkiestruje przepływ i deleguje.
Zmiana kolejnosci krokow = zmiana w JEDNYM miejscu.
"""

from __future__ import annotations

import logging
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Callable

from src.raporty.domain.modele import ModelRaportu, Zamowienie
from src.raporty.domain.porty import (
    EksporterRaportu,
    MagazynAudytu,
    WynikEksportu,
    ZrodloDanych,
)
from src.raporty.domain.reguly import BladDomeny

log = logging.getLogger("raporty.app")


@dataclass(frozen=True)
class WynikUzycia:
    """DTO aplikacji. Co dostaje CLI/HTTP po wywolaniu generuj()."""

    wynik: WynikEksportu
    czas_s: float
    okres: str
    schemat: str
    hash_wsadu: str
    ostrzezenia: tuple[str, ...]


class BazowyRaport(ABC):
    """SZKIELET. Podklasy wypelniaja haczyki _*."""

    identyfikator: str = "raport_bazowy"
    schemat: str = "v1"

    def __init__(
        self,
        *,
        zrodlo: ZrodloDanych,
        eksporter: EksporterRaportu,
        magazyn: MagazynAudytu | None = None,
        zegar: Callable[[], datetime] | None = None,
    ) -> None:
        self._zrodlo = zrodlo
        self._eksporter = eksporter
        self._magazyn = magazyn
        # WSTRZYKNIETY ZEGAR -> deterministycznosc (N6). Bez tego test
        # porownujacy dwa przebiegi bylby niemozliwy.
        self._teraz = zegar or (lambda: datetime.now(timezone.utc))

    # ==================================================================
    # SZKIELET - final. Zmieniaj TYLKO tutaj, nie w podklasach.
    # ==================================================================
    def generuj(self, *, okres: str, cel: str) -> WynikUzycia:
        start = time.perf_counter()
        log.info("raport start id=%s okres=%s format=%s",
                 self.identyfikator, okres, self._eksporter.format)

        # 1. WEJSCIE - jedna sciezka do domeny, wiec jedno miejsce sanitacji.
        zamowienia, hash_wsadu = self._zrodlo.pobierz(okres)
        if not zamowienia:
            raise BladDomeny(f"Brak zamowien w okresie {okres!r}")

        # 2. DOMENA buduje model. Podklasa decyduje, JAKI model.
        model = self._zbierz(zamowienia, okres, hash_wsadu)
        ostrzezenia: list[str] = list(model.ostrzezenia)

        # 3. WYJSCIE - aplikacja nie wie, ze to openpyxl.
        wynik = self._eksporter.eksportuj(model, cel)
        ostrzezenia.extend(wynik.ostrzezenia)

        czas = time.perf_counter() - start

        # 4. OBSERWOWALNOSC - JEDNO zdarzenie z kompletnym kompletem faktow.
        log.info(
            "raport koniec id=%s okres=%s format=%s wierszy=%s arkuszy=%s "
            "plikow=%s bajtow=%s czas_s=%.2f ostrzezen=%s wynik=ok",
            self.identyfikator, okres, wynik.format, wynik.wierszy,
            wynik.arkuszy, len(wynik.sciezki), wynik.bajtow, czas,
            len(ostrzezenia),
        )
        for tresc in ostrzezenia:
            log.warning("ostrzezenie id=%s tresc=%s", self.identyfikator, tresc)

        # 5. AUDYT - poza plikiem xlsx, zeby nie ginal razem z plikiem.
        if self._magazyn is not None:
            self._magazyn.zapisz_zdarzenie({
                "id": self.identyfikator,
                "okres": okres,
                "schemat": model.schemat,
                "generator": model.wersja_generatora,
                "hash_wsadu": hash_wsadu,
                "hash_wyniku": wynik.hash_wyniku,
                "wygenerowano": model.wygenerowano,
                "wierszy": wynik.wierszy,
                "czas_s": round(czas, 3),
                "plikow": len(wynik.sciezki),
                "ostrzezenia": ostrzezenia,
            })

        return WynikUzycia(
            wynik=wynik, czas_s=czas, okres=okres,
            schemat=model.schemat, hash_wsadu=hash_wsadu,
            ostrzezenia=tuple(ostrzezenia),
        )

    # ==================================================================
    # HACZYKI - kazdy raport wypelnia po swojemu.
    # ==================================================================
    @abstractmethod
    def _zbierz(self, zamowienia: list[Zamowienie], okres: str,
                hash_wsadu: str) -> ModelRaportu:
        """Buduje model raportu. Domena liczy, aplikacja orkiestruje."""

    @abstractmethod
    def _arkusze(self) -> tuple[str, ...]:
        """Lista arkuszy - do testu F1 i do _meta."""

    @abstractmethod
    def _metodyka(self) -> tuple[str, ...]:
        """Tekst arkusza 'Metodyka' (F8)."""

    def _zapisuj_dane_strumieniowo(self) -> bool:
        """Haczyk z wartoscia domyslna: czy tworzyc DRUGI plik (write_only)?"""
        return False
```

### Decyzja 3: dlaczego `_zbierz` przyjmuje trzy argumenty, a nie tylko listę

Bo `okres` i `hash_wsadu` **muszą** trafić do modelu (`_meta`, `F9`), a model jest budowany w haczyku. Alternatywa — trzymanie ich w `self` — uczyniłaby obiekt raportu **stanowym** między wywołaniami i zabiłaby determinizm oraz testowalność (dwa równoległe wywołania na tym samym obiekcie mieszałyby dane). Przekazanie jawnych argumentów do czystej funkcji to ta sama zasada, którą stosowaliśmy dla zegara: **stan jest wstrzykiwany, nie trzymany**.

### Pełny kod 3 z 3: `src/raporty/infrastructure/excel/eksporter.py`

```python
"""src/raporty/infrastructure/excel/eksporter.py - ADAPTER openpyxl.

To JEDYNE miejsce, ktore tlumaczy model domeny na komorki, arkusze,
tabele, reguly i wykresy. Wszystko, co wiesz o openpyxl, jest tutaj.
"""

from __future__ import annotations

import hashlib
import logging
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from src.raporty.domain.modele import ModelRaportu
from src.raporty.domain.porty import WynikEksportu
from src.raporty.infrastructure.excel.repository import (
    BladZapisu,
    JednostkaPracy,
    zbuduj_nazwe_wyjscia,
)
from src.raporty.infrastructure.excel.styl import FORMATY, RejestrStylow, zbuduj_rejestr
from src.raporty.infrastructure.excel.wykresy import dodaj_wykres_sprzedaz_marza

log = logging.getLogger("raporty.infra.excel")

# Tlumaczenie format_hint z domeny na styl nazwany skoroszytu.
STYL_DLA_FORMATU = {
    "currency": "RaportWaluta",
    "percent": "RaportProcent",
    "int": "RaportLiczba",
    "date": "RaportData",
    "text": "RaportTekst",
}


class EksporterXlsxOpenpyxl:
    """ADAPTER. Implementuje port EksporterRaportu (strukturalnie)."""

    def __init__(
        self,
        *,
        haslo_ochrony: str | None = None,
        sciezka_logo: str | None = "assets/logo.png",
        polityka_nazw: str = "wersjonuj",
    ) -> None:
        self._haslo = haslo_ochrony
        self._logo = sciezka_logo
        self._polityka = polityka_nazw

    @property
    def format(self) -> str:
        return "xlsx"

    # ==================================================================
    # GLOWNA METODA PORTU
    # ==================================================================
    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu:
        cel_path = Path(cel).with_suffix(".xlsx")
        cel_path.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        wb.remove(wb.active)                    # usuwamy domyslny "Sheet"!

        styl = zbuduj_rejestr()
        styl.dodaj_do(wb)                       # PRZED pierwszym cell.style

        ostrzezenia: list[str] = []

        # --- Arkusze budujemy w kolejnosci prezentacji ---
        self._arkusz_podsumowanie(wb, raport, styl)
        self._arkusz_regiony(wb, raport, styl, ostrzezenia)
        self._arkusz_dynamika(wb, raport, styl)
        if self._logo:
            self._wstaw_logo(wb["Podsumowanie"], self._logo)
        self._arkusz_metodyka(wb, raport, styl)
        self._arkusz_meta(wb, raport, styl, ostrzezenia)

        self._wlasciwosci(wb, raport)
        wb.active = 0                           # uzytkownik widzi Podsumowanie

        # --- Komendy w pamieci (Command, modul 25) ---
        jednostka = JednostkaPracy()
        if self._haslo:
            jednostka.dodaj(
                "ochrona arkuszy wynikowych",
                lambda wb_: self._ochrona_wszystkich(wb_, self._haslo),
            )

        sciezka = zbuduj_nazwe_wyjscia(
            cel_path.parent, cel_path.stem, polityka=self._polityka
        )

        try:
            jednostka.zatwierdz(wb=wb, cel=sciezka, wersja=raport.schemat)
        except BladZapisu:
            raise

        # --- Drugi plik: dane strumieniowo (tylko gdy raport ich wymaga) ---
        wynik = WynikEksportu(
            format="xlsx",
            bajtow=sciezka.stat().st_size,
            arkuszy=len(wb.sheetnames),
            wierszy=len(raport.regiony),
            sciezki=(str(sciezka),),
            hash_wyniku=self._hash(sciezka),
            ostrzezenia=tuple(ostrzezenia),
        )
        return wynik

    # ==================================================================
    # ARKUSZE - kazdy w osobnej metodzie (SRP, modul 26)
    # ==================================================================
    def _arkusz_podsumowanie(self, wb, raport: ModelRaportu, styl):
        ws = wb.create_sheet("Podsumowanie")
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = "1F4E78"

        # Logo wstawia osobna metoda - tu zostawiamy miejsce.
        # Naglowek tekstowy ponizej, zeby logo go nie zakrywalo.
        ws["A3"] = raport.tytul
        ws["A3"].style = "RaportTytul"
        ws["A4"] = f"Okres: {raport.okres}"
        ws["B4"] = f"Wygenerowano: {raport.wygenerowano}"
        ws["A4"].font = Font(italic=True, color="595959")
        ws["B4"].font = Font(italic=True, color="595959")

        wiersz = 6
        for kpi in raport.kpi:
            ws.cell(row=wiersz, column=1, value=kpi.etykieta).style = "RaportKpiEtykieta"
            wartosc = kpi.wartosc
            if kpi.format_hint in ("currency", "percent"):
                wartosc = float(wartosc)      # jedyna konwersja Decimal -> float
            komorka = ws.cell(row=wiersz, column=2, value=wartosc)
            komorka.style = STYL_DLA_FORMATU.get(kpi.format_hint, "RaportTekst")
            komorka.font = Font(bold=True, size=14, color="1F4E78")
            wiersz += 1

        ws["A3"].alignment = None            # reset - tytul wyrownany do lewej
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 26
        ws.row_dimensions[1].height = 50     # miejsce na logo
        return ws

    def _arkusz_regiony(self, wb, raport, styl, ostrzezenia: list[str]):
        ws = wb.create_sheet("Regiony")
        ws.sheet_view.showGridLines = False

        # UWAGA: naglowki w wierszu 4, bo wiersze 1-3 rezerwujemy na logo.
        WIERSZ_NAGLOWKA = 4
        for kolumna, (nazwa, format_hint) in enumerate(raport.kolumny_regiony, start=1):
            ws.cell(row=WIERSZ_NAGLOWKA, column=kolumna, value=nazwa).style = (
                "RaportNaglowekGranat"
            )

        from src.raporty.infrastructure.excel.wykresy import _pole  # lokalny import

        for indeks, agregat in enumerate(raport.regiony, start=WIERSZ_NAGLOWKA + 1):
            for kolumna, (_, format_hint) in enumerate(raport.kolumny_regiony, start=1):
                pole = _pole(agregat, kolumna, raport.kolumny_regiony)
                wartosc = float(pole) if format_hint in ("currency", "percent") else pole
                ws.cell(row=indeks, column=kolumna, value=wartosc).style = (
                    STYL_DLA_FORMATU.get(format_hint, "RaportTekst")
                )

        ostatni = WIERSZ_NAGLOWKA + len(raport.regiony)

        # F3: TABELA z filtrem i stylem.
        from openpyxl.worksheet.table import Table, TableStyleInfo
        litera = get_column_letter(len(raport.kolumny_regiony))
        tabela = Table(
            displayName="TabelaRegiony",
            ref=f"A{WIERSZ_NAGLOWKA}:{litera}{ostatni}",
        )
        tabela.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showRowStripes=True,
            showColumnStripes=False,
            showFirstColumn=False,
            showLastColumn=False,
        )
        ws.add_table(tabela)

        # F6: WYKRES DWUOSIOWY. Marza jest w kolumnie 4 (patrz konfiguracja).
        dodaj_wykres_sprzedaz_marza(
            ws, wiersze=len(raport.regiony),
            anchor=f"G{WIERSZ_NAGLOWKA}",
            kolumna_region=1, kolumna_netto=3, kolumna_marza=4,
        )

        ws.freeze_panes = f"A{WIERSZ_NAGLOWKA + 1}"
        ws.print_title_rows = f"{WIERSZ_NAGLOWKA}:{WIERSZ_NAGLOWKA}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        for kolumna, szerokosc in zip("ABCD", (26, 12, 18, 12), strict=False):
            ws.column_dimensions[kolumna].width = szerokosc
        return ws

    def _arkusz_dynamika(self, wb, raport: ModelRaportu, styl):
        ws = wb.create_sheet("Dynamika")
        ws.sheet_view.showGridLines = False

        naglowki = ("Region", "Poprzedni okres", "Biezacy okres", "Dynamika %")
        for kolumna, naglowek in enumerate(naglowki, start=1):
            ws.cell(row=1, column=kolumna, value=naglowek).style = "RaportNaglowekGranat"

        for indeks, wiersz in enumerate(raport.dynamika, start=2):
            ws.cell(row=indeks, column=1, value=wiersz.region).style = "RaportTekst"
            ws.cell(row=indeks, column=2,
                    value=float(wiersz.netto_poprzedni_pln)).style = "RaportWaluta"
            ws.cell(row=indeks, column=3,
                    value=float(wiersz.netto_biezacy_pln)).style = "RaportWaluta"
            ws.cell(row=indeks, column=4,
                    value=float(wiersz.dynamika_proc)).style = "RaportProcent"

        ostatni = len(raport.dynamika) + 1
        zakres = f"D2:D{ostatni}"

        # F7: PASEK DANYCH. start_value UJEMNY - obsluguje spadki.
        ws.conditional_formatting.add(zakres, DataBarRule(
            start_type="num", start_value=-0.5,
            end_type="num", end_value=0.5,
            color="638EC6", showValue=True,
        ))
        # F5: SPADKI - czerwone tlo i tekst.
        ws.conditional_formatting.add(zakres, CellIsRule(
            operator="lessThan", formula=["0"],
            fill=PatternFill(fill_type="solid",
                             start_color="FCE4E4", end_color="FCE4E4"),
            font=Font(color="9C0006", bold=True),
        ))

        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 24
        for litera in ("B", "C", "D"):
            ws.column_dimensions[litera].width = 18
        return ws

    def _wstaw_logo(self, ws, sciezka: str, *, wysokosc_px: int = 44) -> None:
        """F2. Proporcje liczone, nie zgadywane (modul 16)."""
        from openpyxl.drawing.image import Image as XLImage
        try:
            from PIL import Image as PILImage
        except ImportError:
            log.warning("brak pillow - logo pomijam")
            return

        with PILImage.open(sciezka) as obraz:
            oryginal_w, oryginal_h = obraz.size
        skala = wysokosc_px / oryginal_h
        obraz_xl = XLImage(sciezka)
        obraz_xl.height = wysokosc_px
        obraz_xl.width = int(oryginal_w * skala)
        ws.add_image(obraz_xl, "A1")

    def _arkusz_metodyka(self, wb, raport: ModelRaportu, styl):
        ws = wb.create_sheet("Metodyka")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Metodyka raportu"
        ws["A1"].style = "RaportTytul"
        for indeks, akapit in enumerate(raport.metodyka, start=3):
            komorka = ws.cell(row=indeks, column=1, value=akapit)
            komorka.style = "RaportMetodyka"
        ws.column_dimensions["A"].width = 110
        return ws

    def _arkusz_meta(self, wb, raport, styl, ostrzezenia: list[str]):
        """F9. 'Karton z dwiema datami' (modul 26, sekcja 1.4)."""
        ws = wb.create_sheet("_meta")
        pola = [
            ("schemat", raport.schemat),
            ("generator", raport.wersja_generatora),
            ("wygenerowano", raport.wygenerowano),
            ("okres", raport.okres),
            ("hash_wsadu", raport.hash_wsadu),
            ("regionow", len(raport.regiony)),
            ("ostrzezen", len(ostrzezenia)),
        ]
        for indeks, (etykieta, wartosc) in enumerate(pola, start=1):
            ws.cell(row=indeks, column=1, value=etykieta).style = "RaportMetaEtykieta"
            ws.cell(row=indeks, column=2, value=wartosc)
        for indeks, tresc in enumerate(ostrzezenia, start=len(pola) + 2):
            ws.cell(row=indeks, column=1, value=tresc)
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 60
        ws.sheet_state = "hidden"
        return ws

    def _wlasciwosci(self, wb, raport: ModelRaportu) -> None:
        wb.properties.title = raport.tytul
        wb.properties.creator = "generator raportow"
        wb.properties.lastModifiedBy = "generator raportow"
        wb.properties.category = "Sprzedaz"
        wb.properties.version = raport.wersja_generatora
        wb.properties.description = (
            f"schemat={raport.schemat}; generator={raport.wersja_generatora}; "
            f"wsad={raport.hash_wsadu}; okres={raport.okres}"
        )
        wb.properties.keywords = f"schemat:{raport.schemat};okres:{raport.okres}"

    def _ochrona_wszystkich(self, wb, haslo: str) -> None:
        """F10. Komendy w pamieci - wykonywane przez JednostkePracy."""
        for nazwa in ("Podsumowanie", "Regiony", "Dynamika"):
            ws = wb[nazwa]
            ws.protection.sheet = True
            ws.protection.password = haslo
            ws.protection.enable()

    @staticmethod
    def _hash(sciezka: Path) -> str:
        skrot = hashlib.sha256()
        with sciezka.open("rb") as plik:
            for porcja in iter(lambda: plik.read(65536), b""):
                skrot.update(porcja)
        return skrot.hexdigest()[:16]
```

### Decyzja 4: dlaczego logo jest wstawiane **przed** arkuszem `Metodyka`

Kolejność w `eksportuj()` to: `podsumowanie → regiony → dynamika → logo → metodyka → meta`. Logo **musi** być wstawione po utworzeniu arkusza `Podsumowanie` (bo nie ma czego ozdabiać przed), ale **przed** opera­cjami, które mogłyby przesuwać zakresy (np. wstawianiem wierszy). W praktyce ten projekt nie wstawia wierszy po fakcie — ale **kolejność jest zapisana w kodzie celowo**, bo przy rozszerzeniu o kolejne arkusze każdy programista będzie się zastanawiał, gdzie wstawić logo. Wpisanie tego w jedno miejsce (i opisanie w komentarzu) to ta sama zasada, której broniliśmy w module 26: **decyzje są w kodzie, nie w głowie autora**.

### Decyzja 5: co świadomie zostawiam poza zakresem projektu

- **PDF przez LibreOffice headless.** Rozszerzenie: `EksporterPdfLibreOffice` implementujący ten sam port, wywołujący `subprocess.run(["soffice", "--headless", "--convert-to", "pdf", ...])`. Wchodzi w M11 (nie ma go w projekcie). Powód pominięcia: to zewnętrzny proces z timeoutem, kodem wyjścia i własnym sprzątaniem — materiał na osobny moduł.
- **Kolejka zadań i worker.** Wchodzi wtedy, gdy 200 tys. wierszy trzeba wygenerować asynchronicznie. `eksportuj()` jest **bezstanowy**, więc działa **identycznie** w request/response i w workerze. Decyzja jest w warstwie aplikacji, nie w adapterze (moduł 26, sekcja 2.11).
- **Szyfrowanie pliku.** openpyxl **nie umie** zapisać zaszyfrowanego `.xlsx`. Hasło w `ws.protection` to ochrona struktury, nie szyfrowanie. Jeżeli wymaganie mówi „poufne", to trzeba to nazwać wprost w `README.md` i użyć narzędzia zewnętrznego.
- **Sparkline'y.** openpyxl **nie ma API**. Jeżeli wymaganie brzmi „małe wykresy w komórkach", trzeba albo zmienić bibliotekę na `xlsxwriter`, albo zmienić wymaganie na pasek danych.
- **Interaktywny import (Excel jako formularz wejściowy).** To jest **inny system**. Nasz projekt ma tylko port wejściowy `ZrodloDanych` z plikiem CSV — jeżeli kiedyś pojawi się wymaganie „użytkownik uzupełnia arkusz i przysyła go z powrotem", trzeba zbudować oddzielny port wejściowy z walidacją i kontrolą wersji wiersza, a nie „dopisać się do tej samej ścieżki".

</details>

## 6. Typowe błędy i pułapki

**1. Objaw: „napisałem cały generator od razu — 900 linii w jednym pliku. Uruchamiam i dostaję `AttributeError: 'NoneType' object has no attribute 'append'` w linii 743. Nie wiem, gdzie zacząć."**
→ *Przyczyna:* **brak kamieni milowych.** Kod został napisany „od końca do początku" — z myślą o pełnym raporcie, ale bez etapu, na którym każda część jest sprawdzona. Pierwszy błąd pojawia się po 700 liniach, więc debugowanie oznacza czytanie 700 linii kodu, którego się nie rozumie, bo nikt go nie uruchamiał w częściach. To jest najdroższy możliwy tryb pracy i najczęstszy.
→ *Naprawa:* **kamienie milowe M1–M10, po każdym działający kod i przechodzący test.** Po M1 masz plik z jedną komórką. Po M4 masz cztery arkusze bez formatowania. Zmiany wprowadzasz **po jednej**, za każdym razem uruchamiając testy i otwierając plik w Excelu. Diagnostyka: jeżeli nie umiesz odpowiedzieć na pytanie „co działało ostatnim razem", to znaczy, że nie masz kamieni milowych.

**2. Objaw: „`grep -rn "openpyxl" src/raporty/` zwraca wynik w `application/raport_sprzedazy.py`. Jest tam tylko `from openpyxl.styles import NamedStyle` — bo potrzebowałem jednej stałej ze stylu."**
→ *Przyczyna:* **przeciek abstrakcji przez „jedną niewinną linię".** Tak działa 95 % przecieków: nie przez wielki import, a przez jedną potrzebną stałą, jeden typ do adnotacji, jedną funkcję pomocniczą. Konsekwencja jest natychmiastowa i widoczna dopiero w CI: `test_bez_openpyxl` przestaje przechodzić, bo `application` nie da się zaimportować bez biblioteki.
→ *Naprawa:* **stała należy do warstwy, w której jest definiowana, a nie tam, gdzie się jej używa.** Nazwa stylu jest **konfiguracją prezentacji** → definiuj ją w `infrastructure/excel/styl.py` i przekazuj do aplikacji jako **string w spec-u** albo **w słowniku przekazywanym do adaptera**. Nigdy odwrotnie. Diagnostyka: **job CI bez openpyxl to nie dodatek, a jedyne narzędzie**, które wyłapuje ten błąd automatycznie. Ręcznie tego nie znajdziesz raz na zawsze — przeciek wraca przy następnej zmianie.

**3. Objaw: „Testy przechodzą: sprawdzam, że `ws["D2"].number_format == '#,##0.00 "zl"'`. Ale w Excelu kolumna D pokazuje `1 234,56 zł` — a klient mówi, że powinno być `1 234,56 PLN`."**
→ *Przyczyna:* **test sprawdza kod formatu, nie efektu.** Kod formatu w pliku to jedna rzecz, a **wyświetlanie** to druga — zależy od locale, ustawień systemu i symbolu `[$-415]` (moduł 10). Test asertujący kod formatu jest **przydatny**, ale **nie wystarcza**. To jest ten sam problem, o którym mówiłem w sekcji 1.5: część wymagań da się zbadać tylko okiem.
→ *Naprawa:* **podział weryfikacji na trzy poziomy** (sekcja 1.5). Test sprawdza kod formatu (poziom 1). Weryfikacja struktury pliku sprawdza, że format istnieje w `styles.xml` (poziom 2). Oko człowieka sprawdza **wygląd** (poziom 3). Jeżeli klient chce symbol `PLN` a nie `zł`, zmieniasz **jeden wpis** w `FORMATY` w `styl.py` — i nie musisz ruszać ani testów, ani logiki, ani konfiguracji.

**4. Objaw: „Wykres dwuosiowy wygląda dziwnie: linia marży jest płaska na samym dole wykresu. Wartości marży są w porządku (0,05; 0,08; 0,12), ale linia nie widać."**
→ *Przyczyna:* **zapomniałeś `axId = 200` i `crosses = "max"`.** Bez tego obie serie są rysowane na **jednej** osi wartości — a skala jest wyznaczana przez sprzedaż (setki tysięcy). Wartość 0,05 na osi idącej do 500 000 to praktycznie zero, więc linia leży dokładnie na osi X. Nie ma błędu. Wykres po prostu **nie robi tego, o co prosiłeś**.
→ *Naprawa:* **trzy linie, w tej kolejności** (z przykładu M6):
```python
linia.y_axis.axId = 200              # drugie ID osi wartosci
linia.y_axis.title = "Marza (%)"
linia.y_axis.numFmt = "0.0%"
slupki.y_axis.crosses = "max"        # os sprzedazy przesun do krawedzi
slupki += linia                      # sklej serie
```
Diagnostyka: **otwórz plik w Excelu i kliknij wykres**. Jeżeli w okienku „Formatuj osie" widzisz **jedną** oś, to znaczy, że druga nie istnieje. Test na sam `xl/charts/chart1.xml` tego nie wyłapie — kolejny powód, dlaczego „na oko" jest obowiązkowe.

**5. Objaw: „Pasek danych na wartościach ujemnych wygląda jak plama — wszystkie paski są czerwone i idą od lewej krawędzi, bez różnicy między −3 % a −40 %."**
→ *Przyczyna:* **pasek danych bez `start_value` ujemnego.** Domyślnie `DataBarRule` zakłada zakres od **zera** do maksimum — jeżeli wartości są ujemne, wszystkie trafiają do „lewej" części i wyglądają tak samo. Pasek danych sam w sobie **nie skaluje się przez zero**, jeżeli mu tego nie każesz.
→ *Naprawa:* **ustaw oba końce jawnie.** Dla dynamiki w zakresie ±50 %:
```python
DataBarRule(start_type="num", start_value=-0.5,
            end_type="num", end_value=0.5,
            color="638EC6", showValue=True)
```
Alternatywnie `start_type="min"`, `end_type="max"` — ale wtedy pasek będzie się skalował do **aktualnych** danych i dwa raporty (z innym zakresem) będą wyglądały zupełnie inaczej, co jest **dezorientujące** w porównaniach międzyokresowych. Dla raportu okresowego **zawsze** ustawiaj zakresy liczbowe — są wtedy porównywalne.

**6. Objaw: „Test `test_kontrakt_portu` przechodzi dla `xlsx` i dla `csv`, ale wywala się na `assert wynik.arkuszy == 1`."**
→ *Przyczyna:* **to nie jest błąd testu — to jest błąd portu.** Konkretnie: port obiecuje coś, czego druga implementacja nie może spełnić. Ta asercja przechodzi dla Excela (który ma arkusze) i nie przechodzi dla CSV (który arkuszy nie ma). To jest **właśnie ta informacja**, po którą napisałeś ten test.
→ *Naprawa:* **zdecyduj, czy port obiecuje arkusze, czy nie, i zmień port — nie test.** Trzy opcje: (a) usuń `arkuszy` z `WynikEksportu` (może być liczone po stronie adaptera xlsx w jego własnym logu), (b) zmień `arkuszy` na `stron`/`sekcji` (neutralne dla wszystkich), (c) zdefiniuj, że `arkuszy` to **liczba części logicznych** (CSV zwraca 1 dla jednego pliku, 3 dla trzech plików). Każda opcja jest w porządku — **błędem jest zmienienie testu** tak, żeby przechodził. To jest lekcja z modułu 26: **dwa adaptery obnażają kontrakt**.

**7. Objaw: „Po dwóch tygodniach w `output/` mam 340 plików `raport_sprzedazy_v1.xlsx` … `raport_sprzedazy_v340.xlsx`, z czego 200 ma rozmiar zero bajtów, a dysk się kończy."**
→ *Przyczyna:* **polityka nazw „wersjonuj" bez retencji.** Ktoś napisał pętlę, która szuka wolnego numeru, i uznał, że problem rozwiązany. Problem **przeniósł się** na dysk i na pytanie „który plik jest aktualny": wersjonowanie bez sprzątania zamienia „nadpisanie" na „śmieci". Pliki zerobajtowe to dodatkowy problem — to wynik wcześniejszej awarii, który został zapisany „na wszelki wypadek".
→ *Naprawa:* **trzy rzeczy razem.** (1) **Polityka retencji** — np. „trzymaj ostatnie 30 dni, starsze usuwaj" — skrypt sprzątający w CI albo w samym CLI. (2) **Wybieraj politykę świadomie** (`nadpisz`, `data`, `wersjonuj`) i **nie wdrażaj** domyślnie `wersjonuj` — większość raportów to dokumenty robocze, które powinny być nadpisywane. (3) **Zapis atomowy gwarantuje, że plik zerobajtowy nie powstanie** — `os.replace` podmienia **cały** plik naraz, więc nie ma stanu „plik istnieje, ale jest pusty". Jeżeli masz pliki zerobajtowe, to znaczy, że któryś etap **nie** przechodzi przez `JednostkaPracy` (pułapka 8).

**8. Objaw: „W `output/` znajduję pliki `.{nazwa}.tmp.xlsx` — jeden na każdy błąd, który wystąpił w tym tygodniu."**
→ *Przyczyna:* **brak sprzątania w `finally`.** Najczęściej dlatego, że sprzątanie jest w `except`, a nie w `finally` — więc nie uruchamia się przy `KeyboardInterrupt` albo przy wyjątku **nie-`Exception`** (`SystemExit`). Drugi scenariusz: `unlink` jest w `except`, ale `except` **nie chwyta** danego typu wyjątku (np. `BaseException`).
→ *Naprawa:* **`try/finally` wokół całości.** Wzorzec jest w module 26 (przykład 5) i wygląda tak:
```python
tmp = Path(tmp_name)
try:
    wb.save(tmp)
    # ... weryfikacja, backup, os.replace ...
    return cel
finally:
    tmp.unlink(missing_ok=True)     # ZAWSZE, takze przy KeyboardInterrupt
```
Uwaga na subtelność: `unlink(missing_ok=True)` musi być **po** `os.replace` — jeżeli `os.replace` się udał, tmp **nie istnieje**, i `missing_ok=True` to obsłuży. Diagnostyka: `assert not list(cel.parent.glob(".*.tmp.xlsx"))` w teście atomowości.

**9. Objaw: „Loguję do `output/audyt.jsonl` i po awarii zasilania ostatnia linia jest ucięta w połowie. Parser JSON w narzędziu do analizy się wywala i gubię **wszystko**."**
→ *Przyczyna:* **parser, który zakłada poprawność każdej linii.** `json.loads()` na uciętej linii rzuca wyjątek — a jeżeli Twój analizator robi to w pętli **bez `try/except`**, traci wszystkie poprzednie rekordy. To nie jest wina JSONL: to jest wina parsera, który założył doskonały plik.
→ *Naprawa:* **parsuj odpornie:**
```python
def czytaj_dziennik(sciezka: Path) -> list[dict]:
    wpisy: list[dict] = []
    niepelne = 0
    with sciezka.open(encoding="utf-8") as plik:
        for numer, linia in enumerate(plik, start=1):
            linia = linia.strip()
            if not linia:
                continue
            try:
                wpisy.append(json.loads(linia))
            except json.JSONDecodeError:
                niepelne += 1
                log.warning("dziennik: linia %d niepelna (awaria zapisu?)", numer)
    if niepelne:
        log.warning("dziennik: pominięto %d niepełnych linii", niepelne)
    return wpisy
```
I jeszcze jedno: **JSONL wybrano dlatego, że ucięcie dotyczy tylko ostatniej linii** — gdyby to był jeden wielki JSON, straciłbyś cały plik. Logika wyboru jest w komentarzu w kodzie (bo inaczej ktoś „ulepszy" format przy okazji i problem wróci).

**10. Objaw: „`test_200k_wierszy` przechodził lokalnie w 38 sekundach. W CI przekracza limit czasu na 6 minut i zżera 3 GB RAM."**
→ *Przyczyna:* **trzy różnice środowiska, o których się nie myśli.** (1) CI ma **wolniejszy dysk** — `wb.save()` na tmpfs/HDD jest 5× wolniejszy niż na NVMe. (2) CI może **nie mieć `lxml`** — wtedy openpyxl używa czystego `xml.etree`, co jest **kilkukrotnie wolniejsze** (moduł 19). (3) CI może mieć inne limity pamięci kontenera i **nie ma swapu**.
→ *Naprawa:* **trzy rzeczy.** (1) **`lxml` jako jawna zależność produkcyjna**, nie opcjonalna — w projekcie, który deklaruje 200 000 wierszy w budżecie, brak `lxml` to **niespełnione wymaganie**, a nie optymalizacja. Dodaj test: `assert openpyxl.LXML` albo sprawdzenie `import lxml` na starcie. (2) **Testy `slow` osobno** (`@pytest.mark.slow`, job z dłuższym timeoutem) — żeby nie blokowały szybkiego feedbacku. (3) **Mierz `tracemalloc` w CI** i loguj szczyt pamięci — bez tego nie wiadomo, czy zbliżasz się do limitu, czy już go przekraczasz na danych produkcyjnych.

**11. Objaw: „Pasek danych i czerwone tło na wykresie wyglądają dobrze w `Regiony`, ale Arkusz `Dynamika` po **drugim** uruchomieniu wygląda inaczej — nie ma drugiej reguły."**
→ *Przyczyna:* **nadpisanie `conditional_formatting` dla tego samego zakresu.** `ws.conditional_formatting.add(zakres, reguła)` powinno **dodawać** regułę do zbioru. Ale jeżeli w kodzie jest ścieżka, która dla tego samego zakresu tworzy **nowy** obiekt `ConditionalFormatting` (np. przez `ws.conditional_formatting = ...` w miejscu), tracisz poprzednie reguły. W 3.1.x jest to naprawione, ale pułapka wraca przy migracji ze starszych wersji i przy przestarzałych przykładach z Internetu.
→ *Naprawa:* **zawsze `add()`, nigdy przypisanie.** Diagnostyka **obowiązkowa**: po wczytaniu pliku policz reguły i zakresy:
```python
def policz_reguly(sciezka: Path) -> dict[str, int]:
    wb = load_workbook(sciezka)
    try:
        return {
            nazwa: sum(len(cf.rules) for cf in wb[nazwa].conditional_formatting)
            for nazwa in wb.sheetnames
        }
    finally:
        wb.close()
```
Asercja „`Dynamika` ma 2 reguły" jest częścią testu regresji i **wyłapie** ten błąd automatycznie.

**12. Objaw: „Klient przysłał raport z powrotem w formacie `.xlsx` z naniesionymi zmianami. Skrypt wczytał, dopisał kolumnę i zapisał. Po otwarciu zniknęły sparkline'y i slicery. Klient nie zauważył, ale to jego dane."**
→ *Przyczyna:* **moduł 18 w praktyce.** Wczytano plik, który zawierał elementy nieznane openpyxl, i zapisano — czyli **przepisano** go bez tych elementów. To jest najpoważniejszy praktyczny błąd tego typu projektu, bo **jest cichy**. Nikt nie sprawdza sparkline'ów, bo „przecież plik się otwiera".
→ *Naprawa:* **trzy twarde reguły.** (1) **Nigdy nie modyfikuj pliku, który nie powstał z tego generatora, w miejscu** — zapisuj **nowy** plik z sufiksem `_v2`. (2) **Skanuj przed zapisem** (`przeskanuj_przed_zapisem` z M9) i umieszczaj ostrzeżenia w `_meta` i w logu jako `WARNING`. (3) **Nie dopisuj kolumn do cudzych plików** — jeżeli naprawdę trzeba, to jest to operacja o **wysokim ryzyku**, wymagająca archiwizacji oryginału i jawnego ostrzeżenia w interfejsie. Diagnostyka: narzędzie `scan` podpięte jako osobna komenda CLI do sprawdzania **każdego** pliku wejściowego.

**13. Objaw: „Program przetworzył 200 000 wierszy w 42 sekundy. To mieści się w 60. Ale wynikowy plik ma 2,1 GB i Excel go nie otwiera."**
→ *Przyczyna:* **spełnienie `N1` nie oznacza spełnienia wymagania „użytkownik może otworzyć plik".** 200 000 wierszy × 8 kolumn z formatowaniem to **nie** 2 GB — to raczej 20–40 MB. Skok do 2,1 GB oznacza jedno z trzech: (a) każda komórka ma **osobny** obiekt stylu (nadmiarowy Flyweight), (b) reguły warunkowe są dodane **per wiersz**, a nie na zakres, (c) w pliku są zapisane **wartości pośrednie**, o których nie wiesz (np. ukryty arkusz roboczy z danymi).
→ *Naprawa:* **trzy kontrole.** (a) **Sprawdź rozmiar pliku po każdym kamieniu milowym** — nie po całym projekcie; skok objawi się natychmiast i będziesz wiedzieć, **który etap** go spowodował. (b) **Style przez `cell.style = "Nazwa"`** albo przez **współdzielone** obiekty (rejestr) — nigdy `Font(bold=True)` utworzony w pętli. (c) **Reguła warunkowa na zakres, nie na komórkę** — to jest różnica między **jednym** elementem XML a 200 000. Diagnostyka: `zipfile.ZipFile(...).getinfo("xl/styles.xml").file_size` — jeżeli `styles.xml` ma megabajty, masz problem ze stylami.

**14. Objaw: „Wszystko działa, dopóki nie spróbuję wygenerować raportu na komputerze kolegi. U niego wywala się `PermissionError` przy zapisie, bo ma plik otwarty w Excelu. U mnie działa."**
→ *Przyczyna:* **brak obsługi konkretnego błędu i brak konsekwentnej polityki.** Windows i Linux różnią się **fundamentalnie** w blokowaniu podmiany otwartego pliku: Windows blokuje (`PermissionError`), Linux nie (podmienia „pod spodem"). Kod, który „działa u mnie", jest kodem, który **nie obsługuje najbardziej prawdopodobnego** przypadku błędu w produkcji na Windows — a na Windows siedzi 95 % użytkowników biznesowych.
→ *Naprawa:* **jawna obsługa i jawny komunikat:**
```python
try:
    os.replace(tmp, cel)
except PermissionError as blad:
    raise BladZapisu(
        f"Plik {cel.name} jest otwarty w innym programie "
        f"(najczęściej Excel). Zamknij go i uruchom ponownie, "
        f"albo użyj --polityka data, żeby zapisać pod nową nazwą."
    ) from blad
```
I **test** — nie „otwórz w Excelu", bo tego nie da się zrobić w CI, ale symulacja: `mock.patch("os.replace", side_effect=PermissionError)`. Komunikat **musi** być instrukcją dla użytkownika, a nie tracebackiem.

### Checklista przed wdrożeniem

Przejdź ją **całą** przed wypuszczeniem generatora do ludzi. Każde „nie" to konkretne ryzyko.

- [ ] **1. Czy dane wejściowe są nienaruszone po przebiegu?** Sprawdź hash pliku CSV przed i po. Jeżeli się zmienił — masz poważny błąd (w 99 % przypadków to przypadkowe `wb.save()` na tej samej ścieżce albo dopisywanie do pliku wejściowego).
- [ ] **2. Czy plik wyjściowy **nigdy** nie jest nadpisywany po cichu?** Polityka nazw jest jawna (`nadpisz`/`data`/`wersjonuj`) i widoczna w konfiguracji. Nie ma fallbacku na `raport(1).xlsx`.
- [ ] **3. Czy zapis jest atomowy?** `mkstemp(dir=cel.parent)` + `os.replace` + `finally: unlink`. Test z symulowanym wyjątkiem: plik docelowy nietknięty, brak `.tmp`.
- [ ] **4. Czy backup powstaje **przed** podmianą?** Kolejność `copy2 → replace`, nie odwrotnie. Backup ma timestamp.
- [ ] **5. Czy potrafisz wytłumaczyć, co się stanie, gdy plik jest otwarty w Excelu?** Komunikat dla użytkownika istnieje i zawiera instrukcję.
- [ ] **6. Czy dwa uruchomienia z tym samym wsadem dają ten sam plik (poza znacznikami czasu)?** Test `a.hash_wyniku == b.hash_wyniku` przechodzi.
- [ ] **7. Czy drugie uruchomienie **nie zmienia** istniejącego pliku?** Idempotencja: `podsumowanie_skoroszytu` identyczne.
- [ ] **8. Czy domena i aplikacja importują się **bez** openpyxl?** Job CI bez openpyxl przechodzi. To nie jest opcja — to kontrakt.
- [ ] **9. Czy sanitacja jest na **wejściu** do domeny, nie przy zapisie?** Jedna funkcja, `grep -n "ZNAKI_WYZWALAJACE" src/raporty/` → jedno miejsce w `domain/`.
- [ ] **10. Czy pliki mają wersję generatora i hash wsadu?** Otwórz **losowy** plik wygenerowany trzy miesiące temu i odpowiiedz na pytanie „skąd to i czy jest kompletne" **bez** logu serwera.
- [ ] **11. Czy wiesz, ile wierszy i jak długo trwa generowanie dużego raportu?** Zmierzone `N1` i `N2`. Budżet jest **liczbą**, nie „szybko".
- [ ] **12. Czy program się zachowuje, gdy dane wejściowe mają błąd?** Test „zaburzonego CSV": `=HYPERLINK(...)`, pusty region, brakująca kolumna. Każdy przypadek daje czytelny komunikat, nie cichy pominięty wiersz.
- [ ] **13. Czy wykrywasz utratę funkcji pliku?** Skan `zipfile.namelist()` przed zapisem; ostrzeżenia w `_meta` i w logu jako `WARNING`, nie `INFO`.
- [ ] **14. Czy otworzyłeś plik w Excelu i **obejrzałeś** go?** Wszystkie arkusze, wszystkie wykresy, wszystkie reguły. **To jest jedyny krok, którego nie da się zastąpić testem** — i dlatego jest na końcu checklisty, żeby nie zniknął pod pozorem „mam to pokryte testami".

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Kamienie milowe są częścią metody, nie uprzejmością.** Dziesięć kroków `M1`–`M10`, po każdym **działający kod i przechodzący test**. To jest jedyna różnica między projektem, który się udaje, a projektem, w którym „po trzech dniach mam 700 linii i nie wiem, gdzie jest błąd". Kolejność kroków **musi** być taka: domena bez infrastruktury (M2) → port i adapter (M3) → funkcje (M4–M8) → transakcja (M9) → wydajność (M10). Każda inna kolejność mnoży dług: domena „z myślą o Excelu" przecieka, brak portu czyni M4–M8 nieprzetestowalnymi, brak transakcji czyni M10 niebezpiecznym.

2. **Specyfikacja to ponumerowane wymagania, z których każde ma sposób weryfikacji.** `F1`–`F12` i `N1`–`N10`. Numer pozwala powiedzieć w code review „ta funkcja realizuje F6, a nie realizuje F4". Sposób weryfikacji pozwala odróżnić „test automatyczny" od „weryfikacji na oko". **Nie wszystkie wymagania da się sprawdzić testem** — i specyfikacja musi to mówić **wprost**, bo inaczej powstaje test, który sprawdza kod formatu i nie widzi, że kolumna pokazuje zły symbol waluty.

3. **Architektura ma zarabiać na siebie — i dlatego mierzysz, ile razy kod się wykona.** Poziom 1 (jeden plik, 120 linii) jest **właściwy** dla raportu jednorazowego. Poziom 3 (warstwy, porty, transakcje) jest właściwy dla raportu, który zmieniał się siedem razy w roku. **Nie ma jednej „właściwej" architektury** — jest architektura dopasowana do tempa zmian. Kto zawsze wybiera poziom 3, marnuje czas na projekcie jednorazowym; kto zawsze wybiera poziom 1, traci projekt, gdy zmiany przyjdą. Umiejętność polega na **rozpoznaniu, w którym przypadku jesteś**.

4. **Dwie decyzje architektoniczne, które widzisz dopiero na żywym projekcie.** (a) **Przy dużych danych raport trzeba rozdzielić na dwa pliki** — strumieniowy z surowizną (`write_only`, `lxml`) i bogaty z agregatami (normalny, z wykresem i tabelą). Próba zmieszczenia obu w jednym to wybór między „nie spełniam N1" i „nie spełniam F6". (b) **Reguły biznesowe to zamknięta paleta w kodzie, nie wyrażenia w YAML-u.** Konfiguracja **wybiera nazwę** i podaje parametr. Konfiguracja, która wyraża obliczenia, staje się nowym językiem programowania — bez testów, bez podpowiedzi edytora i bez historii zmian.

5. **„Działa u mnie" to nie jest wymaganie — jest nim „działa tam, gdzie pójdzie".** Transakcyjność (`N3`), determinizm (`N6`), brak nadpisania wejścia (`N5`), sanitacja na wejściu do domeny (`F11`), wykrywanie utraty funkcji (`N10`), jawny komunikat przy otwartym pliku (pułapka 14) — to nie są „dodatki". To są **odpowiedzi na pytania „co się stanie, gdy coś pójdzie źle"**. W systemie produkcyjnym **zawsze** coś idzie źle: plik otwarty w Excelu, kolumna w innym porządku, klient z `=` w nazwie, brak `lxml`, plik źródłowy ze sparkline'ami, dysk pełen, wyjątek w połowie zapisu. Każda z tych sytuacji ma w tym projekcie **nazwane rozwiązanie** — i to jest cała różnica między skryptem a programem.

## 8. Ściągawka modułu

```python
# ==================================================================
# STRUKTURA PROJEKTU (pelna, z sekcji 2.4)
# ==================================================================
# raporty/
#   assets/logo.png
#   config/raport_sprzedazy.yaml, raport_minimal.yaml
#   data/orders.csv                     <- WEJSCIE, tylko odczyt
#   output/                             <- wszystko, co generujemy
#   src/raporty/
#     __main__.py                       <- CLI: python -m raporty --config ...
#     domain/    modele.py reguly.py paleta.py porty.py        (ZERO openpyxl)
#     application/ raport_bazowy.py raport_sprzedazy.py raport_minimal.py komendy.py
#     infrastructure/ excel/ zrodla/ konfig/ fake/ audyt/      (openpyxl TU)
#   tests/domain/ tests/application/ tests/infrastructure/excel/


# ==================================================================
# 1. CLI - JEDYNE miejsce kompozycji (modul 26)
# ==================================================================
# python -m raporty --config config/raport_sprzedazy.yaml
# python -m raporty --config config/raport_minimal.yaml --polityka data

def kompozycja(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="raporty")
    parser.add_argument("--config", required=True)
    parser.add_argument("--polityka", default="wersjonuj",
                        choices=("nadpisz", "data", "wersjonuj"))
    parser.add_argument("--log", default="INFO")
    args = parser.parse_args(argv)

    spec = wczytaj_spec(args.config)             # waliduje PRZED generowaniem
    zrodlo = ZrodloCsv(spec.dane)
    eksporter = EksporterXlsxOpenpyxl(
        haslo_ochrony=spec.arkusze.ochrona_haslo,
        polityka_nazw=args.polityka,
    )
    magazyn = DziennikJsonl(Path("output/audyt.jsonl"))

    przypadek = RaportSprzedazy(zrodlo=zrodlo, eksporter=eksporter,
                                magazyn=magazyn, spec=spec)
    wynik = przypadek.generuj(okres=spec.raport.okres, cel="output/raport_sprzedazy")
    print(f"OK plikow={len(wynik.wynik.sciezki)} wierszy={wynik.wynik.wierszy} "
          f"czas={wynik.czas_s:.1f}s ostrzezen={len(wynik.ostrzezenia)}")
    return 0


# ==================================================================
# 2. KAMIENIE MILOWE - kolejnosc OBOWIAZKOWA
# ==================================================================
# M1  szkielet: Workbook + save                          -> plik istnieje
# M2  domain/modele.py + reguly.py (sanitacja!)          -> testy bez openpyxl
# M3  porty.py + eksporter.py + fake/eksporter.py        -> test kontraktu
# M4  styl.py + arkusze (RejestrStylow, cell.style=)     -> F1
# M5  DataBarRule + CellIsRule                           -> F5, F7
# M6  wykresy.py (axId=200, crosses="max", +=)           -> F6
# M7  Table + Image + Metodyka                           -> F2, F3, F8
# M8  protection + DataValidation + wb.properties        -> F9, F10, F12
# M9  JednostkaPracy + utrata.py + dziennik.jsonl        -> N3, N9, N10
# M10 eksporter_danych.py (write_only) + CI + determinizm -> N1, N2, N4, N6


# ==================================================================
# 3. TRANSKAKCYJNY ZAPIS (N3) - kolejnosc NIENARUSZALNA
# ==================================================================
# 1. komendy w PAMIECI          -> blad = pelny rollback, dysk nietkniety
# 2. mkstemp(dir=cel.parent)    -> TYLKO ten sam wolumen!
# 3. wb.save(tmp)
# 4. load_workbook(tmp, read_only=True) -> weryfikacja + close()!
# 5. skan utraty funkcji        -> ostrzezenia WARNING
# 6. shutil.copy2(cel, backup)  -> PRZED podmiana, z timestampem
# 7. os.replace(tmp, cel)       -> COMMIT, jeden ruch
# 8. finally: tmp.unlink(missing_ok=True)


# ==================================================================
# 4. NAJWAZNIEJSZE FRAGMENTY openpyxl W TYM PROJEKCIE
# ==================================================================
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# NAZWANE STYLE (F1, szybkosc, spojnosc):
wb.add_named_style(NamedStyle(name="RaportWaluta",
                              number_format='#,##0.00 "zl";(#,##0.00)'))
cell.style = "RaportWaluta"

# USAWANIE DOMYSLNEGO ARKUSZA (F1):
wb = Workbook(); wb.remove(wb.active)

# TABELA Z FILTREM (F3):
tb = Table(displayName="TabelaRegiony", ref=f"A4:D{ostatni}")
tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
ws.add_table(tb)

# PASEK DANYCH dla wartosci UJEMNYCH (F7):
ws.conditional_formatting.add("D2:D37", DataBarRule(
    start_type="num", start_value=-0.5, end_type="num", end_value=0.5,
    color="638EC6", showValue=True))

# SPADKI (F5):
ws.conditional_formatting.add("D2:D37", CellIsRule(
    operator="lessThan", formula=["0"],
    fill=PatternFill(fill_type="solid", start_color="FCE4E4", end_color="FCE4E4"),
    font=Font(color="9C0006", bold=True)))

# WYKRES DWUOSIOWY (F6) - trzy kluczowe linie:
marza.y_axis.axId = 200
slupki.y_axis.crosses = "max"
slupki += marza

# KOLOR SERII - graphicalProperties, NIE series.color:
from openpyxl.chart.shapes import GraphicalProperties
series.graphicalProperties = GraphicalProperties(solidFill="1F4E78")

# LOGO z ZACHOWANIEM PROPORCJI (F2):
with PILImage.open(sciezka) as ob:
    orig_w, orig_h = ob.size
skala = 44 / orig_h
img = XLImage(sciezka); img.height = 44; img.width = int(orig_w * skala)
ws.add_image(img, "A1")

# OCHRONA (F10) - locked=False PRZED enable():
cell.protection = Protection(locked=False)
ws.protection.sheet = True
ws.protection.password = haslo
ws.protection.enable()

# WALIDACJA Z ZAKRESU (F12) - nie z literalnej listy (limit 255 znakow!):
dv = DataValidation(type="list", formula1="=Regiony!$A$2:$A$20", allow_blank=True)
ws.add_data_validation(dv); dv.add("B2:B100")

# WLASCIWOSCI DOKUMENTU (F9):
wb.properties.title = raport.tytul
wb.properties.version = raport.wersja_generatora
wb.properties.description = f"schemat={raport.schemat}; wsad={raport.hash_wsadu}"

# STRUMIENIOWY ZAPIS 200k WIERSZY (N1, N2):
wb = Workbook(write_only=True)
ws = wb.create_sheet("Dane")
cell = WriteOnlyCell(ws, value=naglowek)
cell.font = _FONT_NAGLOWKA          # utworzony RAZ, poza petla!
ws.append([cell, ...])
for wiersz in generator:            # generator, nie lista!
    ws.append(list(wiersz))
wb.save(cel)


# ==================================================================
# 5. TESTY - TRZY POZIOMY (sekcja 1.5)
# ==================================================================
# POZIOM 1 - testy automatyczne (pytest, BytesIO, bez plikow):
#   testy domeny i aplikacji -> job CI BEZ openpyxl  (N8)
#   test kontraktu portu parametryzowany po 3 adapterach
#   test atomowosci: wyjatek w polowie -> plik nietkniety, brak .tmp (N3)
#   test determinizmu: a.hash_wyniku == b.hash_wyniku (N6)
#   test idempotencji: podsumowanie_skoroszytu identyczne (N4)
#   test wydajnosci: @pytest.mark.slow, 200k < 60 s, < 500 MB (N1, N2)
#
# POZIOM 2 - weryfikacja struktury pliku (po zapisie):
def podsumowanie_skoroszytu(sciezka) -> dict:
    wb = load_workbook(sciezka, read_only=True)
    try:
        return {
            "arkusze": tuple(wb.sheetnames),
            "wiersze": {n: sum(1 for _ in wb[n].iter_rows()) for n in wb.sheetnames},
            "regul": {n: sum(len(cf.rules) for cf in wb[n].conditional_formatting)
                      for n in wb.sheetnames},
            "tabele": {n: tuple(wb[n].tables) for n in wb.sheetnames},
        }
    finally:
        wb.close()
#   + zipfile: xl/charts/, xl/media/, xl/sparklineGroups (N10)
#
# POZIOM 3 - OKO CZLOWIEKA (obowiazkowe, nie da sie zastapic):
#   F2 rozmiar i pozycja logo | F4 wyglad waluty | F6 dwie osie |
#   F7 paski dla wartosci ujemnych | F8 czytelnosc Metodyki | F10 ochrona dziala


# ==================================================================
# 6. CHECKLISTA PRZED WDROZENIEM (14 pytan, sekcja 6)
# ==================================================================
#   1. Dane wejsciowe nienaruszone (hash przed/po)?
#   2. Polityka nazw jawna, bez "raport(1).xlsx"?
#   3. Zapis atomowy (mkstemp dir= + os.replace + finally unlink)?
#   4. Backup PRZED podmiana, z timestampem?
#   5. Wiadomo, co sie dzieje przy pliku otwartym w Excelu?
#   6. Dwa przebiegi -> ten sam hash_wyniku (N6)?
#   7. Drugie uruchomienie nie zmienia pliku (N4)?
#   8. Domena i aplikacja importuja sie BEZ openpyxl (N8)?
#   9. Sanitacja na WEJSCIU do domeny, jedno miejsce?
#  10. Kazdy plik niesie wersje generatora i hash wsadu (N9)?
#  11. Budzet wydajnosci zmierzony liczbami (N1, N2)?
#  12. Zaburzone dane daja czytelny blad, nie ciche pominniecie?
#  13. Utrata funkcji pliku wykrywana i logowana (N10)?
#  14. Plik OTWARTY w Excelu i OBEJRZANY (poziom 3)?


# ==================================================================
# 7. REGULY, KTORYCH NIE LAMIE SIE NIGDY
# ==================================================================
# - grep -rn "openpyxl" src/raporty/domain/ src/raporty/application/  -> PUSTO
# - grep -n "if \|when \|condition" config/*.yaml                     -> PUSTO
# - tempfile.mkstemp(dir=cel.parent)  - ZAWSZE "dir="
# - data_only=True + wb.save()        - NIGDY (zniszczenie formul, modul 06)
# - keep_vba=True                     - dla .xlsm (modul 18)
# - wb.close()                        - dla kazdego read_only (modul 19)
# - bufor.seek(0)                     - przed kazdym wyslaniem BytesIO
# - os.replace na innym wolumenie     - NIE jest atomowy (sprawdz st_dev)
```

## 9. Co dalej

Skończyłeś kurs. To znaczy: masz za sobą dwadzieścia siedem modułów, od `wb = Workbook()` w module 04 do transakcyjnego generatora z dwoma plikami wyjściowymi i portami. Zanim jednak powiesz sobie „umiem", chcę Ci dać trzy ostatnie rzeczy — bo bez nich wiedza się rozsypie.

**Pierwsza: wiesz, czego openpyxl nie umie.** I to jest równie ważne co to, co umie. Nie ma silnika formuł (moduł 06). Nie ma sparkline'ów (moduł 15). Nie ma API do szyfrowania pliku (moduł 17). Gubi shapes, slicery i Power Query (moduł 18). Nie zapisuje wykresów „wiernie" — odbudowuje je z modelu. To nie są wady do obejścia, to są **granice narzędzia**, i Twoja wartość jako inżyniera polega na tym, że **znasz je na pamięć** i **potrafisz powiedzieć klientowi w dwie minuty**, czy jego pomysł jest wykonalny, czy wymaga innego narzędzia. Osoba, która „umie openpyxl", ale nie zna tych granic, obiecuje rzeczy, których potem nie dostarcza.

**Druga: masz mapę, do której wracasz.** Moduł 28 jest do tego zbudowany. Kiedy za pół roku klient powie „chcę listę rozwijaną z innego arkusza", nie będziesz czytał kursu od początku — otworzysz moduł 17, sekcję z `DataValidation`. To nie jest lenistwo; to jest **właściwy sposób korzystania z narzędzia**, i dlatego moduł 28 ma tabelę „problem → moduł", a nie tylko listę API. Kurs nie jest do przeczytania raz — jest do **używania**.

**Trzecia: dane są ważniejsze od raportu.** To zdanie powtarzałem w modułach 18, 21, 26 i 27 — i powtarzam je ostatni raz, bo jest najważniejsze. **Raport jest projekcją.** Paragonem. Wyświetleniem. Prawda jest w systemie — w bazie, w CSV, w transakcji. Jeżeli kiedykolwiek staniesz przed wyborem „zrobić szybciej" albo „zrobić bezpieczniej", wybierz bezpieczniej. Bo raport można wygenerować ponownie w 10 sekund, a **dane można stracić raz**. Cała ta maszyneria — warstwy, porty, `mkstemp` z `dir=`, `os.replace`, weryfikacja przez `load_workbook`, skan utraty funkcji, hashe wsadu, backupy z timestampem — istnieje **wyłącznie** po to, żeby ta jedna zasada nigdy nie została złamana przez Twój kod.

**I jeszcze jedno, na koniec — o tym, co robić dalej.** Trzy konkretne kroki:

1. **Weź swój najbrzydszy skrypt z openpyxl** i przepisz go przez to, czego się nauczyłeś. Nie po to, żeby go „ulepszyć" — po to, żeby **poczuć w rękach**, co daje port i transakcyjny zapis. Jedno przepisanie nauczy Cię więcej niż trzy przeczytane moduły.
2. **Zbadaj swoje pliki.** Weź trzy `.xlsx` z pracy — jeden wygenerowany przez system, jeden tworzony ręcznie, jeden przysłany przez klienta — i **rozpakuj je** `zipfile`. Zobacz, ile części ma każdy, gdzie są wykresy, czy są sparkline'y, co jest w `_meta` (albo czego tam nie ma). To jest test Twojej wiedzy: jeżeli patrzysz na `namelist()` i rozumiesz, co widzisz, to kurs spełnił swoje zadanie.
3. **Wróć do modułu 28 i zrób z niego własną ściągawkę.** Nie kopiuj — **przepisz te fragmenty, których używasz najczęściej**, w jednym pliku `.py`, który importujesz do każdego projektu. Twoja ściągawka, nie moja. Bo kod, który napisałeś sam, pamięta się inaczej.

Powodzenia. I pamiętaj o jednej rzeczy, gdy będziesz o trzeciej nad ranem debugować raport, który „nie chciał się zapisać": **jeżeli wszystkie komendy wykonały się w pamięci i zapis padł na `os.replace`, to znaczy, że Twój projekt działa dokładnie tak, jak powinien.** Plik docelowy jest nietknięty, oryginał jest w backupie, a błąd jest w logu z pełnym kontekstem. To nie jest awaria — to jest **architektura, która zrobiła to, po co ją zbudowałeś**.