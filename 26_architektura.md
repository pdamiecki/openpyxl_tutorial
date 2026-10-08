# Moduł 26 — Architektura: openpyxl jako adapter wyjściowy, nie centrum wszechświata

> **Część:** IV — Wzorce projektowe i architektura · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–25

## 0. W tym module nauczysz się

- **Ustawisz openpyxl na właściwym miejscu w systemie** — jako **adapter wyjściowy** w warstwie infrastruktury, a nie jako bibliotekę, od której zależy cały program. Zrozumiesz twardą regułę: `import openpyxl` wolno **tylko** w jednym katalogu — i będziesz umiał to sprawdzić jednym `grep`-em.
- **Zbudujesz port `EksporterRaportu` i trzy adaptery** — `EksporterXlsxOpenpyxl`, `EksporterCsv` i `EksporterFake`. Zobaczysz, że **cały raport z modułu 25 da się przetestować bez openpyxl i bez pliku**, bo domena nie wie, że Excel istnieje.
- **Zbudujesz `SkoroszytRepository` i `JednostkaPracy`** — i zrozumiesz, dlaczego „repozytorium" dla pliku to nie baza danych: ta dyscyplina jest wyłącznie po to, żeby **kontrolować, kiedy plik jest przepisywany** (a to jest jedyna operacja, która może zniszczyć dane — moduł 18).
- **Opanujesz transakcję zapisu end-to-end** — `wczytaj → zmodyfikuj → zapisz do tmp → os.replace → backup`, z obsługą wyjątku w połowie, bez częściowo zapisanego pliku i bez nadpisania oryginału przy błędzie.
- **Wstrzykniesz zależność bez frameworka** — konstruktor przyjmuje `EksporterRaportu`, a w testach `EksporterFake` zapisuje do listy. Zobaczysz, że to daje **testy szybkie, deterministyczne i bez plików** — i że DI to jedna linia, nie kontener.
- **Opiszesz raport deklaratywnie** (dataclass/TOML/YAML) i zobaczysz **dokładnie, gdzie leży granica**: konfiguracja opisuje **co**, kod opisuje **jak**. Dowiesz się, kiedy „język konfiguracji" zamienia się w nowy język programowania i dlaczego trzeba wtedy powiedzieć „stop".
- **Domykasz architekturę**: wersjonowanie i migracje szablonów, arkusz `_meta`, idempotencja i deterministyczność, wydajność jako **decyzja** (synchroniczny vs kolejka vs plik + link), bezpieczeństwo **na wejściu do domeny**, obserwowalność i wzorzec „**raport jako projekcja**".
- **Otrzymasz diagram warstw, przepływ żądania i checklistę 10 pytań**, którą przejdziesz przed każdym wdrożeniem generatora raportów.

## 1. Intuicja i analogia

### 1.1. Kuchnia i kurier, czyli cała architektura w jednym zdaniu

W restauracji są **trzy światy**:

- **Kuchnia** — gotuje. Ma przepisy, składa dania, zna produkt. Kuchnia **nie wie i nie chce wiedzieć**, jaką firmą kurierską pojedzie danie ani czy pudełko jest kartonowe czy plastikowe.
- **Kelner / kierownik sali** — decyduje **kiedy** danie wychodzi i **co** zrobić z zamówieniem: „stolik 4 czeka, a stolik 7 odwołał". To jest orkiestracja.
- **Kurier** — dowozi. Zna pudełka, trasy, terminy. Kurier **nie gotuje** i **nie decyduje**, co jest na talerzu.

Jeżeli kuchnia zacznie projektować pudełka, to:
1. każde nowe pudełko wymaga przebudowy kuchni,
2. kucharz przestaje gotować, bo projektuje opakowania,
3. test nowego dania wymaga zamówienia kuriera.

To jest **cała** architektura warstwowa. W naszym kursie:

- **kuchnia** = domena (`Zamowienie`, `pozycje()`, `suma()`, `marza()`, `WierszRaportu`),
- **kierownik sali** = aplikacja (przypadek użycia „wygeneruj raport miesięczny"),
- **kurier** = infrastruktura (`openpyxl`, `xl/media/`, `styles.xml`, `os.replace`).

I **twarda reguła**: w kuchni nie ma pudełek. Jeżeli `import openpyxl` pojawi się w pliku domeny, to znaczy, że kucharz projektuje opakowania — i cały podział przestaje cokolwiek dawać.

### 1.2. Paragon, czyli raport jako projekcja

Robisz zakupy. Istnieje **transakcja** w systemie sklepu: pozycje, ceny, VAT, data, kasa. I istnieje **paragon** — karteczka. Paragon jest **wyświetleniem transakcji**. Jeżeli zgubisz paragon, transakcja nadal istnieje. Jeżeli na paragonie jest błąd, transakcja jest prawdziwa, a karteczka nie.

Nasz raport Excel to **paragon**. Nie magazyn. I z tego wynikają trzy bardzo konkretne konsekwencje, które ta część kursu zamyka:

1. **Excel nie jest źródłem prawdy.** Jeżeli ktoś edytuje raport i przysyła go z powrotem z „poprawkami", to te poprawki są **na paragonie**, a nie w transakcji. Wraca do systemu tylko to, co system wpuści. (Tak, są systemy, w których Excel **jest** interfejsem wejściowym — i wtedy mówimy o tym **wprost** w module 21 i w naszej konfiguracji, z walidacją i sanitacją, a nie „bo to tylko Excel".)
2. **Ten sam wsad → ten sam raport.** Paragon z tej samej transakcji wygląda tak samo. Jeżeli Twój generator daje dwa różne pliki z tych samych danych, to ma **ukryty stan** — czas, `set`, losową kolejność — i nie da się go testować (moduł 20).
3. **Raport jest jednokierunkowy.** Domena → aplikacja → infrastruktura. Nie „domena → Excel → domena", chyba że **świadomie** zbudowałeś drugą ścieżkę na wejściu i **nazwałeś** to importem, z własnym portem i własną walidacją. Nie robimy tego przypadkiem.

### 1.3. Kontakt w ścianie, czyli DI bez frameworka

Wmurowujesz kontakt w ścianę **raz**. Jeżeli chcesz wymienić urządzenie, które do niego podłączasz, po prostu wyciągasz wtyczkę i wkładasz inną. Ale jeżeli **przylutujesz** czajnik do instalacji, to wymiana czajnika oznacza skuwanie tynku.

Klasa, która **sama tworzy** swój `EksporterXlsxOpenpyxl` w konstruktorze, to przylutowany czajnik:

```python
class RaportSprzedazy:                         # ANTYWZORZ
    def __init__(self, dane):
        self._eksporter = EksporterXlsxOpenpyxl()   # ← przylutowane
```

Nie da się jej przetestować bez openpyxl (nie da się podstawić „fałszywego kuriera"), nie da się jej użyć do eksportu CSV, nie da się jej sprawdzić **szybko** w CI. Wystarczy jedno słowo — konstruktor przyjmuje zależność:

```python
class RaportSprzedazy:
    def __init__(self, dane, eksporter: EksporterRaportu):
        self._eksporter = eksporter            # ← wtyczka, nie lut
```

I to jest **cały** wstrzykiwanie zależności. Nie potrzebujesz do tego żadnego frameworka. Potrzebujesz tylko **jednej decyzji projektowej**: „nie tworzę moich zależności, dostaję je". Dobra wiadomość: robisz to już w module 25 — `BazowyRaport.__init__` przyjmuje `format_nazwa`, a `RaportKatalogowy` przyjmuje `eksport_nazwa`. Ten moduł wymieniają **nazwy** (stringi) na **obiekty** i pokazuje, co się dzięki temu zyskuje.

### 1.4. Dwie daty na kartonie, czyli wersjonowanie

Na kartonie mleka są dwie daty: **data produkcji** i **termin przydatności**. Pierwsza mówi, **z czego** to zrobiono, druga — **do kiedy** ma sens. Jeżeli masz karton bez daty, nie wiesz, czy jesteś na dobrej wersji produktu.

Raport ma identyczny problem. Kiedy użytkownik pisze „raport ma złe formaty", musisz umieć odpowiedzieć **bez pytania go o plik**:

- **wersja generatora** (z której wersji kodu powstał),
- **wersja schematu/szablonu** (jakie kolumny i reguły były wtedy aktualne),
- **hash danych wejściowych** (czy raport dotyczy tego samego wsadu),
- **data wygenerowania**, autor, wersja biblioteki.

Wszystko to ląduje w **arkuszu `_meta`** i w `wb.properties` — i to jest **jedyna** rzecz, jaką naprawdę warto dopisać do każdego raportu, jakiego kiedykolwiek wygenerujesz. W tej części kursu robimy to systematycznie.

### 1.5. Jeden pomiar, dwie naprawy, czyli transakcja zapisu

Wyobraź sobie, że przepisujesz z zeszytu do zeszytu **sto stron**. Robisz to tak:

1. bierzesz czysty zeszyt,
2. przepisujesz stronę po stronie,
3. oddajesz do archiwum.

Zauważ, co się dzieje, gdy ktoś wytrąci Ci zeszyt na 60. stronie: masz **nowy zeszyt z 59 stronami** i **stary zeszyt w całości**. Jeżeli zgubisz kartkę przed 60. stroną, orientujesz się natychmiast — bo nowy zeszyt się nie domyka. Nigdy nie dopisujesz do starego zeszytu, bo wtedy zniszczyłbyś jedyną pełną kopię.

`os.replace()` robi dokładnie to samo, tylko na poziomie systemu plików: piszesz **nowy** plik tymczasowy w **tym samym** katalogu, a potem **jednym ruchem** podmieniasz nazwę. Jeżeli pisanie się nie uda — usuwasz tymczasowy i nic się nie stało. Stary plik jest **nietknięty**, bo nigdy go nie otworzyłeś do zapisu. To jest cała transakcyjność, jakiej potrzebujesz w generatorze raportów — i wystarcza ona na 99% przypadków. (Druga ścieżka, `fsync`, dotyczy awarii prądu; o tym powiemy sobie krótko w sekcji 2.4.)

### 1.6. Pułapka deklaratywności, czyli prawo wyrażone w formularzu

Można opisać raport w taki sposób, że **nikt nie pisze kodu**:

```yaml
kolumny:
  - nazwa: "Sprzedaz netto"
    zrodlo: "suma_netto"
    format: '#,##0.00 "zl"'
  - nazwa: "Marza"
    zrodlo: "marza_proc"
    format: "0.0%"
    warunki:
      - gdy: "marza_proc < 0.05"
        tlo: "FCE4E4"
```

To jest **bardzo** kuszące, bo brzmi jak postęp. I bardzo szybko zamienia się w katastrofę, bo po dwóch miesiącach ktoś dopisuje `gdy: "abs(marza_proc - srednia_regionu) > odchylenie_standardowe"` — i mamy **nowy język programowania**, z **jednym** autorem i **zerem** testów. Prawo podatkowe też da się zapisać jako tabelka, ale w końcu trzeba napisać interpretację.

Zasada, którą zapamiętaj, jest jedyną obroną:

> **Konfiguracja opisuje CO. Kod opisuje JAK.**

Konfiguracja: które kolumny, jak nazwane, jakie formaty, które reguły **wybrane z gotowej palety**. Kod: jak liczyć marżę, jak wygląda reguła nietypowa, co się dzieje, gdy nie ma danych. Jeżeli konfiguracja zaczyna **wyrażać obliczenia**, to znak, że to nie konfiguracja — to program, i powinien trafić do kodu, gdzie ma testy i podpowiedzi edytora.

### 1.7. Rura, kubeł i lista, czyli wydajność jako decyzja

Zamawiasz cegły na budowę. Trzy sposoby:

- **Rura**: dostawca podłącza taśmę i cegły jadą na bieżąco. Zaczynasz budować, gdy lecą pierwsze. Ale nie widzisz całej taśmy, nie cofniesz jej i kiedy się kończy, musisz reagować. → **`StreamingResponse`**, strumień z bazy do arkusza.
- **Kubeł**: czekasz, aż przyjadą wszystkie cegły, wtedy budujesz. Wygodnie, widzisz wszystko, możesz poprawić przed budowaniem. Ale kubeł zajmuje miejsce i **masz limit rozmiaru**. → **generowanie w pamięci i `send_file`**.
- **Lista zamówienia i plac budowy**: cegły leżą na placu, a Ty dostajesz **adres**. Idziesz i budujesz, kiedy chcesz. → **plik na dysku + link**, praca w tle, kolejka.

Żadna z tych decyzji **nie dotyczy Excela**. Nie decyduje o niej openpyxl, tylko **twoje ograniczenie**: czas odpowiedzi HTTP, rozmiar pamięci, długość zadania. To jest właśnie architektura: **wybór na podstawie ograniczeń systemu**, nie na podstawie tego, co umie biblioteka. W sekcji 2.11 rozpisujemy ten wybór na tabelę.

### 1.8. Jeden pomysł na cały moduł

Popatrz na to, co zrobiliśmy w poprzednich modułach i co robimy tutaj:

- moduły 03–21 — **API openpyxl**: jak to działa, jak to ugryźć, co się psuje,
- moduły 22–25 — **wzorce**: jak wynieść decyzję i zachowanie do obiektów,
- moduł 26 — **architektura**: **gdzie** te obiekty leżą i **kto ma prawo kogo wołać**.

Jednym zdaniem: **architektura to granice plus kierunek zależności.** Wszystko dalej w tym module jest tego jedną konsekwencją:

> **Zależności płyną do środka.** Domena nie zna nikogo. Aplikacja zna domenę. Infrastruktura zna oba. **Nigdy odwrotnie.**

## 2. Teoria

### 2.1. Trzy warstwy — czym się różnią, nie „co zawierają"

Większość opisów warstw myli, bo wymienia **zawartość** („domena ma modele, aplikacja ma serwisy"). Ale zawartość to kwestia gustu. Różnicą **prawdziwą** jest **reguła zależności i odpowiedzialność za zmianę**.

| Warstwa | Co wie | Czego nie wie | Co ją zmienia |
|---|---|---|---|
| **Domena** | reguły biznesowe, pojęcia: zamówienie, marża, region, próg | nic o plikach, HTTP, bazach, Excelu | **zmiana reguły biznesowej** |
| **Aplikacja** | jak **poprowadzić** przypadek użycia: kiedy domena jest gotowa, kiedy zapisać | nic o formacie pliku i nic o SQL | **zmiana procesu** (nowy krok, nowa walidacja biznesowa) |
| **Infrastruktura** | jak **dowieźć**: openpyxl, CSV, HTTP, `os.replace`, SMTP | **nic o regułach biznesowych** | **zmiana technologii** (nowy format, nowy kanał) |

Test jest mechaniczny i bezlitosny. Zadaj sobie kolejno trzy pytania o każdy plik:

1. **Czy zmieni się, gdy zmieni się formuła marży, VAT albo próg „dużej sprzedaży"?** → **domena**.
2. **Czy zmieni się, gdy zmieni się kolejność kroków w procesie generowania, ale nie reguły i nie format?** → **aplikacja**.
3. **Czy zmieni się, gdy zmieni się biblioteka, format albo miejsce zapisu?** → **infrastruktura**.

I teraz najważniejsze: **plik, który odpowiada „tak" na więcej niż jedno z tych pytań, jest źle umieszczony.** To jest ten mechanizm, który łapie 90% realnych błędów — nie intuicja, nie diagram, tylko trzy pytania o **powód przyszłej zmiany**.

#### Reguła zależności

W kodzie: domena **nie importuje nic** ze swojego projektu poza innymi plikami domeny. Aplikacja importuje domenę. Infrastruktura importuje oba. To jest **jednokierunkowy graf**. Sprawdzenie:

```bash
# Test poziomu 1: zadna warstwa wyzej nie wola warstwy nizej (poza Infra -> wszystko)
grep -rn "openpyxl" src/
grep -rn "from xlsx_writer\|from openpyxl" src/domain/ src/application/
# Oba musza byc PUSTE. Pierwszy wolno miec wynik TYLKO w src/infrastructure/.
```

### 2.2. Ports & Adapters — „port" to nie buzzword, to kontrakt

**Intuicja.** Gniazdko elektryczne. Urządzenie **nie zna elektrowni**. Zna **kształt gniazdka** — i to jest cały kontrakt.

**Analogia.** Przejściówka podróżna (moduł 24, sekcja 2.2) — tylko z drugiej strony: tu **urządzenie** deklaruje, jakiego gniazdka potrzebuje, a to Ty dobierasz przejściówkę.

**Definicja.** **Port** to **kontrakt** (interfejs), w którym nasz kod deklaruje, czego potrzebuje od świata zewnętrznego — napisany w języku **domeny**, nie technologii. **Adapter** to implementacja tego kontraktu dla konkretnej technologii.

Port `EksporterRaportu` — kluczowa własność: **domeny nie ma w jego sygnaturze, ale też nie ma openpyxl.**

```python
from typing import Protocol

class EksporterRaportu(Protocol):
    """PORT wyjsciowy. Kontrakt miedzy aplikacja a infrastruktura.

    Celowo NIE MA tu: Workbook, Worksheet, Cell, save, BytesIO, Path.
    Jest: raport (model), wynik (nasz DTO), sciezka docelowa.
    """
    def eksportuj(self, raport: "ModelRaportu", cel: str) -> "WynikEksportu": ...
```

Trzy porty, które w praktyce wystarczą w każdym generatorze raportów:

| Port | Kontrakt | Kto implementuje |
|---|---|---|
| **`EksporterRaportu`** (wyjście) | „weź model, zapisz pod celem, zwróć wynik" | Xlsx, Csv, Pdf, Fake, Http |
| **`ZrodloDanych`** (wejście) | „daj mi wsad po identyfikatorze/okresie" | baza, API, pliki, `ZrodloFake` |
| **`MagazynRaportow`** (magazyn) | „znajdź wcześniejszy raport, zapisz wynik, wersjonuj" | katalog plików, S3, baza |

I uwaga na **typ zwracany**: `WynikEksportu` to **nasz** DTO. Nie `Workbook`, nie `BytesIO`, nie `Path` z openpyxl. To, co zwraca port, **musi być zrozumiałe bez openpyxl**, inaczej port przecieka (pułapka 1).

#### Dlaczego Ports & Adapters wygrywa w praktyce

Cztery korzyści, które są **weryfikowalne**, a nie deklaratywne:

1. **Testy bez plików.** `EksporterFake` zapisuje model do listy. Test domeny i aplikacji nie wie, że istnieje Excel — i trwa **mikrosekundy**, nie sekundy.
2. **Wymiana formatu bez dotykania logiki.** `EksporterCsv` vs `EksporterXlsxOpenpyxl` — wybór to jedna linia w kompozycji (sekcja 2.5).
3. **Dwie implementacje → uczciwy kontrakt.** Kiedy port ma tylko jeden adapter, kontrakt jest **przypadkowy** (bo pisany pod tę jedną technologię). Druga implementacja to natychmiast obnaża: „o, mój port wymaga, żeby ktoś znał `BytesIO`".
4. **Wydajność i bezpieczeństwo stają się decyzjami** — bo zmieniają się w jednym miejscu (sekcje 2.11, 2.12).

### 2.3. Repository i Unit of Work — uwaga na nadinterpretację

**Intuicja.** Biblioteka. Wypożyczasz książkę, pracujesz z nią, oddajesz. Biblioteka pilnuje, **kiedy** książka wraca na półkę.

**Uczciwe ostrzeżenie, które trzeba postawić na początku:** w klasycznym sensie „Repository" to **wzorzec z baz danych**, gdzie jest sesja, transakcja i odwzorowanie obiektowo-relacyjne. **Dla pliku `.xlsx` repo nie ma transakcji** (bo openpyxl jej nie ma), nie ma zapytań, nie ma tożsamości wierszy. Jeżeli przeniesiesz do swojego kodu cały ciężar „prawdziwego" repository, dostaniesz **gruby wrapper nad `load_workbook` + `save`** i nic więcej, plus fałszywe poczucie, że to jest baza.

Dlatego piszemy to jasno:

> **`SkoroszytRepository` istnieje w naszym projekcie z JEDNEGO powodu: żeby kontrolować, KIEDY plik jest przepisywany.** Bo to jest jedyna operacja, która może zniszczyć dane (moduł 03, 18). Cały sens to: `save` w jednym miejscu, atomowo, z backupem i z logiem.

Jak to wygląda w praktyce:

```python
class SkoroszytRepository(Protocol):
    def wczytaj(self, identyfikator: str) -> "SkoroszytModelu": ...
    def zapisz(self, skoroszyt: "SkoroszytModelu", *, wersja: str) -> "Sciezka": ...
    def usun(self, identyfikator: str) -> None: ...
```

I **`JednostkaPracy`** — odpowiedzialna za to, żeby **na końcu było jedno zapisanie**, nie pięć. Zbierasz zmiany, a dopiero gdy wszystko jest spójne, **commitujesz**:

```text
uow = JednostkaPracy(repo)
uow.wczytaj("szablon_raportu_2026.xlsx")
uow.zmodyfikuj(...)   # zmiany w modelu w pamieci
uow.zmodyfikuj(...)
uow.zatwierdz( wersja="3" )   # JEDEN zapis, atomowo, z backupem
```

**Kiedy to jest nadmiar?** Jeżeli generujesz raport **od zera** i zapisujesz **raz** — nie potrzebujesz ani `SkoroszytRepository`, ani `JednostkaPracy`. Sam `EksporterRaportu` z metodą `eksportuj()` już jest jednostką pracy (jeden zapis na końcu). Repository i UoW pojawiają się dopiero, gdy:

1. **wczytujesz istniejący plik i modyfikujesz go** (moduł 18 — pełne ryzyko utraty),
2. **pracujesz na szablonach** i musisz rozróżnić „szablon" od „wygenerowanego raportu",
3. **zapisujesz kilka skoroszytów jako jedną operację logiczną** (pakiet raportów — wszystkie albo żaden),
4. **wersjonujesz** i musisz wiedzieć, z której wersji szablonu powstał plik.

### 2.4. Transakcyjność i atomowość — pełna sekwencja

To jest sekcja, którą traktuj jak checklistę. Kolejność jest obowiązkowa.

```text
1. ZAMIAR        | sprawdz, ze katalog docelowy istnieje i jest zapisywalny
2. ODCZYT        | jesli modyfikujesz istniejacy plik: wczytaj do PAMIECI (nigdy "w miejscu")
3. MODYFIKACJA   | zmiany w modelu / przez StosKomend (modul 25); zero zapisu do pliku
4. WALIDACJA     | sprawdz wynik PRZED zapisem (metryki, podglad, reguly domeny)
5. ZAPIS TMP     | zapisz do pliku tymczasowego w TYM SAMYM katalogu (mkstemp)
6. WERYFIKACJA   | otworz plik tymczasowy, sprawdz metryki (liczba arkuszy, wierszy)
7. BACKUP        | jesli plik docelowy istnieje: przenies go do .bak z timestampem
8. PODMIANA      | os.replace(tmp, docelowy)  <-- COMMIT, nieodwracalny, jednym ruchem
9. SPRZATANIE     | w razie wyjatku na kazdym kroku: unlink(tmp), przywroc .bak
10. RAPORT       | zwroc wynik (sciezka, bajty, hash, metryki); zapisz audyt JSON
```

Trzy szczegóły techniczne, które trzeba znać:

**`mkstemp` w tym samym katalogu.** `os.replace()` jest atomowy **tylko na tym samym wolumenie**. Jeżeli plik tymczasowy powstanie w `/tmp`, a docelowy jest na innym dysku (albo w innym kontenerze, albo na montowanym wolumenie), to `os.replace()` może **skopiować** plik zamiast podmienić nazwę — i wtedy atomowości nie ma. Dlatego `tempfile.mkstemp(dir=sciezka.parent)`.

**`os.fsync` przed podmianą.** Na wypadek awarii prądu: samo zamknięcie pliku nie gwarantuje, że dane trafiły na dysk. Jeżeli zależy Ci na tym poziomie trwałości (rzadkie w generatorze raportów, częste w systemach rozliczeniowych), dopisz `os.fsync(fd)` **przed** `os.replace()`. Powiedz to użytkownikowi kodu w docstringu — bo bez tego ktoś doda tę linię jako „optymalizację" i usunie, nie wiedząc, czemu tam była.

**Backup przed podmianą, nie po.** Kolejność `backup → replace` chroni przed utratą w sytuacji, gdy podmiana się nie uda. Kolejność `replace → backup` nie chroni przed niczym, bo „backup" jest już kopią **nowego** pliku.

**Czego NIE daje ta transakcja, a co ludzie obiecują:**

- **Brak blokady pliku.** Jeżeli użytkownik ma plik docelowy otwarty w Excelu, `os.replace()` na Windows da `PermissionError`. Obsłuż to **jawnie** (pułapka 4) albo powiedz użytkownikowi „zamknij plik".
- **Brak izolacji.** Dwie instancje aplikacji piszące ten sam raport się nie zobaczą. Jeżeli to możliwe, nazwy plików muszą być **unikalne per uruchomienie** (timestamp, UUID w nazwie tymczasowej).
- **Brak gwarancji metadanych.** Uprawnienia pliku tymczasowego (`mkstemp` tworzy z 0600) nie są automatyczne kopiowane z oryginału. Jeżeli to istotne, ustaw je jawnie albo skopiuj (`shutil.copymode`).

### 2.5. Wstrzykiwanie zależności — bez frameworka, bez ceremonii

Reguły są trzy i mieszczą się w trzech zdaniach:

1. **Klasa nie tworzy swoich zależności.** Dostaje je.
2. **Zależności deklaruje się w konstruktorze**, z adnotacją typu (kontrakt = port), nie w ciele metodach.
3. **Kompozycja na brzegu.** Jedyne miejsce, które **wybiera** implementacje, to `main()` / `cli.py` / `app.py` — jedno miejsce w całym projekcie.

To ostatnie jest kluczowe i najczęściej pomijane. Twój kod może mieć pięćdziesiąt klas z DI — ale jeżeli w **trzydziestu** miejscach ktoś pisze `EksporterXlsxOpenpyxl(...)`, to „zostawiłeś sobie wtyczkę, ale przylutowałeś przewód w trzydziestu punktach". Gruntowna kompozycja to **jeden** plik:

```python
def kompozycja(argv: list[str]) -> int:
    # JEDYNE miejsce, ktore zna konkretne implementacje
    format_wyjscia = argv[0] if argv else "xlsx"
    eksporter: EksporterRaportu = {
        "xlsx": lambda: EksporterXlsxOpenpyxl(),
        "csv": lambda: EksporterCsv(),
        "fake": lambda: EksporterFake(),
    }[format_wyjscia]()

    zrodlo = ZrodloCsv(Path("data/orders.csv"))
    przypadek = GenerujRaportMiesieczny(zrodlo=zrodlo, eksporter=eksporter)
    wynik = przypadek.wykonaj(okres="2026-09", cel=f"output/raport_{format_wyjscia}")
    print(f"OK: {wynik.sciezka} ({wynik.bajtow} B, {wynik.wierszy} wierszy)")
    return 0
```

W testach **wywołujesz ten sam przypadek użycia**, tylko podstawiasz `EksporterFake`. I to jest cała wartość DI: **jeden kod produkcyjny, ten sam testowany** — a nie dwie ścieżki, które „na pewno robią to samo".

### 2.6. Konfiguracja deklaratywna — cztery poziomy i granica

Zamiast od razu skakać w YAML, rozróżnij **cztery** poziomy. W praktyce większość projektów powinna zatrzymać się na drugim.

| Poziom | Gdzie jest definicja raportu | Kiedy wystarcza | Koszt |
|---|---|---|---|
| **1. Kod** | w klasie `RaportSprzedazy._naglowek()` (moduł 25) | 1–3 raporty, jeden zespół | nowy raport = nowy plik `.py`, plus wdrożenie |
| **2. Dataclass** | `RaportSpec` w kodzie Pythona | wiele raportów, ten sam zespół, typowanie i podpowiedzi | nowy raport = nowy obiekt w kodzie; brak wdrożenia dla prostych zmian |
| **3. TOML/YAML** | plik obok kodu | analityk/spEC zmienia kolumny i formaty **bez** wdrożenia | trzeba walidować plik, obsłużyć błędy, wersjonować |
| **4. „DSL"** | wyrażenia, warunki, zależności, pętle w konfiguracji | — **ostrzeżenie** | **to jest nowy język** i wtedy się zatrzymaj |

Reguła graniczna, którą warto powtarzać jak mantrę:

> **Konfiguracja wybiera z gotowej palety. Kod tworzy nowe rzeczy.**

Konfiguracja **może** powiedzieć: „kolumna `Nazwa` idzie na `klient_nazwa`, format `#,##0.00 "zl"`, reguła `ujemny_na_czerwono`". Konfiguracja **nie może** mówić: „jeżeli region jest w bazie VIP, a marża mniejsza niż trzykrotność mediany z poprzedniego kwartału, to czerwone". Bo to jest reguła biznesowa — ma być w domenie, z testem, z autorem i z możliwością debaty o niej.

I jeszcze jedno, bardzo praktyczne: **nie wymyślaj własnego formatu**. TOML albo YAML, z biblioteką, którą wszyscy znają. Ścieżka `tomllib` (Python 3.11+; `tomli` dla starszych) dla plików „tylko odczyt" jest wygodna, bo parsuje bez zależności zewnętrznej. YAML wymaga `pyyaml` i ma zwodniczą składnię (wcięcia, `yes`/`no` jako boole, wersje 1.1 vs 1.2) — używaj go, gdy **naprawdę** potrzebujesz komentarzy i ludzkiej edycji.

### 2.7. Wersjonowanie i migracje

Trzy wersje, które trzeba rozróżniać — i wszystkie trzy umieścić w raporcie:

| Wersja | Co znaczy | Gdzie |
|---|---|---|
| **Wersja generatora** | która wersja **kodu** produkcji raport (np. `git describe --tags`) | `wb.properties.description` + `_meta` |
| **Wersja schematu** | jaki był **kontrakt kolumn/reguł** (np. `raport_v3`) | `_meta`, nazwa arkusza `_meta` |
| **Wersja danych** | jaki był **wsad** (hash, liczba wierszy, zakres dat) | `_meta`, osobny plik sidecar JSON |

Wersjonowanie schematu ma dwa poziomy:

1. **Nazwy plików z datą** (`raport_2026-09_v3.xlsx`) — trywialne, ale skuteczne, gdy chodzi tylko o to, żeby nadpisać tylko własny plik, a nie cudzy.
2. **Polityka deprecjacji** — jeżeli szablony mają wersje, to musisz odpowiedzieć na pytania: **ile wersji obsługuje nowy kod** i **co robisz ze starym szablonem**. Bez tej polityki projektujesz tak, jakby wymagania nigdy się nie zmieniały — a wiesz, że się zmieniają.

Wzorzec migracji, który się sprawdza:

```text
nowy kod czyta stary szablon:
   1. sprawdz _meta.schemat
   2. jesli schemat nieznany -> BladProtokolu z jasnym komunikatem (NIE ciche zgadywanie)
   3. jesli schemat starszy, ale w polityce obslugi -> migracja (nowy plik wynikowy!)
   4. NIGDY nie nadpisuj starego szablonu w miejscu
```

**Kluczowe zdanie tego podpunktu:** migracja szablonu to **nowy plik**, nigdy podmiana „w miejscu". Jeżeli migracja uda się w 80%, zostaniesz z szablonem, który nie jest ani starą, ani nową wersją. To jest najgorszy możliwy stan — dokładnie ten, przed którym ostrzegał moduł 18.

### 2.8. Idempotencja i deterministyczność

**Analogia.** Kalkulator: `2 + 2` daje `4` za każdym razem. Jeżeli wciskasz `=` pięć razy i dostajesz `4, 5, 6, 7, 8`, to nie jest kalkulator — to generator liczb losowych. Twój generator raportów ma być kalkulatorem.

**Idempotencja** — jeżeli uruchomisz skrypt **dwa razy** na tych samych danych, nie dostaniesz podwojonych wierszy i nie „znajdziesz" nieistniejących zmian. To jest wymóg modułu 18 i to jest wymóg tutaj.

**Deterministyczność** — **te same** dane wejściowe dają **ten sam** plik (dopuszczamy wyłącznie *znaczniki czasu*, takie jak `created`/`modified`, których z definicji nie da się wymusić).

Cztery realne źródła niedeterminizmu, wszystkie w Pythonie i wszystkie do naprawienia:

| Źródło | Objaw | Naprawa |
|---|---|---|
| iterowanie po `set`/`dict` bez klucza | kolejność wierszy lub kolumn „raz taka, raz taka" | `sorted(...)`, klucz sortowania jawny |
| `datetime.now()` wewnątrz logiki | plik zmienia się przy każdym uruchomieniu | **wstrzyknij czas** jako parametr |
| kolejność z `glob()`/`os.listdir()` | nieuporządkowane wejście | `sorted()` zawsze |
| DOMYŚLNE sortowanie tekstu zależne od locale | „Ł" sortuje się raz tak, raz tak | `key=lambda s: s.casefold()` |

Dobra wiadomość: **żadne z tych źródeł nie jest po stronie openpyxl**. To są pułapki Twojego kodu — i wszystkie cztery da się znaleźć **testem**, który generuje plik dwa razy i porównuje go po `podsumowanie_skoroszytu()` (moduł 20). To jest ten sam test regresji, przeniesiony na poziom architektury.

### 2.9. Bezpieczeństwo jako warstwa

To jest **powtórka modułu 21 w jednym zdaniu**, ale umieszczona architektonicznie:

> **Sanitacja odbywa się na WEJŚCIU DO DOMENY, nie przy zapisie.**

Dlaczego to ma znaczenie:

- Jeżeli walidujesz w `EksporterXlsxOpenpyxl`, to **druga ścieżka** (`EksporterCsv`, HTTP API) jest niezabezpieczona. Nie dlatego, że ktoś to zrobił źle, ale dlatego, **że o sanitacji pomyślał w jednym miejscu**.
- Jeżeli walidujesz w domenie (funkcja `przyjmij_tekst_zewnetrzny()`), to **każda** ścieżka, która wchodzi do domeny, jest chroniona przez **tę jedną** funkcję.
- Test jest wtedy mechaniczny: `grep -n "def przyjmij_tekst_zewnetrzny" src/domain/` — jedno miejsce. A **każde** wejście do domeny przez port `ZrodloDanych` **musi** przez nie przejść.

Praktyczna wskazówka: **w domenie trzymaj stringi jako stringi**. Jeżeli na wejściu odrzucasz kontrolnie `=`, `+`, `-`, `@`, tabulator i CR — a robisz to **raz**, przy przyjęciu danych — to eksporter openpyxl nie ma już nic do roboty. To jest różnica między „walidowaniem na wyjściu" (kruche, zależne od liczby ścieżek wyjścia) a „walidowaniem na wejściu" (odporne, jedno miejsce zmiany).

### 2.10. Obserwowalność

Raport, którego nie umiesz zdiagnozować po fakcie, jest raportem, któremu **nie ufasz**. A raportowi, któremu nie ufasz, ludzie przestają używać — i wracają do ręcznego Excela. Dlatego obserwowalność **nie jest** opcją.

Cztery minimalne elementy:

**1. Log strukturalny na zdarzenie** — jedno zdarzenie na raport, z kompletem faktów (moduł 25 mówił o metrykach; tu je zbieramy w jedno miejsce):

```text
raport=raport_miesieczny okres=2026-09 wierszy=18430 arkuszy=3 czas_s=4.7
  ostrzezenia=2 bajtow=1204881 wynik=ok
ostrzezenia=["region 'PL' zmapowany z 'polska' (312 wierszy)",
             "plik wejsciowy zawieral sparkline'y - utracone przy zapisie (modul 18)"]
```

Zauważ, że **ostrzeżenia są częścią wyniku**, a nie komunikatem na `print`. Bo wynik idzie do logu, do `_meta` i do odpowiedzi HTTP. W module 25 nazbliżaliśmy się do tego: `metryki` na strategiach, `bledy` na autobusie, `opis` na komendach. Tu to **składamy w jedno**.

**2. Metryki liczbowe** — format nadaje się pod monitoring: `raport_wierszy_total`, `raport_czas_s`, `raport_ostrzezen_total`, `raport_eksportow_total{format="xlsx"}` (jeśli używasz Prometheus/OTLP, odpowiednio `counter`/`histogram`).

**3. Utrata funkcji pliku jako zdarzenie alertujące.** To jest wniosek modułu 18, ale **na poziomie systemu**: jeżeli wczytany plik zawierał `xl/charts/*` albo `xl/drawings/*` albo ostrzeżenie openpyxl o nieznanej części rozszerzenia — to jest **zdarzenie**, które **musi** trafić do logu z odpowiednim poziomem. Bo w tym momencie tracisz dane, o których istnieniu być może nie wiedziałeś. Alert = przyjrzenie się temu plikowi.

**4. Ślad audytowy w pliku.** Arkusz `_meta` — bo log na serwerze może zniknąć, a plik wędruje. W `_meta` wpisujesz: wersję generatora, wersję schematu, hash wsadu, datę, autora, liczbę wierszy, ostrzeżenia. Dzięki temu pytanie „skąd ten plik i czy jest kompletny?" ma odpowiedź **w samym pliku**.

### 2.11. Wydajność jako decyzja architektoniczna

To nie jest decyzja „o openpyxl". To jest decyzja o tym, **gdzie raport jest generowany i jak trafia do użytkownika**. Trzy wzorce, każdy z własną ceną:

| Wzorzec | Kiedy | Zalety | Cena |
|---|---|---|---|
| **Synchroniczny w request/response** | mały raport (< 10 tys. wierszy), interaktywny użytkownik | natychmiastowa odpowiedź, brak infrastruktury | blokuje worker; limit czasu odpowiedzi; retry = ponowne przeliczenie |
| **Strumień z `BytesIO`** | średni raport, chcemy odpowiedź HTTP bez pliku pośredniego | brak pliku tymczasowego, brak sprzątania, brak wycieku danych do katalogu | cały plik w pamięci; brak retry od miejsca przerwania |
| **Zadanie w tle + plik + link** | duży raport, długie generowanie, wielu odbiorców | nie blokuje, można retry, można wysłać e-mail, plik do audytu | trzeba kolejki, statusu, sprzątania, polityki retencji |

Trzy techniczne szczegóły, które trzeba znać przy strumieniu HTTP:

```python
# FastAPI: strumien bez pliku posredniego
from fastapi.responses import StreamingResponse
from io import BytesIO
from openpyxl import Workbook

def raport_http() -> StreamingResponse:
    bufor = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.append(["Data", "Kwota"])
    # ... wypelnij ...
    wb.save(bufor)                       # save() przyjmuje obiekt plikopodobny!
    bufor.seek(0)                        # ← OBOWIAZKOWO: wskaznik na poczatek
    return StreamingResponse(
        bufor,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="raport.xlsx"'},
    )

# Flask: send_file z BytesIO
from flask import send_file
def raport_flask():
    bufor = BytesIO()
    # ... zbuduj i wb.save(bufor) ...
    bufor.seek(0)                        # ← OBOWIAZKOWO
    return send_file(
        bufor,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="raport.xlsx",
    )
```

**Dwa punkty, o których ludzie zapominają:**

- **`seek(0)`.** Bez tego odpowiedź jest pusta, bo wskaźnik stoi na końcu. Objaw: „plik się pobiera, ale ma 0 bajtów" — i to jest pułapka 7.
- **MIME type.** Dla `.xlsx` użyj `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`. `application/octet-stream` też otworzy plik, ale przeglądarki nie zaproponują rozszerzenia/„otwórz w Excelu web".

I jeszcze **limit rozmiaru i cache** — dwa elementy, które w systemie produkcyjnym muszą być **nazwane**:

- **Limit wejścia** (rozmiar pliku, liczba wierszy, liczba arkuszy) — bo bez limitów pierwszy lepszy użytkownik wystawi Ci `MemoryError`.
- **Cache** — jeżeli ten sam raport jest generowany dla 500 użytkowników, klucz cache to `(schemat, hash_wsadu, format)`. Ale **nie wdrażaj cache'a, zanim nie masz metryk** — bo cache to optymalizacja, a optymalizacja bez pomiaru to spekulacja.

### 2.12. Raport jako projekcja — wniosek architektoniczny

Zbierzmy to, co rozsypane po poprzednich punktach:

| Jeśli Excel jest... | Konsekwencja | Gdzie o tym mówiliśmy |
|---|---|---|
| **projekcją** danych z systemu | nigdy nie czytamy go z powrotem jako źródła prawdy | tu (26) |
| **formularzem wejściowym** | musimy mieć port wejściowy, walidację, sanitację, kontrolę wersji | 17, 21, tu |
| **szablonem** | jest prototypem: czytamy go, nie nadpisujemy, zapisujemy **nowy** plik | 18, 23, tu |
| **przechowalnią** danych | to jest błąd projektowy; dane powinny być w bazie | tu (26) |

Jedna decyzja architektoniczna, jedno zdanie:

> **Raport to projekcja. Jest wygenerowany, wysłany, zarchiwizowany i zapomniany. Nie jest stanem.**

Jeżeli w Twoim systemie ludzie edytują raport i przysyłają go z powrotem z „poprawkami", to nie masz generatora raportów — masz **system z Excelem jako interfejsem wejściowym**. To jest uprawniony system, ale wymaga zupełnie innej architektury: port wejściowy, walidacja pól, konflikty wersji, kontrola zmian, wersjonowanie wierszy. Nie buduj go **przypadkiem** — bo skończysz z nerwowym plikiem, który czasem jest raportem, a czasem źródłem prawdy.

### 2.13. Diagram warstw

```text
                    ┌───────────────────────────────────────────────────────────┐
                    │                         BRZEG / KOMPOZYCJA                │
                    │  cli.py  /  app.py (FastAPI)  /  worker.py (kolejka)      │
                    │  ── jedyne miejsce, ktore wybiera implementacje ──        │
                    └───────────────┬───────────────────────────────────────────┘
                                    │ tworzy i przekazuje (DI)
                                    ▼
    ┌───────────────────────────────────────────────────────────────────────────┐
    │  APLIKACJA   src/application/                                             │
    │                                                                            │
    │    GenerujRaportMiesieczny(zrodlo: ZrodloDanych,                           │
    │                           eksporter: EksporterRaportu,                     │
    │                           magazyn: MagazynRaportow)                        │
    │                                                                            │
    │    wykonaj(okres, cel) -> WynikEksportu                                    │
    │       1. dane = zrodlo.pobierz(okres)                                      │
    │       2. model = zbuduj_model_raportu(dane)   <-- ORKIESTRACJA             │
    │       3. wynik = eksporter.eksportuj(model, cel)                           │
    │       4. magazyn.zapisz_metadane(wynik)                                    │
    └───────────────────────────────┬───────────────────────────────────────────┘
                                    │ import (w dol: aplikacja -> domena)
                                    ▼
    ┌───────────────────────────────────────────────────────────────────────────┐
    │  DOMENA   src/domain/                     (ZERO importow zewnetrznych)    │
    │                                                                            │
    │    Zamowienie (dataclass)        ModelRaportu (dataclass)                  │
    │    pozycje(), suma(), marza()    WierszRaportu, Sekcja, Kpi                │
    │    przyjmij_tekst_zewnetrzny()   <-- sanitacja WEJSCIA (modul 21)          │
    │    reguly: duza_sprzedaz(), wymaga_uwagi()                                 │
    │                                                                            │
    │    NIE wie o: openpyxl, plikach, HTTP, YAML, CSV                           │
    └───────────────────────────────────────────────────────────────────────────┘
                                    ▲
                                    │ implementuje porty (w gore: infra -> aplikacja)
                                    │
    ┌───────────────────────────────┴───────────────────────────────────────────┐
    │  INFRASTRUKTURA   src/infrastructure/       <-- TU WOLNO import openpyxl  │
    │                                                                            │
    │    excel/EksporterXlsxOpenpyxl   (implementuje EksporterRaportu)           │
    │    excel/SkoroszytRepository     (wczytaj/zapisz, atomowo, backup)         │
    │    excel/JednostkaPracy          (jeden zapis na koniec)                   │
    │    csv/EksporterCsv              (implementuje EksporterRaportu)           │
    │    pdf/EksporterPdfLibreOffice   (implementuje EksporterRaportu)           │
    │    fake/EksporterFake            (do TESTOW - tez implementuje port!)      │
    │    zrodla/ZrodloCsv, ZrodloBazy  (implementuja ZrodloDanych)               │
    │    konfig/RaportSpecLoader       (TOML/YAML -> RaportSpec)                 │
    │    log/AudytJson                 (JSON sidecar - nie arkusz!)              │
    └───────────────────────────────────────────────────────────────────────────┘

    PRZEPLYW ZADANIA (przyklad HTTP):
      POST /raporty/miesieczny {okres: "2026-09", format: "xlsx"}
        -> cli/app: kompozycja -> wybierz EksporterXlsxOpenpyxl
        -> application.GenerujRaportMiesieczny.wykonaj()
              -> ZrodloDanych.pobierz(okres)          [infra: CSV/baza]
              -> domain: zbuduj model, policz, zwaliduj
              -> EksporterRaportu.eksportuj(model)    [infra: openpyxl]
                    -> JednostkaPracy: tmp -> weryfikacja -> backup -> os.replace
              -> magazyn: zapisz _meta, audyt JSON
        -> odpowiedz: 200 {plik: "raport_2026-09_v3.xlsx", bajtow: 1204881}
```

Zwróć uwagę na jedną rzecz, która jest sednem całej tej części kursu: **strzałki zależności idą w dół i do środka, ale wywołania w czasie wykonania idą na zewnątrz** (aplikacja woła port, którego implementacja leży w infrastrukturze). To jest **odwrócenie zależności**: kod, który „wywołuje" (aplikacja), nie zależy od kodu, który „robi" (infrastruktura). Zależy tylko od **kontraktu**. Dzięki temu testy domeny i aplikacji nie potrzebują openpyxl — nie „dają radę" bez niego, ale **nie mają z nim nic wspólnego**.

### 2.14. Checklista przeglądu architektury

Dziesięć pytań. Jeżeli odpowiesz „nie" na którekolwiek, nie wdrażaj — bo te „nie" zamienią się w awarie w najgorszym momencie.

1. **Czy mogę zmienić format wyjścia bez dotykania domeny?** (Test: `grep -r "openpyxl" src/domain/ src/application/` → pusto.)
2. **Czy mogę przetestować domenę bez plików i bez openpyxl?** (Test: `pytest tests/domain/` na maszynie bez `openpyxl` — przechodzi.)
3. **Czy wiem, co dokładnie stracę, jeśli plik wejściowy był edytowany w Excelu?** (Test: `scan-before-save` z modułu 18 uruchomiony na **każdym** pliku szablonu.)
4. **Czy zapis jest atomowy i nie nadpisuje oryginału przy błędzie?** (Test: zasymuluj wyjątek po zapisaniu tmp; docelowy plik musi być nietknięty.)
5. **Czy wiem, jak zachowa się system, gdy użytkownik ma plik otwarty w Excelu?** (Test: otwórz w Excelu i spróbuj zapisać — masz jawny komunikat czy `PermissionError` bez kontekstu?)
6. **Czy te same dane dają ten sam plik poza znacznikami czasu?** (Test: wygeneruj dwa razy, porównaj `podsumowanie_skoroszytu()`.)
7. **Czy wiem, jaki jest budżet wydajnościowy i co się dzieje, gdy zostanie przekroczony?** (Test: 200k wierszy — mieści się w limicie; jeżeli nie, wiem, że trzeba kolejki.)
8. **Czy sanitacja jest na wejściu do domeny, a nie przy zapisie?** (Test: `grep -n "HYPERLINK\|data_type" src/` — żaden wynik poza infrastrukturą.)
9. **Czy każdy wygenerowany plik niesie swoją wersję, hash wsadu i ostrzeżenia?** (Test: wczytaj losowy plik z archiwum — odpowiedz na „skąd to i czy jest kompletne?".)
10. **Czy wiem, jak wykryję utratę funkcji pliku (sparkline, makra, shapes)?** (Test: `alert` z modułu 25 podnosi zdarzenie na każdym pliku z `xl/charts`, `xl/drawings`, `x14`.)

## 3. Przykłady krok po kroku

Przykłady są **kumulatywne**. Wszystkie leżą w projekcie `xlsx_raporty/`, dobudowywanym przez cały moduł. Na końcu dostajesz aplikację, którą można uruchomić z CLI i przetestować **bez openpyxl**.

Struktura katalogów, do której dążymy:

```text
xlsx_raporty/
├── src/
│   ├── domain/
│   │   ├── __init__.py
│   │   ├── modele.py            # Zamowienie, ModelRaportu, WierszRaportu
│   │   ├── reguly.py            # marza(), duza_sprzedaz(), przyjmij_tekst_zewnetrzny()
│   │   └── ports.py             # EksporterRaportu, ZrodloDanych, MagazynRaportow
│   ├── application/
│   │   ├── __init__.py
│   │   └── generuj_raport.py    # GenerujRaportMiesieczny
│   └── infrastructure/
│       ├── __init__.py
│       ├── excel/
│       │   ├── eksporter.py     # EksporterXlsxOpenpyxl
│       │   ├── repository.py    # SkoroszytRepository + JednostkaPracy
│       │   └── fasada.py        # RaportExcel z modulu 24 (przeniesione)
│       ├── csv/eksporter.py
│       ├── pdf/eksporter.py
│       ├── fake/eksporter.py
│       ├── zrodla/zrodlo_csv.py
│       └── konfig/spec.py
├── config/
│   └── raport_miesieczny.toml
├── tests/
│   ├── domain/
│   ├── application/
│   └── infrastructure/
└── cli.py
```

### Przykład 1 — `src/domain/modele.py` i `src/domain/reguly.py`: domena bez openpyxl

**Problem.** W module 25 logika „marża", „duża sprzedaż" i „stan poniżej minimum" siedziała w podklasach `BazowyRaport`. To dług, o którym sami pisaliśmy. Spłacamy go.

```python
"""src/domain/modele.py - modele domenowe.

ZERO importow zewnetrznych. Zero openpyxl. Zero plikow.
To jest "kuchnia" - gotuje, nie pakuje.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP


@dataclass(frozen=True)
class Zamowienie:
    """Jedno zamowienie. Kwoty w GROSZACH (int) - unikamy float na kwotach.

    Uzasadnienie (modul 05): float ma 15-17 cyfr znaczacych i potrafi
    zgubic grosz na duzej sumie. Kwota w groszach to jednoznaczny int.
    """

    identyfikator: str
    klient: str
    region: str
    data_zamowienia: date
    ilosc: int              # w sztukach, calkowita
    cena_grosze: int        # cena jednostkowa w groszach
    koszt_grosze: int       # koszt jednostkowy w groszach

    @property
    def netto_grosze(self) -> int:
        return self.ilosc * self.cena_grosze

    @property
    def marza_grosze(self) -> int:
        return self.ilosc * (self.cena_grosze - self.koszt_grosze)

    @property
    def marza_proc(self) -> Decimal:
        """Marza proc z dokladnoscia do promila (3 miejsca).

        Decimal, nie float: wyswietlamy to w raporcie finansowym.
        """
        if self.netto_grosze == 0:
            return Decimal("0")
        iloraz = Decimal(self.marza_grosze) / Decimal(self.netto_grosze)
        return iloraz.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)


@dataclass(frozen=True)
class WierszRaportu:
    """Wiersz w raporcie. Domena NIE wie, ze to pojdzie do komorek."""

    klient: str
    region: str
    ilosc: int
    netto_pln: Decimal
    marza_proc: Decimal
    wymaga_uwagi: bool


@dataclass(frozen=True)
class Kpi:
    etykieta: str
    wartosc: object
    format_hint: str              # "currency", "percent", "int" - NIE kod Excela!


@dataclass(frozen=True)
class ModelRaportu:
    """Zupelny opis tego, co ma byc w raporcie. Bez kolorow, bez arkuszy."""

    tytul: str
    okres: str                    # "2026-09"
    naglowki: tuple[str, ...]
    wiersze: tuple[WierszRaportu, ...]
    kpi: tuple[Kpi, ...]
    ostrzezenia: tuple[str, ...] = ()
    schemat: str = "raport_v1"
    hash_wsadu: str = ""
    wygenerowano: str = ""
```

```python
"""src/domain/reguly.py - reguly biznesowe i sanitacja WEJSCIA.

UWAGA: to jest jedyne miejsce, w ktorym odrzucamy podejrzane znaki.
Nie robimy tego przy zapisie do Excela! (modul 21 + sekcja 2.9)
"""

from __future__ import annotations

from decimal import Decimal
from typing import Iterable

from .modele import Kpi, ModelRaportu, WierszRaportu, Zamowienie


class BladDomeny(Exception):
    """Blad reguly biznesowej. Nie jest bledem technicznym."""


# ======================================================================
# SANITACJA WEJSCIA - jedno miejsce, kazda sciezka przez nie przechodzi
# ======================================================================
ZNAKI_WYZWALAJACE = ("=", "+", "-", "@", "\t", "\r")


def przyjmij_tekst_zewnetrzny(wartosc: object, *, pole: str) -> str:
    """Przyjmuje tekst do domeny, chroniac przed formula injection.

    Trzy zasady:
      1. Odcinamy znaki kontrolne.
      2. ODRZUCAMY tekst zaczynajacy sie od znaku wyzwalajacego - nie 'naprawiamy'.
      3. Ograniczamy dlugosc komorki (Excel: 32767 znakow, modul 05).

    Odrzucenie zamiast modyfikacji jest celowe: lepiej jawnie rzucic blad
    niz po cichu zmienic dane klienta.
    """
    if wartosc is None:
        return ""
    tekst = str(wartosc)
    if len(tekst) > 32767:
        raise BladDomeny(f"Pole {pole!r}: tekst dluzszy niz 32767 znakow")
    if tekst and tekst[0] in ZNAKI_WYZWALAJACE:
        raise BladDomeny(
            f"Pole {pole!r}: tekst zaczyna sie od znaku wyzwalajacego "
            f"{tekst[0]!r}. Ten tekst wykonalby sie jako formula w Excelu. "
            "Odrzucamy - nie modyfikujemy danych klienta."
        )
    return "".join(ch for ch in tekst if ch == " " or ch.isprintable()).strip()


# ======================================================================
# REGULY BIZNESOWE
# ======================================================================
def marza_proc(zamowienia: Iterable[Zamowienie]) -> Decimal:
    """Marza zbiorcza = suma marz / suma netto. Nie srednia marz!"""
    netto = sum(z.netto_grosze for z in zamowienia)
    marza = sum(z.marza_grosze for z in zamowienia)
    if netto == 0:
        return Decimal("0")
    return (Decimal(marza) / Decimal(netto)).quantize(Decimal("0.001"))


def wymaga_uwagi(zamowienie: Zamowienie, *, prog_marzy: Decimal) -> bool:
    """Regula biznesowa. Zmienia sie tu, NIE w kodzie Excela."""
    return zamowienie.marza_proc < prog_marzy


def zbuduj_model_raportu(
    zamowienia: list[Zamowienie],
    *,
    okres: str,
    prog_marzy: Decimal = Decimal("0.05"),
    tytul: str = "Raport sprzedazy",
    schemat: str = "raport_v1",
    hash_wsadu: str = "",
    wygenerowano: str = "",
) -> tuple[ModelRaportu, tuple[str, ...]]:
    """Domena buduje model raportu. Zwraca (model, ostrzezenia)."""
    if not zamowienia:
        raise BladDomeny(f"Brak zamowien w okresie {okres!r}")

    wiersze: list[WierszRaportu] = []
    for z in sorted(zamowienia, key=lambda x: (x.region, x.klient, x.identyfikator)):
        wiersze.append(
            WierszRaportu(
                klient=z.klient,
                region=z.region,
                ilosc=z.ilosc,
                netto_pln=(Decimal(z.netto_grosze) / 100).quantize(Decimal("0.01")),
                marza_proc=z.marza_proc,
                wymaga_uwagi=wymaga_uwagi(z, prog_marzy=prog_marzy),
            )
        )

    netto = (Decimal(sum(z.netto_grosze for z in zamowienia)) / 100).quantize(
        Decimal("0.01")
    )
    kpi = (
        Kpi("Pozycji ogolem", len(zamowienia), "int"),
        Kpi("Sprzedaz netto", netto, "currency"),
        Kpi("Marza zbiorcza", marza_proc(zamowienia), "percent"),
        Kpi(
            "Pozycji do przegladu",
            sum(1 for w in wiersze if w.wymaga_uwagi),
            "int",
        ),
    )

    model = ModelRaportu(
        tytul=tytul,
        okres=okres,
        naglowki=("Klient", "Region", "Ilosc", "Sprzedaz netto", "Marza"),
        wiersze=tuple(wiersze),
        kpi=kpi,
        schemat=schemat,
        hash_wsadu=hash_wsadu,
        wygenerowano=wygenerowano,
    )
    return model, ()
```

**Co się dzieje w pamięci.** `ModelRaportu`, `WierszRaportu`, `Kpi` i `Zamowienie` są `frozen=True` — niemutowalne (moduł 24, sekcja 2.1). To znaczy, że raz zbudowany model jest **bezpieczny do przekazywania przez warstwy**: infrastruktura go nie zmieni przypadkiem, a jeśli spróbuje, dostanie `FrozenInstanceError` **natychmiast**, w miejscu błędu — a nie „cichą modyfikację, która gdzieś się objawi".

`Decimal` w kwotach zamiast `float` — bo kwoty trafiają do raportu finansowego, a `float` na sumie 50 000 pozycji potrafi zgubić grosz. To kosztuje trochę wydajności i to jest **świadoma** decyzja architektoniczna: za poprawność płacimy kilkoma procentami czasu.

**Co trafi do pliku.** Na tym etapie **nic**. Ten plik nie ma ani jednego wywołania openpyxl i **nie może** go mieć. To jest jego cała wartość.

### Przykład 2 — `src/domain/ports.py`: porty

```python
"""src/domain/ports.py - PORTY. Kontrakty miedzy warstwami.

Uzywamy typing.Protocol (Python 3.8+), a nie ABC, bo:
  - implementacje nie musza dziedziczyc (wiec moga byc klasami z innej biblioteki),
  - testy nie musza tworzyc podklas "na niby",
  - mypy sprawdza zgodnosc STRUKTURALNIE (metoda o tej nazwie = zgodne).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from .modele import ModelRaportu


@dataclass(frozen=True)
class WynikEksportu:
    """DTO wyniku. Nasz typ. ZERO openpyxl, ZERO Path z openpyxl.

    `sciezka` jest stringiem albo None (gdy zapisujemy do pamieci/strumienia).
    """
    format: str                 # "xlsx" | "csv" | "pdf" | "fake"
    bajtow: int
    arkuszy: int
    wierszy: int
    sciezka: str | None
    hash_wyniku: str = ""
    ostrzezenia: tuple[str, ...] = ()


class EksporterRaportu(Protocol):
    """PORT WYJSCIOWY. To, czego aplikacja potrzebuje od 'swiata zewnetrznego'."""

    @property
    def format(self) -> str: ...

    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu: ...


class ZrodloDanych(Protocol):
    """PORT WEJSCIOWY. Skad bierzemy wsad."""

    def pobierz(self, okres: str) -> tuple[list, str]:
        """Zwraca (lista_zamowien, hash_wsadu)."""
        ...


class MagazynRaportow(Protocol):
    """PORT MAGAZYNU. Metadane, audyt, historia."""

    def zapisz_metadane(self, wynik: WynikEksportu, meta: dict) -> None: ...
    def znajdz(self, okres: str, format: str) -> list[dict]: ...
```

**Dlaczego `Protocol`, a nie `ABC`.** Bo port to **kształt**, a nie dziedziczenie. `EksporterCsv` nie musi dziedziczyć po klasie z domeny — wystarczy, że ma metodę `eksportuj` i właściwość `format`. Dzięki temu:
- implementacja nie zależy od klasy portu (luźne sprzężenie),
- testy mogą „udawać" port jednym `dataclass` z metodą — bez deklarowania dziedziczenia,
- `mypy` sprawdzi zgodność **strukturalnie**, czyli tak, jak naprawdę działa Python.

**Uwaga o `Protocol` i `runtime_checkable`.** `isinstance(x, EksporterRaportu)` zadziała tylko z dekoratorem `@runtime_checkable`. Ale **nie używaj** go do sprawdzania zgodności w produkcji — sprawdzaj **raz**, w kompozycji (`assert`, `raise`), a nie w każdej pętli. Sprawdzanie typu w każdej iteracji to strata czasu i zamaskowany problem projektowy: jeżeli musisz pytać „czy to na pewno eksporter", to znaczy, że ktoś przekazuje coś, czego nie powinien.

### Przykład 3 — `src/infrastructure/fake/eksporter.py`: `EksporterFake` (najważniejszy plik infrastruktury)

To jest plik, który wygląda na mało ważny, a **jest** najważniejszy w całym module.

```python
"""src/infrastructure/fake/eksporter.py - EksporterFake.

Nie zapisuje pliku. Zapisuje MODEL do listy. Dzieki temu caly przypadek
uzycia mozna przetestowac bez openpyxl i bez dysku - w mikrosekundach.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field

from src.domain.modele import ModelRaportu
from src.domain.ports import WynikEksportu


@dataclass
class ZapisFake:
    """Co 'zapisano'. Model + cel. Bez openpyxl."""
    model: ModelRaportu
    cel: str


@dataclass
class EksporterFake:
    """Implementuje EksporterRaportu STRUKTURALNIE (nie dziedziczy)."""

    _zapisy: list[ZapisFake] = field(default_factory=list)
    _awaria: Exception | None = None
    _format: str = "fake"

    # --- konfiguracja do testow ---
    def ustaw_awarie(self, blad: Exception) -> "EksporterFake":
        self._awaria = blad
        return self

    # --- EksporterRaportu ---
    @property
    def format(self) -> str:
        return self._format

    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu:
        if self._awaria is not None:
            blad, self._awaria = self._awaria, None     # awaria jednorazowa
            raise blad

        self._zapisy.append(ZapisFake(model=raport, cel=cel))

        # "bajty" liczymy z reprezentacji modelu - deterministycznie
        surowiec = json.dumps(
            {
                "tytul": raport.tytul,
                "okres": raport.okres,
                "naglowki": list(raport.naglowki),
                "wiersze": [
                    {"klient": w.klient, "region": w.region, "ilosc": w.ilosc,
                     "netto": str(w.netto_pln), "marza": str(w.marza_proc),
                     "uwaga": w.wymaga_uwagi}
                    for w in raport.wiersze
                ],
                "kpi": [(k.etykieta, str(k.wartosc), k.format_hint) for k in raport.kpi],
            },
            sort_keys=True,
            ensure_ascii=False,
        ).encode("utf-8")

        return WynikEksportu(
            format=self._format,
            bajtow=len(surowiec),
            arkuszy=1,
            wierszy=len(raport.wiersze),
            sciezka=None,                      # nic nie zapisalismy na dysk
            hash_wyniku=hashlib.sha256(surowiec).hexdigest()[:16],
            ostrzezenia=raport.ostrzezenia,
        )

    # --- akcesory dla testow (nie sa czescia portu!) ---
    def zapisy(self) -> tuple[ZapisFake, ...]:
        return tuple(self._zapisy)

    def ostatni(self) -> ZapisFake:
        if not self._zapisy:
            raise AssertionError("EksporterFake: nic nie zapisano")
        return self._zapisy[-1]
```

**Co się dzieje w pamięci.** `EksporterFake` trzyma **listę zapisów** — czyli pełną historię tego, co aplikacja przekazała. W teście możesz zapytać nie tylko „czy coś zapisano", ale też „**co dokładnie** zapisano w drugim wywołaniu". To jest moc tego wzorca: **podstawiasz kamerę zamiast kuriera** i widzisz przesyłkę.

**`_awaria` do symulowania błędu.** Metoda `ustaw_awarie()` pozwala przetestować błędy zapisu **bez** udawania się do systemu plików. To jest jedyny sensowny sposób testowania „co się dzieje, gdy zapis padnie" — wersja „otwórz plik w Excelu, żeby zablokować" nie działa w CI.

**Co trafi do pliku.** **Nic.** I to jest punkt. Cały przypadek użycia, który za chwilę zbudujesz, będzie testowany właśnie przez ten plik.

### Przykład 4 — `src/application/generuj_raport.py`: przypadek użycia

```python
"""src/application/generuj_raport.py - ORKIESTRACJA.

Ten plik nie wie, jaki format pliku powstaje ani jak dziala openpyxl.
Wie tylko: pobierz dane, zbuduj model, wyeksportuj, zapisz metadane.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from datetime import datetime, timezone

from src.domain.modele import ModelRaportu
from src.domain.ports import EksporterRaportu, MagazynRaportow, WynikEksportu, ZrodloDanych
from src.domain.reguly import BladDomeny, zbuduj_model_raportu

log = logging.getLogger("raporty.app")


@dataclass(frozen=True)
class WynikUzycia:
    """Co zwraca przypadek uzycia. DTO aplikacji."""
    wynik: WynikEksportu
    czas_s: float
    okres: str
    schemat: str
    hash_wsadu: str


class GenerujRaportMiesieczny:
    """ZALEZNOSCI PRZEZ KONSTRUKTOR. Trzy porty, zero importow z openpyxl."""

    def __init__(
        self,
        *,
        zrodlo: ZrodloDanych,
        eksporter: EksporterRaportu,
        magazyn: MagazynRaportow | None = None,
        zegar=None,
    ) -> None:
        self._zrodlo = zrodlo
        self._eksporter = eksporter
        self._magazyn = magazyn
        # ZEGAR WSTRZYKNIETY - deterministycznosc (sekcja 2.8)!
        self._teraz = zegar or (lambda: datetime.now(timezone.utc))

    def wykonaj(self, *, okres: str, cel: str) -> WynikUzycia:
        start = time.perf_counter()
        log.info("raport start okres=%s format=%s", okres, self._eksporter.format)

        zamowienia, hash_wsadu = self._zrodlo.pobierz(okres)

        model, ostrzezenia = zbuduj_model_raportu(
            zamowienia,
            okres=okres,
            hash_wsadu=hash_wsadu,
            wygenerowano=self._teraz().isoformat(timespec="seconds"),
        )

        wynik = self._eksporter.eksportuj(model, cel)
        czas = time.perf_counter() - start

        # JEDNO zdarzenie logu z kompletnym kompletem faktow (sekcja 2.10)
        log.info(
            "raport koniec okres=%s format=%s wierszy=%s arkuszy=%s bajtow=%s "
            "czas_s=%.2f ostrzezen=%s wynik=%s",
            okres, wynik.format, wynik.wierszy, wynik.arkuszy, wynik.bajtow,
            czas, len(ostrzezenia) + len(wynik.ostrzezenia), "ok",
        )
        for ostrzezenie in (*ostrzezenia, *wynik.ostrzezenia):
            log.warning("ostrzezenie okres=%s tresc=%s", okres, ostrzezenie)

        if self._magazyn is not None:
            self._magazyn.zapisz_metadane(wynik, {
                "okres": okres,
                "schemat": model.schemat,
                "hash_wsadu": hash_wsadu,
                "wygenerowano": model.wygenerowano,
                "wersja_generatora": model.schemat.rsplit("_", 1)[-1],
                "wierszy": wynik.wierszy,
                "ostrzezenia": list(ostrzezenia) + list(wynik.ostrzezenia),
            })

        return WynikUzycia(
            wynik=wynik, czas_s=czas, okres=okres,
            schemat=model.schemat, hash_wsadu=hash_wsadu,
        )
```

**Co się dzieje w pamięci.** Ten plik **nie ma ani jednego** importu openpyxl i **nigdy nie powinien** mieć. Nie wie nawet, że coś takiego jak openpyxl istnieje — wie tylko, że istnieje `EksporterRaportu`, który „weźmie model i coś z nim zrobi". To jest **właściwa** warstwa aplikacji: orkiestruje, liczy czas, loguje, deleguje.

**Zegar jest wstrzyknięty.** `zegar=lambda: datetime(...)` to jedna linia, a kupuje **testowalność deterministyczną**: w teście podajesz `zegar=lambda: datetime(2026, 10, 1, tzinfo=timezone.utc)` i wiesz, co jest w `_meta` bez mockowania modułu. To jest dokładnie ten sam mechanizm, o którym mówiliśmy w module 20 („testy zależne od zegara") — tu widać, jak wygląda jego **architektoniczne** rozwiązanie.

**Co trafi do pliku.** Nic bezpośrednio — ale to jest miejsce, w którym **zaczyna się** cały proces, i miejsce, w którym **kończy się** jego wynik. Wszystkie warstwy poniżej są już „techniczne".

### Przykład 5 — `src/infrastructure/excel/repository.py`: `SkoroszytRepository` + `JednostkaPracy`

```python
"""src/infrastructure/excel/repository.py - atomowy zapis.

Tu wolno uzywac openpyxl i os.replace. Tu JEST ta jedna, jedyna operacja,
ktora moze zniszczyc plik. Dlatego cala dyscyplina jest tutaj.
"""

from __future__ import annotations

import hashlib
import logging
import os
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

log = logging.getLogger("raporty.infra.excel")


class BladZapisu(Exception):
    """Blad infrastruktury zapisu. Zawiera kontekst, nie tylko komunikat."""


def _hash_pliku(sciezka: Path) -> str:
    skrot = hashlib.sha256()
    with sciezka.open("rb") as plik:
        for porcja in iter(lambda: plik.read(65536), b""):
            skrot.update(porcja)
    return skrot.hexdigest()[:16]


@dataclass
class JednostkaPracy:
    """Zbiera zmiany i ZAPISUJE RAZ, atomowo.

    Kolejnosc (sekcja 2.4) - nie zmieniaj jej bez zrozumienia:
      1. sprawdz katalog docelowy
      2. zapisz do tmp w TYM SAMYM katalogu
      3. weryfikuj tmp (otworz i policz arkusze)
      4. backup oryginalu (jezeli jest)
      5. os.replace - COMMIT
      6. sprzatanie zawsze, takze przy wyjatku
    """

    _operacje: list = field(default_factory=list)
    _story: list = field(default_factory=list)

    def dodaj(self, opis: str, operacja) -> "JednostkaPracy":
        """Operacja to callable(wb) -> None. Zmiany w PAMIECI, nie na dysku."""
        self._operacje.append((opis, operacja))
        return self

    def zatwierdz(self, *, wb, cel: str | Path, wersja: str = "1") -> Path:
        cel = Path(cel)
        cel.parent.mkdir(parents=True, exist_ok=True)

        for opis, operacja in self._operacje:
            try:
                operacja(wb)
                self._story.append(opis)
            except Exception as blad:
                raise BladZapisu(
                    f"Operacja {opis!r} nie udala sie w pamieci "
                    f"({type(blad).__name__}: {blad}). Plik na dysku NIE zmieniony."
                ) from blad

        # 2. TMP W TYM SAMYM KATALOGU - atomowosc os.replace
        fd, tmp_name = tempfile.mkstemp(
            prefix=f".{cel.stem}.", suffix=".tmp.xlsx", dir=cel.parent
        )
        os.close(fd)
        tmp = Path(tmp_name)

        try:
            wb.save(tmp)                                # faktyczny zapis

            # 3. WERYFIKACJA - otworz i policz (modul 20, test regresji)
            sprawdzony = load_workbook(tmp, read_only=True)
            arkuszy = len(sprawdzony.sheetnames)
            sprawdzony.close()
            if arkuszy != len(wb.sheetnames):
                raise BladZapisu(
                    f"Weryfikacja tmp: oczekiwano {len(wb.sheetnames)} arkuszy, "
                    f"w pliku jest {arkuszy}. Nie podmieniamy."
                )

            # 4. BACKUP - PRZED podmiana, nigdy po
            if cel.exists():
                znacznik = datetime.now().strftime("%Y%m%d-%H%M%S")
                backup = cel.with_name(f"{cel.stem}.{znacznik}.bak{cel.suffix}")
                shutil.copy2(cel, backup)
                log.info("backup %s -> %s", cel.name, backup.name)

            # 5. COMMIT - jeden ruch, atomowo na tym samym wolumenie
            os.replace(tmp, cel)
            log.info(
                "zapis ok cel=%s wersja=%s bajtow=%s hash=%s",
                cel.name, wersja, cel.stat().st_size, _hash_pliku(cel),
            )
            return cel

        except Exception as blad:
            # 6. SPRZATANIE - tmp NIGDY nie zostaje
            if tmp.exists():
                tmp.unlink(missing_ok=True)
            if isinstance(blad, BladZapisu):
                raise
            raise BladZapisu(
                f"Zapis do {cel.name} nie udal sie "
                f"({type(blad).__name__}: {blad}). Plik docelowy NIE zmieniony."
            ) from blad
```

**Co się dzieje w pamięci.** `JednostkaPracy` najpierw wykonuje **wszystkie** operacje w pamięci (na `wb`), zapisując, co się udało — i gdy którakolwiek padnie, **rzuca bez dotykania dysku**. Zwróć uwagę na ten szczegół: **operacje pamięciowe** są pierwsze, **dyskowe** są drugie. To nie jest przypadkowa kolejność, to jest **transakcja**: albo całość się powiedzie i wtedy piszemy, albo pada przed zapisem i **nic się nie stało**.

**Weryfikacja tmp przez ponowne otwarcie.** Ten krok jest bardzo konkretny: `load_workbook(tmp, read_only=True)` i porównanie liczby arkuszy. To sprawdza, że plik **naprawdę jest czytelny** — bo zdarza się, że `wb.save()` „się uda" i plik jest uszkodzony (np. przy braku miejsca na dysku). **Bez tej weryfikacji** wypchniesz uszkodzony plik i nadpiszesz dobry. Z `read_only=True` otwarcie jest tanie (moduł 19) — nie parsuje wszystkich komórek.

**`read_only=True` i `close()`.** Ten `wb` jest tylko do sprawdzenia liczby arkuszy, więc `read_only` jest właściwym trybem — i `close()` jest **obowiązkowe** (moduł 19, pułapka 1).

**Backup z timestampem, w tym samym katalogu.** `shutil.copy2` (a nie `copy`) zachowuje metadane — czas modyfikacji, uprawnienia. `copy2` przy dużych plikach kopiuje całość, co jest w porządku: backup to nie operacja na ścieżce krytycznej. Ale w systemie z **bardzo** dużymi plikami kopiowanie zajmuje sekundy — dlatego **backup polityki** (ile sztuk trzymamy, czy rolowany `raport.bak`) powinien być ustawieniem, nie stałą.

**Co trafi do pliku.** Docelowy plik z podmienioną zawartością, plik `.bak` z timestampem (jeśli plik istniał wcześniej) i **brak** pliku tymczasowego po zakończeniu. Ten ostatni jest testowalny: `assert not list(cel.parent.glob(".*.tmp.xlsx"))`.

### Przykład 6 — `src/infrastructure/excel/eksporter.py`: `EksporterXlsxOpenpyxl`

To jest adapter. Cała wiedza o openpyxl jest **tutaj** — i to jest ostatni raz, kiedy openpyxl pojawia się w tym module poza tabelą API.

```python
"""src/infrastructure/excel/eksporter.py - adapter xlsx.

To JEDYNY plik w projekcie, ktory tlumaczy model domeny na komorki.
Zasada: format_hint z domeny -> kod formatu Excela. To jest TLUMACZENIE,
nie logika biznesowa.
"""

from __future__ import annotations

import hashlib
import logging
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.domain.modele import ModelRaportu
from src.domain.ports import WynikEksportu
from src.domain.reguly import BladDomeny
from src.infrastructure.excel.repository import JednostkaPracy

log = logging.getLogger("raporty.infra.excel")

# TLUMACZENIE formatow: domena -> kod Excela. W JEDNYM slowniku (modul 10).
FORMATY: dict[str, str] = {
    "currency": '#,##0.00 "zl";(#,##0.00)',
    "percent": "0.0%",
    "int": "#,##0",
    "date": "dd.mm.yyyy",
    "text": "@",
}

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class EksporterXlsxOpenpyxl:
    """ADAPTER. Implementuje EksporterRaportu strukturalnie."""

    def __init__(self, *, autor: str = "generator raportow",
                 wersja_generatora: str = "26.0") -> None:
        self._autor = autor
        self._wersja = wersja_generatora
        self._ostatni_wynik: WynikEksportu | None = None

    @property
    def format(self) -> str:
        return "xlsx"

    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu:
        """Buduje skoroszyt w PAMIECI, ale zapisuje przez JednostkePracy.

        Podzial jest celowy:
          - _zbuduj() jest "czysty" (testowalny: buduje wb w pamieci),
          - JednostkaPracy odpowiada za DYSK (atomowosc, backup, weryfikacja).
        """
        if not raport.wiersze:
            raise BladDomeny(f"Raport {raport.tytul!r} bez wierszy - nie generujemy")

        wb = Workbook()
        wb.remove(wb.active)                        # usun domyslny "Sheet"

        ostrzezenia: list[str] = list(raport.ostrzezenia)

        self._arkusz_dane(wb, raport)
        self._arkusz_podsumowanie(wb, raport)
        self._arkusz_meta(wb, raport, ostrzezenia)

        wb.properties.title = raport.tytul
        wb.properties.creator = self._autor
        wb.properties.description = (
            f"schemat={raport.schemat} okres={raport.okres} "
            f"generator={self._wersja} wsad={raport.hash_wsadu}"
        )
        wb.properties.keywords = f"schemat:{raport.schemat}"

        # DOMYSLNY ARKUSZ = Dane (uzytkownik otwiera plik i widzi dane)
        wb.active = 0

        jednostka = JednostkaPracy()
        jednostka.dodaj("zapis do pliku", lambda _: None)   # wlasciwy zapis ponizej

        sciezka = Path(cel).with_suffix(".xlsx")
        # JednostkaPracy wykonuje zapis; tu przekazujemy GOTOWY wb.
        self._zapisz_atomowo(wb, sciezka)

        wynik = WynikEksportu(
            format="xlsx",
            bajtow=sciezka.stat().st_size,
            arkuszy=len(wb.sheetnames),
            wierszy=len(raport.wiersze),
            sciezka=str(sciezka),
            hash_wyniku=self._hash(sciezka),
            ostrzezenia=tuple(ostrzezenia),
        )
        self._ostatni_wynik = wynik
        return wynik

    # ------------------------------------------------------------------
    # BUDOWA ARKUSZY - "czysta" czesc adaptera
    # ------------------------------------------------------------------
    def _arkusz_dane(self, wb, raport: ModelRaportu):
        ws = wb.create_sheet("Dane")
        ws.sheet_view.showGridLines = False

        # naglowek
        for kolumna, naglowek in enumerate(raport.naglowki, start=1):
            cell = ws.cell(row=1, column=kolumna, value=naglowek)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(fill_type="solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 22

        # wiersze danych
        for indeks, wiersz in enumerate(raport.wiersze, start=2):
            wartosci = (
                wiersz.klient, wiersz.region, wiersz.ilosc,
                float(wiersz.netto_pln), float(wiersz.marza_proc),
            )
            for kolumna, wartosc in enumerate(wartosci, start=1):
                ws.cell(row=indeks, column=kolumna, value=wartosc)

            # formaty: int, currency, percent (modul 10)
            ws.cell(row=indeks, column=3).number_format = FORMATY["int"]
            ws.cell(row=indeks, column=4).number_format = FORMATY["currency"]
            ws.cell(row=indeks, column=5).number_format = FORMATY["percent"]

        ostatni = len(raport.wiersze) + 1
        szerokosci = (30, 10, 10, 16, 10)
        for kolumna, szerokosc in enumerate(szerokosci, start=1):
            ws.column_dimensions[get_column_letter(kolumna)].width = szerokosc

        # REGULA WARUNKOWA: wiersz "wymaga uwagi" -> tlo na calej pozycji.
        # Formuła bez "=" (modul 13), adres $E2 (kolumna ABS, wiersz WZGL.).
        regula = FormulaRule(
            formula=["$E2<0.05"],
            fill=PatternFill(fill_type="solid", bgColor="FCE4E4"),
            font=Font(color="9C0006"),
        )
        ws.conditional_formatting.add(f"A2:E{ostatni}", regula)

        ws.auto_filter.ref = f"A1:E{ostatni}"
        ws.print_title_rows = "1:1"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        return ws

    def _arkusz_podsumowanie(self, wb, raport: ModelRaportu):
        ws = wb.create_sheet("Podsumowanie")
        ws.cell(row=1, column=1, value=raport.tytul).font = Font(bold=True, size=14)
        ws.cell(row=2, column=1, value=f"Okres: {raport.okres}")

        wiersz = 4
        for kpi in raport.kpi:
            ws.cell(row=wiersz, column=1, value=kpi.etykieta).font = Font(bold=True)
            cell = ws.cell(row=wiersz, column=2, value=(
                float(kpi.wartosc) if kpi.format_hint in ("currency", "percent")
                else kpi.wartosc
            ))
            cell.number_format = FORMATY[kpi.format_hint]
            wiersz += 1

        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 18
        return ws

    def _arkusz_meta(self, wb, raport: ModelRaportu, ostrzezenia: list[str]):
        """Karton z dwiema datami (sekcja 1.4)."""
        ws = wb.create_sheet("_meta")
        wiersze = [
            ("schemat", raport.schemat),
            ("wygenerowano", raport.wygenerowano),
            ("generator", self._wersja),
            ("autor", self._autor),
            ("okres", raport.okres),
            ("hash_wsadu", raport.hash_wsadu),
            ("wierszy", len(raport.wiersze)),
            ("ostrzezen", len(ostrzezenia)),
        ]
        for i, (etykieta, wartosc) in enumerate(wiersze, start=1):
            ws.cell(row=i, column=1, value=etykieta).font = Font(bold=True)
            ws.cell(row=i, column=2, value=wartosc)
        for i, tekst in enumerate(ostrzezenia, start=len(wiersze) + 2):
            ws.cell(row=i, column=1, value=tekst)
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 48
        ws.sheet_state = "hidden"     # uzytkownik nie musi tego widziec
        return ws

    # ------------------------------------------------------------------
    # DYSK - delegujemy do JednostkiPracy (sekcja 2.4)
    # ------------------------------------------------------------------
    @staticmethod
    def _zapisz_atomowo(wb, cel: Path) -> Path:
        from src.infrastructure.excel.repository import JednostkaPracy

        return JednostkaPracy().zatwierdz(wb=wb, cel=cel)

    @staticmethod
    def _hash(sciezka: Path) -> str:
        skrot = hashlib.sha256()
        with sciezka.open("rb") as plik:
            for porcja in iter(lambda: plik.read(65536), b""):
                skrot.update(porcja)
        return skrot.hexdigest()[:16]
```

**Co się dzieje w pamięci.** Adapter buduje **trzy arkusze**: `Dane`, `Podsumowanie`, `_meta`. Trzeci jest **ukryty** (`sheet_state = "hidden"`) — bo ma służyć Tobie, a nie użytkownikowi. Zwróć uwagę na `wb.remove(wb.active)`: domyślny `Sheet` z openpyxl **jest usuwany**, bo ma zostać zastąpiony przez nasze `Dane` — bez tego mielibyśmy w pliku cztery arkusze, a nie trzy, i lista arkuszy nie zgadzałaby się z wynikiem `WynikEksportu`.

**Co trafi do pliku.** `output/raport_2026-09.xlsx` z trzema arkuszami:
- `Dane` — nagłówek z niebieskim tłem, freeze panes, autofilter, formaty walutowe i procentowe, reguła warunkowa `$E2<0.05` na tle `FCE4E4`, druk A4 poziomo z powtarzanym nagłówkiem,
- `Podsumowanie` — tytuł i cztery wiersze KPI,
- `_meta` — ukryty arkusz z wersją, hashem i ostrzeżeniami.

**Trzy rzeczy do przemyślenia:**

1. **`format_hint` w domenie, `FORMATY` w infrastrukturze.** Domena mówi „to jest percent", infrastruktura tłumaczy: `"0.0%"`. To jest **właściwy** podział: domena **opisuje intencję**, infrastruktura **wybiera reprezentację**. Jeżeli zmieniasz format waluty na `#,##0.00 €`, zmieniasz **jeden słownik** — a domena nie drgnie.
2. **`_zapisz_atomowo` jest `@staticmethod`** — bo nie korzysta ze stanu adaptera. I to jest dobry sygnał: operacja dysku jest **niezależna** od tego, co buduje adapter. Dwa różne obowiązki, dwa różne obiekty.
3. **Weryfikacja pliku przez `load_workbook` w `JednostkaPracy`.** Zwróć uwagę, że adapter **nie weryfikuje** — to rola jednostki pracy. Podział obowiązków: adapter wie, **jak** zbudować plik, jednostka pracy wie, **jak bezpiecznie** go zapisać. Dobrze to widać w tym, że `JednostkaPracy` mogłaby obsłużyć **każdy** skoroszyt, nie tylko ten konkretny raport.

### Przykład 7 — `tests/application/test_generuj_raport.py`: test bez openpyxl

```python
"""tests/application/test_generuj_raport.py - testy BEZ openpyxl, BEZ plikow.

Uruchomienie tego testu NIE wymaga zainstalowanego openpyxl ani Excela.
"""

from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal

import pytest

from src.application.generuj_raport import GenerujRaportMiesieczny
from src.domain.modele import Zamowienie
from src.domain.reguly import BladDomeny, przyjmij_tekst_zewnetrzny
from src.infrastructure.fake.eksporter import EksporterFake


class ZrodloFake:
    """Implementuje ZrodloDanych strukturalnie. Bez bazy, bez plikow."""

    def __init__(self, zamowienia: list[Zamowienie], hash_wsadu: str = "abc123"):
        self._zamowienia = zamowienia
        self._hash = hash_wsadu

    def pobierz(self, okres: str):
        if okres == "1900-01":
            return [], self._hash
        return list(self._zamowienia), self._hash


ZEGAR_STALY = lambda: datetime(2026, 10, 1, 12, 0, tzinfo=timezone.utc)


def zamowienie(klient="Alfa", region="PL", ilosc=10,
               cena=2500, koszt=2000) -> Zamowienie:
    return Zamowienie(
        identyfikator=f"{klient}-1", klient=klient, region=region,
        data_zamowienia=date(2026, 9, 15), ilosc=ilosc,
        cena_grosze=cena, koszt_grosze=koszt,
    )


@pytest.fixture
def przypadek():
    eksporter = EksporterFake()
    zrodlo = ZrodloFake([zamowienie(), zamowienie("Beta", "DE", 4, 4150, 3000)])
    return GenerujRaportMiesieczny(
        zrodlo=zrodlo, eksporter=eksporter, zegar=ZEGAR_STALY
    ), eksporter


# ----------------------------------------------------------------------
# 1. DOMENA jest testowana w izolacji - bez portow, bez infrastruktury
# ----------------------------------------------------------------------
def test_sanitacja_odrzuca_formule():
    with pytest.raises(BladDomeny, match="znak wyzwalajacy"):
        przyjmij_tekst_zewnetrzny("=HYPERLINK(\"http://zly\")", pole="klient")


@pytest.mark.parametrize("znak", ["=", "+", "-", "@"])
def test_sanitacja_odrzuca_kazdy_znak_wyzwaajacy(znak):
    with pytest.raises(BladDomeny):
        przyjmij_tekst_zewnetrzny(f"{znak}cos", pole="klient")


def test_marza_zbiorcza_to_iloraz_sum_nie_srednia_marz():
    from src.domain.reguly import marza_proc

    male = zamowienie("A", ilosc=1, cena=10000, koszt=5000)      # marza 0.5
    duze = zamowienie("B", ilosc=100, cena=1000, koszt=990)      # marza 0.01
    # srednia marz = 0.255, iloraz sum = (5000+1000)/(10000+100000) = 0.054
    assert marza_proc([male, duze]) == Decimal("0.054")


# ----------------------------------------------------------------------
# 2. APLIKACJA jest testowana przez EksporterFake
# ----------------------------------------------------------------------
def test_wynik_zawiera_metryki(przypadek):
    przypadek_uzycia, eksporter = przypadek
    wynik = przypadek_uzycia.wykonaj(okres="2026-09", cel="output/x")
    assert wynik.wynik.wierszy == 2
    assert wynik.wynik.format == "fake"
    assert wynik.wynik.bajtow > 0
    assert wynik.hash_wsadu == "abc123"


def test_model_trafia_do_eksportera(przypadek):
    przypadek_uzycia, eksporter = przypadek
    przypadek_uzycia.wykonaj(okres="2026-09", cel="output/x")
    zapis = eksporter.ostatni()
    assert zapis.model.tytul == "Raport sprzedazy"
    assert zapis.model.ewiersze[0].klient == "Alfa"


def test_determinizm_dwa_uruchomienia_daja_ten_sam_hash(przypadek):
    przypadek_uzycia, eksporter = przypadek
    a = przypadek_uzycia.wykonaj(okres="2026-09", cel="output/x")
    b = przypadek_uzycia.wykonaj(okres="2026-09", cel="output/x")
    assert a.wynik.hash_wyniku == b.wynik.hash_wyniku


def test_awaria_eksportera_jest_zglaszana(przypadek):
    przypadek_uzycia, eksporter = przypadek
    eksporter.ustaw_awarie(PermissionError("plik otwarty w Excelu"))
    with pytest.raises(PermissionError, match="Excelu"):
        przypadek_uzycia.wykonaj(okres="2026-09", cel="output/x")


def test_brak_danych_to_blad_domeny(przypadek):
    przypadek_uzycia, _ = przypadek
    with pytest.raises(BladDomeny):
        przypadek_uzycia.wykonaj(okres="1900-01", cel="output/x")
```

**Co się dzieje w pamięci.** Ten plik nie importuje **ani jednego** modułu `openpyxl` i nie tworzy **żadnego** pliku. Cała logika — metryki, model, hash, obsługa awarii — jest testowana na `EksporterFake`. Czas uruchomienia: milisekundy. Można go uruchomić w CI, w Dockerze **bez** openpyxl i bez LibreOffice.

**Uwaga o literówce w teście.** W `test_model_trafia_do_eksportera` używam `zapis.model.ewiersze` — to literówka, **celowo** oznaczona jako coś, co `mypy` i test by złapały. W prawdziwym pliku byłoby `zapis.model.wiersze`. Wstawiam ją tutaj jako **demonstrację** jednej rzeczy: testy nie muszą być doskonałe, ale **muszą być uruchomione**.

**Trzy rzeczy do przemyślenia:**

1. **`test_awaria_eksportera_jest_zglaszana` — bez pliku.** Sprawdza, że gdy eksporter zawiedzie, wyjątek **nie jest połknięty**. To jest test z modułu 24 („brama zapisu"), ale tu **naprawdę** można go wykonać szybko.
2. **`ZrodloFake` implementuje port strukturalnie.** Nie dziedziczy po `ZrodloDanych` — bo `Protocol` tego nie wymaga. Trzy linijki kodu, pełnoprawny port.
3. **`ZEGAR_STALY`.** Determinizm daty. Bez tego test `_meta` byłby **niemożliwy**. Ta jedna lambda to cała różnica między „raport, który raz ma datę, a raz inną" a „raport, którego wynik można porównać między uruchomieniami".

### Przykład 8 — `src/infrastructure/konfig/spec.py` + `config/raport_miesieczny.toml`: konfiguracja deklaratywna

```python
"""src/infrastructure/konfig/spec.py - deklaratywny opis raportu.

Zasada (sekcja 2.6): konfiguracja opisuje CO, kod opisuje JAK.
Walidujemy config PRZED uzyciem - nie w trakcie generowania.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

try:
    import tomllib                      # Python 3.11+
except ModuleNotFoundError:              # pragma: no cover
    import tomli as tomllib             # pip install tomli (3.9, 3.10)


class BladKonfiguracji(Exception):
    """Konfiguracja jest niepoprawna. Rzucamy PRZED generowaniem."""


@dataclass(frozen=True)
class KolumnaSpec:
    nazwa: str
    zrodlo: str                          # nazwa pola w modelu domeny
    format: str                          # klucz ze slownika FORMATY
    szerokosc: float = 15.0


@dataclass(frozen=True)
class RaportSpec:
    identyfikator: str
    tytul: str
    schemat: str
    kolumny: tuple[KolumnaSpec, ...]
    prog_marzy: float = 0.05
    druk: str = "landscape"


DOSTEPNE_FORMATY = {"currency", "percent", "int", "date", "text"}
DOSTEPNE_POLA = {"klient", "region", "ilosc", "netto_pln", "marza_proc", "wymaga_uwagi"}


def wczytaj_spec(sciezka: str | Path) -> RaportSpec:
    sciezka = Path(sciezka)
    surowe = tomllib.loads(sciezka.read_text(encoding="utf-8"))

    try:
        identyfikator = surowe["raport"]["identyfikator"]
        tytul = surowe["raport"]["tytul"]
        schemat = surowe["raport"].get("schemat", "raport_v1")
        kolumny_raw = surowe["kolumny"]
    except (KeyError, TypeError) as blad:
        raise BladKonfiguracji(
            f"Brak wymaganego klucza w {sciezka.name}: {blad}. "
            "Wymagane: [raport].identyfikator, [raport].tytul, [[kolumny]]."
        ) from blad

    kolumny: list[KolumnaSpec] = []
    for i, kolumna in enumerate(kolumny_raw, start=1):
        for wymagane in ("nazwa", "zrodlo", "format"):
            if wymagane not in kolumna:
                raise BladKonfiguracji(
                    f"Kolumna {i} nie ma klucza {wymagane!r}"
                )
        if kolumna["format"] not in DOSTEPNE_FORMATY:
            raise BladKonfiguracji(
                f"Kolumna {i}: format {kolumna['format']!r} nieznany. "
                f"Dostepne: {sorted(DOSTEPNE_FORMATY)}"
            )
        if kolumna["zrodlo"] not in DOSTEPNE_POLA:
            raise BladKonfiguracji(
                f"Kolumna {i}: pole {kolumna['zrodlo']!r} nie istnieje w modelu. "
                f"Dostepne: {sorted(DOSTEPNE_POLA)}"
            )
        kolumny.append(KolumnaSpec(
            nazwa=kolumna["nazwa"],
            zrodlo=kolumna["zrodlo"],
            format=kolumna["format"],
            szerokosc=float(kolumna.get("szerokosc", 15.0)),
        ))

    if not kolumny:
        raise BladKonfiguracji("Konfiguracja bez kolumn - raport bylby pusty")

    return RaportSpec(
        identyfikator=identyfikator, tytul=tytul, schemat=schemat,
        kolumny=tuple(kolumny),
        prog_marzy=float(surowe["raport"].get("prog_marzy", 0.05)),
        druk=surowe["raport"].get("druk", "landscape"),
    )
```

```toml
# config/raport_miesieczny.toml
# Konfiguracja opisuje CO - nie wyraza obliczen. Granica z sekcji 2.6.

[raport]
identyfikator = "raport_miesieczny"
tytul = "Raport sprzedazy miesiecznej"
schemat = "raport_v3"
prog_marzy = 0.05
druk = "landscape"

[[kolumny]]
nazwa = "Klient"
zrodlo = "klient"
format = "text"
szerokosc = 30

[[kolumny]]
nazwa = "Region"
zrodlo = "region"
format = "text"
szerokosc = 10

[[kolumny]]
nazwa = "Ilosc"
zrodlo = "ilosc"
format = "int"

[[kolumny]]
nazwa = "Sprzedaz netto"
zrodlo = "netto_pln"
format = "currency"
szerokosc = 16

[[kolumny]]
nazwa = "Marza"
zrodlo = "marza_proc"
format = "percent"
```

**Co się dzieje w pamięci.** `wczytaj_spec` **waliduje wszystko przed** utworzeniem raportu i robi to **zbiorczo**: nieznany format, nieznane pole, brak wymaganego klucza, zero kolumn — każdy błąd ma swój komunikat **z listą dostępnych wartości**. To jest ta sama zasada, co w adapterach z modułu 24: błąd konfiguracji jest **natychmiast czytelny** i mówi, co poprawić.

**Co trafi do pliku.** TOML zmienia to, że **nie wiesz** w kodzie Pythona, które kolumny będą w raporcie — więc `EksporterXlsxOpenpyxl` musi **iterować po `spec.kolumny`**, nie po stałej liście. To jest druga strona deklaratywności: dostajesz elastyczność, ale **płacisz** tym, że część logiki przenosi się do danych. Dlatego walidacja jest tu **tak** szczegółowa — bo bez niej literówka w TOML-u staje się raportem bez kolumny, a nie błędem.

**Trzy rzeczy do przemyślenia:**

1. **`DOSTEPNE_POLA` to zamknięta lista.** Konfiguracja może **wybrać** spośród pól modelu, ale **nie może dodać** nowego pola. Dodanie pola to zmiana kodu (w modelu domeny). To jest **ta** granica z sekcji 2.6, wyrażona w kodzie.
2. **`DOSTEPNE_FORMATY` to zamknięta lista.** Konfiguracja używa kluczy z gotowej palety. Nowy format = nowy wpis w słowniku adaptera — **nie** ciąg znaków z konfiguracji. Gdyby tak było, użytkownik mógłby wpisać **cokolwiek** i Excel dostałby śmieć w `number_format` (moduł 10, pułapka „zły format wyświetlony dosłownie").
3. **`prog_marzy` w TOML.** Tu jest cienka granica: próg jest **parametrem reguły**, a nie **regułą**. Konfiguracja może zmienić próg. Konfiguracja **nie może** zmienić sposobu liczenia ani dodać nowej reguły. To jest świadome: próg jest **decyzją biznesową** zmienianą co kwartał, sposób liczenia marży jest **definicją**, zmienianą świadomie i testem.

### Przykład 9 — `cli.py`: kompozycja na brzegu

```python
"""cli.py - JEDYNE miejsce, w ktorym wybieramy konkretne implementacje.

Caly DI sprowadza sie do tego pliku. Reszta projektu nie tworzy zaleznosci.
"""

from __future__ import annotations

import argparse
import logging
import sys

from src.application.generuj_raport import GenerujRaportMiesieczny
from src.domain.ports import EksporterRaportu
from src.infrastructure.csv.eksporter import EksporterCsv
from src.infrastructure.excel.eksporter import EksporterXlsxOpenpyxl
from src.infrastructure.fake.eksporter import EksporterFake
from src.infrastructure.zrodla.zrodlo_csv import ZrodloCsv


def kompozycja(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generator raportow sprzedazowych")
    parser.add_argument("--okres", required=True, help="np. 2026-09")
    parser.add_argument("--format", default="xlsx", choices=("xlsx", "csv", "fake"))
    parser.add_argument("--wejscie", default="data/orders.csv")
    parser.add_argument("--wyjscie", default="output/raport")
    parser.add_argument("--log", default="INFO")
    args = parser.parse_args(argv)

    logging.basicConfig(level=args.log, format="%(asctime)s %(levelname)s %(message)s")

    # JEDYNE miejsce wyboru implementacji:
    eksportery: dict[str, type] = {
        "xlsx": EksporterXlsxOpenpyxl,
        "csv": EksporterCsv,
        "fake": EksporterFake,
    }
    eksporter: EksporterRaportu = eksportery[args.format]()

    przypadek = GenerujRaportMiesieczny(
        zrodlo=ZrodloCsv(args.wejscie),
        eksporter=eksporter,
    )
    wynik = przypadek.wykonaj(okres=args.okres, cel=args.wyjscie)
    cel = wynik.wynik.sciezka or "(do pamieci)"
    print(f"OK okres={args.okres} format={args.format} "
          f"wierszy={wynik.wynik.wierszy} bajtow={wynik.wynik.bajtow} "
          f"czas={wynik.czas_s:.2f}s cel={cel}")
    return 0


if __name__ == "__main__":
    raise SystemExit(kompozycja(sys.argv[1:]))
```

**Co się dzieje w pamięci.** `cli.py` to **jedyne** miejsce w całym projekcie, w którym widzisz konkretny typ adaptera w słowniku. Wszystkie pozostałe pliki **nie wiedzą**, która implementacja istnieje — widzą tylko porty (`Protocol`). To jest **cała** architektura i jednocześnie cała wartość DI: podmiana implementacji to jedna linia w słowniku `eksportery`.

**Co trafi do pliku.** Zależnie od `--format`: `.xlsx` przez openpyxl, `.csv` przez `csv.writer`, albo **nic** (fake — tylko do tests/dev). Ten sam przypadek użycia, ta sama domena, ta sama logika — **trzy różne kurierów**.

**Trzy rzeczy do przemyślenia:**

1. **Trzy linie `kompozycji`, które decydują o wszystkim.** `ZrodloCsv`, `EksporterXlsxOpenpyxl`, `GenerujRaportMiesieczny`. To jest cały „framework" DI, jakiego potrzebujesz: **funkcja + słownik**.
2. **`--format fake` w produkcji?** Nie do końca. To opcja diagnostyczna: pozwala uruchomić pełny przepływ **bez zapisu** i zobaczyć wynik. Świetna do sprawdzenia, czy problem jest w logice, czy w zapisie.
3. **Ksandry „na brzegu" są dozwolone.** `cli.py` widzi **wszystkie** moduły — i to jest **poprawne**. Granica kompozycji to jedyne miejsce, w którym wolno łączyć warstwy. Test poziomu 1 z sekcji 2.1 **nie obejmuje** `cli.py`, `app.py`, `worker.py` — te są świadomie „brudne".

### Przykład 10 — obserwowalność i raportowanie utraty funkcji

```python
"""src/infrastructure/excel/utrata.py - wykrywanie utraty funkcji pliku.

Zdarzenie "stracimy X przy zapisie" MUSI byc widoczne (sekcja 2.10).
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path

# Czego openpyxl NIE zapisze lub zapisze czesciowo (modul 18):
NIEZAPISYWALNE = {
    "xl/sparklineGroups": "sparkline'y (rozszerzenie x14)",
    "xl/slicers": "slicery",
    "xl/timelines": "timeline'y",
    "xl/pivotCache": "cache tabel przestawnych",
    "xl/model": "model danych (Power Pivot)",
}
CZESCIOWE = {
    "xl/drawings/": "obrazy i kształty (ksztalty moga znikac)",
    "xl/charts/": "wykresy (formatowanie moze byc uproszczone)",
    "xl/vbaProject.bin": "makra (zapisane tylko z keep_vba=True)",
}


@dataclass(frozen=True)
class RaportUtraty:
    sciezka: str
    nieprzepisywalne: tuple[str, ...]
    czesciowe: tuple[str, ...]

    @property
    def ma_alert(self) -> bool:
        return bool(self.nieprzepisywalne)


def przeskanuj_przed_zapisem(sciezka: str | Path) -> RaportUtraty:
    """Skanuje ZIP bez wczytywania calego pliku. Tanio i szybko."""
    sciezka = Path(sciezka)
    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()

    nieprzepisywalne = tuple(
        opis for prefiks, opis in NIEZAPISYWALNE.items()
        if any(n.startswith(prefiks) for n in nazwy)
    )
    czesciowe = tuple(
        opis for prefiks, opis in CZESCIOWE.items()
        if any(n.startswith(prefiks) for n in nazwy)
    )
    return RaportUtraty(str(sciezka), nieprzepisywalne, czesciowe)
```

Użycie w `JednostkaPracy` (fragment, do dopisania w kroku 3 transakcji):

```python
def sprawdz_utrate(sciezka_tmp: Path, ostrzezenia: list[str]) -> None:
    raport = przeskanuj_przed_zapisem(sciezka_tmp)
    if raport.ma_alert:
        ostrzezenia.append(
            "PLIK WEJSCIOWY ZAWIERAL ELEMENTY, KTORE ZOSTANA UTRACONE: "
            + ", ".join(raport.nieprzepisywalne)
        )
    for opis in raport.czesciowe:
        ostrzezenia.append(f"Element odtworzony czesciowo: {opis}")
```

**Co się dzieje w pamięci.** `zipfile.ZipFile` **nie rozpakowuje** pliku — tylko czyta listę nazw z centralnego katalogu ZIP. To jest tanie (mikrosekundy dla plików o rozmiarze 100 MB) i nie zwiększa pamięci. Dlatego skanowanie jest możliwe **przed każdym** zapisem, a nie tylko przy okazji diagnostyki.

**Co trafi do pliku.** Ostrzeżenia trafiają do **`_meta`** (arkusz ukryty) i **do logu** jako zdarzenie `WARNING`. Dzięki temu plik niesie ze sobą informację o tym, co zostało utracone — co jest bardzo praktyczne, bo w momencie gdy plik jest w archiwum, **log może już nie być dostępny**.

**Trzy rzeczy do przemyślenia:**

1. **„Przed zapisem" jest ważne.** Dla plików **odczytywanych**: skanujemy **źródło** (przed `load_workbook`) i porównujemy z tym, co po zapisie. Dla plików **generowanych od zera**: skanujemy tymczasowy plik i sprawdzamy, czy nie ma tam czegoś, czego nie spodziewamy się. W obu przypadkach intencja jest ta sama: **informacja o utracie nie może zniknąć**.
2. **`NIEZAPISYWALNE` vs `CZESCIOWE`.** Ta sama trójdzielność z modułu 03: „potrafi / potrafi częściowo / gubi przy zapisie". Tu jest w **kodzie**, a nie w prozie — bo tylko wtedy **system** może na to zareagować.
3. **`ma_alert` jako osobna właściwość.** Nie każde ostrzeżenie jest alarmem. „Wykres zostanie odbudowany" to `WARNING`. „Sparkline'y znikną" to `ALERT`. Rozróżnienie w kodzie pozwala ustawić **różne poziomy logowania, różne metryki i różne powiadomienia** dla różnych kategorii utraty. To jest obserwowalność robiąca robotę.

## 4. Anatomia API

### Część A — elementy architektury i ich kontrakty

| Element | Co robi | Parametry / kształt | Uwagi |
|---|---|---|---|
| `EksporterRaportu` (`Protocol`) | port wyjściowy | `format` (property), `eksportuj(model, cel) -> WynikEksportu` | **Zero typów openpyxl w sygnaturze**. Test: `grep -n "openpyxl" port.py` → pusto |
| `ZrodloDanych` (`Protocol`) | port wejściowy | `pobierz(okres) -> (lista, hash)` | Jeden punkt wejścia → jedno miejsce sanitacji (sekcja 2.9) |
| `MagazynRaportow` (`Protocol`) | port magazynu | `zapisz_metadane`, `znajdz` | Metadane i audyt, nie sam plik |
| `WynikEksportu` (dataclass) | DTO wyniku | `format, bajtow, arkuszy, wierszy, sciezka, hash_wyniku, ostrzezenia` | **Nasz typ**, nie `Workbook`/`Path` z openpyxl |
| `ModelRaportu`, `WierszRaportu`, `Kpi` | modele domeny | `dataclass(frozen=True)` | Niemutowalne — bezpieczne przechodzenie między warstwami |
| `przyjmij_tekst_zewnetrzny` | sanitacja wejścia | `wartosc, pole` | Jedyne miejsce odrzucania znaków wyzwalających (moduł 21) |
| `zbuduj_model_raportu` | reguła domeny | `zamowienia, okres, ...` | Domena buduje model — nie adapter |
| `GenerujRaportMiesieczny` | przypadek użycia | `__init__(zrodlo, eksporter, magazyn, zegar)` | Zależności przez konstruktor; **wstrzyknięty zegar** dla determinizmu |
| `JednostkaPracy` | transakcja zapisu | `dodaj(opis, op)`, `zatwierdz(wb, cel, wersja)` | Pamięć → tmp → weryfikacja → backup → `os.replace` → sprzątanie |
| `SkoroszytRepository` | dyscyplina dostępu | `wczytaj(id)`, `zapisz(model, wersja)`, `usun(id)` | Kontroluje **kiedy** plik jest przepisywany (sekcja 2.3) |
| `EksporterFake` | adapter testowy | `zapisy()`, `ostatni()`, `ustaw_awarie(blad)` | **Najważniejszy plik infrastruktury** — umożliwia testy bez openpyxl |
| `RaportSpec`, `KolumnaSpec` | model konfiguracji | `dataclass(frozen=True)` | Konfiguracja jako typ — walidowana, wersjonowana |
| `wczytaj_spec` | loader konfiguracji | `sciezka -> RaportSpec` | Waliduje **przed** użyciem; błędy z listą dostępnych wartości |
| `przeskanuj_przed_zapisem` | obserwowalność | `sciezka -> RaportUtraty` | `zipfile.namelist()` — tanie i szybkie (sekcja 2.10) |

### Część B — moduły i obiekty `openpyxl` w tym module

Wszystko poniżej pojawia się **wyłącznie** w `src/infrastructure/excel/`. To jest cała lista i ona **musi** być całą listą — jeżeli w Twoim projekcie openpyxl pojawia się gdzie indziej, to sygnał błędu.

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | nowy skoroszyt w pamięci | — | Bez `wb.active` w sposób zależny od wersji — zawsze jawnie `wb.remove(wb.active)` |
| `wb.remove(wb.active)` | usuwa domyślny `Sheet` | — | Bez tego lista arkuszy ≠ oczekiwana liczba |
| `wb.create_sheet("Dane")` | nowy arkusz o nazwie | `title`, `index` | Kolejność tworzenia = kolejność w pliku |
| `wb.active = 0` | aktywny arkusz po otwarciu | indeks | 0 = pierwszy; użytkownik widzi dane, nie `_meta` |
| `load_workbook(tmp, read_only=True)` | weryfikacja tmp | `read_only` | Tanie otwarcie; **`close()` obowiązkowe** (moduł 19) |
| `ws.freeze_panes = "A2"` | zamrożenie nagłówka | `str` | Formuła pozycji zależy od adresu (moduł 11) |
| `ws.sheet_view.showGridLines = False` | odcięcie siatki | `bool` | Tabela raportowa ma własne obramowanie |
| `ws.sheet_state = "hidden"` | ukrycie arkusza `_meta` | `"hidden"` / `"veryHidden"` | Użytkownik nie widzi, ale plik nosi metadane |
| `ws.print_title_rows = "1:1"` | powtarzanie nagłówka | **format zakresu** | `"1"` **nie zadziała** — moduł 11, pułapka 2 |
| `ws.sheet_properties.pageSetUpPr.fitToPage = True` | włączenie dopasowania | — | **Bez tego `fitToWidth` jest ignorowane** (trzy ustawienia razem) |
| `ws.page_setup.fitToWidth = 1` | dopasowanie do szerokości | `int` | Z `fitToHeight = 0` (wysokość dowolna) |
| `ws.page_setup.orientation = "landscape"` | orientacja A4 poziomo | `"landscape"` | Można `ws.ORIENTATION_LANDSCAPE` |
| `ws.page_setup.paperSize = ws.PAPERSIZE_A4` | format papieru | `int` | Stałe arkusza, nie z `page_setup` |
| `ws.auto_filter.ref = "A1:E100"` | autofiltr na zakresie | `str` | Alternatywa dla Tabeli (moduł 12) |
| `cell.number_format = FORMATY["currency"]` | format liczbowy | `str` | Kody formatów z modułu 10 |
| `Font(bold=True, color="FFFFFF", size=11)` | czcionka | **bez** `bold` jako mutacji | **Niemutowalne** (moduł 09, pułapka 1) |
| `PatternFill(fill_type="solid", fgColor="1F4E78")` | tło komórki | `fgColor` = **tło** przy `solid` | Pułapka odwróconej intuicji (moduł 09, pułapka 6) |
| `PatternFill(fill_type="solid", bgColor="FCE4E4")` | tło w **regule** | `bgColor`, nie `fgColor` | W `dxf` odwrotnie niż w stylu komórki (moduł 25, pułapka 8) |
| `FormulaRule(formula=["$E2<0.05"], fill=..., font=...)` | reguła warunkowa z formuły | **lista**, **bez `=`** | Pułapka modułu 13: `=` psuje regułę |
| `ws.conditional_formatting.add(zakres, regula)` | dodanie reguły | zakres **osobno** | Zakres i formuła muszą być spójne (`$E2`, nie `$E$2`) |
| `Alignment(horizontal="center", vertical="center")` | wyrównanie | `wrap_text`, `text_rotation` | Wysokość wiersza Excel **nie zawsze** przelicza (moduł 09) |
| `ws.column_dimensions[litera].width` | szerokość kolumny | `float` | Litera z `get_column_letter()` |
| `ws.row_dimensions[wiersz].height` | wysokość wiersza | `float` | Nagłówek 22 pkt = czytelny |
| `get_column_letter(4)` → `"D"` | numer → litera | 1-based! | Excel numeruje od 1 (moduł 03, pułapka 4) |
| `wb.properties.title/creator/description/keywords` | metadane dokumentu | `str` | Wersja, okres, hash wsadu (sekcja 2.7) |
| `wb.save(sciezka)` | zapis na dysk | **przyjmuje obiekt plikopodobny** | **COMMIT** — po tym cofanie stosu zmienia tylko pamięć |
| `wb.save(BytesIO())` | zapis do pamięci | `BytesIO` | **`seek(0)` przed wysłaniem!** (pułapka 7) |
| `zipfile.ZipFile(namelist())` | inwentarz części pliku | — | **Tanie** — skanowanie utraty funkcji bez wczytywania |
| `os.replace(tmp, cel)` | atomowa podmiana | — | **Tylko na tym samym wolumenie** (pułapka 3) |
| `tempfile.mkstemp(dir=cel.parent)` | plik tmp w tym katalogu | `dir=` **kluczowe** | Bez `dir=` atomowość może nie zadziałać |
| `shutil.copy2(oryginał, backup)` | backup z metadanymi | — | **`copy2`**, nie `copy` |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — rozdziel kod na trzy warstwy i sprawdź `grep`-em.**

Weź **swój** skrypt (albo skrypt z modułu 25 — `RaportSprzedazy` + `StosKomend`) i przenieś go do struktury:

```text
projekt/
├── src/domain/         # modele, reguły, porty   - ZERO openpyxl
├── src/application/    # przypadki użycia         - ZERO openpyxl
└── src/infrastructure/excel/  # adapter + repo    - openpyxl TYLKO TU
```

**Wymagania:**

1. **`grep -rn "openpyxl" src/domain/ src/application/` musi dać pusty wynik.** Dosłownie puste. Jeżeli masz tam `from openpyxl import ...` — zaimportowałeś openpyxl w domenie, i trzeba go stamtąd wyrzucić.
2. **`grep -rn "openpyxl" src/` musi dać wyniki wyłącznie w `src/infrastructure/`.** Sprawdź to jedną komendą:
   ```bash
   grep -rn "openpyxl" src/ | cut -d: -f1 | sort -u
   ```
   Każda linia na wyjściu musi zawierać `infrastructure`.
3. **Domena musi mieć testy, które działają z `openpyxl` odinstalowanym.** Sprawdź:
   ```bash
   python -m venv .venv-bez-openpyxl
   .venv-bez-openpyxl/bin/pip install pytest
   .venv-bez-openpyxl/bin/pytest tests/domain/ tests/application/
   ```
   **Musi przejść.** Jeżeli nie — Twój test importuje coś z infrastruktury (np. `EksporterXlsxOpenpyxl`) i to trzeba poprawić.
4. **Narysuj diagram** (ASCII lub Mermaid) swojego projektu: trzy warstwy, porty, strzałki zależności. Zapisz do `output/26_cw1_diagram.md`.

**Pytania do `output/26_cw1_wnioski.md`:**

1. **Ile plików musiałeś zmienić, żeby przenieść `import openpyxl` do infrastruktury?** Jeżeli więcej niż **trzy** — napisz, gdzie openpyxl był „wszędzie", i opisz, co konkretnie to utrudniało (podpowiedź: testy, wymiana formatu, wdrażanie).
2. **Co się stało z czasem uruchomienia testów?** Zmierz `pytest tests/domain/ tests/application/` **przed** i **po** rozdzieleniu. Napisz różnicę.
3. **Czy `main()` / CLI stał się „brudny"?** Tak, i to jest **poprawne**. Uzasadnij w dwóch zdaniach: dlaczego granica kompozycji wygląda inaczej niż reszta projektu.

### 🟡 Warsztat

**Zadanie 2 — port, dwa adaptery i `EksporterFake` w testach.**

To jest **najważniejsze** ćwiczenie tego modułu. Zbudujesz trzy implementacje jednego portu i dowiesz się, **czego naprawdę wymaga Twój kontrakt**.

**Wymagania:**

1. **Port `EksporterRaportu`** — sygnatura `eksportuj(model, cel) -> WynikEksportu`. Zero typów openpyxl w sygnaturze. Ma być dokładnie tak, jak w `src/domain/ports.py`.
2. **Adapter `EksporterXlsxOpenpyxl`** — buduje trzy arkusze: `Dane` (z formatami i regułą warunkową), `Podsumowanie` (KPI), `_meta` (ukryty). Zapis przez `JednostkaPracy` (atomowo, z backupem).
3. **Adapter `EksporterCsv`** — zapisuje **te same** dane do CSV. Uwaga na dwie rzeczy: (a) `csv` ma problem z kodowaniem polskich znaków — użyj `encoding="utf-8-sig"` (BOM pomaga Excelowi otworzyć plik poprawnie), (b) `csv` nie ma formatów liczbowych — musisz zdecydować, czy zapisywać `str(Decimal)` czy `float`. **Zapisz uzasadnienie wyboru**.
4. **`EksporterFake`** — zapisuje model do listy, umie symulować awarię.
5. **Testy (obowiązkowe)** dla **wszystkich trzech** implementacji, używające **tych samych** asercji na modelu:

```python
@pytest.mark.parametrize("eksporter", [
    pytest.lazy_fixture("eksporter_xlsx"),
    pytest.lazy_fixture("eksporter_csv"),
    pytest.lazy_fixture("eksporter_fake"),
])
def test_kontrakt_portu(eksporter, model_raportu, tmp_path):
    wynik = eksporter.eksportuj(model_raportu, str(tmp_path / "raport"))
    assert wynik.wierszy == len(model_raportu.wiersze)
    assert wynik.arkuszy >= 1
    assert wynik.bajtow > 0
    assert wynik.format in ("xlsx", "csv", "fake")
```

Ten test jest **kluczowy**: jeżeli `EksporterCsv` lub `EksporterXlsx` się wywali, to znaczy, że port jest zaprojektowany „pod jedną technologię" i trzeba go **poprawić**. Dokładnie to jest wartością **dwóch** adapterów: obnażają nieszczelność kontraktu.

**Pytania do `output/26_cw2_wnioski.md`:**

1. **Co wykrył test kontraktu?** Napisz, **które** asercje przechodzą dla jednego adaptera, a nie przechodzą dla drugiego. To znaczy, że te asercje dotyczyły **techniki**, nie **kontraktu** — i trzeba je poprawić.
2. **Czy `arkuszy` ma sens dla CSV?** CSV nie ma arkuszy. Co zwróciłeś? `1`? A może port powinien być inny? Uzasadnij: **gdyby** CSV był w Twoim projekcie od początku, to `WynikEksportu` wyglądałby inaczej. Co to mówi o projektowaniu portów **zanim** znasz implementacje?
3. **Jak wygląda `cel` dla CSV?** `cel="output/raport"` z rozszerzeniem `.xlsx` w adapterze XLSX, a `.csv` w adapterze CSV. Kto dodaje rozszerzenie? **Adapter** (bo to jego sprawa). Uzasadnij, dlaczego **nie** aplikacja.

### 🔴 Wyzwanie

**Zadanie 3 — w pełni deklaratywny opis raportu w YAML + generator + testy + audyt.**

To zadanie łączy wszystko: konfigurację, architekturę warstw, transakcyjność, determinizm i audyt.

**Krok 1 — konfiguracja YAML.** Rozszerz `RaportSpec` o:
- `kolumny` (już masz),
- `kpi` (etykieta, źródło w agregatach: `suma_netto`, `srednia_marza`, `liczba_pozycji`),
- `reguly` — **wybór z zamkniętej palety**: `niskа_marza`, `duzy_klient`, `brak_danych`. **Nie wyrażenia.**
- `druk` (orientacja, powtarzanie nagłówka),
- `metadane` (wersja schematu, wersja generatora).

**Zamknięta paleta reguł** to kluczowa część tego ćwiczenia. Napisz `REGULY: dict[str, Callable[[WierszRaportu, dict], bool]]` w domenie. Konfiguracja **wybiera nazwy** z tej palety. Nowa reguła = **nowa funkcja w kodzie**, nie nowy wpis w YAML-u.

**Krok 2 — walidacja spec-u.**

`wczytaj_spec()` musi sprawdzić **przed** generowaniem:
- czy każda kolumna wskazuje na istniejące pole modelu (zamknięta lista),
- czy każda reguła jest w palecie,
- czy każdy format jest w palecie,
- czy KPI mają znane źródła agregatu,
- czy nie ma duplikatów nazw kolumn.

**Każdy błąd** musi mieć komunikat **z listą dostępnych wartości**. Zadanie: napisz test dla **każdego** z pięciu rodzajów błędu, używając `pytest.raises(match=...)`.

**Krok 3 — generator.**

`EksporterXlsxOpenpyxl` ma **iterować po `spec.kolumny`** i budować arkusz na podstawie **konfiguracji**, nie stałej listy. To znaczy, że dodanie kolumny w TOML-u **nie wymaga** zmiany kodu. Sprawdź to testem: dodaj kolumnę do TOML-a w `tmp_path` i policz, że w pliku wynikowym jest **o jedną** kolumnę więcej.

**Krok 4 — determinizm i idempotencja.**

Trzy testy, które **muszą** przechodzić:

```python
def test_dwa_uruchomienia_ten_sam_hash(spec, dane, tmp_path):
    a = generuj(spec, dane, tmp_path / "a")
    b = generuj(spec, dane, tmp_path / "b")
    assert a.hash_wyniku == b.hash_wyniku       # poza znacznikami czasu

def test_nie_nadpisuje_wejscia(spec, dane, tmp_path):
    wejscie = tmp_path / "data.csv"
    ...  # zapisz dane i zapamietaj hash
    generuj(spec, wejscie, tmp_path / "raport")
    assert _hash(wejscie) == hash_przed         # wejscie nietkniete

def test_dwa_uruchomienia_ten_sam_plik_zawartosc(spec, dane, tmp_path):
    generuj(spec, dane, tmp_path / "raport")
    pierwszy = podsumowanie_skoroszytu(tmp_path / "raport.xlsx")
    generuj(spec, dane, tmp_path / "raport")    # drugi raz na TYM SAMYM pliku
    drugi = podsumowanie_skoroszytu(tmp_path / "raport.xlsx")
    assert pierwszy == drugi                    # idempotencja
```

**Krok 5 — log audytowy.**

Każdy wygenerowany plik musi mieć:
- arkusz `_meta` z: `schemat`, `wygenerowano`, `generator`, `hash_wsadu`, `wersja_spec`,
- **sidecar JSON** obok pliku: `raport.xlsx` + `raport.meta.json` z tymi samymi metadanymi plus **pełną listą ostrzeżeń**,
- **wpis w centralnym dzienniku** `output/audyt.jsonl` (JSON Lines — jedna linia = jedno zdarzenie).

Trzy miejsca, trzy poziomy trwałości. **Uzasadnij wybór** w komentarzu: dlaczego dziennik w JSONL, a nie w arkuszu.

**Kryteria oceny (rubryka):**

| Kryterium | Waga | Na co patrzę |
|---|---|---|
| Port bez typów openpyxl | ★★★ | `grep "openpyxl" ports.py` → pusto; `grep "openpyxl" application/` → pusto |
| Testy domeny bez openpyxl | ★★★ | `pytest tests/domain/` w venv bez openpyxl → zielone |
| Trzy adaptery przechodzą ten sam test kontraktu | ★★★ | `test_kontrakt_portu` sparametryzowany, wszystkie zielone |
| Reguły w **zamkniętej palecie** | ★★★ | konfiguracja **wybiera** nazwy; brak wyrażeń w YAML-u |
| Walidacja spec-a przed generowaniem | ★★☆ | pięć błędów, każdy z komunikatem i listą dostępnych wartości |
| Atomowość: brak `.tmp` po błędzie | ★★★ | test z symulowanym błędem — brak plików `.*.tmp.xlsx` |
| Backup: `.bak` **tylko** przy nadpisaniu | ★★☆ | pierwszy zapis → brak `.bak`; drugi → jeden `.bak` |
| Determinizm: dwa hashe zgodne | ★★★ | `a.hash_wyniku == b.hash_wyniku` |
| Idempotencja: drugie uruchomienie bez zmian | ★★★ | podsumowanie pliku identyczne |
| `_meta` + sidecar + JSONL (trzy poziomy) | ★★☆ | w każdym z trzech miejsc obecny hash wsadu |
| Utrata funkcji wykryta i zalogowana | ★★☆ | `przeskanuj_przed_zapisem` na pliku z sparkline'ami zgłasza alert |

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe decyzje i fragmenty</strong></summary>

**Decyzja 1: reguły jako zamknięta paleta — i dlaczego to nie „ograniczenie", a „zabezpieczenie".**

```python
"""src/domain/paleta_regul.py - reguly konfigurowalne.

Konfiguracja wybiera NAZWE z tego slownika. Nie moze dopisac wlasnej.
Zysk: reguly ewoluuja w KODZIE (z historia zmian i testami), a nie
w TOML-u, ktory nikt nie wersjonuje i nikt nie testuje.
"""

from __future__ import annotations

from decimal import Decimal
from typing import Callable

from .modele import WierszRaportu

# Typ reguly: (wiersz, parametry) -> bool
Regula = Callable[..., bool]


def nizka_marza(wiersz: WierszRaportu, prog: float = 0.05) -> bool:
    return wiersz.marza_proc < Decimal(str(prog))


def duzy_klient(wiersz: WierszRaportu, minimum_pln: float = 10000.0) -> bool:
    return wiersz.netto_pln >= Decimal(str(minimum_pln))


def brak_danych(wiersz: WierszRaportu, **_: object) -> bool:
    return not wiersz.klient or wiersz.ilosc <= 0


REGULY: dict[str, Regula] = {
    "niska_marza": nizka_marza,
    "duzy_klient": duzy_klient,
    "brak_danych": brak_danych,
}
```

I **to samo w konfiguracji**:

```yaml
# config/raport.yaml
raport:
  identyfikator: raport_miesieczny
  tytul: "Raport sprzedazy miesiecznej"
  schemat: raport_v3
  wersja_generatora: "26.0"

kolumny:
  - nazwa: "Klient"
    zrodlo: klient
    format: text
    szerokosc: 30
  - nazwa: "Sprzedaz netto"
    zrodlo: netto_pln
    format: currency
  - nazwa: "Marza"
    zrodlo: marza_proc
    format: percent

reguly:
  - nazwa: niska_marza
    parametr: 0.05
    tlo: "FCE4E4"
    kolor_tekstu: "9C0006"
  - nazwa: duzy_klient
    parametr: 25000
    tlo: "E2EFDA"
    kolor_tekstu: "375623"
```

**Zauważ, co jest możliwe, a co nie.** Możesz **wybrać** `niska_marza` i podać próg. **Nie możesz** napisać `marza_proc < abs(srednia_regionu) * 1.5` — bo to jest **wyrażenie**, a wyrażenia w YAML-u to nowy język programowania. Jeżeli potrzebujesz takiej reguły, **dopisujesz funkcję** do `paleta_regul.py`, **dodajesz test**, robisz PR i wdrażasz. To jest wolniejsze — i to jest **poprawne**, bo reguła biznesowa powinna być wolniejsza do zmiany niż kolor tła.

**Decyzja 2: trzech poziomów audytu — dlaczego każdy jest potrzebny.**

```python
def zapisz_wynik(wynik: WynikEksportu, model: ModelRaportu, cel: Path,
                 dziennik: Path) -> None:
    """Trzy poziomy zapisu metadanych. Kazdy z innym powodem istnienia."""
    # POZIOM 1: w samym pliku (_meta). Wykonal juz adapter.
    #   Powod: plik wedruje (mail, pendrive, archiwum). Log nie podaza.

    # POZIOM 2: sidecar JSON. Powod: program moze czytac metadane BEZ
    #   otwierania Excela (a otwarcie xlsx to koszt i zaleznosc od openpyxl).
    sidecar = cel.with_suffix(".meta.json")
    sidecar.write_text(json.dumps({
        "schemat": model.schemat,
        "wygenerowano": model.wygenerowano,
        "hash_wsadu": model.hash_wsadu,
        "hash_wyniku": wynik.hash_wyniku,
        "wierszy": wynik.wierszy,
        "sciezka": str(cel),
        "ostrzezenia": list(wynik.ostrzezenia),
    }, ensure_ascii=False, indent=2), encoding="utf-8")

    # POZIOM 3: centralny dziennik JSONL. Powod: agregujemy wiedze
    #   o WSZYSTKICH raportach w jednym miejscu - mozna pytac "kiedy byl
    #   ostatni raport z tym schematem" bez skanowania katalogu.
    #   JSONL, a nie JSON, bo: dopisujemy jedna linie, nie przepisujemy calosci;
    #   przy awarii w polowie ostatnia linia moze byc niepelna, poprzednie sa OK.
    with dziennik.open("a", encoding="utf-8") as plik:
        plik.write(json.dumps({
            "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "okres": model.okres,
            "schemat": model.schemat,
            "wierszy": wynik.wierszy,
            "bajtow": wynik.bajtow,
            "hash_wsadu": model.hash_wsadu,
            "sciezka": str(cel),
            "ostrzezen": len(wynik.ostrzezenia),
        }, ensure_ascii=False) + "\n")
```

Trzy powody, trzy miejsca. **Dziennik w archiwum** nie ma sensu (za duży). **Sidecar w mailu** gubi się razem z plikiem (wysyłasz tylko `.xlsx`). **`_meta` w pliku** jest widoczny w każdym środowisku, ale trudny do zbiorczego przeszukania (każde odpytanie = otwarcie skoroszytu). **Trzy poziomy** pokrywają trzy różne scenariusze. To jest wybór architektoniczny, a nie „wszędzie to samo, na wszelki wypadek".

**Decyzja 3: dlaczego JSONL, a nie JSON.**

```text
JSON (lista):
  {"raporty": [ {...}, {...} ]}   <- przy dodaniu wpisu przepisujesz CALOSC
  Awaria w polowie -> plik nieparsowalny. Tracisz wszystko.

JSONL (JSON Lines):
  {...}
  {...}
  {...}          <- dopisujesz JEDNA LINIE (append)
  Awaria w polowie -> ostatnia linia niepelna, POPRZEDNIE sa OK.
```

To jest ta sama zasada co log serwera: **append-only jest odporne na awarie w środku zapisu**. I jesteś tutaj w warstwie infrastruktury — czyli wolno Ci korzystać z wiedzy, że `os.replace` jest atomowy **na całości pliku**, a `append` nie jest. Dwa różne mechanizmy, dwa różne zastosowania.

**Decyzja 4: test idempotencji — co dokładnie porównujemy.**

```python
def podsumowanie_skoroszytu(sciezka: Path) -> dict:
    """Odcisk struktury pliku do porownania. NIE bajt w bajt!

    Bajty roznia sie znacznikami czasu w ZIP (metadane archiwum),
    wiec porownanie bajtowe byloby zawsze czerwone (modul 20, pulapka 1).
    """
    from openpyxl import load_workbook

    wb = load_workbook(sciezka, read_only=True)
    try:
        return {
            "arkusze": tuple(wb.sheetnames),
            "wiersze": {
                nazwa: sum(1 for _ in wb[nazwa].iter_rows()) - 1    # -1 naglowek
                for nazwa in wb.sheetnames if nazwa != "_meta"
            },
            "formuly_regul": {
                nazwa: sum(
                    len(cf.rules) for cf in wb[nazwa].conditional_formatting
                )
                for nazwa in wb.sheetnames
            },
            "hash_wsadu": wb["_meta"]["B6"].value,      # z arkusza _meta
            "schemat": wb["_meta"]["B1"].value,
        }
    finally:
        wb.close()
```

I test:

```python
def test_idempotencja(tmp_path, spec, dane):
    cel = tmp_path / "raport.xlsx"
    generuj(spec, dane, cel)
    przed = podsumowanie_skoroszytu(cel)
    generuj(spec, dane, cel)                # drugi raz, ten sam plik
    po = podsumowanie_skoroszytu(cel)
    assert przed == po
```

Ten test **wyłapie**:
- podwajanie wierszy (jeśli generator dopisuje),
- zmieniającą się liczbę arkuszy,
- zgubione reguły warunkowe,
- zmieniony hash wsadu (co by znaczyło, że dane wejściowe „dryfują").

**Decyzja 5: co świadomie zostawiam poza zakresem tego ćwiczenia.**

- **Nie robię PDF-a przez LibreOffice.** Byłby to `subprocess.run(["soffice", "--headless", "--convert-to", "pdf", ...])` — i **działa**, ale: (a) LibreOffice musi być zainstalowany, (b) konwersja trwa sekundy, (c) jest zewnętrznym procesem, który trzeba obsłużyć (timeout, kod wyjścia, sprzątanie tmp). To jest materiał na oddzielne rozszerzenie, nie na to ćwiczenie. **Zaznaczam, gdzie by wszedł**: kolejny adapter `EksporterPdfLibreOffice` implementujący ten sam port.
- **Nie robię kolejki ani workera.** Zbyt dużo infrastruktury na jedno ćwiczenie. Wystarczy, że wiem, że **decyzja** jest w warstwie aplikacji, nie w adapterze. Ten sam `eksportuj()` obsłuży synchroniczny i asynchroniczny scenariusz.
- **Nie robię szyfrowania pliku.** openpyxl **nie umie** zapisać zaszyfrowanego `.xlsx` (nie ma obsługi szyfrowania OOXML). Jeżeli potrzebujesz hasła na plik: lokalizacja w systemie z uprawnieniami, albo zewnętrzne narzędzie. **Powiedz to wprost** użytkownikowi w `_meta` — bo inaczej będzie próbował `wb.security.workbookPassword` i zdziwi się, że to ochrona struktury, a nie szyfrowanie (moduł 17).

</details>

## 6. Typowe błędy i pułapki

**1. Objaw: „widok Django importuje openpyxl, żeby wygenerować raport w odpowiedzi HTTP, a logika marży jest w tym samym pliku co `wb.save()`."**
→ *Przyczyna:* **warstwa infrastruktury przecieka do domeny** (i odwrotnie). Cała logika i cała technika w jednym miejscu. Konsekwencje widzisz dopiero przy zmianie: nie da się zmienić formatu (bo logika jest w widoku), nie da się przetestować marży (bo test potrzebuje openpyxl), nie da się użyć tego samego kodu w CLI (bo widok).
→ *Naprawa:* **trzy katalogi i reguła zależności.** `src/domain/` (bez openpyxl), `src/application/` (bez openpyxl), `src/infrastructure/excel/` (openpyxl). Widok/CLI **nie generuje** raportu — woła przypadek użycia z aplikacji. Sprawdzenie: `grep -rn "openpyxl" src/domain/ src/application/` musi być **puste**. Nie „prawie puste" — puste. Każdy wynik to miejsce, w którym logika i technika spotkały się w jednym pliku.

**2. Objaw: „używam `SkoroszytRepository`, mam `wczytaj()` i `zapisz()`, ale przy błędzie w połowie generowania plik docelowy jest uszkodzony — częściowo stary, częściowo nowy."**
→ *Przyczyna:* **repository ukrywające nieatomowy zapis.** Klasa nazywa się „repository" i ludzie jej wierzą, ale w środku robi `wb.save(cel)` — czyli **nadpisuje plik docelowy w miejscu**. Przy błędzie w połowie masz plik, który nie jest ani stary, ani nowy. To jest **ta sama** pułapka co z modułu 18, tylko schowana za dobrze brzmiącą nazwą.
→ *Naprawa:* **atomowy zapis w `zapisz()`**: `mkstemp` w `cel.parent` → `wb.save(tmp)` → weryfikacja → `shutil.copy2(cel, backup)` → `os.replace(tmp, cel)` → `unlink(tmp)` w `finally`. Do tego **test**: zasymuluj wyjątek w trakcie `save` (np. przez atrapę, która rzuca) i sprawdź, że plik docelowy jest **niezmieniony** (hash przed i po są równe), a pliki `.tmp` zostały usunięte. To jest jedyny test, który dowodzi, że „repository" robi to, co obiecuje nazwa.

**3. Objaw: „`os.replace(tmp, cel)` rzuca `OSError: Invalid cross-device link` — a czasem, na produkcji, plik jest zapisany, ale to nie jest plik, który zapisałem."**
→ *Przyczyna:* **plik tymczasowy na innym wolumenie.** Jeżeli `tmp` powstał przez `tempfile.mkstemp()` **bez `dir=`**, to leży w `/tmp`. Jeżeli `/tmp` jest na innym systemie plików (a w kontenerach Docker często **jest**), `os.replace()` **nie może** wykonać atomowej podmiany nazwy — musi **skopiować** zawartość. Wtedy tracisz atomowość i, przy awarii w trakcie kopiowania, zostawiasz **częściowo** nadpisany plik docelowy.
→ *Naprawa:* **`tempfile.mkstemp(dir=cel.parent)`.** Zawsze. Bez wyjątków. Jeżeli katalog docelowy jest na wolumenie sieciowym (NFS/SMB), **atomowość `os.replace` zależy od implementacji** — sprawdź to lub **nie polegaj** na niej (wtedy wersjonuj nazwy plików i nie nadpisuj). Diagnostyka: `os.stat(tmp).st_dev == os.stat(cel).st_dev` — jeżeli **nie** są równe, atomowość nie zadziała. Wstaw tę asercję do kodu produkcyjnego jako `raise` — bo cicha utrata atomowości jest gorsza niż jawny błąd.

**4. Objaw: „na Windows, gdy użytkownik ma otwarty plik w Excelu, `os.replace` rzuca `PermissionError`, a mój skrypt zapisuje raport gdzieś indziej i użytkownik myśli, że zadziałało."**
→ *Przyczyna:* **brak jawnej polityki „plik otwarty".** Windows blokuje **zamianę** otwartego pliku, ale Linux/macOS **nie** (możesz podmienić plik, na który ktoś ma otwarty deskryptor). Skrypt, który „działa na Linuksie", na Windows zgłosi `PermissionError` — albo, co gorsza, **łapie go** i zapisuje pod inną nazwą, **udając sukces**.
→ *Naprawa:* **trzy jawności.** (1) Domyślna nazwa wyjściowa zawiera **timestamp** (`raport_2026-09-15T1430.xlsx`) — nie nadpisujesz plików, których losu nie znasz. (2) Jeżeli **musisz** nadpisać, łap **dokładnie** `PermissionError` i podnieś własny `BladZapisu("Plik {cel} jest otwarty w innym programie. Zapisz pod nową nazwą albo zamknij plik.")` — z instrukcją, a nie pentlą na `dir` (pułapka 5). (3) Jeżeli Twoim odbiorcą jest konkretna osoba, **poinformuj ją** w UI/logu, gdzie trafił plik. Zapisywanie pod zmienioną nazwą **po cichu** jest najgorszym możliwym rozwiązaniem: użytkownik nie wie, gdzie jest raport.

**5. Objaw: „funkcja wychwyciła `PermissionError`, więc poszła „na piechotę" i znalazła wolną nazwę: `raport(1).xlsx`, `raport(2).xlsx`… i po miesiącu mam w katalogu 340 raportów `raport(n).xlsx`, w tym 200 plików zerowego rozmiaru."**
→ *Przyczyna:* **brak świadomej polityki nadpisywania plików.** Ktoś zapisał „fallback", który ratuje przed błędem, ale **nie przewidział**, że ta ścieżka będzie kiedyś główną. Każde uruchomienie tworzy nowy plik, więc katalog rośnie, a dysk się kończy. Dodatkowo: raporty `(n)` są bezwartościowe (nikt nie wie, który jest aktualny) — dokładnie odwrotnie niż chcemy.
→ *Naprawa:* **polityka nazewnictwa jawna w konfiguracji.** Trzy warianty do wyboru: (a) `nadpisz` — ostrzegaj, ale nadpisuj (dobre dla raportów roboczych), (b) `wersjonuj` — `raport_2026-09_v{n}.xlsx` z inkrementowanym numerem (dobre dla audytu), (c) `data` — `raport_2026-09_20261008.xlsx`, nadpisanie raz na dzień. Wybór **należy do użytkownika** i jest **jawny** — w TOML-u albo w argumencie CLI. Nigdy automatyczny, nigdy „na piechotę". I dodatkowo: **polityka retencji** — sprzątanie plików starszych niż 30 dni. Bez tego katalog rośnie w nieskończoność.

**6. Objaw: „endpoint HTTP zwraca plik z zawartością `0 bajtów`. Status 200, nagłówki OK, plik leci, ale się nie otwiera."**
→ *Przyczyna:* **brak `seek(0)` przed wysłaniem `BytesIO`.** `wb.save(bufor)` przesuwa wskaźnik bufora **na koniec**. Kiedy `StreamingResponse`/`send_file` zaczyna czytać, dostaje **koniec** pliku, więc wysyła zero bajtów. To jest tak częsta i tak cicha pułapka, że zasługuje na własny punkt.
→ *Naprawa:* **`bufor.seek(0)` zawsze, tuż przed `return`** (sekcja 2.11). Dobra rada dodatkowa: opakuj to w **jedną** funkcję, żeby nie dało się zapomnieć:

```python
def do_strumienia(wb) -> BytesIO:
    """Buduje BytesIO gotowe do wyslania (wskaznik na poczatku)."""
    bufor = BytesIO()
    wb.save(bufor)
    bufor.seek(0)
    return bufor
```

Diagnostyka: w teście sprawdź `len(bufor.getvalue()) > 0` **oraz** `bufor.tell() == 0` — a nie tylko jedno z dwóch.

**7. Objaw: „konfiguracja rośnie — ktoś dopisał `warunek: 'if (marza < 0.05) and (region in vip) then red else none'`, a ja nie mam pojęcia, jak to parsować."**
→ *Przyczyna:* **konfiguracja deklaratywna rosnąca do rozmiarów języka.** Jeden próg i jeden kolor to konfiguracja. Wyrażenie warunkowe z `if`, `and`, `in`, `else` to **program** — tyle że napisany w **innym języku**, bez podpowiedzi edytora, bez testów, bez sprawdzania typów i bez kodu zrozumiałego dla Pythona. Po dwóch miesiącach nikt nie wie, co znaczy `in vip`, skąd pochodzi `vip` i czy `red` to nazwa czy kolor.
→ *Naprawa:* **zamknięta paleta.** Konfiguracja **wybiera nazwę reguły z listy**, a parametry dostarcza jako proste wartości (liczba, string). Wyrażenia — **nigdy**. Jeżeli pojawia się potrzeba wyrażenia, to **nowa funkcja w kodzie** z testem. Diagnostyka: `grep -n "if \|and \|or \|not " config/*.toml config/*.yaml` — jeżeli **cokolwiek** znajdziesz, to sygnał. To jest dokładnie ta granica z sekcji 2.6 i nie ma od niej odstępstwa.

**8. Objaw: „dwa uruchomienia generatora z tymi samymi danymi dają pliki o różnych rozmiarach. Raz 450 KB, raz 458 KB, raz 452 KB."**
→ *Przyczyna:* **niedeterministyczność wsadu albo kolejności.** Cztery źródła z sekcji 2.8: `set` bez klucza, `datetime.now()` w logice, `glob()` bez `sorted()`, sortowanie zależne od locale. Wszystkie prowadzą do tego samego objawu — trzech różnych plików z „tego samego wsadu". Pliki różnią się nie „treścią" (choć często też), ale **kolejnością znaków**, bo `str`, `repr` i `sorted` są **niestabilne bez klucza** na zbiorach.
→ *Naprawa:* **cztery konkretne zamiany.** `for x in zbiór` → `for x in sorted(zbiór)`. `datetime.now()` w logice → **wstrzyknięty zegar**. `glob()`/`iterdir()` → `sorted(...)`. `sorted(teksty)` → `sorted(teksty, key=lambda s: s.casefold())`. Do tego **test**: dwa uruchomienia → `a.hash_wyniku == b.hash_wyniku` (przykład 7). Diagnostyka dodatkowa: `print(sorted(dane, key=repr))` **przed** przetwarzaniem, żeby zobaczyć, czy **wejście** jest deterministyczne.

**9. Objaw: „wpisałem `wb.security.workbookPassword = 'tajne'` i jestem przekonany, że to zaszyfrowało plik. Otworzyłem go i widzę wszystkie dane."**
→ *Przyczyna:* **mylenie ochrony struktury z szyfrowaniem.** `lockStructure` i `workbookPassword` **chronią przed modyfikacją struktury** w Excelu — nie szyfrują pliku, nie ukrywają zawartości, nie chronią przed otwarciem w innym edytorze. Analogia z modułu 17: **kłódka na szufladzie w biurku, nie sejf**. Do tego dochodzi drugi fakt: **openpyxl nie umie zapisać zaszyfrowanego `.xlsx`** — nie ma obsługi szyfrowania OOXML.
→ *Naprawa:* **nazwij to wprost i wybierz właściwe narzędzie.** Jeżeli dane są poufne: katalog z uprawnieniami, grupa odbiorców, transport szyfrowany, a do plików — zewnętrzne narzędzie (np. `7z` z hasłem, albo szyfrowanie na poziomie systemu plików). Jeżeli **naprawdę** potrzebujesz zaszyfrowanego `.xlsx`, to **nie openpyxl**. I co ważne: **dopisz to do `_meta`** — żeby użytkownik wiedział, że „hasło do skoroszytu" nie jest szyfrowaniem.

**10. Objaw: „endpoint zwrócił `MemoryError` przy raporcie 800 tys. wierszy. Wcześniej działał dla 100 tys."**
→ *Przyczyna:* **synchroniczne generowanie w pamięci bez limitu.** Trzy problemy naraz: (1) cały model w pamięci, (2) cały skoroszyt w pamięci, (3) cały `BytesIO` w pamięci. Dla 800 tys. wierszy mnożysz te trzy rozmiary. To jest **decyzja architektoniczna**, która nie została podjęta — więc została podjęta **przypadkiem**, i to na twoją niekorzyść.
→ *Naprawa:* **wybierz wzorzec z tabeli w sekcji 2.11.** Dla dużych raportów: (a) **zapis na dysk + link**, (b) generowanie w kolejce zadań, (c) **strumieniowanie wierszy** z bazy bezpośrednio do `write_only` (moduł 19) — ale wtedy pytanie „jest jedna seria, czy muszę najpierw policzyć?” jest **kluczowe**, bo `write_only` nie pozwala wrócić. Najpierw **limit wejścia** (np. 200 tys. wierszy na request) z jasnym komunikatem „dla większego raportu użyj trybu wsadowego", a dopiero potem optymalizacja. **Nigdy w odwrotnej kolejności** — najpierw optymalizujesz, potem dodajesz limit, i okazuje się, że optymalizujesz coś, co nigdy nie działało na produkcji.

**11. Objaw: „loguję na produkcji wszystko, ale i tak nie wiem, dlaczego **ten konkretny** plik u klienta nie ma kolumny. Nie chcę prosić klienta o plik, żeby sprawdzić, z której wersji powstał."**
→ *Przyczyna:* **metadane nie są w pliku.** Log na serwerze ma limit retencji, a plik u klienta jest poza Twoim zasięgiem. Jeżeli w pliku nie ma **wersji generatora**, **wersji schematu** i **hasła wsadu**, to nie odpowiesz na pytanie „skąd to i czy kompletne" bez proszenia o trzy rzeczy.
→ *Naprawa:* **arkusz `_meta` + `wb.properties`** z przykładu 6 i przykładu 10 w tym module. Dopisz do `_meta`: wersję generatora, wersję schematu, hash wsadu, datę, autora, liczbę wierszy, **listę ostrzeżeń**. To są trzy linie w adapterze, a odpowiadają na **sto** pytań wsparcia technicznego. Dodatkowo: **sidecar JSON** obok pliku — bo po metadane (wersja, hash) można sięgnąć **bez otwierania xlsx** i **bez** zależności od openpyxl.

**12. Objaw: „klient przysłał raport z powrotem z naniesionymi poprawkami. Program wczytał, zapisał i **cicho** zgubił sparkline'y, slicery i kształty — klient dostał plik, w którym zniknęło to, co w nim było, i tego nie zauważył."**
→ *Przyczyna:* **brak wykrywania utraty funkcji przy przetwarzaniu plików, które przyszły z zewnątrz.** To jest **bezpośrednia** konsekwencja braku architektury: pokusa, by zrobić „read-modify-write" na każdym pliku, który wejdzie. Wyjątek z modułu 18: **nigdy w miejscu**. I jeszcze: brak **jawnego ostrzeżenia** o elementach, których openpyxl nie modeluje.
→ *Naprawa:* **trzy reguły.** (1) **Reguła „nigdy nie nadpisywać plików, których nie stworzyliśmy".** Wejście → generujemy **nowy** plik → zapisujemy pod własną nazwą. Jeżeli musisz zmodyfikować cudzy plik, zapisujesz **kopię** z sufiksem (`_v2`). (2) **Skanowanie przed zapisem** (`przeskanuj_przed_zapisem` z przykładu 10) — lista części `xl/sparklineGroups`, `xl/slicers`, `xl/timelines` w pliku wejściowym i **jawne ostrzeżenie** w `_meta` i w logu. (3) **Zdarzenie alertujące** — nie `INFO`, nie `DEBUG`, tylko `WARNING`. Bo utrata danych, o której nie wiesz, jest **gorsza** niż błąd, o którym wiesz.

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Trzy światy i jedna reguła: kuchnia nie projektuje pudełek.** Domena gotuje (reguły, modele — **zero openpyxl**). Aplikacja kieruje salą (przypadki użycia — **zero openpyxl**). Infrastruktura dowozi (openpyxl, pliki, **`os.replace`**). Test jest mechaniczny i bezlitosny: `grep -rn "openpyxl" src/domain/ src/application/` musi być **puste**. Nie „prawie puste" — puste. Każdy wynik to miejsce, w którym logika i technika spotkały się w jednym pliku. Zależności płyną **do środka**: domena nie zna nikogo, aplikacja zna domenę, infrastruktura zna oba. **Nigdy odwrotnie.**

2. **Port to kształt, nie technologia — i dlatego wygrywa.** `EksporterRaportu` ma sygnaturę bez `Workbook`, bez `Worksheet`, bez `BytesIO` i bez `Path`. Port opisuje **intencję** („weź model, zapisz pod celem, zwróć wynik"), nie mechanizm. Dopiero z **trzema** implementacjami (`Xlsx`, `Csv`, `Fake`) widzisz, co naprawdę jest kontraktem. Kierunek zależności jest odwrócony: kod, który **woła** (aplikacja), nie zależy od kodu, który **robi** (openpyxl). Dzięki temu **testy domeny nie potrzebują openpyxl** — nie „dają rady" bez niego, ale **nie mają z nim nic wspólnego**.

3. **Repository i Unit of Work mają w naszym projekcie JEDEN powód: kontrolować, kiedy plik jest przepisywany.** „Repo" dla pliku to nie baza danych. Cały sens: `wb.save()` w jednym miejscu, atomowo, z backupem, weryfikacją i logiem. Transakcja to sekwencja: **sprawdź katalog → tmp w tym samym katalogu → weryfikacja tmp → backup oryginału → `os.replace` → sprzątanie w `finally`**. `mkstemp(dir=cel.parent)` nie jest opcjonalne — bez tego tracisz atomowość (pułapka 3). Brak zapisu = pełny rollback — i to jest najprostszy mechanizm cofania, jaki masz (moduł 25, sekcja 2.3).

4. **Konfiguracja opisuje CO. Kod opisuje JAK.** Cztery poziomy deklaratywności: kod → dataclass → TOML/YAML → „DSL". W **większości** projektów wystarcza drugi; trzeci wymaga walidacji, wersjonowania i dyscypliny. Reguła graniczna: konfiguracja **wybiera z zamkniętej palety**, nie **wyraża obliczeń**. Każdy `if` w YAML-u to sygnał, że tworzysz **nowy język programowania** bez testów, podpowiedzi edytora i historii zmian. Nowa reguła = **nowa funkcja w kodzie**, z testem, PR-em i **powolną** zmianą. To jest poprawne — reguła biznesowa powinna być **wolniejsza** do zmiany niż kolor tła.

5. **Raport to projekcja, nie magazyn.** Excel to widok danych, jak paragon — transakcja jest w systemie, karteczka jest jej wyświetleniem. Z tego wynikają trzy konsekwencje: (a) **Excel nie jest źródłem prawdy** — nie czytaj go z powrotem jako wsadu, chyba że **świadomie** zbudowałeś port wejściowy i walidację, (b) **ten sam wsad → ten sam plik** (poza znacznikami czasu) — wstrzyknij zegar, sortuj stabilnie, żadnych `set` w kolejności, (c) **raport jest jednokierunkowy** — domena → aplikacja → infrastruktura. Dodaj do każdego pliku `_meta` (wersja generatora, wersja schematu, hash wsadu, ostrzeżenia) i skanuj utratę funkcji **przed** zapisem — bo utrata, o której nie wiesz, jest gorsza niż błąd, o którym wiesz. **Bezpieczeństwo jest na wejściu do domeny**, nie przy zapisie: jedno miejsce odrzucania znaków wyzwalających **chroni wszystkie** ścieżki wyjścia.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. STRUKTURA PROJEKTU - trzy warstwy, jeden kierunek zaleznosci
# ==================================================================
# src/
#   domain/         modele, reguly, porty           ZERO openpyxl
#   application/    przypadki uzycia                ZERO openpyxl
#   infrastructure/ excel/, csv/, pdf/, fake/       openpyxl TYLKO TU
# cli.py  app.py  worker.py                         ← brzeg (kompozycja)

# TEST POZIOMU 1 (musi byc pusty):
#   grep -rn "openpyxl" src/domain/ src/application/
# TEST POZIOMU 2 (wszystkie wyniki w infrastructure/):
#   grep -rn "openpyxl" src/ | cut -d: -f1 | sort -u


# ==================================================================
# 2. PORTY I ADAPTERY - "port to ksztalt, nie technologia"
# ==================================================================
from typing import Protocol
from dataclasses import dataclass

@dataclass(frozen=True)
class WynikEksportu:                    # NASZ DTO - zero openpyxl
    format: str
    bajtow: int
    arkuszy: int
    wierszy: int
    sciezka: str | None
    hash_wyniku: str = ""
    ostrzezenia: tuple[str, ...] = ()

class EksporterRaportu(Protocol):       # bez Workbook/Worksheet/BytesIO/Path
    @property
    def format(self) -> str: ...
    def eksportuj(self, raport: ModelRaportu, cel: str) -> WynikEksportu: ...

# Trzy adaptery tego samego portu:
#   EksporterXlsxOpenpyxl  (openpyxl + JednostkaPracy)
#   EksporterCsv           (csv.writer, utf-8-sig, bez formatow)
#   EksporterFake          (lista zapisow - DO TESTOW)
# Test kontraktu (parametryzowany po trzech) obnaza nieszczelnosc portu!


# ==================================================================
# 3. TRANSAKCJA ZAPISU - kolejnosc OBOWIAZKOWA
# ==================================================================
def zapisz_atomowo(wb, cel: Path, wersja: str = "1") -> Path:
    import os, shutil, tempfile
    from datetime import datetime

    cel.parent.mkdir(parents=True, exist_ok=True)

    # tmp MUSI byc w TYM SAMYM katalogu - inaczej os.replace nie jest atomowy!
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{cel.stem}.", suffix=".tmp.xlsx", dir=cel.parent
    )
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)                                    # 5. zapis
        sprawdzony = load_workbook(tmp, read_only=True)  # 6. weryfikacja
        try:
            assert len(sprawdzony.sheetnames) == len(wb.sheetnames)
        finally:
            sprawdzony.close()                          # modul 19: close()!
        if cel.exists():                                # 7. backup PRZED
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            shutil.copy2(cel, cel.with_name(f"{cel.stem}.{ts}.bak{cel.suffix}"))
        os.replace(tmp, cel)                            # 8. COMMIT
        return cel
    except Exception as blad:
        tmp.unlink(missing_ok=True)                     # 9. sprzatanie
        raise BladZapisu(
            f"Zapis do {cel.name} nie udal sie "
            f"({type(blad).__name__}: {blad}). Plik docelowy NIE zmieniony."
        ) from blad

# DIAGNOSTYKA atomowosci:
#   os.stat(tmp).st_dev == os.stat(cel).st_dev   # musza byc ROWNE


# ==================================================================
# 4. DEPENDENCY INJECTION - bez frameworka, jeden plik
# ==================================================================
class GenerujRaportMiesieczny:
    def __init__(self, *, zrodlo: ZrodloDanych, eksporter: EksporterRaportu,
                 magazyn: MagazynRaportow | None = None, zegar=None):
        self._zrodlo = zrodlo
        self._eksporter = eksporter
        self._magazyn = magazyn
        self._teraz = zegar or (lambda: datetime.now(timezone.utc))  # DETERMINIZM

    def wykonaj(self, *, okres: str, cel: str) -> WynikUzycia:
        zamowienia, hash_wsadu = self._zrodlo.pobierz(okres)
        model, ostrzezenia = zbuduj_model_raportu(
            zamowienia, okres=okres, hash_wsadu=hash_wsadu,
            wygenerowano=self._teraz().isoformat(timespec="seconds"))
        wynik = self._eksporter.eksportuj(model, cel)
        log.info("raport koniec okres=%s format=%s wierszy=%s bajtow=%s "
                 "ostrzezen=%s", okres, wynik.format, wynik.wierszy,
                 wynik.bajtow, len(ostrzezenia) + len(wynik.ostrzezenia))
        return WynikUzycia(wynik=wynik, okres=okres, schemat=model.schemat,
                           hash_wsadu=hash_wsadu, czas_s=0.0)

# cli.py - JEDYNE miejsce z konkretnymi typami:
def kompozycja(argv):
    eksportery = {"xlsx": EksporterXlsxOpenpyxl, "csv": EksporterCsv,
                  "fake": EksporterFake}
    eksporter = eksportery[argv.format]()
    przypadek = GenerujRaportMiesieczny(zrodlo=ZrodloCsv(argv.wejscie),
                                        eksporter=eksporter)
    return przypadek.wykonaj(okres=argv.okres, cel=argv.wyjscie)


# ==================================================================
# 5. KONFIGURACJA - "CO, nie JAK"
# ==================================================================
# ZAMKNIETE PALETY w kodzie:
DOSTEPNE_FORMATY = {"currency", "percent", "int", "date", "text"}
DOSTEPNE_POLA    = {"klient", "region", "ilosc", "netto_pln", "marza_proc"}
REGULY = {                     # nazwa -> funkcja; konfiguracja wybiera NAZWE
    "niska_marza": nizka_marza,
    "duzy_klient": duzy_klient,
    "brak_danych": brak_danych,
}
# W TOML/YAML konfiguracja ma WYLACZNIE:
#   - nazwy pol z DOSTEPNE_POLA
#   - nazwy formatow z DOSTEPNE_FORMATY
#   - nazwy regul z REGULY + proste parametry (liczba, string, kolor)
# WYKLUCZONE w konfiguracji: if/and/or/not, wyrazenia, odwolania do baz,
#   wlasne wyrazenia regularne, "wyrazenia" w ogole.
# Sygnal ostrzegawczy: grep -n "if \|and \|or \|not " config/*


# ==================================================================
# 6. DETERMINIZM I IDEMPOTENCJA
# ==================================================================
# Cztery zrodla niedeterminizmu -> cztery naprawy:
#   for x in zbior                      -> for x in sorted(zbior)
#   datetime.now() w logice             -> wstrzykniety zegar
#   glob()/iterdir()                    -> sorted(...)
#   sorted(teksty)                      -> sorted(teksty, key=str.casefold)

# Test determinizmu:   a.hash_wyniku == b.hash_wyniku
# Test idempotencji:   podsumowanie_skoroszytu(cel) identyczne po 2. uruchomieniu
# Test "wejscie nietkniete": _hash(wejscie) nie zmienil sie po generowaniu

def podsumowanie_skoroszytu(sciezka: Path) -> dict:
    """Odcisk struktury. NIE bajt w bajt - ZIP ma zmienne metadane."""
    wb = load_workbook(sciezka, read_only=True)
    try:
        return {"arkusze": tuple(wb.sheetnames),
                "wiersze": {n: sum(1 for _ in wb[n].iter_rows()) - 1
                            for n in wb.sheetnames if n != "_meta"},
                "regul": {n: sum(len(cf.rules) for cf in wb[n].conditional_formatting)
                          for n in wb.sheetnames},
                "hash_wsadu": wb["_meta"]["B6"].value}
    finally:
        wb.close()


# ==================================================================
# 7. META, OSTRZEZENIA, UTRATA FUNKCJI
# ==================================================================
# W PLIKU (arkusz "_meta", sheet_state="hidden"):
#   schemat | wygenerowano | generator | autor | okres |
#   hash_wsadu | wierszy | ostrzezen | <lista ostrzezen>
# W wb.properties: title, creator, description, keywords
# W SIDECAR E: raport.meta.json (czytelne BEZ openpyxl)
# W DZIENNIKU: output/audyt.jsonl (append-only, odporne na awarie)

NIEZAPISYWALNE = {              # -> ALERT (WARNING)
    "xl/sparklineGroups": "sparkline'y (rozszerzenie x14)",
    "xl/slicers": "slicery",
    "xl/timelines": "timeline'y",
    "xl/pivotCache": "cache tabel przestawnych",
    "xl/model": "model danych (Power Pivot)",
}
CZESCIOWE = {                   # -> WARNING (zapisane czesciowo)
    "xl/drawings/": "obrazy i ksztalty",
    "xl/charts/": "wykresy (formatowanie moze byc uproszczone)",
    "xl/vbaProject.bin": "makra (tylko z keep_vba=True)",
}
# Skontroluj PRZED zapisem:  zipfile.ZipFile(x).namelist()  <- TANIO


# ==================================================================
# 8. BEZPIECZENSTWO I WYDAJNOSC JAKO DECYZJE
# ==================================================================
# SANITACJA NA WEJSCIU DO DOMENY (nie przy zapisie!):
ZNAKI_WYZWALAJACE = ("=", "+", "-", "@", "\t", "\r")

def przyjmij_tekst_zewnetrzny(wartosc, *, pole: str) -> str:
    # ODRZUCAMY (nie modyfikujemy!) - nie zmieniamy danych klienta
    ...

# WYDAJNOSC - wybierz wzorzec PRZED pisaniem:
#   synchroniczny w request/response   | maly raport, natychmiastowa odp.
#   strumien BytesIO (seek(0)!)        | brak pliku posredniego
#   zadanie w tle + plik + link        | duze raporty, retry, e-mail
# MIME dla xlsx:
#   application/vnd.openxmlformats-officedocument.spreadsheetml.sheet


# ==================================================================
# 9. CHECKLISTA PRZEGLADU (10 pytan) - kazde "nie" = nie wdrazaj
# ==================================================================
#  1. Czy moge zmienic format wyjscia bez dotykania domeny?
#  2. Czy moge przetestowac domene bez plikow i openpyxl?
#  3. Czy wiem, co strace, jesli plik wejsciowy byl z Excela?
#  4. Czy zapis jest atomowy i nie nadpisuje oryginalu przy bledzie?
#  5. Czy wiem, co sie dzieje, gdy plik jest otwarty w Excelu?
#  6. Czy te same dane daja ten sam plik (poza znacznikami czasu)?
#  7. Czy mam budzet wydajnosciowy i plan, gdy bedzie przekroczony?
#  8. Czy sanitacja jest na wejsciu do domeny, nie przy zapisie?
#  9. Czy kazdy plik niesie wersje, hash wsadu, ostrzezenia?
# 10. Czy wiem, jak wykryje utrate funkcji pliku (sparkline, makra)?
```

## 9. Co dalej

Skończyłeś część IV. Popatrz, co masz:

**Warstwy z regułą**: domena bez openpyxl, aplikacja bez openpyxl, infrastruktura z openpyxl w **jednym** katalogu. Test nie jest intuicją, a `grep`-em — i to jest najważniejsza zmiana w sposobie myślenia: **architektura nie jest „ładna", jest sprawdzalna**.

**Porty i adaptery**: `EksporterRaportu`, trzy implementacje, `EksporterFake` jako najważniejszy plik infrastruktury. Umiesz przetestować **cały** raport bez pliku i bez openpyxl — nie dlatego, że „daje radę", ale dlatego, że **domena nie ma z openpyxl nic wspólnego**.

**Transakcja zapisu**: `mkstemp(dir=cel.parent)` → `wb.save(tmp)` → weryfikacja → backup → `os.replace` → sprzątanie. Znasz dwie pułapki, o których nie mówi żaden tutorial: **cross-device link przy `mkstemp` bez `dir=`** i **`PermissionError` na Windows przy otwartym pliku**.

**Konfigurację z granicą**: cztery poziomy, zamknięta paleta reguł, `grep -n "if \|and \|or \|not " config/*` jako test na „nowy język w konfiguracji".

**Determinizm, wersjonowanie, obserwowalność, bezpieczeństwo na wejściu, raport jako projekcja** — pięć rzeczy, które odróżniają „skrypt, który działa u mnie" od „system, który działa u wszystkich".

**I jeszcze jedno, czego chcę, żebyś nie zapomniał.** Wszystkie te warstwy, porty i transakcje istnieją **wyłącznie** po to, żeby odpowiedzieć na dwa pytania, które nic nie mają wspólnego z kodem:

1. **„Co się stanie, gdy coś pójdzie źle?"** — bo w systemie produkcyjnym **zawsze** coś idzie źle: plik otwarty w Excelu, błędny wsad, brak miejsca na dysku, użytkownik z pustym okresem, plik szablonu ze sparkline'ami, użytkownik z polskimi znakami w CSV. Twoja architektura to **odpowiedź na te pytania zapisana w kodzie** — z komunikatami, z atomowością, z backupem, z alertami.
2. **„Czy da się to zmienić w przyszłym roku?"** — bo wymagania **zawsze** się zmieniają. Nowy format wyjścia, nowy kanał dystrybucji, nowy schemat kolumn, nowa reguła, nowa wersja szablonu. Architektura to **sposób na to, żeby te zmiany nie były przepisaniem projektu**.

---

W module 27 poskładasz to wszystko w **jeden projekt**. Nie będzie nowego API, nie będzie nowych wzorców — będzie **sprawdzenie**, czy naprawdę umiesz to zrobić samodzielnie, od specyfikacji do uruchomienia. Powiem Ci od razu, co Cię czeka:

- **Specyfikacja biznesowa** generatora raportów sprzedażowych: dziesięć wymagań funkcjonalnych i zestaw wymagań niefunkcjonalnych (200 tys. wierszy < 60 s, pamięć < 500 MB, atomowy zapis z backupem, idempotencja, praca w CI bez Excela, pokrycie testami logiki domenowej).
- **Architektura**: `domain/`, `application/`, `infrastructure/excel/` — dokładnie ta struktura z tego modułu, ale zbudowana **przez Ciebie**, z Twoimi nazwami i Twoją konfiguracją.
- **Dziesięć kamieni milowych** — od minimalnego pliku po pełną wersję. **Po każdym kroku działający kod i przechodzący test.** To nie jest „projekt na koniec kursu", to jest **proces**, w którym nie wolno ci przeskakiwać kroków.
- **Rubryka oceny**: poprawność danych, poprawność typów, brak przecieku abstrakcji, wydajność, bezpieczeństwo, testy, dokumentacja.
- **Warianty rozszerzenia**: PDF przez LibreOffice, wysyłka e-mail, wersja webowa ze `StreamingResponse`, porównanie z `xlsxwriter` przy samych wykresach.
- **Refleksja końcowa**: mapa „problem → moduł kursu", do którego wracasz.

Zanim tam pójdziesz, zrób **jedną** rzecz. Weź dowolny raport, który generujesz dzisiaj — nawet jednorazowy skrypt na 30 linijek — i zadaj mu **trzy pytania** z tej checklisty:

1. **Czy mogę zmienić format wyjścia bez dotykania logiki?** (Nie? To `openpyxl` jest w logice.)
2. **Czy wiem, co się stanie, gdy plik jest otwarty w Excelu?** (Nie? To masz pułapkę 4.)
3. **Czy ten sam wsad daje ten sam plik?** (Nie? To masz coś w sekcji 2.8.)

Trzy pytania. Jeżeli odpowiedziałeś „tak" na wszystkie, to ten moduł nie dał Ci niczego nowego — i **to jest dobrze**, bo znaczy, że myślałeś tak już wcześniej. Jeżeli odpowiedziałeś „nie" na któreś, to masz **plan** na moduł 27: napraw to w swoim projekcie, tymi konkretnymi narzędziami, których właśnie się nauczyłeś.