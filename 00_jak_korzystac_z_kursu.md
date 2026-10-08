# Moduł 00 — Jak korzystać z tego kursu

> **Część:** 0 — Fundamenty i model mentalny · **Poziom:** ⭐ · **Wymaga:** nic (moduł startowy)

## 0. W tym module nauczysz się

- Rozumiesz, dla kogo jest ten kurs i którą ścieżką powinieneś iść: „laik”, „programista” czy „analityk”.
- Wiesz, jak czytać moduły, co oznaczają poziomy ⭐, ⭐⭐, ⭐⭐⭐ i dlaczego kolejność modułów ma znaczenie.
- Umiesz utworzyć **wirtualne środowisko** (venv), zainstalować `openpyxl` razem z opcjonalnymi `pillow` i `lxml` oraz **sprawdzić**, czy biblioteka naprawdę je widzi.
- Rozumiesz, jak zorganizować katalog projektu kursu (`examples/`, `data/`, `output/`, `templates/`) i dlaczego to nie jest „zbędna biurokracja”.
- Znasz **zasadę bezpieczeństwa obowiązującą w całym kursie** — nigdy nie nadpisujemy pliku źródłowego — i wiesz, dlaczego jest to ważniejsze, niż się wydaje.
- Widzisz całą mapę kursu i rozumiesz, po co jest w nim aż 29 modułów.
- Wykonujesz pierwszy „dymny test” (smoke test): tworzysz skoroszyt, zapisujesz go i wczytujesz z powrotem — tylko po to, żeby potwierdzić, że środowisko działa.

## 1. Intuicja i analogia

Wyobraź sobie, że zapisałeś się na kurs stolarstwa.

Pierwszy dzień nie wygląda tak, że dostajesz piłę i heblujesz deskę. Pierwszy dzień wygląda tak: ktoś oprowadza Cię po warsztacie. Pokazuje, gdzie leżą narzędzia, gdzie jest wyciąg, gdzie apteczka, gdzie gaśnica, jak włącza się i wyłącza maszyny. Tłumaczy, dlaczego nie wolno odkładać dłuta ostrzem do siebie i dlaczego nie wkłada się rąk w pobliże piły, gdy ta jeszcze się obraca. Dopiero potem — i to nie od razu — dostajesz pierwszy kawałek drewna.

Ten kurs jest zbudowany dokładnie w ten sposób.

- **Moduły 00–03 to Twój dzień zapoznawczy i zasady BHP.** Nie budujesz tu jeszcze niczego trwałego. Ustawiasz środowisko, dowiadujesz się, czym openpyxl jest, czym nie jest, jak wygląda plik `.xlsx` od środka i jak openpyxl „myśli” o danych. To są moduły niemal pozbawione spektakularnych efektów — i właśnie dlatego są najważniejsze.
- **Moduły 04–17 to nauka narzędzi.** Poznajesz piłę, wiertarkę, hebel i frezarkę: tworzenie skoroszytów, wpisywanie danych, formuły, wiersze i kolumny, arkusze, formatowanie, tabele, formatowanie warunkowe, wykresy, obrazy, walidację i ochronę. Każde narzędzie osobno, na małych przykładach.
- **Moduły 18–21 to zasady bezpieczeństwa przy prawdziwej pracy.** Moduł 18 („modyfikacja istniejących plików bez utraty zawartości”) jest odpowiednikiem zasady „nigdy nie wsadzaj rąk w ruchomą piłę”. To najważniejszy moduł praktyczny w całym kursie. Moduły 19–21 to praca pod presją (duże pliki), kontrola jakości (testy) i bezpieczeństwo (dane od nieznanych osób).
- **Moduły 22–27 to projektowanie mebli na zamówienie.** Dopiero tutaj ktoś prosi Cię o szafę, która ma pasować do konkretnego pokoju, mieć konkretne wymiary i wyglądać dobrze. To moment, w którym proste skrypty przestają wystarczać i trzeba zacząć myśleć o strukturze kodu.
- **Moduł 28 to tablica narzędziowa na ścianie.** Ściągawka, do której wracasz po miesiącach, gdy zapomnisz, jak nazywa się ten jeden parametr.

Druga analogia, tym razem o tym, dlaczego moduł 00 istnieje.

Kurs prawa jazdy zaczyna się od tego, jak ustawić fotel, lusterka i pasy — a nie od wyjechania na autostradę. Nikt rozsądny nie uczy się prowadzić samochodu, którego nie umie uruchomić i nie wie, gdzie ma hamulec. W programowaniu jest identycznie: najczęstszym powodem, dla którego początkujący „nie może uruchomić tutoriala”, nie jest brak wiedzy o bibliotece. Jest nim to, że kod działa w innym Pythonie, niż myśli autor, albo że brakuje pakietu, o którym tutorial milczy, bo u autora po prostu był zainstalowany od lat.

Ten moduł jest właśnie o tym: o ustawieniu fotela i lusterek. Nie zobaczysz tu ani jednego efektownego wykresu. Zobaczysz za to coś cenniejszego — pewność, że cokolwiek napiszesz w kolejnych 28 modułach, uruchomi się u Ciebie dokładnie tak, jak zostało to opisane.

Trzecia analogia, krótka, ale kluczowa — do zrozumienia, po co jest wirtualne środowisko.

Globalna instalacja pakietów to **wspólna kuchnia w akademiku**. Wstawiasz swój słoik na półkę, ale każdy inny mieszkaniec może ten słoik przesunąć, podmienić albo opróżnić. Kiedy wrócisz po tygodniu i coś nie zadziała, nie będziesz wiedział, czy zepsułeś to Ty, czy sąsiad z pokoju obok.

Wirtualne środowisko to **własna szafka z zamkiem**. Wkładasz tam dokładnie te słoiki, które chcesz, w takich wersjach, jakie chcesz, i nikt ich nie dotknie. Gdy projekt umrze, wyrzucasz całą szafkę i nic po nim nie zostaje. To wszystko, co musisz na razie wiedzieć — szczegóły za chwilę.

## 2. Teoria

### 2.1. Dla kogo jest ten kurs

Ten kurs jest napisany dla trzech grup czytelników i wszystkie trzy powinny go przeczytać od początku do końca — różnica polega na tym, ile czasu poświęcą konkretnym modułom.

**Ścieżka „laik”.** Umiesz napisać pętlę, wiesz, co to zmienna, lista i funkcja, ale pojęcia takie jak „iterator”, „niemutowalność” czy „kontekst” nic Ci nie mówią. To całkowicie w porządku — kurs zakłada taką właśnie sytuację i **każde** pojęcie programistyczne tłumaczy od zera, w jednym zdaniu i na analogii z życia codziennego. Twoja rada: nie przeskakuj modułów 01–03. Wydają się teoretyczne, ale to one sprawiają, że moduły 18 i 19 staną się oczywiste, a nie magiczne. Czytaj wolniej, uruchamiaj każdy przykład, otwieraj każdy wygenerowany plik w Excelu i patrz na niego własnymi oczami.

**Ścieżka „programista”.** Piszesz kod zawodowo, znasz wzorce projektowe, testy jednostkowe i pracę z repozytoriami. Możesz szybciej przejść przez moduły 00–04, ale **nie pomijaj modułu 18**. To moduł, w którym openpyxl przestaje być „biblioteką do pisania plików”, a staje się narzędziem o konkretnych, ostrych krawędziach — i to jest wiedza, której nie znajdziesz w pośpiechu w żadnym tutorialu. Twoje zainteresowanie będzie rosło w modułach 20 i 22–26; możesz do nich zajrzeć wcześniej, ale wróć do części o modyfikacji istniejących plików.

**Ścieżka „analityk”.** Pracujesz z danymi, znasz Excela lepiej niż większość programistów, ale kodujesz okazjonalnie. Dla Ciebie najcenniejsze będą moduły 05–14 (typy danych, formuły, formatowanie, tabela, formatowanie warunkowe) oraz moduł 19 (duże pliki). Moduły o wzorcach projektowych możesz potraktować jako „przeczytam, gdy będę budować coś większego”.

Jedno zdanie wspólne dla wszystkich: **ten kurs nie uczy Excela.** Zakłada, że wiesz, co to jest arkusz, wiersz, komórka i formuła. Kurs uczy, jak tym wszystkim sterować z Pythona.

### 2.2. Jak czytać moduły — anatomia jednego pliku

Każdy moduł kursu ma dokładnie dziewięć sekcji i zawsze w tej samej kolejności. Warto wiedzieć, po co one są, żebyś nie czytał ich bezmyślnie od góry do dołu w poszukiwaniu kodu.

| Sekcja | Po co istnieje | Jak z niej korzystać |
|--------|----------------|----------------------|
| **Nagłówek + 0** | Orientacja: część kursu, poziom, wymagania, lista celów | Przeczytaj, żeby wiedzieć, czy jesteś w dobrym miejscu |
| **1. Intuicja i analogia** | Zbudowanie obrazu w głowie, zanim pojawi się API | Czytaj raz, spokojnie. To sekcja „dlaczego”, nie „jak” |
| **2. Teoria** | Właściwe wyjaśnienie mechanizmów, słownictwo, granice | To tu się uczysz. Wracaj tu, gdy coś nie działa |
| **3. Przykłady krok po kroku** | Pełny, uruchamialny kod z komentarzami | Przepisz ręcznie i uruchom. Nie kopiuj-wklej |
| **4. Anatomia API** | Tabela: metoda/klasa → co robi → parametry → uwagi | Ściągawka robocza; wracaj tu w trakcie pisania |
| **5. Ćwiczenia** | Sprawdzenie się trzema poziomami trudności | Rób je. Czytanie kodu to nie umiejętność |
| **6. Typowe błędy i pułapki** | Objaw → przyczyna → naprawa | Sekcja ratunkowa. Zajrzyj tu, zanim zapytasz Google |
| **7. Podsumowanie w 5 punktach** | Kompresja modułu do pięciu zdań | Przeczytaj, gdy wracasz do modułu po tygodniach |
| **8. Ściągawka** | Blok kodu z najważniejszymi wywołaniami | Do skopiowania do własnego notatnika |
| **9. Co dalej** | Łącznik do kolejnego modułu | Orientacja w kursie |

Jedna ważna uwaga o sekcji 3: **przepisuj kod ręcznie, nie kopiuj**. Kopiowanie nie zapisuje niczego w pamięci mięśniowej, a właśnie ta pamięć sprawia, że po dwóch tygodniach napiszesz `ws["A1"] = 5` bez zastanowienia. Jeśli naprawdę musisz kopiować — kopiuj, ale potem **zmodyfikuj** każdy przykład choćby o jedną rzecz. Zmiana wymusza zrozumienie.

### 2.3. Co znaczą poziomy ⭐

| Poziom | Znaczenie | Czego się spodziewać |
|--------|-----------|----------------------|
| ⭐ | Fundament | Mało API, dużo wyjaśnień, wszystko działa od razu |
| ⭐⭐ | Praktyka | Kilka powiązanych klas, więcej pułapek, trzeba już rozumieć model z modułu 03 |
| ⭐⭐⭐ | Zaawansowane | Wiele możliwości na raz, ograniczenia biblioteki, wymóg świadomych decyzji |

Poziomy nie oznaczają „ważności” — moduł 03 jest oznaczony ⭐, a jest najważniejszy w całym kursie. Oznaczają **trudność wejścia**, nie wagę.

### 2.4. Trzy pojęcia, bez których nie ruszymy: interpreter, pip, środowisko

Zanim zainstalujesz cokolwiek, potrzebujesz trzech pojęć. Wyjaśniam je tak, jakbyś nigdy nie programował zawodowo.

**Interpreter Pythona** to program, który czyta Twój plik `.py` i wykonuje jego instrukcje. Analogia: to **kucharz**. Ty dajesz przepis (kod), on go wykonuje. Problem polega na tym, że w jednym komputerze może być zainstalowanych kilku różnych kucharzy — np. `python` z systemu, `python3.11` z Homebrew, `python` z pakietu instalacyjnego, `python` w środku wirtualnego środowiska. I **każdy z nich ma własny zestaw dostępnych składników**. Kiedy mówisz „zainstalowałem openpyxl, a Python go nie widzi”, prawie zawsze znaczy to: „dałem składnik kucharzowi A, a gotuje kucharz B”.

**pip** to instalator pakietów — menadżer składników. Analogia: **dostawca zakupów**. Mówisz „poproszę openpyxl w wersji 3.1.5” i dostawca dostarcza go do wybranej kuchni. Kluczowe: dostawca też jest przypisany do konkretnego kucharza. `pip` uruchomiony samodzielnie może trafić do innego Pythona, niż myślisz. Dlatego w tym kursie konsekwentnie używamy zapisu `python -m pip install ...`, a nie `pip install ...`. Różnica jest ogromna — wyjaśniam ją za chwilę.

**Środowisko (environment)** to zestaw pakietów widzianych przez konkretnego interpreter. Wirtualne środowisko (venv) to odizolowane środowisko przypisane do jednego projektu.

### 2.5. Dlaczego `python -m pip`, a nie `pip`

To jedna z tych drobnych decyzji, które eliminują połowę problemów początkujących.

Gdy wpisujesz `pip install openpyxl`, system szuka w zmiennej `PATH` pierwszego lepszego programu o nazwie `pip` i uruchamia go. Ten `pip` należy do **jakiegoś** Pythona — ale niekoniecznie do tego, którym uruchomisz później swój skrypt.

Gdy wpisujesz `python -m pip install openpyxl`, mówisz jednoznacznie: „weź **ten konkretny** interpreter, który u mnie nazywa się `python`, i uruchom w nim moduł `pip`”. Nie ma miejsca na pomyłkę. Analogia: `pip install` to „poproszę doktora” na dużym szpitalnym korytarzu. `python -m pip install` to „poproszę doktora Kowalskiego, tego z trzeciego piętra”.

W tym kursie **zawsze** używamy drugiej formy. Jeśli pracujesz na systemie, gdzie interpreter nazywa się `python3` (typowe na Linuksie i macOS), używaj `python3 -m pip`.

### 2.6. Wirtualne środowisko — instalacja krok po kroku

Pokażę pełną procedurę dla Windows i dla macOS/Linux. Nie musisz rozumieć każdego słowa z komend — poniżej rozbieram je na części.

**Windows (PowerShell albo cmd):**

```bat
mkdir C:\projekty\kurs-openpyxl
cd C:\projekty\kurs-openpyxl
py -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install openpyxl pillow lxml
```

**macOS / Linux (bash, zsh):**

```bash
mkdir -p ~/projekty/kurs-openpyxl
cd ~/projekty/kurs-openpyxl
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install openpyxl pillow lxml
```

Rozbierzmy to na części:

- `mkdir` (make directory) — tworzy katalog. `mkdir -p` na Linuksie/macOS tworzy także katalogi nadrzędne, jeśli nie istnieją. Analogia: tworzysz pudełko na projekt.
- `cd` (change directory) — przechodzisz do tego pudełka. Od tego momentu wszystkie operacje dotyczą tego katalogu.
- `python3 -m venv .venv` — tworzy wirtualne środowisko w podkatalogu o nazwie `.venv`. Uwaga na kropkę z przodu: to konwencja oznaczająca katalog „ukryty”, którego zwykle nie chcemy oglądać w menedżerze plików. Nazwa `.venv` jest standardem — używaj jej, bo narzędzia (edytory, CI) domyślnie jej szukają.
- `source .venv/bin/activate` (macOS/Linux) lub `.venv\Scripts\activate` (Windows) — **aktywuje** środowisko. Od tego momentu Twoja powłoka systemowa tak ustawia zmienne, że słowo `python` wskazuje na interpreter ze `.venv`, a `pip` — na tego właściwego. W wierszu zachęty zwykle pojawia się `(.venv)`. Analogia: założyłeś fartuch konkretnego stanowiska pracy.
- `python -m pip install --upgrade pip` — aktualizuje samego instalatora. To dobra higiena: stare `pip` potrafi nie obsłużyć nowych formatów pakietów. Nie jest to obowiązkowe, ale warto zrobić raz na początku.
- `python -m pip install openpyxl pillow lxml` — instaluje trzy pakiety jedną komendą.

**Kiedy musisz aktywować środowisko?** Za każdym razem, gdy otwierasz **nowy** terminal i chcesz pracować w tym projekcie. Aktywacja nie jest trwała — dotyczy tylko bieżącego okna. To normalne i zdrowe: jedno okno terminala = jedno środowisko.

**Kiedy nie musisz aktywować?** Gdy używasz edytora, który sam wybrał `.venv` jako interpreter (np. VS Code po kliknięciu „Select Interpreter” albo PyCharm po skonfigurowaniu projektu). Edytor uruchamia wtedy kod z Twojego środowiska, nawet jeśli terminal obok jest „czysty”.

### 2.7. Dlaczego `pillow` i `lxml` są opcjonalne, ale warto je mieć od razu

`openpyxl` ma dwa opcjonalne zależności, które sprawiają, że część funkcji albo działa, albo **wywala wyjątek**. Nagłówek „instaluję tylko openpyxl” brzmi kusząco, ale w praktyce prowadzi do zaskakujących błędów w połowie kursu.

| Pakiet | Do czego służy | Co się dzieje bez niego |
|--------|----------------|--------------------------|
| **`Pillow`** (moduł `PIL`) | Ładowanie, skalowanie i zapisywanie obrazów w komórkach (moduł 16) | `ws.add_image(...)` kończy się błędem importu — praca z obrazami jest niemożliwa |
| **`lxml`** | Szybszy parser XML niż wbudowany w Pythona | Wszystko działa, ale duże pliki (moduły 19) przetwarzają się wyraźnie wolniej |

Oba są **opcjonalne w sensie technicznym** — kurs nie wymaga ich w modułach 01–15. Ale są **obowiązkowe w sensie praktycznym**: 95% realnych raportów albo zawiera logo/obraz, albo ma więcej niż kilka tysięcy wierszy. Zainstaluj je od razu, żeby uniknąć „dlaczego u mnie nie działa” w modułach 16 i 19.

Warto zapamiętać jeszcze jedną rzecz: `openpyxl` **nie zgłasza żadnego ostrzeżenia**, gdy `lxml` brakuje. Po prostu działa wolniej i nikt Ci o tym nie powie. Musisz wiedzieć, że dodałeś `lxml` po to, by mieć przyspieszenie — i sprawdzić, że faktycznie jest.

### 2.8. Jak sprawdzić, czy `openpyxl` widzi `pillow` i `lxml`

Najprostszy, wiarygodny sposób to zapytać Pythona wprost — i w tym samym interpreterze, w którym będziesz uruchamiać kursowe przykłady.

```python
import sys
import importlib.util

# Najpierw dowiadujemy się, KTÓRY to Python. Bez tego cała diagnoza jest zgadywaniem.
print("Wersja Pythona:", sys.version)
print("Ścieżka interpretera:", sys.executable)

# openpyxl musi być zaimportowany, żebyśmy mogli poznać jego wersję.
import openpyxl
print("openpyxl:", openpyxl.__version__)

# find_spec tylko SPRAWDZA, czy pakiet da się znaleźć. Nie importuje go,
# więc jest szybkie i nie wpływa na stan programu.
for nazwa_pakietu, opis in (("PIL", "pillow  - obrazy"),
                            ("lxml",  "lxml    - wydajność XML")):
    znaleziony = importlib.util.find_spec(nazwa_pakietu) is not None
    status = "ZNALEZIONY" if znaleziony else "BRAK"
    print(f"{opis}: {status}")
```

**Co się dzieje w pamięci:** `import openpyxl` ładuje cały pakiet do pamięci i wykonuje kod inicjujący bibliotekę. `find_spec` tego nie robi — tylko sprawdza, czy interpreter potrafi zlokalizować pakiet na dysku. Dzięki temu pytanie „czy mam pillow?” nie kosztuje nas importu całej biblioteki graficznej.

**Co trafi do pliku:** nic. Ten skrypt nic nie zapisuje. To celowe — moduł 00 nie tworzy plików Excel. Pierwszy plik pojawi się w module 04.

Przykładowy wynik (u Ciebie liczby i ścieżki będą inne — to normalne):

```text
Wersja Pythona: 3.12.4 (main, ...)
Ścieżka interpretera: /home/anna/projekty/kurs-openpyxl/.venv/bin/python
openpyxl: 3.1.5
pillow  - obrazy: ZNALEZIONY
lxml    - wydajność XML: ZNALEZIONY
```

Jeśli w linii `Ścieżka interpretera` **nie widzisz** `.venv`, to znak, że uruchamiasz nie ten Python — wróć do sekcji 2.6 albo do pułapki numer 1 w sekcji 6.

### 2.9. Struktura katalogu projektu kursu

Kod kursu trzymamy w czterech katalogach. Wygląda to tak:

```text
kurs-openpyxl/
├── .venv/            # wirtualne środowisko (nie dotykamy, nie wersjonujemy)
├── .gitignore        # pliki, których nie chcemy w repozytorium
├── requirements.txt  # lista zainstalowanych pakietów i wersji
├── data/             # pliki WEJŚCIOWE: dane, które czytamy
├── templates/        # skoroszyty-szablony (wzorce, których nie nadpisujemy)
├── output/           # wszystko, co wygeneruje kod kursu
└── examples/         # nasze skrypty .py, jeden na przykład
```

Po co aż tyle podziałów? Bo mieszanie danych wejściowych z wynikami to najprostsza droga do katastrofy. Przypomnij sobie wspólną kuchnię z akademika: jeśli nie masz osobnej półki na „to, co wchodzi” i „to, co wychodzi”, prędzej czy później zjesz składnik, który miał być na obiad, albo podasz na stół surowe ciasto zamiast upieczonego.

- **`data/`** — pliki wejściowe. Traktuj ten katalog jako **tylko do czytania**. Reguła: nic w `data/` nigdy się nie zmienia w trakcie uruchomienia skryptu.
- **`templates/`** — skoroszyty wzorcowe (np. firmowy raport z logo). Też tylko do czytania. Zasada: szablon nigdy nie jest nadpisywany (wrócimy do tego w module 23).
- **`output/`** — wyjście. Tu wolno pisać do woli. Cały katalog możesz w każdej chwili skasować i wygenerować od nowa.
- **`examples/`** — skrypty Pythona. Jeden plik = jeden przykład. Nazwy: `00_smoke_test.py`, `04_pierwszy_skoroszyt.py` i tak dalej.

Dwa pliki porządkowe:

**`.gitignore`** — jeśli używasz gita, powinien zawierać co najmniej:

```text
.venv/
__pycache__/
*.pyc
output/
```

`output/` ignorujemy, bo to wyniki — generowalne, często duże, niepotrzebne w repozytorium. `.venv/` ignorujemy zawsze, bo to pliki binarne zależne od Twojego systemu.

**`requirements.txt`** — lista pakietów, po której ktoś inny (albo Ty za pół roku, na innym komputerze) odtworzy dokładnie to samo środowisko:

```text
openpyxl==3.1.5
pillow>=10.0
lxml>=5.0
```

Do jej wygenerowania użyjesz `python -m pip freeze > requirements.txt`. Znaki `==` oznaczają „dokładnie ta wersja”, znaki `>=` — „ta albo nowsza”. Kurs przypina `openpyxl==3.1.5`, żeby przykłady działały identycznie u wszystkich.

### 2.10. Konwencja kursu: każdy przykład zapisuje do `output/`

W całym kursie obowiązuje jedna konwencja, którą warto zapamiętać już teraz:

> **Każdy przykład, który coś tworzy, zapisuje wynik w katalogu `output/`, a żaden przykład nie modyfikuje niczego w `data/` ani `templates/`.**

To nie jest ozdobnik. To jest metoda nauki.

openpyxl jest biblioteką, której efektów **nie da się zrozumieć, patrząc tylko na kod**. Możesz napisać trzydzieści linii definiujące formatowanie warunkowe i nie mieć pojęcia, czy zrobiłeś to dobrze — dopóki nie otworzysz pliku w Excelu i nie zobaczysz kolorów. Dlatego każdy przykład w tym kursie kończy się plikiem, który możesz otworzyć, kliknąć i obejrzeć. **Otwieraj te pliki.** Dosłownie, w prawdziwym Excelu (albo LibreOffice Calc, jeśli nie masz Excela — działa równie dobrze do oglądania efektów).

Drugi powód jest bezpieczeństwowy: pokazuję różne rzeczy na plikach, które **łatwo zniszczyć** (moduł 18). Gdyby przykłady operowały na `data/`, już pierwszego dnia mógłbyś stracić dane bezpowrotnie.

### 2.11. Zasada bezpieczeństwa obowiązująca w całym kursie

To zdanie zasługuje na osobną sekcję, bo wraca w całym kursie i jest ważniejsze od wszystkich trików z API:

> **Nigdy nie nadpisujemy pliku źródłowego.**

Konsekwencje praktyczne:

1. Zawsze czytamy z `data/` lub `templates/`, a **piszemy do `output/`** (pod inną nazwą lub w innym katalogu).
2. Zanim uruchomisz jakikolwiek skrypt modyfikujący istniejący plik, zrób kopię. Nie „bo na pewno się uda”. Kopię. Uruchom zgłoszenie w myślach: „jeśli skrypt zniszczy plik, czy mogę go odtworzyć?”. Jeśli odpowiedź brzmi „nie” — nie uruchamiaj go na oryginale.
3. W module 18 nauczysz się **wzorca bezpiecznego zapisu**: zapisu do pliku tymczasowego, atomowej podmiany i backupu z datą w nazwie. To standard, którego będziesz używać w każdej prawdziwej pracy.

Analogia: masz tylko jeden oryginał aktu notarialnego. Zanim wręczysz go komukolwiek, robisz kopię i pracujesz na kopii. Nikt rozsądny nie oddaje oryginału do adnotacji bez zabezpieczenia. W pracy z plikami Excel — zwłaszcza firmowymi raportami „z prawdziwymi danymi” — obwiązuje dokładnie ta sama zasada.

### 2.12. Roadmapa i mapa kursu

Poniżej pełna mapa. Nie musisz jej analizować — przejrzyj ją raz, żeby wiedzieć, dokąd zmierzasz.

| # | Moduł | Temat | Poziom |
|---|-------|-------|--------|
| **00** | `00_jak_korzystac_z_kursu.md` | Organizacja kursu, środowisko, instalacja, roadmapa | — |
| **01** | `01_co_to_jest_openpyxl.md` | Czym openpyxl jest i czym nie jest; wybór biblioteki | ⭐ |
| **02** | `02_anatomia_pliku_xlsx.md` | `.xlsx` = ZIP + XML (OOXML); mapa części pliku | ⭐ |
| **03** | `03_model_mentalny_openpyxl.md` | Graf obiektów, indeksowanie od 1, cykl życia, tryby pracy | ⭐ |
| **04** | `04_pierwszy_skoroszyt.md` | Tworzenie, zapis, odczyt, ścieżki, właściwości dokumentu | ⭐ |
| **05** | `05_komorki_i_wartosci.md` | Typy Python ↔ Excel, daty, `None`, limity, błędy | ⭐ |
| **06** | `06_formuly.md` | Formuły jako tekst, brak silnika obliczeń, `data_only` | ⭐⭐ |
| **07** | `07_wiersze_kolumny_zakresy.md` | Wymiary, iteracja, `append`, insert/delete, adresy | ⭐⭐ |
| **08** | `08_arkusze.md` | Tworzenie, kolejność, kopiowanie, widoczność | ⭐⭐ |
| **09** | `09_formatowanie_podstawy.md` | Model stylu, niemutowalność, `Font`, `Fill`, `NamedStyle` | ⭐⭐ |
| **10** | `10_formaty_liczb_i_daty.md` | `number_format`, kody formatów, waluty, lokalizacja | ⭐⭐ |
| **11** | `11_wyglad_arkusza.md` | Szerokości, scalanie, `freeze_panes`, druk i PDF | ⭐⭐ |
| **12** | `12_tabele.md` | `Table`, style tabel, filtry, odwołania strukturalne | ⭐⭐ |
| **13** | `13_formatowanie_warunkowe_podstawy.md` | Reguły, `dxf`, priorytet, `CellIsRule`, `FormulaRule` | ⭐⭐⭐ |
| **14** | `14_formatowanie_warunkowe_zaawansowane.md` | `ColorScale`, `DataBar`, `IconSet`, reguły własne | ⭐⭐⭐ |
| **15** | `15_wykresy.md` | `Reference`, `Series`, typy wykresów, dwie osie | ⭐⭐⭐ |
| **16** | `16_obrazy_hiperlinki_komentarze.md` | `Image`, anchory, komentarze, hiperlinki | ⭐⭐ |
| **17** | `17_walidacja_danych_i_ochrona.md` | `DataValidation`, listy rozwijane, ochrona arkusza | ⭐⭐⭐ |
| **18** | `18_modyfikacja_istniejacych_plikow.md` | **Edycja bez utraty treści, inwentarz ryzyk, atomowy zapis** | ⭐⭐⭐ |
| **19** | `19_duze_pliki_i_wydajnosc.md` | `read_only`, `write_only`, `lxml`, pamięć, streaming | ⭐⭐⭐ |
| **20** | `20_testowanie_i_jakosc.md` | `pytest`, `BytesIO`, testy w pamięci, CI | ⭐⭐⭐ |
| **21** | `21_bezpieczenstwo.md` | Formula injection, ryzyka ZIP/XML, walidacja wejścia | ⭐⭐⭐ |
| **22** | `22_wzorce_wprowadzenie.md` | Dlaczego skrypty Excel stają się „big ball of mud” | ⭐⭐ |
| **23** | `23_wzorce_konstrukcyjne.md` | Factory, Builder, Prototype, Registry, Flyweight | ⭐⭐⭐ |
| **24** | `24_wzorce_strukturalne.md` | Facade, Adapter, Decorator, Proxy, Composite | ⭐⭐⭐ |
| **25** | `25_wzorce_behawioralne.md` | Strategy, Template Method, Command, Visitor, Specification | ⭐⭐⭐ |
| **26** | `26_architektura.md` | Warstwy, Ports & Adapters, Repository, konfiguracja | ⭐⭐⭐ |
| **27** | `27_projekt_koncowy.md` | Generator raportów sprzedażowych, od specyfikacji do kodu | ⭐⭐⭐ |
| **28** | `28_sciagawka_faq_antywzorce.md` | Ściągawka, FAQ, antywzorce, glosariusz | ⭐ |

Trzy punkty orientacyjne na tej mapie, które warto zapamiętać już teraz:

- **Moduł 03** to najważniejszy moduł kursu. Jeśli go zrozumiesz, moduły 18 i 19 będą dla Ciebie naturalne. Jeśli go przeskoczysz, będziesz się dziwił, dlaczego openpyxl „gubi” dane.
- **Moduł 18** to moduł, po którym zaczniesz pracować bezpiecznie na cudzych plikach. To najważniejszy moduł praktyczny.
- **Moduły 22–26** to jedna spójna historia: ten sam przykład (raport sprzedaży) ewoluuje od 40-linijkowego skryptu do trójwarstwowej aplikacji z portem `EksporterRaportu`. Czytaj je razem, nie osobno.

## 3. Przykłady krok po kroku

Poniższe przykłady układają się w logiczną sekwencję: najpierw tworzymy środowisko, potem weryfikujemy je, potem budujemy strukturę katalogów, potem zaglądamy do środka pliku `.xlsx`, a na końcu robimy pierwszy „dymny test” — czyli sprawdzamy, że cały łańcuch od wpisania kodu do otwarcia pliku w Excelu działa. Każdy przykład jest kompletny: zawiera importy i można go uruchomić samodzielnie.

### Przykład 1 — weryfikacja środowiska (skrypt `examples/00_sprawdz_srodowisko.py`)

Zapisz poniższy kod w pliku `examples/00_sprawdz_srodowisko.py` i uruchom go komendą `python examples/00_sprawdz_srodowisko.py` — **z aktywowanym środowiskiem `.venv`**.

```python
"""Sprawdza, czy środowisko kursu jest gotowe do pracy.

Uruchom:  python examples/00_sprawdz_srodowisko.py
Oczekiwane: wersja Pythona, ścieżka interpretera, wersja openpyxl,
status pakietów pillow i lxml.
"""

import sys
import importlib.util


def sprawdz_pakiet(nazwa_modulu: str) -> bool:
    """Zwraca True, jeśli pakiet jest importowalny w tym interpreterze.

    find_spec() nie importuje pakietu - tylko sprawdza, czy jest znajdowalny.
    """
    return importlib.util.find_spec(nazwa_modulu) is not None


def main() -> None:
    # --- CZĘŚĆ 1: kim jesteśmy? Bez tego cała diagnoza to zgadywanie. ---
    print("=" * 60)
    print("ŚRODOWISKO")
    print("=" * 60)
    print("Wersja Pythona:      ", sys.version.split()[0])
    print("Ścieżka interpretera:", sys.executable)

    # --- CZĘŚĆ 2: czy openpyxl jest dostępny i w jakiej wersji? ---
    try:
        import openpyxl
    except ImportError:
        print("\nBŁĄD: nie mogę zaimportować openpyxl.")
        print("Zainstaluj go: python -m pip install openpyxl==3.1.5")
        return

    print("Wersja openpyxl:     ", openpyxl.__version__)

    # --- CZĘŚĆ 3: pakiety opcjonalne. ---
    print("\nPAKIETY OPCJONALNE")
    for nazwa, opis in (("PIL", "pillow (obrazy)"),
                        ("lxml", "lxml (wydajność XML)")):
        status = "ZNALEZIONY" if sprawdz_pakiet(nazwa) else "BRAK"
        print(f"  {opis:<25} {status}")

    # --- CZĘŚĆ 4: ostrzeżenie, jeśli ścieżka wygląda podejrzanie. ---
    if ".venv" not in sys.executable:
        print("\nUWAGA: uruchamiasz kod spoza wirtualnego środowiska.")
        print("Aktywuj środowisko (.venv) przed dalszą pracą.")


if __name__ == "__main__":
    main()
```

**Co robi ten kod, linijka po linijce:**

- `sys.executable` to pełna ścieżka do interpretera, który **aktualnie wykonuje ten plik**. To najważniejsza linia diagnostyczna w całym kursie. Jeśli widzisz tu `.venv`, wszystko jest w porządku. Jeśli nie — uruchamiasz nie ten Python.
- `sys.version.split()[0]` wyciąga z długiego opisu wersji pierwszy wyraz, czyli np. `3.12.4`. Dzielimy tekst po białych znakach i bierzemy pierwszy element — bo `sys.version` zwraca coś w rodzaju `"3.12.4 (main, Jun 7 2024, ...)"`.
- `try: import openpyxl except ImportError:` — jeśli pakietu nie ma, przechwytujemy wyjątek i kończymy z konkretną instrukcją. To dobra praktyka: komunikat błędu powinien mówić **co zrobić**, nie tylko **co się stało**.
- Blok `if __name__ == "__main__":` — to standardowy wzorzec. Wyjaśnienie w ramce poniżej.

<details>
<summary><strong>Dygresja dla laików: co to jest <code>if __name__ == "__main__"</code>?</strong></summary>

Gdy Python uruchamia plik bezpośrednio (np. `python examples/00_sprawdz_srodowisko.py`), ustawia specjalnej zmiennej o nazwie `__name__` wartość `"__main__"`. Gdyby ten sam plik został **zaimportowany** przez inny plik, `__name__` miałoby inną wartość (nazwę modułu). Wzorzec ten oznacza więc: „wykonaj `main()` tylko wtedy, gdy ten plik jest uruchamiany bezpośrednio, a nie wtedy, gdy ktoś go importuje”. Dzięki temu ten sam plik może być jednocześnie programem i biblioteką — i nie uruchomi niczego niechcący przy imporcie.

Analogia: masz skrypt wypowiedzi. Jeśli jesteś mówcą na scenie (plik główny), wygłaszasz ją. Jeśli jesteś autorem, którego fragment cytuje ktoś inny (import), nie wygłaszasz całej mowy — tylko użyczasz fragmentu.

</details>

**Co się dzieje w pamięci:** `import openpyxl` ładuje moduł do pamięci i wykonuje kod inicjujący. `find_spec` nie importuje niczego — tylko sprawdza, czy pakiet znajduje się na ścieżce wyszukiwania Pythona.

**Co trafi do pliku:** nic. Ten skrypt nie tworzy żadnych plików — jest wyłącznie diagnostyczny.

### Przykład 2 — utworzenie struktury katalogów (skrypt `examples/00_struktura.py`)

Skoro obowiązuje nas konwencja czterech katalogów, warto umieć je utworzyć programowo — i przy okazji poznać `pathlib`, który w całym kursie będzie podstawą pracy ze ścieżkami.

```python
"""Tworzy strukturę katalogów kursu.

Uruchom:  python examples/00_struktura.py
Efekt:    katalogi data/ templates/ output/ examples/ obok tego pliku.
"""

from pathlib import Path

# __file__ to ścieżka TEGO pliku. .resolve() sprowadza ją do postaci
# bezwzględnej (nie zawiera ".." ani dowiązań symbolicznych).
# .parent to katalog, w którym leży ten plik.
# W naszym układzie jest to katalog 'examples/', więc wychodzimy piętro wyżej.
ROOT = Path(__file__).resolve().parent.parent

KATALOGI = ("data", "templates", "output", "examples")


def utworz_strukture(root: Path) -> None:
    for nazwa in KATALOGI:
        katalog = root / nazwa                      # operator "/" łączy ścieżki!
        katalog.mkdir(parents=True, exist_ok=True)  # tworzy, ale nie krzyczy, jeśli już istnieje
        print(f"gotowe: {katalog}")


def main() -> None:
    print(f"Katalog główny projektu: {ROOT}\n")
    utworz_strukture(ROOT)
    print("\nStruktura gotowa.")


if __name__ == "__main__":
    main()
```

**Trzy rzeczy warte zapamiętania:**

1. `Path(__file__).resolve().parent.parent` — to najbezpieczniejszy sposób ustalenia, gdzie leży projekt. Działa niezależnie od tego, z jakiego katalogu uruchomisz skrypt. **Nigdy** nie zakładaj, że „uruchomię z katalogu projektu i ścieżki względne zadziałają”. To najczęstsza przyczyna „u mnie działa, na serwerze nie działa”.
2. Operator `/` na obiektach `Path` łączy ścieżki. `Path("a") / "b"` daje `Path("a/b")` (na Windows: `a\b`). To wygodniejsze i bezpieczniejsze niż ręczne sklejanie tekstu z ukośnikami.
3. `mkdir(parents=True, exist_ok=True)` — `parents=True` tworzy brakujące katalogi nadrzędne, `exist_ok=True` pozwala wywołać funkcję wielokrotnie bez błędu. Bez `exist_ok=True` drugie uruchomienie skryptu kończy się wyjątkiem `FileExistsError`. Analogia: „zrób kanapkę, chyba że już jest” zamiast „zrób kanapkę i krzycz, jeśli jest”.

**Co się dzieje w pamięci:** obiekt `Path` to lekka reprezentacja ścieżki; operacje `mkdir` faktycznie wykonują wywołania systemowe.

**Co trafi do pliku:** powstają **katalogi**, nie pliki. Żadnych plików Excel jeszcze nie tworzymy.

### Przykład 3 — zwiastun modułu 02: zaglądamy do środka `.xlsx` (skrypt `examples/00_zajrzyj_do_xlsx.py`)

To najbardziej zaskakujący fragment modułu 00. Plik `.xlsx` **nie jest** plikiem binarnym w takim sensie, w jakim jest nim zdjęcie czy film. To zwykłe archiwum ZIP — dokładnie takie samo, jak plik `.zip`, który tworzysz prawym przyciskiem myszy.

Sprawdź to sam, bez Pythona: skopiuj dowolny plik `.xlsx` do `data/`, zmień jego rozszerzenie z `.xlsx` na `.zip` i spróbuj go rozpakować wbudowanym narzędziem systemowym (Windows: „Wyodrębnij wszystko”; macOS: dwuklik; Linux: `unzip`). Zobaczysz katalog z plikami XML. Jeśli to zrobiłeś, wróciłeś właśnie z przyszłości — dokładnie tym zajmiemy się w module 02, ale teraz zrobimy to samo programowo.

```python
"""Pokazuje, że .xlsx to archiwum ZIP z plikami XML.

Uruchom:  python examples/00_zajrzyj_do_xlsx.py data/dowolny_plik.xlsx
Uwaga:    jeśli nie masz pod ręką pliku .xlsx, ten skrypt sam utworzy
          prosty plik testowy w output/ i obejrzy jego zawartość.
"""

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"


def utworz_plik_testowy() -> Path:
    """Tworzy minimalny skoroszyt, żeby przykład działał od pierwszego uruchomienia."""
    from openpyxl import Workbook  # import lokalny: potrzebny tylko w tej funkcji

    OUTPUT.mkdir(parents=True, exist_ok=True)
    sciezka = OUTPUT / "00_do_podgladu.xlsx"

    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Tajemnica tego pliku zaraz się wyda."
    ws["A2"] = 42
    wb.save(sciezka)

    return sciezka


def wypisz_zawartosc(sciezka: Path) -> None:
    print(f"Oglądam: {sciezka}\n")
    # 'with' zapewnia, że archiwum zostanie zamknięte, nawet gdy poleci wyjątek.
    with zipfile.ZipFile(sciezka) as archiwum:
        nazwy = archiwum.namelist()

    print(f"W środku znajduje się {len(nazwy)} części:\n")
    for nazwa in nazwy:
        print("  ", nazwa)


def main() -> None:
    if len(sys.argv) > 1:
        sciezka = Path(sys.argv[1])
        if not sciezka.exists():
            print(f"Nie ma takiego pliku: {sciezka}")
            return
    else:
        print("Nie podano pliku - tworzę testowy.\n")
        sciezka = utworz_plik_testowy()

    wypisz_zawartosc(sciezka)


if __name__ == "__main__":
    main()
```

Przykładowy wynik:

```text
Oglądam: /home/anna/projekty/kurs-openpyxl/output/00_do_podgladu.xlsx

W środku znajduje się 8 części:

   [Content_Types].xml
   _rels/.rels
   xl/workbook.xml
   xl/_rels/workbook.xml.rels
   xl/worksheets/sheet1.xml
   xl/styles.xml
   docProps/core.xml
   docProps/app.xml
```

Przeczytaj tę listę uważnie. To **mapa przyszłego modułu 02**. Już teraz widzisz, że:

- `xl/workbook.xml` to spis arkuszy (nasz „zeszyt”),
- `xl/worksheets/sheet1.xml` to zawartość pierwszego arkusza („kartka”),
- `xl/styles.xml` to definicje wyglądu („legenda kolorów”),
- `docProps/core.xml` to metadane (autor, tytuł),
- `[Content_Types].xml` i `_rels/` to spis i powiązania pakietu.

Zauważ też, że **nie ma osobnego pliku na każdą komórkę**. Cały arkusz to jeden plik XML. To bardzo ważna obserwacja — wrócimy do niej przy wydajności.

**Co się dzieje w pamięci:** `zipfile.ZipFile` otwiera archiwum i czyta jedynie spis części (central directory) — nie rozpakowuje zawartości do pamięci. Dlatego ta operacja jest błyskawiczna nawet dla dużych plików.

**Co trafi do pliku:** w tym przykładzie tworzymy plik `output/00_do_podgladu.xlsx` **tylko po to**, by mieć na czym pracować. To pierwszy plik Excel w kursie — choć formalnie należy do modułu 04.

### Przykład 4 — pierwszy „dymny test” (skrypt `examples/00_smoke_test.py`)

„Smoke test” (test dymny) to termin z realnej inżynierii: po złożeniu urządzenia włącza się je na chwilę i sprawdza, czy **nie leci z niego dym**. Nie testujemy poprawności każdego obwodu — tylko to, że urządzenie w ogóle działa. W naszym przypadku sprawdzimy, że openpyxl umie: utworzyć skoroszyt, zapisać go i wczytać z powrotem.

```python
"""Pierwszy test dymny kursu.

Cel: potwierdzić, że łańcuch (utworzenie -> zapis -> odczyt) działa
w Twoim środowisku. To NIE jest jeszcze nauka API - szczegóły w module 04.
"""

from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"

SCIEZKA = OUTPUT / "00_smoke_test.xlsx"


def main() -> None:
    OUTPUT.mkdir(parents=True, exist_ok=True)

    # --- KROK 1: utworzenie nowego, pustego skoroszytu w pamięci.
    wb = Workbook()
    ws = wb.active            # domyślny arkusz - więcej o tym w module 04
    ws.title = "Test"

    ws["A1"] = "openpyxl działa"
    ws["A2"] = 42
    ws["A3"] = True

    # --- KROK 2: zapis na dysk. Bez tego w pamięci jest tylko obiekt,
    #             a na dysku - nic.
    wb.save(SCIEZKA)
    print(f"Zapisano: {SCIEZKA}")

    # --- KROK 3: odczyt z dysku do NOWEGO obiektu.
    wb2 = load_workbook(SCIEZKA)
    ws2 = wb2["Test"]

    print("Odczytane wartości:")
    for adres in ("A1", "A2", "A3"):
        wartosc = ws2[adres].value
        print(f"  {adres}: {wartosc!r}  (typ: {type(wartosc).__name__})")


if __name__ == "__main__":
    main()
```

Oczekiwany wynik:

```text
Zapisano: /home/anna/projekty/kurs-openpyxl/output/00_smoke_test.xlsx
Odczytane wartości:
  A1: 'openpyxl działa'  (typ: str)
  A2: 42  (typ: int)
  A3: True  (typ: bool)
```

Zwróć uwagę na `{wartosc!r}` — to formatowanie w f-stringu, które wypisuje wartość w **reprezentacji** (z apostrofami wokół tekstu, bez zamiany `True` na coś innego). To świetny nawyk diagnostyczny: od razu widać, czy `42` to liczba (`42`), czy tekst (`'42'`). Ta różnica będzie jednym z najważniejszych problemów modułu 05.

**Co się dzieje w pamięci:** `Workbook()` tworzy **cały model skoroszytu w pamięci**. W tym momencie na dysku nie ma jeszcze nic. `wb.save(SCIEZKA)` serializuje ten model do pliku. `load_workbook(SCIEZKA)` tworzy **nowy, niezależny** model w pamięci na podstawie pliku. `wb` i `wb2` to dwa różne obiekty, choć opisują ten sam zbiór danych.

**Co trafi do pliku:** plik `output/00_smoke_test.xlsx` zawierający jeden arkusz o nazwie `Test`, trzy komórki (`A1` tekst, `A2` liczba, `A3` wartość logiczna) i standardowe metadane.

Otwórz ten plik w Excelu. Zobaczysz dokładnie to, co wypisał skrypt. To pierwszy moment w kursie, w którym możesz dosłownie **otworzyć efekt swojej pracy** — i dlatego ten test jest tu, a nie w module 04. Chodzi o to, żebyś od pierwszej chwili zobaczył, że „kod Pythona” i „plik Excel” to dwa końce tego samego procesu.

### Przykład 5 — porządki: co robić z zamieszaniem w wersjach (skrypt `examples/00_freeze.py`)

Ostatni przykład jest organizacyjny, ale rozwiązuje realny problem: gdy za tydzień zapragniesz odtworzyć działające środowisko, musisz wiedzieć, co dokładnie było zainstalowane.

```python
"""Zapisuje listę zainstalowanych pakietów do requirements.txt.

Uruchom:  python examples/00_freeze.py
Efekt:    plik requirements.txt w katalogu głównym projektu.
"""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WYJSCIE = ROOT / "requirements.txt"


def main() -> None:
    # Uruchamiamy pip w TYM SAMYM interpreterze, który wykonuje ten skrypt.
    # sys.executable to gwarancja, że nie pomylimy środowisk.
    wynik = subprocess.run(
        [sys.executable, "-m", "pip", "freeze"],
        capture_output=True,
        text=True,
        check=True,
    )
    WYJSCIE.write_text(wynik.stdout, encoding="utf-8")
    print(f"Zapisano {len(wynik.stdout.splitlines())} pakietów do {WYJSCIE}")


if __name__ == "__main__":
    main()
```

Dwie rzeczy, na które warto zwrócić uwagę:

- `subprocess.run` uruchamia inny program (tu: `pip`) z wnętrza Pythona. `capture_output=True` przechwytuje jego wyjście, `text=True` zwraca je jako tekst zamiast bajtów, `check=True` rzuca wyjątek, jeśli program zakończy się błędem.
- Używamy `sys.executable` — czyli **tego samego** interpretera, który wykonuje nasz skrypt. To gwarancja, że notujemy wersje pakietów faktycznie widzianych przez nasz kod, a nie przez jakiś systemowy Python.

**Co się dzieje w pamięci:** wynik `pip freeze` trafia do zmiennej jako tekst.

**Co trafi do pliku:** `requirements.txt` w katalogu głównym projektu. To plik, który możesz wrzucić do repozytorium i dzięki któremu ktoś inny odtworzy Twoje środowisko komendą `python -m pip install -r requirements.txt`.

## 4. Anatomia API

Ten moduł nie uczy API openpyxl — i to celowo. Poniższa tabela zbiera **narzędzia i polecenia**, które wprowadziliśmy, żebyś miał je w jednym miejscu. To swoista „skrzynka z narzędziami na start”.

| Element | Co robi | Parametry / składnia | Uwagi |
|---------|---------|----------------------|-------|
| `python -m venv .venv` | Tworzy wirtualne środowisko w katalogu `.venv` | ścieżka katalogu | Nazwa `.venv` to konwencja; narzędzia szukają jej domyślnie |
| `.venv\Scripts\activate` / `source .venv/bin/activate` | Aktywuje środowisko w bieżącej powłoce | — | Windows vs macOS/Linux; działa tylko w danym oknie terminala |
| `python -m pip install <pakiet>` | Instaluje pakiet w **tym konkretnym** interpreterze | `nazwa==wersja`, `-r requirements.txt` | Zawsze używaj formy `python -m pip`, nie samego `pip` |
| `python -m pip install --upgrade pip` | Aktualizuje instalator | — | Warto raz na początku projektu |
| `python -m pip freeze` | Wypisuje zainstalowane pakiety z wersjami | `> requirements.txt` | Podstawa odtwarzalności środowiska |
| `sys.version` | Pełny opis wersji Pythona | — | `sys.version.split()[0]` daje sam numer, np. `3.12.4` |
| `sys.executable` | Pełna ścieżka do interpretera | — | **Najważniejsza linia diagnostyczna** |
| `importlib.util.find_spec("PIL")` | Sprawdza, czy pakiet jest znajdowalny | nazwa modułu | Nie importuje pakietu; zwraca `None`, gdy brak |
| `openpyxl.__version__` | Wersja openpyxl | — | Oczekujemy `3.1.5` |
| `import openpyxl` | Ładuje bibliotekę do pamięci | — | Sam import niczego nie tworzy |
| `Path(__file__).resolve().parent` | Katalog, w którym leży plik | — | Fundament bezpiecznych ścieżek |
| `Path.mkdir(parents=True, exist_ok=True)` | Tworzy katalog (wielokrotnie bezpiecznie) | `parents`, `exist_ok` | Bez `exist_ok` drugie wywołanie rzuca wyjątek |
| `Path / "nazwa"` | Łączy ścieżki | operator `/` | Przenośne między systemami |
| `zipfile.ZipFile(path)` | Otwiera archiwum ZIP (także `.xlsx`) | `with` zalecane | `.namelist()` zwraca listę części |
| `Workbook()` | Nowy, pusty skoroszyt w pamięci | — | Nic nie zapisuje na dysk |
| `wb.save(path)` | Zapisuje model do pliku `.xlsx` | ścieżka lub obiekt plikopodobny | **Nadpisuje istniejący plik bez ostrzeżenia** |
| `load_workbook(path)` | Wczytuje plik do modelu w pamięci | ścieżka; flagi opisane w module 18 | Tworzy niezależny obiekt |
| `wb.active` | Domyślny/aktywny arkusz | — | Nowy skoroszyt ma jeden arkusz |
| `ws["A1"].value` | Odczyt wartości z komórki | notacja A1 | Zapis: `ws["A1"] = ...` |
| `if __name__ == "__main__":` | Wykonaj `main()` tylko przy bezpośrednim uruchomieniu | — | Wzorzec, nie biblioteka |

**Uwaga praktyczna:** kolumny „Parametry” celowo nie rozwijam tu do pełnych sygnatur — to zadanie modułów tematycznych. Tu chodzi o to, żebyś rozpoznał narzędzie, nie żeby znał wszystkie jego ostrza.

## 5. Ćwiczenia

### 🟢 Rozgrzewka — środowisko działa

**Zadanie 1.** Utwórz wirtualne środowisko w dowolnym katalogu, zainstaluj `openpyxl` (przypięte do `3.1.5`), `pillow` i `lxml`. Następnie uruchom skrypt z przykładu 1 i upewnij się, że:
- w `Ścieżka interpretera` widnieje `.venv`,
- `Wersja openpyxl` to `3.1.5`,
- oba pakiety opcjonalne są oznaczone jako `ZNALEZIONY`.

**Zadanie 2.** Uruchom skrypt z przykładu 4 (dymny test). Otwórz wygenerowany plik `output/00_smoke_test.xlsx` w Excelu (albo LibreOffice Calc) i sprawdź, czy widzisz w nim:
- arkusz o nazwie `Test`,
- tekst `openpyxl działa` w komórce `A1`,
- liczbę `42` w `A2`, wyrównaną do prawej (bo to liczba, nie tekst),
- `TRUE` w `A3`.

**Zadanie 3.** Otwórz wygenerowany plik w zwykłym edytorze tekstu (VS Code, Notatnik). Zauważ, że widzisz „śmieci”. To dobrze — to znak, że plik jest archiwum ZIP, a nie tekstem. Wróć do tego wspomnienia w module 02.

### 🟡 Warsztat — diagnoza i struktura

**Zadanie 4.** Celowo zepsuj środowisko: otwórz nowy terminal, **nie aktywuj** `.venv` i uruchom `python examples/00_sprawdz_srodowisko.py`. Zapisz, co się stało. Czy:
- (a) skrypt się uruchomił i pokazał inny interpreter,
- (b) skrypt wyrzucił `ModuleNotFoundError: No module named 'openpyxl'`?

Wynik zależy od tego, czy globalnie masz zainstalowany openpyxl. **Wyciągnij z tego wniosek dla siebie** i zapisz go w notatce: „skrypt uruchamiam zawsze z aktywnym środowiskiem, bo…”.

**Zadanie 5.** Zmodyfikuj skrypt z przykładu 2 tak, aby tworzył dodatkowy katalog `output/archiwum/2026/`. Użyj `mkdir(parents=True, exist_ok=True)`. Następnie uruchom go dwa razy i upewnij się, że drugie uruchomienie nie kończy się błędem.

**Zadanie 6.** Napisz skrypt, który liczy, ile części (plików XML) jest w każdym pliku `.xlsx` z katalogu `data/`. Dla każdego pliku wypisz jego nazwę i liczbę części. Podpowiedź: użyj `zipfile.ZipFile(path).namelist()`.

### 🔴 Wyzwanie — od zera do zautomatyzowanego setupu

**Zadanie 7.** Napisz skrypt `examples/00_setup.py`, który:
1. Sprawdza wersję Pythona i przerywa pracę, jeśli jest starsza niż 3.9 (użyj `sys.version_info`).
2. Sprawdza, czy istnieje katalog `.venv`. Jeśli nie — wypisuje czytelny komunikat z instrukcją, jak go utworzyć.
3. Sprawdza, czy zainstalowany jest `openpyxl` we wersji `3.1.x` (użyj `openpyxl.__version__` i porównania tekstowego lub `tuple(map(int, ...))`).
4. Sprawdza obecność `PIL` i `lxml` i wypisuje rekomendację, jeśli któregoś brakuje.
5. Tworzy brakujące katalogi projektu (`data`, `templates`, `output`, `examples`).
6. Zapisuje raport ze sprawdzenia do pliku `output/00_setup_report.txt` z datą i godziną.

Skrypt ma kończyć się kodem wyjścia 0 przy sukcesie i 1 przy wykryciu problemu blokującego (użyj `sys.exit(1)`). Kody wyjścia przydadzą się w module 20, gdy będziemy uruchamiać testy w automatyzacji.

<details>
<summary><strong>Szkic rozwiązania zadania 7</strong></summary>

```python
"""Sprawdza i raportuje gotowość środowiska do kursu."""

import importlib.util
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "output"


def sprawdz_python() -> tuple[bool, str]:
    ok = sys.version_info >= (3, 9)
    opis = f"Python {sys.version.split()[0]} ({sys.executable})"
    return ok, opis


def sprawdz_openpyxl() -> tuple[bool, str]:
    if importlib.util.find_spec("openpyxl") is None:
        return False, "openpyxl: BRAK - zainstaluj: python -m pip install openpyxl==3.1.5"
    import openpyxl
    ok = openpyxl.__version__.startswith("3.1.")
    return ok, f"openpyxl: {openpyxl.__version__}"


def sprawdz_opcjonalne() -> list[tuple[bool, str]]:
    wyniki = []
    for nazwa, opis in (("PIL", "pillow - obrazy"), ("lxml", "lxml - wydajność")):
        ok = importlib.util.find_spec(nazwa) is not None
        status = "ZNALEZIONY" if ok else "BRAK (zalecana instalacja)"
        wyniki.append((ok, f"{opis}: {status}"))
    return wyniki


def glowny() -> int:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    linie: list[str] = []
    blad = False

    ok, opis = sprawdz_python()
    blad |= not ok
    linie.append(("[OK] " if ok else "[BLAD] ") + opis)

    ok, opis = sprawdz_openpyxl()
    blad |= not ok
    linie.append(("[OK] " if ok else "[BLAD] ") + opis)

    for ok, opis in sprawdz_opcjonalne():
        linie.append(("[OK] " if ok else "[UWAGA] ") + opis)

    # Katalogi
    for nazwa in ("data", "templates", "output", "examples"):
        (ROOT / nazwa).mkdir(parents=True, exist_ok=True)
    linie.append("[OK] Struktura katalogów gotowa")

    raport = (
        f"Raport środowiska - {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        + "-" * 60 + "\n"
        + "\n".join(linie) + "\n"
    )
    (OUTPUT / "00_setup_report.txt").write_text(raport, encoding="utf-8")
    print(raport)
    return 1 if blad else 0


if __name__ == "__main__":
    sys.exit(glowny())
```

Kluczowe decyzje w tym rozwiązaniu:
- `importlib.util.find_spec` przed importem — pozwala na czytelny komunikat zamiast wyjątku.
- `import openpyxl` wewnątrz funkcji — import dopiero po sprawdzeniu, że pakiet istnieje.
- Raport zapisywany do `output/` — zgodnie z konwencją kursu; sam raport nie jest plikiem Excel, więc nie koliduje z modułem 04.
- `sys.exit(glowny())` — kod wyjścia ustawiany na podstawie wyniku; gotowe pod CI.

</details>

## 6. Typowe błędy i pułapki

Poniższe pułapki to najczęstsze przyczyny tego, że „tutorial nie działa”. Zapisuję je w formacie **objaw → przyczyna → naprawa**, żebyś mógł je szybko diagnozować w przyszłości.

**1. „Zainstalowałem openpyxl, ale Python go nie widzi” (objaw) → pakiet trafił do innego Pythona (przyczyna) → zawsze instaluj przez `python -m pip`, zawsze w aktywnym środowisku; sprawdź `sys.executable` (naprawa).**

To pułapka numer jeden. Klasyczny scenariusz: student instaluje `pip install openpyxl` (systemowo), potem aktywuje venv i uruchamia skrypt — i dostaje `ModuleNotFoundError`. Przyczyna: instalacja poszła do globalnego Pythona, a kod uruchamia się w venv, gdzie pakietu nie ma. Rozwiązanie: po aktywacji venv **zawsze** używaj `python -m pip install`. I dodaj do nawyku sprawdzanie `sys.executable` na początku każdej sesji, w której coś nie działa.

**2. „Uruchamiam skrypt bez aktywacji — raz działa, raz nie” (objaw) → zależność od tego, czy globalny Python ma przypadkiem ten sam pakiet (przyczyna) → aktywuj środowisko przed każdą sesją pracy; w edytorze wskaż interpreter `.venv` (naprawa).**

Ten błąd jest groźniejszy od poprzedniego, bo daje nieregularne objawy. Dziś działa (bo globalnie masz openpyxl), jutro na innym komputerze nie działa. Dlatego w skrypcie z przykładu 1 dodałem ostrzeżenie o braku `.venv` w `sys.executable`. To drobny kod, a oszczędza godziny.

**3. „Mój stary projekt czyta pliki `.xls` i po aktualizacji przestał działać” (objaw) → konflikt ze starymi bibliotekami `xlrd`/`xlwt` oraz zmiana API w ekosystemie (przyczyna) → rozdziel środowiska dla starego i nowego projektu; do `.xls` nadal potrzebujesz `xlrd`, do `.xlsx` — openpyxl (naprawa).**

Wyjaśnienie techniczne: `openpyxl` **nie obsługuje formatu `.xls`** (Excel 97–2003). To osobny, binarny format, obsługiwany przez `xlrd` (odczyt). Dodatkowo nowsze wersje `xlrd` **przestały** obsługiwać `.xlsx` — więc nie jest już zamiennikiem dla openpyxl. Jeśli w Twoim środowisku współistnieją stare skrypty z `xlrd`/`xlwt` i nowe z openpyxl, najlepiej trzymać je w **osobnych wirtualnych środowiskach**. Analogia: nie próbujesz naprawiać zabytkowego zegara mechanicznego nowoczesnym kluczem dynamometrycznym; potrzebne są dwa zestawy narzędzi w dwóch szufladach.

**4. „Tutorial, który czytam, używa `ws.cell(row, col)` w inny sposób albo w ogóle nie działa” (objaw) → tutorial napisany dla openpyxl 2.x lub innej wersji (przyczyna) → sprawdź wersję biblioteki w tutorialu; dopasuj do 3.1.x lub świadomie zaadaptuj (naprawa).**

Wersja 3.1 wprowadziła kilka zmian, które łamią stare przykłady. Najczęściej spotykane: usunięto moduł `openpyxl.pandas` (dawniej `openpyxl.reader.pandas`), zmieniono API nazw zdefiniowanych (`DefinedName` zamiast `NamedRange`), usunięto stare metody `get_named_range`/`add_named_range`/`remove_named_range`. Jeśli tutorial z 2018 roku używa którejś z tych form, to nie Twoja wina, że nie działa — to kwestia wersji. Zasada: **zawsze sprawdź wersję, której dotyczy materiał**, zanim zaczniesz debugować swój kod.

**5. „Polskie znaki w wyjściu z konsoli wyglądają jak krzaki, plik zapisuje się bez ogonków” (objaw) → kodowanie konsoli lub brak `encoding="utf-8"` przy zapisie plików tekstowych (przyczyna) → ustaw `encoding="utf-8"` przy każdym zapisie pliku tekstowego; na Windows starsze konsole domyślnie używają innego kodowania (naprawa).**

Ten kurs pisze wszystko w UTF-8 i konsekwentnie używa `encoding="utf-8"` przy `write_text`. Pliki `.xlsx` są odporne na ten problem — kodowanie obsługuje format. Problem dotyczy plików tekstowych (`.txt`, `.csv`) oraz konsoli. Jeśli na Windows widzisz „krzaki” w konsoli, sprawdź, czy używasz Windows Terminal / PowerShell 7 (obsługują UTF-8 domyślnie) lub ustaw stronę kodową `chcp 65001`.

**6. „Skrypt działa, gdy uruchamiam go z katalogu projektu, ale nie działa z innego katalogu” (objaw) → ścieżki względne zależne od bieżącego katalogu roboczego (przyczyna) → zawsze buduj ścieżki od `Path(__file__).resolve().parent` (naprawa).**

Bieżący katalog roboczy (CWD) to miejsce, z którego uruchomiono skrypt — nie musi być tym samym co katalog, w którym leży skrypt. To rozróżnienie myli prawie wszystkich początkujących. Dlatego wzorzec `ROOT = Path(__file__).resolve().parent.parent` jest w tym kursie wszechobecny. Analogia: dokumenty w Twoim biurze są zawsze w tej samej szufladzie, niezależnie od tego, jakim wejściem wszedłeś do budynku.

**7. „Dodałem obraz do arkusza i dostałem błąd importu” (objaw) → brak pakietu `pillow` (przyczyna) → `python -m pip install pillow` (naprawa).**

openpyxl umie manipulować obrazami tylko wtedy, gdy jest obecny `Pillow`. Gdy go brak, `ws.add_image` zgłasza `ImportError: You must install Pillow to fetch image objects`. To zwiastun modułu 16 — ale lepiej mieć ten pakiet zainstalowany już teraz, żeby nie tracić czasu na diagnozę później.

**8. „Testowy skrypt nadpisał mój cenny plik” (objaw) → zapis do ścieżki wejściowej (przyczyna) → zasada: czytaj z `data/`, pisz do `output/`; nigdy nie zapisuj pod nazwą pliku wejściowego (naprawa).**

To pułapka, która boli w prawdziwej pracy. `wb.save(path)` **nadpisuje istniejący plik bez żadnego ostrzeżenia**. Nie ma potwierdzenia, nie ma kopii zapasowej. Jeśli zapiszesz skoroszyt wczytany z `templates/firmowy_raport.xlsx` pod tą samą nazwą, szablon przepadnie. Dlatego cały kurs trzyma żelazną zasadę oddzielenia `data`/`templates` od `output`. W module 18 nauczysz się dodatkowo pisać do pliku tymczasowego i atomowo podmieniać plik docelowy — ale zasada „nigdy nie nadpisuj źródła” pozostaje.

**9. „Mam kilka wersji Pythona i nie wiem, którą wybrać” (objaw) → brak świadomości, który interpreter jest uruchamiany (przyczyna) → ustal jedną wersję dla kursu (3.9+) i pilnuj, żeby `python` w Twoim terminalu wskazywał na interpreter z `.venv` (naprawa).**

Jeśli po uruchomieniu `python --version` widzisz starą wersję, a po aktywacji `.venv` — nową, to znaczy, że system działa poprawnie. Ale jeśli `python` w ogóle nie wskazuje na venv, użyj pełnej ścieżki: `.venv/bin/python` (macOS/Linux) lub `.venv\Scripts\python` (Windows). To najbezpieczniejszy sposób w skryptach automatyzacji.

**10. „Skrypt zapisuje plik, ale Excel mówi, że plik jest uszkodzony” (objaw) → plik został zapisany z rozszerzeniem `.xlsm` bez zachowania makr (`keep_vba=True`) albo z niezgodnym rozszerzeniem (przyczyna) → używaj `.xlsx` dla zwykłych plików, a dla `.xlsm` wczytuj z `keep_vba=True` (naprawa).**

Ten błąd wystąpi na późniejszym etapie, ale warto go znać już teraz, bo dotyczy samej konwencji nazywania plików. Jeśli plik zawiera makra, ich zachowanie wymaga jawnej flagi — bez niej Excel zgłosi uszkodzenie struktury. Szczegóły w module 18.

**11. „Otwieram plik, edytuję go ręcznie w Excelu i uruchamiam skrypt — tracę zmiany” (objaw) → plik był otwarty w Excelu w trakcie zapisu skryptu lub zapis nastąpił, gdy Excel trzymał blokadę (przyczyna) → zamykaj plik w Excelu przed uruchomieniem skryptu, który go zapisuje (naprawa).**

Nowoczesne systemy nie zawsze blokują zapis w sposób widoczny dla skryptu, więc zapis „obok otwartego Excela” może zakończyć się uszkodzeniem lub nadpisaniem bez ostrzeżenia. W module 18 wrócimy do tego jako do jednej z pułapek produkcyjnych. Na razie zasada jest prosta: **jedno narzędzie naraz**.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Kurs to warsztat, nie zestaw trików.** Moduły 00–03 to BHP i poznanie narzędzi; dopiero potem budujesz. Jeśli pominiesz fundament, większość późniejszych ostrzeżeń będzie dla Ciebie niejasna.

2. **Środowisko jest częścią kodu.** Wirtualne środowisko (`.venv`) to Twój prywatny zestaw narzędzi. Używaj `python -m pip`, sprawdzaj `sys.executable`, zapisuj wersje do `requirements.txt`. Prawie wszystkie problemy „u mnie nie działa” to problemy środowiska, nie biblioteki.

3. **`openpyxl` to narzędzie, a `pillow` i `lxml` to jego opcjonalne kończyny.** Bez `pillow` nie ma obrazów, bez `lxml` jest wolniej — a openpyxl nie powie Ci o tym ani słowem. Zainstaluj oba od razu.

4. **Cztery katalogi, jedna żelazna zasada.** `data/` i `templates/` są tylko do czytania; piszesz wyłącznie do `output/`. **Nigdy nie nadpisuj pliku źródłowego.** Ta zasada nie jest ozdobnikiem — towarzystwo ubezpieczeniowe Twojej pracy.

5. **Plik `.xlsx` to archiwum ZIP z XML-ami.** Widziałeś to już w przykładzie 3. Ta jedna obserwacja wyjaśni połowę ograniczeń openpyxl opisanych w dalszych modułach.

## 8. Ściągawka modułu

```python
# ------------------------------------------------------------------
# 1. ŚRODOWISKO (w powłoce, nie w Pythonie)
# ------------------------------------------------------------------
# python -m venv .venv
# source .venv/bin/activate            # macOS / Linux
# .venv\Scripts\activate               # Windows
# python -m pip install --upgrade pip
# python -m pip install openpyxl==3.1.5 pillow lxml
# python -m pip freeze > requirements.txt

# ------------------------------------------------------------------
# 2. DIAGNOSTYKA (w Pythonie)
# ------------------------------------------------------------------
import sys
import importlib.util

print(sys.version.split()[0])          # np. "3.12.4"
print(sys.executable)                  # ŚCIEŻKA interpretera - kluczowa!
print(".venv" in sys.executable)       # True => jesteśmy w środowisku

import openpyxl
print(openpyxl.__version__)            # oczekiwane "3.1.5"

print(importlib.util.find_spec("PIL"))   # None => brak pillow
print(importlib.util.find_spec("lxml"))  # None => brak lxml

# ------------------------------------------------------------------
# 3. STRUKTURA PROJEKTU
# ------------------------------------------------------------------
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent   # NIE zależ od CWD!
for nazwa in ("data", "templates", "output", "examples"):
    (ROOT / nazwa).mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------------
# 4. PODGLĄD WNĘTRZA .xlsx (zwiastun modułu 02)
# ------------------------------------------------------------------
import zipfile

with zipfile.ZipFile(ROOT / "output" / "plik.xlsx") as archiwum:
    for czesc in archiwum.namelist():
        print(czesc)
# Zwykle zobaczysz m.in.:
#   [Content_Types].xml, _rels/.rels,
#   xl/workbook.xml, xl/worksheets/sheet1.xml,
#   xl/styles.xml, docProps/core.xml

# ------------------------------------------------------------------
# 5. TEST DYMNY (pełne API dopiero w module 04)
# ------------------------------------------------------------------
from openpyxl import Workbook, load_workbook

wb = Workbook()                    # model skoroszytu w PAMIĘCI
ws = wb.active
ws["A1"] = "openpyxl działa"
ws["A2"] = 42
wb.save(ROOT / "output" / "test.xlsx")   # teraz powstaje PLIK

wb2 = load_workbook(ROOT / "output" / "test.xlsx")
print(wb2["Sheet"]["A1"].value)    # 'openpyxl działa'
print(wb2["Sheet"]["A2"].value)    # 42

# ------------------------------------------------------------------
# 6. ŻELAZNE ZASADY KURSU
# ------------------------------------------------------------------
# - czytam z data/ i templates/, piszę do output/
# - NIGDY nie nadpisuję pliku źródłowego
# - zawsze sprawdzam sys.executable, gdy coś nie działa
# - otwieram każdy wygenerowany plik w Excelu i patrzę na niego
```

## 9. Co dalej

Masz działające środowisko, znasz strukturę projektu i wiesz, jak nie zniszczyć sobie pracy. To wszystko, czego trzeba, żeby bezpiecznie wejść w materiał merytoryczny.

W **module 01** odpowiemy na pytanie, które powinno poprzedzać naukę każdej biblioteki: *czym openpyxl właściwie jest i kiedy jest dobrym wyborem, a kiedy złym?* Zobaczysz tabelę porównawczą z `xlsxwriter`, `pandas`, `xlwings` i `pyexcel`, dowiesz się, czego openpyxl nigdy nie zrobi (i dlaczego to nie wada, a świadoma decyzja konstrukcyjna), oraz obalimy najczęstszy mit: że openpyxl umie liczyć formuły.

Przygotuj do modułu 01 jeden plik `.xlsx` w katalogu `data/` — dowolny, nawet wygenerowany przez nasz skrypt z przykładu 3. Będzie nam potrzebny do obserwacji.