Zweryfikuję kluczowe fakty techniczne dla tego modułu: testowanie w pamięci przez `BytesIO`, doświadczenia z typowaniem `mypy` wobec openpyxl, determinizm plików `.xlsx` oraz zachowanie właściwości używanych w asercjach.

Dobre informacje. Zweryfikuję jeszcze kilka szczegółów API, których będę używał w asercjach testowych: iterację po tabelach i formatowaniu warunkowym, zachowanie `print_area` oraz strukturę walidacji.

# Moduł 20 — Testowanie i jakość

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–19

## 0. W tym module nauczysz się

- **Zrozumiesz, dlaczego plik `.xlsx` wydaje się niemożliwy do przetestowania** — i dlaczego to tylko pozór. Po tym module testowanie raportu będzie dla Ciebie tańsze niż otwieranie go w Excelu.
- **Opanujesz `BytesIO` jako „przymierzalnię"**: zapiszesz skoroszyt do pamięci, wczytasz go z powrotem i sprawdzisz — **bez ani jednego bajtu na dysku**, w milisekundach.
- **Zbudujesz pełny warsztat `pytest`**: fixtures, fabrykę skoroszytu, `tmp_path`, asercje na **wartościach, typach, formatach, zakresach i liczbie reguł** — a nie na bajtach.
- **Napiszesz `podsumowanie_skoroszytu()`** — „odcisk palca" pliku, który zamieni nieporównywalny plik binarny w czytelny tekst i stanie się **testem regresji** pilnującym każdej zmiany w raporcie.
- **Poznasz cztery strategie porównywania plików** — i zrozumiesz, dlaczego porównanie bajt w bajt **zawsze** będzie czerwone, choćby kod był doskonały.
- **Wykorzystasz testy właściwościowe** (`hypothesis`) przeciwko czyścicielom danych: zamiast wymyślać przykłady, dasz maszynie wygenerować tysiąc najtrudniejszych.
- **Zamienisz „wydaje się wolne" w twardy budżet** — „100 000 wierszy < 10 s" jako test, który pilnuje wydajności z modułu 19.
- **Uruchomisz cały zestaw testów w CI bez Excela** (bo Excel nie jest do niczego potrzebny), dodasz `ruff`, `mypy`, `pre-commit` i **listę kontrolną „definicja ukończenia raportu"**.
- **Zrozumiesz, dlaczego struktura projektu decyduje o testowalności** — i dlaczego oddzielenie „co ma być w raporcie" od „jak trafia do pliku" jest warunkiem wstępnym wszystkiego, co przyjdzie w modułach 24–26.

## 1. Intuicja i analogia

### 1.1. Hamulce, które pozwalają jechać szybciej

Zacznijmy od pozornego paradoksu.

Kierowca rajdowy zapytany, po co mu hamulce, mógłby odpowiedzieć: „żeby się zatrzymywać". Ale to nieprawda. **Hamulce nie służą do zatrzymywania — służą do tego, żebyś mógł jechać szybko.** Bez hamulców jedziesz wolno, bo każdy zakręt to rosnące ryzyko, którego nie kontrolujesz. Z hamulcami wchodzisz w zakręt szybciej, bo wiesz, że **w każdej chwili możesz wyhamować dokładnie tam, gdzie chcesz**.

Testy w kodzie działają dokładnie tak samo. Brzmi to jak dodatkowa praca — i na krótką metę nią jest. Ale ich prawdziwym celem nie jest „pilnowanie, żeby nic się nie zepsuło". Ich celem jest to, żebyś **miał odwagę zmieniać kod szybko**:

> **Raport bez testów to raport, w którym nie masz odwagi nic zmienić.**

Rozpoznasz taki kod od razu. Leży w repozytorium od dwóch lat. Wszyscy wiedzą, że trzeba poprawić nagłówek i dodać kolumnę, ale nikt tego nie robi — bo „ostatnio, jak Kasia dodała kolumnę, to zniknęły wykresy i klient to zauważył". Nikt nie pamięta, dlaczego zniknęły. Nikt nie wie, co jeszcze może zniknąć. Więc kod **nie ewoluuje**. Nie dlatego, że jest dobry, ale dlatego, że jest **zbyt ryzykowny w dotykaniu**.

Testy zmieniają to w jedną rzecz: **odwagę**. Zmieniasz kod, uruchamiasz `pytest`, dostajesz odpowiedź w trzy sekundy. Jeśli zielono — idziesz dalej. Jeśli czerwono — test powie Ci **co i gdzie** przestało działać, po nazwie. To jest hamulec, który pozwala jechać szybciej.

### 1.2. Dlaczego przyjazd do pliku `.xlsx` wygląda jak droga w mgle

A teraz problem, który sprawia, że ludzie nie testują raportów. Postaw się w sytuacji, gdzie masz dwa pliki:

- `raport_oczekiwany.xlsx`
- `raport_wygenerowany.xlsx`

Chcesz sprawdzić, czy są takie same. Otwierasz oba w Excelu i... **dwadzieścia arkuszy, trzy tysiące komórek, pięć wykresów, dwadzieścia reguł formatowania warunkowego**. Porównywanie tego okiem to droga w gęstej mgle: wiesz, że coś może być nie tak, ale nie wiesz gdzie, a weryfikacja każdej komórki ręcznie zajmie Ci pół dnia.

Co więcej, nawet gdybyś chciał porównać pliki **maszynowo**, masz problem, którego większość ludzi nie przewiduje:

- Pliki `.xlsx` to archiwa ZIP (moduł 02). ZIP zapisuje **znaczniki czasu** dla każdego wpisu.
- openpyxl **nadpisuje czas modyfikacji skoroszytu bieżącym czasem przy każdym `save()`** — w `docProps/core.xml`. Całkowicie niezależnie od tego, co ustawisz w `wb.properties.modified`.
- Do tego kompresja DEFLATE może w różnych wersjach `zlib` wyprodukować **inne bajty** z tych samych danych.

Wniosek jest twardy i trzeba go przyswoić raz na zawsze:

> **Dwa uruchomienia tego samego, doskonałego kodu NIGDY nie dadzą identycznych bajtów.**

To jest dokładnie ten moment, w którym początkujący mówi „no to się nie da testować" i odpuszcza. Ale to jest **zły wniosek z prawdziwej przesłanki**. Prawdziwy wniosek brzmi:

> **Nie porównuj plików. Porównuj to, co one ZNACZĄ.**

I to jest cały ten moduł.

### 1.3. Odcisk palca zamiast zdjęcia

Analogia, która rozwiąże cały problem: **odcisk palca**.

Nie da się porównać dwóch ludzi, patrząc na nich — są zbyt złożeni: miliardy cech, z których każda może się różnić. Ale można porównać ich **odciski palców**: krótkie, stabilne, unikalne podsumowanie tego, kim są.

Dokładnie tak traktuj raport. Zamiast porównywać plik (kompletnie nieporównywalny, z metadanymi i kompresją), zbuduj jego **odcisk** — krótkie, tekstowe podsumowanie, które mówi:

```text
Arkusze: Dane, Podsumowanie, Metodyka
Dane         : 2 431 wierszy × 8 kolumn, 18 442 komórki z danymi, 0 reguł CF
Podsumowanie : 14 wierszy × 4 kolumny, 56 komórek z danymi, 3 reguły CF
Metodyka     : 9 wierszy × 2 kolumny, 18 komórek z danymi, 0 reguł CF
Tabele       : SprzedazTabela (A1:H2432)
Obrazy       : 1 (xl/media/image1.png)
Reguły CF    : 3 (colorScale, cellIs, expression)
Fonty w pliku: 4
```

Ten odcisk ma dwie cudowne właściwości:

1. **Jest stabilny** — nie zmienia się między uruchomieniami, bo nie zawiera metadanych ani kompresji.
2. **Jest porównywalny przez `diff`** — a gdy się zmieni, `diff` **powie Ci dokładnie, która linia i która liczba**, nie „pliki się różnią".

I jeszcze trzecia, najważniejsza:

3. **Zapisujesz go raz do pliku tekstowego i staje się on punktem odniesienia na zawsze.** To się nazywa **golden file** („złoty plik"). Uruchamiasz test za pół roku — jeśli odcisk jest inny niż zapisany, test jest czerwony i pokazuje różnicę. Jeśli nie — nikt nie ruszył raportu.

To jest twój **hamulec**.

### 1.4. `BytesIO` — przymierzalnia

Druga analogia, bez której ten moduł nie ma sensu.

Wyobraź sobie sklep z ubraniami. Możesz **wyjść z przymierzalni w ubraniu i od razu je kupić**, płacąc kartą, stojąc w kolejce, czekając na paragon. Albo możesz **stanąć przed lustrem w przymierzalni**, obejrzeć się z każdej strony, odwrócić się, sprawdzić kieszenie — i dopiero wtedy zdecydować. W drugim przypadku nic nie kupiłeś, nic nie zapłaciłeś, nie stałeś w kolejce, a mimo to **wiesz dokładnie, jak wyglądasz**.

`BytesIO` to **lustro w przymierzalni** dla Twojego skoroszytu.

Zamiast:

```python
wb.save("output/test.xlsx")          # wyszedłeś z przymierzalni do kasy
inny = load_workbook("output/test.xlsx")   # i musisz wrócić do sklepu, żeby obejrzeć
assert inny["Dane"]["A1"].value == "Produkt"
# ... i jeszcze posprzątaj plik testowy!
```

robisz:

```python
bufor = io.BytesIO()
wb.save(bufor)                        # zapis do PAMIĘCI, na dysk nic nie trafia
inny = load_workbook(io.BytesIO(bufor.getvalue()))
assert inny["Dane"]["A1"].value == "Produkt"
# koniec. Nic nie zostało na dysku.
```

**Co się dzieje w pamięci.** `BytesIO` to **plik, który istnieje wyłącznie w RAM** — bufor bajtów z interfejsem pliku (`read`, `write`, `seek`). openpyxl nie robi różnicy między tym buforem a prawdziwym plikiem, bo `wb.save()` przyjmuje **obiekt plikopodobny** (ang. *file-like object*), czyli wszystko, co umie `write`. `BytesIO` to umie.

**Co trafi do pliku.** Nic. Dosłownie nic. Żaden bajt nie wyląduje na dysku. To jest ta różnica, która sprawia, że testy raportów mogą trwać **milisekundy** i nie brudzą katalogu `output/`.

I konsekwencja, którą docenisz dopiero w praktyce: **test w `BytesIO` nie może „nie posprzątać"**. Nie ma pliku, który zostanie po nieudanym teście. Nie ma katalogu `tmp/`, który rośnie. Nie ma sytuacji „u mnie działa, bo mam stary plik testowy". Test jest albo zielony, albo czerwony — i nie zostawia śmieci.

### 1.5. Fixtures to `mise en place`, a CI to druga para oczu

Dwie analogie na koniec, obie konkretne.

**Fixture** (wymawiane „fiksczer") to w `pytest` stan przygotowany przed testem. Najlepsza analogia nie jest informatyczna — jest kuchenna. W profesjonalnej restauracji kucharz przed rozpoczęciem serwisu robi **`mise en place`**: warzywa pokrojone, sosy gotowe, patelnie rozgrzane, przyprawy odmierzone w miseczkach. Gdy wchodzi zamówienie, nikt nie kroi cebuli od zera — **sięga po gotowe**. Fixture to dokładnie to: stanowisko przygotowane **raz**, używane przez **wiele testów**, posprzątane **po** nich. Bez tego każdy test zaczynałby od krojenia cebuli.

**CI** (Continuous Integration) to **druga para oczu, która nigdy nie jest zmęczona**. Ty uruchamiasz testy u siebie, na swoim Pythonie 3.12, z zainstalowanym `lxml`, na Windowsie, w polskim locale. CI uruchamia je na świeżym kontenerze, w kilku wersjach Pythona, bez `lxml`, w angielskim locale, o trzeciej w nocy, na cudzym commicie. I właśnie dlatego **CI znajduje błędy, których Ty nie znajdziesz nigdy** — bo szuka ich w warunkach, których nie masz na co dzień.

I jeszcze jedno, kluczowe dla tego kursu: **CI nie potrzebuje Excela**. Ani jednego kroku instalacji Office, ani jednego wywołania COM, ani jednego okna. To nie jest ciekawostka — to jest dowód, że openpyxl naprawdę nie uruchamia Excela (moduł 01). Twój raport da się zbudować i sprawdzić na maszynie bez systemu okien. Zapamiętaj to zdanie — wrócimy do niego w 2.12.

### 1.6. Trzy kategorie prawdy dla tego modułu

Zgodnie z konwencją kursu — tabela, wokół której zbudowany jest moduł:

| Kategoria | W kontekście testowania |
|---|---|
| **openpyxl potrafi** | `wb.save(BytesIO())`, `load_workbook(BytesIO)`, odczyt wartości, typów, `number_format`, tabel (`ws.tables`), reguł CF (`.rules`), walidacji, obszaru druku, scalenia, ustawień strony — wszystko nadaje się na asercję |
| **openpyxl potrafi częściowo** | porównywanie plików: **nie bajt w bajt**, ale po odcisku/hashu części/modelu; `ws.print_area` zwraca **kwalifikowaną, absolutną** postać, nie to, co wpisałeś; `len(ws.conditional_formatting)` liczy **zakresy, nie reguły**; style w `read_only` są niedostępne |
| **openpyxl gubi / nie ma** | deterministycznych bajtów (`docProps/core.xml` + znaczniki ZIP), możliwości przypięcia `properties.modified`, oficjalnego narzędzia testowego (piszesz własne), odczytu stylów w `read_only`, controlli nad utraconymi sparkline'ami (możesz tylko udokumentować, że giną) |

## 2. Teoria

### 2.1. Trzy poziomy pewności — co właściwie chcesz wiedzieć

Zanim napiszesz pierwszy test, musisz odpowiedzieć na pytanie: **co ma być sprawdzone?** Bo „testować raport" to nie jedna rzecz, a trzy różne, na trzech różnych poziomach:

| Poziom | Pytanie | Przykład asercji | Koszt |
|---|---|---|---|
| **1. Dane** | Czy liczby są tam, gdzie mają być? | `ws["D2"].value == 1200.5` | bardzo tani |
| **2. Typy i formaty** | Czy to są liczby, a nie tekst? Czy data to data? | `isinstance(v, float)`, `cell.number_format == "#,##0.00"` | tani |
| **3. Struktura i wygląd** | Czy są tabele, reguły, filtry, zamrożone okna, obszar druku? | `ws.tables["T1"].ref == "A1:D100"` | tani, jeśli mierzysz liczby |
| **4. Cały dokument** | Czy raport jako całość nie zmienił się niepostrzeżenie? | odcisk skoroszytu vs `golden` | tani, jeśli zautomatyzowany |

Zauważ, co **nie** jest poziomem: „czy wygląda ładnie". Tego nie da się automatycznie sprawdzić i **nie powinieneś próbować** — to zadanie dla oka (moduł 18 nazywał to „weryfikacją obowiązkową"). Test automatyzuje **liczby i strukturę**, a nie gust.

Praktyczna zasada: **poziom 1 i 2 to Twój codzienny test. Poziom 3 to Twoja sieć bezpieczeństwa. Poziom 4 to Twoja polisa ubezpieczeniowa na przyszłość.**

### 2.2. BytesIO — anatomia i szczegóły, które gryzą

Wiemy już, że `wb.save()` przyjmuje obiekt plikopodobny. Ale diabeł siedzi w szczegółach, a trzy z nich trzeba znać **zanim** się na nie natkniesz.

**Szczegół 1 — `getvalue()` vs `read()`.** `BytesIO` ma wewnętrzny „kursor" pozycji, tak jak plik.

```python
import io
from openpyxl import Workbook

bufor = io.BytesIO()
wb = Workbook()
wb.active["A1"] = "x"
wb.save(bufor)

# Kursor stoi na KONCU bufora (bo tam skonczyl zapis):

print(len(bufor.getvalue()))   # 4329  <- getvalue() zwraca CALY bufor, ignoruje kursor
print(bufor.read())            # b''   <- read() czyta OD KURSORA, wiec nic!

bufor.seek(0)                  # przewin na poczatek
print(len(bufor.read()))       # 4329  <- teraz dziala
```

**Wniosek:** `getvalue()` jest „odporne" na pozycję kursora i dlatego jest wygodniejsze w testach. `read()` (i przekazanie bufora do `send_file`, `StreamingResponse`) **wymaga `seek(0)`**. To rozróżnienie jest źródłem najbardziej frustrującego błędu w tym module: **plik „ma poprawną długość", ale wysyła się jako pusty**. Dlaczego? Bo nikt nie przewinął bufora. Zapamiętaj: *`getvalue()` w testach, `seek(0)` przy strumieniowaniu*.

**Szczegół 2 — `save()` domyka archiwum, ale nic nie „przewija".** `wb.save(bufor)` buduje ZIP **do końca** — a to znaczy, że zapisuje „katalog centralny" ZIP-a (spis treści archiwum), który w formacie ZIP leży **na końcu pliku** (moduł 02). Dopóki `save()` się nie skończy, plik jest **niekompletny** i nie da się go otworzyć. Dlatego:

- **`wb.save(bufor)` jest wystarczające** — openpyxl sam zamyka archiwum.
- **Nie próbuj „podglądać" bufora w trakcie zapisu.** Nie ma czego podglądać.
- **Jeśli budujesz ZIP ręcznie** (`ZipFile` + `ExcelWriter`, np. żeby wstawić plik do archiwum w locie), **musisz sam zamknąć `ZipFile`** przed odczytem. Inaczej `load_workbook` powie Ci `BadZipFile` — i to jest ten „niewidzialny" błąd, o którym mówi się „bufor obcięty".

**Szczegół 3 — `with io.BytesIO() as buf` bywa pułapką.** Kontekst menedżer (`with`) **zamyka** bufor przy wyjściu z bloku. A zamknięty `BytesIO` nie oddaje danych:

```python
# ❌ ZLE
with io.BytesIO() as bufor:
    wb.save(bufor)
bajty = bufor.getvalue()      # ValueError: I/O operation on closed file

# ✅ DOBRZE
bufor = io.BytesIO()
wb.save(bufor)
bajty = bufor.getvalue()      # dziala
```

`BytesIO` nie wymaga zamykania i w testach **nie ma powodu** używać `with`. Zamknięcie ma sens tylko przy ogromnych buforach, gdy chcesz **natychmiast** zwolnić pamięć — i wtedy `getvalue()` robisz **wewnątrz** bloku.

**Szczegół 4 — `load_workbook(BytesIO(...))` też wymaga `close()`.** To powtórka z modułu 19, ale tu wraca, bo testy uruchamiasz tysiące razy:

```python
wb = load_workbook(io.BytesIO(bajty))
try:
    ...
finally:
    wb.close()
```

**Co się dzieje w pamięci.** `load_workbook` na buforze ma **podwójny koszt**: bufor z bajtami ZIP-a (nieskompresowane dane? nie — skompresowane, więc mały) **plus** pełny model skoroszytu w RAM (moduł 19: ~50× rozmiar **rozpakowanej** treści). Dla pliku testowego 50 wierszy to nic. Ale gdy testujesz plik 5 000 wierszy — to już setki megabajtów na jeden test, i uruchomienie 40 takich testów z rzędu może być odczuwalne.

**Praktyczna konsekwencja:** **testy generują małe skoroszyty.** Nie testujesz „czy wygeneruje się raport 500 000 wierszy" — testujesz **strukturę** na 50 wierszach, a wydajność osobno (2.10). To jest zasada 3 z piramidy testów: setki małych, szybkich testów i garść dużych, rzadkich.

### 2.3. Co warto testować — katalog asercji

To jest serce modułu. Poniższa tabela to gotowy „checklist" asercji. Każdy wiersz to konkretne wywołanie, które możesz skopiować.

| Co sprawdzasz | Jak to wygląda w kodzie | Uwaga |
|---|---|---|
| Obecność i nazwy arkuszy | `wb.sheetnames == ["Dane", "Podsumowanie"]` | lista zachowuje kolejność |
| Kolejność arkuszy | `wb.sheetnames[0] == "Podsumowanie"` | to jest część UX raportu |
| Aktywny arkusz | `wb.active.title` | |
| Wartość w krytycznej komórce | `ws["B2"].value == 1200.5` | |
| Kwota jako float | `isinstance(ws["B2"].value, float)` | **test przeciwko źródłu prawdy** |
| Data jako `datetime` | `isinstance(ws["A2"].value, datetime)` | nie jako tekst! |
| Format liczby | `ws["B2"].number_format == "#,##0.00"` | testuj **kod**, nie wygląd |
| Brak `None` tam, gdzie ma być wartość | `assert ws["C2"].value is not None` | |
| Formatowanie warunkowe — zakres | `str(cf.sqref) == "D2:D100"` | `sqref` to `MultiCellRange` |
| Formatowanie warunkowe — liczba reguł | `sum(len(cf.rules) for cf in ws.conditional_formatting) == 3` | **nie `len()`!** patrz 2.3.1 |
| Typ reguły | `rule.type == "cellIs"` | `"expression"` dla `FormulaRule` |
| Priorytet reguły | `rule.priority == 1` | mniejsza liczba = ważniejsza |
| Walidacja danych | `dv.type == "list"`, `"B2:B100" in str(dv.sqref)` | |
| Tabela — nazwa i zakres | `ws.tables["Sprzedaz"].ref == "A1:H100"` | |
| Tabela — styl | `ws.tables["Sprzedaz"].tableStyleInfo.name == "TableStyleMedium9"` | |
| Scalony nagłówek | `"A1:D1" in {str(r) for r in ws.merged_cells.ranges}` | `ranges` to zbiór obiektów |
| Zamrożone okna | `str(ws.freeze_panes) == "B2"` | `str()` dla bezpieczeństwa |
| Filtr automatyczny | `ws.auto_filter.ref == "A1:H100"` | |
| Orientacja strony | `ws.page_setup.orientation == "landscape"` | |
| Obszar druku | `"$A$1:$F$50" in ws.print_area` | **patrz 2.3.2!** |
| Powtarzany nagłówek druku | `ws.print_title_rows == "1:1"` | |
| Dopasowanie do stron | `ws.sheet_properties.pageSetUpPr.fitToPage is True` | + `fitToWidth`, `fitToHeight` |
| Szerokość kolumny | `ws.column_dimensions["A"].width == 25` | |
| Właściwości dokumentu | `wb.properties.creator == "Automat"` | ślad audytowy |
| Brak formuł tam, gdzie mają być wartości | `[c for r in ws.iter_rows() for c in r if c.data_type == "f"] == []` | |
| Liczba obrazów | `len(ws._images)` — **albo** policz `xl/media/*` w ZIP | `_images` jest prywatne! |
| Liczba wykresów | `len(ws._charts)` — **albo** policz `xl/charts/*` | jak wyżej |

#### 2.3.1. Trzy pułapki pomiarowe, które psują asercje

To są trzy miejsca, gdzie **intuicja mówi jedno, a API drugie**. Każde z nich sam kiedyś zaliczysz — lepiej teraz.

**Pułapka 1: `len(ws.conditional_formatting)` to NIE liczba reguł.** Spójrzmy w źródło `ConditionalFormattingList`:

```python
def __len__(self):
    return len(self._cf_rules)
```

`_cf_rules` to słownik, którego **kluczem jest zakres**, a **wartością lista reguł**. Czyli `len()` zwraca **liczbę zakresów**, nie reguł. Jeśli masz jedną regułę na `B2:B100` i trzy reguły na `D2:D100`, `len(ws.conditional_formatting)` zwróci **2**. Prawidłowa liczba reguł to:

```python
liczba_regul = sum(len(cf.rules) for cf in ws.conditional_formatting)
```

**Pułapka 2: `ws.print_area` nie zwraca tego, co wpisałeś.** Getter robi:

```python
@property
def print_area(self):
    self._print_area.title = self.title
    return str(self._print_area)
```

a `PrintArea.__str__` produkuje **pełną, kwalifikowaną, absolutną** postać:

```text
Dane!$A$1:$F$10     # albo 'Dane 2026'!$A$1:$F$10 gdy nazwa wymaga cudzysłowu
```

Więc `assert ws.print_area == "A1:F10"` **zawsze będzie czerwone**. To nie błąd openpyxl — to konsekwencja tego, że obszar druku przechodzi przez ten sam mechanizm co `quote_sheetname` i `absolute_coordinate` (moduł 07). Testuj tak:

```python
assert "$A$1:$F$10" in (ws.print_area or "")
```

**Pułapka 3: `ws.freeze_panes`, `ws.print_title_rows`, `ws.print_area` bywają `None` lub `""`.** Gdy nie są ustawione, nie dostajesz wyjątku — dostajesz `None` (albo pusty string). Test `assert ws.print_title_rows == "1:1"` na arkuszu bez tytułów druku da czytelne `AssertionError: None != '1:1'`. To dobrze — ale tylko jeśli wiesz, że `None` jest **prawidłową** odpowiedzią, a nie usterką. W testach pomocnicze normalizacje (`str(x) if x else None`) ratują przed fałszywymi alarmami.

#### 2.3.2. Testowanie „nie ma `None` tam, gdzie ma być wartość"

To jedna z najcenniejszych asercji w całym module — i jedna z tych, których ludzie nie piszą, bo wydaje się zbyt prosta.

Pamiętasz z modułu 05: **pusta komórka to nie zero**. A teraz wyobraź sobie, że raport spadł na kolumnę, która miała zawierać kwoty, ale reguła agregująca zwróciła `None` dla 300 rekordów. Raport **wnie otworzy się bez ostrzeżenia**. Kwoty będą po prostu puste. Excel nic nie powie. Użytkownik pomyśli, że „nie było sprzedaży". Test tego nie przepuści:

```python
def test_zadna_kwota_nie_jest_pusta(raport_bytes):
    wb = load_workbook(io.BytesIO(raport_bytes))
    try:
        ws = wb["Sprzedaz"]
        braki = [
            c.coordinate
            for wiersz in ws.iter_rows(min_row=2)
            for c in wiersz
            if c.column == 4 and c.value is None
        ]
        assert not braki, f"Brak kwoty w komórkach: {braki[:10]}"
    finally:
        wb.close()
```

Zwróć uwagę na komunikat: **`braki[:10]`** — żeby przy 300 pustych komórkach test powiedział Ci *gdzie*, a nie wyrzucił listę na ekran. To drobiazg, ale w praktyce decyduje, czy test pomaga, czy przeszkadza.

### 2.4. Fixtures — `mise en place` dla testów

`pytest` daje Ci dwa gotowe fixture'y, których będziesz używał w każdym teście plików:

- **`tmp_path`** — świeży, **unikalny** katalog tymczasowy dla **każdego** wywołania testu, w którym możesz pisać pliki; `pytest` sam go posprząta. Używaj go **tylko** tam, gdzie kod *wymaga* ścieżki.
- **`capsys`, `monkeypatch`, `caplog`** — do wychwytywania wyjścia, podmieniania rzeczy i czytania logów. Przydatne, gdy chcesz przetestować, że kod **ostrzegł** przed utratą danych (moduł 18).

Ale prawdziwa siła to **własne fixture'y**.

#### 2.4.1. „Fabryka jako fixture" — wzorzec, który zmieni Twoje testy

Naiwne podejście to fixture zwracający **gotowy** skoroszyt:

```python
@pytest.fixture
def prosty_skoroszyt():
    wb = Workbook()
    wb.active["A1"] = "gotowe"
    return wb
```

Problem: każdy test dostaje **ten sam** skoroszyt i **nie może go zmienić**. Test na pusty arkusz i test na arkusz z 500 wierszami nie mają jak powstać.

Rozwiązanie to **fixture zwracające funkcję** („factory as fixture"): fixture dostarcza **narzędzie do budowania**, a każdy test buduje dokładnie to, czego potrzebuje. W kuchni: nie dostajesz gotowej potrawy — dostajesz **stanowisko `mise en place`**, z którego składasz to, co akurat zamówiłeś.

```python
# conftest.py
import io
import pytest
from openpyxl import Workbook


@pytest.fixture
def zbuduj_skoroszyt():
    """Fabryka: zwraca funkcje, ktora buduje skoroszyt w pamieci
    i oddaje jego BAJTY (nie obiekt - bo bajty sa niezmienne!)."""
    def build(rows=None, sheet="Dane", ustaw=None):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        for wiersz in (rows or []):
            ws.append(wiersz)
        if ustaw:
            ustaw(wb, ws)          # hak na modyfikacje (styl, tabela, reguly...)
        bufor = io.BytesIO()
        wb.save(bufor)
        return bufor.getvalue()    # bajty: kazdy test dostaje SWOJA kopie
    return build
```

Dlaczego **bajty, a nie obiekt `Workbook`**? Z trzech powodów, wszystkie praktyczne:

1. **Niezmienność.** Bajty są niemutowalne. Test nie może „przypadkiem" dopisać czegoś do wspólnego obiektu i zepsuć następnego testu. Obiekt `Workbook` byłby **współdzielonym stanem** — a to najczęstsze źródło „testów, które przechodzą tylko w kolejności alfabetycznej".
2. **Wierność.** Bajty to **dokładnie to**, co kod produkcyjny dostanie ze S3, z uploadu albo z HTTP. Testujesz prawdziwą ścieżkę, nie uproszczoną.
3. **Kontekst użycia.** Jeśli testowany kod przyjmuje **ścieżkę albo obiekt plikopodobny** (a powinien!), to bajty w `BytesIO` są jego naturalnym wejściem.

Hak `ustaw` jest tu kluczowy. Bez niego fabryka byłaby sztywna. Z nim — każdy test dokłada to, czego potrzebuje:

```python
def test_z_naglowkiem_w_tabeli(zbuduj_skoroszyt):
    from openpyxl.worksheet.table import Table, TableStyleInfo

    def dodaj_tabele(wb, ws):
        tab = Table(displayName="Sprzedaz", ref="A1:C4")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws.add_table(tab)

    bajty = zbuduj_skoroszyt(
        rows=[["Data", "Region", "Kwota"], ["2026-01-01", "Północ", 100.0]],
        ustaw=dodaj_tabele,
    )
    ...  # teraz test sprawdza tabele
```

#### 2.4.2. Jedna zasada, którą złamiesz i tego pożałujesz

> **Fixture nie może zawierać asercji.**

Fixture **przygotowuje**. Test **sprawdza**. Jeśli fixture ma `assert`, to błąd w niej wygląda jak błąd w kodzie produkcyjnym — i tracisz pół godziny na szukanie buga tam, gdzie go nie ma. Rozdzielenie „przygotowania" od „sprawdzenia" to dokładnie ta sama zasada, którą w 2.11 przeniesiemy na poziom całego projektu.

### 2.5. Odcisk skoroszytu — `podsumowanie_skoroszytu()`

Teraz narzędzie numer jeden tego modułu. Zbudujemy funkcję, która bierze ścieżkę do `.xlsx` i zwraca **słownik** z odciskiem — a potem drugą, która zamienia ten słownik w **czytelny tekst do porównania**.

Odcisk ma trzy warstwy, każda z innego miejsca:

**Warstwa 1 — inwentarz części ZIP (struktura pliku).** To jest bezpośrednie zastosowanie modułu 02. Zamiast sięgać do prywatnych atrybutów (`ws._images`), **liczymy części archiwum** — a to jest publiczne, stabilne i wierne temu, co naprawdę jest w pliku:

```python
with zipfile.ZipFile(sciezka) as zf:
    nazwy = zf.namelist()
    media = [n for n in nazwy if n.startswith("xl/media/")]
    wykresy = [n for n in nazwy if n.startswith("xl/charts/chart")]
    tabele = [n for n in nazwy if n.startswith("xl/tables/")]
```

**Warstwa 2 — model skoroszytu (treść).** Wczytujemy plik i liczymy to, co opisuje **znaczenie** dokumentu: nazwy arkuszy, wymiary, liczbę komórek z danymi, liczbę i typy reguł, tabele, scalenia, ustawienia druku.

**Warstwa 3 — stylistyka (z `xl/styles.xml`).** Liczymy, ile **definicji** stylów jest w pliku: fontów, wypełnień, obramowań, kodów formatów. Robimy to **parsując XML**, a nie przez prywatne `wb._fonts` — i to jest trzecie zastosowanie modułu 02 w tym module:

```python
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
root = ET.fromstring(zf.read("xl/styles.xml"))
fonty = len(root.findall(f"{NS}font"))
```

**Dlaczego liczba definicji stylów jest dobrym składnikiem odcisku?** Bo Excel **deduplikuje** style: dwa identyczne `Font(bold=True, size=11, name="Calibri")` stają się **jednym** wpisem. Jeśli Twoja zmiana w kodzie sprawiła, że zamiast 4 fontów jest 40 — to znak, że gdzieś tworzysz obiekt stylu w pętli (pułapka wydajnościowa z modułu 09!). Odcisk to złapie jako zmianę liczby, a nie jako „coś się różni".

#### 2.5.1. Co **wykluczamy** z odcisku i dlaczego

To jest miejsce, w którym odcisk albo staje się użyteczny, albo bezużyteczny. Musisz **świadomie wykluczyć części zmienne**:

| Część / właściwość | Dlaczego wykluczamy |
|---|---|
| `docProps/core.xml` | openpyxl wpisuje tam `dcterms:modified` = czas bieżący **przy każdym `save()`**. Nie da się tego przypiąć. |
| `docProps/app.xml` | zawiera m.in. wersję aplikacji; zmienia się między wersjami openpyxl |
| znaczniki czasu wpisów ZIP | ustawiane na „teraz" przy każdym zapisie |
| `xl/calcChain.xml` | kolejność przeliczania formuł; bywa odbudowywana |

**Wszystko inne zostaje.** I to jest całe `podsumowanie_skoroszytu` — **lista tego, co trwałe**, nie „cały plik".

### 2.6. Cztery strategie porównywania plików

Gdy chcesz porównać dwa skoroszyty, masz dokładnie cztery drogi. Warto znać wszystkie, bo każda nadaje się do czegoś innego.

#### Strategia 1 — bajt w bajt (nie używać)

```python
# ❌ TO NIGDY NIE ZADZIALA
assert Path("a.xlsx").read_bytes() == Path("b.xlsx").read_bytes()
```

**Dlaczego nie działa** — trzy niezależne powody:

1. `docProps/core.xml` ma `dcterms:modified` = czas zapisu (zmienia się przy każdym `save()`).
2. Każdy wpis ZIP ma znacznik czasu. W formacie ZIP to pole ma rozdzielczość **2 sekund** — więc nawet dwa zapisy w tej samej sekundzie mogą wypaść różnie.
3. Kompresja DEFLATE może dać inne bajty na różnych wersjach `zlib` (laptop vs serwer). Te same dane wejściowe, **inny plik**.

Wniosek: **kto pisze test bajt w bajt, ten pisze test, który zawsze jest czerwony.** To nie jest „trudny test" — to **nie-test**.

#### Strategia 2 — hash **rozpakowanych** części (zalecana do CI)

Zamiast hashować plik, hashowało się **treść jego części**, w kolejności nazw, **bez kompresji** i **bez części zmiennych**:

```python
import hashlib
import zipfile

CZESCI_ZMIENNE = {"docProps/core.xml", "docProps/app.xml", "xl/calcChain.xml"}

def odcisk_bajtowy(sciezka) -> str:
    """Hash tresci czesci - odporny na kompresje i znaczniki ZIP."""
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in sorted(zf.namelist()):
            if nazwa in CZESCI_ZMIENNE:
                continue
            h.update(nazwa.encode("utf-8"))
            h.update(b"\x00")
            h.update(zf.read(nazwa))
            h.update(b"\x00")
    return h.hexdigest()
```

**Co to daje:** ten sam kod na dwóch maszynach, z różnymi wersjami `zlib`, da **ten sam hash**. A jednocześnie **każda realna zmiana** w pliku (inna wartość, inny format, inna reguła) zmieni hash. To jest strategia do **CI i porównań między środowiskami**.

**Czego nie daje:** **nie powie Ci, co się zmieniło.** Hash to liczba — „zmieniło się" albo „nie". Dlatego hash jest dobry jako **sygnał**, a odcisk tekstowy (strategia 4) jako **diagnoza**. W praktyce używasz obu: hash w CI jako szybką bramkę, odcisk w testach jako źródło informacji.

**Pułapka:** `xl/styles.xml` może mieć inne **uporządkowanie atrybutów** między wersjami openpyxl. Znormalizuj XML, jeśli porównujesz między wersjami biblioteki:

```python
from xml.etree import ElementTree as ET

def znormalizuj(xml_bajty: bytes) -> bytes:
    """C14N - kanoniczna postac XML (Python 3.8+)."""
    return ET.canonicalize(xml_data=xml_bajty).encode("utf-8") + b"\n"
```

#### Strategia 3 — porównanie po modelu (komórka po komórce)

Wczytaj **oba** pliki i porównaj ich zawartość w Pythonie:

```python
def porownaj_modele(sciezka_a, sciezka_b) -> list[str]:
    """Zwraca liste ROZNIC (puste = identyczne modele)."""
    roznice = []
    wa = load_workbook(sciezka_a)
    wb_ = load_workbook(sciezka_b)
    try:
        if wa.sheetnames != wb_.sheetnames:
            roznice.append(f"arkusze: {wa.sheetnames} != {wb_.sheetnames}")
            return roznice
        for nazwa in wa.sheetnames:
            sa, sb = wa[nazwa], wb_[nazwa]
            if (sa.max_row, sa.max_column) != (sb.max_row, sb.max_column):
                roznice.append(f"[{nazwa}] wymiary {sa.max_row}x{sa.max_column} "
                               f"!= {sb.max_row}x{sb.max_column}")
            for wiersz_a, wiersz_b in zip(sa.iter_rows(), sb.iter_rows()):
                for ca, cb in zip(wiersz_a, wiersz_b):
                    if ca.value != cb.value:
                        roznice.append(f"[{nazwa}] {ca.coordinate}: "
                                       f"{ca.value!r} != {cb.value!r}")
    finally:
        wa.close()
        wb_.close()
    return roznice
```

**Zaleta:** dostajesz **konkretne różnice z adresami komórek**. **Wada:** pamięć ×2 (dwa modele naraz — moduł 19!), a przy dużych plikach to boli. Stosuj dla małych plików i **jednorazowej diagnostyki**, nie w pętli CI.

#### Strategia 4 — snapshot tekstowy (golden file) — **ta, której chcesz używać**

Zamieniasz skoroszyt w **tekst** (jedna linia = jedna niepusta komórka + jej format) i porównujesz z **zapisanym w repozytorium** plikiem wzorcowym. Różnica pokazuje się jako `diff` z nazwami komórek.

```text
Dane|A1|wartosc='Produkt'|typ=s|fmt=General|b=1|fill=solid:001F4E78|align=center|br=-
Dane|B1|wartosc='Kwota'|typ=s|fmt=General|b=1|fill=solid:001F4E78|align=center|br=-
Dane|A2|wartosc='Kawa'|typ=s|fmt=General|b=0|fill=-|align=-|br=-
Dane|B2|wartosc=19.99|typ=n|fmt=#,##0.00|b=0|fill=-|align=-|br=-
```

**Trzy zalety, które czynią to zwycięzcą:**

1. **Czytelne w `git diff`.** Zmieniłeś format kwoty z `#,##0.00` na `#,##0`? `diff` pokaże **dokładnie tę linię**, z adresem `Dane|B2`.
2. **Nadaje się do przeglądu.** Ten sam `diff`, który widzi test, widzi recenzent w Pull Requeście. Zmiana snapshotu **razem z kodem** to **dokumentacja zamierzonej zmiany wyglądu raportu**.
3. **Nie zależy od kompresji ani metadanych.** Tekst jest tekstem.

**Wada:** snapshot trzeba **utrzymywać**. Gdy intencjonalnie zmieniasz raport, snapshot trzeba **zaktualizować** — i tu jest miejsce na przełącznik `--update-golden` (2.6.1).

#### 2.6.1. Wzorzec `--update-golden`

Klasyczny błąd przy golden testach: „test się wywalił, więc nadpisałem wzorzec". To zamienia test w pieczątkę — zawsze zielony, nic nie sprawdza. Lekarstwo: **aktualizacja wzorca musi być osobną, świadomą decyzją**, a nie automatycznym odruchem.

```python
# conftest.py
import difflib
from pathlib import Path
import pytest

KATALOG_WZORCOW = Path(__file__).parent / "golden"


def pytest_addoption(parser):
    parser.addoption(
        "--update-golden",
        action="store_true",
        default=False,
        help="Nadpisz pliki wzorcowe (uzyj swiadomie i przejrzyj git diff!).",
    )


@pytest.fixture
def golden(request):
    """Fixture porownujacy tekst ze wzorcem w tests/golden/<nazwa>.txt."""
    def porownaj(nazwa: str, tekst: str) -> None:
        KATALOG_WZORCOW.mkdir(parents=True, exist_ok=True)
        plik = KATALOG_WZORCOW / f"{nazwa}.txt"

        if request.config.getoption("--update-golden"):
            plik.write_text(tekst, encoding="utf-8")
            return

        if not plik.exists():
            plik.write_text(tekst, encoding="utf-8")
            pytest.fail(
                f"Brak wzorca - utworzono {plik}.\n"
                "Przejrzyj go OKIEM, potem zatwierdź: git add " + str(plik)
            )

        oczekiwany = plik.read_text(encoding="utf-8")
        if tekst == oczekiwany:
            return

        roznica = "".join(
            difflib.unified_diff(
                oczekiwany.splitlines(keepends=True),
                tekst.splitlines(keepends=True),
                fromfile=f"{nazwa}.txt (wzorzec)",
                tofile=f"{nazwa}.txt (aktualny)",
            )
        )
        pytest.fail(f"Raport zmienił się względem wzorca:\n\n{roznica}")

    return porownaj
```

Zwróć uwagę na **trzy decyzje projektowe**:

1. **Brak wzorca = `pytest.fail`, nie `pass`.** Nowy snapshot **musi zostać obejrzany**, zanim zostanie zaufany. To zabezpieczenie przed sytuacją „dodałem test, test przeszedł, wszystko super" — a snapshot zawiera błędy, bo kod był błędny.
2. **`--update-golden` nie porównuje, tylko pisze.** Nie ma trybu „porównaj i napraw". Świadoma zmiana.
3. **Komunikat mówi, co zrobić** (`git add ...`). Test ma być **instrukcją**, nie zagadką.

### 2.7. Testy właściwościowe — gdy nie znasz wszystkich przypadków

Do tej pory pisaliśmy testy **przykładowe**: „dam wejście X, oczekuję wyniku Y". To działa, gdy znasz przypadki. Ale co z funkcją, która ma „jakoś" oczyścić dane z Excela — gdzie wejść jest nieskończenie wiele?

Analogia: **kontrola jakości w fabryce**. Nie sprawdzasz **każdej** śruby — to niemożliwe. Sprawdzasz **losowo wybraną próbkę**, ale z **narzędziem pomiarowym**, które wie, jak wygląda wadliwa śruba. I co więcej: dobra kontrola jakości **wybiera najtrudniejsze egzemplarze** — te o skrajnych wymiarach, nie te „typowo dobre".

`hypothesis` robi dokładnie to. Zamiast podawać przykłady, opisujesz **niezmiennik** (właściwość), która ma być zawsze prawdziwa, i strategię generowania danych. `hypothesis` **sama szuka kontrprzykładów**, a gdy je znajdzie — **„kurczy" je** (ang. *shrink*) do najmniejszego, jaki potrafi.

```python
from hypothesis import given, strategies as st

@given(st.one_of(
    st.none(),
    st.integers(),
    st.floats(allow_nan=False, allow_infinity=False),
    st.text(max_size=20),
))
def test_czyszczenie_kwoty_oddaje_float_albo_none(wartosc):
    """NIEZMIENNIK: niezależnie od wejścia wynik jest float albo None.

    'Żadna wartość nie jest tekstem, jeśli miała być liczbą.'
    """
    wynik = czysc_kwote(wartosc)
    assert wynik is None or isinstance(wynik, float)
```

I druga, mocniejsza właściwość — **idempotencja**:

```python
@given(st.text(alphabet="0123456789., ", max_size=12))
def test_czyszczenie_jest_idempotentne(tekst):
    """NIEZMIENNIK: czyszczenie dwa razy = czyszczenie raz."""
    raz = czysc_kwote(tekst)
    assert czysc_kwote(raz) == raz
```

**Dlaczego te dwie właściwości są tak cenne w kontekście Excela?**

Pierwsza łapie całą klasę błędów, którą w modułach 05 i 06 nazywaliśmy „koszmarem analityka": **tekst udający liczbę**. Jeśli funkcja czyszcząca zwróci `"1 200,50"` jako `str`, nie jako `float`, to dalsze sumowanie w Pythonie wybuchnie — a **w Excelu nie wybuchnie**, tylko da ciche zero. Test właściwościowy nie pozwoli tej klasie błędu przejść.

Druga łapie subtelniejsze: funkcja, która przy pierwszym przebiegu zamienia `"1,5"` na `1.5`, a przy drugim na `1.5` **innym** przebiegiem (np. `"1.5"` → `1` bo kropka jest „separatorem tysięcy"). Idempotencja to najlepsza pojedyncza właściwość dla każdej funkcji czyszczącej.

**Praktyczne uwagi:**

- **`st.floats(allow_nan=False, allow_infinity=False)`** — domyślnie `hypothesis` generuje `NaN` i `inf`, a te psują porównania (`NaN != NaN`!). Wyłącz je świadomie, chyba że **chcesz** je testować.
- **`assume(warunek)`** — gdy odrzucasz część wygenerowanych przypadków. Używaj rzadko; `.filter()` jest wydajniejszy.
- **`@settings(max_examples=200)`** — więcej ziaren, wolniejszy test. Domyślnie 100 wystarczy do złapania większości.
- **`@settings(deadline=None)`** — gdy test jest wolny (np. buduje skoroszyt), `hypothesis` domyślnie krzyczy o przekroczeniu 200 ms na przykład. Wyłączaj świadomie, nie „na wszelki wypadek".
- **Nie używaj `hypothesis` do testowania Excela samego w sobie.** Używaj go do **własnej logiki** — czyszczenia, normalizacji, konwersji. Generowanie 100 skoroszytów na test to marnowanie sekund.

### 2.8. Testy pułapek — kontrakt na ograniczenia

Ten typ testu jest specyficzny i bardzo niedoceniany. **Nie sprawdza, że coś działa. Sprawdza, że coś nadal NIE działa — w znany nam sposób.**

Dlaczego to ma sens? Bo w module 18 ostrzegaliśmy: **openpyxl gubi część danych**. Jeśli nie zapiszesz tego jako testu, to za rok ktoś przyjdzie i powie: „hej, ale my możemy po prostu wczytać ten plik z ludzkim wykresem i zapisać — przecież openpyxl to obsługuje". I wprowadzi błąd, bo **zapomni, że wykres jest odbudowywany z modelu openpyxl i część formatowania ginie**.

Test pułapki wygląda tak:

```python
def test_formuly_zapisywane_przez_openpyxl_nie_maja_wartosci():
    """DOKUMENTUJE OGRANICZENIE (moduł 06):
    openpyxl nie ma silnika obliczeń, więc plik zapisany przez nas
    ma formułę bez cache'owanej wartości. Odczyt z data_only=True
    da None - i to jest OCZEKIWANE, nie bug do naprawy.
    """
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["A2"] = 3
    ws["A3"] = "=SUM(A1:A2)"          # formula, nie wynik

    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()), data_only=True)
    try:
        assert wb2.active["A3"].value is None, (
            "Gdyby tu była wartość, openpyxl zacząłby liczyć formuły - "
            "a to się nie stanie. Ten test chroni przed złudzeniem."
        )
    finally:
        wb2.close()
```

**Docblock wyjaśnia DLACZEGO.** To jest sedno: test pułapki nie jest po to, żeby „sprawdzić" — jest po to, żeby **przekazać wiedzę następnemu programiście** (i przyszłemu Tobie). Bez `"""DOKUMENTUJE OGRANICZENIE..."""` taki test wygląda jak błąd i ktoś go „naprawi".

Trzy testy pułapek, które **musisz mieć**:

1. **Formuła nie ma wartości** (powyżej) — chroni przed oczekiwaniem obliczeń.
2. **`None` to nie zero** — chroni przed `if kwota:` zamiast `if kwota is not None:`.
3. **`=` w danych użytkownika staje się formułą** — chroni przed wstrzyknięciem (moduł 21!). Ten test jest jednocześnie **testem bezpieczeństwa**:

```python
def test_rowna_sie_w_danych_staje_sie_formula_i_trzeba_to_zneutralizowac():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "=1+1"                 # tak wpisze to prosty kod
    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        # ⚠️ To jest FORMULA, nie tekst. Excel ja wykona!
        assert wb2.active["A1"].data_type == "f"
    finally:
        wb2.close()

    # Neutralizacja: wymus typ tekstowy PRZED zapisem.
    # (sprawdź w dokumentacji swojej wersji - to zachowanie jest
    #  na styku modelu i zapisu)
    wb3 = Workbook()
    ws3 = wb3.active
    ws3["A1"] = "=1+1"
    ws3["A1"].data_type = "s"         # "s" = string
    bufor3 = io.BytesIO()
    wb3.save(bufor3)

    wb4 = load_workbook(io.BytesIO(bufor3.getvalue()))
    try:
        assert wb4.active["A1"].data_type == "s"
        assert wb4.active["A1"].value == "=1+1"
    finally:
        wb4.close()
```

### 2.9. Testy wydajności jako budżet

W module 19 mierzyliście. Ale **pomiar bez progu nie chroni przed niczym**. Wiesz, że teraz jest 4 sekundy — ale nie masz mechanizmu, który krzyknie, gdy za pół roku ktoś doda `ws.cell(...)` w pętli i zrobi się 40 sekund.

Rozwiązanie: **budżet jako test**. Analogia: **termin dostawy**. Nie „w miarę szybko", ale „do 14:00". Konkretna liczba, którą albo dotrzymujesz, albo nie.

```python
import time
import pytest
from openpyxl import Workbook


@pytest.mark.slow
@pytest.mark.performance
def test_budzet_zapisu_100k_wierszy(tmp_path):
    """BUDZET: 100 000 wierszy w trybie write_only < 10 s.

    Jeśli ten test pada, sprawdź:
      1. czy lxml jest aktywny (from openpyxl import LXML),
      2. czy nie tworzysz obiektow Font w petli (modul 09),
      3. czy nie uzywasz trybu normalnego zamiast write_only (modul 19).
    """
    cel = tmp_path / "duzo.xlsx"

    t0 = time.perf_counter()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Dane")
    ws.append(["ID", "Kwota"])
    for i in range(1, 100_001):
        ws.append([i, i * 0.5])
    wb.save(cel)
    wb.close()
    czas = time.perf_counter() - t0

    assert czas < 10.0, (
        f"Zapis 100k wierszy zajął {czas:.2f} s przy budżecie 10 s.\n"
        f"lxml aktywny: {__import__('openpyxl').LXML}"
    )
```

**Trzy decyzje projektowe, które czynią ten test użytecznym:**

1. **Marker `slow` + `performance`.** Rejestrujesz je w konfiguracji i możesz je pomijać lokalnie (`pytest -m "not slow"`), a uruchamiać w CI. Bez tego test wydajnościowy dokłada sekundę przy **każdym** `pytest`, i po tygodniu ludzie zaczynają go usuwać.
2. **Komunikat z konkretnym budżetem i czasem.** „4.2 s < 10 s" — a nie „assert 4.2 < 10".
3. **Komunikat podpowiada, jak naprawić.** Trzy typowe przyczyny wprost w komunikacie. To jest ta sama zasada, co w `--update-golden`: **test ma być instrukcją**.

**Czego NIE testować wydajnościowo:**

- **Nie testuj budżetu w milisekundach.** CI ma zmienne obciążenie (inny runner, zimny cache, kolizje przy równoległych jobach). „< 0.5 s" będzie migać jak choinka. Budżet ustawiaj **z zapasem 2–3×** wobec tego, co realnie mierzysz.
- **Nie testuj na danych, których nie da się wygenerować szybko.** Test, który sam zajmuje 30 sekund na **przygotowanie** danych, jest bezużyteczny.
- **Nie testuj pamięci w CI** (`tracemalloc` w CI jest nieprzewidywalny). Pamięć mierz lokalnie i traktuj jako wynik do przeglądu, nie jako asercję.

### 2.10. Struktura projektu — warunek wstępny testowalności

I teraz najważniejszy fragment teoretyczny tego modułu. Bo prawie wszystko, co napisałem wyżej, **da się zrobić dopiero wtedy**, gdy kod ma właściwą strukturę.

Zobacz, jak wygląda **nietestowalny** kod:

```python
# ❌ NIETESTOWALNE
def generuj_raport():
    dane = pobierz_z_api()                     # I/O - nie da sie podmienic
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaz"
    ws.append(["Data", "Region", "Kwota"])
    for d in dane:
        ws.append([d["data"], d["region"], d["kwota"]])
        if d["kwota"] > 1000:
            ws[d["data"].cell if False else "A1"].fill = PatternFill(...)   # i tak dalej
    wb.save("output/raport.xlsx")               # zapis na dysk - nie da sie usunac
    return "output/raport.xlsx"
```

Dlaczego nie da się tego przetestować?

- **Woła sieć** — test zależy od zewnętrznego API.
- **Czyta zegar / rzeczywiste dane** — nie da się ustalić, co „powinno" być.
- **Pisze na dysk** — test zostawia śmieci i zależy od stanu.
- **Miesza trzy światy**: pobieranie danych, logika biznesowa, renderowanie. Zmiana w renderowaniu może zepsuć pobieranie.

A teraz wersja **testowalna** — ta sama funkcjonalność, inna struktura:

```python
# ✅ TESTOWALNE
from dataclasses import dataclass
from datetime import date
from openpyxl import Workbook


@dataclass(frozen=True)
class WierszSprzedazy:
    """Model domenowy - ZERO openpyxl."""
    data: date
    region: str
    kwota: float


@dataclass(frozen=True)
class SpecyfikacjaRaportu:
    """CO ma byc w raporcie, nie JAK wyglada."""
    tytul: str
    naglowki: tuple[str, ...]
    wiersze: tuple[WierszSprzedazy, ...]
    dzien: date                       # ⬅ czas WSTRZYKNIETY, nie z zegara


def zbuduj_specyfikacje(tytul: str, wiersze, dzien: date) -> SpecyfikacjaRaportu:
    """Czysta logika: zamienia dane wejsciowe w opis raportu."""
    posortowane = tuple(sorted(wiersze, key=lambda w: (w.data, w.region)))
    return SpecyfikacjaRaportu(
        tytul=tytul,
        naglowki=("Data", "Region", "Kwota"),
        wiersze=posortowane,
        dzien=dzien,
    )


def zapisz_xlsx(spec: SpecyfikacjaRaportu, cel) -> None:
    """Renderowanie: TYLKO tutaj pojawia sie openpyxl."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaż"
    ws.append(list(spec.naglowki))
    for w in spec.wiersze:
        ws.append([w.data, w.region, w.kwota])
    wb.properties.creator = "Automat"
    wb.properties.created = spec.dzien
    wb.save(cel)
    wb.close()
```

I teraz cała różnica:

```python
# Test logiki - bez openpyxl, bez dysku, w mikrosekundach:
def test_specyfikacja_sortuje_wiersze():
    wiersze = [
        WierszSprzedazy(date(2026, 3, 1), "Południe", 100.0),
        WierszSprzedazy(date(2026, 1, 1), "Północ", 200.0),
    ]
    spec = zbuduj_specyfikacje("Q1", wiersze, date(2026, 4, 1))
    assert spec.wiersze[0].region == "Północ"          # wcześniejsza data
    assert spec.dzien == date(2026, 4, 1)

# Test renderowania - openpyxl, ale bez dysku i sieci:
def test_zapis_uklada_kolumny_w_kolejnosci_naglowkow():
    spec = SpecyfikacjaRaportu(
        tytul="Q1",
        naglowki=("Data", "Region", "Kwota"),
        wiersze=(WierszSprzedazy(date(2026, 1, 1), "Północ", 200.0),),
        dzien=date(2026, 4, 1),
    )
    bufor = io.BytesIO()
    zapisz_xlsx(spec, bufor)

    wb = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        ws = wb.active
        assert ws["B2"].value == "Północ"
    finally:
        wb.close()
```

**Zauważ trzy rzeczy, które zrobiły tę różnicę:**

1. **`dzien` jest parametrem, nie `date.today()`.** Zegar został **wstrzyknięty** (ang. *dependency injection* — dosłownie „podanie zależności z zewnątrz"). To rozwiązuje pułapkę „testy zależne od zegara" z modułu 19 i z sekcji 6.
2. **`zapisz_xlsx` przyjmuje `cel`, nie `"output/raport.xlsx"`.** Obiekt plikopodobny albo ścieżka — decyduje **wywołujący**. Dzięki temu test może podać `BytesIO` i **test automatycznie nie pisze na dysk**.
3. **`openpyxl` pojawia się w JEDNEJ funkcji.** Możesz tę jedną funkcję przetestować raz, a **całą logikę** — osobno, taniej, bez plików. To jest fundament modułu 26: **`import openpyxl` wolno tylko w warstwie infrastruktury**.

> **Zasada, którą zapamiętaj na cały kurs:** jeśli w funkcji widzisz jednocześnie `date.today()`, `requests.get(...)` i `wb.save("stała/ścieżka")`, to nie napiszesz dla niej testu. Nigdy. Ani dziś, ani za rok.

### 2.11. CI/CD bez Excela

Teraz uruchomimy to wszystko na maszynie, która nie ma Excela i nigdy go nie miała.

```yaml
# .github/workflows/testy.yml
name: testy

on:
  push:
    branches: [main]
  pull_request:

jobs:
  test:
    runs-on: ubuntu-latest          # ⬅ ZADNEGO Excela na tym obrazie. I nie trzeba.
    strategy:
      fail-fast: false              # jeden padniety Python nie ukrywa reszty
      matrix:
        python-version: ["3.9", "3.10", "3.11", "3.12"]
        # 3.13+ dodaj po sprawdzeniu changelogu openpyxl dla swojej wersji

    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}
          cache: pip

      - name: Instalacja zależności
        run: |
          python -m pip install --upgrade pip
          pip install "openpyxl==3.1.5" pillow lxml
          pip install pytest pytest-cov hypothesis

      - name: Weryfikacja, że lxml NAPRAWDĘ działa
        run: |
          python -c "from openpyxl import LXML; assert LXML, 'lxml nieaktywny w CI!'; print('lxml OK')"

      - name: Testy z pokryciem
        run: pytest --cov=raporty --cov-report=term-missing --cov-fail-under=80

      - name: Kontrola jakości statycznej
        run: |
          pip install ruff mypy types-openpyxl
          ruff check .
          ruff format --check .
          mypy raporty
```

**Pięć decyzji w tym pliku, z których każda ma powód:**

1. **`pip install "openpyxl==3.1.5"` — przypięta wersja.** Bez tego CI może pobrać inną wersję niż Twój laptop i „nagle" zaczną padać testy. Przypięcie to nie konserwatyzm — to **porównywalność**.
2. **`python -c "assert LXML"` — osobny krok.** To jest ten „test cichej regresji wydajności" z modułu 19. Bez niego trader: obraz zmieni się tak, że `lxml` przestanie być aktywny, raporty zwolnią się trzykrotnie, **a CI nadal będzie zielone**. Ten jeden krok jest wart więcej niż wszystkie optymalizacje razem.
3. **Macierz wersji Pythona.** Testy przechodzące na 3.12 mogą padać na 3.9 — najczęściej z powodu składni (`X | Y` w typach wymaga `from __future__ import annotations`) albo zmian w `datetime`/`zipfile`. Macierz to wyłapie **przed** użytkownikiem.
4. **`--cov-fail-under=80`.** Pokrycie jako **bramka**, nie jako ozdoba. Ale uwaga — to jest **wskaźnik, nie cel** (patrz pułapka 14 w sekcji 6).
5. **Brak jakiegokolwiek kroku instalacji Office.** To nie przeoczenie — to **dowód**. Cały raport powstaje i jest weryfikowany na maszynie bez Excela.

**Pułapka numer jeden w CI — `fail-fast: false`.** Domyślnie GitHub Actions **przerywa całą macierz**, gdy jeden job padnie. Efekt: naprawiasz błąd na 3.9, pushujesz, widzisz błąd na 3.10, naprawiasz, pushujesz, widzisz błąd na 3.11... `fail-fast: false` daje Ci **wszystkie** błędy naraz.

**Pułapka numer dwa — locale.** Twój laptop ma polskie ustawienia regionalne, CI ma `C.UTF-8`. Jeśli gdziekolwiek **porównujesz wyrenderowany tekst** (np. „1 200,50 zł"), CI wybuchnie, a Ty nie zrozumiesz, dlaczego. Lekarstwo: **testuj `number_format` (kod), nie jego renderowanie.** Kod `#,##0.00` jest identyczny w każdym locale; wynik jego zastosowania — nie.

### 2.12. Jakość statyczna — `ruff`, `mypy`, `pre-commit`

Testy sprawdzają **zachowanie**. Ale część błędów da się złapać **zanim kod w ogóle się uruchomi**. Trzy narzędzia, które zamykają resztę dziury:

#### `ruff` — linter i formater w jednym

`ruff` w kilka sekund sprawdza styl, kolejność importów, nieużywane zmienne, nadmiernie złożone konstrukcje. Zastępuje `flake8`, `isort`, `pyupgrade` i `black` — i jest napisany w Rust, więc jest szybki jak błyskawica.

```toml
# pyproject.toml
[tool.ruff]
target-version = "py39"
line-length = 100

[tool.ruff.lint]
# E/W - bledy i ostrzezenia pycodestyle
# F   - pyflakes (nieuzywane importy, nazwy)
# I   - sortowanie importow (isort)
# UP  - modernizacja skladni (pyupgrade)
# B   - bugbear (czeste pulapki logiczne)
# SIM - uproszczenia
# PTH - uzywaj pathlib zamiast os.path
select = ["E", "W", "F", "I", "UP", "B", "SIM", "PTH"]
ignore = [
    "E501",   # dlugosc linii - zostawiona formaterowi
]
```

Uwaga na wersjonowanie: w `ruff >= 0.2` sekcja `select` przenosi się pod `[tool.ruff.lint]` (jak wyżej). W starszych wersjach było `[tool.ruff]` → `select`. Sprawdź wersję.

**Jedna reguła z tej listy jest szczególnie cenna dla nas:** `B` (bugbear) łapie m.in. porównania z `None` przez `==`, nieużywane zmienne w pętlach i — co ważne w naszym kontekście — **mutable default arguments**. A `PTH` zmusza do `pathlib` zamiast `os.path`, co jest zgodne z konwencją z modułu 04.

#### `mypy` — typy jako dokumentacja, którą można zweryfikować

I tu wchodzimy w temat, w którym trzeba być **uczciwym**.

openpyxl **nie ma pełnych, wbudowanych adnotacji typów**. Biblioteka powstała przed erą powszechnego typowania i ma bardzo dynamiczne API (np. `wb["Dane"]` może zwrócić `Worksheet`, `ReadOnlyWorksheet`, `WriteOnlyWorksheet` albo `Chartsheet` — zależnie od trybu!). Dlatego powstał **osobny pakiet stubów**:

```bash
pip install types-openpyxl
```

`types-openpyxl` to pakiet z **plikami `.pyi`** (stubami) dostarczającymi adnotacje **bez modyfikowania samej biblioteki**. Wersja `3.1.5.xxxx` odpowiada `openpyxl==3.1.5` — **dobieraj wersje parami**, bo stuby opisują konkretną wersję API.

Oto realny, udokumentowany problem, na który się natkniesz:

```python
from openpyxl import load_workbook

wb = load_workbook("raport.xlsx")
ws = wb.active                  # typ: Worksheet | ReadOnlyWorksheet | WriteOnlyWorksheet | Chartsheet
naglowki = [c.value for c in ws[1]]   # ⚠️ mypy: Chartsheet nie wspiera indeksowania!
```

Stuby są **poprawne**, ale przez to niewygodne: `wb.active` naprawdę może zwrócić `Chartsheet`. Rozwiązanie, które rekomenduje sam zespół typeshed:

```python
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

wb = load_workbook("raport.xlsx")
ws = wb.active
assert isinstance(ws, Worksheet)      # ⬅ "zawężenie typu" - i mypy jest spokojny
naglowki = [c.value for c in ws[1]]

wb.close()
```

**To nie jest hack.** To jest **sprawdzenie założenia**. Jeśli w pliku raportu `active` okaże się `Chartsheet`, to znaczy, że coś jest nie tak z raportem — i `assert` powie Ci to **natychmiast**, w miejscu, gdzie potrzebujesz `Worksheet`. Test i typ sprawdzają się nawzajem.

**A co z `# type: ignore`?** Używaj go — ale **świadomie i z kodem błędu**:

```python
# ❌ ZLE: cicha tasma na czujniku dymu
ws["A1"] = wartosc_produktu   # type: ignore

# ✅ DOBRZE: wiem, co wylaczam i dlaczego
# types-openpyxl opisuje wb.defined_names jako Incomplete (stuby czesciowe),
# wiec dostep do .items() nie jest sprawdzany - modul 26 tego nie potrzebuje.
for nazwa, definicja in wb.defined_names.items():   # type: ignore[union-attr]
    print(nazwa)
```

Analogia: `# type: ignore` to **taśma na czujniku dymu**. Czasem jest potrzebna (masz uzasadnienie), ale **musisz wiedzieć, że ją założyłeś** i **musisz wiedzieć, na którym czujniku**. `ignore` bez kodu błędu i bez uzasadnienia to taśma założona „na wszelki wypadek" — i to ona pewnego dnia sprawi, że pożar przejdzie niezauważony.

**Praktyczna zasada:** włączaj `mypy` **stopniowo**. Zamiast `--strict` na cały projekt (co da 400 błędów i zniechęci wszystkich), zacznij od:

```toml
# pyproject.toml
[tool.mypy]
python_version = "3.9"
# Na start: sprawdzaj tylko to, co ma adnotacje
disallow_untyped_defs = false
warn_unused_configs = true
warn_redundant_casts = true
warn_unused_ignores = true        # ⬅ KLUCZOWE: krzyczy o zbednych # type: ignore
```

**`warn_unused_ignores = true` to najważniejsza opcja na tej liście.** Sprawia, że gdy zaktualizujesz `types-openpyxl` i problem zniknie, `mypy` **powie Ci, że `# type: ignore` jest już niepotrzebny**. Bez tego Twoje taśmy na czujnikach zostaną tam **na zawsze**.

#### `pre-commit` — bo nikt nie pamięta

**`pre-commit`** uruchamia narzędzia **automatycznie przed każdym commitem**. Nie musisz pamiętać o `ruff` — on sam sprawdzi i naprawi to, co umie naprawić automatycznie.

```yaml
# .pre-commit-config.yaml
repos:
  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.6.9            # podmień na aktualną rewizję
    hooks:
      - id: ruff
        args: [--fix]       # poprawia, co umie
      - id: ruff-format

  - repo: https://github.com/pre-commit/mirrors-mypy
    rev: v1.11.2           # podmień na aktualną rewizję
    hooks:
      - id: mypy
        additional_dependencies: ["types-openpyxl"]
        args: [--config-file=pyproject.toml]
        files: ^raporty/    # nie sprawdzaj testów

  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v4.6.0            # podmień na aktualną rewizję
    hooks:
      - id: end-of-file-fixer
      - id: trailing-whitespace
      - id: check-added-large-files
        args: [--maxkb=500]   # ⬅ nie commituj plikow .xlsx!
```

**Ostatni hook jest szczególnie ważny w naszym kontekście.** Bez `check-added-large-files` ktoś (być może Ty) scommituje wygenerowany `raport_2026.xlsx` — i za rok repozytorium waży gigabajt, a historia gita jest nieodwracalnie zanieczyszczona. **Raporty to artefakty, nie kod źródłowy.**

**Instalacja i użycie:**

```bash
pip install pre-commit
pre-commit install            # od teraz dziala przy kazdym git commit
pre-commit run --all-files    # jednorazowo na calym repo
```

### 2.13. Definicja ukończenia raportu

To jest lista, którą powinieneś powiesić nad biurkiem. Nie „testy przechodzą", ale **konkretny stan artefaktu, który wysyłasz człowiekowi**.

| # | Kryterium | Jak to sprawdzić |
|---|---|---|
| 1 | **Plik otwiera się bez ostrzeżeń Excela** | otwórz ręcznie. Jedyny krok, którego nie zautomatyzujesz w pełni. |
| 2 | **Liczby są liczbami, nie tekstem** | test: `isinstance(v, (int, float))` dla kolumn liczbowych |
| 3 | **Daty są datami** | test: `isinstance(v, (datetime, date))` |
| 4 | **Nie ma `#REF!` ani `#N/A`** | test: brak wartości zaczynających się od `#` |
| 5 | **Nie ma `#DIV/0!`** | jak wyżej (openpyxl czyta błąd jako **string**) |
| 6 | **Formaty liczb budżetu są spójne** | test: `number_format` z rejestru stylów (moduł 09) |
| 7 | **Reguły formatowania warunkowego są na miejscu** | test: liczba reguł + `str(cf.sqref)` |
| 8 | **Tabele mają właściwe `ref`** | test: `ws.tables["X"].ref` |
| 9 | **Ustawienia druku są kompletne** | test: `print_area`, `print_title_rows`, `fitToPage` |
| 10 | **Właściwości dokumentu są ustawione** | test: `wb.properties.creator is not None` |
| 11 | **Nie ma pustych komórek tam, gdzie ma być wartość** | test z 2.3.2 |
| 12 | **Raport jest deterministyczny** | zbuduj dwa razy, `odcisk_bajtowy` równy |
| 13 | **Dane użytkownika nie stały się formułami** | test pułapki z 2.8 |
| 14 | **Wzorzec (`golden`) jest aktualny i zatwierdzony** | `git status` czysty |
| 15 | **Użytkownik dostał dokument gotowy do wysłania** | otwórz, przejrzyj, wyślij |

**Kryterium 1 i 15 są nieusuwalne.** Reszta to Twój hamulec, żeby kryterium 1 i 15 zajmowały pięć sekund, a nie pół godziny.

### 2.14. Trzy kategorie prawdy — powtórka operacyjna

| Kategoria | W kontekście testowania |
|---|---|
| **potrafi** | `BytesIO` w obie strony, asercje na wartościach/typach/formatach, iteracja reguł CF i walidacji, `ws.tables`, odczyt ustawień druku, `tmp_path`, `hypothesis`, hash części, snapshot tekstowy, CI bez Excela |
| **potrafi częściowo** | porównanie plików (hash części i model — tak; bajty — nie), `ws.print_area` (kwalifikowana postać!), `len(ws.conditional_formatting)` (liczy zakresy), `read_only` (bez stylów i reguł), `mypy` (tylko ze stubami i zawężaniem typu) |
| **gubi / nie ma** | deterministycznych bajtów, możliwości przypięcia `properties.modified`, oficjalnego helpera testowego, dostępu do stylów w `read_only`, sygnału o utraconych sparkline'ach poza ostrzeżeniem |

## 3. Przykłady krok po kroku

### Przykład 1 — pierwszy test w `BytesIO` (🟢)

Budujemy najprostszą możliwą rzecz: skoroszyt w pamięci, zapis do bufora, odczyt z bufora, dwie asercje. Ten plik jest **szablonem** dla wszystkich pozostałych.

```python
"""Modul 20 - pierwszy test: skoroszyt w pamieci, bez dysku.

Uruchom:  pytest tests/test_20_bytesio.py -v
"""

from __future__ import annotations

import io

from openpyxl import Workbook, load_workbook


def utworz_skoroszyt() -> bytes:
    """Buduje skoroszyt w pamieci i zwraca jego BAJTY.

    Nic nie trafia na dysk. Bajty sa niemutowalne, wiec kazdy test
    dostaje wlasna, niezalezna kopie - zero wspoldzielonego stanu.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    ws["A1"] = "Produkt"
    ws["B1"] = "Kwota"
    ws["A2"] = "Kawa"
    ws["B2"] = 19.99

    bufor = io.BytesIO()
    wb.save(bufor)                   # save() przyjmuje obiekt plikopodobny
    # getvalue() zwraca caly bufor, niezaleznie od pozycji kursora
    return bufor.getvalue()


def test_a1_zawiera_naglowek_produkt():
    """Czy wartosc w komorce jest tym, co chcielismy?"""
    bajty = utworz_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Dane"]
        assert ws["A1"].value == "Produkt"
    finally:
        wb.close()                   # zwolnij uchwyt ZIP (modul 19)


def test_kwota_jest_floatem_a_nie_tekstem():
    """NAJWAZNIEJSZY test w calym kursie.

    Jesli kwota jest tekstem, to Excel pokaze ja w lewo wyrownana,
    formuly ja pominą, a sortowanie bedzie alfabetyczne.
    Ten test to lapie.
    """
    bajty = utworz_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Dane"]
        wartosc = ws["B2"].value

        assert isinstance(wartosc, float), (
            f"Kwota jest typu {type(wartosc).__name__}, a powinna byc float. "
            "Prawdopodobnie zapisano ja jako tekst."
        )
        assert wartosc == 19.99
    finally:
        wb.close()


def test_wymiary_sa_zgodne_z_oczekiwaniem():
    """Wymiary to najtanszy sposob wykrycia 'raport sie rozjechal'."""
    bajty = utworz_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Dane"]
        assert ws.max_row == 2
        assert ws.max_column == 2
    finally:
        wb.close()
```

**Co się dzieje w pamięci.** `Workbook()` tworzy pusty model (~kilka KB). Po wpisaniu czterech komórek `ws._cells` ma cztery obiekty `Cell`. `wb.save(bufor)` serializuje cały model do XML-a i pakuje go do ZIP-a **wewnątrz bufora** — bufor ma teraz ~4–5 KB. `load_workbook(io.BytesIO(bajty))` tworzy **drugi, niezależny model** — bo `bajty` to niezmienny `bytes`, a `BytesIO` to tylko „nakładka" na nie.

**Co trafi do pliku.** Nic. Ale w buforze jest **kompletny, poprawny plik `.xlsx`** — ten sam, który powstałby na dysku, bajt po bajcie (poza znacznikami czasu).

**Trzy rzeczy do przemyślenia:**

1. **`getvalue()` a `read()`.** Sprawdź sam: dodaj `print(bufor.read())` przed `getvalue()` i zobacz, co dostaniesz. Potem dodaj `bufor.seek(0)` i sprawdź ponownie. Zrozumienie pozycji kursora to 90% wszystkich problemów z „pustym plikiem w odpowiedzi HTTP" (moduł 26).
2. **`isinstance(wartosc, float)` — dlaczego `float`, a nie `(int, float)`?** Bo `19.99` zapisane przez openpyxl wróci jako `float`, a `19` jako `int` (moduł 05). Świadome zawężenie typu w asercji jest **dokumentacją Twojego kontraktu**: „ta kolumna ma być zmiennoprzecinkowa". Jeśli kiedyś wróci jako `int`, to znaczy, że dane źródłowe się zmieniły — i chcesz o tym wiedzieć.
3. **`finally: wb.close()` — nawet w teście.** To trzeci raz w kursie, ale tu ma dodatkowy wymiar: `load_workbook` na buforze tworzy `ZipFile`, który trzyma **referencję do bufora**. W testach uruchamianych setki razy z rzędu brak `close()` to rosnąca lista otwartych uchwytów — i test, który `passes` lokalnie, a w CI na 4 000 testów wywala się na `Too many open files`.

### Przykład 2 — fixture, fabryka i pięć testów (🟡)

Teraz warsztat. Budujemy `conftest.py` z fabryką i piszemy pięć testów na realnym scenariuszu raportu sprzedażowego.

```python
# conftest.py  <- pytest znajduje go automatycznie w katalogu testow i wyzej
"""Fixture'y wspolne dla calego zestawu testow."""

from __future__ import annotations

import io

import pytest
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.datetime import to_excel


# Domyslne dane: swiadomie MALE (5 wierszy), bo testy maja byc szybkie.
DOMYSLNE_WIERSZE = [
    ["Data", "Region", "Kwota", "Status"],
    ["2026-01-05", "Północ", 1200.50, "OK"],
    ["2026-01-06", "Południe", 860.00, "OK"],
    ["2026-01-07", "Wschód", 2450.75, "Pilne"],
    ["2026-01-08", "Północ", 0.00, "OK"],
]


@pytest.fixture
def zbuduj_skoroszyt():
    """FABRYKA JAKO FIXTURE.

    Nie zwraca skoroszytu - zwraca FUNKCJE, ktora go buduje.
    Dzieki temu kazdy test tworzy dokladnie taki arkusz,
    jakiego potrzebuje, a nie taki, jaki wymyslilismy z gory.
    """
    def build(rows=None, sheet="Sprzedaż", ustaw=None) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet

        for wiersz in (rows if rows is not None else DOMYSLNE_WIERSZE):
            ws.append(wiersz)

        if ustaw is not None:
            ustaw(wb, ws)          # hak: test dokłada style/reguly/tabele

        bufor = io.BytesIO()
        wb.save(bufor)
        return bufor.getvalue()    # bajty = niemutowalne = bezpieczne

    return build


@pytest.fixture
def wczytaj_raport():
    """Pomocnik: wczytuje bajty i GWARANTUJE close().

    Uzycie:
        with wczytaj_raport(bajty) as ws:
            assert ...
    """
    from contextlib import contextmanager
    from openpyxl import load_workbook

    @contextmanager
    def _wczytaj(bajty: bytes, nazwa: str = "Sprzedaż"):
        wb = load_workbook(io.BytesIO(bajty))
        try:
            yield wb[nazwa]
        finally:
            wb.close()

    return _wczytaj
```

```python
# tests/test_20_raport.py
"""Piec testow na raporcie sprzedazowym."""

from __future__ import annotations

import datetime as dt
import io

from openpyxl import load_workbook


# --- TEST 1: struktura ------------------------------------------------------
def test_ma_dokladnie_jeden_arkusz_o_wlasciwej_nazwie(zbuduj_skoroszyt):
    bajty = zbuduj_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        assert wb.sheetnames == ["Sprzedaż"], (
            f"Oczekiwano jednego arkusza 'Sprzedaż', jest {wb.sheetnames}"
        )
    finally:
        wb.close()


# --- TEST 2: wartosci -------------------------------------------------------
def test_wiersz_danych_zgadza_sie_z_wejsciem(zbuduj_skoroszyt):
    bajty = zbuduj_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Sprzedaż"]
        assert ws["B2"].value == "Północ"
        assert ws["C2"].value == 1200.50
        assert ws["D2"].value == "OK"
    finally:
        wb.close()


# --- TEST 3: typy -----------------------------------------------------------
def test_daty_sa_datami_a_kwoty_floatami(zbuduj_skoroszyt):
    """DATY jako datetime, NIE jako tekst. Kwoty jako float, NIE jako tekst.

    To jest test 'przeciwko zrodlu prawdy' - chroni przed cicha konwersja
    (modul 05: '1' != 1).
    """
    bajty = zbuduj_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Sprzedaż"]
        # Uwaga: w DOMYSLNE_WIERSZE data jest podana jako string "2026-01-05".
        # Po round-tripie openpyxl zwroci ja jako string - bo tak ja wpisalismy.
        # Test dokumentuje ten fakt swiadomie:
        assert isinstance(ws["A2"].value, str) or isinstance(ws["A2"].value, dt.datetime)

        # A teraz test na prawdziwym datetime:
        from openpyxl.utils.datetime import to_excel
        wiersze = [
            ["Data", "Kwota"],
            [dt.date(2026, 1, 5), 100.0],
        ]
        bajty2 = zbuduj_skoroszyt(rows=wiersze, sheet="T2")

        wb2 = load_workbook(io.BytesIO(bajty2))
        try:
            ws2 = wb2["T2"]
            assert isinstance(ws2["A2"].value, (dt.datetime, dt.date)), (
                f"Data przyszła jako {type(ws2['A2'].value).__name__}"
            )
            assert isinstance(ws2["B2"].value, float)
        finally:
            wb2.close()
    finally:
        wb.close()


# --- TEST 4: formaty --------------------------------------------------------
def test_kolumna_kwoty_ma_format_walutowy(zbuduj_skoroszyt):
    """Testuj KOD FORMATU, nie jego wyglad.

    '#,##0.00' jest identyczny w kazdym locale.
    '1 200,50 zl' - nie jest (CI ma C.UTF-8, Ty masz pl_PL).
    """
    def sformatuj(wb, ws):
        from openpyxl.styles import Font, Alignment
        for komorka in ws[1]:
            komorka.font = Font(bold=True)
            komorka.alignment = Alignment(horizontal="center")
        for wiersz in ws.iter_rows(min_row=2, min_col=3, max_col=3):
            for komorka in wiersz:
                komorka.number_format = "#,##0.00"

    bajty = zbuduj_skoroszyt(ustaw=sformatuj)

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Sprzedaż"]
        assert ws["C1"].number_format == "General"        # naglowek
        assert ws["C2"].number_format == "#,##0.00"       # dane
        assert ws["C5"].number_format == "#,##0.00"
        assert ws["C1"].font.bold is True
        assert ws["C1"].alignment.horizontal == "center"
    finally:
        wb.close()


# --- TEST 5: brak brakow ----------------------------------------------------
def test_zadna_komorka_w_kolumnie_kwot_nie_jest_pusta(zbuduj_skoroszyt):
    """Jedna z najcenniejszych asercji - patrz sekcja 2.3.2."""
    bajty = zbuduj_skoroszyt()

    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Sprzedaż"]
        braki = [
            c.coordinate
            for wiersz in ws.iter_rows(min_row=2, min_col=3, max_col=3)
            for c in wiersz
            if c.value is None
        ]
        assert not braki, f"Brak kwoty w komorkach: {braki[:10]}"
    finally:
        wb.close()
```

**Co się dzieje w pamięci.** Fabryka `zbuduj_skoroszyt` tworzy i **porzuca** `Workbook` za każdym wywołaniem — po zapisie do bufora obiekt `wb` wychodzi z zakresu i idzie do `gc`. W pamięci zostają tylko `bytes` (kilka KB). `load_workbook(io.BytesIO(bajty))` tworzy **nowy**, niezależny model, który jest zamykany w `finally`. Szczyt pamięci na jeden test: kilkaset KB.

**Co trafi do pliku.** Nic. Ale **każdy z pięciu testów dostał własną, świeżą kopię** dokumentu — mimo że fabryka jest **jedna**.

**Trzy rzeczy do przemyślenia:**

1. **Dlaczego fabryka zwraca `bytes`, a nie `io.BytesIO`?** Bo `bytes` jest **niemutowalne**. Gdyby zwracała bufor, test mógłby zrobić `bufor.seek(0); bufor.read()` i **zepsuć pozycję kursora** dla następnego użycia. Z `bytes` każdy `BytesIO(bajty)` tworzy **świeży** bufor. To jest cała idea fixture'u: brak wspólnego, mutowalnego stanu.
2. **Dlaczego `wczytaj_raport` jest context managerem?** Bo wymusza `close()`. To jest konwersja **dobrego nawyku w gwarancję**. W testach, które ktoś napisze za rok, `close()` zadziała sam. Zapamiętaj tę technikę — to `@contextmanager` ze standardowego `contextlib`.
3. **Test 3 jest napisany nieporadnie — i to celowo.** Zobacz, jak walczy z tym, że dane wejściowe mają datę jako **string**. To jest **realistyczny problem**: większość danych „z Excela" przychodzi jako tekst. I pokazuje, dlaczego **dane wejściowe trzeba jawnie konwertować** — bo openpyxl tego nie zrobi za Ciebie (moduł 05: Excel nie wie, co miałaś na myśli).

### Przykład 3 — `podsumowanie_skoroszytu()` i test regresji z modułu 11 (🔴)

Teraz narzędzie numer jeden. Budujemy odcisk skoroszytu z trzech warstw opisanych w 2.5, generujemy raport gotowy do druku (jak w module 11) i pilnujemy go golden testem.

```python
# tests/odcisk.py
"""Odcisk skoroszytu - 'fingerprint' pliku .xlsx.

Trzy warstwy:
  1. inwentarz czesci ZIP   (struktura - z modulu 02)
  2. model skoroszytu       (tresc przez openpyxl)
  3. stylistyka z styles.xml (liczba DEFINICJI stylow)
"""

from __future__ import annotations

import hashlib
import zipfile
from datetime import datetime, date
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Czesci ZMIENNE - wykluczone z odcisku i z hasha.
# openpyxl wpisuje do docProps/core.xml biezacy czas przy KAZDYM save(),
# wiec te czesci zmieniaja sie miedzy uruchomieniami bez powodu.
CZESCI_ZMIENNE = frozenset({
    "docProps/core.xml",       # dcterms:modified = "teraz"
    "docProps/app.xml",        # wersja aplikacji
    "xl/calcChain.xml",        # kolejnosc przeliczania formul
})


# ============================================================
# WARSTWA 1: inwentarz czesci ZIP
# ============================================================
def _inwentarz_czesci(zf: zipfile.ZipFile) -> dict:
    nazwy = zf.namelist()
    return {
        "media": sorted(n for n in nazwy if n.startswith("xl/media/")),
        "wykresy": sorted(n for n in nazwy if n.startswith("xl/charts/chart")),
        "tabele_zip": sorted(n for n in nazwy if n.startswith("xl/tables/")),
        "arkusze_zip": sorted(n for n in nazwy if n.startswith("xl/worksheets/sheet")),
        "ma_shared_strings": "xl/sharedStrings.xml" in nazwy,
        "ma_vba": "xl/vbaProject.bin" in nazwy,
    }


# ============================================================
# WARSTWA 3: stylistyka z xl/styles.xml (parsujemy XML, nie prywatne atrybuty)
# ============================================================
def _stylistyka(zf: zipfile.ZipFile) -> dict:
    try:
        dane = zf.read("xl/styles.xml")
    except KeyError:
        return {}
    root = ET.fromstring(dane)

    def ile_dzieci(tag: str) -> int:
        el = root.find(f"{NS}{tag}")
        return len(el) if el is not None else 0

    return {
        "fonty": ile_dzieci("fonts"),
        "wypelnienia": ile_dzieci("fills"),
        "obramowania": ile_dzieci("borders"),
        "formaty_liczb": ile_dzieci("numFmts"),
        "style_komorek": ile_dzieci("cellXfs"),          # <xf> w cellXfs
    }


# ============================================================
# WARSTWA 2: model skoroszytu
# ============================================================
def _arkusz(ws) -> dict:
    reguly = [
        (str(cf.sqref), r.type, r.priority)
        for cf in ws.conditional_formatting
        for r in cf.rules
    ]
    walidacje = [
        (dv.type, str(dv.sqref))
        for dv in ws.data_validations.dataValidation
    ]
    komorki = sum(
        1
        for wiersz in ws.iter_rows()
        for c in wiersz
        if c.value is not None
    )
    return {
        "nazwa": ws.title,
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "komorki_z_danymi": komorki,
        "reguly_cf": len(reguly),                       # ⬅ NIE len(ws.conditional_formatting)!
        "typy_regul_cf": sorted({t for _, t, _ in reguly}),
        "walidacje": len(walidacje),
        "typy_walidacji": sorted({t for t, _ in walidacje}),
        "tabele": sorted(ws.tables.keys()),             # dict-like
        "scalenia": sorted(str(r) for r in ws.merged_cells.ranges),
        "zamrozone": str(ws.freeze_panes) if ws.freeze_panes else None,
        "obszar_druku": ws.print_area or None,          # ⬅ kwalifikowana postac!
        "powtarzane_wiersze": ws.print_title_rows,
        "powtarzane_kolumny": ws.print_title_cols,
        "orientacja": ws.page_setup.orientation,
        "fitToPage": bool(
            ws.sheet_properties.pageSetUpPr
            and ws.sheet_properties.pageSetUpPr.fitToPage
        ),
        "fitToWidth": ws.page_setup.fitToWidth,
        "fitToHeight": ws.page_setup.fitToHeight,
        "ukryte_kolumny": sorted(
            k for k, d in ws.column_dimensions.items() if d.hidden
        ),
        "widocznosc": ws.sheet_state,
    }


def podsumowanie_skoroszytu(sciezka: str | Path) -> dict:
    """Zwraca ODCISK skoroszytu jako slownik (deterministyczny)."""
    sciezka = Path(sciezka)

    with zipfile.ZipFile(sciezka) as zf:
        inwentarz = _inwentarz_czesci(zf)
        stylistyka = _stylistyka(zf)

    wb = load_workbook(sciezka)
    try:
        propiedades = {
            "tytul": wb.properties.title,
            "autor": wb.properties.creator,
            "utworzono": (
                wb.properties.created.isoformat()
                if isinstance(wb.properties.created, (datetime, date))
                else None
            ),
        }
        arkusze = [_arkusz(ws) for ws in wb.worksheets]
    finally:
        wb.close()

    return {
        "arkusze": arkusze,
        "wlasciwosci": propiedades,
        "style": stylistyka,
        "czesc": inwentarz,
    }


def odcisk_tekstowy(podsumowanie: dict) -> str:
    """Zamienia odcisk w TEKST - czytelny w git diff."""
    linie: list[str] = []

    linie.append("=== ARKUSZE ===")
    for a in podsumowanie["arkusze"]:
        linie.append(
            f"{a['nazwa']:<14} {a['max_row']:>6}x{a['max_column']:<3} "
            f"komorki={a['komorki_z_danymi']:<5} "
            f"cf={a['reguly_cf']}({','.join(a['typy_regul_cf']) or '-'}) "
            f"dv={a['walidacje']}({','.join(a['typy_walidacji']) or '-'})"
        )
        linie.append(
            f"{'':<14} tabele={a['tabele'] or '-'} "
            f"scalenia={a['scalenia'] or '-'} "
            f"zamrozone={a['zamrozone']}"
        )
        linie.append(
            f"{'':<14} druk: obszar={a['obszar_druku'] or '-'} "
            f"wiersze={a['powtarzane_wiersze'] or '-'} "
            f"orientacja={a['orientacja']} "
            f"fitToPage={a['fitToPage']} "
            f"({a['fitToWidth']}x{a['fitToHeight']})"
        )

    linie.append("=== STYLE (z xl/styles.xml) ===")
    for klucz, wartosc in sorted(podsumowanie["style"].items()):
        linie.append(f"{klucz:<16} {wartosc}")

    linie.append("=== CZESCI PLIKU ===")
    czesc = podsumowanie["czesc"]
    for klucz in sorted(czesc):
        linie.append(f"{klucz:<18} {czesc[klucz]}")

    linie.append("=== WLASCIWOSCI ===")
    for klucz, wartosc in sorted(podsumowanie["wlasciwosci"].items()):
        linie.append(f"{klucz:<16} {wartosc}")

    return "\n".join(linie) + "\n"


def odcisk_bajtowy(sciezka: str | Path) -> str:
    """SHA-256 z TRESCI rozpakowanych czesci (bez kompresji i metadanych).

    Odporny na roznice zlib miedzy maszynami - nadaje sie do CI.
    """
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in sorted(zf.namelist()):
            if nazwa in CZESCI_ZMIENNE:
                continue
            h.update(nazwa.encode("utf-8"))
            h.update(b"\x00")
            h.update(zf.read(nazwa))
            h.update(b"\x00")
    return h.hexdigest()
```

Teraz raport do przetestowania — ten z modułu 11 (gotowy do druku A4 poziomo, z powtarzanym nagłówkiem, freeze panes, autofiltrem i formatowaniem warunkowym):

```python
# tests/raport_11.py
"""Raport 'gotowy do druku' - material do testu regresji z modulu 20."""

from __future__ import annotations

import io
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

DZIEN = date(2026, 1, 15)          # ⬅ WSTRZYKNIETY czas, nie date.today()

OBRAMOWANIE_CIENKIE = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def zbuduj_raport(dzien: date = DZIEN) -> bytes:
    """Buduje raport, ktory ma przejsc test regresji.

    KLUCZOWE: `dzien` jest PARAMETREM. Bez tego raport nie bylby
    deterministyczny, a test regresji migalby jak choinka.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sprzedaż"

    # --- naglowek scalony ---
    ws.merge_cells("A1:D1")
    komorka_tytulu = ws["A1"]
    komorka_tytulu.value = f"Raport sprzedaży — {dzien.isoformat()}"
    komorka_tytulu.font = Font(size=14, bold=True, color="FFFFFF")
    komorka_tytulu.fill = PatternFill("solid", fgColor="1F4E78")
    komorka_tytulu.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # --- wiersz naglowkowy tabeli ---
    naglowki = ("Data", "Region", "Kwota", "Marża")
    ws.append(naglowki)
    for k in ws[2]:
        k.font = Font(bold=True, color="FFFFFF")
        k.fill = PatternFill("solid", fgColor="2F5496")
        k.alignment = Alignment(horizontal="center")
        k.border = OBRAMOWANIE_CIENKIE

    # --- dane ---
    dane = [
        (date(2026, 1, 5), "Północ", 1200.50, 0.12),
        (date(2026, 1, 6), "Południe", 860.00, 0.08),
        (date(2026, 1, 7), "Wschód", 2450.75, 0.19),
        (date(2026, 1, 8), "Północ", 0.00, 0.00),
        (date(2026, 1, 9), "Zachód", 3120.00, 0.24),
    ]
    for data_, region, kwota, marza in dane:
        ws.append([data_, region, kwota, marza])

    # formaty: kolumna C = waluta, kolumna D = procent
    for wiersz in ws.iter_rows(min_row=3, max_row=ws.max_row):
        wiersz[0].number_format = "yyyy-mm-dd"
        wiersz[2].number_format = "#,##0.00"
        wiersz[3].number_format = "0.0%"
        for c in wiersz:
            c.border = OBRAMOWANIE_CIENKIE

    # --- wymiary kolumn ---
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10

    # --- UX: freeze panes, filtr, brak linii siatki ---
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A2:D{ws.max_row}"

    # --- TABELA (kontrakt z uzytkownikiem, modul 12) ---
    tab = Table(displayName="SprzedazTabela", ref=f"A2:D{ws.max_row}")
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)

    # --- FORMATOWANIE WARUNKOWE (modul 13) ---
    # 1) kwota powyzej progu -> czerwone tlo
    ws.conditional_formatting.add(
        f"C3:C{ws.max_row}",
        CellIsRule(
            operator="greaterThan",
            formula=["2000"],
            fill=PatternFill("solid", fgColor="FFC7CE"),
            font=Font(color="9C0006"),
        ),
    )
    # 2) caly wiersz na czerwono, gdy marza < 10%
    ws.conditional_formatting.add(
        f"A3:D{ws.max_row}",
        FormulaRule(formula=["$D3<0.1"], fill=PatternFill("solid", fgColor="FFF2CC")),
    )

    # --- DRUK: gotowy do wyslania (modul 11) ---
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True       # ⬅ trzy razem!
    ws.print_area = "A1:D7"
    ws.print_title_rows = "2:2"                            # naglowek na kazdej stronie

    # --- WLASCIWOSCI: slad audytowy ---
    wb.properties.title = "Raport sprzedaży 2026-01"
    wb.properties.creator = "Automat raportowy"
    wb.properties.created = dzien
    wb.properties.modified = dzien

    bufor = io.BytesIO()
    wb.save(bufor)
    return bufor.getvalue()
```

I test regresji:

```python
# tests/test_20_regresja.py
"""Test regresji: raport z modulu 11 musi byc stabilny."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from openpyxl import load_workbook

from odcisk import odcisk_bajtowy, odcisk_tekstowy, podsumowanie_skoroszytu
from raport_11 import zbuduj_raport


@pytest.fixture
def raport_na_dysku(tmp_path) -> Path:
    """JEDYNE miejsce w testach, gdzie plik trafia na dysk.

    Bo `podsumowanie_skoroszytu` przyjmuje SCIEZKE (openpyxl czyta archiwum
    losowo, a my potrzebujemy inwentarza czesci i parsowania styles.xml).
    tmp_path jest unikalny dla kazdego wywolania testu.
    """
    cel = tmp_path / "raport.xlsx"
    cel.write_bytes(zbuduj_raport())
    return cel


def test_odcisk_strukturalny_zgadza_sie_z_wzorcem(golden, raport_na_dysku):
    """Zloty test: kazda zmiana w raporcie musi byc SWIADOMA."""
    odcisk = podsumowanie_skoroszytu(raport_na_dysku)
    golden("raport_11_odcisk", odcisk_tekstowy(odcisk))


def test_raport_jest_deterministyczny(raport_na_dysku, tmp_path):
    """Dwa uruchomienia = ten sam hash TRESCI czesci.

    Nie hash pliku! (docProps/core.xml i znaczniki ZIP zawsze sie roznia)
    Hash ROZPAKOWANYCH czesci, bez metadanych.
    """
    drugi = tmp_path / "raport2.xlsx"
    drugi.write_bytes(zbuduj_raport())

    assert odcisk_bajtowy(raport_na_dysku) == odcisk_bajtowy(drugi), (
        "Raport nie jest deterministyczny! Sprawdz, czy gdzies nie ma "
        "date.today(), time.time(), set() w kolejnosci danych albo "
        "random bez ziarna."
    )


def test_ustawienia_druku_sa_kompletne(raport_na_dysku):
    """Test z modulu 11: trzy ustawienia fitToPage MUSZA byc razem."""
    wb = load_workbook(raport_na_dysku)
    try:
        ws = wb["Sprzedaż"]

        assert ws.page_setup.orientation == "landscape"
        assert ws.page_setup.paperSize == ws.PAPERSIZE_A4

        # UWAGA: print_area zwraca KWALIFIKOWANA, ABSOLUTNA postac!
        # Nie "A1:D7", a np. "Sprzedaż!$A$1:$D$7".
        obszar = ws.print_area or ""
        assert "$A$1:$D$7" in obszar, f"Obszar druku: {obszar!r}"

        assert ws.print_title_rows == "2:2"

        # fitToPage sklada sie z TRZECH czesci - wszystkie naraz:
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True
        assert ws.page_setup.fitToWidth == 1
        assert ws.page_setup.fitToHeight == 0
    finally:
        wb.close()


def test_liczba_regul_cf_jest_poprawna(raport_na_dysku):
    """PULAPKA: len(ws.conditional_formatting) liczy ZAKRESY, nie REGULY."""
    wb = load_workbook(raport_na_dysku)
    try:
        ws = wb["Sprzedaż"]

        # Zle - to liczba ZAKRESOW (tu: 2):
        liczba_zakresow = len(ws.conditional_formatting)
        # Dobrze - to liczba REGUL:
        liczba_regul = sum(len(cf.rules) for cf in ws.conditional_formatting)

        assert liczba_zakresow == 2
        assert liczba_regul == 2

        # Wyciagnij konkretne reguly z ich zakresami i typami:
        inwentarz = sorted(
            (str(cf.sqref), r.type, r.priority)
            for cf in ws.conditional_formatting
            for r in cf.rules
        )
        assert inwentarz == [
            ("A3:D7", "expression", 2),
            ("C3:C7", "cellIs", 1),
        ], f"Nieoczekiwany inwentarz regul: {inwentarz}"
    finally:
        wb.close()


def test_tabela_ma_wlasciwy_zakres_i_styl(raport_na_dysku):
    wb = load_workbook(raport_na_dysku)
    try:
        ws = wb["Sprzedaż"]
        assert list(ws.tables.keys()) == ["SprzedazTabela"]

        tab = ws.tables["SprzedazTabela"]
        assert tab.ref == "A2:D7"                    # 5 wierszy danych + naglowek
        assert tab.tableStyleInfo.name == "TableStyleMedium2"
        assert tab.tableStyleInfo.showRowStripes is True
    finally:
        wb.close()


def test_scalenia_i_zamrozenie(raport_na_dysku):
    wb = load_workbook(raport_na_dysku)
    try:
        ws = wb["Sprzedaż"]
        scalenia = {str(r) for r in ws.merged_cells.ranges}
        assert scalenia == {"A1:D1"}, f"Scalenia: {scalenia}"

        # str() dla bezpieczenstwa - typ zwracany bywa Cell albo str
        assert str(ws.freeze_panes) == "A3"
    finally:
        wb.close()


def test_style_sa_zdeduplikowane(raport_na_dysku):
    """Excel deduplikuje style. 5 wierszy danych NIE moze dac 50 fontow.

    Ten test chroni przed pulapka wydajnosciowa z modulu 09:
    tworzeniem obiektu Font w kazdym obiegu petli.
    """
    odcisk = podsumowanie_skoroszytu(raport_na_dysku)
    style = odcisk["style"]

    assert style["fonty"] <= 5, (
        f"W pliku jest {style['fonty']} DEFINICJI fontow. "
        "To podejrzanie duzo - sprawdz, czy nie tworzysz Font() w petli."
    )
    assert style["wypelnienia"] <= 6
    assert style["formaty_liczb"] <= 5
```

**Co się dzieje w pamięci.** `zbuduj_raport()` trzyma w pamięci cały model (kilkaset komórek — nieistotne). `podsumowanie_skoroszytu(sciezka)` otwiera ZIP, czyta trzy części, parsuje `styles.xml` (mały plik), a potem wczytuje pełny model — i **zamyka go w `finally`**. `odcisk_bajtowy` **nie wczytuje modelu w ogóle** — czyta tylko części ZIP-a. Dlatego `test_raport_jest_deterministyczny` jest najszybszym testem w pliku.

**Co trafi do pliku.** Po pierwszym uruchomieniu testu `golden` utworzy `tests/golden/raport_11_odcisk.txt` i **celowo go obleje** (`pytest.fail`), żebyś go przejrzał. Po `git add` i drugim uruchomieniu — test jest zielony. Ten plik trafia do repozytorium i jest **kontraktem na wygląd raportu**.

**Co zawiera snapshot (fragment):**

```text
=== ARKUSZE ===
Sprzedaż            7x   4 komorki=27    cf=2(cellIs,expression) dv=0(-)
                     tabele=['SprzedazTabela'] scalenia=['A1:D1'] zamrozone=A3
                     druk: obszar=Sprzedaż!$A$1:$D$7 wiersze=2:2 orientacja=landscape fitToPage=True (1x0)
=== STYLE (z xl/styles.xml) ===
fonty             3
formaty_liczb     4
obramowania       2
style_komorek     6
wypelnienia       4
=== CZESCI PLIKU ===
arkusze_zip       ['xl/worksheets/sheet1.xml']
ma_shared_strings True
ma_vba            False
media             []
tabele_zip        ['xl/tables/table1.xml']
wykresy           []
```

**Trzy rzeczy do przemyślenia:**

1. **`odcisk_bajtowy` nie wczytuje modelu.** To nie jest optymalizacja — to **decyzja projektowa**. Hash jest po to, żeby porównać **treść części**; wczytanie modelu dałoby ten sam wynik za wyższą cenę i zależałoby od **wersji openpyxl**. Hash części jest odporny na wersję (poza porządkiem atrybutów XML) i dlatego nadaje się do CI na różnych maszynach.
2. **`test_style_sa_zdeduplikowane` to test wydajności przebrany za test struktury.** Nie mierzy czasu — mierzy **liczbę definicji stylów w pliku**. A to jest dokładnie ta miara, która rośnie, gdy tworzysz `Font(...)` w pętli. Ten test **nigdy nie będzie migał** (bo nie zależy od sprzętu), a złapie dokładnie tę regresję, którą moduł 09 opisywał jako „marnujesz czas i pamięć".
3. **Ten test zostanie czerwony po każdej zmianie raportu — i to jest jego wartość.** Ale musisz nauczyć się z nim pracować: **nie** nadpisujesz snapshotu, dopóki nie przeczytasz `diff` i nie uznasz, że zmiana jest **zamierzona**. Wtedy `pytest --update-golden`, `git diff tests/golden/`, i **commitujesz snapshot razem z kodem**. Recenzent widzi **kod i skutek** w jednym Pull Requeście.

### Przykład 4 — testy pułapek, `hypothesis` i budżet wydajności (🔴)

Trzy ostatnie klocki w jednym pliku.

```python
# tests/test_20_pulapki.py
"""Testy, ktore DOKUMENTUJA ograniczenia. Nie naprawiaj ich - czytaj."""

from __future__ import annotations

import io

import pytest
from openpyxl import Workbook, load_workbook


# ============================================================
# PULAPKA 1: formula nie ma wartosci (modul 06)
# ============================================================
def test_formula_zapisana_przez_openpyxl_nie_ma_wartosci():
    """DOKUMENTUJE OGRANICZENIE.

    openpyxl NIE MA silnika obliczen. Zapisany plik ma formule,
    ale bez 'cache'owanej' wartosci. Odczyt z data_only=True da None.

    Gdyby ten test kiedys przeszedl na assert == 5, znaczy to,
    ze openpyxl zaczal liczyc formuly - co sie nie zdarzy.
    Ten test chroni przed ZLUDZENIEM.
    """
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 2
    ws["A2"] = 3
    ws["A3"] = "=SUM(A1:A2)"

    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()), data_only=True)
    try:
        assert wb2.active["A3"].value is None
    finally:
        wb2.close()

    # A bez data_only widzimy LITERAL formuly:
    wb3 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        assert wb3.active["A3"].value == "=SUM(A1:A2)"
    finally:
        wb3.close()


# ============================================================
# PULAPKA 2: None to nie zero (modul 05)
# ============================================================
def test_pusta_komorka_to_none_a_zero_to_zero():
    """DOKUMENTUJE: pusta komorka != 0.

    Konsekwencja: `if kwota:` jest BLEDEM. Prawidlowo: `if kwota is not None:`.
    """
    wb = Workbook()
    ws = wb.active
    ws.append(["A", None])        # brak wartosci
    ws.append(["B", 0])           # wartosc zerowa

    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        ws2 = wb2.active
        pusta = ws2["B1"].value
        zero = ws2["B2"].value

        assert pusta is None
        assert zero == 0
        assert zero is not None

        # A teraz dowod, dlaczego to ma znaczenie:
        assert bool(pusta) is False
        assert bool(zero) is False        # ⬅ 0 tez jest falsy!
        # Dlatego `if kwota:` wyrzuca ZAROWNO braki, JAK I zera.
        # Prawidlowo:
        assert (pusta is not None) is False
        assert (zero is not None) is True
    finally:
        wb2.close()


# ============================================================
# PULAPKA 3: '=' w danych uzytkownika staje sie formula (modul 21)
# ============================================================
def test_rowna_sie_w_danych_staje_sie_formula():
    """DOKUMENTUJE RYZYKO WSTRZYKNIECIA FORMULY.

    Jesli zapiszesz tekst uzytkownika bezposrednio, a zaczyna sie od '=',
    openpyxl zapisze FORMULE. Excel ja wykona przy otwarciu.
    """
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "=1+1"                  # tak wpisze to naiwny kod

    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        assert wb2.active["A1"].data_type == "f", (
            "To jest formula! Excel ja wykona przy otwarciu pliku."
        )
        assert wb2.active["A1"].value == "=1+1"
    finally:
        wb2.close()


def test_neutralizacja_rowna_sie_przez_wymuszenie_typu_tekstowego():
    """NEUTRALIZACJA: wymus data_type = 's' przed zapisem.

    Uwaga: to zachowanie jest na styku modelu i zapisu.
    Sprawdz w dokumentacji swojej wersji openpyxl.
    """
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "=1+1"
    ws["A1"].data_type = "s"           # "s" = string

    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        return_typ = wb2.active["A1"].data_type
        return_val = wb2.active["A1"].value
        assert return_typ == "s", (
            f"Po wymuszeniu 's' typ wrocil jako {return_typ!r}. "
            "Sprawdz w dokumentacji swojej wersji openpyxl."
        )
        assert return_val == "=1+1"
    finally:
        wb2.close()


# ============================================================
# PULAPKA 4: bardzo dlugi tekst
# ============================================================
def test_bardzo_dlugi_tekst_przechodzi_przez_openpyxl():
    """DOKUMENTUJE GRANICE.

    Limit Excela na zawartosc komorki to 32 767 znakow.
    openpyxl sam tego limitu NIE egzekwuje - wiec plik z dluzszym
    tekstem powstanie, ale Excel moze go odrzucic przy otwarciu.

    Ten test pilnuje, zebysmy nie wprowadzali takich danych
    "przypadkiem" - jesli test padnie, ktos zmienil limit.
    """
    dlugi = "x" * 40_000               # powyzej limitu Excela!

    wb = Workbook()
    ws = wb.active
    ws["A1"] = dlugi

    bufor = io.BytesIO()
    wb.save(bufor)

    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        assert wb2.active["A1"].value == dlugi
    finally:
        wb2.close()

    # Wniosek: openpyxl zapisal, ale Excel moze odmowic.
    # Dla dlugich tekstow: podziel na wiele komorek albo skroc.
```

I testy właściwościowe:

```python
# tests/test_20_wlasciwosci.py
"""Testy wlasciwosciowe dla funkcji czyszczacej dane."""

from __future__ import annotations

import pytest
from hypothesis import given, settings, strategies as st


def czysc_kwote(wartosc):
    """Zamienia DOWOLNY zapis kwoty na float albo None.

    Obsluguje: None, liczby, "1 200,50", "1.234,56", "1200.50 zl", smieci.
    """
    if wartosc is None:
        return None
    if isinstance(wartosc, bool):      # bool jest podtypem int - lapiemy osobno
        return None
    if isinstance(wartosc, (int, float)):
        return float(wartosc)

    tekst = str(wartosc).strip()
    for smiec in ("zł", "PLN", "EUR", "\u00a0"):
        tekst = tekst.replace(smiec, "")
    tekst = tekst.replace(" ", "").strip()
    if not tekst:
        return None

    przecinek = tekst.count(",")
    kropka = tekst.count(".")
    if przecinek == 1 and kropka == 0:
        tekst = tekst.replace(",", ".")
    elif przecinek > 0 and kropka > 0:
        tekst = tekst.replace(".", "").replace(",", ".")

    try:
        return float(tekst)
    except (ValueError, TypeError):
        return None


# --- Wlasciwosc 1: niezmiennik typu ----------------------------------------
@given(st.one_of(
    st.none(),
    st.booleans(),
    st.integers(),
    st.floats(allow_nan=False, allow_infinity=False),
    st.text(max_size=20),
))
@settings(max_examples=300)
def test_wynik_jest_floatem_albo_none(wartosc):
    """NIEZMIENNIK: 'zadna wartosc nie jest tekstem, jesli miala byc liczba'."""
    wynik = czysc_kwote(wartosc)
    assert wynik is None or isinstance(wynik, float), (
        f"czysc_kwote({wartosc!r}) zwrocilo {wynik!r} typu {type(wynik).__name__}"
    )


# --- Wlasciwosc 2: idempotencja --------------------------------------------
@given(st.text(alphabet="0123456789., ", max_size=12))
@settings(max_examples=300)
def test_czyszczenie_jest_idempotentne(tekst):
    """NIEZMIENNIK: czyszczenie dwa razy = czyszczenie raz."""
    raz = czysc_kwote(tekst)
    assert czysc_kwote(raz) == raz, (
        f"czysc_kwote({tekst!r}) = {raz!r}, "
        f"ale czysc_kwote({raz!r}) = {czysc_kwote(raz)!r}"
    )


# --- Wlasciwosc 3: formaty europejskie -------------------------------------
@given(
    st.integers(min_value=0, max_value=999_999),
    st.integers(min_value=0, max_value=99),
)
@settings(max_examples=200)
def test_format_europejski_jest_rozumiany(czesc_calkowita, grosze):
    """NIEZMIENNIK: "1.234,56" == 1234.56, niezaleznie od wartosci."""
    tekst = f"{czesc_calkowita:,}".replace(",", ".") + f",{grosze:02d}"
    wynik = czysc_kwote(tekst)
    assert wynik is not None, f"Nie udalo sie odczytac {tekst!r}"
    assert wynik == pytest.approx(czesc_calkowita + grosze / 100, abs=0.005)
```

I budżet wydajności z fixture'em sterującym markerami:

```toml
# pyproject.toml (fragment)
[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-ra --strict-markers"
markers = [
    "slow: testy trwajace dluzej niz sekunde (pomin lokalnie: -m 'not slow')",
    "performance: testy budzetu wydajnosci",
]
```

```python
# tests/test_20_budzet.py
"""Budzet wydajnosci jako test - zabezpieczenie przed cicha regresja."""

from __future__ import annotations

import time

import pytest
from openpyxl import Workbook, load_workbook

import openpyxl


@pytest.mark.slow
@pytest.mark.performance
def test_budzet_zapisu_100k_wierszy(tmp_path):
    """BUDZET: 100 000 wierszy w write_only < 10 s.

    Jesli pada, sprawdz w tej kolejnosci:
      1. czy lxml jest aktywny,
      2. czy nie tworzysz obiektow Font/Fill w petli,
      3. czy naprawde uzywasz write_only (a nie trybu normalnego).
    """
    cel = tmp_path / "duzo.xlsx"

    t0 = time.perf_counter()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Dane")
    ws.append(["ID", "Kwota"])
    for i in range(1, 100_001):
        ws.append([i, i * 0.5])
    wb.save(cel)
    wb.close()
    czas = time.perf_counter() - t0

    assert czas < 10.0, (
        f"Zapis 100k wierszy: {czas:.2f} s (budzet 10 s). "
        f"lxml aktywny: {openpyxl.LXML}"
    )


@pytest.mark.slow
@pytest.mark.performance
def test_budzet_odczytu_100k_wierszy_read_only(tmp_path):
    """BUDZET: odczyt 100 000 wierszy w read_only < 8 s."""
    cel = tmp_path / "duzo.xlsx"

    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title="Dane")
    for i in range(1, 100_001):
        ws.append([i, i * 0.5])
    wb.save(cel)
    wb.close()

    t0 = time.perf_counter()
    wb2 = load_workbook(cel, read_only=True)
    try:
        suma = 0.0
        for wiersz in wb2["Dane"].iter_rows(values_only=True):
            suma += wiersz[1] or 0.0
    finally:
        wb2.close()
    czas = time.perf_counter() - t0

    assert czas < 8.0, f"Odczyt 100k wierszy: {czas:.2f} s (budzet 8 s)"
```

**Co się dzieje w pamięci.** Testy pułapek działają na skoroszytach o jednej–dwóch komórkach — koszt zerowy. Testy właściwościowe generują 300 przypadków, ale **każdy jest stringiem lub liczbą** — bez openpyxl, więc koszt to mikrosekundy. Testy budżetu celowo tworzą duże pliki — dlatego są oznaczone `slow` i `performance`: **dokładają sekundę do zestawu, więc muszą być wyłączalne**.

**Co trafi do pliku.** W testach pułapek i właściwości — nic (wszystko w `BytesIO`). W testach budżetu — plik w `tmp_path`, który `pytest` **sam posprząta** po zakończeniu sesji.

**Trzy rzeczy do przemyślenia:**

1. **`test_budzet_*` nie jest testem wydajności — jest testem REGRESJI wydajności.** Nie interesuje go, czy jesteś szybki, ale czy **nie stałeś się nagle trzy razy wolniejszy**. Dlatego budżet ustawia się z zapasem. Test, który działa na granicy, w CI będzie migał — a test, który miga, jest bezużyteczny.
2. **`st.text(alphabet="0123456789., ", max_size=12)`** — to jest **celowe zawężenie strategii**. Gdybyśmy dali `st.text()` bez argumentów, `hypothesis` generowałaby chińskie znaki, emoji i znaki sterujące — i test padłby na czymś, co prawdopodobnie nigdy nie przyjdzie z Excela. **Test właściwościowy też ma zakres.** Świadome zawężenie do „alfabetu, który realnie widzisz w komórkach" jest ważniejsze niż „testowanie wszystkiego".
3. **Test pułapki 4 z 40 000 znaków ma sens tylko, jeśli ktoś go przeczyta.** Więc jego docblock musi być napisany tak, żeby **przekazywał wiedzę**. Zobacz, jak zaczyna się każdy z tych testów: `"""DOKUMENTUJE OGRANICZENIE."""`. Trzy słowa, które mówią następnemu programiście: „nie naprawiaj tego, to jest zamierzone".

## 4. Anatomia API

| Klasa / funkcja | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `io.BytesIO()` | Bufor binarny w pamięci (plikopodobny) | — | `getvalue()` ignoruje kursor, `read()` nie |
| `bufor.getvalue()` | Zwraca całe bajty bufora | — | Bezpieczne niezależnie od pozycji kursora |
| `bufor.seek(0)` | Przewija kursor na początek | `offset` | **Wymagane** przed `read()` i przed strumieniowaniem |
| `wb.save(obiekt)` | Zapis do pliku lub bufora | ścieżka **albo** obiekt z `write()` | Zamyka ZIP; nadpisuje istniejące pliki bez ostrzeżenia |
| `load_workbook(BytesIO(bajty))` | Wczytanie z pamięci | — | Pamiętaj o `wb.close()` |
| `wb.close()` | Zwalnia uchwyt ZIP | — | Obowiązkowe po `load_workbook` |
| `pytest.fixture` | Definiuje stan przygotowany przed testem | — | **Nigdy nie zawiera asercji** |
| Fabryka jako fixture | Fixture zwracający funkcję budującą | — | Zwracaj **bajty**, nie obiekt `Workbook` |
| `tmp_path` | Unikalny katalog tymczasowy na test | — | Tylko tam, gdzie kod **wymaga** ścieżki |
| `pytest.mark.parametrize` | Jeden test, wiele wariantów | `argnames`, `argvalues`, `ids` | `ids=lambda f: f.__name__` daje czytelne nazwy |
| `pytest.raises(Wyjątek, match="...")` | Asercja na wyjątek | `match` = regex na komunikat | Testuj też **komunikat**, nie tylko typ |
| `pytest.approx(x, abs=...)` | Porównanie liczb zmiennoprzecinkowych | `rel`, `abs` | **Zawsze** dla `float` |
| `wb.sheetnames` | Lista nazw arkuszy (w kolejności) | — | Lista, nie zbiór — kolejność ma znaczenie |
| `ws.max_row`, `ws.max_column` | Wymiary | — | Bywają zawyżone przez formatowanie (moduł 07) |
| `cell.data_type` | Typ komórki | — | `n` liczba, `s` tekst, `d` data, `f` formuła, `b` bool, `e` błąd |
| `cell.number_format` | Kod formatu liczby | — | Testuj **kod**, nie renderowanie (locale!) |
| `cell.font` / `.fill` / `.border` / `.alignment` | Styl komórki | — | Dostępne tylko w trybie normalnym |
| `ws.tables` | Dict-like tabel arkusza | — | `.keys()`, `.values()`, `.items()` → `[(nazwa, ref)]` |
| `Table.ref` | Zakres tabeli | — | Zawiera nagłówek |
| `Table.tableStyleInfo.name` | Nazwa stylu tabeli | — | np. `"TableStyleMedium9"` |
| `ws.conditional_formatting` | Iterowalny po **zakresach** | — | Każdy element ma `.sqref` i `.rules` |
| `cf.rules` | Lista reguł w zakresie | — | **Prawidłowa liczba reguł** = `sum(len(cf.rules) ...)` |
| `rule.type`, `.priority`, `.operator`, `.formula`, `.stopIfTrue` | Cechy reguły | — | `"expression"` = `FormulaRule` |
| `ws.data_validations.dataValidation` | Lista walidacji | — | `dv.type`, `dv.sqref`, `dv.formula1` |
| `ws.merged_cells.ranges` | Zbiór scalonych zakresów | — | `{str(r) for r in ...}` |
| `ws.freeze_panes` | Zamrożone okna | — | Ustawiane stringiem; czytaj `str(...)` |
| `ws.auto_filter.ref` | Zakres filtra automatycznego | — | String, np. `"A1:H100"` |
| `ws.print_area` | Obszar druku | — | **Zwraca kwalifikowaną, absolutną postać!** |
| `ws.print_title_rows` | Powtarzane wiersze druku | — | Format `"1:1"`; `None` gdy nieustawione |
| `ws.page_setup.orientation` | Orientacja | `"portrait"`/`"landscape"` | |
| `ws.page_setup.fitToWidth/fitToHeight` | Dopasowanie do stron | — | Wymaga `pageSetUpPr.fitToPage = True` |
| `ws.sheet_properties.pageSetUpPr.fitToPage` | Włącznik fitToPage | — | **Trzy ustawienia razem** (moduł 11) |
| `ws.column_dimensions[k].width` | Szerokość kolumny | — | Testuj, że jest **ustawiona** |
| `wb.properties.creator/title/created` | Właściwości dokumentu | — | Ślad audytowy; `modified` jest nadpisywane |
| `zipfile.ZipFile.namelist()` | Lista części archiwum | — | Podstawa inwentarza ZIP (moduł 02) |
| `zipfile.ZipFile.read(nazwa)` | Treść części | — | Do hashowania i parsowania |
| `ET.canonicalize(xml_data=...)` | Kanoniczna postać XML | Python 3.8+ | Normalizacja przed porównaniem |
| `hashlib.sha256()` | Hash części | — | Hashuj **rozpakowane** części, nie plik |
| `hypothesis.given` / `strategies` | Generowanie danych wejściowych | `st.integers`, `st.text`, `st.floats` | `st.floats(allow_nan=False)` |
| `hypothesis.settings` | Konfiguracja (`max_examples`, `deadline`) | — | `deadline=None` dla wolnych testów |
| `pytest_addoption` + `--update-golden` | Świadoma aktualizacja wzorca | — | Osobna decyzja, nie odruch |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — test w `BytesIO` na własnym skoroszycie.**

Napisz `tests/test_20_cw1.py`, który zawiera **jeden test** i **jedną funkcję pomocniczą**.

Wymagania:

1. Funkcja `utworz_skoroszyt() -> bytes`:
   - tworzy `Workbook`,
   - nadaje arkuszowi nazwę `"Test"`,
   - wpisuje `A1 = "Nazwa"`, `B1 = "Wartość"`,
   - wpisuje `A2 = "Alfa"`, `B2 = 12.5`,
   - zapisuje do `io.BytesIO()`, **zwraca `getvalue()`**.
2. Test `test_b1_zawiera_wartosc_A1()`:
   - wczytuje bajty przez `load_workbook(io.BytesIO(...))`,
   - w `try/finally` zamyka skoroszyt,
   - sprawdza `ws["A1"].value == "Nazwa"`.
3. Test `test_typ_wartosci()`:
   - sprawdza `isinstance(ws["B2"].value, float)`,
   - sprawdza `ws["B2"].value == 12.5`.
4. Uruchom `pytest tests/test_20_cw1.py -v`.
5. **Dodatkowo** — w tym samym pliku, w osobnym teście `test_kursor_bufora()`, udowodnij eksperymentalnie różnicę między `getvalue()` a `read()`:
   - zapisz skoroszyt do bufora,
   - `assert bufor.read() == b""` (bo kursor na końcu),
   - `bufor.seek(0)`,
   - `assert len(bufor.read()) > 0`.

W pliku `output/20_cw1_wnioski.md` odpowiedz:

1. **Ile bajtów ma Twój skoroszyt?** Wypisz `len(bajty)` i porównaj z rozmiarem pliku, który powstałby na dysku (zapisz raz do `tmp_path` i sprawdź). Czy zgadzają się co do bajtu? Dlaczego nie muszą?
2. **Co by się stało, gdybyś zapomniał `wb.close()`?** Wyjaśnij, odwołując się do tego, co `load_workbook` trzyma otwarte (moduł 19).
3. **Dlaczego fabryka zwraca `bytes`, a nie `io.BytesIO`?** Odpowiedz w dwóch zdaniach, używając słowa „mutowalny".
4. **Ile trwał cały zestaw testów?** Uruchom `pytest --durations=5` i wypisz wynik. Porównaj z hipotetycznym testem, który zapisuje i wczytuje **prawdziwe pliki** z dysku.
5. **Co by było, gdyby `wb.save(bufor)` nie zamykał archiwum ZIP?** Wyjaśnij, odwołując się do „katalogu centralnego" z modułu 02.

### 🟡 Warsztat

**Zadanie 2 — fixture, fabryka i pięć testów.**

Napisz `conftest.py` i `tests/test_20_cw2.py` dla raportu **faktur**.

Wymagania dla `conftest.py`:

1. Fixture `zbuduj_skoroszyt` — fabryka zwracająca `bytes`, z hakiem `ustaw` (jak w Przykładzie 2).
2. Stała `DOMYSLNE_WIERSZE` z nagłówkiem i **4 wierszami** faktur: `(numer, data, kontrahent, netto, vat, brutto)`.
3. Fixture `wczytaj_faktury` — **context manager** (przez `contextlib.contextmanager`), który wczytuje bajty i **zawsze** woła `close()`.

Wymagania dla testów (**dokładnie pięć**):

1. `test_arkusz_ma_wlasciwa_nazwe_i_jeden_arkusz` — sprawdza `wb.sheetnames`.
2. `test_naglowki_sa_w_pierwszym_wierszu` — sprawdza **wszystkie sześć** nagłówków w kolejności.
3. `test_kwoty_netto_sa_floatami` — iteruje po kolumnie netto (wiersze 2+) i sprawdza `isinstance(..., float)` dla każdej.
4. `test_brak_pustych_komorek_w_obliczeniach` — sprawdza, że w kolumnach netto, vat i brutto **nie ma `None`**, i wypisuje **do 10 adresów** braków w komunikacie.
5. `test_brutto_zgadza_sie_z_netto_plus_vat` — sprawdza arytmetykę `brutto == pytest.approx(netto + vat)` dla każdego wiersza.

W pliku `output/20_cw2_wnioski.md` odpowiedz:

1. **Dlaczego w komunikacie testu 4 wypisujesz tylko 10 pierwszych adresów?** Wyjaśnij, jak wyglądałby test bez tego ograniczenia przy 400 pustych komórkach.
2. **Dlaczego użyłeś `pytest.approx`, a nie `==`, w teście 5?** Co konkretnie mogłoby pójść nie tak bez tego?
3. **Co się stanie, gdy jeden test w `wczytaj_faktury` rzuci wyjątek?** Śledź kod i odpowiedz: czy `close()` zostanie wywołane? Dlaczego to działa?
4. **Dodaj test, który pada, i zobacz komunikat.** Napisz (następnie usuń) asercję `assert ws["A1"].value == "BŁĄD"`. Wklej dokładny komunikat `pytest` i wskaż, który fragment mówi Ci, **gdzie** szukać.
5. **Czy pięciu testów to dużo czy mało dla tego raportu?** Wymień **trzy** rzeczy, które warto jeszcze przetestować, i uzasadnij każdą jednym zdaniem (podpowiedź: 2.3).

### 🔴 Wyzwanie

**Zadanie 3 — `podsumowanie_skoroszytu()` i test regresji dla modułu 11.**

Napisz `tests/odcisk.py` (dokładnie tak, jak w Przykładzie 3) i `tests/test_20_cw3.py` z golden testem **dla własnego raportu**.

Wymagania dla `odcisk.py`:

1. `podsumowanie_skoroszytu(sciezka) -> dict` z **trzema warstwami**:
   - inwentarz części ZIP (`xl/media/`, `xl/charts/chart*`, `xl/tables/`, `xl/worksheets/sheet*`, `sharedStrings`, `vbaProject`),
   - model skoroszytu (arkusze, wymiary, komórki z danymi, reguły CF, walidacje, tabele, scalenia, freeze, obszar druku, powtarzane wiersze, orientacja, `fitToPage`),
   - stylistyka parsowana z `xl/styles.xml` (fonty, wypełnienia, obramowania, formaty liczb, `cellXfs`).
2. `odcisk_tekstowy(podsumowanie) -> str` — czytelny tekst, **jedna linia na właściwość**.
3. `odcisk_bajtowy(sciezka) -> str` — SHA-256 z **rozpakowanych części**, z **wykluczeniem** `docProps/core.xml`, `docProps/app.xml`, `xl/calcChain.xml`.
4. Stała `CZESCI_ZMIENNE` z komentarzem **wyjaśniającym, dlaczego** te części są wykluczone.

Wymagania dla `tests/test_20_cw3.py`:

1. Fixture `raport_na_dysku(tmp_path)` — buduje raport i **zwraca ścieżkę** (jedno z niewielu miejsc, gdzie dotykamy dysku; uzasadnij to komentarzem!).
2. Fixture `golden` z `--update-golden` (Przepisz go lub zaimportuj z `conftest.py`).
3. Test `test_odcisk_zgadza_sie_z_wzorcem`.
4. Test `test_raport_jest_deterministyczny` — buduje raport **dwa razy**, porównuje `odcisk_bajtowy`.
5. Test `test_liczba_regul_cf_jest_poprawna` — **z komentarzem ostrzegawczym** o pułapce `len(ws.conditional_formatting)`.
6. Test `test_ustawienia_druku_sa_kompletne` — sprawdza **cztery** rzeczy: orientację, obszar druku (przez `in`, nie `==`!), `print_title_rows` i **trzyczęściowy** `fitToPage`.
7. Test `test_style_sa_zdeduplikowane` — asercja na **liczbę definicji fontów** w `styles.xml`.
8. Raport musi być **deterministyczny**: `dzien` jako parametr, żadnego `date.today()`, żadnego `set()` w kolejności danych, żadnego `random` bez ziarna.

Wymagania eksperymentalne (i odpowiedzi w `output/20_cw3_wnioski.md`):

1. **Uruchom test regresji dwa razy.** Pierwszy raz — celowo czerwony (brak wzorca). Skopiuj komunikat, wykonaj `git add tests/golden/`, uruchom ponownie. Wklej oba wyniki.
2. **Zmień format kwoty** z `#,##0.00` na `#,##0` w raporcie. Uruchom `pytest -v`. **Wklej pełny `diff`**, który wyświetli fixture `golden`. Odpowiedz: **ile linii** snapshotu się zmieniło? Czy `diff` wskazuje **adres komórki**? Wykonaj `--update-golden`, potem `git diff tests/golden/`.
3. **Usuń jedną regułę CF** i sprawdź, **ile** asercji padnie. Wyjaśnij: dlaczego odcisk i `test_liczba_regul_cf` padają **oba**, skoro mierzą to samo? (Odpowiedź: bo mierzą **na dwóch poziomach** — struktura i model. Zaskoczony? To jest „obrona w głąb".)
4. **Dodaj `date.today()`** zamiast `dzien`. Uruchom `test_raport_jest_deterministyczny` **trzy razy**. Czy pada za każdym razem? Wyjaśnij, dlaczego może **przejść raz, a paść raz** (podpowiedź: rozdzielczość 2 sekund w znacznikach ZIP i to, gdzie `date.today()` trafia — do XML, nie do znaczników!).
5. **Porównaj czas** `odcisk_bajtowy` z `podsumowanie_skoroszytu` na tym samym pliku (`time.perf_counter` ×100). Wyjaśnij różnicę: co robi jedno, czego nie robi drugie?
6. **Dodaj obraz `xl/media/`** (moduł 16) do raportu. Czy `test_style_sa_zdeduplikowane` go widzi? A `test_odcisk_zgadza_sie_z_wzorcem`? Wyjaśnij, gdzie obraz pojawia się w odcisku i dlaczego `styles.xml` go nie widzi.

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe fragmenty i uzasadnienia</strong></summary>

```python
"""Odcisk skoroszytu - kluczowe fragmenty i uzasadnienia."""

from __future__ import annotations

import hashlib
import zipfile
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# UZASADNIENIE WYKLUCZENIA:
#   docProps/core.xml  -> openpyxl wpisuje dcterms:modified = biezacy czas
#                          przy KAZDYM save(). Nie da sie tego przypiac
#                          przez wb.properties.modified (jest nadpisywane).
#   docProps/app.xml   -> wersja aplikacji; zmienia sie miedzy wersjami.
#   xl/calcChain.xml   -> kolejnosc przeliczania formul; bywa odbudowywana.
CZESCI_ZMIENNE = frozenset({
    "docProps/core.xml",
    "docProps/app.xml",
    "xl/calcChain.xml",
})


def _inwentarz(zf: zipfile.ZipFile) -> dict:
    nazwy = zf.namelist()
    return {
        "media": sorted(n for n in nazwy if n.startswith("xl/media/")),
        "wykresy": sorted(n for n in nazwy if n.startswith("xl/charts/chart")),
        "tabele_zip": sorted(n for n in nazwy if n.startswith("xl/tables/")),
        "arkusze_zip": sorted(n for n in nazwy if n.startswith("xl/worksheets/sheet")),
        "ma_shared_strings": "xl/sharedStrings.xml" in nazwy,
        "ma_vba": "xl/vbaProject.bin" in nazwy,
    }


def _stylistyka(zf: zipfile.ZipFile) -> dict:
    """Parsujemy XML, a NIE prywatne wb._fonts.

    Powod: prywatne atrybuty moga zniknac miedzy wersjami openpyxl,
    a format XLSX jest stabilny. Dodatkowo styles.xml mowi nam,
    ile DEFINICJI stylow jest w pliku (Excel je deduplikuje!).
    """
    try:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    except KeyError:
        return {}

    def ile(tag: str) -> int:
        el = root.find(f"{NS}{tag}")
        return len(el) if el is not None else 0

    return {
        "fonty": ile("fonts"),
        "wypelnienia": ile("fills"),
        "obramowania": ile("borders"),
        "formaty_liczb": ile("numFmts"),
        "style_komorek": ile("cellXfs"),
    }


def _arkusz(ws) -> dict:
    """Jedna linia odcisku na arkusz."""
    reguly = [
        (str(cf.sqref), r.type, r.priority)
        for cf in ws.conditional_formatting
        for r in cf.rules
    ]
    return {
        "nazwa": ws.title,
        "wymiary": (ws.max_row, ws.max_column),
        # PULAPKA: len(ws.conditional_formatting) == liczba ZAKRESOW.
        # Liczba REGUL to sum() po cf.rules dla wszystkich cf.
        "reguly_cf": len(reguly),
        "typy_regul_cf": sorted({t for _, t, _ in reguly}),
        "komorki_z_danymi": sum(
            1 for r in ws.iter_rows() for c in r if c.value is not None
        ),
        "tabele": sorted(ws.tables.keys()),
        "scalenia": sorted(str(r) for r in ws.merged_cells.ranges),
        "zamrozone": str(ws.freeze_panes) if ws.freeze_panes else None,
        # PULAPKA: print_area zwraca KWALIFIKOWANA, ABSOLUTNA postac,
        # np. "Sprzedaż!$A$1:$D$7" - nie "A1:D7".
        "obszar_druku": ws.print_area or None,
        "powtarzane_wiersze": ws.print_title_rows,
        "orientacja": ws.page_setup.orientation,
        "fitToPage": bool(
            ws.sheet_properties.pageSetUpPr
            and ws.sheet_properties.pageSetUpPr.fitToPage
        ),
    }


def podsumowanie_skoroszytu(sciezka: str | Path) -> dict:
    sciezka = Path(sciezka)
    with zipfile.ZipFile(sciezka) as zf:
        inwentarz = _inwentarz(zf)
        stylistyka = _stylistyka(zf)

    wb = load_workbook(sciezka)
    try:
        arkusze = [_arkusz(ws) for ws in wb.worksheets]
    finally:
        wb.close()

    return {"arkusze": arkusze, "style": stylistyka, "czesc": inwentarz}


def odcisk_tekstowy(pod: dict) -> str:
    linie = ["=== ARKUSZE ==="]
    for a in pod["arkusze"]:
        linie.append(
            f"{a['nazwa']} | {a['wymiary']} | komorki={a['komorki_z_danymi']} | "
            f"cf={a['reguly_cf']} {a['typy_regul_cf']} | tabele={a['tabele']} | "
            f"scalenia={a['scalenia']} | freeze={a['zamrozone']} | "
            f"druk={a['obszar_druku']} | tytuly={a['powtarzane_wiersze']} | "
            f"{a['orientacja']} | fitToPage={a['fitToPage']}"
        )
    linie.append("=== STYLE ===")
    for k, v in sorted(pod["style"].items()):
        linie.append(f"{k} | {v}")
    linie.append("=== CZESCI ===")
    for k, v in sorted(pod["czesc"].items()):
        linie.append(f"{k} | {v}")
    return "\n".join(linie) + "\n"


def odcisk_bajtowy(sciezka: str | Path) -> str:
    """SHA-256 TRESCI rozpakowanych czesci.

    BEZ wczytywania modelu (szybko) i BEZ kompresji (odporne na zlib).
    To jest wersja do CI i do porownan miedzy maszynami.
    """
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in sorted(zf.namelist()):
            if nazwa in CZESCI_ZMIENNE:
                continue
            h.update(nazwa.encode("utf-8"))
            h.update(b"\x00")
            h.update(zf.read(nazwa))
            h.update(b"\x00")
    return h.hexdigest()
```

**Kluczowe decyzje projektowe:**

- **Trzy warstwy, nie jedna.** Inwentarz ZIP mówi „co jest w pliku" (obrazy, wykresy, tabele). Model mówi „co to znaczy". Stylistyka mówi „jak jest narysowane i ile definicji powstało". Gdybyśmy mieli tylko jedną warstwę, odcisk byłby albo niepełny, albo zależny od implementacji.

- **`_stylistyka` parsuje XML, nie `wb._fonts`.** Trzy powody: (a) `_fonts` jest **prywatne** i może zniknąć; (b) format XLSX jest **stabilny** — to standard; (c) `styles.xml` mówi prawdę o **pliku**, a nie o modelu. To jest wprost zastosowanie modułu 02: **plik to archiwum XML, i tak też można go badać**.

- **`odcisk_bajtowy` nie wczytuje modelu.** Nie z lenistwa — z **izolacji**. Hash części nie zależy od wersji openpyxl (poza porządkiem atrybutów XML, który można znormalizować przez `ET.canonicalize`). Hash modelu zależałby od tego, jak openpyxl akurat reprezentuje `CellRange`. W CI, gdzie wersja biblioteki bywa inna niż lokalnie, ta różnica jest kluczowa.

- **`CZESCI_ZMIENNE` jako `frozenset` z komentarzem przy każdym wpisie.** Bez wyjaśnienia **dlaczego** — wykluczenie wygląda jak obejście, a nie decyzja. Za rok ktoś doda tam kolejną część „bo test miga", a rok później odcisk przestanie cokolwiek wykrywać. **Komentarz jest tu równie ważny jak kod.**

- **`_arkusz` liczy `reguly_cf` przez `sum(len(cf.rules))`, nie przez `len()`.** To jest ta pułapka z 2.3.1. Wersja przez `len()` dałaby **liczbę zakresów** — czyli odcisk mówiłby „2 reguły", a raport miałby 5. Test przeszedłby, a snapshot kłamałby.

- **`obszar_druku` zapisywany jako `ws.print_area or None`.** Getter zwraca `""` dla pustego obszaru, a `""` w odcisku jest mylące. Normalizacja do `None` czyni odcisk **jednoznacznym**: albo obszar jest, albo go nie ma.

- **Czego tu świadomie nie ma:** porównania `wb.properties.modified` (jest nadpisywane przy `save()` — pułapka!) i wyliczania sparkline'ów (openpyxl ich nie modeluje — moduł 18; można by je policzyć z `xl/worksheets/sheetN.xml` przez `extLst`, ale to materiał na osobne zadanie i wymaga ostrożności). **Odcisk musi być odporny na to, co zmienne, i czuły na to, co istotne.** Dodanie pola, które zależy od zegara, zniszczyłoby cały mechanizm.

</details>

## 6. Typowe błędy i pułapki

**1. „Test porównujący pliki bajt w bajt zawsze jest czerwony" (objaw) → openpyxl wpisuje do `docProps/core.xml` czas bieżący przy **każdym** `save()`, a każdy wpis ZIP dostaje znacznik czasu. Do tego kompresja DEFLATE może dać inne bajty na innych wersjach `zlib` (2.6, strategia 1) (przyczyna) → **Porównuj hash rozpakowanych części** (`odcisk_bajtowy` z 2.5) albo snapshot tekstowy. Nigdy bajty pliku (naprawa).**

**2. „Zwracam plik przez HTTP i wychodzi pusty, choć ma poprawny `Content-Length`" (objaw) → **brak `bufor.seek(0)` przed przekazaniem bufora.** Po `wb.save(bufor)` kursor stoi na **końcu** bufora, więc `read()` nic nie zwraca (2.2, szczegół 1) (przyczyna) → **`bufor.seek(0)` przed `send_file`/`StreamingResponse`.** W testach używaj `getvalue()`, które jest odporne na pozycję kursora (naprawa).**

**3. „`ValueError: I/O operation on closed file' przy `bufor.getvalue()`" (objaw) → **użyto `with io.BytesIO() as bufor:` i próbowano sięgnąć po `getvalue()` po wyjściu z bloku** (2.2, szczegół 3) (przyczyna) → **Nie używaj `with` z `BytesIO` w testach.** Jeśli chcesz zwolnić pamięć, rób `getvalue()` **wewnątrz** bloku (naprawa).**

**4. „`BadZipFile` przy `load_workbook(io.BytesIO(bajty))`" (objaw) → **ZIP nie został domknięty.** Dzieje się tak, gdy budujesz archiwum ręcznie (`ZipFile` + `ExcelWriter`) i nie zamkniesz `ZipFile`. Katalog centralny ZIP-a leży **na końcu pliku** (moduł 02) — bez niego archiwum jest niekompletne (2.2, szczegół 2) (przyczyna) → **`wb.save()` zamyka archiwum sam — użyj go.** Jeśli budujesz ręcznie, zamknij `ZipFile` **przed** odczytem (naprawa).**

**5. „`AssertionError: 2 != 5' na liczbie reguł formatowania warunkowego" (objaw) → **`len(ws.conditional_formatting)` zwraca liczbę ZAKRESÓW, nie reguł.** `_cf_rules` to słownik `{zakres: [reguły]}` (2.3.1, pułapka 1) (przyczyna) → **`sum(len(cf.rules) for cf in ws.conditional_formatting)`** (naprawa).**

**6. „`AssertionError: \"Sprzedaż!$A$1:$D$7\" != \"A1:D7\"'" (objaw) → **`ws.print_area` zwraca pełną, kwalifikowaną, absolutną postać** — getter składa ją z `quote_sheetname` i `absolute_coordinate` (2.3.1, pułapka 2) (przyczyna) → **Testuj `assert \"$A$1:$D$7\" in (ws.print_area or \"\")`**, a nie równość (naprawa).**

**7. „Test przechodzi sam, ale pada w zestawie z innymi" (objaw) → **współdzielony, mutowalny stan.** Najczęściej fixture zwracający **obiekt** `Workbook`, który jeden test modyfikuje, a drugi potem widzi zmieniony (2.4.1) (przyczyna) → **Fixture zwracająca `bytes`** (niemutowalne). Każdy test dostaje **świeżą** kopię. Nigdy nie zwracaj `wb` ani `ws` z fixture'u (naprawa).**

**8. „`fixture 'zbuduj_skoroszyt' not found' mimo że jest w `conftest.py`" (objaw) → **`conftest.py` leży poza drzewem katalogów testu** — `pytest` szuka go **w katalogu testu i wyżej**, w stronę `rootdir` (2.4) (przyczyna) → **Umieść `conftest.py` w `tests/` (podzbiór testów) albo w katalogu głównym (wszystkie testy).** Sprawdź `pytest --fixtures` (naprawa).**

**9. „Fixture padła i nie wiem, czy to błąd w kodzie, czy w fixture" (objaw) → **fixture zawiera asercje.** Wtedy błąd „przygotowania" wygląda jak błąd „kodu pod testem" (2.4.2) (przyczyna) → **Fixture PRZYGOTOWUJE, test SPRAWDZA.** Żadnych `assert` w fixture'ach (naprawa).**

**10. „Test działa u mnie 3.12, pada w CI na 3.9" (objaw) → **składnia typów `X | Y` albo `list[str]` bez `from __future__ import annotations`** (2.11) (przyczyna) → **Dodaj `from __future__ import annotations` na początku KAŻDEGO pliku** albo używaj `Optional[X]`, `List[X]`. Macierz Pythona w CI wyłapie to przed Tobą (naprawa).**

**11. „Test przechodzi lokalnie, ale pada w CI na porównaniu kwoty → tekst renderowany zależnie od locale" (objaw) → **CI ma `C.UTF-8` (albo inny locale), a Twój laptop `pl_PL`.** Test porównywał **wyrenderowany** string zamiast **kodu formatu** (2.11, 2.12) (przyczyna) → **Testuj `cell.number_format` (`\"#,##0.00\"`), nie jego wynik.** Kod formatu jest identyczny w każdym locale (naprawa).**

**12. „`AssertionError` w teście, który przechodził wczoraj, a kod się nie zmienił" (objaw) → **test zależny od zegara** (`date.today()`, `datetime.now()`), losowości albo kolejności zbioru (`set()`) (2.10, moduł 19) (przyczyna) → **Wstrzyknij czas jako parametr** (`def zbuduj(dzien: date)`). Sortuj dane **stabilnie**. Nigdy nie iteruj po `set()` tam, gdzie kolejność trafia do pliku (naprawa).**

**13. „`--update-golden` i już — test zielony" (objaw) → **nadpisałeś wzorzec bez przeczytania `diff`.** Wtedy golden test przestaje cokolwiek sprawdzać (2.6.1) (przyczyna) → **Kolejność: pad → przeczytaj `diff` → oceń, czy zmiana zamierzona → `--update-golden` → `git diff tests/golden/` → commituj snapshot RAZEM z kodem.** Recenzent musi widzieć **skutek** zmiany (naprawa).**

**14. „Pokrycie 95%, a produkcja się sypie" (objaw) → **pokrycie mierzy, ile linii WYKONANO — nie ile ZWERYFIKOWANO.** Test, który wykonuje kod bez ani jednej asercji, daje 100% pokrycia i zero pewności (2.11) (przyczyna) → **Pokrycie jako `--cov-fail-under` (bramka na minimum), nie jako cel.** Patrz na `--cov-report=term-missing` — **czerwone linie są cenniejsze niż procent** (naprawa).**

**15. „Zaktualizowałem openpyxl i `mypy` zaczął pluć 200 błędami" (objaw) → **stuby `types-openpyxl` opisują KONKRETNĄ wersję API.** Niezmieniona wersja stubów wobec nowej wersji biblioteki daje rozjazd (2.12) (przyczyna) → **Wersjonuj parę: `openpyxl==3.1.5` + `types-openpyxl==3.1.5.*`.** Aktualizuj oba **razem**. Włącz `warn_unused_ignores`, by niepotrzebne `# type: ignore` same się zgłaszały (naprawa).**

**16. „`# type: ignore` wszędzie i nie wiem już, co wyłączyłem" (objaw) → **`# type: ignore` bez kodu błędu i bez uzasadnienia** (2.12) (przyczyna) → **Zawsze z kodem (`# type: ignore[union-attr]`) i zdaniem `# dlaczego`.** Włącz `warn_unused_ignores = true` — gdy problem zniknie, `mypy` powie Ci, że taśma jest już zbędna (naprawa).**

**17. „Test wydajności miga w CI — raz zielony, raz czerwony" (objaw) → **budżet ustawiony na granicy realnego czasu**, a CI ma zmienne obciążenie (inny runner, zimny cache, równoległe joby) (2.9) (przyczyna) → **Budżet z zapasem 2–3× wobec lokalnego pomiaru.** Oznacz `@pytest.mark.slow` i uruchamiaj w CI, pomijaj lokalnie (`-m \"not slow\"`) (naprawa).**

**18. „`pytest` trwa 90 sekund i nikt go już nie uruchamia" (objaw) → **wszystkie testy wydajnościowe i duże skoroszyty uruchamiane bezwarunkowo** (2.9, Przykład 4) (przyczyna) → **Markery `slow`/`performance` + `addopts` z rozsądkiem.** Rejestruj markery (`--strict-markers`), żeby literówka w nazwie markera była błędem, a nie cichym pominięciem (naprawa).**

**19. „`Too many open files` po 4 000 testów" (objaw) → **brak `wb.close()` w testach.** Każdy `load_workbook` (także na `BytesIO`) tworzy `ZipFile` trzymający uchwyt; w pętli testowej to kumuluje się (moduł 19, Przykład 1) (przyczyna) → **Fixture jako context manager wymuszający `close()`** (`wczytaj_raport` z Przykładu 2) — nawyk staje się gwarancją (naprawa).**

**20. „Test formuły przechodzi i `A3` jest `None`, więc ktoś to „naprawił" i dodał obliczanie" (objaw) → **test pułapki bez docblocka wyjaśniającego, że `None` jest OCZEKIWANE** (2.8) (przyczyna) → **Każdy test pułapki zaczynaj od `\"\"\"DOKUMENTUJE OGRANICZENIE.`** i wyjaśnij, dlaczego tak ma być. Test ma **przekazywać wiedzę**, nie tylko sprawdzać (naprawa).**

**21. „`hypothesis` generuje chińskie znaki i test pada na czymś nierealnym" (objaw) → **`st.text()` bez `alphabet`** generuje cały Unicode (2.7, Przykład 4) (przyczyna) → **Zawęź strategię do alfabetu, który realnie występuje** (`st.text(alphabet=\"0123456789., \")`). Test właściwościowy **też ma zakres** (naprawa).**

**22. „`hypothesis` krzyczy o przekroczeniu `deadline`, choć test jest poprawny" (objaw) → **domyślny limit 200 ms na przykład** jest za ostry dla funkcji budującej skoroszyt (2.7) (przyczyna) → **`@settings(deadline=None)` ze świadomym komentarzem.** Nie wyłączaj globalnie — wyłączaj **punktowo**, żeby nie ukryć realnych spowolnień w innych testach (naprawa).**

**23. „Tester `NaN != NaN`, więc test równości kwoty pada bez powodu" (objaw) → **`hypothesis` domyślnie generuje `NaN` i `inf`**, a w IEEE 754 `NaN != NaN` (2.7) (przyczyna) → **`st.floats(allow_nan=False, allow_infinity=False)`.** Chyba że **chcesz** je testować — wtedy jawnie (naprawa).**

**24. „W repozytorium leży `raport_2026.xlsx` i historia gita waży 800 MB" (objaw) → **brak `check-added-large-files` w `pre-commit`,** ktoś scommitował wygenerowany raport (2.12) (przyczyna) → **Hook `check-added-large-files` z `--maxkb=500`** + `output/` w `.gitignore`. **Raporty to artefakty, nie źródła.** A co istotne: **golden snapshots (`.txt`) MAJĄ trafiać do repo** — one są źródłem (naprawa).**

**25. „Zmieniam format raportu, a test regresji jest czerwony — więc wyłączam test" (objaw) → **test regresji wykonuje dokładnie to, do czego go stworzono** — wykrył **zamierzoną** zmianę (2.6.1) (przyczyna) → **Nie wyłączaj testu. Przeczytaj `diff`, potwierdź, że zmiana jest zamierzona, `--update-golden`, `git diff`, zatwierdź snapshot z kodem.** Golden test, który nigdy nie jest czerwony, **nic nie robi** (naprawa).**

**26. „`assert isinstance(ws, Worksheet)` — po co, skoro to oczywiste?" (objaw) → **`wb.active` może zwrócić `Chartsheet` albo `ReadOnlyWorksheet`;** stuby są tego świadome i `mypy` ma rację (2.12) (przyczyna) → **Traktuj `assert isinstance(...)` jako SPRAWDZENIE ZAŁOŻENIA, nie hack.** Jeśli raport ma `Chartsheet` tam, gdzie spodziewasz się `Worksheet`, **chcesz o tym wiedzieć natychmiast** (naprawa).**

**27. „Testy są, ale nikt nie wie, co ma być sprawdzone" (objaw) → **brak „definicji ukończenia"** — lista kryteriów jest w głowie, nie w repozytorium (2.13) (przyczyna) → **Zapisz checklistę z 2.13 jako `docs/DEFINICJA_UKONCZENIA.md` i linkuj ją w Pull Requeście.** Kryterium trudne do zautomatyzowania („otwiera się bez ostrzeżeń Excela") zostaje w checkliście jako krok ręczny (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Hamulce, nie kajdany.** Testy nie służą do zatrzymywania — służą do tego, żebyś **miał odwagę zmienić kod szybko**. Raport bez testów to raport, w którym nikt nie odważy się poprawić nagłówka. Zadaniem testu jest odpowiedź w trzy sekundy: „zielono" albo „ta konkretna linia się zmieniła".

2. **`BytesIO` to przymierzalnia, nie kasa.** `wb.save(bufor)` zapisuje **pełny, poprawny** plik do pamięci. `load_workbook(io.BytesIO(bajty))` czyta go z powrotem. Nic nie trafia na dysk. Pamiętaj o dwóch szczegółach: **`getvalue()`** jest odporne na pozycję kursora (używaj w testach), a **`seek(0)`** jest wymagane przed `read()` i strumieniowaniem. I **nie używaj `with` z `BytesIO`**, jeśli `getvalue()` wołasz po wyjściu z bloku.

3. **Nie porównuj plików — porównuj ich znaczenie.** Plik `.xlsx` **nigdy** nie będzie identyczny bajt w bajt (metadane `docProps/core.xml` + znaczniki ZIP + kompresja). Porównuj **hash rozpakowanych części** (odporny na `zlib`, dobry w CI) albo **snapshot tekstowy** (czytelny w `git diff`, dobry do diagnozy). Odcisk = inwentarz ZIP + model + liczba definicji stylów z `styles.xml`. Wykluczaj **świadomie**: `docProps/core.xml`, `docProps/app.xml`, `xl/calcChain.xml`.

4. **Trzy pułapki pomiarowe, które złapiesz dopiero po wpadce.** `len(ws.conditional_formatting)` liczy **zakresy**, nie reguły (`sum(len(cf.rules) ...)`). `ws.print_area` zwraca **kwalifikowaną, absolutną** postać (`"Sprzedaż!$A$1:$D$7"`, nie `"A1:D7"`). Style w `read_only` **nie istnieją** — do testów wyglądu potrzebujesz trybu normalnego.

5. **Wstrzyknij czas, oddziel logikę od zapisu, i zrób z tego bramkę.** `date.today()` w środku funkcji to gwarancja testów migających. Trzy światy — dane, prezentacja, zapis — muszą być rozdzielone, żeby dało się je testować osobno. Test wydajności to **budżet z zapasem 2–3×**, nie pomiar na granicy. Pokrycie to **minimum**, nie cel. `# type: ignore` zawsze z kodem błędu i uzasadnieniem, plus `warn_unused_ignores`. I na koniec: **CI działa bez Excela** — i to jest najlepszy dowód, że cały kurs od modułu 01 mówi prawdę.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY
# ==================================================================
import io
import time
import hashlib
import zipfile
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from openpyxl import Workbook, load_workbook, LXML
from openpyxl.worksheet.worksheet import Worksheet


# ==================================================================
# 2. SKOROSZYT W PAMIECI - bez dysku
# ==================================================================
def zbuduj() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    ws.append(["Kolumna", "Wartość"])
    ws.append(["Alfa", 12.5])
    bufor = io.BytesIO()
    wb.save(bufor)                     # save() przyjmuje obiekt plikopodobny
    return bufor.getvalue()            # IGNORUJE pozycje kursora

def sprawdz(bajty: bytes):
    wb = load_workbook(io.BytesIO(bajty))
    try:
        ws = wb["Dane"]
        assert ws["B2"].value == 12.5
        assert isinstance(ws["B2"].value, float)
    finally:
        wb.close()                     # OBOWIAZKOWE

# UWAGA: bufor.read() bez seek(0) zwroci b""
# bufor.seek(0); bufor.read()  <- dopiero teraz dziala
# UWAGA: NIE uzywaj `with io.BytesIO() as b:` gdy getvalue() poza blokiem!


# ==================================================================
# 3. FIXTURE-FABRYKA (conftest.py)
# ==================================================================
@pytest.fixture
def zbuduj_skoroszyt():
    def build(rows=None, sheet="Dane", ustaw=None) -> bytes:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        for w in (rows or []):
            ws.append(w)
        if ustaw:
            ustaw(wb, ws)
        b = io.BytesIO()
        wb.save(b)
        return b.getvalue()            # BAJTY = niemutowalne = bezpieczne
    return build
# ZASADA: fixture PRZYGOTOWUJE, test SPRAWDZA. Zero assert w fixture.


# ==================================================================
# 4. ASERCJE, KTORE WARTO MIEĆ
# ==================================================================
# struktura
assert wb.sheetnames == ["Dane", "Podsumowanie"]
assert ws.max_row == 6 and ws.max_column == 4

# wartosci i TYPY (najwazniejsze)
assert ws["B2"].value == 12.5
assert isinstance(ws["B2"].value, float)
assert isinstance(ws["A2"].value, (datetime, date))

# formaty - KOD, nie renderowanie (locale!)
assert ws["B2"].number_format == "#,##0.00"

# brak brakow (2.3.2)
braki = [c.coordinate for r in ws.iter_rows(min_row=2, min_col=3, max_col=3)
         for c in r if c.value is None]
assert not braki, f"Brak kwoty: {braki[:10]}"

# formatowanie warunkowe - PULAPKA: len() liczy ZAKRESY
liczba_regul = sum(len(cf.rules) for cf in ws.conditional_formatting)
assert liczba_regul == 2
inwentarz = sorted((str(cf.sqref), r.type, r.priority)
                   for cf in ws.conditional_formatting for r in cf.rules)

# tabele
assert list(ws.tables.keys()) == ["SprzedazTabela"]
assert ws.tables["SprzedazTabela"].ref == "A2:D7"

# druk - PULAPKA: print_area zwraca KWALIFIKOWANA, ABSOLUTNA postac
assert "$A$1:$D$7" in (ws.print_area or "")
assert ws.print_title_rows == "2:2"
assert ws.page_setup.orientation == "landscape"
assert ws.sheet_properties.pageSetUpPr.fitToPage is True
assert ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 0

# freeze i scalenia - str() dla bezpieczenstwa
assert str(ws.freeze_panes) == "A3"
assert {str(r) for r in ws.merged_cells.ranges} == {"A1:D1"}

# walidacje
for dv in ws.data_validations.dataValidation:
    assert dv.type in {"list", "whole", "decimal", "date", "textLength", "custom"}

# liczby zmiennoprzecinkowe - ZAWSZE approx
assert ws["C2"].value == pytest.approx(100.0 + 23.0, abs=0.01)


# ==================================================================
# 5. ZLOTY TEST - wzorzec i diff
# ==================================================================
def pytest_addoption(parser):
    parser.addoption("--update-golden", action="store_true", default=False)

@pytest.fixture
def golden(request):
    import difflib
    katalog = Path(__file__).parent / "golden"
    def porownaj(nazwa: str, tekst: str) -> None:
        katalog.mkdir(parents=True, exist_ok=True)
        plik = katalog / f"{nazwa}.txt"
        if request.config.getoption("--update-golden"):
            plik.write_text(tekst, encoding="utf-8")
            return
        if not plik.exists():
            plik.write_text(tekst, encoding="utf-8")
            pytest.fail(f"Brak wzorca - utworzono {plik}. Przejrzyj OKIEM, git add.")
        oczekiwany = plik.read_text(encoding="utf-8")
        if tekst == oczekiwany:
            return
        pytest.fail("Raport zmienil sie:\n" + "".join(difflib.unified_diff(
            oczekiwany.splitlines(keepends=True),
            tekst.splitlines(keepends=True),
            fromfile=f"{nazwa}.txt (wzorzec)", tofile=f"{nazwa}.txt (aktualny)")) )
    return porownaj


# ==================================================================
# 6. ODCISK SKOROSZYTU - hash czesci (odporny na zlib)
# ==================================================================
CZESCI_ZMIENNE = frozenset({
    "docProps/core.xml",      # openpyxl wpisuje dcterms:modified = "teraz"
    "docProps/app.xml",       # wersja aplikacji
    "xl/calcChain.xml",       # kolejnosc przeliczania formul
})

def odcisk_bajtowy(sciezka) -> str:
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in sorted(zf.namelist()):
            if nazwa in CZESCI_ZMIENNE:
                continue
            h.update(nazwa.encode()); h.update(b"\x00")
            h.update(zf.read(nazwa)); h.update(b"\x00")
    return h.hexdigest()

def inwentarz_zip(sciezka) -> dict:
    with zipfile.ZipFile(sciezka) as zf:
        nazwy = zf.namelist()
        return {
            "media": [n for n in nazwy if n.startswith("xl/media/")],
            "wykresy": [n for n in nazwy if n.startswith("xl/charts/chart")],
            "tabele": [n for n in nazwy if n.startswith("xl/tables/")],
            "ma_vba": "xl/vbaProject.bin" in nazwy,
        }

def stylistyka(sciezka) -> dict:
    """Licz definicje stylow z XML, nie z prywatnych atrybutow."""
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(sciezka) as zf:
        root = ET.fromstring(zf.read("xl/styles.xml"))
    def ile(tag):
        el = root.find(f"{NS}{tag}")
        return len(el) if el is not None else 0
    return {t: ile(t) for t in ("fonts", "fills", "borders", "numFmts", "cellXfs")}


# ==================================================================
# 7. TESTOWANIE WLASCIWOSCIOWE (hypothesis)
# ==================================================================
from hypothesis import given, settings, strategies as st

@given(st.one_of(st.none(), st.integers(),
                 st.floats(allow_nan=False, allow_infinity=False),   # NaN!
                 st.text(max_size=20)))
@settings(max_examples=300)
def test_wynik_jest_floatem_albo_none(wartosc):
    wynik = czysc_kwote(wartosc)
    assert wynik is None or isinstance(wynik, float)

@given(st.text(alphabet="0123456789., ", max_size=12))   # ZAWEZ alfabet!
def test_czyszczenie_jest_idempotentne(tekst):
    raz = czysc_kwote(tekst)
    assert czysc_kwote(raz) == raz


# ==================================================================
# 8. BUDZET WYDAJNOSCI
# ==================================================================
@pytest.mark.slow
@pytest.mark.performance
def test_budzet_100k(tmp_path):
    cel = tmp_path / "d.xlsx"
    t0 = time.perf_counter()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Dane")
    for i in range(1, 100_001):
        ws.append([i, i * 0.5])
    wb.save(cel); wb.close()
    assert time.perf_counter() - t0 < 10.0

# pyproject.toml:
# [tool.pytest.ini_options]
# addopts = "-ra --strict-markers"
# markers = ["slow: dluzsze niz 1 s", "performance: budzet wydajnosci"]


# ==================================================================
# 9. CI (GitHub Actions) - BEZ EXCELA
# ==================================================================
# - uses: actions/setup-python@v5
#   with: { python-version: "${{ matrix.python-version }}", cache: pip }
# - run: pip install "openpyxl==3.1.5" pillow lxml pytest pytest-cov hypothesis
# - run: python -c "from openpyxl import LXML; assert LXML, 'lxml nieaktywny!'"
# - run: pytest --cov=raporty --cov-report=term-missing --cov-fail-under=80
# strategy: { fail-fast: false, matrix: { python-version: ["3.9","3.12"] } }


# ==================================================================
# 10. JAKOSC STATYCZNA
# ==================================================================
# pip install ruff mypy types-openpyxl pre-commit
# ruff check . ; ruff format --check .
# mypy raporty
#
# pyproject.toml:
#   [tool.mypy] warn_unused_ignores = true   <- KLUCZOWE
#
# Zawężanie typu (bo wb.active to unia):
#   ws = wb.active
#   assert isinstance(ws, Worksheet)
#
# # type: ignore ZAWSZE z kodem i uzasadnieniem:
#   wb.defined_names.items()  # type: ignore[union-attr]  # stuby: Incomplete


# ==================================================================
# 11. PULAPKI - DO ZAPAMIETANIA
# ==================================================================
# 1) Bajt w bajt ZAWSZE czerwone -> hash czesci albo snapshot.
# 2) getvalue() tak, read() wymaga seek(0).
# 3) len(ws.conditional_formatting) = ZAKRESY, nie reguly!
# 4) ws.print_area = "Sheet!$A$1:$D$7" - nie "A1:D7".
# 5) Fixture zwraca BAJTY, nie obiekt Workbook.
# 6) Fixture nie zawiera assert.
# 7) Testuj number_format (kod), nie renderowanie (locale).
# 8) date.today() w srodku funkcji = testy migajace.
# 9) Test pułapki ZAWSZE z docblockiem "DOKUMENTUJE OGRANICZENIE".
# 10) CI bez Excela - i to jest dowod, ze nie potrzebujesz Office.
# 11) Budzet wydajnosci z zapasem 2-3x, marker slow.
# 12) Pokrycie to bramka minimum, nie cel.
```

## 9. Co dalej

Zatrzymajmy się na chwilę i spójrzmy, jaką drogę przebyłeś.

W **module 18** nauczyłeś się, że openpyxl **nie edytuje pliku w miejscu** — że zapis to przepisanie, a wszystko, czego biblioteka nie umie odczytać, nie trafi do nowego pliku. W **module 19** zobaczyłeś, że **regresja wydajności też jest cicha** — wyłączony `lxml` nie daje błędu, tylko „na serwerze jest wolniej". A w tym module **nauczyłeś się te dwie ciche awarie łapać** — odciskiem, golden testem, testem pułapki, budżetem wydajności i assertem na `LXML` w CI.

I teraz jesteś w punkcie, w którym pojawia się pytanie, które w module 20 **nie mogło** być zadane, bo nie miałbyś na nie odpowiedzi:

> **Skoro umiem już przetestować swój kod — to dlaczego mój kod tak trudno się testuje?**

Zauważ, ile wysiłku kosztowało Cię **rozdzielenie** `zbuduj_specyfikacje` od `zapisz_xlsx` w 2.10. Ile trzeba było się nagimnastykować, żeby funkcja **nie wołała `date.today()`**. Ile razy musiałeś w testach sięgać do `ws.tables[...]`, `ws.conditional_formatting`, `str(ws.print_area)` — czyli **do API openpyxl**, żeby sprawdzić coś, co jest **regułą biznesową** („kwoty mają być walutowe", „wiersze po terminie na czerwono").

To nie jest przypadkowe. **To jest sygnał architektoniczny.** Gdy test musi znać `openpyxl`, żeby sprawdzić logikę — to znaczy, że **logika i openpyxl są zbyt blisko siebie**.

I tu zaczyna się ostatnia, najbardziej dojrzała część kursu.

W **module 21** zajmiemy się bezpieczeństwem — i zobaczysz, że **plik Excel jest kodem wykonywalnym w oczach Excela**. `=HYPERLINK(...)` w komórce od użytkownika to nie tekst, to **program**, który Excel uruchomi. Zapowiedź tego masz już w teście pułapki z 2.8 — teraz zbudujemy **centralną, przetestowaną funkcję sanitacji**, bo wiesz już, jak ją przetestować.

W **module 22** odpowiemy sobie na pytanie, dlaczego skrypt, który zaczynał się od piętnastu linii, po roku ma dziewięćset i nikt nie potrafi w nim zmienić nagłówka. Zobaczysz sześć sygnałów ostrzegawczych — i jeden z nich **już znasz**: „nie da się tego przetestować". To nie jest skutek uboczny złej architektury. To jest **objaw**.

A potem wejdziemy w **moduły 23–26** — wzorce konstrukcyjne (23), strukturalne (24), behawioralne (25) i architekturę warstwową (26). I wtedy zobaczysz, po co był cały ten moduł, bo:

- **Facade** (moduł 24) to odpowiedź na to, że musiałeś znać `ws.tables[...]`, żeby sprawdzić prostą rzecz. Fasada ukryje openpyxl za własnym API — a test fasady będzie testem **tego API**, nie biblioteki.
- **Ports & Adapters** (moduł 26) to odpowiedź na to, że `openpyxl` przeciekał do logiki. Port `EksporterRaportu` pozwoli Ci **przetestować całą aplikację bez openpyxl** — z adapterem `FakeEksporter`, który zapisuje do listy Pythona.
- **Repository i Unit of Work** (moduł 26) to odpowiedź na to, że **nie kontrolowałeś, kiedy plik jest przepisywany.** A wiesz już, że przepisanie jest kosztowne i ryzykowne (moduł 18).
- **Strategy i Specification** (moduł 25) to odpowiedź na `ws.conditional_formatting.add(...)` w pięciu miejscach. Reguły opisane deklaratywnie, przetestowane **raz**, zastosowane **wszędzie**.

I najważniejsze: **wszystko, co zbudowałeś w tym module, zostanie z Tobą.** `BytesIO`, fixture-fabryka, `podsumowanie_skoroszytu`, golden test, testy właściwościowe, budżet wydajności, `ruff`, `mypy`, `pre-commit`, CI bez Excela. To nie są narzędzia „do raportów" — to są narzędzia **inżyniera**. Będą działać dokładnie tak samo, gdy za rok będziesz pisać coś zupełnie innego.

Zaczynamy od pytania, które w tym module pojawiło się mimochodem, a w następnym stanie się centralne: **co się dzieje, gdy dane w komórce są wrogie?** Moduł 21 odpowie: *„Sprawdź to testem. Ale najpierw zneutralizuj — bo Excel wykona to, co znajdzie w komórce."*