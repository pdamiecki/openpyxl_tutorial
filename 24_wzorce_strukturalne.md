# Moduł 24 — Wzorce strukturalne: warstwa, przez którą nie przecieka openpyxl

> **Część:** IV — Wzorce projektowe i architektura · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–23

## 0. W tym module nauczysz się

- **Zbudujesz `RaportExcel` — fasadę**, która ukrywa `Workbook`, `Table`, `BarChart`, `save` i cały openpyxl. Nauczysz się **testu jakości fasady**, który wykrywa dwie różne choroby: przeciek typów (`ws`, `cell`, `Font` w sygnaturach) i przeciek *pojęć* (adresy `A1`, litery kolumn, kolejność operacji).
- **Zobaczysz, czym różni się fasada od „przepakowania"** — obiekt, przez który i tak trzeba znać openpyxl, tylko dłużej się go woła. Nauczysz się trzech poziomów testu, z których każdy wyłapuje inną wersję tego samego błędu.
- **Zbudujesz dwa adaptery** — „lista + nagłówki" (dane z CSV) i „obiekt → kolumny po nazwie pola" — i **domkniesz kompromis `getattr`**, który świadomie odłożyliśmy w module 23. Zobaczysz `typing.Protocol` jako kontrakt i dowiesz się, dlaczego `hasattr` jest tam w praktyce ważniejszy niż `isinstance`.
- **Poznasz dekoratory w trzech postaciach** — pomiar czasu, log zapisu, walidacja danych przed zapisem — oraz dekorator klasy `@audytowany`. Zobaczysz **dekorator, który połyka wyjątek zapisu**, i dowiesz się, dlaczego to najgorszy możliwy sposób „uodparniania" kodu.
- **Zbudujesz dwa proxy**: `LazyWorkbook`, który **wczytuje plik dopiero przy pierwszym dostępie**, i `BramaZapisu`, która odmawia zapisu bez zatwierdzenia. Zrozumiesz, dlaczego proxy **nie przechwytuje metod specjalnych** (`len()`, `with`, iteracji) i jak to naprawić.
- **Zbudujesz Composite — drzewo sekcji raportu** (`Sekcja`, `GrupaSekcji`), które dodaje podrozdziały rekurencyjnie. Kluczowa lekcja: **Composite buduje model, a nie pisze do arkusza.** Zobaczysz, jak to samo drzewo renderować do Excela **i** do Markdowna — i dlaczego to jest prawdziwy dowód, że wzorzec działa.
- **Domykasz warstwę Front Controllerem** `eksportuj(spec, dane, cel)` — jednym wejściem, przez które przechodzą wszystkie żądania eksportu. To ten sam wzorzec, który w aplikacji webowej stoi za `StreamingResponse`.

## 1. Intuicja i analogia

### 1.1. Recepcja hotelu, czyli fasada

Wyobraź sobie hotel. Przychodzisz na recepcję i mówisz: „pokój na dwie noce, ze śniadaniem". **Nie rozmawiasz osobno z kucharzem** („czy jutro o 7:30 będą jajka?"), **z konserwatorem** („czy winda do trzeciego piętra działa?"), **z księgową** („czy moja karta przejdzie?") ani **z pokojówką** („czy pościel jest już zmieniona?"). Mówisz jedno zdanie do jednej osoby. Recepcja **wie**, do kogo zadzwonić.

I teraz rzecz najważniejsza: **recepcja nie jest tylko kolejką do ludzi.** Recepcja ma **własne pojęcia** — „pokój", „doba", „rezerwacja", „gość". Nie zna pojęcia „zaworek w instalacji hydraulicznej na drugim piętrze". Jeżeli recepcjonista zacznie odpowiadać na pytania typu „a jaki jest numer rury doprowadzającej wodę do pokoju 314?", to znaczy, że recepcja przestała być recepcją — **stała się listą telefonów do wszystkich działów**.

To jest dokładnie różnica między fasadą a „przepakowaniem". Fasada ukrywa nie tylko **wywołania**, ale i **pojęcia**. Przepakowanie ukrywa wywołania, a pojęcia zostawia na wierzchu:

```python
# PRZEPAKOWANIE: jeden poziom wyzej, ale dalej trzeba znac openpyxl
class RaportExcel:
    def arkusz(self, tytul: str, index: int) -> Worksheet:   # ← zwraca ws!
        return self._wb.create_sheet(title=tytul, index=index)

    def ustaw_styl(self, ws: Worksheet, komorka: str, font: Font) -> None:  # ← Font!
        ws[komorka].font = font
```

Popatrz na sygnatury. `Worksheet`, `Font`, `komorka: str` (`"A1"`), `index: int`. **Cały openpyxl jest tu nadal obecny** — tylko przeszedł przez jedną dodatkową warstwę. Użytkownik tej klasy musi znać numer indexu arkusza, notację A1 i klasę `Font`. Więc do czego służy ta klasa? Do niczego. **To nie jest fasada. To jest opakowanie w folię.**

Prawdziwa fasada mówi tak:

```python
raport.arkusz("dane", tytul="Dane", tab_color="1072BA")
raport.naglowek("dane", ["Data", "Klient", "Region", "Ilosc", "Cena"])
raport.wiersze("dane", wiersze, formaty=["dd.mm.yyyy", None, None, "#,##0", '#,##0.00 "zl"'])
raport.tabela("dane", nazwa="T_Dane")
raport.wykres_slupkowy("dane", kolumna_kategorii="Klient", kolumny_wartosci=["Cena"])
raport.zapisz("output/raport.xlsx")
```

**Zero openpyxl.** Zero `Worksheet`, zero `Font`, zero `A1`, zero indexów. Kolumny adresujesz **po nazwie nagłówka**, nie po literze. Wykres mówisz „po kliencie, wartości z ceny" — nie mówisz `Reference(ws, min_col=5, min_row=4, max_col=5, max_row=200)`. Jeżeli w tym momencie pomyślałeś „no dobrze, ale gdzie ja wpiszę ten kolor `FFC000`?" — to znaczy, że **zrozumiałeś**. Kolor nie jest sprawą wołającego. Kolor jest sprawą fasady i motywu.

I jeszcze jedna własność recepcji, o której łatwo zapomnieć: **recepcja pilnuje kolejności.** Nie zamówisz śniadania na jutro, jeśli najpierw nie masz rezerwacji. Fasada też pilnuje: nie ustawisz wykresu na arkuszu, który nie istnieje, i nie dodasz tabeli, jeśli nagłówek nie został jeszcze napisany. To jest **egzekwowanie protokołu** — i to jest połowa wartości fasady.

### 1.2. Przejściówka, czyli adapter

Jedziesz do Anglii z polską ładowarką. **Nie kupujesz nowego telefonu.** Kupujesz przejściówkę za dwadzieścia złotych: jedno gniazdko po jednej stronie, drugie po drugiej. **Nie zmieniasz urządzenia — zmieniasz interfejs.**

Adapter w kodzie robi dokładnie to. Twoje dane są w kształcie, w jakim przyszły ze świata (obiekt `Zamowienie`, wiersz z CSV, rekord z JSON-a). Arkusz oczekuje **krotek wartości w ustalonej kolejności**. Adapter stoi pomiędzy i tłumaczy jedno na drugie.

I zwróć uwagę na rzecz, która jest sednem tej analogii: **przejściówka nie zmienia napięcia.** Nie „poprawia" Twojego telefonu, nie konwertuje wartości. Adapter też nie powinien: jeżeli `cena` w CSV ma przecinek dziesiętny, to konwersja jest jego zadaniem — ale jeżeli `cena` jest **ujemna**, a biznes mówi, że nie może być, to **nie jest zadaniem adaptera** udawać, że jest zero. Adapter przekształca **kształt**, nie **znaczenie**.

### 1.3. Opakowanie prezentu, czyli dekorator

Masz prezent. Wkładasz go w papier, potem w torebkę, potem przyklejasz kokardę. **Prezent jest ten sam.** Zmienił się tylko sposób, w jaki go dostajesz — i każda warstwa zewnętrzna robi **jedną** rzecz: papier zdobi, torebka niesie, kokarda informuje.

Dekorator robi dokładnie to: **opakowuje funkcję w funkcję**, nie zmieniając jej wnętrza. Walidacja przed, log po, pomiar czasu wokół. I **kluczowa własność dobrego opakowania: papier nie chowa prezentu.** Jeżeli po zdjęciu papieru okaże się, że prezentu nie ma — to nie było opakowanie, a **podmiana**. W kodzie nazywa się to „dekorator, który połyka wyjątek" i jest to najgorszy błąd w tym module (pułapka 2).

### 1.4. Sekretarka, czyli proxy

Masz sekretarkę. Ludzie dzwonią do niej, a nie do Ciebie. Ona:
- **odbiera i przekazuje** (albo nie — jeśli jesteś w spotkaniu),
- **filtruje** (nie łączy z panem, który „koniecznie musi teraz"),
- **notuje**, kto dzwonił i kiedy,
- **może odmówić** — i to jest jej najważniejsza funkcja.

Proxy w kodzie robi to samo: stoi między wołającym a prawdziwym obiektem i **decyduje, czy przekazać**. `LazyWorkbook` przekazuje dopiero, gdy obiekt jest potrzebny — czyli ładuje plik przy pierwszym dostępie, a nie przy tworzeniu. `BramaZapisu` **odmawia**, jeśli nie ma zatwierdzenia.

I tu jest subtelność, którą poznać trzeba: **sekretarka nie odbiera telefonów na linii bezpośredniej.** Jeżeli w firmie istnieje numer do Twojego biurka, który dzwoni bokiem, sekretarka tego nie widzi. W Pythonie jest **dokładnie tak samo**: metody specjalne (`len()`, `with`, iteracja, `[]`) są wyszukiwane **na typie**, a nie przez `__getattr__`. Więc proxy, które przekazuje wszystko przez `__getattr__`, **nie przechwyci** `len(proxy)`. To nie jest błąd Pythona — to jest jego reguła, i o niej mówi pułapka 8.

### 1.5. Spis treści, czyli Composite

Otwierasz dokument techniczny. Ma spis treści. W spisie treści widzisz:

```
1. Wprowadzenie ...................... 3
2. Metodyka .......................... 5
   2.1. Źródła danych ................ 6
   2.2. Filtry ....................... 7
3. Wyniki ............................ 9
```

Zauważ dwie rzeczy. Po pierwsze, **rozdział 2 zawiera podrozdziały** — czyli rozdział może zawierać rozdziały. Po drugie, **każdy element ma numer strony** — czy to pojedynczy rozdział, czy grupa z podrozdziałami. Spis treści **traktuje je tak samo**, mimo że jedno jest liściem, a drugie gałęzią.

To jest **Composite**. Drzewo, w którym liść i gałąź mają ten sam interfejs. I dzięki temu możesz napisać jedno:

```python
for element in raport.elementy:
    element.renderuj(cel)          # dziala tak samo dla Akapitu i dla GrupySekcji
```

**Dlaczego to jest tu ważne?** Bo raport ma strukturę: sekcja „Metodyka" zawiera podsekcje „Źródła danych" i „Filtry", a sekcja „Wyniki" zawiera trzy tabele i dwa wykresy. Bez Composite napisałbyś `if typ == "sekcja": ... elif typ == "akapit": ...` — i to zagnieżdżone rekurencyjnie. Z Composite masz **jedną pętlę**.

I najważniejsza lekcja tego modułu, ukryta w tej analogii: **spis treści tworzy się PRZED drukiem.** Najpierw wiesz, co będzie w dokumencie i na której stronie, a potem drukujesz. Composite, który **od razu pisze do arkusza**, to spis treści tworzony w trakcie drukowania — nie policzysz stron, nie przestawisz rozdziałów i nie sprawdzisz, czy się zmieszczą.

### 1.6. Most, czyli Bridge

Jedno zdanie, bo w module 25 zobaczysz to jako pełnoprawną Strategię: **oddziel to, CO piszesz, od tego, JAK to wygląda.** Raport ma sekcje i dane (abstrakcja). Wygląd — gęsty, schowkowy, finansowy — jest osobną osią (implementacja). Zmieniasz jedną oś, druga stoi. Most: dwie drabiny, które można przesuwać niezależnie, a nie jedna drabina z przyklejonymi szczeblami na sztywno.

### 1.7. Test przecieku: jedna linia, która mówi wszystko

Zanim wejdziemy w teorię, dostajesz narzędzie. Uruchom je w swoim projekcie **teraz**:

```bash
grep -rn "openpyxl" src/ --include="*.py" | grep -v "/infrastructure/"
```

**Jak zinterpretować wynik:**

| Wynik | Co to znaczy |
|---|---|
| brak trafień | Warstwa jest czysta. openpyxl żyje tylko w infrastrukturze |
| trafienia tylko w komentarzach/docstringach | W porządku — to nie jest zależność |
| `from openpyxl.styles import Font` w innym pliku | **Przeciek stylów.** Fasada przepuszcza `Font` do wołającego (sekcja 2.1) |
| `ws[` albo `["A1"]` w innym pliku | **Przeciek struktury.** Ktoś trzyma `ws` i pisze do komórek poza fasadą (pułapka 1) |
| `load_workbook(...)` w innym pliku | **Przeciek cyklu życia.** Ktoś otwiera pliki sam, obok fasady |
| `except Exception` wokół `save()` w innym pliku | **Przeciek odpowiedzialności.** Fasada nie egzekwuje swojego protokołu (pułapka 2) |

Zapamiętaj to narzędzie, bo wrócimy do niego trzy razy: raz w teście jakości (sekcja 2.1), raz przy adapterach (2.2), raz w ćwiczeniach.

## 2. Teoria

### 2.1. Facade: definicja, trzy poziomy i granica

**Intuicja.** Jeden obiekt, jedno wejście, jedno słownictwo.

**Analogia.** Recepcja hotelu (sekcja 1.1).

**Definicja.** **Facade** to klasa, która udostępnia **uproszczony, spójny interfejs** do złożonego podsystemu. Wołający nie zna podsystemu, nie zna jego typów i nie zna jego kolejności wywołań.

Ale teraz najważniejsza część, bo tu większość ludzi się zatrzymuje o poziom za wcześnie. **Test jakości fasady ma trzy poziomy.** Wszystkie trzy muszą przejść.

#### Poziom 1 — brak typów openpyxl w sygnaturach

Test mechaniczny, który zrobisz w sekundę:

```bash
grep -n "def " src/raport.py | grep -E "Workbook|Worksheet|Cell|Font|PatternFill|Alignment|Table|Chart"
```

Wynik powinien być **pusty**. Jeżeli nie jest — masz „przepakowanie" i wiesz to w jednej linii.

Ale ten test jest **za słaby**, bo można go oszukać. Oto fasada, która przechodzi poziom 1 i nadal jest bezużyteczna:

```python
class RaportExcel:
    def wstaw(self, arkusz: str, wiersz: int, kolumna: int, wartosc) -> None:
        ...
    def sformatuj(self, arkusz: str, od_wiersza: int, do_wiersza: int,
                  od_kolumny: int, do_kolumny: int, nazwa_stylu: str) -> None:
        ...
```

Zero typów openpyxl! A mimo to wołający musi wiedzieć: że istnieje podział na wiersze i kolumny, że kolumny mają numery, że style są gdzieś zarejestrowane, że „sformatuj" potrzebuje czterech współrzędnych zamiast nazwy kolumny. **Pojęcia openpyxl przeciekły, mimo że typy nie.**

#### Poziom 2 — brak pojęć openpyxl w API

Przetestuj to pytaniem: **„gdyby openpyxl jutro został zastąpiony przez `xlsxwriter` albo przez usługę HTTP, ile linii w warstwie aplikacji bym zmienił?"**

- Jeżeli w warstwie aplikacji występuje adres `"A1"` — **przeciek**. Adres to pojęcie Excela.
- Jeżeli występuje `index=0` przy tworzeniu arkusza — **przeciek**. Kolejność arkuszy to pojęcie Excela.
- Jeżeli występuje litera kolumny (`"C"`) — **przeciek**. Litery to pojęcie Excela.
- Jeżeli występuje `ws.max_row` — **przeciek**. Wymiar arkusza to pojęcie Excela.

Prawidłowa odpowiedź to adresowanie **po nazwie kolumny**:

```python
raport.naglowek("dane", ["Data", "Klient", "Region", "Ilosc", "Cena"])
raport.wiersze("dane", dane, formaty=[...])
raport.wykres_slupkowy("dane", kolumna_kategorii="Klient", kolumny_wartosci=["Cena"])
```

„Klient" jest tu **pojęciem biznesowym**, nie pozycją w siatce. Gdyby tabela miała kolumnę „Klient" na miejscu trzecim, a nie drugim, wołający się nie zmienia. **To jest cały sens fasady.**

> **Wyjątkowo uczciwa uwaga.** Adres `"A1"` jako *argument* nie jest zbrodnią. Czasem fasada ma metodę `komentarz(arkusz, komorka: str, tekst: str)` i to jest sensowne — bo czasem naprawdę chcesz wskazać konkretną komórkę. **Zbrodnią jest, gdy adres jest podstawowym sposobem adresowania danych tabelarycznych.** W praktyce: jeżeli 90% twoich wywołań to „wiersz danych", a adresy `A1` pojawiają się tylko przy dwóch metadanych — jesteś czysty.

#### Poziom 3 — wołający nie zna protokołu

Najgłębszy poziom i ten, który odróżnia fasadę prawdziwą od dobrej. **Czy wołający musi wiedzieć, że tabelę dodaje się po danych, a wykres po tabeli? Czy musi wiedzieć, że `add_table` wymaga zakresu z nagłówkiem? Czy musi wiedzieć, że `BarChart` z `set_categories` działa inaczej niż bez?**

Nie. To wszystko jest **protokołem fasady** i fasada ma go egzekwować. Sposoby:
- **Wykrywanie stanu** — `wykres_slupkowy` sprawdza, czy nagłówek istnieje i czy wiersze przyszły. Jeśli nie — rzuć `BladRaportu` z komunikatem, który mówi, **co zrobić najpierw**.
- **Kolejność wymuszona konstrukcją** — sekcje budowane przez metodę, która najpierw tworzy arkusz, potem nagłówek, potem dane (przykład 1).
- **Walidacja na wejściu** — dekorator `@waliduj_dane` (sekcja 2.3) zamiast sprawdzania w każdym miejscu.

#### Co fasada POWINNA wystawiać — i jedna decyzja, którą trzeba podjąć świadomie

Fasada musi wystawiać **wszystko, czego potrzebuje wołający, i nic więcej**. W praktyce lista bywa taka: tworzenie arkusza, nagłówek, wiersze, tekst i KPI, tabela, wykres, ustawienia druku, zapis. To jest sześć–osiem metod i one załatwiają 95% raportów.

Problem pojawia się przy tym 5%, gdy potrzebujesz czegoś, czego fasada nie ma — na przykład **nietypowej osi wykresu** albo **wykresu kołowego z eksplozją wycinka**. I tu masz dwie uczciwe opcje:

**Opcja A — rozszerz fasadę o metodę.** `raport.wykres(..., typ="pie", eksploduj=...)`. Zbierzesz coraz więcej argumentów i po pół roku fasada ma czterdzieści metod z sześcioma argumentami każda. To nie jest katastrofa — to jest **znany koszt**, o którym mówi sekcja o antywzorcach. Ale trzeba go pilnować.

**Opcja B — świadoma klapa bezpieczeństwa (escape hatch).** Jedna metoda, która **jawnie** oddaje openpyxl, z docstringiem, że wolno jej użyć tylko w warstwie infrastruktury:

```python
    def _na_poziomie_openpyxl(self):
        """KLAPA BEZPIECZENSTWA. Zwraca Workbook.

        WOLNO uzyc TYLKO w warstwie infrastruktury i w TESTACH.
        W warstwie aplikacji to jest blad architektury (patrz modul 26).
        Swiadomie lamie kontrakt fasady - zamiast tego rozszerz API.
        """
        return self._wb
```

To jest metoda prywatna, więc poziom 1 testu (sygnatury publiczne) ją przepuszcza — a jej nazwa i docstring krzyczą, co robi. **Klapa bezpieczeństwa, która nie jest nazwana klapą bezpieczeństwa, staje się furtką, przez którą wszyscy wchodzą.**

I jeszcze jedno, bardzo praktyczne: **testy mają prawo znać openpyxl.** To nie jest przeciek. Test, który wczytuje wynik przez `load_workbook` i sprawdza komórki, jest **poprawnym testem** — bo test nie jest kodem produkcyjnym, tylko klientem z kluczem. Reguła „openpyxl tylko w infrastrukturze" dotyczy **kodu, który trafia na produkcję**. Zapisz to, bo to rozbraja wiele dylematów: nie musisz budować fasady tak, żeby dała się przetestować **bez** openpyxl. Musisz tylko zadbać, żeby **produkcja** go nie potrzebowała.

### 2.2. Adapter: dwa warianty i domknięcie kompromisu z modułu 23

**Intuicja.** Twoje dane mają jeden kształt, arkusz oczekuje drugiego. Ktoś musi to przetłumaczyć.

**Analogia.** Przejściówka do gniazdka (sekcja 1.2).

**Definicja.** **Adapter** to obiekt, który zamienia interfejs jednego typu danych na interfejs oczekiwany przez odbiorcę. W naszym przypadku: zamienia obiekt/CSV na **krotkę wartości w kolejności kolumn arkusza**.

#### Adapter „obiekt → kolumny po nazwie pola"

To ten, który w module 23 zostawił nam niedokończony wątek `getattr`. Przypomnijmy problem:

```python
# z modulu 23 - dziala, ale nie ma ZADNEGO kontraktu
wartosc = getattr(w, kolumna.pole)
```

Jeżeli `kolumna.pole == "wartosc"` i `WierszRaportu` nie ma pola `wartosc`, to `AttributeError` wybuchnie **w środku pętli po 5000 wierszy** — czyli po trzech sekundach pracy, bez informacji, która kolumna i dlaczego.

Rozwiązanie to **`typing.Protocol`** — kontrakt na **kształt**, a nie na dziedziczenie:

```python
from typing import Protocol

class PozycjaRaportu(Protocol):
    """Kształt, jakiego wymaga adapter atrybutowy.

    To NIE jest klasa bazowa - niczego nie dziedziczysz.
    To opis: 'cokolwiek, co ma te pola, nadaje sie'.
    """
    data: date
    klient: str
    region: str
    ilosc: int
    cena: float
```

Dzięki temu statyczny analizator (`mypy`) sprawdzi, że przekazujesz obiekt o właściwym kształcie, i podpowie Ci pola w edytorze. **`WierszRaportu` nic nie musi dziedziczyć** — wystarczy, że ma te pola. To jest *structural typing*: ważny jest kształt, nie rodowód.

**I teraz uczciwa część, która odróżnia teorię od praktyki.** `typing.Protocol` z oznaczeniem `@runtime_checkable` **nie sprawdza pól danych przy `isinstance`** — a w nowszych wersjach Pythona (3.12+) `isinstance` na protokole z polami danych potrafi rzucić `TypeError`. Protokoły z polami są więc **kontraktem dla analizy statycznej**, a nie dla sprawdzania w locie. W praktyce oznacza to, że:

- **Chcesz kontroli w czasie działania** → sprawdzaj ręcznie `hasattr` i rzuć własny wyjątek z listą brakujących pól.
- **Chcesz kontroli w edytorze i w CI** → wystarczy `Protocol` i `mypy`.
- **Oba naraz** → `Protocol` + metoda `sprawdz(obiekt)`, którą adapter wywołuje raz, przy pierwszym wierszu.

Trzeci wariant jest najlepszy, bo daje **komunikat w miejscu, w którym można go zrozumieć**:

```python
def sprawdz(self, obiekt) -> None:
    brakujace = [k.pole for k in self._kolumny if not hasattr(obiekt, k.pole)]
    if brakujace:
        raise BladAdaptera(
            f"Obiekt {type(obiekt).__name__} nie ma pol: {brakujace}. "
            f"Wymagany kształt: {[k.pole for k in self._kolumny]}"
        )
```

Zwróć uwagę, że komunikaty są tu **dwa razy lepsze niż wyjątki Pythona**: mówią, jakiej klasy był obiekt, czego brakowało i czego oczekiwano. To ta sama zasada, co `try/except KeyError` w fabryce z modułu 23.

#### Adapter „lista + nagłówki"

Dane z CSV przychodzą jako lista list plus wiersz nagłówków. Trzy strategie:

1. **Nagłówki z pliku, kolejność narzucona** — najbezpieczniejsza. Mówisz „kolumna o nagłówku `Cena` to pole `cena`", a kolejność w CSV nie ma znaczenia.
2. **Nagłówki nieużywane, kolejność narzucona** — najszybsza, ale krucha: przestawienie kolumn w CSV cicho psuje raport. **Nigdy tego nie rób bez sprawdzenia nagłówków.**
3. **Nagłówki używane, mapowanie jawne** — najbardziej elastyczna, wymaga słownika `{naglowek_csv: pole_domenowe}`.

W przykładzie 2 zrobimy wariant 1 z **weryfikacją**: adapter sprawdza, czy wszystkie wymagane nagłówki istnieją w pliku i rzuca wyjątek z listą brakujących. To ta sama filozofia: **błąd ma wybuchnąć w miejscu, w którym da się go zrozumieć** — a nie po cichu zostawić pustą kolumnę w raporcie dla klienta.

#### Antywzorzec: adapter z czterdziestoma argumentami

Jeżeli Twój adapter wygląda tak:

```python
def adaptuj(dane, kolumna1, kolumna2, kolumna3, format1, format2, format3,
            naglowek1, naglowek2, naglowek3, szerokosc1, szerokosc2, szerokosc3, ...):
```

to **nie jest adapter**. To jest konfiguracja rozsypana w sygnaturze. Adapter przyjmuje **specyfikację kolumn** (`Sequence[Kolumna]`) — jeden obiekt, który opisuje wszystkie kolumny — i tyle. W module 23 zbudowaliśmy `Kolumna`; teraz ona wreszcie pokazuje swoją wartość.

### 2.3. Decorator: trzy zastosowania i jedna żelazna zasada

**Intuicja.** Dołóż warstwę wokół funkcji, nie zmieniając funkcji.

**Analogia.** Opakowanie prezentu (sekcja 1.3).

**Definicja.** **Decorator** to funkcja, która przyjmuje funkcję i zwraca funkcję. Wywołanie `@dekorator` nad definicją to to samo co `f = dekorator(f)`. **Funkcja się nie zmienia — zmienia się to, co się dzieje wokół niej.**

Trzy miejsca, w których dekorator jest naturalny w naszej warstwie:

| Dekorator | Kiedy się wykonuje | Co robi |
|---|---|---|
| `@mierz_czas` | wokół wywołania | mierzy czas i zapisuje metrykę (`finally` — także przy wyjątku) |
| `@loguj_zapis` | wokół wywołania | zapisuje fakt zapisu: plik, kiedy, ile arkuszy, rozmiar |
| `@waliduj_dane` | **przed** wywołaniem | sprawdza dane i **rzuca wyjątek**, zamiast wpuszczać je dalej |
| `@audytowany` | na klasie | owija wszystkie metody publiczne fasady i notuje ich wywołania |

**Kolejność ma znaczenie.** `@waliduj_dane` powinno być **jak najbliżej funkcji** — wtedy walidacja jest pierwszą rzeczą, która się dzieje, i mierzenie czasu mierzy także walidację. `@mierz_czas` powinien być **na wierzchu**, żeby mierzyć całość. Analogia: najpierw sprawdzasz, czy prezent jest w pudełku (walidacja), potem pakujesz (log), potem oznaczasz czas (pomiar). Odwrotna kolejność mierzy czas pakowania papieru, w którym nic nie ma.

#### Żelazna zasada: dekorator nie polyka wyjątków

Najbardziej szkodliwy dekorator, jaki można napisać:

```python
# ANTYWZORZ - nigdy tak nie rob
def waliduj_dane(funkcja):
    @functools.wraps(funkcja)
    def opakowanie(dane, *args, **kwargs):
        try:
            sprawdz(dane)
        except Exception:
            return None                # ← cichy powrot: nic sie nie zapisze
        return funkcja(dane, *args, **kwargs)
    return opakowanie
```

Co tu się dzieje w praktyce? Wywołujesz `zapisz_raport(dane, "raport.xlsx")`. Dane są złe. Dekorator łapie błąd, zwraca `None`, funkcja **nie zapisuje pliku**. Twój kod sprawdza `if wynik is not None`? Nie sprawdza — bo w normalnym przepływie wynik jest ścieżką. Więc nikt nie zauważa, że **raportu nie ma**. Dopiero na spotkaniu o 10:00 ktoś mówi „a gdzie ten raport?".

To jest ten sam błąd, co „podmiana zamiast opakowania" z sekcji 1.3. **Dekorator, który połyka błąd, nie czyni kodu odpornym — czyni go niemy.** Odporność polega na tym, że błąd wybucha **w miejscu, w którym można go zrozumieć**, a nie na tym, że nie wybucha wcale.

Poprawna wersja ma **jawny wyjątek z listą problemów**:

```python
class BladWalidacjiDanych(Exception):
    def __init__(self, problemy: list[str]):
        self.problemy = problemy
        super().__init__(f"Dane odrzucone ({len(problemy)} problemow): "
                         + "; ".join(problemy))
```

Ten wyjątek jest **wyjątkiem naszej warstwy**, nie openpyxl. To ważne: wołający fasady nie powinien nigdy łapać `ValueError` z wnętrza openpyxl. Powinien łapać `BladRaportu`, `BladWalidacjiDanych`, `BladAdaptera`.

#### `functools.wraps` — nie ozdoba

Bez `@functools.wraps(funkcja)` Twój dekorator **gubi nazwę i docstring** funkcji. Skutki są konkretne i dotkliwe:

- `pytest` w raportach pokaże `test_zapisu` jako `opakowanie`;
- `monkeypatch` i `patch` po nazwie nie znajdą funkcji;
- dokumentacja i podpowiedzi edytora pokażą `*args, **kwargs`;
- debugger pokaże zagnieżdżoną funkcję bez kontekstu.

Jedna linia. Zawsze.

### 2.4. Proxy: leniwe i chronione

**Intuicja.** Obiekt, który wygląda jak prawdziwy, ale najpierw coś sprawdza albo coś odkłada.

**Analogia.** Sekretarka (sekcja 1.4).

**Definicja.** **Proxy** to obiekt o tym samym interfejsie co obiekt docelowy, który kontroluje dostęp do niego: opóźnia utworzenie, ogranicza operacje, notuje użycie.

#### `LazyWorkbook` — wczytaj, gdy naprawdę trzeba

**Intuicja.** `load_workbook()` czyta ZIP, parsuje XML-e i buduje pełny model w pamięci. Przy szablonie z siedmioma arkuszami i 200 000 komórek to sekundy i setki megabajtów. Jeżeli okaże się, że raport powstanie od zera (bo dane nie pasują do szablonu), to **cała ta praca poszła na marne**.

**Rozwiązanie.** Proxy, które udaje wczytany skoroszyt, ale ładuje go przy **pierwszym dostępie**:

```python
class LazyWorkbook:
    def __init__(self, sciezka):
        self._sciezka = Path(sciezka)
        self._wb = None                  # ← jeszcze NIE wczytane
        self._pobrania = 0

    def _zaladuj(self):
        if self._wb is None:
            self._wb = load_workbook(self._sciezka, read_only=True)
            self._pobrania += 1
        return self._wb

    def __getattr__(self, nazwa):
        return getattr(self._zaladuj(), nazwa)      # ← przekazanie dalej
```

**Dwukrotne sprawdzenie jest tu sednem**: `self._wb = None` w konstruktorze **nie ładuje pliku**, a `_zaladuj()` ładuje tylko raz, choćbyś sięgał po `sheetnames` sto razy. To ten sam mechanizm, co cache rejestru stylów z modułu 23 — leniwe, jednorazowe, z metryką.

**I teraz pułapka, dla której ta klasa istnieje w kursie.** `__getattr__` **nie przechwytuje metod specjalnych**. Dlaczego? Bo Python szuka `__len__`, `__iter__`, `__enter__` **na typie** (`type(obj)`), a nie przez atrybuty instancji. Twoje `self._zaladuj()._images` przejdzie przez `__getattr__`, ale `len(proxy)` — nie. I `with proxy:` — też nie, o ile nie zdefiniujesz `__enter__` jawnie.

Dlatego w `LazyWorkbook` **musisz** jawnie napisać:

```python
    def close(self):
        if self._wb is not None:
            self._wb.close()

    def __enter__(self):
        return self

    def __exit__(self, *wyjatki):
        self.close()
        return False
```

To jest ta „linia bezpośrednia" z analogii o sekretarce. **Nie poprawisz tego sprytnym `__getattr__`. Poprawisz to jawną deklaracją.**

#### `BramaZapisu` — zapis dopiero po zatwierdzeniu

**Intuicja.** W wielu procesach raport nie może trafić na dysk, dopóki człowiek go nie zaakceptuje. Prawnik sprawdza kwoty, kierownik sprawdza progi. Jeżeli skrypt zapisuje od razu, **nie ma na to miejsca w procesie**.

**Rozwiązanie.** Proxy, które **odmawia** — i to jest jego główna funkcja:

```python
class BramaZapisu:
    def __init__(self, fasada):
        self._fasada = fasada
        self._zatwierdzone = False
        self._kto: str | None = None
        self._historia: list[tuple[str, str]] = []

    def zatwierdz(self, kto: str) -> None:
        self._zatwierdzone = True
        self._kto = kto
        self._historia.append(("zatwierdzenie", kto))

    def zapisz(self, cel):
        if not self._zatwierdzone:
            raise BladZapisu(
                "Zapis zablokowany. Najpierw zatwierdz(...) - "
                "raport idzie do odbiorcy zewnetrznego."
            )
        return self._fasada.zapisz(cel)

    def __getattr__(self, nazwa):
        return getattr(self._fasada, nazwa)
```

Zauważ dwie rzeczy. Po pierwsze, **czytanie przechodzi swobodnie** (`podglad` przez `__getattr__`), a **zapis jest bramkowany**. To znaczy, że użytkownik może wygenerować raport, obejrzeć go, sprawdzić kwoty — i dopiero wtedy go zapisać. Po drugie, **historia zatwierdzeń** to gotowy ślad audytowy. Nie musisz pisać osobnego logu — brama jest logiem, bo przez nią przechodzi każdy zapis.

**Uczciwa granica.** Brama pilnuje **Twojego kodu**, nie użytkownika. Ktoś może wywołać `raport._fasada.zapisz(...)` i obejść bramę. To nie jest zabezpieczenie kryptograficzne (o tym mówi moduł 17: hasło w OOXML to kłódka na szufladzie, nie sejf). Brama to **kontrakt w kodzie**: „w tym systemie zapis idzie przez zatwierdzenie". Kontrakt, który łatwo złamać świadomie i trudno złamać przypadkiem — i to jest dokładnie tyle, ile potrzebujesz.

### 2.5. Composite: drzewo, które buduje model zamiast pisać do arkusza

**Intuicja.** Raport ma strukturę drzewa: sekcje zawierają podsekcje, podsekcje zawierają akapity i tabele. Chcesz jedną pętlę, która przejdzie całe drzewo.

**Analogia.** Spis treści z rozdziałami i podrozdziałami (sekcja 1.5).

**Definicja.** **Composite** to wzorzec, w którym pojedynczy element i grupa elementów mają **ten sam interfejs**, dzięki czemu struktura drzewiasta obsługuje się rekurencyjnie, jednolitym kodem.

```python
class Sekcja(ABC):
    @abstractmethod
    def tytul(self) -> str: ...
    @abstractmethod
    def renderuj(self, cel) -> None: ...
    @abstractmethod
    def policz_wiersze(self) -> int: ...     # ← planowanie!
```

Trzy klasy konkretne:

- `Akapit` — kilka linii tekstu (liść),
- `TabelaDanych` — nagłówki, wiersze, formaty (liść),
- `ParaKPI` — etykieta i wartość (liść),
- `GrupaSekcji` — **lista sekcji, i sama jest sekcją** (gałąź).

#### Kluczowe rozróżnienie: budowa modelu ≠ renderowanie

To jest **pułapka 4** z tego modułu i najważniejsza rzecz w całej sekcji. Popatrz na zły wariant:

```python
# ANTYWZORZ: sekcja pisze DO ARKUSZA w konstruktorze
class SekcjaWynikow:
    def __init__(self, ws, tytul, dane):
        ws.cell(row=1, column=1, value=tytul)   # ← pisze od razu!
        for i, w in enumerate(dane, start=2):
            ws.cell(row=i, column=1, value=w)
```

Co przez to tracisz? **Wszystko, co daje planowanie:**

1. **Nie policzysz wierszy przed renderowaniem.** A to jest potrzebne, żeby sprawdzić, czy raport zmieści się na dwóch stronach A4 (`fitToWidth=1`) albo na jednej stronie dashboardu.
2. **Nie sprawdzisz warunku bez skutku.** „Pokaż sekcję wyników tylko, jeśli sprzedaż spadła" — w wariancie złym sprawdzenie warunku polega na **wpisaniu i wycofaniu**, a wycofanie to już nie to samo (moduł 23, Builder: nie da się cofnąć kroku).
3. **Nie przetestujesz struktury bez pliku.** Test musiałby tworzyć `Workbook`, wołać sekcję i czytać komórki. A wystarczyłoby `assert raport.policz_wiersze() == 12`.
4. **Nie przemeblujesz.** Chcesz zamienić kolejność sekcji? W drzewie to `sekcje.reverse()`. W arkuszu to przenoszenie komórek, aktualizacja tabel, formatowania warunkowego i wykresów — czyli katastrofa.
5. **Nie wyrenderujesz tego samego do dwóch miejsc.** Raport do Excela i podgląd w Markdownie (albo w logu, albo w wiadomości e-mail) — z modelu tak, z „pisania do arkusza" nie.

Dlatego **Composite buduje model** (czyste obiekty Pythona, bez openpyxl), a **renderowanie** to osobny krok, który przechodzi drzewo i wywołuje **cel renderowania**.

I tu pojawia się rzecz, która sprawia, że ten wzorzec wreszcie wygląda jak coś, a nie jak biurokracja. **Cel renderowania to protokół.** Fasada może być jednym celem, a Markdown drugim:

```python
class CelRaportu(Protocol):
    def podnaglowek(self, tekst: str, poziom: int) -> None: ...
    def akapit(self, tekst: str) -> None: ...
    def tabela(self, naglowki, wiersze, formaty=None) -> None: ...
    def para(self, etykieta: str, wartosc, format_liczby=None) -> None: ...
```

Teraz masz **dwa** cele: `CelExcel` (deleguje do fasady) i `CelMarkdown` (zbiera linie tekstu). I zysk jest natychmiastowy:

- raport renderujesz do Excela w produkcji,
- **ten sam raport** renderujesz do Markdowna w teście — bez plików, bez openpyxl, w milisekundach,
- mógłbyś dodać `CelHTML` dla wersji webowej, **nie dotykając ani jednej klasy sekcji**.

To jest dokładnie ten sam mechanizm, co `EksporterRaportu` z modułu 26 — tylko tu, w wersji dla sekcji. Zwróć uwagę, że ten protokół realizuje też **Bridge** z sekcji 1.6: drzewo sekcji to „co piszemy", cel renderowania to „jak to wygląda".

### 2.6. Front Controller i Serializer: jedno wejście

**Intuicja.** Masz pięć miejsc w aplikacji, które generują raporty. W każdym ktoś zapomniał czegoś innego: w jednym nie było walidacji, w drugim nie było logu, w trzecim ktoś zapisał do pliku wejściowego, w czwartym nie było pomiaru czasu, w piątym — było, ale w formacie, którego nikt nie czyta.

Front Controller rozwiązuje to jednym ruchem: **wszystkie żądania eksportu przechodzą przez jedną funkcję.** Nie dlatego, żeby była ładna — dlatego, że **wszystko, co musi się zdarzyć zawsze, zdarza się w jednym miejscu**.

```python
def eksportuj(spec, dane, cel, *, adapter=None, zatwierdzenie=None) -> Path:
    """Front Controller. Jedno wejscie do eksportu raportu.

    ZAWSZE: wybor adaptera -> walidacja -> pomiar czasu -> zapis atomowy -> log.
    """
```

To ta sama funkcja, która w aplikacji webowej zwraca bajty zamiast pliku — i wtedy nazywa się `eksportuj_do_bajtow` i wchodzi w `StreamingResponse` (moduł 26).

**Uczciwe ostrzeżenie.** Front Controller bardzo łatwo staje się **god function** — funkcją, która robi wszystko i ma 300 linii. Zabezpieczenie: funkcja powinna **wywoływać** kroki, a nie je **implementować**. Sekcja 3, przykład 6 pokazuje wersję, w której `eksportuj` ma 30 linii, bo całą pracę wykonują: fabryka adaptera, walidator, fasada, brama. Front Controller to **spis kroków**, nie ich treść.

**Serializer** to nazwa na konwersję `spec + dane → bajty`. W naszym przypadku serializatorem jest fasada plus adapter: `spec` mówi, jakie kolumny, adapter tłumaczy dane, fasada pisze, `do_bajtow()` daje bajty. Warto tę nazwę znać, bo w webach mówi się o tym „serializacja odpowiedzi" — a problem jest **identyczny** (moduł 26, `StreamingResponse`).

### 2.7. Tabela decyzyjna: który wzorzec na jaki problem

| Problem w kodzie | Wzorzec | Sygnał, że to problem | Koszt | Kiedy NIE |
|---|---|---|---|---|
| wołający zna `ws`, `Font`, `A1` | **Facade** | `grep openpyxl` poza infrastrukturą | średni | projekt jednorazowy, jeden skrypt |
| fasada zwraca `ws` | **popraw fasada** | `-> Worksheet` w sygnaturze publicznej | niski | — **zawsze** |
| 40 argumentów w funkcji eksportu | **Adapter + spec** | sygnatura dłuższa niż 5 argumentów | średni | gdy naprawdę 3 argumenty |
| walidacja/log/pomiar rozproszone | **Decorator** | te same 5 linii `try/except/finally` w 6 miejscach | niski | jedno miejsce |
| szablon wczytywany, a nieużywany | **Proxy `LazyWorkbook`** | `load_workbook` przed sprawdzeniem warunku | niski | gdy szablon i tak zawsze jest potrzebny |
| brak kontroli, kto zapisał raport | **Proxy `BramaZapisu`** | raport idzie na zewnątrz bez zatwierdzenia | niski | narzędzie wewnętrzne dla jednej osoby |
| strukturę raportu trzeba przewidzieć | **Composite** | `if/elif` po typie sekcji, rekurencyjnie | średni | raport liniowy, jedna tabela |
| to samo do XLSX i do podglądu | **Composite + cele** | logika renderowania wpleciona w budowę | średni | tylko jeden format na zawsze |
| wygląd zależy od odbiorcy | **Bridge (→ Strategy, moduł 25)** | `if odbiorca == "zarzad"` w środku generowania | średni | jeden odbiorca |
| 5 miejsc generuje raporty, każde inaczej | **Front Controller** | różne formaty logów i walidacji | niski | jedno miejsce |

## 3. Przykłady krok po kroku

Pięć przykładów. Ten sam raport sprzedaży, ewolucja od modułu 22: najpierw fasada, potem adaptery, potem dekoratory, potem proxy, na końcu composite i Front Controller razem.

### Przykład 1 — Facade `RaportExcel`

**Problem.** W module 23 skończyliśmy z `RaportBuilder`, który zwraca `Workbook`. Ten `Workbook` wędruje do wołającego, a wołający robi `wb["Dane"]["A1"]`. **openpyxl jest na wolności.**

**Rozwiązanie.** Fasada, która trzyma `Workbook` **w środku** i nigdy go nie oddaje.

```python
"""Modul 24, przyklad 1: Fasada RaportExcel.

Kontrakt fasady:
  1. Zadna metoda PUBLICZNA nie przyjmuje i nie zwraca typu openpyxl.
  2. Adresowanie po NAZWIE kolumny, nie po literze.
  3. Fasada pilnuje KOLEJNOSCI (naglowek przed danymi, tabela po danych).
  4. Testy maja prawo znac openpyxl. Produkcja - nie.

Uruchom: python examples/24_01_fasada.py
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterable, Sequence
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# =====================================================================
# WYJATKI NASZEJ WARSTWY - wolajacy NIGDY nie lapie bledow openpyxl
# =====================================================================
class BladRaportu(Exception):
    """Bazowy blad warstwy raportowej."""


class BladProtokolu(BladRaportu):
    """Wolajacy zlamal kolejnosc: np. tabela przed naglowkiem."""


def _bezpieczna_nazwa(nazwa: str) -> str:
    """Excel wymaga nazwy tabeli bez spacji i znakow specjalnych."""
    czysta = "".join(z if (z.isalnum() or z == "_") else "_" for z in nazwa)
    return czysta if czysta and not czysta[0].isdigit() else f"T_{czysta}"


class RaportExcel:
    """FASADA. Publiczne API nie zawiera openpyxl. Nie zwraca Workbooka."""

    def __init__(self, tytul: str, autor: str = "system") -> None:
        self._wb = Workbook()                # ← trzymany w srodku
        del self._wb["Sheet"]                # domyslny arkusz nie jest potrzebny
        self._wb.properties.title = tytul
        self._wb.properties.creator = autor

        # stan wewnetrzny: zeby egzekwowac PROTOKOL (poziom 3 testu jakosci)
        self._nazwy: dict[str, str] = {}                 # klucz -> tytul w Excelu
        self._kolumny: dict[str, tuple[str, ...]] = {}   # klucz -> naglowki
        self._pierwszy_danych: dict[str, int] = {}
        self._ostatni: dict[str, int] = {}
        self._tablice: set[str] = set()

    # ------------------------------------------------------------------
    # ARKUSZE
    # ------------------------------------------------------------------
    def arkusz(
        self,
        klucz: str,
        tytul: str | None = None,
        *,
        tab_color: str | None = None,
        freeze_wiersz: int = 0,
        siatka: bool = True,
    ) -> None:
        if klucz in self._nazwy:
            raise BladRaportu(
                f"Arkusz {klucz!r} juz istnieje. Istniejace: {sorted(self._nazwy)}"
            )
        tytul_docelowy = tytul or klucz
        if tytul_docelowy in self._wb.sheetnames:
            raise BladRaportu(
                f"Tytul arkusza {tytul_docelowy!r} jest juz uzywany. "
                f"Obecne: {self._wb.sheetnames}"
            )
        ws = self._wb.create_sheet(title=tytul_docelowy)
        if tab_color:
            ws.sheet_properties.tabColor = tab_color
        if freeze_wiersz:
            ws.freeze_panes = f"A{freeze_wiersz + 1}"
        ws.sheet_view.showGridLines = siatka
        self._nazwy[klucz] = tytul_docelowy
        self._ostatni[klucz] = 0

    def _ws(self, klucz: str):
        nazwa = self._nazwy.get(klucz)
        if nazwa is None:
            raise BladRaportu(
                f"Nie ma arkusza {klucz!r}. Sa: {sorted(self._nazwy)}"
            )
        return self._wb[nazwa]

    # ------------------------------------------------------------------
    # TRESC - adresowanie PO NAZWIE KOLUMNY tam, gdzie to ma sens
    # ------------------------------------------------------------------
    def naglowek(
        self,
        klucz: str,
        etykiety: Sequence[str],
        *,
        tlo: str = "1072BA",
        kolor_tekstu: str = "FFFFFF",
    ) -> None:
        if klucz in self._kolumny:
            raise BladProtokolu(
                f"Arkusz {klucz!r} ma juz naglowek: {self._kolumny[klucz]}"
            )
        ws = self._ws(klucz)
        wiersz = self._ostatni[klucz] + 1
        for i, etykieta in enumerate(etykiety, start=1):
            c = ws.cell(row=wiersz, column=i, value=str(etykieta))
            c.font = Font(bold=True, color=kolor_tekstu)
            c.fill = PatternFill(fill_type="solid", fgColor=tlo)
            c.alignment = Alignment(horizontal="center", vertical="center")
        self._kolumny[klucz] = tuple(str(e) for e in etykiety)
        self._pierwszy_danych[klucz] = wiersz + 1
        self._ostatni[klucz] = wiersz

    def wiersze(
        self,
        klucz: str,
        dane: Iterable[Sequence[object]],
        *,
        formaty: Sequence[str | None] | None = None,
    ) -> int:
        if klucz not in self._kolumny:
            raise BladProtokolu(
                f"Arkusz {klucz!r} nie ma naglowka. Najpierw naglowek(...)."
            )
        oczekiwane = len(self._kolumny[klucz])
        ws = self._ws(klucz)
        licznik = 0
        for wartosci in dane:
            if len(wartosci) != oczekiwane:
                raise BladRaportu(
                    f"Wiersz o dlugosci {len(wartosci)} != {oczekiwane} kolumn "
                    f"w arkuszu {klucz!r}. Wartosci: {wartosci!r}"
                )
            wiersz = self._ostatni[klucz] + 1
            for i, wartosc in enumerate(wartosci, start=1):
                c = ws.cell(row=wiersz, column=i, value=wartosc)
                if formaty and formaty[i - 1]:
                    c.number_format = formaty[i - 1]
            self._ostatni[klucz] = wiersz
            licznik += 1
        return licznik

    def para(self, klucz: str, etykieta: str, wartosc,
             format_liczby: str | None = None) -> None:
        """Wiersz dwoch komorek: etykieta + wartosc. Do KPI i podsumowan."""
        ws = self._ws(klucz)
        wiersz = self._ostatni[klucz] + 1
        ws.cell(row=wiersz, column=1, value=etykieta).font = Font(bold=True)
        c = ws.cell(row=wiersz, column=2, value=wartosc)
        if format_liczby:
            c.number_format = format_liczby
        self._ostatni[klucz] = wiersz

    def tytul(self, klucz: str, tekst: str, *, rozmiar: int = 16) -> None:
        ws = self._ws(klucz)
        wiersz = self._ostatni[klucz] + 1
        ws.cell(row=wiersz, column=1, value=tekst).font = Font(bold=True, size=rozmiar)
        self._ostatni[klucz] = wiersz
```

```python
    # ------------------------------------------------------------------
    # TABELE - fasada pamieta zakres, wolajacy podaje TYLKO nazwe
    # ------------------------------------------------------------------
    def tabela(self, klucz: str, nazwa: str, *, styl: str = "TableStyleMedium9") -> None:
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        if klucz in self._tablice:
            raise BladProtokolu(f"Arkusz {klucz!r} ma juz tabele")
        pierwszy = self._pierwszy_danych[klucz] - 1      # ← wiersz naglowka
        ostatni = self._ostatni[klucz]
        if ostatni < pierwszy:
            raise BladRaportu(f"Arkusz {klucz!r} nie ma danych do tabeli")
        n = len(self._kolumny[klucz])
        ref = f"A{pierwszy}:{get_column_letter(n)}{ostatni}"
        tabela = Table(displayName=_bezpieczna_nazwa(nazwa), ref=ref)
        tabela.tableStyleInfo = TableStyleInfo(name=styl, showRowStripes=True)
        self._ws(klucz).add_table(tabela)
        self._tablice.add(klucz)

    def szerokosci(self, klucz: str, szerokosci: Sequence[float]) -> None:
        ws = self._ws(klucz)
        for i, szer in enumerate(szerokosci, start=1):
            ws.column_dimensions[get_column_letter(i)].width = szer

    # ------------------------------------------------------------------
    # WYKRES - kolumny PO NAZWIE, nie po numerze
    # ------------------------------------------------------------------
    def wykres_slupkowy(
        self,
        klucz: str,
        tytul: str,
        *,
        kolumna_kategorii: str,
        kolumny_wartosci: Sequence[str],
        pozycja: tuple[int, int] = (2, 8),        # (wiersz, kolumna) - NIE "H2"
        szerokosc_cm: float = 14.0,
        wysokosc_cm: float = 8.0,
    ) -> None:
        naglowki = self._kolumny.get(klucz)
        if not naglowki:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")

        def indeks(nazwa: str) -> int:
            if nazwa not in naglowki:
                raise BladRaportu(
                    f"Nie ma kolumny {nazwa!r} w arkuszu {klucz!r}. "
                    f"Dostepne: {list(naglowki)}"
                )
            return naglowki.index(nazwa) + 1

        ws = self._ws(klucz)
        pierwszy = self._pierwszy_danych[klucz]
        ostatni = self._ostatni[klucz]
        wykres = BarChart()
        wykres.type = "col"
        wykres.title = tytul
        wykres.y_axis.title = "Wartosc"
        wykres.legend.position = "b"

        dane = Reference(
            ws,
            min_col=indeks(kolumny_wartosci[0]),
            min_row=pierwszy - 1,
            max_col=indeks(kolumny_wartosci[-1]),
            max_row=ostatni,
        )
        wykres.add_data(dane, titles_from_data=True)
        kategorie = Reference(
            ws,
            min_col=indeks(kolumna_kategorii),
            min_row=pierwszy,
            max_row=ostatni,
        )
        wykres.set_categories(kategorie)
        wykres.width = szerokosc_cm
        wykres.height = wysokosc_cm
        ws.add_chart(wykres, ws.cell(row=pozycja[0], column=pozycja[1]).coordinate)

    # ------------------------------------------------------------------
    # WYJSCIE - jedyne dwie metody, ktore dotykaja plikow
    # ------------------------------------------------------------------
    def do_bajtow(self) -> bytes:
        """Serializacja do bajtow - do HTTP, do testow, do wysylki."""
        from io import BytesIO

        bufor = BytesIO()
        self._wb.save(bufor)
        return bufor.getvalue()

    def zapisz(self, cel: str | Path) -> Path:
        """Zapis ATOMOWY: plik tymczasowy + os.replace."""
        cel = Path(cel)
        cel.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=cel.parent)
        os.close(fd)
        try:
            self._wb.save(tmp)
            os.replace(tmp, cel)                 # atomowe na tym samym wolumenie
        except BaseException:
            Path(tmp).unlink(missing_ok=True)
            raise
        return cel

    def __repr__(self) -> str:
        return f"<RaportExcel arkusze={list(self._nazwy)}>"
```

**Co się dzieje w pamięci.** `RaportExcel` tworzy **jeden** `Workbook` w konstruktorze i trzyma go w polu prywatnym `self._wb`. Wołający nie ma do niego żadnego dostępu publicznego. Ślady stanu (`_nazwy`, `_kolumny`, `_pierwszy_danych`, `_ostatni`, `_tablice`) istnieją z konkretnego powodu: **żeby fasada mogła egzekwować protokół i odpowiadać na pytania bez zaglądania do arkusza przez `ws.max_row`**. To jest o wiele tańsze niż czytanie `max_row` przy każdym wywołaniu i o wiele bardziej przewidywalne (moduł 07: `max_row` kłamie po formatowaniu, a `read_only` wymiarów może wcale nie mieć).

I zwróć uwagę na jedną rzecz w `_ws`: fasada nigdy nie oddaje obiektu arkusza na zewnątrz. `_ws` jest prywatne i istnieje tylko w środku. **Jeżeli pojawiłoby się publiczne `def arkusz_ws(...) -> Worksheet`, cała warstwa by się rozsypała** — bo każdy zacząłby pisać `raport.arkusz_ws("dane")["A1"] = ...` i po tygodniu nie byłoby już fasady (pułapka 1).

**Co trafi do pliku** przy takim użyciu:

```python
def main() -> None:
    raport = RaportExcel("Raport sprzedazy - styczen 2026", autor="dzial BI")

    raport.arkusz("dane", tytul="Dane", tab_color="1072BA",
                  freeze_wiersz=3, siatka=False)
    raport.naglowek("dane", ["Klient", "Region", "Ilosc", "Cena"])

    dane = [
        ("Alfa sp. z o.o.", "PL", 10, 25.0),
        ("Beta S.A.",       "DE",  4, 41.5),
        ("Gamma sp. k.",    "PL", 22, 25.0),
        ("Alfa sp. z o.o.", "PL",  7, 63.0),
    ]
    ile = raport.wiersze(
        "dane", dane,
        formaty=[None, None, "#,##0", '#,##0.00 "zl"'],
    )
    raport.tabela("dane", nazwa="T_Dane")
    raport.szerokosci("dane", [24, 10, 10, 12])
    raport.wykres_slupkowy(
        "dane", "Sprzedaz wg klienta",
        kolumna_kategorii="Klient", kolumny_wartosci=["Cena"],
        pozycja=(2, 6),
    )

    raport.arkusz("podsumowanie", tytul="Podsumowanie", tab_color="70AD47", siatka=False)
    raport.tytul("podsumowanie", "Podsumowanie")
    raport.para("podsumowanie", "Liczba pozycji", ile, "#,##0")
    raport.para("podsumowanie", "Wartosc lacznie",
                sum(w[2] * w[3] for w in dane), '#,##0.00 "zl"')

    cel = raport.zapisz("output/24_01_fasada.xlsx")
    print("Zapisano:", cel, "| wierszy:", ile, "|", len(raport.do_bajtow()), "bajtow")


if __name__ == "__main__":
    main()
```

Do pliku trafią dwa arkusze. W `Dane`: nagłówek pogrubiony, biały tekst na niebieskim tle, cztery wiersze danych z formatami (tysięczny i walutowy), **prawdziwa tabela Excela** `T_Dane` na zakresie `A1:D5`, szerokości kolumn, wykres słupkowy zakotwiczony w komórce `F2`. W `Podsumowanie`: tytuł i dwa wiersze KPI. Do `xl/tables/table1.xml` trafi definicja tabeli, do `xl/charts/chart1.xml` — wykres, do `xl/drawings/drawing1.xml` — kotwica wykresu, a do `xl/styles.xml` — cztery wpisy `font` i dwa `fill`.

**Trzy rzeczy do przemyślenia:**

1. **Fasada pilnuje protokołu.** Wywołanie `raport.tabela("dane", ...)` przed `naglowek(...)` da `BladProtokolu` z komunikatem `"Arkusz 'dane' nie ma naglowka"`. Bez fasady openpyxl **pozwoliłby** na `add_table` z zakresem, który nie istnieje, a Excel przy otwarciu powiedziałby „uszkodzony plik". **Fasada zamienia błąd Excela w błąd Pythona** — i to jest jej połowa wartości.
2. **`pozycja: tuple[int, int]` zamiast `"H2"`.** To nie pedanteria. `(2, 6)` nie wymaga od wołającego znajomości notacji A1 ani umiejętności przeliczania `F` → `6`. Gdybyś chciał iść dalej, mógłbyś przyjąć `pozycja=("dane", "cena")` i wyliczać kotwicę samodzielnie. **To pokazuje, że test poziomu 2 można zaostrzać w nieskończoność — i że trzeba wiedzieć, kiedy przestać.** W praktyce `(wiersz, kolumna)` jako krotka liczb jest dobrym kompromisem.
3. **`do_bajtow()` obok `zapisz()`.** Nie jest ozdobą. To jest dokładnie ta funkcja, której potrzebuje aplikacja webowa (moduł 26, `StreamingResponse`) i której potrzebuje test (moduł 20, weryfikacja bez dysku). **Warto ją mieć od początku**, bo dodana później wymaga refaktoryzacji.

### Przykład 2 — Adapter: dwa wejścia, jedno wyjście

**Problem.** Ten sam raport ma powstać z dwóch źródeł: z listy obiektów domenowych (`WierszRaportu` z modułu 23) i z CSV-a od księgowości. Fasada oczekuje **krotek wartości w kolejności kolumn** — więc ktoś musi je zbudować.

**Rozwiązanie.** Dwa adaptery o **tym samym interfejsie** (`naglowki`, `formaty`, `wiersze`), żeby Front Controller nie musiał wiedzieć, z którym pracuje.

```python
"""Modul 24, przyklad 2: Adaptery - obiekt/CSV -> wiersze arkusza."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Protocol

from openpyxl import Workbook  # tylko w tym pliku-eksperymencie; w projekcie: infrastruktura


class BladAdaptera(Exception):
    """Dane wejsciowe nie maja kształtu wymaganego przez adapter."""


# =====================================================================
# KONTRAKT (statyczny) - ksztalt, jakiego wymaga adapter atrybutowy
# =====================================================================
class PozycjaRaportu(Protocol):
    data: date
    klient: str
    region: str
    ilosc: int
    cena: float


# =====================================================================
# KOLUMNA - jeden obiekt opisuje WSZYSTKO o kolumnie (modul 23)
# =====================================================================
@dataclass(frozen=True)
class Kolumna:
    etykieta: str
    szerokosc: float
    pole: str                            # nazwa pola w obiekcie ALBO naglowek w CSV
    format_liczby: str | None = None
    konwersja: Callable[[object], object] | None = None


KOLUMNY = (
    Kolumna("Data",    12, "data",   "dd.mm.yyyy"),
    Kolumna("Klient",  24, "klient"),
    Kolumna("Region",   8, "region"),
    Kolumna("Ilosc",    8, "ilosc",  "#,##0"),
    Kolumna("Cena",    10, "cena",   '#,##0.00 "zl"'),
)
```

```python
# =====================================================================
# ADAPTER 1: obiekt domenowy -> krotki, po NAZWACH pol
# =====================================================================
class AdapterAtrybutowy:
    def __init__(self, kolumny: tuple[Kolumna, ...] = KOLUMNY) -> None:
        self._kolumny = kolumny

    @property
    def naglowki(self) -> tuple[str, ...]:
        return tuple(k.etykieta for k in self._kolumny)

    @property
    def formaty(self) -> tuple[str | None, ...]:
        return tuple(k.format_liczby for k in self._kolumny)

    def szerokosci(self) -> tuple[float, ...]:
        return tuple(k.szerokosc for k in self._kolumny)

    def sprawdz(self, obiekt) -> None:
        """Jawna kontrola - bo isinstance na Protocol z polami jest zawodne."""
        brakujace = [k.pole for k in self._kolumny if not hasattr(obiekt, k.pole)]
        if brakujace:
            raise BladAdaptera(
                f"Obiekt {type(obiekt).__name__} nie ma pol: {brakujace}. "
                f"Wymagane: {[k.pole for k in self._kolumny]}"
            )

    def wiersze(self, dane):
        """Generator - NIE buduje listy w pamieci (modul 19)."""
        for i, obiekt in enumerate(dane):
            if i == 0:
                self.sprawdz(obiekt)          # raz, na pierwszym obiekcie
            yield tuple(getattr(obiekt, k.pole) for k in self._kolumny)


# =====================================================================
# ADAPTER 2: CSV (lista slownikow) -> krotki, po NAGLOWKACH z pliku
# =====================================================================
class AdapterTabelaryczny:
    """Przyjmuje wiersze jako Mapping[naglowek_csv, tekst].

    KLUCZOWE: wymaga, zeby naglowki w pliku byly zgodne z oczekiwanymi.
    Bez tego przestawienie kolumn w CSV cicho psuje raport.
    """

    def __init__(self, kolumny: tuple[Kolumna, ...] = KOLUMNY) -> None:
        self._kolumny = kolumny

    @property
    def naglowki(self) -> tuple[str, ...]:
        return tuple(k.etykieta for k in self._kolumny)

    @property
    def formaty(self) -> tuple[str | None, ...]:
        return tuple(k.format_liczby for k in self._kolumny)

    def szerokosci(self) -> tuple[float, ...]:
        return tuple(k.szerokosc for k in self._kolumny)

    def sprawdz_naglowki(self, obecne: set[str]) -> None:
        oczekiwane = {k.pole for k in self._kolumny}
        brakujace = oczekiwane - obecne
        if brakujace:
            raise BladAdaptera(
                f"W pliku CSV brakuje kolumn: {sorted(brakujace)}. "
                f"Znalezione: {sorted(obecne)}"
            )

    def wiersze(self, rekordy):
        for rekord in rekordy:
            wartosci = []
            for k in self._kolumny:
                surowa = rekord[k.pole]
                wartosci.append(k.konwersja(surowa) if k.konwersja else surowa)
            yield tuple(wartosci)


def jako_date(tekst: object) -> date:
    if isinstance(tekst, date):
        return tekst
    return datetime.strptime(str(tekst).strip(), "%Y-%m-%d").date()


def jako_float(tekst: object) -> float:
    return float(str(tekst).replace(",", ".").replace("\u00a0", "").strip())


def jako_int(tekst: object) -> int:
    return int(jako_float(tekst))


KOLUMNY_CSV = (
    Kolumna("Data",    12, "data",   "dd.mm.yyyy", jako_date),
    Kolumna("Klient",  24, "klient"),
    Kolumna("Region",   8, "region"),
    Kolumna("Ilosc",    8, "ilosc",  "#,##0",       jako_int),
    Kolumna("Cena",    10, "cena",   '#,##0.00 "zl"', jako_float),
)
```

```python
# =====================================================================
# FABRYKA ADAPTERA - Front Controller nie musi znac typow wejsciowych
# =====================================================================
def wybierz_adapter(dane, pierwszy=None):
    """Zwraca adapter dopasowany do KSRZTALTU danych.

    Nie sprawdzamy typu listy - sprawdzamy typ PIERWSZEGO elementu,
    bo to on mowi, czy mamy obiekty, czy wiersze z CSV.
    """
    if pierwszy is None:
        dane = list(dane)
        if not dane:
            raise BladAdaptera("Brak danych do zaadaptowania")
        pierwszy = dane[0]

    if isinstance(pierwszy, dict):
        adapter = AdapterTabelaryczny(KOLUMNY_CSV)
        adapter.sprawdz_naglowki(set(pierwszy))
        return adapter, dane
    if hasattr(pierwszy, "klient") and hasattr(pierwszy, "cena"):
        return AdapterAtrybutowy(), dane
    raise BladAdaptera(
        f"Nie wiem, jak zaadaptowac element typu {type(pierwszy).__name__}. "
        "Obslugiwane: dict (CSV) i obiekt z polami klient/cena."
    )


def wczytaj_csv(sciezka: Path, separator: str = ";") -> list[dict]:
    with sciezka.open(encoding="utf-8-sig", newline="") as f:
        return [dict(wiersz) for wiersz in csv.DictReader(f, delimiter=separator)]
```

**Co się dzieje w pamięci — trzy rzeczy warte nazwania.**

1. **`AdapterAtrybutowy.wiersze` jest generatorem, nie listą.** Wyjaśnienie w jednym zdaniu: **generator to funkcja, która produkuje wartości po jednej, „leniwie"** — analogia: **taśma produkcyjna, która podaje kolejne elementy na żądanie, zamiast wypluć cały zapas do magazynu naraz.** Dzięki temu 500 000 rekordów z bazy nigdy nie jest w całości w pamięci — fasada dostaje je po jednym i od razu zapisuje. To jest ten sam mechanizm, o którym mówi moduł 19.

2. **`sprawdz()` wywołuję tylko dla pierwszego obiektu.** To jest świadoma decyzja o kompromisie: sprawdzanie `hasattr` dla każdego z 500 000 obiektów kosztowałoby czas, a w praktyce wszystkie obiekty w liście mają ten sam typ. Jeżeli dane są heterogeniczne (lista `mix` obiektów różnych klas), to **jest błąd danych, nie adaptera** — i wybuchnie na drugim obiekcie jako `AttributeError`, a ja to akceptuję, bo wiadomo wtedy dokładnie, co się stało.

3. **`AdapterTabelaryczny` nie zamienia nagłówków na pola.** `k.pole` wskazuje **nagłówek w CSV** (`"data"`, `"klient"`), nie nazwę atrybutu. To znaczy, że ten sam obiekt `Kolumna` opisuje dwie różne rzeczy w dwóch adapterach — i to **nie jest** eleganckie. Świadomie zostawiam to niedociągnięcie, bo pokazuje realny dylemat: albo dwa różne typy kolumn (duplikacja), albo jedno pole pełniące dwie role (mniej jasne). W prawdziwym projekcie rozdzieliłbyś to na `KolumnaObiektowa` i `KolumnaTabelaryczna`, a wspólny interfejs wyciągnąłbyś w protokół. **Zauważ ten problem w ćwiczeniu 🟡** — to jest dokładnie ten rodzaj decyzji, którego nie ma w podręcznikach.

**Co trafi do pliku.** Identyczny plik jak z przykładu 1 — i to jest cała wartość adaptera. **Źródło danych może się zmienić bez zmiany fasady.**

```python
def main() -> None:
    from examples_23 import WierszRaportu  # dataclass z modulu 23

    obiekty = [
        WierszRaportu(date(2026, 1, 5), "Alfa sp. z o.o.", "PL", 10, 25.0),
        WierszRaportu(date(2026, 1, 9), "Beta S.A.", "DE", 4, 41.5),
    ]
    from_csv = [
        {"data": "2026-01-05", "klient": "Alfa sp. z o.o.", "region": "PL",
         "ilosc": "10", "cena": "25,00"},
        {"data": "2026-01-09", "klient": "Beta S.A.", "region": "DE",
         "ilosc": "4", "cena": "41,50"},
    ]

    for etykieta, dane in (("obiekty", obiekty), ("csv", from_csv)):
        adapter, rekordy = wybierz_adapter(dane)
        print(etykieta, "->", adapter.naglowki)
        for wiersz in adapter.wiersze(rekordy):
            print("   ", wiersz)
```

Wynik z obu źródeł jest identyczny: `('Data', 'Klient', 'Region', 'Ilosc', 'Cena')` i te same krotki. **Adapter jest przejściówką — i przejściówka nie zmienia napięcia** (sekcja 1.2): obie ścieżki dają te same wartości logiczne. Gdyby CSV zawierał kolumnę `id_klienta` zamiast `klient`, adapter rzuci `BladAdaptera` z listą brakujących nagłówków — zamiast cicho wygenerować raport z pustą kolumną.

### Przykład 3 — Decorator: walidacja, log, pomiar, audyt

**Problem.** Chcemy, żeby **każdy** eksport był walidowany, logowany i zmierzony. W module 22 te trzy rzeczy były rozproszone po funkcjach.

**Rozwiązanie.** Dekoratory — z żelazną zasadą z sekcji 2.3: **nie połykać wyjątków**.

```python
"""Modul 24, przyklad 3: Decoratory warstwy raportowej."""

from __future__ import annotations

import functools
import logging
import time
from collections.abc import Callable, Sequence
from pathlib import Path

DZIENNIK: list[dict] = []          # w projekcie: logger, nie lista
METRYKI: list[tuple[str, float]] = []
log = logging.getLogger("raporty")


class BladWalidacjiDanych(Exception):
    def __init__(self, problemy: Sequence[str]) -> None:
        self.problemy = tuple(problemy)
        super().__init__(
            f"Dane odrzucone ({len(self.problemy)} problemow): "
            + "; ".join(self.problemy)
        )


# ---------------------------------------------------------------------
# 1. POMIAR CZASU - mierzy TAKZE przy wyjatku (finally!)
# ---------------------------------------------------------------------
def mierz_czas(funkcja: Callable) -> Callable:
    @functools.wraps(funkcja)                     # ← nazwa i docstring zostaja
    def opakowanie(*args, **kwargs):
        start = time.perf_counter()
        try:
            return funkcja(*args, **kwargs)
        finally:                                  # ← mierzy tez bledy
            czas = time.perf_counter() - start
            METRYKI.append((funkcja.__qualname__, czas))
            log.debug("%s: %.4f s", funkcja.__qualname__, czas)
    return opakowanie


# ---------------------------------------------------------------------
# 2. LOG ZAPISU - notuje fakt, nie zmienia semantyki
# ---------------------------------------------------------------------
def loguj_zapis(funkcja: Callable) -> Callable:
    @functools.wraps(funkcja)
    def opakowanie(self, cel, *args, **kwargs):
        cel = Path(cel)
        istnial = cel.exists()
        rozmiar_przed = cel.stat().st_size if istnial else 0

        wynik = funkcja(self, cel, *args, **kwargs)   # ← wyjatek leci DALEJ

        DZIENNIK.append({
            "operacja": funkcja.__name__,
            "plik": cel.name,
            "nadpisal": istnial,
            "rozmiar_bajtow": cel.stat().st_size,
            "przyrost": cel.stat().st_size - rozmiar_przed,
        })
        return wynik                                  # ← zwraca to, co funkcja
    return opakowanie


# ---------------------------------------------------------------------
# 3. WALIDACJA - PRZED wywolaniem; rzuca, NIE polyka
# ---------------------------------------------------------------------
def waliduj_dane(walidator: Callable[[Sequence], list[str]]) -> Callable:
    """Dekorator z ARGUMENTEM: walidator zwraca liste problemow (pusta = OK)."""
    def dekorator(funkcja: Callable) -> Callable:
        @functools.wraps(funkcja)
        def opakowanie(dane, *args, **kwargs):
            problemy = walidator(dane)
            if problemy:
                raise BladWalidacjiDanych(problemy)    # ← JAWNY wyjatek warstwy
            return funkcja(dane, *args, **kwargs)
        return opakowanie
    return dekorator


def walidator_sprzedazy(dane: Sequence) -> list[str]:
    problemy: list[str] = []
    if not dane:
        problemy.append("brak danych wejsciowych")
        return problemy
    for i, w in enumerate(dane[:50], start=1):        # probka, nie calosc
        if getattr(w, "ilosc", 0) < 0:
            problemy.append(f"wiersz {i}: ujemna ilosc ({w.ilosc})")
        if getattr(w, "cena", 0) < 0:
            problemy.append(f"wiersz {i}: ujemna cena ({w.cena})")
        if not str(getattr(w, "klient", "")).strip():
            problemy.append(f"wiersz {i}: pusty klient")
        if len(problemy) >= 20:
            problemy.append("... (dalsze problemy pominięte)")
            break
    return problemy
```

```python
# ---------------------------------------------------------------------
# 4. DEKORATOR KLASY: @audytowany - owija wybrane metody publiczne
# ---------------------------------------------------------------------
def audytowany(*metody: str) -> Callable:
    """Dekorator KLASY. Owija wskazane metody publiczne i notuje wywolania.

    Uwaga: setattr(cls, ...) daje ZWYKLA funkcje - dziala jak metoda,
    bo Python wiaze ja przy dostepie przez instancje.
    """
    def dekorator(cls):
        for nazwa in metody:
            oryginal = getattr(cls, nazwa)

            @functools.wraps(oryginal)
            def opakowanie(self, *args, __oryginal=oryginal, __nazwa=nazwa, **kwargs):
                log.info("wywolanie: %s(%r, %r)", __nazwa, args[:2], list(kwargs))
                wynik = __oryginal(self, *args, **kwargs)
                log.info("zakonczone: %s", __nazwa)
                return wynik

            setattr(cls, nazwa, opakowanie)
        return cls
    return dekorator
```

**Co się dzieje w pamięci — i jeden genialny szczegół, który trzeba zrozumieć.**

Zwróć uwagę na `__oryginal=oryginal, __nazwa=nazwa` w domknięciu. To **domyślne argumenty z podwójnym podkreśleniem**, i to nie jest kaprys. Problem: w pętli `for nazwa in metody` tworzysz **funkcje w pętli**, a każda z nich odwołuje się do zmiennych `oryginal` i `nazwa` z **zasięgu zewnętrznego**. Bez domyślnych argumentów **wszystkie** opakowania widziałyby **ostatnią** wartość tych zmiennych — czyli jedno opakowanie wywołane trzy razy na ostatniej metodzie. To klasyczna pułapka „późnego wiązania w domknięciach" (ang. *late binding*).

Analogia: **rozdajesz pracownikom numery telefonów do klientów, ale zapisujesz je na tablicy.** Wszyscy patrzą na tablicę — więc po ostatniej zmianie wszyscy dzwonią do ostatniego klienta. **Domyślny argument tworzy prywatną kartkę** dla każdego pracownika. I to nie jest trick — to jest jedyny poprawny sposób na domknięcie w pętli.

**Zastosowanie — i kolejność dekoratorów:**

```python
import os
import tempfile
from openpyxl import Workbook

class Eksporter:
    @mierz_czas                                    # ← NAJBLIZEJ wywolania: mierzy wszystko
    @loguj_zapis                                   # ← notuje skutek
    def zapisz(self, cel: str | Path, wb: Workbook) -> Path:
        cel = Path(cel)
        cel.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=cel.parent)
        os.close(fd)
        try:
            wb.save(tmp)
            os.replace(tmp, cel)
        except BaseException:
            Path(tmp).unlink(missing_ok=True)
            raise
        return cel


@waliduj_dane(walidator_sprzedazy)                 # ← NAJBLIZEJ funkcji
@mierz_czas
def generuj_raport(dane, cel) -> Path:
    """Walidacja -> generowanie -> zapis. Kazdy krok widoczny w metrykach."""
    wb = Workbook()
    ws = wb.active
    ws.append(["Data", "Klient", "Region", "Ilosc", "Cena"])
    for w in dane:
        ws.append([w.data, w.klient, w.region, w.ilosc, w.cena])
    return Eksporter().zapisz(cel, wb)


@audytowany("zapisz", "do_bajtow")                 # ← na klasie: owija obie metody
class RaportAudytowany:
    def zapisz(self, cel): ...
    def do_bajtow(self) -> bytes: ...
```

**Dlaczego `@waliduj_dane` jest najbliżej funkcji:** jeżeli dasz je na wierzchu, to czas mierzony przez `@mierz_czas` będzie **zawierał walidację** — i przy złych danych zobaczysz metrykę „generowanie trwało 12 sekund", które w rzeczywistości było walidowaniem 500 000 rekordów. **Metryka, która kłamie, jest gorsza niż brak metryki.** Analogia z sekcji 2.3: najpierw sprawdzasz zawartość pudełka, potem je pakujesz.

**Co trafi do pliku.** Plik wynikowy z pięcioma kolumnami, zapisany atomowo. Ale **ważniejsze jest to, co trafia do `DZIENNIK` i `METRYKI`**:

```python
[
    {"operacja": "zapisz", "plik": "raport.xlsx", "nadpisal": False,
     "rozmiar_bajtow": 4281, "przyrost": 4281},
]
[("generuj_raport", 0.0087), ("Eksporter.zapisz", 0.0071)]
```

I jeszcze test na to, że **dekorator nie połyka wyjątku** — bo to jest kontrakt, nie konwencja:

```python
import pytest

def test_zle_dane_nie_zapisuja_pliku(tmp_path):
    zle = [WierszRaportu(date(2026, 1, 1), "", "PL", -5, 10.0)]
    cel = tmp_path / "nie_mialo_powstac.xlsx"

    with pytest.raises(BladWalidacjiDanych) as info:
        generuj_raport(zle, cel)

    assert "ujemna ilosc" in str(info.value)
    assert not cel.exists(), "Plik NIE MOZE powstac przy blednych danych"
```

Ten test jest ważniejszy niż cały dekorator. Sprawdza **dwie** rzeczy naraz: że walidacja rzuca (nie połyka) i że **nic nie trafiło na dysk**. Gdyby dekorator połykał wyjątek, test padłby na pierwszej asercji i powiedziałby Ci to wprost.

### Przykład 4 — Proxy: leniwy odczyt i brama zapisu

**Problem.** (1) Szablon jest duży i nie zawsze potrzebny. (2) W procesie musimy mieć możliwość obejrzenia raportu **przed** zapisem.

**Rozwiązanie.** Dwa proxy — jedno opóźnia, drugie broni dostępu.

```python
"""Modul 24, przyklad 4: Proxy - leniwe i chronione."""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook


class BladZapisu(Exception):
    """Zapis zablokowany przez brame."""


class LazyWorkbook:
    """Proxy leniwe. Plik wczytywany przy PIERWSZYM dostepie.

    Uwaga na dwie pulapki:
      1. __getattr__ NIE przechwytuje metod specjalnych (len, with, iter).
         Dlatego close/__enter__/__exit__ sa zdefiniowane JAWNIE.
      2. __getattr__ na obiekcie, ktory jeszcze nie ma _wb w __dict__,
         moze wpadac w nieskonczona rekurencje przy copy/pickle.
    """

    def __init__(self, sciezka: str | Path) -> None:
        object.__setattr__(self, "_sciezka", Path(sciezka))
        object.__setattr__(self, "_wb", None)
        object.__setattr__(self, "_pobrania", 0)

    def _zaladuj(self):
        wb = object.__getattribute__(self, "_wb")
        if wb is None:
            sciezka = object.__getattribute__(self, "_sciezka")
            wb = load_workbook(sciezka, read_only=True)
            object.__setattr__(self, "_wb", wb)
            object.__setattr__(self, "_pobrania", object.__getattribute__(self, "_pobrania") + 1)
        return wb

    def __getattr__(self, nazwa):
        # wywolywane TYLKO dla atrybutow nieznalezionych normalnie
        return getattr(self._zaladuj(), nazwa)

    # --- metody specjalne: JAWNIE, bo __getattr__ ich nie obsluzy ---
    def close(self) -> None:
        wb = object.__getattribute__(self, "_wb")
        if wb is not None:
            wb.close()

    def __enter__(self):
        return self

    def __exit__(self, *wyjatki):
        self.close()
        return False

    def czy_wczytany(self) -> bool:
        return object.__getattribute__(self, "_wb") is not None

    def ile_razy_wczytano(self) -> int:
        return object.__getattribute__(self, "_pobrania")

    def __repr__(self) -> str:
        stan = "wczytany" if self.czy_wczytany() else "leniwy"
        return f"<LazyWorkbook {self._sciezka.name} ({stan})>"


class BramaZapisu:
    """Proxy chronione. Czytanie swobodne, zapis tylko po zatwierdzeniu."""

    def __init__(self, fasada) -> None:
        self._fasada = fasada
        self._zatwierdzone = False
        self._kto: str | None = None
        self._historia: list[tuple[str, str]] = []

    # --- bramkowana operacja ---
    def zapisz(self, cel):
        if not self._zatwierdzone:
            raise BladZapisu(
                "Zapis zablokowany: raport nie zostal zatwierdzony. "
                "Wywolaj zatwierdz('imie i nazwisko') po weryfikacji."
            )
        cel = self._fasada.zapisz(cel)
        self._historia.append(("zapis", str(cel)))
        return cel

    def zatwierdz(self, kto: str) -> None:
        if not kto.strip():
            raise ValueError("Zatwierdzenie wymaga podania osoby")
        self._zatwierdzone = True
        self._kto = kto
        self._historia.append(("zatwierdzenie", kto))

    def historia(self) -> tuple[tuple[str, str], ...]:
        return tuple(self._historia)

    # --- wszystko inne przechodzi swobodnie (czytanie, budowa) ---
    def __getattr__(self, nazwa):
        return getattr(self._fasada, nazwa)

    def __repr__(self) -> str:
        stan = f"zatwierdzone przez {self._kto}" if self._zatwierdzone else "OTWARTE"
        return f"<BramaZapisu {stan}>"
```

**Co się dzieje w pamięci — i dlaczego `object.__setattr__`.**

Zwróć uwagę, że w `LazyWorkbook` nie używam zwykłego `self._wb = None`, a `object.__setattr__(self, "_wb", None)`. Powód jest subtelny, ale prawdziwy: **gdyby jakakolwiek operacja próbowała sięgnąć po `_wb` zanim zostanie ustawione** (na przykład `copy.copy` albo `pickle`, który odtwarza obiekt bez wywoływania `__init__`), `__getattr__` zostałby wywołany, ten by sięgnął po `self._zaladuj()`, `_zaladuj` po `self._wb`… i **masz nieskończoną rekurencję**. `object.__getattribute__`/`object.__setattr__` omijają mechanizm atrybutów i gwarantują, że nie wejdziemy w tę pętlę.

Analogia: to jest jak **zostawienie klucza zapasowego u sąsiada.** Sekretarka (`__getattr__`) nie może mieć klucza do własnego biurka, bo gdyby go zgubiła, musiałaby poprosić samą siebie. Klucz musi być u kogoś poza systemem.

**Demonstracja leniwości — i to jest najciekawszy moment przykładu:**

```python
def demo_leniwosci(szablon: Path) -> None:
    import time

    leniwy = LazyWorkbook(szablon)
    print("po utworzeniu:", leniwy.czy_wczytany())          # False - NIC nie czytano

    start = time.perf_counter()
    nazwy = leniwy.sheetnames                               # ← TU wczytuje
    print("pierwsze sheetnames:", round(time.perf_counter() - start, 4), "s")
    print("wczytany:", leniwy.czy_wczytany(), "| pobran:", leniwy.ile_razy_wczytano())

    start = time.perf_counter()
    nazwy2 = leniwy.sheetnames                              # ← z cache
    print("drugie sheetnames:", round(time.perf_counter() - start, 6), "s")

    leniwy.close()
```

Typowy wynik:

```text
po utworzeniu: False
pierwsze sheetnames: 0.0341 s
wczytany: True | pobran: 1
drugie sheetnames: 0.000002 s
```

Pierwsze wywołanie kosztuje 34 ms (wczytanie ZIP-a i zbudowanie modelu), drugie — 2 mikrosekundy. **Ten stosunek jest całym sensem `LazyWorkbook`.** Jeżeli w Twoim procesie szablon jest potrzebny tylko w 30% przypadków, to w pozostałych 70% właśnie zaoszczędziłeś 34 ms i kilkadziesiąt megabajtów. Przy 1000 raportów dziennie: 34 sekundy i gigabajty alokacji.

A `BramaZapisu` działa tak:

```python
def demo_bramy(raport) -> None:
    brama = BramaZapisu(raport)
    brama.para("podsumowanie", "Wartosc lacznie", 1249.0)   # ← przechodzi
    brama.tabela("dane", nazwa="T_Dane")                    # ← przechodzi

    try:
        brama.zapisz("output/24_04_nie_powinien.xlsx")
    except BladZapisu as blad:
        print("Zablokowane:", blad)                          # ← TU lapiemy

    brama.zatwierdz("Anna Kowalska")
    brama.zapisz("output/24_04_zatwierdzony.xlsx")           # ← teraz przechodzi
    print(brama.historia())
```

**Co trafi do pliku.** Tylko plik z ostatniej linii. Ten pierwszy **nie powstanie** — bo brama odmówiła. Do `historia()` trafi `(("zatwierdzenie", "Anna Kowalska"), ("zapis", "..."))`. To jest gotowy ślad audytowy: **kto, kiedy i co zatwierdził.**

**Trzy rzeczy do przemyślenia:**

1. **Brama nie jest zabezpieczeniem, tylko kontraktem.** Ktoś z dostępem do `raport._wb` może ją obejść. Prawdziwe zabezpieczenie to uprawnienia na poziomie systemu (moduł 17). Brama chroni **przed nieuwagą**, nie przed intencją — i to jest dokładnie tyle, ile w tym procesie wystarcza.
2. **`LazyWorkbook` nie jest darmowy.** Trzyma otwarty plik (w `read_only` — **trzeba go zamknąć**), a pierwszy dostęp jest wolniejszy od normalnego wczytania **o obowiązek sprawdzenia, czy się wczytał**. Jeżeli szablon jest potrzebny **zawsze**, leniwość to **czysty narzut**. Nie używaj proxy leniwego dlatego, że brzmi nowocześnie.
3. **Zamknięcie jest tu ryzykiem numer jeden.** W `read_only` openpyxl trzyma otwarty `ZipFile`. Kiedy `LazyWorkbook` ląduje w zmiennej, o której nikt nie pamięta, plik zostaje otwarty aż do końca procesu — a na Windows to potrafi zablokować plik. **Dlatego `close()` jest jawnie, a `__enter__`/`__exit__` — też.** To pułapka 3.

### Przykład 5 — Composite: drzewo sekcji i dwa cele renderowania

**Problem.** Raport „Metodyka" ma sekcje z podsekcjami. Część sekcji zależy od warunku („pokaż ostrzeżenie, jeśli sprzedaż spadła"). Chcemy też móc **policzyć wiersze przed renderowaniem** — żeby wiedzieć, czy raport zmieści się na dwóch stronach.

**Rozwiązanie.** Composite budujący **model**, plus dwa cele renderowania: Excel i Markdown.

```python
"""Modul 24, przyklad 5: Composite - drzewo sekcji raportu."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Protocol, Sequence


class CelRaportu(Protocol):
    """Cel renderowania. Fasada RaportExcel tez go spelnia (patrz CelExcel)."""
    def podnaglowek(self, tekst: str, poziom: int) -> None: ...
    def akapit(self, tekst: str) -> None: ...
    def tabela(self, naglowki: Sequence[str], wiersze, formaty=None) -> None: ...
    def para(self, etykieta: str, wartosc, format_liczby: str | None = None) -> None: ...


# =====================================================================
# WZORZEC COMPOSITE: Sekcja (liść) i GrupaSekcji (gałąź) - TEN SAM interfejs
# =====================================================================
class Sekcja(ABC):
    @abstractmethod
    def tytul(self) -> str: ...

    @abstractmethod
    def renderuj(self, cel: CelRaportu) -> None: ...

    @abstractmethod
    def policz_wiersze(self) -> int:
        """Ile wierszy ZAJMIE ta sekcja. Bez renderowania!"""


@dataclass(frozen=True)
class Akapit(Sekcja):
    teksty: tuple[str, ...]

    def tytul(self) -> str:
        return self.teksty[0] if self.teksty else "(pusty)"

    def renderuj(self, cel: CelRaportu) -> None:
        for tekst in self.teksty:
            cel.akapit(tekst)

    def policz_wiersze(self) -> int:
        return len(self.teksty)


@dataclass(frozen=True)
class ParaKPI(Sekcja):
    etykieta: str
    wartosc: object
    format_liczby: str | None = None

    def tytul(self) -> str:
        return self.etykieta

    def renderuj(self, cel: CelRaportu) -> None:
        cel.para(self.etykieta, self.wartosc, self.format_liczby)

    def policz_wiersze(self) -> int:
        return 1


@dataclass(frozen=True)
class TabelaDanych(Sekcja):
    naglowki: tuple[str, ...]
    wiersze: tuple[tuple, ...]
    formaty: tuple[str | None, ...] | None = None
    _tytul: str = "Tabela"

    def tytul(self) -> str:
        return self._tytul

    def renderuj(self, cel: CelRaportu) -> None:
        cel.tabela(self.naglowki, self.wiersze, self.formaty)

    def policz_wiersze(self) -> int:
        return 1 + len(self.wiersze)          # naglowek + dane


@dataclass(frozen=True)
class GrupaSekcji(Sekcja):
    """GALAZ. Lista sekcji - a sama JEST sekcja."""
    nazwa: str
    sekcje: tuple[Sekcja, ...] = ()

    def tytul(self) -> str:
        return self.nazwa

    def renderuj(self, cel: CelRaportu, poziom: int = 1) -> None:
        cel.podnaglowek(self.nazwa, poziom)
        for sekcja in self.sekcje:
            if isinstance(sekcja, GrupaSekcji):
                sekcja.renderuj(cel, poziom + 1)   # ← REKURENCJA
            else:
                sekcja.renderuj(cel)

    def policz_wiersze(self) -> int:
        # naglowek grupy + suma podsekcji (rekurencyjnie)
        return 1 + sum(s.policz_wiersze() for s in self.sekcje)

    def policz_sekcje(self) -> int:
        return sum(
            1 + (s.policz_sekcje() if isinstance(s, GrupaSekcji) else 0)
            for s in self.sekcje
        )
```

**Co się dzieje w pamięci — i to jest sedno.**

Zwróć uwagę, że **`Sekcja` nie wie nic o openpyxl**. Nie ma `ws`, nie ma komórek, nie ma arkuszy. To są **czyste obiekty Pythona** — `frozen` dataclassy, w całości w pamięci. Możesz je tworzyć, przestawiać, filtrować, kopiować, serializować do JSON-a, liczyć. I dopiero na końcu mówisz „renderuj".

A teraz cel renderowania. Oto `CelExcel`, który deleguje do fasady:

```python
class CelExcel:
    """Adapter: protokol CelRaportu -> metody fasady RaportExcel.

    To jest jednoczesnie BRIDGE: drzewo sekcji nie wie, do czego renderuje.
    """

    def __init__(self, raport, arkusz: str) -> None:
        self._raport = raport
        self._arkusz = arkusz
        self.liczba_tabel = 0

    def podnaglowek(self, tekst: str, poziom: int) -> None:
        rozmiar = max(11, 16 - 2 * poziom)
        self._raport.tytul(self._arkusz, tekst, rozmiar=rozmiar)

    def akapit(self, tekst: str) -> None:
        self._raport.tytul(self._arkusz, tekst, rozmiar=11)

    def tabela(self, naglowki, wiersze, formaty=None) -> None:
        self._raport.naglowek(self._arkusz, naglowki)
        self._raport.wiersze(self._arkusz, wiersze, formaty=formaty)
        self.liczba_tabel += 1

    def para(self, etykieta, wartosc, format_liczby=None) -> None:
        self._raport.para(self._arkusz, etykieta, wartosc, format_liczby)
```

I `CelMarkdown`, który zbiera linie tekstu — **bez openpyxl, bez plików**:

```python
class CelMarkdown:
    """Ten sam protokol, ZERO openpyxl. Uzywany w TESTACH i do podgladu."""

    def __init__(self) -> None:
        self.linie: list[str] = []

    def podnaglowek(self, tekst: str, poziom: int) -> None:
        self.linie.append(f"{'#' * (poziom + 1)} {tekst}")
        self.linie.append("")

    def akapit(self, tekst: str) -> None:
        self.linie.append(tekst)
        self.linie.append("")

    def tabela(self, naglowki, wiersze, formaty=None) -> None:
        self.linie.append("| " + " | ".join(str(n) for n in naglowki) + " |")
        self.linie.append("|" + "---|" * len(naglowki))
        for w in wiersze:
            self.linie.append("| " + " | ".join(str(x) for x in w) + " |")
        self.linie.append("")

    def para(self, etykieta, wartosc, format_liczby=None) -> None:
        self.linie.append(f"- **{etykieta}:** {wartosc}")

    def tekst(self) -> str:
        return "\n".join(self.linie)
```

**I teraz dowód, że wzorzec działa.** Ten sam obiekt raportu, dwie ścieżki:

```python
def zbuduj_raport(dane) -> GrupaSekcji:
    """Buduje MODEL. Zero openpyxl. Zero arkusza."""
    # Uwaga: `dane` to wiersze arkusza jako krotki, a nie obiekty z DTO

    suma = sum(w[2] * w[3] for w in dane)

    sekcja_dane = GrupaSekcji(
        "Dane zrodlowe",
        (
            Akapit(("Ponizsza tabela zawiera pozycje sprzedazy wg klienta.",)),
            TabelaDanych(
                naglowki=("Klient", "Region", "Ilosc", "Cena"),
                wiersze=tuple((w[0], w[1], w[2], w[3]) for w in dane),
                formaty=(None, None, "#,##0", '#,##0.00 "zl"'),
                _tytul="Pozycje sprzedazy",
            ),
        ),
    )

    elementy: list[Sekcja] = [sekcja_dane]

    # SEKCJA WARUNKOWA - decyzja podejmowana na MODELU, nie na arkuszu
    if suma > 1000:
        elementy.append(
            GrupaSekcji(
                "Uwagi",
                (Akapit((f"Wartosc sprzedazy ({suma:.2f} zl) przekracza prog 1000 zl.",
                         "Raport wymaga akceptacji kierownika.")),),
            )
        )
    else:
        elementy.append(Akapit(("Sprzedaz ponizej progu - brak dzialan.",)))

    elementy.append(
        GrupaSekcji(
            "Podsumowanie",
            (ParaKPI("Liczba pozycji", len(dane), "#,##0"),
             ParaKPI("Wartosc lacznie", suma, '#,##0.00 "zl"')),
        )
    )

    return GrupaSekcji("Raport sprzedazy - styczen 2026", tuple(elementy))


def main() -> None:
    from pathlib import Path
    import sys
    sys.path.insert(0, "examples")
    from examples_24_01_fasada import RaportExcel

    dane = [
        ("Alfa sp. z o.o.", "PL", 10, 25.0),
        ("Beta S.A.",       "DE",  4, 41.5),
        ("Gamma sp. k.",    "PL", 22, 25.0),
    ]

    drzewo = zbuduj_raport(dane)

    # --- PLANOWANIE: bez renderowania, bez openpyxl ---
    print("Sekcji w raporcie:", drzewo.policz_sekcje())
    print("Wierszy w raporcie:", drzewo.policz_wiersze())

    # --- PODGLAD: ten sam model do Markdowna (do testow, do maila) ---
    podglad = CelMarkdown()
    drzewo.renderuj(podglad)
    Path("output/24_05_podglad.md").write_text(podglad.tekst(), encoding="utf-8")

    # --- PRODUKCJA: ten sam model do Excela ---
    raport = RaportExcel("Raport sprzedazy", autor="dzial BI")
    raport.arkusz("raport", tytul="Raport", siatka=False)
    cel = CelExcel(raport, "raport")
    drzewo.renderuj(cel)
    raport.szerokosci("raport", [24, 8, 8, 12])
    raport.zapisz("output/24_05_composite.xlsx")


if __name__ == "__main__":
    main()
```

**Co trafi do pliku.** Arkusz `Raport`, w którym sekcje pojawiają się w kolejności drzewa: nagłówek główny „Raport sprzedaży — styczeń 2026", potem „Dane źródłowe" z akapitem i tabelą, potem „Uwagi" (bo suma 1249,50 > 1000 — warunek zadziałał), na końcu „Podsumowanie" z dwoma parami KPI. Do `output/24_05_podglad.md` trafi **ten sam raport w Markdownie** — i możesz go wyświetlić w terminalu, wkleić do maila albo porównać w teście.

**Trzy rzeczy do przemyślenia:**

1. **Warunek `suma > 1000` jest podejmowany na MODELU.** Nie „wpisuję, sprawdzam, wycofuję". Sprawdzam liczbę, która jest już policzona w `zbuduj_raport`, i dodaję albo nie dodaję sekcję. **W wersji antywzorcowej musiałbym to napisać na arkuszu — i wtedy nie mogę cofnąć.** Można by było co prawda usunąć wiersze, ale nie formatowanie warunkowe, tabelę ani wykres, które zdążyłyby się przykleić (moduł 07 o `delete_rows`).
2. **`policz_wiersze()` działa przed renderowaniem.** To jest praktyczna wartość: możesz sprawdzić, czy raport zmieści się na X stronach, i **podjąć decyzję** — albo wyłączyć sekcję „Dane źródłowe" w wersji dla zarządu, albo wygenerować osobny plik z załącznikiem. Bez modelu byłoby to niemożliwe.
3. **Testowanie bez openpyxl to nie trik — to dowód.** Napisz test, który renderuje drzewo do `CelMarkdown` i sprawdza, że w wyniku jest sekcja „Uwagi" przy sumie > 1000 oraz że **jej nie ma** przy sumie poniżej. Ten test nie tworzy pliku, nie wczytuje pliku, i wykonuje się w mikrosekundach. **Gdyby Composite pisał od razu do arkusza, tego testu nie dałoby się napisać.** To jest praktyczna definicja tego, czym różni się model od efektu.

### Przykład 6 — Front Controller: jedno wejście

**Problem.** Mamy już wszystko: fasadę, adaptery, dekoratory, proxy, composite. Ale **każde miejsce** w aplikacji nadal musi pamiętać o kolejności: wybierz adapter → zwaliduj → zbuduj raport → renderuj → zapisz → zaloguj. Za tydzień ktoś o czymś zapomni.

**Rozwiązanie.** `eksportuj()` — jedna funkcja, do której trafiają **wszystkie** żądania eksportu.

```python
"""Modul 24, przyklad 6: Front Controller + Serializer."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from examples_24_01_fasada import RaportExcel
from examples_24_02_adaptery import wybierz_adapter
from examples_24_03_dekoratory import (
    METRYKI, DZIENNIK, mierz_czas, waliduj_dane, walidator_sprzedazy,
)
from examples_24_04_proxy import BramaZapisu
from examples_24_05_composite import CelExcel, CelMarkdown, zbuduj_raport


@dataclass(frozen=True)
class WynikEksportu:
    """DTO wyniku - warstwa aplikacji nie dostaje typow openpyxl."""
    cel: Path
    arkusze: tuple[str, ...]
    wierszy: int
    sekcji: int
    bajtow: int
    czas_s: float


@mierz_czas
def eksportuj(
    dane,
    cel: str | Path,
    *,
    podglad_markdown: bool = False,
    zatwierdzenie: str | None = None,
) -> WynikEksportu:
    """FRONT CONTROLLER. Jedno wejscie do eksportu raportu.

    Kroki (funkcja je WYWOLUJE, nie IMPLEMENTUJE):
      1. adapter dopasowany do ksztaltu danych,
      2. walidacja danych,
      3. zbudowanie MODELU (drzewa sekcji) - bez openpyxl,
      4. renderowanie do fasady,
      5. zapis przez brame (opcjonalnie),
      6. zwrot DTO z metrykami.
    """
    import time
    start = time.perf_counter()

    adapter, rekordy = wybierz_adapter(dane)          # 1
    rekordy = list(rekordy)

    # 2 - walidacja na krotkach (adapter juz znormalizowal dane)
    _sprawdz_surowe(rekordy)

    drzewo = zbuduj_raport(rekordy)                   # 3 - MODEL
    sekcji = drzewo.policz_sekcje()

    raport = RaportExcel("Raport sprzedazy", autor="system")   # 4
    raport.arkusz("raport", tytul="Raport", siatka=False)
    cel_renderowania = CelExcel(raport, "raport")
    drzewo.renderuj(cel_renderowania)
    raport.szerokosci("raport", [24, 8, 8, 12])

    if podglad_markdown:                              # 4b - ten sam model
        cel_md = CelMarkdown()
        drzewo.renderuj(cel_md)
        Path(cel).with_suffix(".md").write_text(cel_md.tekst(), encoding="utf-8")

    # 5 - zapis
    brama = BramaZapisu(raport)
    if zatwierdzenie:
        brama.zatwierdz(zatwierdzenie)
    wynik_path = brama.zapisz(cel)

    bajtow = len(raport.do_bajtow())

    return WynikEksportu(                             # 6
        cel=wynik_path,
        arkusze=tuple(raport._nazwy.values()),
        wierszy=len(rekordy),
        sekcji=sekcji,
        bajtow=bajtow,
        czas_s=time.perf_counter() - start,
    )


def _sprawdz_surowe(rekordy) -> None:
    """Walidacja krotek (adapter juz znormalizowal typy)."""
    problemy: list[str] = []
    if not rekordy:
        raise BladWalidacjiDanych(("brak danych wejsciowych",))
    for i, w in enumerate(rekordy[:50], start=1):
        klient, region, ilosc, cena = w
        if not str(klient).strip():
            problemy.append(f"wiersz {i}: pusty klient")
        if ilosc < 0:
            problemy.append(f"wiersz {i}: ujemna ilosc ({ilosc})")
        if cena < 0:
            problemy.append(f"wiersz {i}: ujemna cena ({cena})")
        if len(problemy) >= 20:
            problemy.append("... (dalsze problemy pominiete)")
            break
    if problemy:
        raise BladWalidacjiDanych(problemy)
```

**Co się dzieje w pamięci.** `eksportuj` **nie implementuje** ani adaptera, ani walidacji, ani renderowania, ani zapisu. On je **wywołuje**. To jest zabezpieczenie przed `god function`: jeżeli widzisz w `eksportuj` więcej niż dwadzieścia linii logiki, to znaczy, że któryś krok został wpisany **w treści**, a nie wywołany. Zwróć uwagę, że funkcja ma **cztery** bloki — adapter, walidacja, render, zapis. I że każdy z nich da się zastąpić **bez zmiany tej funkcji**.

I jeszcze jedno: zwracany jest **`WynikEksportu`** — DTO z modułu 23. Warstwa aplikacji nie dostaje `Workbooka`, nie dostaje `Worksheet`, nie dostaje openpyxl. Dostaje **ścieżkę, liczby i czas**. To jest dokładnie ten poziom czystości, którego wymaga moduł 26.

**Co trafi do pliku.** `output/raport.xlsx` (zapisany przez bramę, więc z zatwierdzeniem albo z wyjątkiem), opcjonalnie `output/raport.md`. Zwracane DTO:

```text
WynikEksportu(cel=output/raport.xlsx, arkusze=('Raport',),
              wierszy=3, sekcji=3, bajtow=5281, czas_s=0.0312)
```

**Trzy rzeczy do przemyślenia:**

1. **Front Controller to `spis kroków`, nie ich treść.** Cztery linie: `wybierz_adapter`, `_sprawdz_surowe`, `zbuduj_raport`, `BramaZapisu.zapisz`. Jeżeli w Twoim `eksportuj` jest `ws.cell(...)` — krok jest wpisany w treść, i po miesiącu ta funkcja ma 200 linii (pułapka 11).
2. **`eksportuj` jest udekorowane `@mierz_czas`.** To znaczy, że **każdy** eksport trafia do `METRYKI` — niezależnie od tego, kto go wywoła. To jest różnica między „wdrożyliśmy pomiar" a „mamy pomiar w jednym miejscu, o którym trzeba pamiętać".
3. **Wersja webowa różni się jedną linią.** Zamiast `brama.zapisz(cel)` robisz `raport.do_bajtow()` i zwracasz `bytes`. **Cała pozostała logika jest identyczna.** To jest ostateczny dowód, że warstwa jest dobrze zbudowana: zmiana kanału wyjścia nie wymaga dotknięcia niczego innego (moduł 26).

## 4. Anatomia API

### Część A — konstrukcje Pythona używane w tym module

| Konstrukcja | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `abc.ABC` + `@abstractmethod` | klasa abstrakcyjna | — | **Nie da się utworzyć instancji** — wymusza implementację w podklasach. Podstawa `Sekcja` |
| `typing.Protocol` | kontrakt strukturalny | — | „Cokolwiek, co ma ten kształt". **Bez dziedziczenia** — podstawowe dla `PozycjaRaportu` i `CelRaportu` |
| `@runtime_checkable` | `isinstance` na protokole | — | **Uwaga:** sprawdza **metody**, a pola danych w nowszych Pythonach (3.12+) potrafią rzucić `TypeError`. Wolimy `hasattr` ręcznie |
| `__getattr__(self, nazwa)` | przechwytuje **nieznalezione** atrybuty | — | Wywoływane tylko, gdy normalne wyszukanie zawiodło. **Nie łapie metod specjalnych** |
| `__getattribute__` | przechwytuje **wszystkie** atrybuty | — | Wymaga ostrożności — łatwo o rekurencję. Użyliśmy `object.__getattribute__` |
| `object.__setattr__` / `object.__getattribute__` | omijają mechanizm atrybutów | — | Klucz zapasowy u sąsiada: chroni `LazyWorkbook` przed nieskończoną rekurencją |
| `__enter__` / `__exit__` | protokół `with` | — | **Muszą być jawne** w proxy — `__getattr__` ich nie obsłuży |
| `functools.wraps` | kopiuje `__name__`, `__doc__`, `__wrapped__` | — | Bez tego `pytest`, `monkeypatch` i debugger tracą nazwę funkcji |
| domknięcie + domyślny argument | zamrożenie wartości z pętli | `__x=x` | **Late binding w pętlach** — bez tego wszystkie opakowania widzą ostatnią wartość |
| `@property` | pole wyliczane | — | `getattr(obj, "wartosc")` **wywołuje** property — dlatego kolumny wyliczane działają |
| `collections.abc.Sequence` | typ dla „listy" | — | Adapter przyjmuje `Sequence`, więc działa i `list`, i `tuple` |
| `dataclass(frozen=True)` | niemutowalna sekcja | — | Zawartość przez `tuple`, bo `frozen` chroni tylko przypisanie (pułapka z modułu 23) |
| `isinstance(x, dict)` | rozpoznanie wiersza CSV | — | W fabryce adaptera: rozpoznajemy **kształt**, nie typ listy |
| `hasattr` | sprawdzenie obecności pola | — | Zastępuje `runtime_checkable` dla pól danych. Czytelny komunikat w `sprawdz()` |
| generator (`yield`) | produkcja wierszy leniwie | — | „Taśma produkcyjna" — nie trzyma całych danych w pamięci (moduł 19) |
| `functools.singledispatch` | wybór implementacji po typie | — | Alternatywa dla ręcznej fabryki adaptera, gdy wariantów jest dużo |
| `inspect.signature` | odczyt sygnatury | — | Podstawa **automatycznego** testu jakości fasady (patrz 🟢) |
| `logging` zamiast `print` | kanał diagnostyczny | — | `log.info`, `log.debug` — testowalne, wyłączalne, z poziomami |
| `time.perf_counter` | pomiar czasu | — | **Monotoniczny** — właściwy do pomiarów. `finally`, żeby mierzyć też błędy |

### Część B — `openpyxl` w kontekście tego modułu

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | nowy skoroszyt | `write_only`, `iso_dates` | **Fasada tworzy go raz i trzyma prywatnie** |
| `wb.properties.title` / `.creator` | metadane dokumentu | — | Ślad audytowy. Ustawiane w konstruktorze fasady |
| `wb.create_sheet(title=, index=)` | nowy arkusz | `title`, `index` | Fasada **nie wystawia `index`** — kolejność wynika z wywołań |
| `wb.sheetnames` | lista tytułów | — | Fasada używa do sprawdzenia kolizji nazw |
| `ws.sheet_properties.tabColor` | kolor zakładki | `"RRGGBB"` | Argument fasady, nie wołającego z `PatternFill` |
| `ws.sheet_view.showGridLines` | linie siatki | `bool` | W fasadzie: `siatka: bool = True` |
| `ws.freeze_panes` | zamrożenie okien | `"A4"` | Fasada liczy `f"A{freeze_wiersz + 1}"` — wołający podaje **liczbę** |
| `ws.cell(row=, column=, value=)` | dostęp do komórki | `row`, `column`, `value` | **Tylko wewnątrz fasady.** Nigdy nie wychodzi na zewnątrz |
| `cell.coordinate` | adres komórki jako `"F2"` | — | Wykorzystane **wewnętrznie** do zakotwiczenia wykresu |
| `ws.add_table(Table(displayName=, ref=))` | tabela Excela | `ref` **zawiera nagłówek** | `displayName` bez spacji i znaków specjalnych (moduł 12) |
| `TableStyleInfo(name=, showRowStripes=)` | styl tabeli | — | Ustawiane po konstrukcji `Table` |
| `ws.add_chart(chart, anchor)` | wstawienie wykresu | `anchor` — adres `"F2"` | Kotwica = komórka lewego-górnego narożnika |
| `BarChart()` | wykres słupkowy | — | `type = "col"`, `title`, `y_axis.title`, `legend.position` |
| `Reference(ws, min_col=, min_row=, max_col=, max_row=)` | zakres dla wykresu | — | Fasada liczy **indeksy z nazw kolumn** — to jest sedno poziomu 2 |
| `chart.add_data(ref, titles_from_data=)` | dodanie serii | `titles_from_data` | Bez `titles_from_data` legenda pokaże `Series1` |
| `chart.set_categories(ref)` | etykiety osi X | — | Bez tego wykres pokaże `1, 2, 3` |
| `chart.width` / `chart.height` | wymiary wykresu | cm | Fasada wystawia `szerokosc_cm`/`wysokosc_cm` |
| `ws.column_dimensions[litera].width` | szerokość kolumny | — | Fasada tłumaczy indeks na literę przez `get_column_letter` |
| `openpyxl.utils.get_column_letter(n)` | `5` → `"E"` | — | Jedyna zamiana indeksu na literę — i **tylko wewnątrz fasady** |
| `wb.save(path)` | zapis | ścieżka lub obiekt plikopodobny | Fasada zapisuje **atomowo** (`mkstemp` + `os.replace`) |
| `wb.save(BytesIO())` | zapis do pamięci | — | Podstawa `do_bajtow()` — dla HTTP i testów |
| `wb._nazwy` (nasze pole) | nie openpyxl | — | Stan własny fasady — **tak, to jest właściwe miejsce na ten stan** |
| `ws._images`, `wb._fonts` | **atrybuty prywatne** | — | Diagnostyka. **Nie buduj na nich logiki produkcyjnej** |
| `load_workbook(path, read_only=True)` | odczyt strumieniowy | `read_only`, `data_only`, `keep_vba` | W `LazyWorkbook`. **Wymaga `wb.close()`!** (pułapka 3) |
| `wb.close()` | zamknięcie uchwytów | — | **Obowiązkowe** przy `read_only` — inaczej wyciek i blokada pliku |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — fasada z czterema metodami i automatycznym testem jakości.**

Weź `examples/24_01_fasada.py`. Zredukuj go do **czterech metod publicznych**:

```python
class RaportMinimalny:
    def arkusz(self, klucz: str, tytul: str | None = None) -> None: ...
    def naglowek(self, klucz: str, etykiety: Sequence[str]) -> None: ...
    def wiersze(self, klucz: str, dane, *, formaty=None) -> int: ...
    def zapisz(self, cel: str | Path) -> Path: ...
```

Cztery. Nie pięć. Nie sześć. **Nic więcej.**

**Wymagania:**

1. **Żadna metoda nie przyjmuje ani nie zwraca typu openpyxl.** Sprawdź to `grep`-em:
   ```bash
   grep -n "def " examples/24_cw1.py | grep -E "Workbook|Worksheet|Cell|Font"
   ```
   Wynik: pusty.
2. **`_ws` jest prywatne.** Sprawdź, że nie ma publicznego akcesora do arkusza.
3. **Od drugiego wywołania `naglowek` na tym samym arkuszu → wyjątek.** Bo nagłówek pisze się raz.
4. **Wiersz o złej długości → wyjątek z komunikatem zawierającym oczekiwaną liczbę kolumn.**
5. **Zapis jest atomowy** (`mkstemp` + `os.replace`).

**Test, który jest gwiazdą tego ćwiczenia.** Napisz test, który **mechanicznie** sprawdza poziom 1 jakości fasady — tak, żeby nie trzeba było o nim pamiętać:

```python
import inspect
import pytest

WOLNO = {"self", "cls", "args", "kwargs"}
ZAKAZANE_NAZWY = {"ws", "wb", "worksheet", "workbook", "cell", "komorka_ws"}
ZAKAZANE_TYPY = {
    "Workbook", "Worksheet", "Cell", "Font", "PatternFill",
    "Alignment", "Border", "Table", "BarChart", "Reference",
}

def _metody_publiczne(cls):
    for nazwa, obiekt in inspect.getmembers(cls, inspect.isfunction):
        if not nazwa.startswith("_"):
            yield nazwa, inspect.signature(obiekt)

def test_fasada_nie_wypuszcza_openpyxl():
    bledy = []
    for nazwa, sygnatura in _metody_publiczne(RaportMinimalny):
        for parametr in sygnatura.parameters.values():
            if parametr.name in ZAKAZANE_NAZWY:
                bledy.append(f"{nazwa}: parametr {parametr.name!r} przecieka openpyxl")

            # UWAGA: przy `from __future__ import annotations` adnotacje sa STRINGAMI
            tekst = str(parametr.annotation)
            for zakazany in ZAKAZANE_TYPY:
                if zakazany in tekst:
                    bledy.append(f"{nazwa}: adnotacja {parametr.annotation!r} przecieka openpyxl")
    assert not bledy, "Przeciek w API fasady:\n  - " + "\n  - ".join(bledy)
```

**Uwaga na jedną rzecz, którą ten test pokazuje wprost.** Jeżeli używasz `from __future__ import annotations`, adnotacje są **stringami** — więc `str(parametr.annotation)` działa, ale `parametr.annotation is Worksheet` **nie**. To pierwsza rzecz do sprawdzenia, gdy test „nie widzi" oczywistego przecieku.

**W pliku `output/24_cw1_wnioski.md` odpowiedz:**

1. **Ile metod naprawdę potrzebujesz do wygenerowania sensownego raportu?** Cztery wystarczyły? Gdybyś miał dodać piątą — jaką i dlaczego **nie** jest to `komorka(klucz, adres, wartosc)`?
2. **Czy test `test_fasada_nie_wypuszcza_openpyxl` da się oszukać?** Wymyśl sposób, w którym fasada przechodzi test, a mimo to przecieka (podpowiedź: przeciek przez **stan**, nie przez sygnaturę — np. metoda `arkusze()` zwracająca... co?).
3. **Poziom 2 i poziom 3 — czy Twój minimalny `RaportMinimalny` je przechodzi?** Adresowanie jest po nazwie arkusza (`klucz`) — a nie po indeksie. Ale `naglowek` przyjmuje listę stringów, więc **jest** zależny od kolejności kolumn. Sprawdź, czy to jest przeciek i uzasadnij.

### 🟡 Warsztat

**Zadanie 2 — adapter dla dwóch formatów wejściowych.**

To jest **to samo** zadanie co przykład 2, ale z poprawką, którą tam świadomie zostawiłem jako niedociągnięcie.

**Wymagania:**

1. **Rozdziel dwa typy kolumn.** W przykładzie 2 jedno pole `Kolumna.pole` pełniło dwie role (nazwa atrybutu obiektu **i** nagłówek w CSV). Rozdziel to:
   ```python
   @dataclass(frozen=True)
   class KolumnaObiektowa:
       etykieta: str
       szerokosc: float
       pole: str                      # nazwa atrybutu
       format_liczby: str | None = None

   @dataclass(frozen=True)
   class KolumnaTabelaryczna:
       etykieta: str
       szerokosc: float
       naglowek: str                  # naglowek w pliku
       format_liczby: str | None = None
       konwersja: Callable[[object], object] | None = None
   ```

2. **Znajdź wspólny protokół** dla obu i zdefiniuj go jako `KolumnaSpec` (`Protocol`), żeby fasada mogła przyjąć dowolną z nich.

3. **Trzy wejścia, nie dwa:**
   - `list[WierszRaportu]` (obiekty domenowe),
   - `list[dict]` z CSV — **nagłówki muszą się zgadzać** (rozdzielacz `;`),
   - `list[list]` — dane bez nagłówków, gdzie **kolejność jest narzucona**. To ten wariant, o którym mówiłem „kruchy". Zaimplementuj go, ale:
     - **dodaj wymóg potwierdzenia**: `AdapterPozycyjny(kolumny, potwierdzone_zrodlo="ksiegowosc_2026.csv")`. Bez potwierdzenia → wyjątek.
     - **udokumentuj ryzyko** w docstringu.

4. **Walidacja przed konwersją.** `sprawdz_naglowki(obecne)` dla CSV; `sprawdz(obiekt)` dla obiektów; **dla pozycyjnego**: `sprawdz_liczbe_kolumn(n)`. Każda rzuca `BladAdaptera` z komunikatem mówiącym, **czego brakuje i co jest dostępne**.

5. **Napisz trzy testy:**
   - ten sam zestaw danych w trzech formatach daje **identyczne krotki**;
   - brakujący nagłówek w CSV → `BladAdaptera` z nazwą brakującej kolumny;
   - `AdapterPozycyjny` bez `potwierdzone_zrodlo` → wyjątek.

**Pomiar (obowiązkowy).** Zmierz i zapisz w `output/24_cw2_pomiar.md`:

- czas adaptacji 100 000 rekordów dla każdego z trzech adapterów,
- pamięć szczytową (`tracemalloc`) dla wersji z listą i dla wersji z generatorem.

To drugie jest najciekawsze: **generator powinien dać wyraźnie mniejszy szczyt pamięci** — a jeżeli nie dał, sprawdź, czy przypadkiem nie robisz `list(adapter.wiersze(dane))` „na wszelki wypadek". To bardzo częsty błąd: generator, który jest natychmiast materializowany, to lista ze zbędnym opakowaniem.

**W pliku `output/24_cw2_wnioski.md` odpowiedz:**

1. **Czy rozdzielenie `KolumnaObiektowa` / `KolumnaTabelaryczna` było dobrą decyzją?** Ile linii kodu to kosztowało, a ile czytelności kupiło?
2. **Jak zmieniłby się wspólny protokół `KolumnaSpec`, gdyby doszedł czwarty format (JSON z zagnieżdżeniami)?** Gdzie postawiłbyś granicę — kiedy adapter przestaje być adapterem, a staje się parserem?
3. **Konwersje (`jako_date`, `jako_float`) — czy to zadanie adaptera?** Uzasadnij. Sprawdź swoją odpowiedź na pytaniu: jeżeli `cena` w CSV ma wartość `"abc"`, czy to błąd **formatowania danych** (adapter), czy błąd **danych** (walidacja)? Kto z nich powinien o tym zdecydować?
4. **Trzy adaptery, trzy testy — czy test „identyczne krotki" naprawdę coś sprawdza?** Tak, ale **co konkretnie**? Co by się musiało zepsuć, żeby ten test padł?
5. **Adapter pozycyjny z potwierdzeniem** — czy to dobre rozwiązanie, czy biurokracja? Co by się stało w Twoim zespole, gdybyś wymagał `potwierdzone_zrodlo` przy każdym użyciu? Kiedy wymaganie staje się przeszkodą, którą ludzie obchodzą?

### 🔴 Wyzwanie

**Zadanie 3 — Composite: raport z sekcjami generowanymi warunkowo.**

Największe ćwiczenie tego modułu. Zbudujesz **pełny generator raportu** oparty na Composite, z sekcjami warunkowymi, planowaniem i **dwoma celami renderowania**.

**Krok 1 — rozbuduj drzewo sekcji.**

Do `Sekcja` z przykładu 5 dodaj trzy nowe typy:

```python
@dataclass(frozen=True)
class WykresSlupkowy(Sekcja):
    tytul_wykresu: str
    kolumna_kategorii: str
    kolumny_wartosci: tuple[str, ...]

@dataclass(frozen=True)
class Formatka(Sekcja):
    """Sekcja, ktora ma INNY wyglad niz zwykly akapit - np. ramka ostrzezenia."""
    tytul_ramki: str
    tresc: tuple[str, ...]

@dataclass(frozen=True)
class WarunkowaSekcja(Sekcja):
    """Opakowuje sekcje i renderuje ja TYLKO gdy warunek jest prawdziwy."""
    warunek: bool
    sekcja: Sekcja
```

**Wymagania dla `WarunkowaSekcja`:**

1. **`renderuj` nic nie robi, gdy `warunek` jest `False`.** Zero wywołań do celu.
2. **`policz_wiersze` zwraca 0, gdy warunek jest fałszywy.** To jest kluczowe — **planowanie musi widzieć warunek**. Inaczej `policz_wiersze()` skłamie i cała wartość modelu przepada.
3. `tytul()` zwraca tytuł zawiniętej sekcji (dla diagnostyki).

**Krok 2 — pełny generator z warunkami biznesowymi.**

Napisz `zbuduj_raport_biznesowy(dane, *, odbiorca: str) -> Sekcja`, który buduje drzewo z takimi warunkami:

| Warunek | Sekcja |
|---|---|
| zawsze | „Dane źródłowe" z tabelą |
| zawsze | „Podsumowanie" z KPI |
| `suma > 5000` | „Alert: wysoka sprzedaż" (Formatka) |
| `suma < 500` | „Alert: sprzedaż poniżej normy" (Formatka) |
| którykolwiek region ma stratę (cena < 10) | „Analiza ryzyka" z tabelą |
| `odbiorca == "zarzad"` | „Wnioski dla zarządu" (3 akapity) |
| `odbiorca != "zarzad"` | „Załącznik: szczegóły" z pełną tabelą |
| `len(dane) > 100` | **NIE** pokazuj pełnej tabeli, tylko podsumowanie per region |

Zwróć uwagę na dwa ostatnie warunki: **ten sam raport dla różnych odbiorców** to zapowiedź wzorca Bridge i Strategy z modułu 25.

**Krok 3 — dwa cele i porównanie.**

Napisz `podglad_markdown(drzewo) -> str`, która używa `CelMarkdown`. **Ważne:** tabela w Markdownie przy 200 wierszach zrobi się ogromna — więc `CelMarkdown` powinien mieć limit (`max_wierszy_tabeli=20`) i po jego przekroczeniu pisać `... (N kolejnych wierszy pominięto)`. **Dodaj to jako parametr konstruktora**, bo to jest decyzja prezentacji, nie sekcji.

**Krok 4 — planowanie i dopasowanie do strony.**

1. Policz `drzewo.policz_wiersze()` dla trzech przypadków: 5, 50 i 500 wierszy danych.
2. Napisz funkcję `dopasuj_do_strony(drzewo, max_wierszy: int) -> Sekcja`, która **wyłącza** najmniej ważne sekcje, żeby zmieścić raport w limicie. Priorytet: KPI > podsumowanie > tabele > akapity > formatki. Jeżeli po wyłączeniu wszystkiego nadal nie mieści się — **rzuć wyjątek**, nie generuj raportu, który „wyjdzie na 40 stron".
3. Zapisz trzy pliki dla trzech rozmiarów danych i **sprawdź w Excelu**, czy `print_area` i liczba stron są sensowne.

**Krok 5 — test bez openpyxl.**

To jest **najważniejszy krok tego zadania** i dowód, że Composite został dobrze zbudowany:

```python
def test_warunkowa_sekcja_znika_z_modelu():
    drzewo = zbuduj_raport_biznesowy(duze_dane(), odbiorca="analityk")
    # suma < 500 -> nie ma alertu o normie
    md = podglad_markdown(drzewo)
    assert "ponizej normy" not in md
    assert "Dane zrodlowe" in md

def test_policz_wiersze_widzi_warunek():
    male = zbuduj_raport_biznesowy(male_dane(), odbiorca="analityk")
    duze = zbuduj_raport_biznesowy(duze_dane(), odbiorca="analityk")
    assert male.policz_wiersze() < duze.policz_wiersze()

def test_widok_zarzadu_nie_zawiera_pelnej_tabeli():
    drzewo = zbuduj_raport_biznesowy(duze_dane(), odbiorca="zarzad")
    md = podglad_markdown(drzewo)
    assert "Wnioski dla zarzadu" in md
    assert "Zalacznik: szczegoly" not in md
```

Te testy **nie tworzą żadnego pliku** i nie potrzebują openpyxl. Jeżeli nie potrafisz ich napisać, to znaczy, że gdzieś w `Sekcja` trafił `ws` — i trzeba wrócić do kroku 1.

**Kryteria oceny (rubryka):**

| Kryterium | Waga | Na co patrzę |
|---|---|---|
| Zero openpyxl w klasach sekcji | ★★★ | `grep openpyxl` w pliku sekcji → pusto |
| `policz_wiersze` widzi warunki | ★★★ | testy z kroku 5 przechodzą |
| Warunek `odbiorca` działa | ★★☆ | dwa raporty, dwa różne drzewa |
| Limit tabeli w podglądzie | ★★☆ | `max_wierszy_tabeli` jako parametr celu, nie sekcji |
| `dopasuj_do_strony` z wyjątkiem | ★★☆ | brak cichego „wyjdzie jak wyjdzie" |
| Oba cele renderują to samo drzewo | ★★★ | Excel i Markdown z jednego modelu |
| Raport otwiera się bez ostrzeżeń | ★★★ | **sprawdź okiem** |

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe decyzje i fragmenty</strong></summary>

**Decyzja 1: `WarunkowaSekcja` musi wpływać na plan, nie tylko na render.**

```python
@dataclass(frozen=True)
class WarunkowaSekcja(Sekcja):
    """Opakowuje sekcje. Wylaczona = nie istnieje dla CALEJ reszty systemu.

    To jest caly sens: policz_wiersze() tez ja widzi jako zero.
    Bez tego planowanie klamie, a raport "na dwie strony" wychodzi na siedem.
    """
    warunek: bool
    sekcja: Sekcja

    def tytul(self) -> str:
        return self.sekcja.tytul()

    def renderuj(self, cel: CelRaportu, poziom: int = 1) -> None:
        if not self.warunek:
            return                                  # ← ZERO wywolan do celu
        if isinstance(self.sekcja, GrupaSekcji):
            self.sekcja.renderuj(cel, poziom)
        else:
            self.sekcja.renderuj(cel)

    def policz_wiersze(self) -> int:
        return self.sekcja.policz_wiersze() if self.warunek else 0

    def policz_sekcje(self) -> int:
        return 1 + (self.sekcja.policz_sekcje()
                    if self.warunek and isinstance(self.sekcja, GrupaSekcji) else 0)
```

Zwróć uwagę na symetrię: **dwa `if self.warunek`** — jeden w `renderuj`, jeden w `policz_wiersze`. **Trzeci pojawi się, gdy dodasz `policz_sekcje`.** To jest sygnał, że `WarunkowaSekcja` mogłaby być abstrakcyjna z metodą `_aktywna()` — ale przy trzech miejscach **to jest nadmiar**. Reguła trzech z modułu 22 działa też tutaj.

**Decyzja 2: drzewo budowane z warunkami w jednym miejscu.**

```python
PROG_WYSOKIEJ = 5000.0
PROG_NISKIEJ = 500.0
MAX_WIERSZY_TABELI_W_RAPORCIE = 100

def zbuduj_raport_biznesowy(dane, *, odbiorca: str) -> Sekcja:
    """Buduje MODEL. Zero openpyxl, zero arkusza, zero plikow."""
    if not dane:
        raise ValueError("Brak danych - nie buduje sie raportu z niczego")

    suma = sum(cena * ilosc for _, _, ilosc, cena in
               ((w[0], w[1], w[2], w[3]) for w in dane))
    regiony: dict[str, float] = {}
    for _, region, ilosc, cena in dane:
        regiony[region] = regiony.get(region, 0.0) + ilosc * cena

    elementy: list[Sekcja] = []

    # --- 1. Dane zrodlowe: pelna tabela TYLKO dla malych zbiorow ---
    if len(dane) <= MAX_WIERSZY_TABELI_W_RAPORCIE:
        elementy.append(GrupaSekcji(
            "Dane zrodlowe",
            (TabelaDanych(
                naglowki=("Klient", "Region", "Ilosc", "Cena"),
                wiersze=tuple(tuple(w) for w in dane),
                formaty=(None, None, "#,##0", '#,##0.00 "zl"'),
                _tytul="Pozycje sprzedazy",
            ),),
        ))
    else:
        elementy.append(GrupaSekcji(
            "Dane zrodlowe",
            (Akapit((f"Zbior {len(dane)} pozycji - pokazano zestawienie per region.",
                     "Pelne dane znajduja sie w zalaczniku."),),
             TabelaDanych(
                 naglowki=("Region", "Wartosc"),
                 wiersze=tuple((r, round(v, 2)) for r, v in sorted(regiony.items())),
                 formaty=(None, '#,##0.00 "zl"'),
                 _tytul="Sprzedaz per region",
             )),
        ))

    # --- 2. Alerty: JEDEN z nich moze sie pokazac, nigdy oba ---
    alert: Sekcja | None = None
    if suma > PROG_WYSOKIEJ:
        alert = Formatka(
            "Alert: wysoka sprzedaz",
            (f"Wartosc {suma:.2f} zl przekracza prog {PROG_WYSOKIEJ:.0f} zl.",
             "Wymagana akceptacja kierownika przed wysylka do klienta."),
        )
    elif suma < PROG_NISKIEJ:
        alert = Formatka(
            "Alert: sprzedaz ponizej normy",
            (f"Wartosc {suma:.2f} zl jest ponizej progu {PROG_NISKIEJ:.0f} zl.",
             "Sprawdz, czy dane zrodlowe sa kompletne."),
        )
    if alert is not None:
        elementy.append(alert)

    # --- 3. Analiza ryzyka: warunek na DANYCH, nie na sumie ---
    ryzykowne = tuple((w[0], w[2], w[3]) for w in dane if w[3] < 10)
    elementy.append(WarunkowaSekcja(
        warunek=bool(ryzykowne),
        sekcja=GrupaSekcji("Analiza ryzyka", (
            Akapit(("Pozycje z cena jednostkowa ponizej 10 zl:",)),
            TabelaDanych(
                naglowki=("Klient", "Ilosc", "Cena"),
                wiersze=ryzykowne,
                formaty=(None, "#,##0", '#,##0.00 "zl"'),
                _tytul="Pozycje niskocenne",
            ),
        )),
    ))

    # --- 4. Podsumowanie ---
    elementy.append(GrupaSekcji("Podsumowanie", (
        ParaKPI("Liczba pozycji", len(dane), "#,##0"),
        ParaKPI("Wartosc lacznie", suma, '#,##0.00 "zl"'),
        ParaKPI("Regionow", len(regiony), "#,##0"),
    ),))

    # --- 5. Rozgalezienie po ODBIORCY (zapowiedz Bridge, modul 25) ---
    if odbiorca == "zarzad":
        region_max = max(regiony, key=regiony.get)
        elementy.append(GrupaSekcji("Wnioski dla zarzadu", (
            Akapit((f"Najwiekszy region: {region_max} "
                    f"({regiony[region_max]:.2f} zl).",)),
            Akapit(("Sprzedaz skoncentrowana w regionie o najwyzszym udziale - "
                    "ryzyko zaleznosci od jednego rynku.",)),
            Akapit(("Rekomendacja: dywersyfikacja kanalow sprzedazy.",)),
        ),))
    else:
        elementy.append(GrupaSekcji("Zalacznik: szczegoly", (
            Akapit(("Zestawienie pelne - do weryfikacji przez analityka.",)),
            TabelaDanych(
                naglowki=("Klient", "Region", "Ilosc", "Cena"),
                wiersze=tuple(tuple(w) for w in dane),
                formaty=(None, None, "#,##0", '#,##0.00 "zl"'),
                _tytul="Pelne dane",
            ),
        ),))

    return GrupaSekcji("Raport sprzedazy - styczen 2026", tuple(elementy))
```

**Decyzja 3: `CelMarkdown` z limitem tabeli jako decyzja CELU, nie sekcji.**

```python
class CelMarkdown:
    def __init__(self, *, max_wierszy_tabeli: int = 20) -> None:
        self.linie: list[str] = []
        self._max = max_wierszy_tabeli

    def tabela(self, naglowki, wiersze, formaty=None) -> None:
        wiersze = list(wiersze)
        self.linie.append("| " + " | ".join(str(n) for n in naglowki) + " |")
        self.linie.append("|" + "---|" * len(naglowki))
        for w in wiersze[: self._max]:
            self.linie.append("| " + " | ".join(str(x) for x in w) + " |")
        if len(wiersze) > self._max:
            self.linie.append(
                f"| ... pominięto {len(wiersze) - self._max} wierszy |"
                + " |" * (len(naglowki) - 1)
            )
        self.linie.append("")
```

**Dlaczego limit należy do celu, a nie do sekcji:** bo sekcja nie wie, że trafi do Markdowna. **Ta sama tabela ze 100 wierszami w Excelu jest w porządku, a w Markdownie jest katastrofą.** Decyzja o obcięciu zależy od **medium**, nie od treści. To ta sama zasada co w Bridge: „co piszemy" jest w jednej osi, „jak to wygląda" w drugiej.

**Decyzja 4: `dopasuj_do_strony` — wyjątek zamiast cichego przydługiego raportu.**

```python
PRIORYTETY = ("ParaKPI", "TabelaDanych", "Akapit", "Formatka", "WykresSlupkowy")

def dopasuj_do_strony(drzewo: Sekcja, max_wierszy: int) -> Sekcja:
    """Wylacza najmniej wazne sekcje, az raport zmiesci sie w limicie."""
    if drzewo.policz_wiersze() <= max_wierszy:
        return drzewo
    if not isinstance(drzewo, GrupaSekcji):
        return drzewo          # pojedynczej sekcji nie da sie okroic

    # usuwamy od najmniej waznych typow, od konca
    do_usuniecia = list(drzewo.sekcje)
    for typ in reversed(PRIORYTETY):
        if not do_usuniecia:
            break
        do_usuniecia = [s for s in do_usuniecia if type(s).__name__ != typ]
        kandydat = GrupaSekcji(drzewo.nazwa, tuple(do_usuniecia))
        if kandydat.policz_wiersze() <= max_wierszy:
            return kandydat

    raise ValueError(
        f"Raport nie miesci sie w {max_wierszy} wierszach nawet po usunieciu "
        f"wszystkich sekcji opcjonalnych. Zmniejsz zakres danych albo "
        f"zwieksz limit - NIE generuj raportu, ktory 'wyjdzie jak wyjdzie'."
    )
```

**Uczciwe ograniczenie tego szkicu:** `dopasuj_do_strony` działa **na jednym poziomie**. Dla zagnieżdżonego drzewa trzeba rekurencji albo spłaszczenia. Zostawiam to jako rozszerzenie — ale zwracam uwagę na **wyjątek**: to jest ta sama filozofia, co `BladBudowy` w module 23. **Raport, który nie mieści się w limicie, nie jest raportem „troszkę za długim" — jest błędem do naprawienia** i lepiej go nie generować niż wysłać.

**Decyzja 5: testy, które nie wiedzą o openpyxl.**

```python
import pytest

def male_dane():
    return [("A", "PL", 1, 5.0), ("B", "DE", 2, 6.0)]           # suma 17

def duze_dane():
    return [("A", "PL", 100, 60.0), ("B", "DE", 50, 40.0)]      # suma 8000

def test_warunek_wylacza_sekcje_z_renderowania():
    md = podglad_markdown(zbuduj_raport_biznesowy(male_dane(), odbiorca="analityk"))
    assert "ponizej normy" in md
    assert "wysoka sprzedaz" not in md

def test_policz_wiersze_widzi_warunek():
    """To jest test, ktory NIE DALBY SIE NAPISAC, gdyby Sekcja pisala do arkusza."""
    male = zbuduj_raport_biznesowy(male_dane(), odbiorca="analityk")
    duze = zbuduj_raport_biznesowy(duze_dane(), odbiorca="analityk")
    assert male.policz_wiersze() < duze.policz_wiersze()

def test_widok_zarzadu_nie_ma_zalacznika():
    md = podglad_markdown(zbuduj_raport_biznesowy(duze_dane(), odbiorca="zarzad"))
    assert "Wnioski dla zarzadu" in md
    assert "Zalacznik: szczegoly" not in md

def test_limit_tabeli_jest_obcieciem_celu():
    drzewo = GrupaSekcji("X", (TabelaDanych(
        naglowki=("A",), wiersze=tuple((i,) for i in range(100)), _tytul="T"),))
    cel = CelMarkdown(max_wierszy_tabeli=5)
    drzewo.renderuj(cel)
    assert "pominięto 95 wierszy" in cel.tekst()

def test_raport_nie_miesci_sie_wybucha():
    """Fail fast: brak cichego generowania przydlugiego raportu."""
    drzewo = zbuduj_raport_biznesowy(duze_dane(), odbiorca="analityk")
    with pytest.raises(ValueError, match="nie miesci sie"):
        dopasuj_do_strony(drzewo, max_wierszy=3)
```

**Test, który jest tu najważniejszy:** `test_policz_wiersze_widzi_warunek`. On nie sprawdza funkcji — sprawdza **architekturę**. Jeżeli model jest dobrze zbudowany, ten test ma sens. Jeżeli gdzieś w środku jest `ws`, ten test nie da się napisać. **Test jako kontrola architektury** — to jest dokładnie ta rola, którą testy powinny pełnić, i o której mówi moduł 20.

**Czego tu świadomie nie robię:**

- **Nie robię z `dopasuj_do_strony` pełnego algorytmu plecakowego.** To zadanie optymalizacyjne i niepotrzebne. Prosty, czytelny, „odcinaj od najmniej ważnych" wystarcza w 95% przypadków.
- **Nie dodaję `WykresSlupkowy` do `CelMarkdown` jako obrazka.** W Markdownie zrobiłbym `podnaglowek` + notkę „(wykres w wersji XLSX)". To pokazuje kolejny raz: **ten sam model, różne medium, różne możliwości** — i cel decyduje, co zrobić, gdy czegoś nie da się pokazać.
- **Nie rozdzielam `odbiorca` na osobne klasy raportów.** Jeszcze nie. To jest dokładnie `Bridge`/`Strategy` — i należy do modułu 25, gdzie zobaczysz, jak z tych trzech `if odbiorca ==` zrobić wymienną strategię.

</details>

## 6. Typowe błędy i pułapki

**1. Objaw: „fasada działa, ale po trzech miesiącach ktoś w warstwie aplikacji pisze `raport._na_poziomie_openpyxl()['Dane']['A1'] = ...`, i cała warstwa się rozsypała."**
→ *Przyczyna:* **przeciek abstrakcji przez zwracanie `ws`.** Fasada (albo jej „klapa bezpieczeństwa") oddaje obiekt arkusza lub skoroszytu na zewnątrz. Pierwsze użycie jest usprawiedliwione („tylko jeden wykres", „tylko jedna formuła"), drugie już nie. Po tygodniu openpyxl jest z powrotem na wolności, a fasada służy jako „ten obiekt, przez który muszę przejść, żeby dostać `ws`".
→ *Naprawa:* **fasada nigdy nie zwraca `ws` w metodzie publicznej.** `_ws` prywatne, `_na_poziomie_openpyxl` — z nazwą i docstringiem, które krzyczą. Regułę egzekwuj **testem** (ćwiczenie 🟢) i **grepem** (sekcja 1.7). Diagnostyka: `grep -rn "\[[\"']A1[\"']\]" src/ | grep -v infrastructure` — jeżeli coś znajduje poza testami, masz przeciek.

**2. Objaw: „skrypt kończy się bez błędu, mówi『gotowe』, a raportu nie ma na dysku."**
→ *Przyczyna:* **dekorator, który połyka wyjątek zapisu.** `try: return funkcja(...) except Exception: return None` albo `log.warning(...); return`. Funkcja zgłasza sukces, choć nic nie zapisała. Ten błąd jest szczególnie paskudny, bo **pojawia się w każdym wywołaniu funkcji, nie tylko w błędnym** — więc wygląda jak „odporność na błędy", a jest **niemym trybem awarii**. Analogia z sekcji 1.3: papier, który po zdjęciu okazuje się pusty.
→ *Naprawa:* **dekorator nigdy nie połyka.** Dekorator może **dodać** (log, metrykę, walidację), ale **nie może odejmować** — ani wyniku, ani wyjątku. Jeżeli chcesz „nie wywalać całego procesu", łap wyjątek **w miejscu, w którym wiesz, co z nim zrobić** (na przykład zbierz do listy błędów i zwróć raport ze statusem), a nie w dekoratorze. I napisz test, który sprawdza, że plik **nie powstał** przy błędnych danych (przykład 3) oraz że wyjątek **poleciał** — nie tylko, że funkcja „zwróciła coś".

**3. Objaw: „po kilkuset wygenerowanych raportach proces zjada 3 GB pamięci i nie może otworzyć kolejnego pliku — a na Windows plik zostaje zablokowany."**
→ *Przyczyna:* **proxy, które nie zamyka `wb` w `read_only`.** W trybie `read_only` openpyxl trzyma **otwarty `ZipFile`**. Proxy leniwe (`LazyWorkbook`) może żyć długo, a nikt nie pamięta o `close()`. Gorzej: jeżeli proxy jest trzymane w jakiejś strukturze zbiorczej (cache, konfiguracja), uchwyt żyje przez cały proces. A na Windows otwarty uchwyt **blokuje plik przed usunięciem i nadpisaniem**.
→ *Naprawa:* **`close()` na proxy i jawny `__enter__`/`__exit__`** — nawet jeżeli wydaje się, że nie są potrzebne. Używaj `with`-a w każdym miejscu, gdzie proxy żyje dłużej niż chwilę. Rozważ `weakref.finalize`, jeżeli proxy może zostać porzucone. I pamiętaj: **`__getattr__` nie obsłuży `__exit__`** — musi być jawnie. Diagnostyka: `lsof`/`handle.exe` albo po prostu `leniwy.ile_razy_wczytano()` w logu — jeżeli proxy żyje długo, a metryka rośnie, masz wyciek.

**4. Objaw: „Composite działa, ale chcę przestawić kolejność sekcji i nie mogę. Chciałbym policzyć wiersze przed renderowaniem i nie mogę. Chciałbym wyrenderować ten sam raport do podglądu i nie mogę. Chciałbym wyłączyć sekcję, gdy brak danych, a sekcja już się częściowo wpisała."**
→ *Przyczyna:* **Composite, który zamiast budować model od razu pisze do arkusza.** To jest **jedna** przyczyna, z której wynikają cztery objawy. Sekcja w konstruktorze (albo w `renderuj`, który dostaje `ws` zamiast celu) wpisuje komórki. Wtedy „model" i „efekt" są tym samym, więc nie ma czego planować, przestawiać, testować ani wycofywać.
→ *Naprawa:* **Composite buduje model, renderowanie jest osobnym krokiem.** Klasa `Sekcja` nie może znać `ws`, `Workbook` ani `cell`. Renderowanie przechodzi drzewo i woła **cel** (`CelRaportu` — protokół). Jeżeli Twoja `Sekcja` ma pole `ws` albo argument `ws` w `__init__` — **masz ten błąd**. Diagnostyka: `grep -n "ws\|Workbook\|cell" sekcje.py` — powinno być pusto (poza docstringiem).

**5. Objaw: „fasada ma czterdzieści metod, każda z sześcioma argumentami, i nadal nie da się przez nią zrobić wykresu kołowego bez sięgnięcia do openpyxl."**
→ *Przyczyna:* **fasada, która nie jest fasadą, a przepakowaniem.** Kontrakt „bez openpyxl w sygnaturach" został spełniony formalnie, ale **nie spełniono poziomu 2 i 3**: fasada przyjmuje indeksy arkuszy, litery kolumn, parametry stylu jako stringi, a protokół wywołań jest przepisany wprost na kolejność operacji openpyxl. Miejsce, w którym miała żyć wiedza, **zostało puste** — więc wszyscy przychodzą z własną.
→ *Naprawa:* **zaostrz testy do poziomu 2 i 3.** Adresowanie po nazwie kolumny (nie po literze), brak `index=`, brak `ws.max_row` w sygnaturach, protokół egzekwowany wyjątkami (nie dokumentacją). Jeżeli fasada rośnie ponad dwadzieścia metod — podziel ją na **tematyczne fasady** (arkusze i treść / formatowanie / wykresy / druk), **a nie rozszerzaj jednej**. I zrób jedną rzecz konsekwentnie: **jeżeli jakiś przypadek naprawdę wymaga openpyxl, niech to będzie jawna klapa bezpieczeństwa z docstringiem, a nie dwudziesta pierwsza metoda z dwudziestoma argumentami.**

**6. Objaw: „kolega woła `adaptuj(dane, "Data", "Klient", "Region", "Ilosc", "Cena", "dd.mm.yyyy", "#,##0", ...)` i nie wie, który argument jest który."**
→ *Przyczyna:* **adapter z czterdziestoma argumentami.** Konfiguracja została rozsypana w sygnaturze zamiast skupiona w obiekcie. To ten sam błąd, co „magiczne liczby" z modułu 22, tylko w przebraniu listy argumentów — a jego objaw jest gorszy, bo **argumenty są pozycyjne i mają ten sam typ**: `float` 12, 24, 8 to szerokości czy formaty?
→ *Naprawa:* **specyfikacja kolumn jako obiekt.** `Kolumna("Data", 12, "data", "dd.mm.yyyy")` — nazwane pola, domyślne wartości, walidacja w `__post_init__`. Adapter przyjmuje `Sequence[Kolumna]` i **nic więcej**. Kryterium praktyczne: **jeżeli funkcja ma więcej niż cztery–pięć argumentów pozycyjnych tego samego typu, to znaczy, że brakuje obiektu.** Wtedy rozważ `Sequence[Kolumna]` albo `RaportSpec` (moduł 23).

**7. Objaw: „dekorator `@loguj_zapis` działa, ale po jego dodaniu plik wynikowy ma inną nazwę niż wcześniej, a przy zapisie z wyjątkiem i tak powstaje plik częściowy."**
→ *Przyczyna:* **decorator, który zmienia semantykę zapisu.** Dekorator miał **obserwować**, a zaczął **decydować**: dodał sufiks daty („bo wygodniej"), zapisał do innego katalogu („bo tak jest w firmie"), zalogował i **uznał zapis za udany** przed `os.replace`, albo — gorzej — zapisał plik tymczasowy i zostawił go przy wyjątku. To jest nie ta sama funkcja, tylko inna funkcja o tej samej nazwie.
→ *Naprawa:* **dekorator ma być przezroczysty w semantyce.** Obserwuje, mierzy, waliduje — ale **argumenty i wynik przekazuje bez zmian**. Jeżeli chcesz zmienić nazwę pliku, zmień nazwę **w wołającym** albo dodaj **jawny parametr**. I każde `try` w dekoratorze **musi mieć `except` z `raise` albo `finally`** — nigdy `except: pass`. Diagnostyka: napisz test, który woła funkcję z dekoratorem i bez, i sprawdza, że **wynik jest identyczny**. Jeżeli nie jest — dekorator zmienia semantykę.

**8. Objaw: „`LazyWorkbook` działa dla `sheetnames`, ale `len(leniwy)` rzuca `TypeError: object of type 'LazyWorkbook' has no len()`, a `with leniwy:` nie wywołuje `close()`."**
→ *Przyczyna:* **proxy, które traci metody specjalne.** `__getattr__` jest wywoływany **tylko wtedy, gdy normalne wyszukiwanie atrybutu zawiodło** — i **tylko dla atrybutów instancji**. Metody specjalne (`__len__`, `__iter__`, `__enter__`, `__getitem__`, `__contains__`) Python szuka **na typie**, czyli na klasie `LazyWorkbook`, nie na obiekcie docelowym. To ta sama reguła, o której mówi analogia z sekcji 1.4 (linia bezpośrednia do biurka). Dotyczy to **każdego** proxy — nie tylko naszego.
→ *Naprawa:* **deklaruj jawne przekazanie dla metod specjalnych, których potrzebujesz.** Minimum: `__enter__`, `__exit__`, `close`. Jeżeli potrzebujesz iteracji — `__iter__`. Jeżeli `len()` — `__len__`. Jeżeli indeksowania — `__getitem__`. Każda z nich ma jedną linijkę (`return self._zaladuj().__xyz__()`), ale **musi być napisana ręcznie**, bo `__getattr__` jej nie złapie. Alternatywnie użyj `__getattribute__` (przechwytuje wszystko), ale **płacisz za to wydajnością i ryzykiem rekurencji**.

**9. Objaw: „dodałem `@mierz_czas`, ale w raportach `pytest` wszystkie testy nazywają się `opakowanie`, a `monkeypatch` nie znajduje funkcji."**
→ *Przyczyna:* **brak `functools.wraps`.** Dekorator zwraca nową funkcję o nazwie `opakowanie` i bez docstringa. `pytest` używa `__name__` do nazywania testów, `monkeypatch`/`patch` do znajdowania atrybutów, `inspect` do sygnatur, a debugger do etykiet. Wszystkie te narzędzia są **bezużyteczne**, choć kod „działa".
→ *Naprawa:* **`@functools.wraps(funkcja)` na każdym opakowaniu.** Jedna linia. Diagnostyka: `print(opakowanie.__name__)` — powinno być nazwą oryginału. Natychmiastowy test: `assert f.__name__ == "oryginalna_nazwa"` na wybranej funkcji. **Uwaga dodatkowa:** `functools.wraps` kopiuje też `__wrapped__`, dzięki czemu `inspect.signature` **widzi oryginalną sygnaturę** — i Twój test jakości fasady z ćwiczenia 🟢 będzie działał poprawnie nawet przez warstwę dekoratorów.

**10. Objaw: „użyłem `LazyWorkbook` do szablonu, który i tak jest potrzebny w każdym uruchomieniu — i raporty generują się wolniej niż przedtem."**
→ *Przyczyna:* **proxy leniwe użyte tam, gdzie nic się nie oszczędza.** Leniwość ma sens **tylko wtedy**, gdy istnieje ścieżka, w której obiekt nie jest potrzebny. Jeżeli szablon jest potrzebny w 100% przypadków, płacisz narzut (sprawdzenie `if wb is None` przy każdym dostępie, dodatkowa warstwa w każdym wywołaniu, obowiązek `close()`) bez żadnej korzyści. **To samo dotyczy `write_only` w module 19**: świetne dla generowania od zera, katastrofalne dla modyfikacji istniejącego pliku.
→ *Naprawa:* **zmierz to i podejmij decyzję na podstawie liczby.** Dodaj metrykę: `leniwy.ile_razy_wczytano()` po uruchomieniu. Jeżeli po 1000 raportów wartość wynosi 1000 — **leniwość nie działa, usuń ją.** Jeżeli 300 — działa i zostaw. **Wzorce są narzędziami, nie deklaracjami ideowymi** (moduł 22). Diagnostyka: jeżeli optymalizacja nie ma **mierzonego** zysku, prawdopodobnie go nie ma.

**11. Objaw: „`eksportuj()` ma 240 linii, siedem zagnieżdżonych `if`-ów i trzy różne miejsca, w których robi `ws.cell(...)`."**
→ *Przyczyna:* **Front Controller jako god function.** Nazwa sugeruje „jedno wejście", a ktoś zaprojektował **jedną funkcję, która robi wszystko**. Kroki, które miały być **wywoływane**, zostały **wpisane w treści**. Objaw diagnostyczny: `eksportuj` zawiera nazwy openpyxl.
→ *Naprawa:* **funkcja WYWOŁUJE kroki, nie IMPLEMENTUJE ich.** Każdy krok to osobna funkcja albo metoda. Kryterium praktyczne: **jeżeli w `eksportuj` wystąpi słowo `openpyxl`, `ws` albo `cell` — krok został wpisany w treści.** Docelowa długość: 30–40 linii. Diagnostyka: `wc -l src/eksport.py` + `grep -c "if "`. Więcej niż pięć `if`-ów w jednej funkcji to sygnał, że część logiki jest nie na swoim miejscu.

**12. Objaw: „używam `RaportExcel`, ale żeby przenieść wykres, muszę znać współrzędne, i tak sięgam po dokumentację openpyxl."**
→ *Przyczyna:* **fasada, która nie przemyślała swojego słownictwa.** Fasada ma chronić przed pojęciami openpyxl — ale każde wywołanie, które wymaga od wołającego „zajrzenia do dokumentacji openpyxl", jest **przeciekiem**, nawet jeżeli w sygnaturze nie ma żadnego typu z openpyxl. Wykorzystanie `pozycja=(2, 6)` zamiast `"H2"` to poprawa (poziom 2), ale `(2, 6)` nadal wymaga rozumienia, że to wiersz i kolumna.
→ *Naprawa:* **przejrzyj API fasady i zapytaj: „jak wytłumaczyłbym to koledze, który nie zna Excela?"** Jeżeli odpowiedź wymaga słów „wiersz", „kolumna", „współrzędna", „kotwica" — jest przeciek pojęciowy. Wtedy zamień na słownictwo biznesowe: `gdzie="pod tabela"`, `strona="prawa"` i pozwól fasadzie wyliczyć pozycję. **To nie zawsze jest możliwe, ale jest to pytanie, które trzeba zadać** — a granica jest sprawą Twojej domeny. W naszym raporcie sprzedaży `pozycja=(wiersz, kolumna)` jest akceptowalna, bo raport ma siatkę i mówienie o wierszach jest naturalne. W innym systemie może nie być.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Fasada ukrywa nie tylko wywołania, ale i pojęcia.** Recepcja hotelu nie jest listą telefonów do działów — ma własne słownictwo (`pokój`, `rezerwacja`) i nie odpowiada na pytania o zawory hydrauliczne. Przetestuj swoją fasadę na **trzech poziomach**: (1) brak typów openpyxl w sygnaturach, (2) brak pojęć openpyxl w API (adres `A1`, litery kolumn, `index`, `max_row`), (3) wołający nie zna protokołu (fasada egzekwuje kolejność). Poziom 1 sprawdzisz `grep`-em w sekundę — i to jest jedyny poziom, który da się sprawdzić mechanicznie. Dwa pozostałe są Twoją pracą.

2. **Adapter tłumaczy kształt, nie znaczenie.** Przejściówka nie zmienia napięcia. Adapter buduje krotki wartości w kolejności kolumn — i **nie decyduje o tym, czy dane są sensowne**. Kompromis `getattr` z modułu 23 domykamy **`typing.Protocol`** (kontrakt statyczny) plus **`hasattr` w metodzie `sprawdz()`** (kontrola w locie z komunikatem, który mówi, czego brakuje i czego oczekiwano). I sprawdzaj **raz**, na pierwszym obiekcie — nie na każdym z 500 000.

3. **Dekorator owija, nie podmienia.** Papier zdobi prezent, torebka go niesie — po zdjęciu jest ten sam prezent. Trzy zastosowania: pomiar czasu (`finally`, więc mierzy też błędy), log zapisu (obserwuje, nie zmienia nazwy pliku) i walidacja **przed** wywołaniem (rzuca `BladWalidacjiDanych`, **nie połyka**). Kolejność ma znaczenie: walidacja najbliżej funkcji, pomiar na wierzchu. **Dekorator, który połyka wyjątek, nie czyni kodu odpornym — czyni go niemym.** I zawsze `@functools.wraps`.

4. **Proxy kontroluje dostęp, ale nie przechwyci linii bezpośredniej.** Sekretarka nie odbiera telefonów dzwoniących wprost na biurko — a `__getattr__` nie obsłuży metod specjalnych, bo Python szuka ich na typie. `LazyWorkbook` kupuje czas i pamięć przez odroczenie `load_workbook` **tylko wtedy, gdy istnieje ścieżka, w której plik nie jest potrzebny** — inaczej to czysty narzut. `BramaZapisu` zmienia proces (raport wymaga zatwierdzenia) i tworzy ślad audytowy. Oba muszą mieć **jawne `close()` i `__enter__`/`__exit__`**, bo w `read_only` openpyxl trzyma otwarty uchwyt.

5. **Composite buduje MODEL, a renderowanie to osobny krok — i to jest cała różnica między architekturą a improwizacją.** Spis treści tworzy się przed drukiem. Sekcja nie może znać `ws`. Dzięki temu możesz policzyć wiersze przed renderowaniem, wyłączyć sekcję bez śladu, przestawić rozdziały i **wyrenderować ten sam model do dwóch różnych celów** — do Excela i do Markdowna. Ten drugi cel nie jest ozdobą: **to dowód, że model naprawdę jest modelem**. A `eksportuj()` zamyka całość jako Front Controller, który kroki **wywołuje**, a nie **implementuje**.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. TEST JAKOSCI FASADY - trzy poziomy, trzy narzedzia
# ==================================================================
# Poziom 1 (mechaniczny, w sekundzie):
#     grep -n "def " src/raport.py | grep -E "Workbook|Worksheet|Cell|Font|Table|Chart"
#     -> musi byc PUSTE
#
# Poziom 1 (testem, na kazdym uruchomieniu CI):
import inspect

ZAKAZANE_NAZWY = {"ws", "wb", "worksheet", "workbook", "cell"}
ZAKAZANE_TYPY = {"Workbook", "Worksheet", "Cell", "Font", "PatternFill",
                 "Alignment", "Border", "Table", "BarChart", "Reference"}

def test_fasada_nie_wypuszcza_openpyxl(fasada_cls):
    for nazwa, metoda in inspect.getmembers(fasada_cls, inspect.isfunction):
        if nazwa.startswith("_"):
            continue
        for p in inspect.signature(metoda).parameters.values():
            assert p.name not in ZAKAZANE_NAMES, f"{nazwa}: {p.name}"
            # UWAGA: `from __future__ import annotations` -> adnotacje sa STRINGAMI
            assert not any(t in str(p.annotation) for t in ZAKAZANE_TYPY)

# Poziom 2 (recenzja): brak A1, liter kolumn, index=, ws.max_row w API
# Poziom 3 (recenzja): wolajacy nie musi znac KOLEJNOSCI operacji


# ==================================================================
# 2. FASADA - szkielet
# ==================================================================
class RaportExcel:
    """Fasada. Trzyma Workbook PRYWATNIE. Nie zwraca ws. Egzekwuje protokol."""

    def __init__(self, tytul: str, autor: str = "system") -> None:
        self._wb = Workbook()
        del self._wb["Sheet"]
        self._wb.properties.title = tytul
        self._wb.properties.creator = autor
        # wlasny stan - zeby NIE czytac ws.max_row i moc egzekwowac protokol
        self._nazwy: dict[str, str] = {}
        self._kolumny: dict[str, tuple[str, ...]] = {}
        self._pierwszy_danych: dict[str, int] = {}
        self._ostatni: dict[str, int] = {}
        self._tablice: set[str] = set()

    def _ws(self, klucz: str):
        """PRYWATNE. Nigdy nie wychodzi na zewnatrz."""
        nazwa = self._nazwy.get(klucz)
        if nazwa is None:
            raise BladRaportu(f"Nie ma arkusza {klucz!r}. Sa: {sorted(self._nazwy)}")
        return self._wb[nazwa]

    def arkusz(self, klucz, tytul=None, *, tab_color=None,
               freeze_wiersz=0, siatka=True) -> None: ...
    def naglowek(self, klucz, etykiety, *, tlo="1072BA") -> None:
        if klucz in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} ma juz naglowek")
        ...
    def wiersze(self, klucz, dane, *, formaty=None) -> int:
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        ...
    def tabela(self, klucz, nazwa, *, styl="TableStyleMedium9") -> None:
        # ref = zakres OD WIERSZA NAGLOWKA do ostatniego (headerRowCount=1)
        ...
    def para(self, klucz, etykieta, wartosc, format_liczby=None) -> None: ...
    def tytul(self, klucz, tekst, *, rozmiar=16) -> None: ...
    def szerokosci(self, klucz, szerokosci) -> None: ...
    def wykres_slupkowy(self, klucz, tytul, *, kolumna_kategorii,
                        kolumny_wartosci, pozycja=(2, 8),
                        szerokosc_cm=14.0, wysokosc_cm=8.0) -> None:
        # kolumny PO NAZWIE -> indeks przez naglowki.index(nazwa) + 1
        ...

    def do_bajtow(self) -> bytes:
        from io import BytesIO
        bufor = BytesIO(); self._wb.save(bufor); return bufor.getvalue()

    def zapisz(self, cel) -> Path:
        cel = Path(cel); cel.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=cel.parent); os.close(fd)
        try:
            self._wb.save(tmp); os.replace(tmp, cel)
        except BaseException:
            Path(tmp).unlink(missing_ok=True); raise
        return cel

    def _na_poziomie_openpyxl(self):
        """KLAPA BEZPIECZENSTWA. Tylko infrastruktura i TESTY. Nigdy aplikacja."""
        return self._wb


# ==================================================================
# 3. ADAPTER - dwa ksztalty, jeden interfejs
# ==================================================================
class PozycjaRaportu(Protocol):        # kontrakt STATYCZNY
    data: date; klient: str; region: str; ilosc: int; cena: float

class AdapterAtrybutowy:
    def __init__(self, kolumny: tuple[Kolumna, ...]) -> None: ...
    @property
    def naglowki(self) -> tuple[str, ...]: ...
    @property
    def formaty(self) -> tuple[str | None, ...]: ...

    def sprawdz(self, obiekt) -> None:
        """JAWNA kontrola - isinstance na Protocol z polami jest zawodne."""
        brakujace = [k.pole for k in self._kolumny if not hasattr(obiekt, k.pole)]
        if brakujace:
            raise BladAdaptera(
                f"Obiekt {type(obiekt).__name__} nie ma pol: {brakujace}. "
                f"Wymagane: {[k.pole for k in self._kolumny]}"
            )

    def wiersze(self, dane):                     # GENERATOR, nie lista!
        for i, obiekt in enumerate(dane):
            if i == 0:
                self.sprawdz(obiekt)             # raz, nie na kazdym
            yield tuple(getattr(obiekt, k.pole) for k in self._kolumny)


def wybierz_adapter(dane, pierwszy=None):
    """Rozpoznaj KSZTALT, nie typ listy."""
    if pierwszy is None:
        dane = list(dane)
        if not dane: raise BladAdaptera("Brak danych")
        pierwszy = dane[0]
    if isinstance(pierwszy, dict):               # CSV
        adapter = AdapterTabelaryczny(KOLUMNY_CSV)
        adapter.sprawdz_naglowki(set(pierwszy))
        return adapter, dane
    if hasattr(pierwszy, "klient") and hasattr(pierwszy, "cena"):
        return AdapterAtrybutowy(), dane
    raise BladAdaptera(f"Nieznany ksztalt: {type(pierwszy).__name__}")


# ==================================================================
# 4. DEKORATORY - owijaja, NIE podmieniaja
# ==================================================================
def mierz_czas(f):
    @functools.wraps(f)                              # ← ZAWSZE
    def opakowanie(*args, **kwargs):
        start = time.perf_counter()
        try:
            return f(*args, **kwargs)
        finally:                                     # ← mierzy tez bledy
            METRYKI.append((f.__qualname__, time.perf_counter() - start))
    return opakowanie

def waliduj_dane(walidator):                         # dekorator z ARGUMENTEM
    def dekorator(f):
        @functools.wraps(f)
        def opakowanie(dane, *args, **kwargs):
            problemy = walidator(dane)
            if problemy:
                raise BladWalidacjiDanych(problemy)   # ← RZUCA, nie polyka!
            return f(dane, *args, **kwargs)
        return opakowanie
    return dekorator

def audytowany(*metody):                             # dekorator KLASY
    def dekorator(cls):
        for nazwa in metody:
            oryginal = getattr(cls, nazwa)
            @functools.wraps(oryginal)
            # DOMYSLNE ARGUMENTY = zamrozenie wartosci z petli (late binding!)
            def opakowanie(self, *a, __o=oryginal, __n=nazwa, **kw):
                log.info("wywolanie: %s", __n)
                return __o(self, *a, **kw)
            setattr(cls, nazwa, opakowanie)
        return cls
    return dekorator

# KOLEJNOSC: walidacja NAJBLIZEJ funkcji, pomiar na wierzchu:
#     @mierz_czas
#     @waliduj_dane(walidator)
#     def generuj_raport(dane, cel): ...

# ANTYWZORZ - NIGDY:
#     except Exception: return None      # <- niemy tryb awarii


# ==================================================================
# 5. PROXY - leniwe i chronione
# ==================================================================
class LazyWorkbook:
    def __init__(self, sciezka):
        object.__setattr__(self, "_sciezka", Path(sciezka))   # ← omija __getattr__
        object.__setattr__(self, "_wb", None)
        object.__setattr__(self, "_pobrania", 0)

    def _zaladuj(self):
        wb = object.__getattribute__(self, "_wb")
        if wb is None:
            wb = load_workbook(object.__getattribute__(self, "_sciezka"), read_only=True)
            object.__setattr__(self, "_wb", wb)
        return wb

    def __getattr__(self, nazwa):        # TYLKO nieznalezione atrybuty
        return getattr(self._zaladuj(), nazwa)

    # METODY SPECJALNE - JAWNIE (__getattr__ ich NIE obsluzy!)
    def close(self) -> None: ...
    def __enter__(self): return self
    def __exit__(self, *e): self.close(); return False

class BramaZapisu:
    def __init__(self, fasada) -> None:
        self._fasada, self._zatwierdzone, self._historia = fasada, False, []

    def zapisz(self, cel):
        if not self._zatwierdzone:
            raise BladZapisu("Zapis zablokowany. Najpierw zatwierdz(kto).")
        cel = self._fasada.zapisz(cel)
        self._historia.append(("zapis", str(cel))); return cel

    def zatwierdz(self, kto) -> None: ...
    def __getattr__(self, nazwa):        # czytanie przechodzi swobodnie
        return getattr(self._fasada, nazwa)


# ==================================================================
# 6. COMPOSITE - model, NIE arkusz
# ==================================================================
class CelRaportu(Protocol):              # most miedzy drzewem a medium
    def podnaglowek(self, tekst: str, poziom: int) -> None: ...
    def akapit(self, tekst: str) -> None: ...
    def tabela(self, naglowki, wiersze, formaty=None) -> None: ...
    def para(self, etykieta, wartosc, format_liczby=None) -> None: ...

class Sekcja(ABC):
    @abstractmethod
    def tytul(self) -> str: ...
    @abstractmethod
    def renderuj(self, cel: CelRaportu, poziom: int = 1) -> None: ...
    @abstractmethod
    def policz_wiersze(self) -> int: ...      # ← PLANOWANIE widzi warunki

@dataclass(frozen=True)
class Akapit(Sekcja):        teksty: tuple[str, ...]
@dataclass(frozen=True)
class ParaKPI(Sekcja):       etykieta: str; wartosc: object; format_liczby=None
@dataclass(frozen=True)
class TabelaDanych(Sekcja):  naglowki: tuple; wiersze: tuple; formaty=None

@dataclass(frozen=True)
class GrupaSekcji(Sekcja):   # GALEZ - a sama JEST sekcja
    nazwa: str
    sekcje: tuple[Sekcja, ...] = ()

    def renderuj(self, cel, poziom=1):
        cel.podnaglowek(self.nazwa, poziom)
        for s in self.sekcje:
            if isinstance(s, GrupaSekcji):
                s.renderuj(cel, poziom + 1)          # ← REKURENCJA
            else:
                s.renderuj(cel)

    def policz_wiersze(self) -> int:
        return 1 + sum(s.policz_wiersze() for s in self.sekcje)

@dataclass(frozen=True)
class WarunkowaSekcja(Sekcja):
    warunek: bool
    sekcja: Sekcja
    def renderuj(self, cel, poziom=1):
        if not self.warunek: return                  # ZERO wywolan do celu
        ...
    def policz_wiersze(self) -> int:
        return self.sekcja.policz_wiersze() if self.warunek else 0   # ← KLUCZ!

# DWA CELE: ten sam model, rozne media
class CelExcel:       # deleguje do RaportExcel
class CelMarkdown:    # list[str] -> tekst; ZERO openpyxl -> TESTY bez plikow

# ANTYWZORZ (NIGDY):
#     class SekcjaWynikow:
#         def __init__(self, ws, tytul, dane):
#             ws.cell(row=1, column=1, value=tytul)   # <- pisze od razu!


# ==================================================================
# 7. FRONT CONTROLLER - spis krokow, nie ich tresc
# ==================================================================
@mierz_czas
def eksportuj(dane, cel, *, podglad_markdown=False, zatwierdzenie=None):
    adapter, rekordy = wybierz_adapter(dane)          # 1
    _sprawdz_surowe(list(rekordy))                    # 2
    drzewo = zbuduj_raport(list(rekordy))             # 3 - MODEL
    raport = RaportExcel("Raport"); raport.arkusz("raport")   # 4
    drzewo.renderuj(CelExcel(raport, "raport"))
    brama = BramaZapisu(raport)                       # 5
    if zatwierdzenie: brama.zatwierdz(zatwierdzenie)
    return WynikEksportu(cel=brama.zapisz(cel), ...)  # 6 - DTO, nie Workbook

# KRYTERIUM: jezeli w eksportuj wystepuje slowo openpyxl|ws|cell
#            -> krok zostal WPISANY w tresc, a nie wywolany. Split.


# ==================================================================
# 8. PULAPKI DO ZAPAMIETANIA
# ==================================================================
# - fasada zwracajaca ws -> cala warstwa sie rozsypuje (grep A1!)
# - dekorator z `except: return None` -> niemy tryb awarii
# - proxy bez jawnego close/__enter__ -> wyciek + blokada pliku na Windows
# - composite piszacy do arkusza -> nie policzysz, nie przestawisz, nie przetestujesz
# - brak functools.wraps -> pytest/monkeypatch/debugger slepna
# - domkniecie w petli BEZ domyslnego argumentu -> late binding (wszystkie == ostatnia)
# - __getattr__ NIE lapie metod specjalnych (len, with, iter) - jawne deklaracje!
# - object.__setattr__/__getattribute__ chroni przed nieskonczona rekurencja
# - lazy uzywane tam, gdzie obiekt jest zawsze potrzebny -> czysty narzut
# - Table.ref ZAWIERA wiersz naglowka (headerRowCount=1)
# - displayName tabeli BEZ spacji i znakow specjalnych
# - wykres bez set_categories -> os X pokaze 1,2,3 zamiast nazw
# - Reference poza zakresem danych -> pusty wykres, BEZ bledu
# - zapis atomowy: mkstemp + os.replace (NIGDY tempfile.mktemp)
# - testy MAJA PRAWO znac openpyxl. Produkcja - nie.
# - config late binding: nie ma, ale `from __future__ import annotations`
#   czyni adnotacje STRINGAMI -> test jakości fasady musi uzywac str()
```

## 9. Co dalej

Zatrzymaj się i policz, co masz. To jest **największy moduł kursu**, bo zbiera warstwę, którą zapowiedziały moduły 22 i 23.

Masz **fasadę `RaportExcel`**, która trzyma `Workbook` w środku, egzekwuje protokół wyjątkami (`BladProtokolu`, gdy tabela przed nagłówkiem), adresuje kolumny **po nazwie**, liczy pozycje wykresu z nazw, nie z liter, i zapisuje atomowo. Masz **trzy poziomy testu jakości** — i wiesz, że mechaniczny `grep` sprawdza tylko pierwszy z nich. Masz **adaptery** z protokołem `PozycjaRaportu` i metodą `sprawdz()`, która mówi dokładnie, czego brakuje. Masz **dekoratory**, które owijają i nie podmieniają, z `functools.wraps`, kolejnością („walidacja najbliżej funkcji"), i twardą zasadą: **dekorator nigdy nie połyka wyjątku**. Masz **proxy leniwe** z jawnym `close()` (bo `__getattr__` nie obsłuży `__exit__`) i **proxy chroniące** z historią zatwierdzeń. I masz **Composite**, które buduje model, policzy wiersze **widząc warunki**, da się przestawić i **wyrenderować do dwóch różnych mediów** — a ten drugi render jest dowodem, że model naprawdę jest modelem.

I jest jedna rzecz, którą chcę, żebyś zapamiętał z tego modułu ponad wszystko, bo jest najczęstszym błędem w tej warstwie:

> **Model to nie efekt.** Spis treści tworzy się **przed** drukiem. Sekcja raportu nie może znać `ws`. Jeżeli `Twoja sekcja` ma pole `ws` albo argument `ws` w konstruktorze — nie masz Composite, a masz generator z nazwami klas. Wtedy nie policzysz wierszy, nie wyłączysz warunku bez śladu, nie przestawisz rozdziałów i nie napiszesz testu, który trwa mikrosekundy.

W module 25 wchodzimy w **wzorce behawioralne** — i robimy to na **dokładnie tym samym przykładzie**. Tu zostały dwa otwarte wątki, oba celowo:

1. **`if odbiorca == "zarzad"` w `zbuduj_raport_biznesowy`.** W przykładzie 5 i w ćwiczeniu 🔴 powtarza się kilka razy. **To jest zapowiedź Strategy** — zobaczysz, jak z tych rozgałęzień zrobić wymienny obiekt, który „wie, co pokazać zarządowi".
2. **`dopasuj_do_strony` z priorytetami jako listą krotek.** To jest zapowiedź **Template Method** — zobaczysz `BazowyRaport.generuj()` jako szkielet (przygotuj → nagłówek → dane → podsumowanie → formatuj → zapisz) z haczykami, które podklasa wypełnia treścią.

A oprócz tego zobaczysz:

- **Command** — operacje edycyjne jako obiekty (`KomendaUstawKomorke`, `KomendaDodajWiersz`) i **stos komend z `undo()`**. To największa wartość tego modułu: po raz pierwszy w kursie zobaczysz **cofanie zmian** i pełny **log audytowy „kto i co zmienił"**, zbudowany na dokładnie tych operacjach, które piszesz od modułu 04.
- **Visitor** — „kontroler z checklistą chodzący po hali". Ten sam szkielet przejścia przez arkusz, różne raporty: znajdź puste kolumny, wykryj formuły, zbierz metryki, zanonimizuj dane.
- **Specification** — reguły deklaratywne, które łączą się przez `and_`, `or_`, `not_`. I tu jest elegancki most, którego jeszcze nie było: **te same obiekty reguł filtrują dane przed zapisem I generują reguły formatowania warunkowego** z modułów 13–14. Jedna definicja „duża sprzedaż" dla filtra i dla koloru.
- **Chain of Responsibility** — łańcuch transformacji wiersza: oczyszczenie → walidacja → normalizacja → konwersja typów. Naturalne uzupełnienie adapterów z tego modułu.

I jeszcze jedno, co robimy w module 25, a co jest bezpośrednio związane z tym modułem: **`Command` z `undo()` jest praktycznym dowodem na to, że `BramaZapisu` miała sens.** Bo zapis po zatwierdzeniu ma sens tylko wtedy, gdy zmiany są **jeszcze w pamięci** — a więc gdy fasada trzyma `Workbook` prywatnie, model jest osobno, a komendy można cofać. **Cała ta warstwa, którą dziś zbudowałeś, istnieje właśnie po to, żeby moduł 25 miał na czym pracować.**

Zanim tam pójdziesz, zrób dwie rzeczy na **swoim** kodzie, nie na ćwiczeniach:

1. Uruchom `grep -rn "openpyxl" src/ --include="*.py" | grep -v "/infrastructure/"`. **Zapisz wynik.** Jeżeli coś znalazło — to nie jest porażka, to jest **pierwszy przeciek, który widzisz świadomie**. Zanotuj, co by trzeba było zrobić, żeby go usunąć, i oceń, czy warto. Czasem nie warto. Ale **musisz wiedzieć, że tam jest**.
2. Weź dowolną swoją funkcję, która generuje raport, i policz jej argumenty. Jeżeli jest ich więcej niż pięć i mają ten sam typ — **to jest Twoje ćwiczenie 🟡 w wersji produkcyjnej**. Nie musisz od razu przepisywać wszystkiego: wystarczy, że zobaczysz, ile wiedzy o raporcie siedzi w sygnaturze zamiast w obiekcie.