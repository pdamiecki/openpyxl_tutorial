# Moduł 23 — Wzorce konstrukcyjne: tworzenie obiektów tanio i kontrolowanie

> **Część:** IV — Wzorce projektowe i architektura · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–22

## 0. W tym module nauczysz się

- **Zrozumiesz, dlaczego tworzenie obiektów to osobny problem** — i że „zwykłe" `wb.create_sheet()` rozsiane po dwudziestu miejscach to nie niewinna duplikacja, a rozproszenie wiedzy, którą trudno zmienić i nie da się przetestować.
- **Opanujesz fabrykę (`ArkuszFabryka`, `StylFabryka`)** jako jedno miejsce, w którym żyje konwencja wyglądu raportu — i odróżnisz Factory Method od Abstract Factory, żeby wiedzieć, kiedy ta druga jest przesadą.
- **Zbudujesz `RaportBuilder` z płynnym API i zobaczysz, dlaczego Builder *musi* odkładać materializację** — ten, który pisze do arkusza w trakcie budowy, traci cały sens istnienia.
- **Nauczysz się wzorca Prototype na szablonach skoroszytów** — pieczątka firmowa, nadruk gotowy, dopisujesz treść. Zobaczysz twardą regułę *„szablon nigdy nie jest nadpisywany"* i dowiesz się eksperymentalnie, że `copy_worksheet` **gubi obrazy i wykresy**.
- **Zbudujesz `StyleRegistry` z cache i zrozumiesz Flyweight** — dlaczego `Font` tworzony raz i użyty w 20 000 komórek daje ten sam plik, ale o wiele mniej pracy dla Pythona. Zmierzysz to sam.
- **Poznasz pułapkę Singletonu** — globalny rejestr wygląda elegancko, dopóki nie spróbujesz napisać testu. Nauczysz się **wstrzykiwać rejestr** zamiast go importować.
- **Domykasz model kreacji przez konfigurację i DTO**: `RaportSpec` opisujący raport jako dane oraz `WierszRaportu` jako `dataclass` z walidacją w konstruktorze — to ostatni klocek przed wzorcami strukturalnymi z modułu 24.

## 1. Intuicja i analogia

### 1.1. Zakład produkcyjny

Wyobraź sobie **zakład produkcyjny**. Nie rzemieślnika w warsztacie, który robi jeden egzemplarz od zera, rzeźbiąc każdy element osobnym narzędziem i mierząc „na oko". Prawdziwy zakład produkcyjny — taki, który robi dziesięć tysięcy sztuk.

Co sprawia, że ta produkcja jest szybka i **powtarzalna**? Nie talent pracowników. Cztery rzeczy:

1. **Formy i foremki.** Nie rzeźbisz za każdym razem tego samego kształtu. Wlewace materiał w formę, a forma jest zawsze taka sama. Jeżeli trzeba poprawić kształt, **poprawia się formę, a nie każdy egzemplarz**.
2. **Instrukcja montażu.** Rzeczy złożone z wielu części mają **ustaloną kolejność kroków** i na końcu **kontrolę kompletności**. Nie da się złożyć obudowy po zamknięciu panelu — i instrukcja to wie, a robotnik po dziesięciu godzinach już nie.
3. **Wzorzec do odciskania.** Pieczątka, matryca, stempel. Nadruk jest gotowy; Ty dopisujesz tylko treść zmienną. Wzorzec **nigdy się nie zużywa**, bo odciskasz z niego, a nie po nim piszesz.
4. **Magazyn zamiast produkcji na zapas.** Element, który produkujesz raz i używasz tysiąc razy (**matryca drukarska**), trzymasz gotowy na półce. Element, który jest drogi albo rzadko potrzebny, robisz **dopiero gdy zamówienie jest potwierdzone**.

I teraz najważniejsze zdanie tego modułu:

> **Wzorce konstrukcyjne to cztery elementy tego zakładu: foremka, instrukcja, wzorzec i magazyn.**

Bo widzisz, problem, który rozwiązują, nie brzmi „obiektów jest za dużo". Problem brzmi: **tworzenie obiektu to decyzja, która rozprasza wiedzę, kosztuje czas i potrafi zostawić obiekt w złym stanie.** Zajmijmy się kolejno tymi trzema rzeczami, bo to one są treścią tego modułu.

### 1.2. Trzy koszty tworzenia obiektu

**Koszt 1: rozproszenie wiedzy.**

Zobacz, ile wiedzy o firmie mieści się w tym fragmencie:

```python
dane = wb.create_sheet("Dane", 0)
dane.sheet_properties.tabColor = "1072BA"
dane.freeze_panes = "A4"
dane.sheet_view.showGridLines = False
dane.column_dimensions["A"].width = 12
dane.column_dimensions["B"].width = 24
```

Cztery **decyzje o firmie** i dwie **decyzje o technice**:
- „arkusz danych jest pierwszy" (decyzja firmy),
- „arkusz danych ma niebieską zakładkę `1072BA`" (decyzja firmy),
- „linie siatki są wyłączone w arkuszach danych" (konwencja firmy),
- „kolumna A ma 12, B ma 24" (konwencja firmy),
- „`create_sheet` przyjmuje index jako drugi argument" (technika — tego nie da się zmienić),
- „`tabColor` siada w `sheet_properties`" (technika).

Teraz wyobraź sobie, że te sześć linii występuje w **dwudziestu miejscach**, bo każdy raport je ma. I przychodzi szef: „zmieniamy niebieski na granatowy, wszystkie arkusze danych".

Ile masz do zrobienia? **Nie wiesz.** I to jest pierwszy koszt. Nie „dwadzieścia edycji". **Niepewność.** Musisz znaleźć wszystkie miejsca, w których powstaje arkusz danych, i nie masz żadnego sposobu, żeby wiedzieć, czy któreś nie jest ukryte w `if`-ie, w funkcji w innym pliku, albo w kopii wklejonej przez kogoś „na chwilę".

Analogia: to jakby w zakładzie produkcyjnym każdy pracownik miał **własną foremkę**, zrobioną według własnego pomysłu. Dopóki nikt nie patrzy, wygląda to samo. Ale gdy zmieniasz projekt, musisz obejść halę i przekonać każdego z osobna — a przynajmniej znaleźć wszystkie foremki.

**Koszt 2: czas i pamięć.**

Tu niespodzianka, bo intuicja podpowiada, że to jest główne zmartwienie wzorców. A nie jest. Ale jest realne.

W module 09 mówiliśmy, że style są niemutowalne: nie poprawiasz nadrukowanej etykiety, przyklejasz nową. No i właśnie — **jeżeli przyklejasz nową etykietę 20 000 razy, to tworzysz 20 000 identycznych etykiet.** Python nie wie, że są identyczne. Zobacz:

```python
for i in range(1, 20_001):
    ws.cell(row=i, column=1, value=i).font = Font(bold=True, color="FFC000", size=11)
```

To tworzy **20 000 obiektów `Font`**. Każdy z nich zajmuje pamięć, każdy musi zostać policzony (`__hash__`), żeby openpyxl mógł sprawdzić, czy taki styl już zna. Efekt uboczny jest pouczający: **plik wyjściowy będzie dokładnie taki sam** jak przy użyciu jednego współdzielonego obiektu, bo openpyxl przy zapisie deduplikuje style w `styles.xml`. Czyli marnujesz pamięć i czas Pythona, a w pliku tego nie widać. Zmierzysz to w przykładzie 4.

**Koszt 3: obiekt w złym stanie.**

To najciekawszy koszt i najczęstsza przyczyna istnienia Buildera.

Zobacz, jak łatwo stworzyć raport, który „wygląda skończony", a nie jest:

```python
def zrob_raport(cel):
    wb = Workbook()
    ws = wb.create_sheet("Dane")
    ws.append(["Data", "Klient", "Kwota"])       # naglowek
    zapisz_wiersze(ws, dane)                     # dane
    wb.save(cel)                                 # ← zapisujemy!
```

Gdzie tu jest `Podsumowanie`? Nie ma — bo ktoś zapomniał wywołać funkcję, która je dodaje. I **nic tego nie złapie.** Nie ma `None`, nie ma wyjątku, nie ma zera w komórce. Jest po prostu raport bez podsumowania, który trafia do księgowości.

Analogia: to jak próba złożenia szafy bez ostatniej kontroli kompletności. Wszystkie śrubki są wkręcone, drzwi się otwierają... i dopiero przy przeprowadzce okazuje się, że **brakuje dwóch tylnych paneli usztywniających**. Nie widzisz ich, dopóki nie zdarzy się coś, czego nie przewidziałeś.

Właśnie dlatego Builder kończy się metodą `build()`, która **weryfikuje kompletność**, zanim cokolwiek trafi do pliku. To jest analogia **instrukcji montażu mebla z IKEA** z zarysu tego modułu: kroki są ponumerowane, kolejność ma znaczenie, a na końcu jest lista kontrolna „czy nic nie zostało".

### 1.3. Cztery analogie, które zostaną z Tobą

Każdy wzorzec z tego modułu ma jedną dobrą analogię. Dobrze — jedna dobra jest warta trzy słabe.

| Wzorzec | Analogia | Sedno |
|---|---|---|
| **Factory** | **foremka** w zakładzie | Poprawiasz formę, a nie każdy odlew. Konwencja w jednym miejscu |
| **Builder** | **instrukcja montażu IKEA** | Kolejność i kompletność mają znaczenie; `build()` = „sprawdź, czy nic nie zostało" |
| **Prototype** | **pieczątka firmowa** | Nadruk gotowy, dopisujesz treść. **Wzorzec nigdy nie jest nadpisywany** |
| **Registry / Flyweight** | **drukarnia z gotowymi matrycami** | Nie rzeźbisz czcionki dla każdego wyrazu; matryca leży na półce i czeka |
| **Konfiguracja** | **przepis vs gotowe danie** | Przepis mówi *co*; kucharz wie *jak* |

Zauważ, że brakuje jednej: **Singleton**. I słusznie — bo Singleton nie jest dobrym wzorcem do opisania analogią z zakładu produkcyjnego. Singleton to **jedna wspólna tablica na hali, na której wszyscy piszą kredą**. Genialna i straszna jednocześnie: każdy ma dostęp, ale **nikt nie pamięta, kto co dopisał**, a gdy dwóch pracowników pisze równocześnie, tablica kłamie. Dlatego w tym module pokażę Ci rejestr stylów i **od razu odradzę** robienie z niego Singletona. To celowe.

### 1.4. Gdzie to wszystko stoi wobec modułu 22

W module 22 zbudowałeś `RaportXlsx` i `Specyfikacja` — i twierdziłeś, że nowy raport to nowe **dane**, a nie nowy `elif`. To była prawda. Ale przyznaj się: `Specyfikacja` opisywała **kolumny i reguły**. Nie opisywała tego, co dzieje się, gdy **arkuszy jest siedem** i każdy powstaje w innej funkcji.

Ten moduł zamyka tę lukę. Możesz na niego patrzeć tak:

- **moduł 22** — *co* jest w raporcie (dane, reguły, kolumny),
- **moduł 23** — *jak powstają obiekty*, z których raport jest zrobiony (arkusze, style, struktura),
- **moduł 24** — *jak nie wypuścić openpyxl na zewnątrz* (elegancka warstwa).

Wracamy tu do trzech sygnałów z sekcji 2.1 modułu 22, które wtedy tylko nazwaliśmy:

| Sygnał z modułu 22 | Wzorzec z tego modułu, który go zamyka |
|---|---|
| powtórzone importy i bloki formatowania (30×) | **Factory** — jedna fabryka arkuszy i stylów |
| literały bez nazw: kolory w 40 miejscach | **Registry / Flyweight** — rejestr stylów + motyw |
| magiczne indeksy i rozjechana kolejność | **Factory + Builder** — indeks wynika z definicji, kolejność z instrukcji |

## 2. Teoria

### 2.1. Fabryka: czym jest i kiedy wygrywa

**Intuicja.** Potrzebujesz nowego arkusza. Zamiast mówić „stwórz arkusz, nazwij go, daj mu index, ustaw kolor zakładki, wyłącz linie siatki, ustaw sześć szerokości" — mówisz **„daj mi arkusz danych"**. Cała wiedza, co to znaczy „arkusz danych", żyje w **jednym miejscu**.

**Analogia.** Foremka. Nie opisujesz kształtu za każdym razem; wkładasz materiał w formę.

**Definicja.** Wzorzec **Factory Method** to sposób na tworzenie obiektów, w którym **decyzja o tym, co dokładnie powstać, jest oddzielona od miejsca użycia**. Wołający mówi *jakiego rodzaju* obiekt chce, a nie *jak* go zbudować.

```python
# WERSJA BEZ FABRYKI - wiedza rozsiana
ws = wb.create_sheet("Dane", 0)
ws.sheet_properties.tabColor = "1072BA"
ws.freeze_panes = "A4"
ws.sheet_view.showGridLines = False
# + 5 linii szerokosci

# WERSJA Z FABRYKA - wiedza w jednym miejscu
ws = fabryka.utworz(wb, "dane")
```

**Konsekwencja praktyczna.** Zmiana koloru zakładki arkuszy danych to zmiana **jednej linii w definicji**. Zmiana konwencji „linie siatki wyłączone" obowiązuje we wszystkich raportach **natychmiast**. A co ważniejsze: `fabryka.utworz(wb, "dane")` jest **testowalne** — możesz sprawdzić w teście, że arkusz o kluczu `"dane"` ma wyłączone linie siatki i szerokość kolumny A równą 12, **nie otwierając ani jednego pliku wyjściowego** w Excelu.

**Factory Method vs Abstract Factory.** Ta różnica bywa rozdmuchana, więc podam ją krótko i praktycznie:

| | Factory Method | Abstract Factory |
|---|---|---|
| Co produkuje | **jeden** produkt (albo jeden z wariantów) | **rodzinę zgodnych** produktów |
| Klasyczny przykład | `fabryka.utworz(wb, "dane")` | `FabrykaRaportuStandardowego` produkująca *pasujące do siebie* (arkusz danych, arkusz podsumowania, styl) |
| Kiedy ma sens | prawie zawsze | gdy istnieją **całe rodziny**, które nie mogą się mieszać |

I teraz szczera rada, w duchu modułu 22: **w automatyzacji raportów potrzebujesz Factory Method.** Abstract Factory ma sens, gdy istnieje ryzyko pomieszania elementów dwóch różnych rodzin — na przykład „arkusz standardowy z motywem audytowym". W praktyce zamiast dwóch rodzin fabryk używa się **jednego motywu**, który jest przekazywany do fabryki. Robimy tak w przykładzie 1 i mówię wprost: `ArkuszFabryka` + `StylFabryka` **razem** pełnią rolę lekkiej Abstract Factory, bo produkują spójną parę (arkusz + pasujący styl) — ale implementujemy to jako dwie małe klasy z jednym źródłem motywu, a nie jako hierarchię fabryk.

**Uczciwa granica.** Fabryka jest przesadą, gdy tworzysz obiekt **raz** i nigdzie indziej. `wb = Workbook()` na początku skryptu nie potrzebuje fabryki. Fabryka ma sens, gdy to samo tworzenie występuje co najmniej trzy razy (reguła trzech z modułu 22) lub gdy konwencja ma być **egzekwowana** („każdy arkusz musi mieć zakładkę").

### 2.2. Builder: instrukcja, która sprawdza kompletność

**Intuicja.** Niektóre obiekty nie powstają „jednym ruchem". Powstają w **wielu krokach**, które muszą być wykonane **w odpowiedniej kolejności**, a na końcu trzeba sprawdzić, czy niczego nie brakuje.

**Analogia.** Instrukcja montażu mebla z IKEA. Krok po kroku, z numerkami, a na końcu: „sprawdź, czy wszystkie elementy z listy są na miejscu". Nigdy nie zaczynasz od przykręcenia tylnej płyty.

**Definicja.** **Builder** to wzorzec, w którym konstrukcję złożonego obiektu rozkładasz na **łańcuch nazwanych kroków** (`dodaj_arkusz`, `naglowek`, `wiersze`), a obiekt tworzysz dopiero na końcu (`build()`). Metoda `build()` ma **prawnie zagwarantować**, że obiekt jest kompletny — albo rzucić wyjątek.

```python
(RaportBuilder("Raport Q3")
    .dodaj_arkusz("Dane").naglowek(["Data", "Kwota", "Klient"]).wiersze(dane)
    .dodaj_arkusz("Podsumowanie").kpi("Suma", suma).tabela(zestawienie)
    .zapisz("output/q3.xlsx"))
```

To się nazywa **API płynne** (*fluent API*): każda metoda zwraca `self`, więc możesz je łańcuchować. Zysk jest czysto czytelniczy — cały raport widzisz jako **jedną pionową listę kroków**, a nie jako dwadzieścia rozproszonych linii.

**Kiedy Builder wygrywa:**

- **Kolejność ma znaczenie** — nagłówek przed danymi, tabela po danych,
- **Kompletność ma znaczenie** — raport bez podsumowania to błąd, a nie wariant,
- **Budowa jest wieloetapowa i rozłożona w czasie** — krok 1 robisz z pliku A, krok 2 z pliku B,
- **Chcesz mieć warianty tej samej budowy** — z tymi samymi krokami, innym zestawem.

**Kiedy Builder przegrywa (i to jest większość przypadków):**

- **Jeden arkusz, jedna tabela.** Builder każe Ci napisać 80 linii kodu, żeby obsłużyć 5 linii przypisań. To jest dokładnie ten `Builder` do trzech kolumn, o którym mówi pułapka z modułu 22.
- **Kolejność nie ma znaczenia.** Jeśli kroki są przemienne, Builder niczego nie pilnuje.
- **Nie ma czego weryfikować.** Jeśli „niekompletny" obiekt jest poprawny (np. raport z jednym arkuszem i opcjonalnym drugim), to `_waliduj()` będzie pustą metodą — znak, że wzorzec jest niepotrzebny.

### 2.3. Najważniejsza rzecz o Builderze: NIE pisz do arkusza w trakcie budowy

To jest pułapka, którą zapowiedział moduł 22 i najczęstszy błąd przy pierwszym zetknięciu ze wzorcem. Zobacz **wersję naiwną**:

```python
class RaportBuilder:
    def __init__(self, tytul):
        self.wb = Workbook()              # ← skoroszyt powstaje OD RAZU
        del self.wb["Sheet"]

    def dodaj_arkusz(self, nazwa):
        self.ws = self.wb.create_sheet(nazwa)   # ← arkusz powstaje OD RAZU
        return self

    def naglowek(self, etykiety):
        for i, e in enumerate(etykiety, start=1):
            self.ws.cell(row=1, column=i, value=e)   # ← zapis DO ARKUSZA
        return self
```

Co tu poszło nie tak? **Budowa i materializacja są splecione.** Skutki:

1. **Nie sprawdzisz kompletności „na sucho".** Kiedy wywołasz `.naglowek(...)`, arkusz już jest zapisany w modelu. Nie ma etapu, na którym możesz zapytać „czy raport jest kompletny?" **przed** rozpoczęciem pisania.
2. **Nie da się cofnąć kroku.** Jeżeli po `.dodaj_arkusz("Dane")` wołający zmieni zdanie i zrobi `.dodaj_arkusz("Podsumowanie")`, pierwszy arkusz już istnieje i jest pusty. Builder nie ma historii kroków — ma tylko skutki.
3. **Nie da się zbudować w pamięci „na dwa razy".** Jeżeli dane przychodzą z dwóch miejsc, po pierwszym `.wiersze(dane_a)` nie wiesz jeszcze, czy drugie źródło zadziała.
4. **`build()` nie ma czego weryfikować**, bo wszystko już się stało. Zostaje mu najwyżej `return self.wb` — metoda pusta.

**Wersja poprawna: Builder gromadzi PLAN, a materializuje dopiero `build()`.**

```python
class RaportBuilder:
    def __init__(self, tytul):
        self._tytul = tytul
        self._arkusze = []                # ← PLAN: lista opisow, nie arkuszy
        self._biezacy = None

    def build(self):
        self._waliduj()                   # ← weryfikacja PRZED czymkolwiek
        wb = Workbook()
        del wb["Sheet"]
        for plan in self._arkusze:        # ← materializacja
            ...
        return wb
```

Analogia: **instrukcja montażu a gotowy mebel.** Instrukcja to lista kroków na papierze — możesz ją sprawdzić, poprawić, porównać z listą części, zanim tkniesz śrubokręt. Dopiero `build()` bierze śrubokręt do ręki.

**Konsekwencja praktyczna, która jest prawdziwym zyskiem:** `build()` zwraca `Workbook`, **nie plik**. To znaczy, że możesz go przetestować dokładnie tak, jak uczył moduł 20:

```python
wb = RaportBuilder("Test").dodaj_arkusz("Dane").naglowek(["A"]).wiersze([[1]]).build()
assert wb["Dane"]["A1"].value == "A"
```

Zero plików, zero dysku, milisekundy. I osobno, na końcu, `zapisz()` = `build()` + `save()`. **Builder, który pisze w trakcie budowy, odbiera Ci tę możliwość.** Nie przez przypadek — przez konstrukcję.

### 2.4. Prototype: pieczątka, która nigdy nie jest nadpisywana

**Intuicja.** Niektóre skoroszyty są tak bogate wizualnie (logo, nagłówki, ustawienia wydruku, motyw), że odtworzenie ich kodem jest bez sensu. Robisz je **raz, ręcznie w Excelu** — i używasz jako wzorca.

**Analogia.** Pieczątka firmowa. Nadruk jest gotowy. Ty dopisujesz tylko treść zmienną. I kluczowa właściwość: **odciskasz z pieczątki, a nie piszesz po pieczątce**. Tuszujesz ją i przyciskasz w nowym miejscu.

**Definicja.** **Prototype** to wzorzec, w którym nowy obiekt powstaje przez **skopiowanie istniejącego wzorca** i modyfikację kopii, zamiast przez budowanie od zera.

```python
wb = load_workbook("templates/raport.xlsx")   # wzorzec wczytany do PAMIECI
ws = wb["Dane"]
# ... dopisujesz TYLKO komorki, ktore maja sie zmieniac ...
wb.save("output/raport_2026_10.xlsx")          # ← NOWY plik!
```

**Kluczowa reguła:** szablon **nigdy** nie jest celem zapisu. To brzmi banalnie, ale zdarza się nagminnie — wystarczy zmienna, która raz wskazuje na szablon, a raz na wynik. Dlatego reguła powinna być **egzekwowana w kodzie**, nie tylko w głowie:

```python
def raport_z_szablonu(szablon, cel, ...):
    if Path(szablon).resolve() == Path(cel).resolve():
        raise ValueError("Szablon nie moze byc celem zapisu!")
```

Zapisz to sobie jako wzorzec do powtarzania: **„reguła, której nie ma w kodzie, jest życzeniem"**.

**Pułapka, którą musisz znać z modułu 18.** Szablon to plik, który najczęściej **został zrobiony ręcznie w Excelu**. A to znaczy, że może zawierać rzeczy, których openpyxl nie modeluje:

| Element szablonu | Co się dzieje przy `load_workbook` + `save` |
|---|---|
| formatowanie komórek, formaty liczb, scalenia | **przechodzi** |
| tabela Excela, filtr, formatowanie warunkowe | **przechodzi** (z wyjątkami wersyjnymi) |
| obrazy w arkuszu, którego **nie kopiujesz** | **przechodzą** |
| wykresy | **przechodzą, ale są odbudowywane z modelu openpyxl** — nietypowe formatowanie może zostać uproszczone |
| **sparkline'y** (rozszerzenie `x14`) | **GUBIONE** — openpyxl ostrzega przy wczytaniu |
| kształty, pola tekstowe, formanty | **GUBIONE** — openpyxl ich nie modeluje |
| slicery, Power Query, Power Pivot | **GUBIONE** |
| makra VBA | **tylko z `keep_vba=True`** i tylko zachowane, nieedytowalne |

I teraz **druga pułapka, specyficzna dla Prototype**: skoro wzorcem jest szablon, to naturalna myśl brzmi „wezmę szablon i **skopiuję arkusz** dla każdego regionu". To wygląda elegancko i **nie działa**:

```python
wb = load_workbook(szablon)
for region in ["PL", "DE", "FR"]:
    kopia = wb.copy_worksheet(wb["Szablon"])
    kopia.title = region
    # ...
```

`copy_worksheet` kopiuje **wartości, style, hyperlinki i komentarze** — i część atrybutów arkusza. **Nie kopiuje obrazów i wykresów.** Więc jeżeli w szablonie jest logo (a zwykle jest), kopia będzie pusta w tym miejscu. Zmierzymy to w przykładzie 3 przez porównanie długości list `_images`.

**Trzy wyjścia z tego problemu:**

1. **Nie kopiuj arkusza — używaj go bezpośrednio.** Wczytaj szablon, wypełnij **istniejący** arkusz, zapisz jako nowy plik. Logo przechodzi, bo nikt go nie kopiował — ono jest w arkuszu, którego używasz.
2. **Wstaw logo kodem.** Jeżeli kopiujesz arkusze, dodaj obraz programowo po skopiowaniu (`ws.add_image(...)`) — ścieżka logo musi być dostępna w runtime.
3. **Wygeneruj szablon w kodzie** (`przygotuj_szablon()`) zamiast ręcznie — wtedy wiesz dokładnie, co w nim jest, i możesz go odtworzyć po zmianie biblioteki. Koszt: tracisz wygodę ręcznego projektowania w Excelu.

### 2.5. Singleton i Registry: matryca, która leży na półce

**Intuicja.** Tworzysz `Font(bold=True, color="FFFFFF", size=11)` każdego dnia, w każdym raporcie, w każdej pętli. To ten sam obiekt. Trzymaj go **raz** i podawaj z półki.

**Analogia.** Drukarnia z gotowymi matrycami. Nie rzeźbisz czcionki dla każdego wyrazu w gazecie. Matryca leży na półce, jest zrobiona raz, a używasz jej tysiące razy. Jeżeli trzeba zmienić krój — **wymieniasz matrycę, a nie każdy wydruk**.

**Definicja.** **Registry** to obiekt, który przechowuje gotowe do użycia elementy i wydaje je na żądanie, tworząc (i zapamiętując) tylko te, których jeszcze nie ma. To wariant wzorca **Flyweight** (współdzielenie ciężkich obiektów) w wersji „z pamięcią".

```python
class StyleRegistry:
    def font(self, **cechy) -> Font:
        klucz = tuple(sorted(cechy.items()))
        gotowy = self._fonts.get(klucz)
        if gotowy is None:
            gotowy = Font(**cechy)
            self._fonts[klucz] = gotowy
        return gotowy
```

**Konsekwencja praktyczna.** 20 000 wywołań `registry.font(bold=True, color="FFC000")` tworzy **jeden** obiekt `Font`. Czas i pamięć Pythona spadają, a plik wyjściowy **jest taki sam** — bo openpyxl i tak deduplikuje style przy zapisie.

#### Flyweight: co jest wspólne, a co osobne

Flyweight opiera się na jednym rozróżnieniu, które warto znać nie tylko dla openpyxl:

- **Stan wewnętrzny (intrinsic)** — to, co jest **identyczne** dla wszystkich egzemplarzy. Tu: zestaw cech czcionki (`bold`, `color`, `size`). To jest współdzielone.
- **Stan zewnętrzny (extrinsic)** — to, co się **różni**. Tu: **która komórka** używa tej czcionki.

`Font(bold=True, color="FFC000")` to stan wewnętrzny. `ws["A1"]` to stan zewnętrzny. Dlatego jeden obiekt `Font` można bezpiecznie przypisać do stu tysięcy komórek — **i to jest bezpieczne właśnie dlatego, że style są niemutowalne** (moduł 09). Gdybyś mógł zrobić `font.color = "FF0000"`, to jedna zmiana pomalowałaby na czerwono sto tysięcy komórek naraz. Niemutowalność nie jest ograniczeniem — jest **warunkiem możliwości współdzielenia**.

#### Pułapka Singletonu: tablica na hali

Skoro rejestr jest przydatny, kusi, żeby zrobić z niego **globalny singleton**:

```python
# ZACHĘCAJĄCO WYGLĄDA, ALE MA KONSEKWENCJE
STYLE = StyleRegistry()      # jeden modul, jeden obiekt

def pisz_raport():
    ws["A1"].font = STYLE.font(bold=True)    # ← skad wiemy, skad?
```

Dlaczego to jest problem? Bo **rejestr to nie tylko cache — rejestr to konfiguracja.** Jeżeli w rejestrze żyją kolory firmowe i motyw, to globalny singleton znaczy:

1. **Testy przecież oddziałują na siebie.** Test A zmieni kolor w rejestrze, test B zobaczy ten kolor. Kolejność wykonania testów wpływa na wynik. To najgorszy rodzaj błędu: **przechodzi u Ciebie, pada w CI**.
2. **Nie da się mieć dwóch motywów naraz.** Raport dla działu A ma inne barwy niż dla działu B? Globalny rejestr ma kolor **jeden**.
3. **Ukryta zależność.** Patrząc na `pisz_raport()`, nie wiesz, **skąd** bierze się wygląd. Funkcja wygląda na czystą, a jest zależna od globalnego stanu. To jest dokładnie to, przed czym ostrzegała sekcja 2.3 modułu 22 (kolumna „co płacisz").

**Rozwiązanie: wstrzykiwanie rejestru.** Rejestr jest obiektem, który **przekazujesz** funkcji:

```python
def pisz_raport(ws, style: StyleRegistry) -> None:
    ws["A1"].font = style.font_naglowka()      # ← JAWNA zaleznosc
```

Teraz widać z sygnatury, że funkcja potrzebuje stylów. W testach przekazujesz rejestr z motywem testowym, w produkcji — firmowy. Możesz mieć dwa naraz.

**Kiedy Singleton (albo moduł jako singleton) jest w porządku?** Wtedy, gdy przechowywana rzecz jest **naprawdę bezstanowa i niemutowalna**. Na przykład:

```python
# openpyxl.constants - to jest singleton i to jest OK
CALENDAR_WINDOWS_1900 = "windows"
```

Stała to nie stan. Rejestr **stylów** to stan, bo ma cache i motyw. Natomiast możesz bezpiecznie trzymać globalnie **same definicje** (np. `DEFINICJE` z przykładu 1), bo są `frozen` i nikt ich nie zmienia w trakcie działania.

> **Reguła kciuka:** moduł jako singleton dla **danych** — tak. Singleton dla **obiektów z cache albo konfiguracją** — prawie nigdy. I pamiętaj: w Pythonie **moduł już jest singletonem** (import wykonuje się raz), więc nie potrzebujesz do tego klasy z `__new__`.

### 2.6. Object Pool i leniwe tworzenie: użyj, gdy naprawdę widać zysk

**Intuicja.** Niektóre rzeczy są drogie. Tworzysz je, gdy naprawdę są potrzebne, nie „na wszelki wypadek".

**Analogia.** Zakład produkcyjny nie uruchamia linii na wszystkie 30 wariantów wyrobu, zanim dostanie zamówienia. Uruchamia linię **gdy zamówienie przychodzi**.

**Definicja.** **Leniwe tworzenie** (lazy initialization) to odkładanie kosztu utworzenia obiektu do chwili pierwszego użycia. **Object Pool** to trzymanie gotowych obiektów do ponownego użycia, żeby nie tworzyć ich wielokrotnie.

**Gdzie to ma sens w openpyxl?** Trzy miejsca:

1. **Nie wczytuj szablonu, jeśli go nie użyjesz.** `load_workbook()` czyta cały ZIP i buduje model w pamięci. Jeżeli raport ma być wygenerowany od zera (np. dla danych bez regionu), wczytywanie szablonu to zmarnowane sekundy i megabajty.
2. **Nie twórz arkusza, dla którego nie ma danych.** Scenariusz: 30 arkuszy (po jednym na region), ale w danym miesiącu sprzedaż była tylko w 12 regionach. Leniwe tworzenie daje 12 arkuszy zamiast 30.
3. **Nie licz tego, czego nie zapiszesz.** Jeżeli nie wiesz, czy raport będzie potrzebował zestawienia krzyżowego, policz je dopiero, gdy okaże się potrzebne.

**Muszę tu być bardzo uczciwy, bo łatwo sprzedać leniwość jako cud:**

> **W openpyxl leniwe tworzenie **arkuszy** daje realny zysk tylko wtedy, gdy część arkuszy faktycznie zostaje pusta.** W trybie normalnym `wb.create_sheet()` tworzy obiekt `Worksheet` z pustym słownikiem komórek — to jest **tanie**. Prawdziwy koszt siedzi w komórkach. Jeżeli wszystkie 30 arkuszy ma po 100 000 wierszy, leniwość arkuszy nie da **nic**: dominuje koszt komórek, a te musisz stworzyć wszystkie.

Prawdziwe oszczędności w dużych plikach to moduł 19: `read_only`, `write_only`, `lxml`, `append` zamiast przypisywania po jednej komórce. Leniwe tworzenie to narzędzie **do rzadkich, warunkowych gałęzi** — nie do wydajności jako takiej.

**Ważne w `write_only`:** w trybie strumieniowym **kolejność jest nieodwracalna**. Nie możesz wrócić do arkusza, który już wypchnąłeś. To znaczy, że Builder z tego modułu jest **szczególnie** przydatny przy `write_only`, bo plan masz gotowy **przed** pierwszym `append` — a nie odkrywasz go w trakcie.

### 2.7. Konfiguracja zamiast kodu: przepis vs gotowe danie

**Intuicja.** Jeżeli raport różni się od innego tylko listą kolumn, tytułem i formatami — to **nie jest inny kod**, to inne **dane**.

**Analogia.** Przepis kulinarny a gotowe danie. Przepis opisuje *co* ma powstać. Kucharz wie *jak* to zrobić. Jeżeli zmieniasz danie, zmieniasz przepis albo wybierasz inny — nie przekwalifikowujesz kucharza.

**Definicja.** `RaportSpec` to `frozen` dataclass opisujący raport: tytuł, arkusze, kolumny, motyw, progi. Z niego powstaje skoroszyt.

Zysk jest bezpośrednio mierzalny i to jest **ta sama miara OCP**, którą poznałeś w module 22: **ile linii dopisałeś do istniejącego kodu, żeby powstał nowy raport?** Nowy raport = nowa instancja `RaportSpec` → zero linii w generatorze.

**I ostrzeżenie, które muszę powtórzyć, bo moduł 22 je zapowiedział:** granica między konfiguracją a kodem jest cienka. Konfiguracja opisuje **co**. Kod opisuje **jak**.

```python
# DOBRZE: to jest "co"
RaportSpec(nazwa="sprzedaz", tytul="Sprzedaz Q3",
           kolumny=(Kolumna("Data", 12, "dd.mm.yyyy"), ...))

# ZLE: to jest "jak" w przebraniu danych
RaportSpec(filtr="ilosc > 10 and region in ['PL','DE']")     # ← mini-jezyk
RaportSpec(sortowanie="desc(wartosc), asc(klient)")          # ← mini-jezyk
RaportSpec(agregacja="sum(ilosc*cena) group by region")      # ← SQL w stringu
```

Ten drugi wariant wygląda na elegancki, a jest **pułapką**: tworzysz język wyrażeń bez parsera, bez typów, bez podpowiedzi edytora i bez testów. Po dwóch miesiącach konfiguracja ma `if`, `and`, zagnieżdżenia i wymaga debuggera — a wtedy **przestała być konfiguracją**.

> **Test: jeżeli konfiguracja wymaga debuggera, to nie jest konfiguracja.** Operatory (`>`, `and`, `or`), pętle i warunki zostają w kodzie jako **nazwane funkcje** w świecie danych (moduł 22).

Dlatego w przykładzie 5 `RaportSpec` opisuje **kolumny i tytuły**, a filtrowanie i agregacja są **funkcjami**. W module 26 zobaczysz wersję YAML-ową — wraz z pełnym ostrzeżeniem o tym samym ryzyku.

### 2.8. DTO: dlaczego nie przekazywać słowników i list

**Intuicja.** Jeżeli między warstwami przekazujesz `{"data": ..., "klient": ..., "cena": ...}`, to **kształt danych istnieje tylko w Twojej głowie**. Kompilator, edytor i `mypy` nie wiedzą nic o tym słowniku.

**Analogia.** Karta wyrobu w zakładzie. Nie przekazujesz luźnej kartki z odręcznymi bazgrołami — przekazujesz **formularz** z nazwanymi polami. Każdy wie, gdzie co wpisać, i od razu widać, gdy pole jest puste.

**Definicja.** **DTO** (*Data Transfer Object*) to obiekt, który nie ma logiki — tylko niesie dane z jednego miejsca w drugie.

```python
@dataclass(frozen=True)
class WierszRaportu:
    data: date
    klient: str
    region: str
    ilosc: int
    cena: float

    def __post_init__(self) -> None:
        if self.ilosc < 0:
            raise ValueError(f"ilosc nie moze byc ujemna: {self.ilosc!r}")
        if not self.klient.strip():
            raise ValueError("klient nie moze byc pusty")
```

**Cztery konkretne zyski wobec `dict`:**

1. **Literówka to błąd w miejscu powstania.** `WierszRaportu(klient="...")` — jeżeli napiszesz `kliient=`, dostaniesz `TypeError` **natychmiast**, przy tworzeniu obiektu. `{"kliient": ...}` żyje dalej i wybucha dopiero wtedy, gdy `w.get("klient", "")` cicho zwróci pusty string — czyli jako pusta kolumna w raporcie, którą klient zobaczy po tygodniu.
2. **Walidacja w konstruktorze.** `__post_init__` uruchamia się przy tworzeniu. Błędne dane **nie wchodzą** do systemu.
3. **Typy są widoczne.** Edytor podpowiada pola, `mypy` sprawdza.
4. **Nie da się przez przypadek dopisać pola.** Bez `frozen=True` ktoś w połowie pipeline'u doda `wiersz.uwagi = "..."` i to pole pojedzie dalej — niewidoczne w żadnej definicji.

**Kiedy NIE używać DTO?** Gdy dane naprawdę są dynamiczne i nie mają stałego kształtu (np. wynik zapytania o nieznanych kolumnach). Wtedy `dict` jest właściwy. Ale **dane, które trafiają do raportu o znanym układzie, mają znany kształt** — więc DTO pasuje niemal zawsze.

### 2.9. Tabela decyzyjna: który wzorzec na jaki problem

Ta tabela jest najważniejszym praktycznym elementem modułu. Wróć do niej, gdy będziesz się zastanawiać, „czy już".

| Problem, który widzisz w kodzie | Wzorzec | Sygnał, że to problem | Koszt wdrożenia | Kiedy NIE |
|---|---|---|---|---|
| te same 5–7 linii konfiguracji arkusza w 20 miejscach | **Factory** | sygnał 1 z modułu 22 | niski | gdy tworzysz arkusz raz w jednym miejscu |
| ta sama czcionka/wypełnienie konstruowane w 40 miejscach | **StylFabryka + Registry** | sygnał 2 | niski | gdy kolor jest używany w 2 komórkach |
| kolory/progi bez nazw, „magiczne" | **Konfiguracja (`Motyw`, `RaportSpec`)** | sygnały 2 i 3 | niski | — **zawsze warto**, to najtańszy wzorzec |
| raport składany w 6 krokach, kolejność i kompletność mają znaczenie | **Builder** | brak walidacji kompletności | średni | jeden arkusz, jedna tabela, kroki przemienne |
| raport, w którym „brak podsumowania" przechodzi niezauważony | **Builder z `build()`/`_waliduj()`** | brak testów na kompletność | średni | gdy „niekompletny" jest poprawny |
| bogaty szablon robiony ręcznie w Excelu | **Prototype** | 200 linii kodu odtwarzających wygląd | niski | gdy nic w szablonie nie trzeba zachować |
| arkusz używany jako „wzór dla każdego regionu" | **Prototype + kopia lub reużycie** | duplikacja logiki wypełniania | niski | pamiętaj: `copy_worksheet` gubi obrazy i wykresy |
| ten sam `Font` w tysiącach komórek; mierzysz czas i pamięć | **Flyweight / Registry** | spadek wydajności | niski | gdy plik jest mały — zysk niemierzalny |
| globalna konfiguracja stylów, „wszędzie dostępna" | **Wstrzykiwany Registry** | testy padają zależnie od kolejności | niski | gdy rejestr jest bezstanowy i `frozen` |
| 30 arkuszy, część pusta, szablon ładowany niepotrzebnie | **Lazy** | marnotrawstwo czasu i pamięci | niski | gdy wszystkie gałęzie są zawsze używane |
| przekazywanie `dict`/`list` między warstwami | **DTO (`WierszRaportu`)** | literówka = pusta kolumna | niski | dane naprawdę dynamiczne |
| nowy raport = `elif` w generatorze | **Konfiguracja (`RaportSpec`)** | złamane OCP (moduł 22) | średni | — **zawsze warto** |
| Builder ma 200 linii, a raport 10 wierszy | **nic. Usuń Builder.** | nadmiar abstrakcji | — | to jest ten przypadek |

## 3. Przykłady krok po kroku

Cztery przykłady, ten sam raport sprzedaży, kontynuacja modułu 22. **Każdy jest uruchamialny** i każdy zapisuje pliki do `output/`, żebyś mógł otworzyć je w Excelu i *zobaczyć* efekt.

### Przykład 1 — Fabryka: koniec z „20 miejscami"

**Problem do rozwiązania.** W module 22 nasz kod wyglądał tak:

```python
ws = wb.create_sheet(title=spec.tytul_arkusza)
ws.cell(row=1, column=1, value=spec.tytul).font = Font(bold=True, size=16)
ws.freeze_panes = "A4"
```

Dla jednego raportu to w porządku. Dla siedmiu raportów, każdy z trzema–pięcioma arkuszami, te same linie się powtarzają. Dochodzi kolor zakładki, wyłączenie linii siatki, szerokości, `sheet_state`. **To jest sygnał 1 z modułu 22 w pełnej krasie.**

**Rozwiązanie.** Oddzielamy **definicję** arkusza od **kodu, który go tworzy**.

```python
"""Modul 23, przyklad 1: ArkuszFabryka + StylFabryka.

Problem: 6 linii konfiguracji arkusza powtorzone w 20 miejscach.
Rozwiazanie: jedna fabryka, ktora WIE, jak wyglada 'arkusz danych'.

Uruchom: python examples/23_01_fabryka.py
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

if TYPE_CHECKING:
    from openpyxl.workbook.workbook import Workbook as TypWorkbook
    from openpyxl.worksheet.worksheet import Worksheet


# =====================================================================
# 1. MOTYW - wszystkie decyzje o wygladzie w JEDNYM miejscu
# =====================================================================
@dataclass(frozen=True)
class Motyw:
    tlo_naglowka: str = "FFC000"
    tekst_naglowka: str = "FFFFFF"
    tekst_alertu: str = "FF0000"
    rozmiar_tytulu: int = 16


# =====================================================================
# 2. DEFINICJE ARKUSZY - to sa "foremki", czyli dane, nie kod
# =====================================================================
@dataclass(frozen=True)
class DefinicjaArkusza:
    nazwa: str                          # jak bedzie widoczna w Excelu
    tab_color: str | None = None
    freeze: str | None = None
    show_grid: bool = True
    szerokosci: tuple[float, ...] = ()
    stan: str = "visible"               # visible | hidden | veryHidden


DEFINICJE: dict[str, DefinicjaArkusza] = {
    "dane": DefinicjaArkusza(
        nazwa="Dane",
        tab_color="1072BA",
        freeze="A4",
        show_grid=False,                # konwencja firmy: brak siatki w danych
        szerokosci=(12.0, 24.0, 8.0, 16.0, 8.0, 10.0),
    ),
    "podsumowanie": DefinicjaArkusza(
        nazwa="Podsumowanie",
        tab_color="70AD47",
        show_grid=False,
        szerokosci=(24.0, 16.0),
    ),
    "wykresy": DefinicjaArkusza(
        nazwa="Wykresy",
        tab_color="C00000",
        show_grid=False,
    ),
    "konfiguracja": DefinicjaArkusza(
        nazwa="_Konfiguracja",
        tab_color="808080",
        stan="veryHidden",              # widoczna w pliku, ukryta przed uzytkownikiem
        szerokosci=(28.0, 40.0),
    ),
}
```

**Co się dzieje w pamięci.** `DEFINICJE` to klucz → `DefinicjaArkusza`. Klucz (`"dane"`) to **nazwa logiczna** używana w kodzie; `DefinicjaArkusza.nazwa` (`"Dane"`) to **nazwa widoczna w Excelu**. To rozdzielenie jest ważniejsze, niż wygląda: możesz zmienić tytuł arkusza w pliku **bez dotykania ani jednej linii kodu**. `DefinicjaArkusza` jest `frozen`, więc jedna instancja może być bezpiecznie współdzielona przez wszystkie raporty w procesie — tak jak `Motyw` z modułu 22.

Teraz fabryka stylów i fabryka arkuszy.

```python
# =====================================================================
# 3. FABRYKA STYLOW (to jest jednoczesnie Registry - sekcja 2.5)
# =====================================================================
class StylFabryka:
    """Tworzy i CACHE'uje obiekty stylu. Jedna instancja = jedna matryca.

    Kazdy obiekt Font powstaje RAZ i jest wydawany wielokrotnie.
    Bez tego 20 000 komorek = 20 000 obiektow Font (przyklad 4).
    """

    def __init__(self, motyw: Motyw | None = None) -> None:
        self.motyw = motyw or Motyw()
        self._font_cache: dict[tuple, Font] = {}
        self._fill_cache: dict[tuple, PatternFill] = {}
        self._align_cache: dict[tuple, Alignment] = {}
        self.licznik_alokacji = 0      # metryka: ile obiektow NAPRAWDE powstalo

    def font(self, **cechy) -> Font:
        klucz = tuple(sorted(cechy.items()))       # wartosci musza byc hashable!
        gotowy = self._font_cache.get(klucz)
        if gotowy is None:
            gotowy = Font(**cechy)
            self._font_cache[klucz] = gotowy
            self.licznik_alokacji += 1
        return gotowy

    def fill(self, kolor: str, fill_type: str = "solid") -> PatternFill:
        klucz = (fill_type, kolor)
        gotowy = self._fill_cache.get(klucz)
        if gotowy is None:
            gotowy = PatternFill(fill_type=fill_type, fgColor=kolor)
            self._fill_cache[klucz] = gotowy
            self.licznik_alokacji += 1
        return gotowy

    def align(self, **cechy) -> Alignment:
        klucz = tuple(sorted(cechy.items()))
        gotowy = self._align_cache.get(klucz)
        if gotowy is None:
            gotowy = Alignment(**cechy)
            self._align_cache[klucz] = gotowy
            self.licznik_alokacji += 1
        return gotowy

    # --- gotowe "pieczatki" wynikajace z motywu ---
    def font_naglowka(self) -> Font:
        return self.font(bold=True, color=self.motyw.tekst_naglowka)

    def font_tytulu(self) -> Font:
        return self.font(bold=True, size=self.motyw.rozmiar_tytulu)

    def font_alertu(self) -> Font:
        return self.font(color=self.motyw.tekst_alertu)

    def fill_naglowka(self) -> PatternFill:
        return self.fill(self.motyw.tlo_naglowka)
```

**Co się dzieje w pamięci — i jedna uczciwa pułapka.** Kluczem cache jest `tuple(sorted(cechy.items()))`. To wymaga, żeby **wartości były hashowalne**. `bold=True`, `size=16`, `color="FFFFFF"` — tak. Ale gdyby ktoś napisał `font(color=["FF0000"])` (lista), dostaniesz `TypeError: unhashable type`. To jest cena cache'u: **przejmujesz odpowiedzialność za klucz**. Dobrze, że wyjdzie przy pierwszym wywołaniu, a nie po tygodniu.

I druga rzecz: `self.licznik_alokacji` **nie rośnie** przy trafieniu w cache. To nie ozdoba — to metryka, której użyjemy w ćwiczeniu 🟡 i w przykładzie 4, żeby **udowodnić**, że cache działa. Bez metryki „optymalizacja" jest wiarą.

```python
# =====================================================================
# 4. FABRYKA ARKUSZY - jedyne miejsce w projekcie z create_sheet
# =====================================================================
class ArkuszFabryka:
    """Tworzy arkusze wedlug definicji. Nic wiecej nie potrafi - i tak ma byc."""

    def __init__(
        self,
        definicje: dict[str, DefinicjaArkusza],
        style: StylFabryka,
    ) -> None:
        self._definicje = definicje
        self.style = style                      # arkusz dostaje swoj styl "w pakiecie"
        self.utworzone: list[str] = []

    def utworz(
        self,
        wb: TypWorkbook,
        klucz: str,
        index: int | None = None,
    ) -> Worksheet:
        try:
            def_ = self._definicje[klucz]
        except KeyError:
            znane = sorted(self._definicje)
            raise KeyError(
                f"Nieznany typ arkusza: {klucz!r}. Znane: {znane}"
            ) from None

        ws = wb.create_sheet(title=def_.nazwa, index=index)

        if def_.tab_color:
            ws.sheet_properties.tabColor = def_.tab_color
        if def_.freeze:
            ws.freeze_panes = def_.freeze
        if not def_.show_grid:
            ws.sheet_view.showGridLines = False
        for i, szer in enumerate(def_.szerokosci):
            ws.column_dimensions[get_column_letter(i + 1)].width = szer
        if def_.stan != "visible":
            ws.sheet_state = def_.stan

        self.utworzone.append(def_.nazwa)
        return ws
```

**Trzy detale warte uwagi:**

1. **`try/except KeyError` z listą znanych kluczy.** Zamiast surowego `KeyError: 'danne'` dostajesz komunikat, który mówi, co było złe **i co jest dostępne**. To jest ta różnica między „kodem, który rzuca wyjątek" a „kodem, który tłumaczy błąd". Kosztuje trzy linie i ratuje dziesięć minut debugowania o 23:00.
2. **Fabryka nic nie wie o treści raportu.** Nie ma tu ani jednej decyzji o sprzedaży, kwotach ani regułach. Fabryka wie wyłącznie o **kształcie arkusza**. To jest zasada **S** z modułu 22 w praktyce: jedna odpowiedzialność.
3. **`self.utworzone`** to lista kontrolna — przydatna w testach („fabryka utworzyła dokładnie te arkusze") i w logowaniu.

**Użycie i „co trafi do pliku":**

```python
def main() -> None:
    style = StylFabryka(Motyw())
    fabryka = ArkuszFabryka(DEFINICJE, style)

    wb = Workbook()
    del wb["Sheet"]                       # domyslny arkusz nie jest nam potrzebny

    ws_dane = fabryka.utworz(wb, "dane", index=0)
    ws_dane["A1"] = "Raport sprzedazy - styczen 2026"
    ws_dane["A1"].font = style.font_tytulu()

    ws_podsum = fabryka.utworz(wb, "podsumowanie", index=1)
    ws_podsum["A1"] = "Podsumowanie"
    ws_podsum["A1"].font = style.font_naglowka()

    fabryka.utworz(wb, "konfiguracja", index=2)

    wb.properties.title = "Raport sprzedazy"

    cel = Path("output/23_01_fabryka.xlsx")
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)
    print("Arkusze:", fabryka.utworzone)
    print("Obiektow stylu utworzono:", style.licznik_alokacji)


if __name__ == "__main__":
    main()
```

Do pliku trafią trzy arkusze plus zerwany domyślny `Sheet`. **Do `xl/workbook.xml` trafią wpisy w kolejności `Dane`, `Podsumowanie`, `_Konfiguracja`** — bo podaliśmy `index`. Do `xl/worksheets/sheetN.xml` trafią kolor zakładki i `sheetState="veryHidden"` dla `_Konfiguracja` (a nie do głównego pliku — te atrybuty siedzą w XML arkusza). Do `xl/styles.xml` trafią **dwa** wpisy `font` — pogrubiony rozmiar 16 i pogrubiony biały. `licznik_alokacji` pokaże dokładnie tyle, ile jest wpisów w `styles.xml`. **To jest ta miara, której nie widzisz w pliku, a która mówi, czy cache działa.**

### Przykład 2 — Builder: instrukcja z kontrolą kompletności

**Problem.** Mamy teraz fabrykę arkuszy, ale **kolejność i kompletność** raportu nadal żyją w głowie wołającego. Jeżeli zapomni wywołać funkcję dodającą podsumowanie, nikt tego nie zauważy.

**Rozwiązanie.** Builder, który **gromadzi plan** (a nie pisze do arkuszy) i materializuje go dopiero w `build()` — po walidacji.

```python
"""Modul 23, przyklad 2: RaportBuilder.

Builder NIE PISZE do arkuszy w trakcie budowy. Gromadzi PLAN i materializuje
go w build() - dopiero po weryfikacji kompletnosci. Dzieki temu build()
zwraca Workbook, ktory mozna przetestowac bez zapisu na dysk (modul 20).

Uruchom: python examples/23_02_builder.py
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


class BladBudowy(Exception):
    """Podniesiony przez build(), gdy raport jest niekompletny.

    Wlasny wyjatek zamiast ValueError: dzieki temu wolajacy moze go
    odroznic od zwyklego bledu wartosci, a komunikat ma liste problemow.
    """


def _nazwa_tabeli(nazwa: str) -> str:
    """Excel wymaga nazwy tabeli bez spacji i znakow specjalnych."""
    bezpieczna = re.sub(r"[^0-9A-Za-z_]", "_", nazwa)
    if not bezpieczna or bezpieczna[0].isdigit():
        bezpieczna = "T_" + bezpieczna
    return bezpieczna


# =====================================================================
# PLAN - neutralny opis tego, co ma powstac. Zero openpyxl.
# =====================================================================
@dataclass
class _PlanArkusza:
    nazwa: str
    naglowek: tuple[str, ...] | None = None
    wiersze: list[tuple] = field(default_factory=list)
    kpi: list[tuple[str, object, str | None]] = field(default_factory=list)
    tabela: list[tuple] | None = None
    szerokosci: tuple[float, ...] | None = None

    def czy_pusty(self) -> bool:
        return not self.wiersze and not self.kpi and self.tabela is None


# =====================================================================
# BUILDER
# =====================================================================
class RaportBuilder:
    """Instrukcja montazu raportu. Kroki -> plan -> build() -> Workbook."""

    def __init__(self, tytul: str, style=None) -> None:
        self._tytul = tytul
        self._style = style
        self._arkusze: list[_PlanArkusza] = []
        self._biezacy: _PlanArkusza | None = None

    # ---------------- API plynne ----------------
    def dodaj_arkusz(self, nazwa: str, szerokosci=None) -> "RaportBuilder":
        plan = _PlanArkusza(
            nazwa=nazwa,
            szerokosci=tuple(szerokosci) if szerokosci else None,
        )
        self._arkusze.append(plan)
        self._biezacy = plan
        return self

    def naglowek(self, etykiety) -> "RaportBuilder":
        self._wymagaj_arkusza().naglowek = tuple(etykiety)
        return self

    def wiersze(self, dane) -> "RaportBuilder":
        self._wymagaj_arkusza().wiersze.extend(tuple(w) for w in dane)
        return self

    def kpi(self, etykieta: str, wartosc, fmt: str | None = None) -> "RaportBuilder":
        self._wymagaj_arkusza().kpi.append((etykieta, wartosc, fmt))
        return self

    def tabela(self, dane) -> "RaportBuilder":
        """Pierwszy wiersz `dane` traktujemy jako naglowek tabeli Excela."""
        self._wymagaj_arkusza().tabela = [tuple(w) for w in dane]
        return self

    # ---------------- materializacja ----------------
    def build(self) -> Workbook:
        self._waliduj()                       # kontrolа kompletosci PRZED praca
        wb = Workbook()
        del wb["Sheet"]
        for i, plan in enumerate(self._arkusze):
            ws = wb.create_sheet(title=plan.nazwa, index=i)
            self._wypelnij(ws, plan)
        wb.properties.title = self._tytul
        wb.properties.creator = "RaportBuilder"
        return wb

    def zapisz(self, cel) -> Path:
        cel = Path(cel)
        cel.parent.mkdir(parents=True, exist_ok=True)
        self.build().save(cel)
        return cel

    # ---------------- wewnetrzne ----------------
    def _wymagaj_arkusza(self) -> _PlanArkusza:
        if self._biezacy is None:
            raise BladBudowy(
                "Najpierw wywolaj dodaj_arkusz(...), potem naglowek/wiersze/kpi/tabela."
            )
        return self._biezacy

    def _waliduj(self) -> None:
        problemy: list[str] = []
        if not self._arkusze:
            problemy.append("raport nie ma zadnego arkusza")

        widziane: set[str] = set()
        for plan in self._arkusze:
            if plan.nazwa in widziane:
                problemy.append(f"zduplikowana nazwa arkusza: {plan.nazwa!r}")
            widziane.add(plan.nazwa)

            if len(plan.nazwa) > 31:
                problemy.append(f"nazwa arkusza dluzsza niz 31 znakow: {plan.nazwa!r}")
            if re.search(r"[:\\/?*\[\]]", plan.nazwa):
                problemy.append(f"niedozwolony znak w nazwie arkusza: {plan.nazwa!r}")

            if plan.czy_pusty():
                problemy.append(f"arkusz {plan.nazwa!r} jest pusty")

            if plan.naglowek and plan.wiersze:
                n = len(plan.naglowek)
                zle = [i for i, w in enumerate(plan.wiersze, start=1) if len(w) != n]
                if zle:
                    problemy.append(
                        f"arkusz {plan.nazwa!r}: {len(zle)} wiersz(y) maja inna "
                        f"liczbe kolumn niz naglowek ({n}); pierwszy zly: {zle[0]}"
                    )

            if plan.wiersze and plan.naglowek is None:
                problemy.append(
                    f"arkusz {plan.nazwa!r} ma dane, ale nie ma naglowka"
                )

        if problemy:
            raise BladBudowy(
                "Raport niekompletny:\n  - " + "\n  - ".join(problemy)
            )
```

```python
    def _wypelnij(self, ws, plan: _PlanArkusza) -> None:
        wiersz = 1

        if plan.naglowek:
            for i, etykieta in enumerate(plan.naglowek, start=1):
                c = ws.cell(row=wiersz, column=i, value=etykieta)
                if self._style:
                    c.font = self._style.font_naglowka()
                    c.fill = self._style.fill_naglowka()
                    c.alignment = self._style.align(horizontal="center")
            wiersz += 1

        for wartosci in plan.wiersze:
            for i, v in enumerate(wartosci, start=1):
                ws.cell(row=wiersz, column=i, value=v)
            wiersz += 1

        for etykieta, wartosc, fmt in plan.kpi:
            c1 = ws.cell(row=wiersz, column=1, value=etykieta)
            c2 = ws.cell(row=wiersz, column=2, value=wartosc)
            if fmt:
                c2.number_format = fmt
            if self._style:
                c1.font = self._style.font(bold=True)
            wiersz += 1

        if plan.tabela is not None:
            start = wiersz
            for wartosci in plan.tabela:
                for i, v in enumerate(wartosci, start=1):
                    ws.cell(row=wiersz, column=i, value=v)
                wiersz += 1
            n = len(plan.tabela[0])
            ref = f"A{start}:{get_column_letter(n)}{wiersz - 1}"
            tab = Table(displayName=_nazwa_tabeli(plan.nazwa), ref=ref)
            tab.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9", showRowStripes=True
            )
            ws.add_table(tab)

        if plan.szerokosci:
            for i, szer in enumerate(plan.szerokosci, start=1):
                ws.column_dimensions[get_column_letter(i)].width = szer
```

**Użycie — dokładnie jak w zarysie:**

```python
def main() -> None:
    dane = [
        ("2026-01-05", "Alfa sp. z o.o.", 10, 25.0),
        ("2026-01-09", "Beta S.A.",        4, 41.5),
        ("2026-01-17", "Gamma sp. k.",    22, 25.0),
    ]
    zestawienie = [
        ("Region", "Wartosc"),
        ("PL", 1490.0),
        ("DE", 166.0),
    ]

    cel = (
        RaportBuilder("Raport Q1 2026")
        .dodaj_arkusz("Dane", szerokosci=(12, 24, 8, 10))
        .naglowek(["Data", "Klient", "Ilosc", "Cena"])
        .wiersze(dane)
        .dodaj_arkusz("Podsumowanie", szerokosci=(18, 14))
        .kpi("Liczba transakcji", len(dane), "#,##0")
        .kpi("Wartosc lacznie", sum(r[2] * r[3] for r in dane), '#,##0.00 "zl"')
        .tabela(zestawienie)
        .zapisz("output/23_02_builder.xlsx")
    )
    print("Zapisano:", cel)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Do momentu `.zapisz(...)` **nie istnieje ani jeden obiekt openpyxl**. Cała praca to dopisywanie do list `_PlanArkusza`. To znaczy, że przez cały czas budowy możesz:
- **zapytać o plan** (`builder._arkusze` — do diagnostyki),
- **wycofać się** (po prostu nie wywołuj `build()`),
- **zmienić zdanie** w środku,
- **przetestować kompletność** bez tworzenia skoroszytu.

I najważniejsze: **walidacja jest przed pracą, nie po niej.** Jeżeli podasz 3 kolumny w nagłówku i 4-elementowe wiersze, `build()` podniesie `BladBudowy` z komunikatem `"arkusz 'Dane': 3 wiersz(y) maja inna liczbe kolumn niz naglowek (3); pierwszy zly: 1"` **i nie powstanie ani jedna komórka**. Zamiast 500 wierszy w Excelu z przesuniętymi kolumnami.

**Co trafi do pliku.** Dwa arkusze. W `Dane` — nagłówek z żółtym wypełnieniem i wyśrodkowaniem, trzy wiersze danych, szerokości kolumn. W `Podsumowanie` — dwa wiersze KPI (z formatem walutowym i tysięcznym) oraz **prawdziwa tabela Excela** (`Table`) na zakresie zestawienia, z nazwą `T_Podsumowanie` (bo `_nazwa_tabeli` zamienia polskie znaki) i stylem `TableStyleMedium9`. Do `xl/tables/table1.xml` trafi definicja tabeli z `displayName="T_Podsumowanie"` i `ref="A4:B6"` — a do `xl/workbook.xml` relacja do niej.

**Trzy rzeczy do przemyślenia:**

1. **Walidujesz w `build()`, ale wyjątek ma listę problemów, nie pierwszy z nich.** To celowe. Użytkownik, który dostał pięć błędów naraz, naprawi je w jednym przebiegu. Użytkownik, który dostaje błąd po błędzie, naprawi ich pięć w pięciu przebiegach.
2. **`_nazwa_tabeli` istnieje, bo Excel ma swoje zasady, o których openpyxl nie ostrzeże.** Nazwa tabeli z odstępem przejdzie przez openpyxl bez mrugnięcia okiem, a Excel zgłosi „uszkodzony plik". To ten sam mechanizm co `quote_sheetname` w module 07: **biblioteka nie waliduje wszystkiego — walidacja jest Twoją pracą.**
3. **`build()` zwraca `Workbook`, a nie plik.** Zapamiętaj tę decyzję, bo w module 24 stanie się regułą: **granica między „zbuduj" a „zapisz" to granica między testowalnym a nietestowalnym.**

### Przykład 3 — Prototype: szablon, który nigdy nie jest nadpisywany

**Problem.** Raport ma logo, firmowy nagłówek, ustawienia wydruku i motyw, który dział marketingu robił w Excelu dwa dni. Odtworzenie tego kodem to 200 linii i szansa, że wyjdzie „prawie tak samo".

**Rozwiązanie.** Szablon jako pieczątka.

```python
"""Modul 23, przyklad 3: Prototype - szablon skoroszytu.

Pokazuje TRZY rzeczy:
  1. jak uzyc szablonu (i ZE NIGDY nie jest on celem zapisu),
  2. ze copy_worksheet GUBI obrazy i wykresy,
  3. jak wstawic logo programowo, gdy trzeba kopiowac arkusze.

Uruchom: python examples/23_03_prototype.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import Font

KATALOG = Path("output")
SZABLON = KATALOG / "23_03_szablon.xlsx"
LOGO = Path("assets/logo.png")          # jesli nie istnieje, pomijamy obraz


# =====================================================================
# 0. Przygotowanie szablonu. W PRAWDZIWYM projekcie robi sie to RECZNIE
#    w Excelu i wrzuca do repozytorium. Tu - programowo, zeby przyklad
#    byl samowystarczalny.
# =====================================================================
def przygotuj_szablon(cel: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"

    ws["A1"] = "RAPORT SPRZEDAZY"
    ws["A1"].font = Font(bold=True, size=18, color="1072BA")
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 28

    naglowki = ["Data", "Klient", "Region", "Produkt", "Ilosc", "Cena"]
    for i, etykieta in enumerate(naglowki, start=1):
        ws.cell(row=3, column=i, value=etykieta).font = Font(bold=True)

    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    ws.print_area = "A1:F40"
    ws.page_setup.orientation = "landscape"

    if LOGO.exists():
        wstaw_logo(ws, LOGO, "H1", wysokosc_px=48)

    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)
    return cel


def wstaw_logo(ws, sciezka: Path, komorka: str, wysokosc_px: int = 48) -> None:
    """Wstawia obraz, skalujac go Z ZACHOWANIEM PROPORCJI.

    Uwaga na jednostki: img.width/img.height sa w PIKSELACH, a nie w punktach.
    """
    img = XlImage(str(sciezka))            # wymaga Pillow: pip install pillow
    skala = wysokosc_px / img.height
    img.height = int(img.height * skala)
    img.width = int(img.width * skala)
    ws.add_image(img, komorka)
```

```python
# =====================================================================
# 1. PROTOTYPE - poprawnie: szablon ONLY DO ODCZYTU, wynik to NOWY plik
# =====================================================================
def raport_z_szablonu(szablon: Path, cel: Path, dane) -> Path:
    if Path(szablon).resolve() == Path(cel).resolve():
        raise ValueError(
            "Szablon nie moze byc celem zapisu. To jest regula, nie zyczenie."
        )

    wb = load_workbook(szablon)            # wzorzec trafia do PAMIECI
    ws = wb["Dane"]                        # uzywamy ISTNIEJACEGO arkusza

    wiersz = 4
    for w in dane:
        ws.cell(row=wiersz, column=1, value=w["data"]).number_format = "dd.mm.yyyy"
        ws.cell(row=wiersz, column=2, value=w["klient"])
        ws.cell(row=wiersz, column=3, value=w["region"])
        ws.cell(row=wiersz, column=4, value=w["produkt"])
        ws.cell(row=wiersz, column=5, value=w["ilosc"])
        ws.cell(row=wiersz, column=6, value=w["cena"]).number_format = '#,##0.00 "zl"'
        wiersz += 1

    cel = Path(cel)
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)                           # NOWY plik; szablon nietkniety
    return cel


# =====================================================================
# 2. ANTYWZORZ: copy_worksheet GUBI obrazy i wykresy
# =====================================================================
def raport_przez_kopie(szablon: Path, cel: Path, regiony) -> Path:
    wb = load_workbook(szablon)
    wzor = wb["Dane"]

    print(f"  we wzorcu obrazow: {len(wzor._images)}   <- atrybut prywatny, "
          "wylacznie do diagnostyki!")

    for region in regiony:
        kopia = wb.copy_worksheet(wzor)
        kopia.title = region
        print(f"  w kopii {region!r} obrazow: {len(kopia._images)}")

    cel = Path(cel)
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)
    return cel


# =====================================================================
# 3. POPRAWNIE, gdy MUSISZ kopiowac: wstaw logo programowo po kopii
# =====================================================================
def raport_z_kopiami_i_logo(szablon: Path, cel: Path, regiony) -> Path:
    wb = load_workbook(szablon)
    wzor = wb["Dane"]

    for region in regiony:
        kopia = wb.copy_worksheet(wzor)
        kopia.title = region
        if LOGO.exists():
            wstaw_logo(kopia, LOGO, "H1", wysokosc_px=48)   # ← rekompensata

    cel = Path(cel)
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)
    return cel


def main() -> None:
    dane = [
        {"data": "2026-01-05", "klient": "Alfa sp. z o.o.", "region": "PL",
         "produkt": "Widzet A", "ilosc": 10, "cena": 25.0},
        {"data": "2026-01-09", "klient": "Beta S.A.", "region": "DE",
         "produkt": "Widzet B", "ilosc": 4, "cena": 41.5},
    ]

    przygotuj_szablon(SZABLON)
    print("Szablon:", SZABLON)

    print("\n--- Prototype (poprawnie) ---")
    print(" Wynik:", raport_z_szablonu(SZABLON, KATALOG / "23_03_raport.xlsx", dane))

    print("\n--- Antywzorz: copy_worksheet ---")
    raport_przez_kopie(SZABLON, KATALOG / "23_03_kopie_bez_logo.xlsx", ["PL", "DE"])

    print("\n--- Kopie + logo programowo ---")
    raport_z_kopiami_i_logo(SZABLON, KATALOG / "23_03_kopie_z_logo.xlsx", ["PL", "DE"])


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `load_workbook(SZABLON)` czyta ZIP, parsuje XML-e i buduje **pełny model w pamięci**: komórki, style, ustawienia strony, i — jeżeli istnieje — obrazy (`ws._images`). Ten model jest **kopią**, nie widokiem na plik. Dlatego `wb.save(cel)` zapisuje kopię **pod inną nazwą** i szablon na dysku pozostaje absolutnie nietknięty.

Warto zwrócić uwagę na jedną rzecz z testu diagnostycznego: `wzor._images` ma **jedną** pozycję, a `kopia._images` ma **zero**. To nie jest „bug openpyxl" — to udokumentowane ograniczenie `copy_worksheet`, o którym mówi moduł 08. Jawnie używam tu atrybutu prywatnego `_images` i **nazwę to prywatnym atrybutem** — w kodzie produkcyjnym zamiast tego napisałbyś test oparty na strukturze pliku (`xl/media/`) albo na liczbie wpisów w `xl/drawings/_rels`. Ale do jednorazowej diagnostyki to jest najszybsza droga i warto wiedzieć, że istnieje.

**Co trafi do pliku.** Trzy pliki wynikowe:

| Plik | Arkusze | Logo |
|---|---|---|
| `23_03_raport.xlsx` | `Dane` | **jest** (nikt go nie kopiował) |
| `23_03_kopie_bez_logo.xlsx` | `Dane`, `PL`, `DE` | tylko w `Dane` |
| `23_03_kopie_z_logo.xlsx` | `Dane`, `PL`, `DE` | we wszystkich trzech |

I jeszcze jedna rzecz, o której trzeba powiedzieć wprost: **trzy pliki wynikowe różnią się od szablonu zawartością `xl/media/`**. W pierwszym pliku są dwa obrazki (szablonowy + nic), w drugim — jeden obraz, w trzecim — trzy. `xl/drawings/drawing1.xml` opisuje anchory. To pokazuje, że **obraz w openpyxl nie jest „w komórce" — jest przypięty do siatki**, więc ma osobną pozycję w pliku i nie „przesuwa się" razem z komórkami przy sortowaniu (moduł 16).

**Trzy rzeczy do przemyślenia:**

1. **Reguła „szablon nigdy nie jest nadpisywany" jest w kodzie, nie w głowie.** `if szablon.resolve() == cel.resolve(): raise` — dwie linie, które chronią najdroższy plik w projekcie. **Reguła, której nie ma w kodzie, jest życzeniem.**
2. **Szablon, który przechodzi przez openpyxl, może stracić treść.** Jeżeli dział marketingu wstawi do szablonu **sparkline'y** albo **pole tekstowe**, przy `load_workbook` openpyxl ostrzeże („Sparkline Group extension is not supported and will be removed"), a przy `save` te elementy **znikną z pliku wynikowego**. Dlatego szablon trzeba **przetestować** — wczytaj go i zapisz na kopię, otwórz w Excelu i **obejrzyj**. To jest obowiązek, nie ostrożność.
3. **Makra (`keep_vba`) wymagają spójnego rozszerzenia.** Jeżeli szablon jest `.xlsm`, wczytaj go z `keep_vba=True` i **zapisz z rozszerzeniem `.xlsm`**. openpyxl VBA tylko zachowuje — nie edytuje — a zmiana rozszerzenia na `.xlsx` przy zachowanym `vbaProject.bin` prowadzi do pliku, z którym Excel nie wie, co zrobić.

### Przykład 4 — Registry i Flyweight: zmierz, nie zgaduj

**Problem.** Budujemy raport 20 000 komórek i nie wiemy, ile kosztuje tworzenie stylu „w locie". **Podejrzewamy**, że dużo. Sprawdźmy.

```python
"""Modul 23, przyklad 4: StyleRegistry (Flyweight) - pomiar.

Cel: DOWIESC, a nie WIERZYC. Pokazujemy, ze:
  - bez rejestru powstaje N obiektow Font,
  - z rejestrem powstaje K obiektow (K << N),
  - plik wyjsciowy jest praktycznie IDENTYCZNY (bo openpyxl dedupuje
    style w styles.xml przy zapisie).

Uruchom: python examples/23_04_registry.py
"""

from __future__ import annotations

import time
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

KOLORY = ["FFC000", "1072BA", "70AD47", "C00000", "FFFFFF"]
ROZMIARY = [8, 9, 10, 11, 12, 14, 16]
N = 20_000


class StyleRegistry:
    """Rejestr = cache + fabryka. WSTRZYKIWANY, nie globalny (sekcja 2.5)."""

    def __init__(self) -> None:
        self._fonts: dict[tuple, Font] = {}
        self.utworzone = 0

    def font(self, **cechy) -> Font:
        klucz = tuple(sorted(cechy.items()))
        gotowy = self._fonts.get(klucz)
        if gotowy is None:
            gotowy = Font(**cechy)
            self._fonts[klucz] = gotowy
            self.utworzone += 1
        return gotowy

    def ile_w_cache(self) -> int:
        return len(self._fonts)


def bez_rejestru(ws) -> int:
    licznik = 0
    for i in range(1, N + 1):
        c = ws.cell(row=i, column=1, value=i)
        c.font = Font(               # ← NOWY obiekt w KAZDEJ iteracji
            bold=True,
            color=KOLORY[i % len(KOLORY)],
            size=ROZMIARY[i % len(ROZMIARY)],
        )
        licznik += 1
    return licznik


def z_rejestrem(ws, rejestr: StyleRegistry) -> int:
    for i in range(1, N + 1):
        c = ws.cell(row=i, column=1, value=i)
        c.font = rejestr.font(       # ← obiekt z cache
            bold=True,
            color=KOLORY[i % len(KOLORY)],
            size=ROZMIARY[i % len(ROZMIARY)],
        )
    return rejestr.utworzone


def zmierz(etykieta: str, funkcja):
    wb = Workbook()
    ws = wb.active
    start = time.perf_counter()
    wynik = funkcja(ws)
    czas = time.perf_counter() - start

    cel = Path(f"output/23_04_{etykieta}.xlsx")
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)

    # style "wewnatrz" openpyxl - atrybut prywatny, tylko diagnostyka
    print(
        f"{etykieta:12s} | {czas:6.3f} s | obiektow Font: {wynik:6d} "
        f"| fonts w wb: {len(wb._fonts):3d} | plik: {cel.stat().st_size/1024:7.1f} kB"
    )
    return czas


def main() -> None:
    rejestr = StyleRegistry()
    print(f"Komorki: {N} | unikalnych kombinacji stylu: "
          f"{len(KOLORY) * len(ROZMIARY)}\n")
    zmierz("bez", bez_rejestru)
    zmierz("z", lambda ws: z_rejestrem(ws, rejestr))
    print(f"\nW cache rejestru: {rejestr.ile_w_cache()} obiektow Font.")


if __name__ == "__main__":
    main()
```

**Czego się spodziewać.** Konkretne liczby zależą od sprzętu, ale **kształt wyniku jest zawsze taki sam**:

```text
Komorki: 20000 | unikalnych kombinacji stylu: 35

bez          |  1.850 s | obiektow Font:  20000 | fonts w wb:  35 | plik:   218.4 kB
z            |  1.120 s | obiektow Font:     35 | fonts w wb:  35 | plik:   218.4 kB

W cache rejestru: 35 obiektow Font.
```

Trzy rzeczy, które z tego wynikają i które są **sedno tego przykładu**:

1. **Rozmiar pliku jest identyczny.** `fonts w wb: 35` w obu przypadkach — bo openpyxl przy zapisie indeksuje style i **deduplikuje** je w `styles.xml`. Twoja „optymalizacja" **nie zmieniła pliku**. Zmieniła tylko pracę Pythona.
2. **Liczby `20000` vs `35` pokazują, że problem był realny.** 19 965 obiektów `Font` powstało niepotrzebnie — każdy z nich musiał zostać policzony (`__hash__`) i porównany. To jest ten „koszt 2" z sekcji 1.2.
3. **Zysk czasu jest realny, ale nie dramatyczny** przy 20 000 komórek (u mnie ~40%). Przy 500 000 komórek zaczyna być wyraźny. **Uczciwie: to nie jest optymalizacja, która ratuje projekt. To higiena.** Wartość rejestru jest większa w tym, że **style mają nazwy i pochodzą z motywu** — a to jest już problem poprawności, nie wydajności.

**Co się dzieje w pamięci — i gdzie rejestr żyje.** `StyleRegistry` to instancja, którą **tworzysz i przekazujesz**. W `main()` jest jedna, ale **nie jest globalna**. To znaczy, że w testach możesz mieć drugą:

```python
def test_motyw_zastepczy():
    rejestr_testowy = StyleRegistry()
    # ... i rejestr produkcyjny NIE zostanie zanieczyszczony
```

To jest dokładnie ta różnica, o której mówiła sekcja 2.5. Wyobraź sobie, że `StyleRegistry` jest singletonem: powyższy test zmieniłby globalny stan i **wpłynąłby na inne testy**. To najgorszy rodzaj błędu — przechodzi u Ciebie, pada w CI.

### Przykład 5 — Konfiguracja, DTO i leniwość razem

Ostatni przykład składa wszystko: `RaportSpec` (konfiguracja), `WierszRaportu` (DTO z walidacją), `LeniweArkusze` (leniwe tworzenie) i fabrykę. To jest wersja, którą **naprawdę** napisałbyś w projekcie.

```python
"""Modul 23, przyklad 5: RaportSpec + WierszRaportu + leniwosc.

Trzy wzorce spotykaja sie w jednym miejscu:
  - RaportSpec  -> konfiguracja zamiast kodu (przepis vs danie),
  - WierszRaportu -> DTO z walidacja w konstruktorze,
  - LeniweArkusze -> leniwe tworzenie arkuszy.

Uruchom: python examples/23_05_spec.py
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =====================================================================
# 1. DTO - nosnik danych z walidacja
# =====================================================================
@dataclass(frozen=True)
class WierszRaportu:
    data: date
    klient: str
    region: str
    ilosc: int
    cena: float

    def __post_init__(self) -> None:
        if self.ilosc < 0:
            raise ValueError(f"ilosc nie moze byc ujemna: {self.ilosc!r}")
        if not self.klient.strip():
            raise ValueError("klient nie moze byc pusty")
        if self.cena < 0:
            raise ValueError(f"cena nie moze byc ujemna: {self.cena!r}")

    @property
    def wartosc(self) -> float:
        return self.ilosc * self.cena


# =====================================================================
# 2. KONFIGURACJA - raport opisany jako DANE
# =====================================================================
@dataclass(frozen=True)
class Kolumna:
    naglowek: str
    szerokosc: float
    pole: str                       # nazwa atrybutu w WierszRaportu
    format_liczby: str | None = None


@dataclass(frozen=True)
class RaportSpec:
    nazwa: str
    tytul: str
    arkusze: tuple[str, ...]
    kolumny: tuple[Kolumna, ...]
    tytul_arkusza: str = "Dane"


KOLUMNY_SPRZEDAZ = (
    Kolumna("Data",    12, "data",   "dd.mm.yyyy"),
    Kolumna("Klient",  24, "klient"),
    Kolumna("Region",   8, "region"),
    Kolumna("Ilosc",    8, "ilosc",  "#,##0"),
    Kolumna("Cena",    10, "cena",   '#,##0.00 "zl"'),
)

SPEC_SPRZEDAZ = RaportSpec(
    nazwa="sprzedaz",
    tytul="Sprzedaz - styczen 2026",
    arkusze=("dane", "podsumowanie"),
    kolumny=KOLUMNY_SPRZEDAZ,
)

SPEC_ALERTY = RaportSpec(
    nazwa="alerty",
    tytul="Transakcje wymagajace uwagi",
    arkusze=("dane",),
    kolumny=(
        Kolumna("Klient",  24, "klient"),
        Kolumna("Region",   8, "region"),
        Kolumna("Cena",    10, "cena", '#,##0.00 "zl"'),
    ),
    tytul_arkusza="Alerty",
)
```

```python
# =====================================================================
# 3. LENIWE ARKUSZE - tworzymy dopiero, gdy mamy dane
# =====================================================================
class LeniweArkusze:
    """Object Pool w miniaturze: arkusz powstaje przy pierwszym uzyciu.

    Sens: raport ma 30 potencjalnych arkuszy (po jednym na region), ale
    w danym miesiacu dane sa tylko dla 12. Reszta ma NIE POWSTAC.
    """

    def __init__(self, wb: Workbook) -> None:
        self._wb = wb
        self._gotowe: dict[str, object] = {}

    def arkusz(self, nazwa: str, szerokosci=()):
        ws = self._gotowe.get(nazwa)
        if ws is None:
            ws = self._wb.create_sheet(title=nazwa)
            for i, szer in enumerate(szerokosci, start=1):
                ws.column_dimensions[get_column_letter(i)].width = szer
            self._gotowe[nazwa] = ws
        return ws

    def ile_utworzono(self) -> int:
        return len(self._gotowe)


# =====================================================================
# 4. GENERATOR - czyta spec i DTO, nic wiecej
# =====================================================================
def generuj(wiersze, spec: RaportSpec) -> Workbook:
    wb = Workbook()
    del wb["Sheet"]
    leniwe = LeniweArkusze(wb)

    szerokosci = tuple(k.szerokosc for k in spec.kolumny)
    ws = leniwe.arkusz(spec.tytul_arkusza, szerokosci)

    ws.cell(row=1, column=1, value=spec.tytul).font = Font(bold=True, size=16)
    for i, kolumna in enumerate(spec.kolumny, start=1):
        ws.cell(row=3, column=i, value=kolumna.naglowek).font = Font(bold=True)

    for i, w in enumerate(wiersze, start=4):
        for j, kolumna in enumerate(spec.kolumny, start=1):
            wartosc = getattr(w, kolumna.pole)
            c = ws.cell(row=i, column=j, value=wartosc)
            if kolumna.format_liczby:
                c.number_format = kolumna.format_liczby

    ws.freeze_panes = "A4"
    wb.properties.title = spec.tytul
    return wb
```

I jeszcze **arkusze per region, tworzone leniwie** — tu widać sens wzorca:

```python
def generuj_per_region(wiersze, regiony_z_danymi) -> Workbook:
    """30 potencjalnych regionow, ale tylko te z danymi dostana arkusz."""
    wb = Workbook()
    del wb["Sheet"]
    leniwe = LeniweArkusze(wb)

    for region in regiony_z_danymi:            # ← TYLKO te, ktore maja dane
        ws = leniwe.arkusz(region, szerokosci=(14, 24, 10))
        ws.cell(row=1, column=1, value=f"Region {region}").font = Font(bold=True)
        for i, w in enumerate(
            (w for w in wiersze if w.region == region), start=3
        ):
            ws.cell(row=i, column=1, value=w.data).number_format = "dd.mm.yyyy"
            ws.cell(row=i, column=2, value=w.klient)
            ws.cell(row=i, column=3, value=w.wartosc).number_format = '#,##0.00 "zl"'

    print(f"Utworzono arkuszy: {leniwe.ile_utworzono()} z 30 mozliwych")
    return wb


def main() -> None:
    wiersze = [
        WierszRaportu(date(2026, 1, 5), "Alfa sp. z o.o.", "PL", 10, 25.0),
        WierszRaportu(date(2026, 1, 9), "Beta S.A.", "DE", 4, 41.5),
        WierszRaportu(date(2026, 1, 17), "Gamma sp. k.", "PL", 22, 25.0),
        WierszRaportu(date(2026, 1, 23), "Alfa sp. z o.o.", "PL", 7, 63.0),
    ]

    katalog = Path("output")
    katalog.mkdir(parents=True, exist_ok=True)

    # Raport 1: pelna specyfikacja
    generuj(wiersze, SPEC_SPRZEDAZ).save(katalog / "23_05_sprzedaz.xlsx")

    # Raport 2: NOWA specyfikacja, ZERO nowego kodu w `generuj`
    alerty = [w for w in wiersze if w.cena > 30]
    generuj(alerty, SPEC_ALERTY).save(katalog / "23_05_alerty.xlsx")

    # Raport 3: leniwe arkusze per region
    generuj_per_region(wiersze, ["PL", "DE"]).save(katalog / "23_05_regiony.xlsx")

    # DTO waliduje w konstruktorze - blad powstaje TU, nie w Excelu
    try:
        WierszRaportu(date(2026, 1, 1), "", "PL", 5, 10.0)
    except ValueError as blad:
        print("Zlapano przy tworzeniu obiektu:", blad)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Trzy rzeczy warte nazwania.

1. **`WierszRaportu(..., ilosc=-1, ...)` rzuca wyjątek natychmiast** — w momencie tworzenia obiektu, jeszcze przed jakimkolwiek przetwarzaniem. To jest sedno DTO: **błędne dane nie wchodzą do systemu.** Gdyby to był dict, `ilosc=-1` przejechałoby przez cały pipeline i wylądowało w Excelu albo (gorzej) w formule.

2. **`getattr(w, kolumna.pole)`** — tu wracamy do świadomego kompromisu z modułu 22 (sekcja o przykładzie 2). Kolumna wskazuje pole po nazwie, więc kolumna **nie może być obliczona**. To jest spójne: `SPEC_ALERTY` i `SPEC_SPRZEDAZ` używają tylko pól surowych. Gdybyś chciał „Wartosc = ilosc × cena" w specyfikacji, musiałbyś przejść na `Callable` (jak w module 22) **albo** dodać `@property` i wskazać je po nazwie — `getattr` zadziała na `property`! To subtelna i ważna własność: `getattr(w, "wartosc")` **wywołuje** `@property`. Więc `Kolumna("Wartosc", 12, "wartosc", '#,##0.00 "zl"')` **działa**. Ograniczenie nie jest tak twarde, jak wyglądało.

3. **`LeniweArkusze` tworzy arkusz przy pierwszym `arkusz(nazwa)`.** Dla trzech wierszy danych różnicy nie zobaczysz. Ale w scenariuszu „30 potencjalnych regionów, dane w 12" zobaczysz **12 arkuszy zamiast 30** — a przy `write_only` to jest różnica między raportem, który się generuje, a takim, który się nie generuje.

**Co trafi do pliku.** Trzy skoroszyty. `23_05_sprzedaz.xlsx` ma arkusz `Dane` z pięcioma kolumnami, `23_05_alerty.xlsx` — arkusz `Alerty` z trzema, `23_05_regiony.xlsx` — arkusze `PL` i `DE`. **Zero zmian w funkcji `generuj`** między raportem pierwszym a drugim.

**Trzy rzeczy do przemyślenia:**

1. **`RaportSpec.arkusze = ("dane", "podsumowanie")` w `SPEC_SPRZEDAZ` jest nieużywane w `generuj`.** To jest błąd w tym przykładzie i zostawiam go celowo. Zauważ go: specyfikacja obiecuje dwa arkusze, a generator tworzy jeden. **To jest dokładnie ten rodzaj rozjazdu, przed którym chroni Builder z przykładu 2** — bo Builder ma `_waliduj()`, a `RaportSpec` sam z siebie nie ma. Wniosek: **konfiguracja nie zastępuje walidacji.** Jeżeli spec mówi „dwa arkusze", musi istnieć coś, co to sprawdzi.
2. **`WierszRaportu` nie wie o openpyxl** i słusznie. Możesz go przekazać do walidacji, do agregacji, do wysyłki mailem — wszędzie. To jest świat danych z modułu 22, wreszcie w formie, z którą nie da się zrobić literówki.
3. **`generuj` nadal ma `getattr` i to jest jego granica.** Pełne rozwiązanie (generyk + protokół) to moduł 24. Ale **nie potrzebujesz go jeszcze** — i to jest cała mądrość modułu 22 w jednym zdaniu.

## 4. Anatomia API

### Część A — konstrukcje Pythona używane w tym module

| Konstrukcja | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `@dataclass(frozen=True)` | niemutowalny obiekt danych | — | Podstawa `Motyw`, `DefinicjaArkusza`, `RaportSpec`, `Kolumnа`. `frozen` chroni **przypisanie** — w polach używaj `tuple`, nie `list` (pułapka z modułu 22) |
| `@dataclass(slots=True)` | `__slots__` automatycznie | — | **Wymaga Pythona 3.10+.** Na 3.9 pomiń albo zdefiniuj `__slots__` ręcznie (przy `frozen` i domyślnych wartościach łatwo się potknąć) |
| `field(default_factory=...)` | domyślna wartość per instancja | `default_factory` | Konieczne dla `list`/`dict` w `_PlanArkusza` |
| `__post_init__` | walidacja po `__init__` | — | Działa też z `frozen=True` — możesz **rzucać wyjątek**, nie możesz przypisywać |
| `@property` | pole wyliczane | — | **`getattr(obj, "wartosc")` wywołuje `@property`** — dlatego `Kolumna(pole="wartosc")` działa |
| `__hash__` / `__eq__` z `frozen=True` | obiekt nadaje się na klucz / do cache | — | Dlatego `DefinicjaArkusza` może być kluczem słownika |
| `**cechy` + `tuple(sorted(cechy.items()))` | klucz cache z argumentów | — | **Wartości muszą być hashowalne** — lista w `color=` wybuchnie `TypeError` |
| `functools.cached_property` | policz raz, zapamiętaj | — | Python 3.8+. Uwaga: trzyma referencję do `self` — nie używaj na obiektach, których chcesz szybko zwalniać |
| `functools.lru_cache` | cache funkcji po argumentach | `maxsize`, `typed` | **Nie owijaj metod** — cache trzyma `self` i wycieka pamięć. Owijaj **funkcje modułowe** |
| `classmethod` | alternatywny konstruktor | — | Klasyczny zapis Factory Method: `RaportSpec.z_pliku_json(...)` |
| `@dataclass` na planie (`_PlanArkusza`) | struktura pośrednia | — | **Plan nie zawiera typów openpyxl** — to warunek testowalności |
| podkreślenie `_nazwa` | konwencja prywatności | — | Python nie egzekwuje; to umowa między Tobą a zespołem |
| `re.sub(r"[^0-9A-Za-z_]", "_", s)` | sanityzacja nazwy | — | Konieczne dla `Table(displayName=...)` i nazw arkuszy |

### Część B — `openpyxl` w kontekście tego modułu

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | nowy skoroszyt | `write_only`, `iso_dates` | W Builderze wywoływane **dopiero w `build()`** |
| `wb.create_sheet()` | nowy arkusz | `title`, `index` | **Jedno miejsce w projekcie:** fabryka. Nazwa ≤31 znaków, bez `: \ / ? * [ ]`, unikalna |
| `wb.remove(ws)` / `del wb["Nazwa"]` | usuń arkusz | — | Uwaga na definiowane nazwy i tabele wskazujące na usunięty arkusz (moduł 08) |
| `wb.copy_worksheet(ws)` | kopia arkusza | — | Kopiuje **wartości, style, hyperlinki, komentarze**. **NIE kopiuje obrazów ani wykresów.** Nie działa między skoroszytami i w `read_only`/`write_only` |
| `ws.sheet_properties.tabColor` | kolor zakładki | string `"RRGGBB"` | Tanie, mocno poprawia UX raportu |
| `ws.sheet_state` | widoczność | `"visible"` \| `"hidden"` \| `"veryHidden"` | `veryHidden` ukrywa też przed oknem „Unhide" — dobre dla arkuszy konfiguracyjnych |
| `ws.sheet_view.showGridLines` | linie siatki | `bool` | Konwencja firmy, nie technika |
| `ws.freeze_panes` | zamrożenie okien | `"A4"` | Wynika z definicji arkusza, nie z ręcznej decyzji przy każdej kolumnie |
| `ws.column_dimensions[litera].width` | szerokość | — | Ustawiane z `DefinicjaArkusza.szerokosci` w pętli |
| `load_workbook(path, keep_vba=, data_only=)` | wczytanie szablonu | `read_only`, `data_only`, `keep_vba`, `keep_links`, `rich_text` | **Prototype.** `.xlsm` wymaga `keep_vba=True` **i** zapisu z `.xlsm` |
| `wb.add_named_style(NamedStyle(...))` | styl nazwany skoroszytu | obiekt `NamedStyle` | Widoczny w **galerii stylów Excela** — użytkownik może go edytować globalnie. Alternatywa dla rejestru, gdy plik będzie edytowany przez ludzi. W 3.1.x API zostało przebudowane — sprawdź w dokumentacji swojej wersji |
| `NamedStyle(name=, font=, fill=, border=, alignment=, number_format=, protection=)` | definicja stylu nazwanego | jak wyżej | **Jest mutowalna.** Nie dodawaj tej samej instancji do dwóch skoroszytów — zobacz pułapkę 5 |
| `cell.style = "NazwaStylu"` | przypisanie stylu nazwanego | string | Nazwa musi istnieć **w tym skoroszycie**; style nazwane są **związane ze skoroszytem** |
| `Font(**cechy)` | czcionka | `bold`, `color`, `size`, `italic`, `name`, `underline`, … | **Niemutowalna w użyciu.** Buduj nowy obiekt albo `copy()` (moduł 09) |
| `PatternFill(fill_type=, fgColor=)` | wypełnienie | `fill_type`, `fgColor`, `bgColor` | Dla `solid` kolor tła to **`fgColor`** |
| `Alignment(**cechy)` | wyrównanie | `horizontal`, `vertical`, `wrap_text`, … | — |
| `Font.__hash__` | hashowanie po cechach | — | Dlatego `IndexedList` w openpyxl deduplikuje style przy zapisie |
| `Image(path)` + `ws.add_image(img, anchor)` | obraz | `str`/plikopodobny, kotwica | **Wymaga `pillow`.** Jednostki to **piksele**, nie punkty |
| `Table(displayName=, ref=)` + `TableStyleInfo(...)` | tabela Excela | `ref` **z nagłówkiem**, `headerRowCount` (domyślnie 1) | `displayName` bez spacji i znaków specjalnych. `tableStyleInfo` ustaw **po** konstrukcji |
| `wb.properties.title` / `.creator` | metadane | — | Ślad audytowy. Ustawiaj **w jednym miejscu** — w generatorze lub Builderze |
| `wb._fonts`, `ws._images` | **atrybuty prywatne** | — | **Wyłącznie diagnostyka.** Nie buduj na nich logiki produkcyjnej |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — fabryka nagłówków arkuszy.**

Cel: poczuć, że fabryka to **jedna decyzja w jednym miejscu**.

Weź `examples/23_01_fabryka.py`. Rozbuduj `ArkuszFabryka` o metodę **`naglowek`**, która pisze wiersz nagłówków zgodnie z konwencją firmy:

```python
def naglowek(
    self,
    ws,
    etykiety,
    wiersz: int = 1,
    wersaliki: bool = True,
    wyrownanie: str = "center",
) -> None:
    """Pisze naglowek zgodnie z konwencja firmy (dane, nie kod)."""
```

**Wymagania:**

1. Nagłówek zawsze w **pogrubionej** czcionce z motywu (`style.font_naglowka()`).
2. Zawsze z wypełnieniem z motywu (`style.fill_naglowka()`).
3. `wersaliki=True` → `"Data"` staje się `"DATA"`.
4. Wyrównanie domyślnie `"center"`.
5. **Zero literałów** w tej metodzie — czcionka, kolor i wyrównanie pochodzą ze `StylFabryka`.
6. Metoda **nie wie**, co to za arkusz. Nie ma `if klucz == "dane"`.

**Kryteria odbioru:**

- `grep -n '"FFC000"\|"FFFFFF"\|"FF0000"' examples/23_01_fabryka.py` → trafienia **tylko** w `Motyw`.
- `grep -n 'Font(' examples/23_01_fabryka.py` → trafienia **tylko** w `StylFabryka.font`.
- Uruchomienie skryptu dwa razy z rzędu daje ten sam plik (idempotencja, moduł 18).

**W pliku `output/23_cw1_wnioski.md` odpowiedz:**

1. **Ile miejsc zmieniłbyś, żeby nagłówki były pisane wersalikami na czerwono?** (Policz linie.)
2. **Ile miejsc byś zmienił w wersji, w której każdy raport pisze nagłówek sam?** (Napisz wersję bez fabryki — dosłownie — i policz.)
3. **Czy `naglowek` powinno być w `ArkuszFabryka`?** Odpowiedz, używając trzech pytań z modułu 22 (sekcja 2.5). Czy zapis nagłówka jest częścią „tworzenia arkusza", czy osobną odpowiedzialnością? Gdzie postawiłbyś granicę — i **dlaczego to jest pytanie o SRP**?

### 🟡 Warsztat

**Zadanie 2 — Builder: raport trzyarkuszowy z walidacją.**

Rozbuduj `RaportBuilder` z przykładu 2 tak, żeby obsłużył raport z **trzema arkuszami** i **kontrolował kompletność na poziomie biznesowym**.

**Wymagania:**

1. **Trzy arkusze:** `Dane` (tabela z danymi), `Podsumowanie` (KPI + tabela zestawienia), `Metodyka` (kilka akapitów tekstu wyjaśniającego, jak policzono raport).

2. **Nowa metoda `.akapit(tekst: str, pogrubiony: bool = False)`** — dopisuje akapit tekstu w bieżącym arkuszu.

3. **Walidacja biznesowa w `_waliduj()`** — nowa metoda `_waliduj_merytorycznie()`, która sprawdza:
   - **`Dane` ma co najmniej 3 wiersze** (raport z 1–2 wierszami to zwykle błąd doboru danych),
   - **`Podsumowanie` ma co najmniej 2 KPI** (inaczej jest bezużyteczne),
   - **`Metodyka` ma co najmniej 1 akapit** (raport bez metodyki jest nieaudytowalny),
   - jeżeli arkusz `Podsumowanie` istnieje, **suma KPI „Wartosc lacznie" zgadza się z liczbą** podaną jawnie — dodaj metodę `.oczekiwana_wartosc(amount: float)` i sprawdź, czy różnica jest mniejsza niż 0,01.

To ostatnie jest najważniejsze: **walidacja merytoryczna, nie techniczna.** Sprawdzasz nie to, czy Excel to przyjmie, a to, czy **raport mówi prawdę**.

4. **`build()` musi wywoływać OBIE walidacje** i mieć komunikat z sekcjami: osobno problemy techniczne, osobno merytoryczne.

**Kryteria odbioru:**

- Test, że brak `Metodyka` daje `BladBudowy` z komunikatem zawierającym `"Metodyka"`.
- Test, że `build()` **nie tworzy żadnego obiektu openpyxl**, jeśli walidacja padnie. Sprawdź to tak:

```python
def test_brak_metodyki_nie_tworzy_workbooka(monkeypatch):
    """Jesli walidacja pada, Workbook NIE MOZE powstac."""
    wywolania = []
    prawdziwy = Workbook

    def podszywacz(*args, **kwargs):
        wywolania.append(args)
        return prawdziwy(*args, **kwargs)

    monkeypatch.setattr("openpyxl.Workbook", podszywacz)
    with pytest.raises(BladBudowy):
        RaportBuilder("X").dodaj_arkusz("Dane").naglowek(["A"]).wiersze([]).build()
    assert wywolania == []
```

Ten test jest dowodem, że Builder naprawdę odkłada materializację — a nie tylko tak wygląda.

- Pomiar: **ile czasu zajmuje `build()` bez zapisu** przy 50 000 wierszach? Porównaj z pełnym cyklem `build() + save()`. Zapisz liczby.

**W pliku `output/23_cw2_wnioski.md` odpowiedz:**

1. **Czy `_waliduj_merytorycznie` należy do Buildera?** A może do świata danych (moduł 22)? Uzasadnij, gdzie jest naturalne miejsce dla „raport musi mieć min. 3 wiersze".
2. **Co jest trudniejsze: walidacja techniczna czy merytoryczna?** Dlaczego merytoryczna jest trudniejsza do napisania i **dlaczego mimo to trzeba ją pisać**?
3. **Ile linii ma `RaportBuilder`, a ile miałby równoważny kod bez Buildera?** Wpisz obie liczby. Policzyłeś koszt i korzyść — a teraz pytanie: **przy jakiej liczbie arkuszy Builder przestaje się opłacać?**
4. **Test `monkeypatch` — czy naprawdę udowadnia, że `build()` nic nie tworzy?** Znajdź w tym teście słaby punkt (podpowiedź: `monkeypatch.setattr("openpyxl.Workbook", ...)` patchuje **nazwę w module**, a `RaportBuilder` zaimportował `Workbook` **do swojego modułu**). Napraw test.

### 🔴 Wyzwanie

**Zadanie 3 — Prototype: szablon + logo + tabela, zapis jako nowy plik.**

Największe wyzwanie tego modułu, bo łączy trzy mechanizmy i **wymaga sprawdzenia w Excelu**.

**Krok 1 — szablon.**

Zbuduj szablon, który zawiera:
- **logo** (plik `assets/logo.png`; jeżeli go nie masz — wygeneruj PNG w kodzie, na przykład `matplotlib` albo `PIL`),
- **firmowy nagłówek** w wierszu 1 (scalony, duża czcionka, kolor firmowy),
- **wiersz nagłówków kolumn** w wierszu 3, gotowy do uzupełnienia,
- **ustawienia wydruku**: A4 poziomo, `print_area`, powtarzanie nagłówków wierszy 1–3 (`print_title_rows = "1:3"`),
- **`freeze_panes = "A4"`**, wyłączone linie siatki, **niebieska zakładka**,
- **arkusz `_Konfiguracja`** ze `sheet_state = "veryHidden"` i wersją szablonu (np. `"wersja_szablonu"` / `"1.0"`).

Zapisz szablon jako `output/23_cw3_szablon.xlsx` **jednym skryptem** `przygotuj_szablon()`.

**Krok 2 — kompletny generator na szablonie.**

Napisz `raport_z_szablonu(szablon, cel, wiersze, spec)`, który:
- **odmawia** zapisu do szablonu (`raise ValueError`),
- wczytuje szablon,
- **wypełnia istniejący arkusz** `Dane` (nie kopiuje go),
- **dodaje nowy arkusz** `Podsumowanie` **z tabelą Excela** (`Table` + `TableStyleInfo`),
- **dopisuje logo do arkusza `Podsumowanie`** programowo (bo nowy arkusz powstał w kodzie i logo w nim nie ma),
- zapisuje jako **nowy plik** z datą w nazwie: `raport_2026-01.xlsx`,
- **zapisuje atomowo** (moduł 18: `mkstemp` + `os.replace`),
- zwraca `Path`.

**Krok 3 — test „nie zniszczyłem szablonu".**

To najważniejszy krok. Napisz test, który:
1. wylicza **odcisk struktury szablonu** przed użyciem (`hashlib.sha256` z posortowanej listy nazw części ZIP-a: `zipfile.ZipFile(szablon).namelist()`, plus rozmiar pliku),
2. uruchamia `raport_z_szablonu(...)`,
3. wylicza odcisk szablonu **po** użyciu,
4. **asyeruje, że odciski są identyczne.**

To jest test, który **fizycznie udowadnia**, że nie nadpisałeś wzorca. Bez niego „nie nadpisuję szablonu" jest obietnicą.

**Krok 4 — test utraty logo przy `copy_worksheet`.**

Napisz test, który:
- wczytuje szablon i **kopiuje** arkusz `Dane` przez `wb.copy_worksheet(...)`,
- zapisuje wynik,
- otwiera wynik jako ZIP i **sprawdza liczbę plików w `xl/media/`**,
- porównuje ją z liczbą w szablonie,
- **asyeruje różnicę** (kopia ma mniej, bo logo w niej nie ma) — **i dodaje komentarz w kodzie**, że to jest udokumentowane zachowanie openpyxl, nie błąd.

**Krok 5 — wersjonowanie szablonu.**

W `_Konfiguracja` szablonu trzymaj `wersja_szablonu`. W `raport_z_szablonu` sprawdź tę wersję i **rzuć wyjątek, jeśli jest inna niż oczekiwana**. Uzasadnij w komentarzu, dlaczego to jest zabezpieczenie przed cichą zmianą szablonu przez kogoś z działu marketingu.

**Krok 6 — pomiar.**

Zmierz i zapisz w `output/23_cw3_pomiar.md`:

1. Czas wczytania szablonu z logo vs. bez logo (przygotuj dwa szablony).
2. Rozmiar pliku szablonu z logo vs. bez.
3. Rozmiar pliku wynikowego po wypełnieniu 10 000 wierszy — z logo i bez.

**Kryteria oceny (rubryka):**

| Kryterium | Waga | Na co patrzę |
|---|---|---|
| Szablon nietknięty (test odcisku) | ★★★ | czy odcisk przed i po jest identyczny |
| Logo obecne w obu arkuszach wyniku | ★★★ | czy jest w `Dane` **i** w `Podsumowanie` |
| Zapis atomowy bez `mktemp` | ★★☆ | `mkstemp` + `os.replace` + `try/finally` |
| Tabela Excela z poprawną nazwą | ★★☆ | brak spacji, brak znaków diakrytycznych w `displayName` |
| Walidacja wersji szablonu | ★★☆ | czy nieoczekiwana wersja daje błąd |
| Pomiar udokumentowany | ★☆☆ | liczby, nie „jest szybciej" |
| Raport otwiera się bez ostrzeżeń Excela | ★★★ | **sprawdź to okiem, nie w teście** |

**W pliku `output/23_cw3_wnioski.md` odpowiedz:**

1. **Czy odcisk szablonu był identyczny?** Jeżeli nie — co dokładnie się zmieniło i **dlaczego to jest ważne** (podpowiedź: `[Content_Types].xml`? znaczniki czasu w ZIP? a może naprawdę zapisałeś do szablonu?).
2. **Ile plików jest w `xl/media/` w szablonie, a ile w wyniku?** Wyjaśnij różnicę.
3. **Czy logo jest w `Podsumowanie`?** Jeżeli tak — **jak je tam wstawiłeś** i ile to kosztowało linii kodu? Jeżeli nie — co zrobiłeś źle?
4. **Co się stanie, gdy w szablonie będziesz miał sparkline'y?** Sprawdź to świadomie: zrób (ręcznie w Excelu, jeżeli masz) szablon ze sparkline'ami, użyj go i **zapisz, co zobaczyłeś**. Wklej komunikat openpyxl.
5. **Twoja wersja szablonu — ile razy ją podniesiesz w tym roku?** Jak to kontrolujesz? Napisz procedurę: kto podnosi wersję, gdzie, jak to sprawdzasz przed wypuszczeniem raportu.
6. **Dlaczego `_Konfiguracja` jest `veryHidden`, a nie `hidden`?** W jakim scenariuszu zwykły użytkownik odkryłby ustawienia, których nie powinien widzieć?
7. **Czy twój szablon jest prototypem, czy konfiguracją?** To jest pytanie z pogranicza wzorców: czy wersja szablonu powinna siedzieć **w szablonie** (jak zrobiłeś), czy **w kodzie** (jak stała `WERSJA_SZABLONU`)? Uzasadnij — i wskaż, kiedy która opcja jest lepsza.

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe decyzje i fragmenty</strong></summary>

**Decyzja 1: szablon jest artefaktem, a nie czymś, co się generuje przy każdym uruchomieniu.**

```python
WERSJA_SZABLONU_OCZEKIWANA = "1.0"

def sprawdz_wersje_szablonu(wb, oczekiwana: str) -> None:
    """Zabezpieczenie przed cicha zmiana szablonu przez kogos z zewnatrz."""
    if "_Konfiguracja" not in wb.sheetnames:
        raise ValueError(
            "Szablon nie ma arkusza '_Konfiguracja' - to nie jest nasz szablon."
        )
    ws = wb["_Konfiguracja"]
    znaleziona = None
    for wiersz in ws.iter_rows(values_only=True):
        if wiersz and wiersz[0] == "wersja_szablonu":
            znaleziona = str(wiersz[1])
            break
    if znaleziona != oczekiwana:
        raise ValueError(
            f"Zla wersja szablonu: oczekiwano {oczekiwana!r}, "
            f"znaleziono {znaleziona!r}. Zaktualizuj skrypt albo szablon."
        )
```

**Dlaczego `ValueError`, a nie ciche ostrzeżenie:** bo mówimy o **raporcie**, który trafi do księgowości. Jeżeli szablon zmienił się nieoczekiwanie (ktoś zmienił kolumnę, wyłączył `print_title_rows`, dorzucił wiersz nagłówka), to raport będzie **wyglądał dobrze i był błędny**. Cicha zmiana szablonu to dokładnie ten rodzaj błędu, który „nikt nie zauważy". **Fail fast.**

**Decyzja 2: dwa zabezpieczenia „szablon nie jest celem" — jedno na ścieżce, jedno w środku.**

```python
def raport_z_szablonu(szablon: Path, cel: Path, wiersze, oczekiwana_wersja="1.0") -> Path:
    szablon = Path(szablon).resolve()
    cel_abs = Path(cel).resolve()

    if szablon == cel_abs:
        raise ValueError(
            f"Szablon {szablon} nie moze byc celem zapisu. "
            "To jest regula w kodzie, nie zyczenie w dokumentacji."
        )
    if not szablon.exists():
        raise FileNotFoundError(f"Brak szablonu: {szablon}")

    wb = load_workbook(szablon)          # to jest kopia w PAMIECI
    sprawdz_wersje_szablonu(wb, oczekiwana_wersja)

    ws = wb["Dane"]
    wiersz = 4
    for w in wiersze:
        ws.cell(row=wiersz, column=1, value=w.data).number_format = "dd.mm.yyyy"
        ws.cell(row=wiersz, column=2, value=w.klient)
        ws.cell(row=wiersz, column=3, value=w.region)
        ws.cell(row=wiersz, column=4, value=w.ilosc).number_format = "#,##0"
        ws.cell(row=wiersz, column=5, value=w.cena).number_format = '#,##0.00 "zl"'
        wiersz += 1

    # Nowy arkusz - logo nie ma go automatycznie, bo powstalo w kodzie
    ws2 = wb.create_sheet("Podsumowanie")
    if LOGO.exists():
        wstaw_logo(ws2, LOGO, "A1", wysokosc_px=32)

    ws2["A3"] = "Podsumowanie"
    ws2["A3"].font = Font(bold=True, size=14)
    ws2["A4"] = "Wartosc lacznie"
    ws2["B4"] = sum(w.wartosc for w in wiersze)
    ws2["B4"].number_format = '#,##0.00 "zl"'

    tab = Table(displayName="T_Podsumowanie", ref="A4:B5")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws2.add_table(tab)

    zapisz_atomowo(wb, cel_abs)
    return cel_abs


def zapisz_atomowo(wb, cel: Path) -> None:
    """Zapis atomowy (moduly 18 i 21). NIGDY tempfile.mktemp."""
    cel.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=cel.parent)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, cel)             # atomowe na tym samym wolumenie
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise
```

Zauważ jedną rzecz: **`Table(displayName="T_Podsumowanie", ref="A4:B5")`** — `ref` zawiera nagłówek tabeli (`A4:B4`). To wynika z domyślnego `headerRowCount=1`. Jeżeli podasz `ref="A5:B5"`, Excel pokaże tabelę bez nagłówka, a dane w `A4` zostaną poza tabelą. To najczęstszy błąd przy pierwszej tabeli.

**Decyzja 3: test odcisku szablonu — i jego uczciwe ograniczenia.**

```python
import hashlib
import zipfile

def odcisk_szablonu(sciezka: Path) -> str:
    """Odcisk struktury pliku. NIE 'bajt w bajt' (ZIP ma znaczniki czasu).

    Porownujemy: nazwy czesci + ich rozmiary. To wystarczy, zeby wykryc
    dopisanie danych, zmiane czesci albo utrate elementu.
    """
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for info in sorted(zf.infolist(), key=lambda i: i.filename):
            h.update(info.filename.encode("utf-8"))
            h.update(str(info.file_size).encode("ascii"))
    return h.hexdigest()


def test_szablon_nie_zostal_zmieniony(tmp_path):
    szablon = tmp_path / "szablon.xlsx"
    przygotuj_szablon(szablon)

    przed = odcisk_szablonu(szablon)
    raport_z_szablonu(szablon, tmp_path / "raport.xlsx", dane_testowe())
    po = odcisk_szablonu(szablon)

    assert przed == po, "Szablon zostal zmodyfikowany! Zlamana regula Prototype."
```

**Uczciwe ograniczenie tego testu:** porównuje **nazwy części i rozmiary**, nie zawartość XML. Ktoś mógłby zmienić wartość komórki **nie zmieniając rozmiaru pliku** (np. `123` → `456`) i test by tego nie złapał. **Mocniejsza wersja** porównuje sumy kontrolne poszczególnych części:

```python
def odcisk_szablonu_dokladny(sciezka: Path) -> str:
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for nazwa in sorted(zf.namelist()):
            h.update(nazwa.encode("utf-8"))
            h.update(hashlib.sha256(zf.read(nazwa)).digest())
    return h.hexdigest()
```

**Ale ta wersja ma swój koszt: może się różnić bez żadnej realnej zmiany**, jeżeli Excel zmieni kolejność atrybutów XML albo dorzuci nieznaczącą spację. Dlatego w praktyce używa się obu: **dokładny** do kontroli „nie piszę do szablonu" (bo tam zmiana jest realna), a **strukturalny** do testów regresji raportu. Zapiszesz to jako wniosek.

**Decyzja 4: dlaczego `wstaw_logo` zachowuje proporcje.**

```python
def wstaw_logo(ws, sciezka: Path, komorka: str, wysokosc_px: int = 48) -> None:
    img = XlImage(str(sciezka))
    skala = wysokosc_px / img.height        # PIL podaje wysokosc w pikselach
    img.height = int(round(img.height * skala))
    img.width = int(round(img.width * skala))
    ws.add_image(img, komorka)
```

Bez tego ludzie wpisują `img.width = 200` i logo wychodzi **rozciągnięte** — bo proporcje ustala się raz na zawsze dla danej grafiki, a nie raz na miejsce wstawienia. To ten sam błąd co „dopasuję wykres ręcznie" w module 15: wynik wygląda *prawie* dobrze.

I jedna uwaga, która ratuje w CI: **logo nie może być tylko na Twoim dysku.** `XlImage(str(sciezka))` otwiera plik **przy wstawianiu**, ale openpyxl wczytuje zawartość przy `save`. Jeżeli ścieżka nie istnieje w środowisku CI, raport powstanie **bez logo** i test przejdzie — a testem patrzącym na `xl/media/` **nie przejdzie**. Dlatego w repo trzymaj `assets/logo.png` i **testuj liczbę plików w `xl/media/`**, a nie tylko „czy się zapisało".

**Decyzja 5: pomiar, który warto umieścić w repozytorium jako „baseline".**

| Metryka | Szablon bez logo | Szablon z logo |
|---|---|---|
| Rozmiar szablonu | ~5 kB | ~5 kB + rozmiar PNG |
| Czas `load_workbook(szablon)` | ~5 ms | ~6 ms |
| Plik wynikowy (10 000 wierszy) | ~250 kB | ~250 kB + PNG |

Wniosek z tego pomiaru jest **ważniejszy niż liczby**: **obraz żyje w `xl/media/` i jest kopiowany do pliku wynikowego w całości.** Więc jeżeli ktoś wstawi do szablonu logo 4 MB w formacie BMP, **każdy** wygenerowany raport będzie ważył 4 MB więcej. To jest realny problem w firmach wysyłających raporty mailem. **Wniosek: rozmiar mediów w szablonie podlega przeglądowi** — tak samo jak kolory i progi.

**Czego tu świadomie nie robię:**

- **Nie używam `copy_worksheet` do arkuszy per region.** Wiem, że gubi logo, więc wybieram wariant „wypełnij istniejący arkusz" (`Dane`) i „stwórz nowy w kodzie" (`Podsumowanie`) — a logo wstawiam programowo tylko tam, gdzie go brakuje. **To jest wybór świadomy, nie obejście problemu.**
- **Nie dodaję konfiguracji YAML.** Wersja szablonu w arkuszu `_Konfiguracja` to „konfiguracja w pliku, który jest danymi" — a granica jest tu nieostra i **to jest pytanie z wniosków**. Świadomie zostawiam je otwarte, bo w module 26 zobaczysz pełną odpowiedź.
- **Nie robię z `wstaw_logo` fabryki.** Występuje dwa razy. **Reguła trzech czeka.**

</details>

## 6. Typowe błędy i pułapki

**1. Objaw: „mój `StyleRegistry` jest globalny i testy zaczęły padać w losowej kolejności."**
→ *Przyczyna:* **Singleton ukrywający konfigurację.** Rejestr to nie tylko cache — trzyma też motyw i kolory firmowe. Jako obiekt globalny **wpływa na wszystkie wywołania w procesie**. Test, który zmieni kolor w rejestrze (albo utworzy styl, którego inny test nie spodziewa się znaleźć w cache), zmienia zachowanie kolejnych testów. Ten błąd jest szczególnie paskudny, bo **przechodzi lokalnie i pada w CI**, gdzie kolejność jest inna.
→ *Naprawa:* **wstrzykuj rejestr.** `def pisz(ws, style: StyleRegistry)`. W testach: `StyleRegistry(Motyw(tekst_alertu="000000"))`. Globalnie trzymaj **tylko bezstanowe, niemutowalne definicje** (`Motyw`, `DefinicjaArkusza`) — te są `frozen`, więc są bezpieczne. Diagnostyka: jeżeli `pytest -p no:randomly` przechodzi, a z randomizacją pada — **na 90% masz stan globalny.**

**2. Objaw: „Builder ma 200 linii, ma `build()`, ale nie rozumiem, co mi to daje — kod w środku wygląda jak zwykły generator raportu."**
→ *Przyczyna:* **Builder, który pisze do arkusza w trakcie budowy.** Każde wywołanie `.naglowek(...)` tworzy komórki w już istniejącym `Workbook`. Efekt: (a) `build()` nie ma czego weryfikować, (b) nie da się zbudować w pamięci „na dwa razy", (c) nie da się cofnąć kroku, (d) **nie da się przetestować kompletności przed rozpoczęciem pracy**. Wzorzec istnieje tylko z nazwy.
→ *Naprawa:* **Builder gromadzi PLAN, materializuje w `build()`.** Plan to `list[_PlanArkusza]` — neutralne obiekty bez typów openpyxl. `Workbook()` powstaje **wewnątrz `build()`, po walidacji**. Test na to: `monkeypatch` na `Workbook` + sprawdzenie, że przy błędnej budowie **nie powstał ani jeden obiekt**. Sygnał ostrzegawczy: jeżeli w polach Builder'a jest `self.wb` albo `self.ws` — **masz ten błąd**.

**3. Objaw: „szablon z sparkline'ami i makrami po przejściu przez openpyxl otwiera się, ale czegoś mu brakuje."**
→ *Przyczyna:* **szablon był tworzony ręcznie w Excelu i zawiera rzeczy, których openpyxl nie modeluje.** Sparkline'y (rozszerzenie `x14`) są **gubione** — openpyxl ostrzega przy wczytaniu i nie zapisuje ich z powrotem. Kształty, pola tekstowe, formanty — **gubione**. Makra są zachowane **tylko** z `keep_vba=True`, i tylko zachowane (nieedytowalne). Slicery, Power Query, Power Pivot — poza modelem.
→ *Naprawa:* **testuj szablon, zanim go wdrożysz.** Procedura: wczytaj szablon, zapisz na kopię, otwórz kopię w Excelu i **obejrzyj**. Jeżeli coś zginęło — albo usuń to z szablonu, albo zdobądź szablon „czysty technicznie" (generowany kodem), albo przenieś generowanie na `xlwings`/COM. I dopisz do tego **arkusz `_Konfiguracja` z wersją** — żebyś wiedział, **kiedy** szablon się zmienił. To najczęstsza „cicha” pułapka całego kursu, dlatego moduł 18 mówi o niej trzy razy.

**4. Objaw: „skopiowałem `NamedStyle` z jednego skoroszytu do drugiego i styl się nie nalicza — a po tygodniu kolory w raporcie A zaczęły się zmieniać razem z raportem B."**
→ *Przyczyna:* **`NamedStyle` jest związany ze skoroszytem i jest mutowalny.** `cell.style = "Naglowek"` działa tylko wtedy, gdy nazwa istnieje **w rejestrze tego konkretnego skoroszytu**. Przeniesienie samego obiektu `NamedStyle` przez `add_named_style` w drugim skoroszycie **doda tę samą instancję** — więc zmiana `styl.font` w jednym skoroszocie zmienia ją **w obu**. To nie jest niezależna kopia, a współdzielony obiekt.
→ *Naprawa:* **nie przenoś instancji — przenoś definicję.** Zbuduj `NamedStyle` od nowa w każdym skoroszycie z tych samych `Motyw`/`DefinicjaArkusza`:

```python
def dodaj_styl_nazwany(wb, motyw: Motyw) -> None:
    wb.add_named_style(NamedStyle(
        name="Naglowek",
        font=Font(bold=True, color=motyw.tekst_naglowka),
        fill=PatternFill(fill_type="solid", fgColor=motyw.tlo_naglowka),
    ))
```

Diagnostyka: `grep -rn "add_named_style" src/` — jeżeli w dwóch miejscach **ten sam obiekt** jest dodawany do dwóch skoroszytów, masz ten błąd. I uwaga dodatkowa: **w openpyxl 3.1.x API nazwanych stylów zostało przebudowane** — jeżeli używasz starszej wersji, zachowanie może się różnić. Sprawdź w dokumentacji swojej wersji.

**5. Objaw: „owinąłem metodę w `@lru_cache` dla wydajności, a aplikacja zjada pamięć i nie zwalnia skoroszytów."**
→ *Przyczyna:* **`lru_cache` na metodzie trzyma `self`.** Cache zbudowany na metodzie instancji pamięta `self` jako klucz — więc **każdy skoroszyt pozostaje w pamięci na zawsze**, dopóki cache żyje. To najczęstszy wyciek pamięci w Pythonie i jest niewidoczny, dopóki nie zobaczysz `MemoryError` po trzech godzinach pracy serwera.
→ *Naprawa:* **nie owijaj metod.** Owijaj **funkcje modułowe**, które przyjmują **proste, hashowalne wartości**:

```python
@lru_cache(maxsize=512)
def _font(bold: bool, size: int, kolor: str) -> Font:
    return Font(bold=bold, size=size, color=kolor)
```

Alternatywnie — tak jak w tym module — **własny cache w instancji** (`self._fonts`). Zaleta własnego cache'u: znika razem z obiektem. Zaleta `lru_cache`: ma `maxsize`, więc nie rośnie bez ograniczeń. **W rejestrze stylów ogranicz rozmiar sam**, bo kombinacji jest skończenie wiele (`len(KOLORY) * len(ROZMIARY)` = 35) i nie ma ryzyka. W innych przypadkach `lru_cache(maxsize=...)` jest właściwy.

**6. Objaw: „wygenerowałem 30 arkuszy (po jednym na region), a dane są w 12 — i plik jest wolny, a połowa arkuszy jest pusta."**
→ *Przyczyna:* **tworzenie wszystkiego „na wszelki wypadek".** Arkusze powstają w pętli po liście **potencjalnych** regionów, a nie po regionach, które **faktycznie mają dane**. To jest brak leniwego tworzenia (sekcja 2.6).
→ *Naprawa:* **klucz w leniwej mapie, nie w pętli.** Wzorzec `LeniweArkusze` z przykładu 5: `arkusz(nazwa)` tworzy dopiero przy pierwszym żądaniu. I **uczciwe zastrzeżenie**: leniwość arkuszy pomaga naprawdę tylko wtedy, gdy część arkuszy zostaje pusta. Jeżeli wszystkie 30 arkuszy ma po 100 000 wierszy, **leniwość nic nie da** — dominuje koszt komórek (moduł 19). Diagnostyka: `print(len(wb.sheetnames), "arkuszy z 30 mozliwych")`.

**7. Objaw: „napisałem `RaportSpec` w YAML, ale żeby dodać warunek `i co najwyżej 3 regiony`, musiałem wpisać fragment Pythona do YAML-a."**
→ *Przyczyna:* **konfiguracja deklaratywna, która urosła do języka programowania.** To ta sama pułapka, którą zapowiedział moduł 22 (sekcja 2.7), tylko w nowym przebraniu. Objaw: w konfiguracji pojawiają się operatory, nawiasy, `and`, `or`, a wkrótce „funkcje pomocnicze" w YAML-u.
→ *Naprawa:* **konfiguracja opisuje CO, kod opisuje JAK.** Format, nagłówek, szerokość, kolor, tytuł — konfiguracja. Operatory, warunki, pętle, agregacje — **nazwane funkcje w świecie danych** (moduł 22). Test: **jeżeli konfiguracja wymaga debuggera, to nie jest konfiguracja.** I jeszcze jedno kryterium, które warto zapamiętać: jeżeli musisz napisać w konfiguracji komentarz wyjaśniający, **w jakiej kolejności** są stosowane reguły — to znaczy, że konfiguracja stała się programem.

**8. Objaw: „dodałem pole `akcent: bool` do `Motyw` i nagle w `StylFabryka` mam `if motyw.akcent: ...` w pięciu miejscach."**
→ *Przyczyna:* **flaga w konfiguracji to zakamuflowany `if` w kodzie.** Przenieśliśmy decyzję do danych, ale rozgałęzienie zostało w logice (pułapka 6 z modułu 22 w przebraniu wzorca konstrukcyjnego). Motyw mówi „używam akcentu", a fabryka musi wiedzieć, **co to znaczy i gdzie**.
→ *Naprawa:* **flaga → osobny obiekt albo osobny zestaw gotowych „pieczątek".** Zamiast `if motyw.akcent` zrób `MotywA` i `MotywB` albo dodaj do `Motyw` gotowe kombinacje:

```python
class StylFabryka:
    def fill_paska(self) -> PatternFill:
        return self.fill(self.motyw.tlo_paska)     # wartosc, nie flaga
```

Jeżeli musisz mieć „dwa warianty wyglądu", to są **dwie instancje `Motyw`**, a nie jedna z flagą. Kryterium: `grep -n "motyw\." examples/` — jeżeli którakolwiek linia ma `if`, to jest podejrzana.

**9. Objaw: „napisałem `@dataclass(frozen=True)` na `RaportSpec`, ale `spec.kolumny.append(...)` działa i psuje wszystkie raporty."**
→ *Przyczyna:* **`frozen=True` blokuje przypisanie, nie mutację zawartości.** `spec.kolumny = x` rzuci `FrozenInstanceError`; `spec.kolumny.append(x)` — **nie rzuci**, jeżeli `kolumny` to `list`. Powtórka pułapki 9 z modułu 22, ale tu dotyczy konkretnie specyfikacji i fabryk — więc konsekwencje są większe (współdzielony obiekt, jeden raport psuje drugi).
→ *Naprawa:* **typy niemutowalne w polach.** `tuple[Kolumna, ...]` zamiast `list`. `Mapping` zamiast `dict`. `frozenset` zamiast `set`. I `frozen=True` **na całym łańcuchu** (`Kolumnа`, `Motyw`, `DefinicjaArkusza`, `RaportSpec`), bo `tuple` chroni tylko jeden poziom. **Test:** napisz test, który próbuje zmodyfikować specyfikację i oczekuje wyjątku — to zamienia konwencję w kontrakt.

**10. Objaw: „mój `create_sheet` w fabryce wywala się na drugim raporcie: `ValueError: Worksheet ... already exists`."**
→ *Przyczyna:* **nazwy arkuszy są unikalne w skoroszycie i fabryka o tym nie wie.** Jeżeli `DEFINICJE["dane"].nazwa == "Dane"` i wołasz `utworz(wb, "dane")` dwa razy na tym samym skoroszycie, openpyxl zgłosi konflikt. Podobnie jeżeli ktoś ręcznie utworzył arkusz „Dane" przed fabryką.
→ *Naprawa:* **fabryka powinna wykrywać konflikt i tłumaczyć go zrozumiale.** Sprawdź `wb.sheetnames` (moduł 08) i rzuć wyjątek z komunikatem, który mówi, **co już jest**:

```python
if def_.nazwa in wb.sheetnames:
    raise ValueError(
        f"Arkusz {def_.nazwa!r} juz istnieje w skoroszycie. "
        f"Obecne arkusze: {wb.sheetnames}"
    )
```

Alternatywa: fabryka przyjmuje parametr `nadpisz=False`. **Ale nie rób z tego automatycznego unikania kolizji** („dopiszę cyfrę") — bo wtedy dostaniesz `Dane`, `Dane2`, `Dane3` i nikt nie będzie wiedział, który jest prawdziwy. **Konflikt nazw to błąd w logice, nie w nazewnictwie.**

**11. Objaw: „schowałem arkusz `_Konfiguracja` przez `veryHidden`, a i tak użytkownik go znalazł i zmienił wartości."**
→ *Przyczyna:* **`veryHidden` to zaciemnienie, nie zabezpieczenie.** Arkusz „very hidden" nie pojawia się w oknie „Unhide", ale **jest w pliku** i każdy, kto użyje Pythona albo edytora OOXML, go odczyta. A jeżeli jest w kopii `.xlsx` wysłanej mailem — **jest tam**.
→ *Naprawa:* **nie trzymaj w skoroszycie niczego, czego nie chcesz wysyłać.** Konfiguracja to dane w kodzie, nie ukryty arkusz w pliku klienta. Jeżeli musisz mieć arkusz konfiguracyjny do wewnętrznego użytku — trzymaj go w **szablonie** (który nie wychodzi na zewnątrz), a raport wynikowy generuj **bez niego** albo z pustą zawartością. Diagnostyka przed wysyłką: `zipfile.ZipFile(plik).namelist()` + sprawdzenie `sheetnames`, czy nie ma arkuszy, których nie powinno być. To ten sam audyt, o którym mówi moduł 21.

**12. Objaw: „używam `wb.remove(wb["Dane"])`, żeby wyczyścić skoroszyt szablonu, i Excel zgłasza `#REF!` w formułach."**
→ *Przyczyna:* **usunięcie arkusza nie usuwa odwołań.** Definiowane nazwy, formuły w innych arkuszach, tabele i reguły formatowania warunkowego mogą wskazywać na usunięty arkusz. To analogia z modułu 08: „usunięcie strony z indeksem" — odwołania zostają i prowadzą donikąd. openpyxl **nie ostrzega**.
→ *Naprawa:* **usuń odwołania przed usunięciem arkusza** albo **nie usuwaj go wcale**. W praktyce: jeżeli szablon ma arkusz, którego nie potrzebujesz w wyniku, **nie usuwaj go w kodzie — wygeneruj wynik na bazie innego szablonu** albo wyczyść zawartość (i pozostaw arkusz pusty, ale istniejący). To kolejny argument za tym, żeby szablon miał **arkusze, które naprawdę są potrzebne**. Sprawdź po usunięciu, że `wb.defined_names` nie zawiera odwołań do usuniętej nazwy — w 3.1.x API definiowanych nazw zostało przebudowane (`DefinedName`, `wb.defined_names.add(...)`), więc sprawdź w dokumentacji swojej wersji.

**13. Objaw: „Builder był świetny w testach, ale w produkcji generuje pliki 400 MB i pada na pamięci."**
→ *Przyczyna:* **Builder gromadzi CAŁY plan w pamięci.** Przy 50 000 wierszy plan trzyma **kopię tuple'i z danymi** — czyli te same dane w drugiej strukturze. Jeżeli dane wejściowe też są w pamięci (lista `WierszRaportu`), masz **podwojenie**. Do tego `build()` tworzy pełny model openpyxl w RAM przed zapisem.
→ *Naprawa:* **Builder dla struktury, strumień dla masy.** Builder planuje **arkusze i nagłówki** (to jest małe i wymaga walidacji), a **wiersze zapisuj strumieniowo** — najlepiej w trybie `write_only` (moduł 19). Praktyczny podział: Builder buduje `Workbook(write_only=True)`, a `.wiersze()` **nie gromadzi** danych, tylko zapisuje je od razu przez `append`. Ale wtedy **walidacja kompletności musi zadziałać przed pierwszym `append`** — czyli plan arkuszy musi być gotowy, zanim ruszy jakikolwiek wiersz. To jest dokładnie ten przypadek, w którym Builder ma sens **dlatego**, że `write_only` jest nieodwracalny. Zmierz: `tracemalloc` przed i po.

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Wzorce konstrukcyjne to cztery elementy zakładu produkcyjnego: foremka, instrukcja, wzorzec i magazyn.** Foremka (**Factory**) — raz opisujesz kształt, używasz tysiące razy, zmieniasz w jednym miejscu. Instrukcja (**Builder**) — kroki w kolejności i kontrola kompletności na końcu. Wzorzec (**Prototype**) — pieczątka, która **nigdy nie jest nadpisywana**. Magazyn (**Registry/Flyweight**) — matryca leży na półce, nie rzeźbisz czcionki dla każdego wyrazu. I wszystko to po to, żeby rozwiązać trzy koszty: rozproszoną wiedzę, zmarnowany czas i obiekt w złym stanie.

2. **Builder musi gromadzić PLAN i materializować go w `build()` — inaczej traci cały sens.** Jeżeli w Builderze jest `self.wb`, to nie jest Builder, a generator raportu z łańcuchowanym API. Poprawny Builder pozwala: **zwalidować kompletność przed rozpoczęciem pracy**, **zbudować skoroszyt w pamięci bez dotykania dysku** (moduł 20) i **cofnąć się bez sprzątania po sobie**. `build()` zwraca `Workbook`, `zapisz()` = `build()` + `save()`. **Ta granica to granica między testowalnym a nietestowalnym.**

3. **`copy_worksheet` gubi obrazy i wykresy — i to jest udokumentowane, nie błąd.** Szablon jako Prototype działa dlatego, że wypełniasz **istniejący** arkusz, a nie kopię. Jeżeli musisz kopiować, rekompensujesz: logo wstawiasz programowo po kopii. I pamiętaj, że szablon zrobiony **ręcznie w Excelu** może zawierać sparkline'y, kształty, slicery i Power Query — **openpyxl je gubi przy zapisie**. Szablon trzeba **przetestować okiem**, a nie „uruchomić i założyć, że jest dobrze".

4. **Registry to cache, ale też konfiguracja — więc nie rób z niego Singletona.** Jako obiekt globalny: przecieka między testami, uniemożliwia dwa motywy naraz i ukrywa zależność w sygnaturze funkcji. **Wstrzykuj rejestr.** Globalnie trzymaj tylko to, co jest naprawdę bezstanowe i `frozen` (`Motyw`, `DefinicjaArkusza`) — i pamiętaj, że w Pythonie **moduł już jest singletonem**, więc nie potrzebujesz do tego klasy z `__new__`. A rozmiar pliku **się nie zmieni** — bo openpyxl i tak deduplikuje style w `styles.xml`. Rejestr kupuje **czas Pythona i nazwy**, nie kilobajty.

5. **Konfiguracja i DTO domykają kreację: obiekt powstaje z DANYCH, nie z kodu.** `RaportSpec` opisuje, *co* ma powstać; `WierszRaportu` opisuje, *jakie dane* tam trafią, i **waliduje się w konstruktorze** — więc błędne dane nie wchodzą do systemu, a literówka w nazwie pola wybucha natychmiast, zamiast stać się pustą kolumną w Excelu. Granica jest twarda i łatwa do zapamiętania: **konfiguracja opisuje CO, kod opisuje JAK.** Jeśli konfiguracja wymaga debuggera, to już nie jest konfiguracja. A jeżeli specyfikacja obiecuje dwa arkusze, a generator tworzy jeden — **konfiguracja nie zastępuje walidacji**.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. WYBOR WZORCA - tabela decyzyjna w wersji minimalnej
# ==================================================================
# 5-7 linii konfiguracji arkusza w 3+ miejscach      -> FACTORY
# kolory/progi bez nazw                              -> KONFIGURACJA (najtansze!)
# czcionka/wypelnienie konstruowane w 40 miejscach   -> STYLFABRYKA + REGISTRY
# 6 krokow, kolejnosc i kompletność maja znaczenie   -> BUILDER z _waliduj()
# bogaty szablon z Excela                            -> PROTOTYPE (load + save as NEW)
# "wzor dla kazdego regionu"                         -> PROTOTYPE + reuzycie/kopia
# ten sam Font w tysiacach komorek                   -> FLYWEIGHT (Registry)
# globalny rejestr z konfiguracja                    -> NIE. Wstrzyknij rejestr.
# 30 potencjalnych arkuszy, czesc pusta              -> LAZY
# dict/list miedzy warstwami                         -> DTO (dataclass + __post_init__)
# nowy raport = elif w generatorze                   -> KONFIGURACJA (RaportSpec)


# ==================================================================
# 2. FACTORY - jedna fabryka, definicje jako dane
# ==================================================================
@dataclass(frozen=True)
class DefinicjaArkusza:
    nazwa: str                       # nazwa widoczna w Excelu
    tab_color: str | None = None
    freeze: str | None = None
    show_grid: bool = True
    szerokosci: tuple[float, ...] = ()
    stan: str = "visible"            # visible | hidden | veryHidden

DEFINICJE = {
    "dane":        DefinicjaArkusza("Dane", "1072BA", "A4", False,
                                    (12.0, 24.0, 8.0, 16.0, 8.0, 10.0)),
    "podsumowanie": DefinicjaArkusza("Podsumowanie", "70AD47", None, False, (24.0, 16.0)),
}

class ArkuszFabryka:
    def utworz(self, wb, klucz, index=None):
        def_ = self._definicje[klucz]                 # KeyError -> komunikat z lista
        ws = wb.create_sheet(title=def_.nazwa, index=index)
        if def_.tab_color:  ws.sheet_properties.tabColor = def_.tab_color
        if def_.freeze:     ws.freeze_panes = def_.freeze
        if not def_.show_grid: ws.sheet_view.showGridLines = False
        for i, szer in enumerate(def_.szerokosci):
            ws.column_dimensions[get_column_letter(i + 1)].width = szer
        if def_.stan != "visible": ws.sheet_state = def_.stan
        return ws


# ==================================================================
# 3. STYLE REGISTRY / FLYWEIGHT - cache + nazwy
# ==================================================================
class StyleRegistry:
    """WSTRZYKIWANY, nie globalny. Cache w instancji = znika z obiektem."""
    def __init__(self, motyw=None):
        self.motyw = motyw or Motyw()
        self._fonts: dict[tuple, Font] = {}
        self.utworzone = 0                       # metryka: dowod, ze cache dziala

    def font(self, **cechy) -> Font:
        klucz = tuple(sorted(cechy.items()))     # wartosci MUSZA byc hashable
        gotowy = self._fonts.get(klucz)
        if gotowy is None:
            gotowy = Font(**cechy)
            self._fonts[klucz] = gotowy
            self.utworzone += 1
        return gotowy

    # gotowe "pieczatki" z motywu
    def font_naglowka(self): return self.font(bold=True, color=self.motyw.tekst_naglowka)
    def font_tytulu(self):   return self.font(bold=True, size=self.motyw.rozmiar_tytulu)
    def font_alertu(self):   return self.font(color=self.motyw.tekst_alertu)

# NIE TAK (Singleton ukrywajacy konfiguracje):
#     STYLE = StyleRegistry()          # globalny stan -> testy przeciekaja
# TAK (wstrzykiwanie):
#     def pisz(ws, style: StyleRegistry) -> None: ...


# ==================================================================
# 4. BUILDER - plan -> walidacja -> materializacja
# ==================================================================
@dataclass
class _PlanArkusza:                      # neutralny: ZERO typow openpyxl
    nazwa: str
    naglowek: tuple[str, ...] | None = None
    wiersze: list[tuple] = field(default_factory=list)
    kpi: list[tuple] = field(default_factory=list)
    tabela: list[tuple] | None = None
    szerokosci: tuple[float, ...] | None = None

    def czy_pusty(self) -> bool:
        return not self.wiersze and not self.kpi and self.tabela is None

class RaportBuilder:
    def __init__(self, tytul, style=None):
        self._tytul = tytul
        self._style = style
        self._arkusze: list[_PlanArkusza] = []   # ← PLAN, nie skoroszyt!
        self._biezacy = None

    # API plynne
    def dodaj_arkusz(self, nazwa, szerokosci=None): ...
    def naglowek(self, etykiety): ...
    def wiersze(self, dane): ...
    def kpi(self, etykieta, wartosc, fmt=None): ...
    def tabela(self, dane): ...

    def build(self) -> Workbook:
        self._waliduj()                      # ← PRZED jakakolwiek praca
        wb = Workbook(); del wb["Sheet"]
        for i, plan in enumerate(self._arkusze):
            self._wypelnij(wb.create_sheet(title=plan.nazwa, index=i), plan)
        return wb                            # ← zwraca Workbook, NIE plik

    def zapisz(self, cel) -> Path:
        cel = Path(cel); cel.parent.mkdir(parents=True, exist_ok=True)
        self.build().save(cel); return cel

    def _waliduj(self) -> None:
        problemy = []
        # zduplikowane nazwy | >31 znakow | niedozwolone znaki
        # arkusze puste | naglowek bez danych (i odwrotnie)
        # wiersze o innej liczbie kolumn niz naglowek
        if problemy:
            raise BladBudowy("Raport niekompletny:\n  - " + "\n  - ".join(problemy))

# TEST, ze Builder NAPRAWDE nie tworzy Workbooka przed walidacja:
#     monkeypatch.setattr("mojmodul.Workbook", podszywacz)   # patchuj w MOIM module!
#     with pytest.raises(BladBudowy): builder.build()
#     assert wywolania == []


# ==================================================================
# 5. PROTOTYPE - szablon
# ==================================================================
def raport_z_szablonu(szablon: Path, cel: Path, dane):
    szablon, cel = Path(szablon).resolve(), Path(cel).resolve()
    if szablon == cel:
        raise ValueError("Szablon nie moze byc celem zapisu!")   # regula W KODZIE

    wb = load_workbook(szablon)              # kopia w PAMIECI
    ws = wb["Dane"]                          # uzywamy ISTNIEJACEGO arkusza
    # ... dopisujesz TYLKO to, co sie zmienia ...
    cel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(cel)                             # NOWY plik
    return cel

# .xlsm:  load_workbook(p, keep_vba=True)  + save z rozszerzeniem .xlsm
# .xltx:  dziala, ale zapis tworzy zwykly skoroszyt - trzymaj szablon jako .xlsx

# Test odcisku (dowod, ze szablon jest nietkniety):
def odcisk_szablonu(sciezka: Path) -> str:
    h = hashlib.sha256()
    with zipfile.ZipFile(sciezka) as zf:
        for info in sorted(zf.infolist(), key=lambda i: i.filename):
            h.update(info.filename.encode()); h.update(str(info.file_size).encode())
    return h.hexdigest()


# ==================================================================
# 6. LOGO - skala z zachowaniem proporcji (jednostki: PIKSELE!)
# ==================================================================
def wstaw_logo(ws, sciezka: Path, komorka: str, wysokosc_px: int = 48) -> None:
    img = XlImage(str(sciezka))          # wymaga pillow
    skala = wysokosc_px / img.height
    img.height = int(round(img.height * skala))
    img.width = int(round(img.width * skala))
    ws.add_image(img, komorka)

# DIAGNOSTYKA (prywatne atrybuty - tylko jednorazowo!):
#     print(len(ws._images))           # obraz NIE idzie przez copy_worksheet
#     print(len(wb._fonts))            # ile stylow NA PRAWDE widzi openpyxl


# ==================================================================
# 7. DTO + KONFIGURACJA
# ==================================================================
@dataclass(frozen=True)
class WierszRaportu:
    data: date; klient: str; region: str; ilosc: int; cena: float

    def __post_init__(self) -> None:      # walidacja w konstruktorze
        if self.ilosc < 0: raise ValueError(f"ilosc < 0: {self.ilosc!r}")
        if not self.klient.strip(): raise ValueError("klient pusty")

    @property
    def wartosc(self) -> float:           # getattr(obj, "wartosc") to WYWOLA
        return self.ilosc * self.cena

@dataclass(frozen=True)
class RaportSpec:
    nazwa: str; tytul: str
    kolumny: tuple[Kolumna, ...]          # TUPLE, nie list (frozen chroni tylko przypisanie)
    tytul_arkusza: str = "Dane"

# GRANICA:  konfiguracja = CO  |  kod = JAK
#   DOBRZE:  Kolumna("Cena", 10, "cena", '#,##0.00 "zl"')
#   ZLE:     filtr="ilosc > 10 and region in ['PL','DE']"
# TEST: jezeli konfiguracja wymaga debuggera, to nie jest konfiguracja.


# ==================================================================
# 8. LENIWE ARKUSZE
# ==================================================================
class LeniweArkusze:
    def __init__(self, wb: Workbook) -> None:
        self._wb = wb; self._gotowe: dict[str, object] = {}

    def arkusz(self, nazwa, szerokosci=()):
        ws = self._gotowe.get(nazwa)
        if ws is None:
            ws = self._wb.create_sheet(title=nazwa)
            for i, szer in enumerate(szerokosci, start=1):
                ws.column_dimensions[get_column_letter(i + 1)].width = szer
            self._gotowe[nazwa] = ws
        return ws

# UCZCIWIE: leniwosc arkuszy pomaga TYLKO gdy czesc zostaje pusta.
# Gdy wszystkie 30 arkuszy ma po 100k wierszy - dominuje koszt KOMOREK.


# ==================================================================
# 9. PULAPKI DO ZAPAMIETANIA
# ==================================================================
# - Builder z polem self.wb/self.ws = Builder, ktory niczego nie pilnuje
# - copy_worksheet NIE kopiuje obrazow ani wykresow (udokumentowane)
# - Szablon z Excela: sparkline'y, shapes, slicery, PowerQuery -> GUBIONE
# - lru_cache na METODZIE trzyma self -> wyciek pamieci (owijaj funkcje!)
# - NamedStyle jest zwiazany ze skoroszytem i MUTOWALNY (nie przenos instancji)
# - frozen=True chroni przypisanie, NIE mutacje zawartosci (tuple/frozenset!)
# - klucz cache z **cechy wymaga wartosci HASHOWALNYCH
# - flaga w konfiguracji (motyw.akcent) = zakamuflowany if w kodzie
# - veryHidden to zaciemnienie, nie zabezpieczenie
# - usuniecie arkusza nie usuwa odwolan -> #REF! w formulach
# - create_sheet z istniejaca nazwa -> ValueError (nie "unikaqj kolizji"!)
# - Table(displayName=...) bez spacji; ref ZAWIERA wiersz naglowka
# - rozmiar pliku sie NIE zmieni po dodaniu Registry (openpyxl dedupuje style)
```

## 9. Co dalej

Zatrzymaj się i policz, co masz po tym module. **Umiesz już tworzyć obiekty w sposób kontrolowany i tani** — i to jest cała treść, ale wartość jest większa, niż się wydaje.

Masz **fabrykę**, która zamieniła dwadzieścia rozproszonych miejsc w jedną definicję. Masz **Builder**, który gromadzi plan i **weryfikuje kompletność przed rozpoczęciem pracy** — a przy okazji daje Ci `build()` zwracający `Workbook`, czyli testowalność z modułu 20 gratis. Masz **Prototype** z twardą regułą w kodzie („szablon nie jest celem zapisu") i z **testem odcisku**, który tę regułę udowadnia. Masz **rejestr stylów**, który jest wstrzykiwany, a nie globalny — i umiesz zmierzyć, czy naprawdę działa (`licznik_alokacji`). I masz **konfigurację oraz DTO**, które domykają kreację: obiekt powstaje z danych, a nie z kodu.

Zauważ jedną rzecz, bo jest ważna. **W tym module dwa razy wróciłeś do modułu 18** — raz przy szablonie (sparkline'y, makra, `keep_vba`), raz przy regule „szablon nigdy nie jest nadpisywany" (atomowy zapis, `mkstemp` + `os.replace`). To nie przypadek. **Wzorce konstrukcyjne i bezpieczeństwo plików to ten sam temat opowiedziany z dwóch stron**: jeden mówi „jak powstaje", drugi „co się stanie, gdy powstanie źle".

I jeszcze jedno, co warto zapamiętać z przykładu 4, bo jest to lekcja szersza niż openpyxl: **rozmiar pliku wynikowego był identyczny**. Rejestr nie kupił kilobajtów. Kupił **czas Pythona i nazwy**. Jeżeli kiedyś zobaczysz optymalizację, która „wygląda dobrze", ale nie umiesz powiedzieć, **co konkretnie kupuje** — zmierz. W tym module metryki są wbudowane w kod (`licznik_alokacji`, `ile_w_cache`, `utworzone`) i to nie jest ozdoba.

W module 24 wchodzimy w **wzorce strukturalne** i robimy to na **dokładnie tym samym przykładzie** — raporcie sprzedaży, który ewoluuje od 60-linijkowego skryptu w moduł 22, przez fabrykę i Builder tutaj, aż do trójwarstwowej aplikacji w module 26. Zobaczysz:

- **Facade (`RaportExcel`)** — jedno wejście ukrywające `Workbook`, tabele, wykresy i zapis. Zobaczysz **test jakości fasady**: jeżeli w jej sygnaturach są typy openpyxl (`ws`, `cell`, `Font`), to nie jest fasada, a „przepakowanie".
- **Adapter** — mapowanie modelu domenowego na wiersz arkusza i rozwiązanie **kompromisu `getattr` z tego modułu** (generyk albo protokół), który świadomie odłożyliśmy.
- **Decorator** — `@waliduj_dane`, `@loguj_zapis`, `@mierz_czas` wokół zapisu, bez dotykania zapisu.
- **Proxy** — `LazyWorkbook`, który ładuje plik dopiero przy pierwszym dostępie, i proxy blokujące zapis bez zatwierdzenia. **Tu dokończymy wątek leniwości**, który w tym module zacząłeś.
- **Composite** — sekcje raportu, gdzie rozdział może zawierać podrozdziały. Zobaczysz, dlaczego Composite **nie może** budować się przez pisanie do arkusza (ta sama pułapka co Builder z tego modułu, tylko ostrzejsza).

I zobaczysz tam coś, czego ten moduł nie mógł pokazać: **co się dzieje, gdy fasada zaczyna przeciekać**. Bo najczęstszy błąd przy fasadzie to zwrócenie z niej `ws` „na chwilę, bo akurat potrzebuję". Wtedy cała warstwa przestaje działać, a Ty tego nie zauważasz przez trzy miesiące.

Ale zanim tam pójdziemy, zrób jedno. Weź swój **największy skrypt generujący raport** i policz w nim dwie rzeczy:

1. **Ile razy występuje `create_sheet`?** Jeżeli więcej niż dwa razy — masz kandydata na fabrykę (i to jest Twoje ćwiczenie 🟢, tylko na realnym kodzie).
2. **Ile razy tworzysz obiekt stylu wewnątrz pętli?** `grep -n "Font(\|PatternFill(\|Alignment(" skrypt.py`. Jeżeli w pętli — masz kandydata na rejestr, a przykład 4 pokaże Ci, ile to kosztuje.

I zastosuj jedną regułę z tego modułu, **nawet jeżeli nic nie refaktoryzujesz**: jeżeli tylko jedna rzecz z modułu 23 ma zostać z Tobą na stałe, niech to będzie ta:

> **Jeżeli Builder ma pole `self.wb` albo `self.ws` — to nie jest Builder.** Ma pole `self._arkusze` z planem. Bo dopiero wtedy `build()` ma co sprawdzić.