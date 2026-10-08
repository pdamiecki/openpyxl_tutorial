# Moduł 25 — Wzorce behawioralne: elastyczne zachowanie bez setek `if`-ów

> **Część:** IV — Wzorce projektowe i architektura · **Poziom:** ⭐⭐⭐ · **Wymaga:** modułów 00–24

## 0. W tym module nauczysz się

- **Zamienisz łańcuchy `if/elif` na wymienne strategie** i zrozumiesz rzecz, którą większość książek pomija: **Strategy nie usuwa decyzji — przenosi ją w jedno miejsce.** Zobaczysz, jak wybór z 5 typów raportów × 3 wariantów eksportu redukuje się z 15 rozgałęzień do `5 + 3` klas i **jednego** słownika.
- **Zbudujesz `BazowyRaport` metodą szablonową** — szkielet `generuj()` (przygotuj → nagłówek → dane → podsumowanie → formatowanie → wykresy → zapis) plus haczyki `_naglowek()`, `_wiersze()`, `_podsumowanie()`. Nauczysz się odróżnić **haczyk wymagany** (abstrakcyjny) od **haczyka opcjonalnego** (domyślnie nic nie robi) i zobaczysz, dlaczego to rozróżnienie jest jedynym zabezpieczeniem przed pułapką „raport bez podsumowania".
- **Zbudujesz komendy z cofaniem** — `KomendaUstawKomorke`, `KomendaDodajWiersz`, `KomendaSformatuj` oraz `StosKomend` z `cofnij()`, `ponow()`, makrem i **logiem audytowym**. Zrozumiesz, dlaczego to niemutowalność stylów z modułu 09 sprawia, że cofanie formatowania jest trywialne, a cofanie `insert_rows` — niebezpieczne.
- **Zbudujesz cztery gości (Visitor)** — `SzukajPustychKolumn`, `WykryjFormuly`, `ZbierzMetryki`, `Anonimizuj`. Zrozumiesz **podwójne wywołanie zwrotne** (double dispatch) i — najważniejsze — dlaczego gość, który **modyfikuje to, co odwiedza**, psuje iterację i gubi komórki. Rozwiązanie: **dwie fazy** — najpierw zbieraj, potem zastosuj, najlepiej jako **komendy** z poprzedniego punktu.
- **Zbudujesz specyfikacje deklaratywne** z kompozycją `&`, `|`, `~` i zobaczysz **najpiękniejszy most tego kursu**: te same obiekty reguł **filtrują dane przed zapisem** i **generują reguły formatowania warunkowego** z modułów 13–14. Nauczysz się też uczciwie mówić, **które reguły da się przetłumaczyć na formułę Excela, a których nie**.
- **Zobaczysz Observer i Chain of Responsibility w wersji, która ma sens** — i uczciwą uwagę, kiedy są nadmiarem. Dowiesz się, dlaczego `logging` wystarcza w małym projekcie, a kiedy autobus zdarzeń zaczyna zarabiać na siebie.
- **Otrzymasz tabelę decyzyjną** „problem → wzorzec → mini-przykład" oraz **12 pułapek** tej warstwy, w tym cztery wskazane w zadaniu: nadmiarowe `if/elif`, Template Method z pustymi haczykami, Command modyfikujący skoroszyt przed `wykonaj()` i Specification, która robi wejście/wyjście w trakcie ewaluacji.

## 1. Intuicja i analogia

### 1.1. Wymienne końcówki do wkrętarki, czyli Strategy

Masz **jedną wkrętarkę**. Do niej **pięć końcówek**: krzyżak, płaski, torx, imbus, nasadka. Kiedy zmieniasz śrubę, **nie kupujesz nowej wkrętarki** — wymieniasz końcówkę. Wkrętarka (mechanizm, silnik, rączka) jest jedna. Końcówki różnią się **wyłącznie sposobem pracy na tym samym materiale**.

Zwróć uwagę na trzy rzeczy z tej analogii, bo wszystkie trzy dotyczą kodu:

1. **Końcówki mają identyczne mocowanie.** Nie „podobne" — **identyczne**. To nie przypadek, to **kontrakt**. Bez wspólnego mocowania końcówki byłyby bezużyteczne. W kodzie: wspólna metoda `zastosuj()`.
2. **Wkrętarka nie wie, którą końcówkę ma.** Wkrętarka kręci. **Ty** decydujesz, którą końcówkę założysz. W kodzie: kontekst nie zna strategii, strategia nie zna kontekstu.
3. **Istnieje moment wyboru.** Sięgnięcie po końcówkę do skrzynki. To jest miejsce, w którym **musi** być jakaś decyzja — i teraz najważniejsze:

> **Strategy nie usuwa `if`-a. On przenosi go w jedno miejsce.**

Jeżeli masz pięć typów raportu i trzy warianty eksportu, to w wersji naiwnej masz piętnaście rozgałęzień **rozsypanych po całym kodzie** — w generatorze, w formacie, w zapisie. W wersji ze strategiami masz **jedno** miejsce, w którym wybierasz końcówkę:

```python
strategia = STRATEGIE_FORMATOWANIA[specyfikacja.format]     # ← JEDEN wybor
```

I to nie jest jeden `if`, tylko **jeden słownik** — a więc nawet nie `if`. Ta sama liczba decyzji (wciąż musisz wiedzieć, że „raport zarządczy" idzie gęsto, a „audytowy" finansowo), ale zapisana **raz**, w widocznym miejscu, a nie piętnaście razy w środku pętli.

### 1.2. Formularz urzędowy, czyli Template Method

Wchodzisz do urzędu i dostajesz **formularz**. Formularz ma **nadrukowany szkielet**: dane osobowe, adres, dochód, podpis, data. Urząd **narzucił kolejność i obecność** pól. Ale **treść pól wypełniasz Ty**. Urzędnik nie mówi Ci „wypełnij pole dochodu, jeśli ci pasuje" — pole jest tam zawsze. Natomiast jest rubryka „uwagi", która **może zostać pusta** i wtedy formularz jest nadal poprawny.

To jest cała różnica, o której mówi ten wzorzec:

- **Szkielet (`generuj()`)** — kolejność kroków. Nadrukowany. **Podklasa go nie zmienia.**
- **Haczyki wymagane (`_naglowek()`, `_wiersze()`)** — pola obowiązkowe. Podklasa **musi** je wypełnić, a jeżeli nie wypełni, Python **nie pozwoli utworzyć obiektu** (`@abstractmethod`).
- **Haczyki opcjonalne (`_podsumowanie()`, `_formatowanie()`)** — rubryka „uwagi". Domyślnie pusta i to jest w porządku.

I jeszcze jedno, co nazywa się **zasadą Hollywood**: *„nie dzwoń do nas, my zadzwonimy do ciebie"*. Podklasa **nie woła** `generuj()`. Ona tylko dostarcza treść. Szablon woła ją we właściwym momencie. To odwrócenie sterowania jest sednem wzorca — i jego najczęstszą pomyłką. Jeżeli Twoja podklasa woła `super()._dane()` we własnym `main()`, to znaczy, że nie używasz szablonu, a tylko dziedziczysz.

### 1.3. Ctrl+Z, czyli Command

Naciskasz `Ctrl+Z`. **Cofnąłeś zmianę.** Ale jak to możliwe, że edytor to potrafi, a Twój skrypt nie? Bo edytor **nie zapisuje efektu** — zapisuje **operację**. Każde naciśnięcie klawisza to obiekt: „wstaw znak `a` na pozycji 155", „usuń znak na pozycji 155", „zmień czcionkę akapitu 3". I **każdy taki obiekt ma swoją odwrotność**.

Twój skrypt w module 24 modyfikował komórki bezpośrednio:

```python
ws["C5"] = 999          # ← efekt. Gdzie jest odwrotność? Nie ma jej.
```

Command mówi: nie zapisuj efektu, **zapisz żądanie**. Obiekt „ustaw C5 na 999" **wie, jak się wykonać i jak się cofnąć**. I dopiero wtedy masz:

- **cofanie** (`Ctrl+Z`),
- **audyt** — kto, co i kiedy (`DZIENNIK` komend),
- **makro** — komenda złożona z komend (i to jest Composite z modułu 24),
- **transakcję** — „zastosuj wszystko, sprawdź, potem zapisz".

Ostatni punkt jest dla nas najważniejszy i **bardzo praktyczny**: openpyxl **nie ma transakcji**. Ale Ty możesz je mieć — bo dopóki nie wywołasz `wb.save()`, **nic nie trafiło na dysk**. Twoją transakcją jest więc: **stos komend w pamięci + jeden atomowy zapis na końcu**. Jeżeli coś pójdzie źle w połowie — cofasz stos i **nie zapisujesz**. Plik na dysku pozostaje nietknięty, bo nigdy go nie dotknąłeś.

### 1.4. Kontroler z checklistą, czyli Visitor

Chodzisz po hali produkcyjnej z **checklistą**. Trasa jest **zawsze ta sama**: hala A → hala B → magazyn → kontrola jakości. Ale za każdym razem masz **inną listę**:

- poniedziałek: „szukam pustych stanowisk",
- wtorek: „szukam maszyn używanych bez przerwy dłużej niż 8 godzin",
- środa: „liczę metryki: ile stanowisk, ile maszyn, ile osób".

**Trasa ta sama, lista inna.** To jest Visitor: przejście po strukturze jest **stałe**, a **reakcja na to, co się znajdzie**, jest wymienna.

I jest jeden krytyczny szczegół, którego większość podręczników nie podkreśla, a który jest w tym kursie **kluczowy**: **kontroler z checklistą niczego nie zmienia, dopóki nie wróci do biura.** Jeżeli w trakcie obchodu zaczniesz przestawiać maszyny, to:
- nie wiesz już, co widziałeś (bo widok się zmienił),
- nie policzysz nic wiarygodnie,
- a jeśli przestawiasz **to, po czym chodzisz**, to możesz się zgubić.

W kodzie jest dokładnie tak samo. Gość, który w trakcie iterowania **usuwa wiersze**, gubi komórki i psuje własną pętlę (moduł 07 ostrzegał o tym przy `insert_rows`). Dlatego wzorzec Visitor w tym kursie ma **dwie fazy**:

1. **Przemieć i zbierz** — pełny obchód, zero modyfikacji, wynikiem jest lista znalezisk.
2. **Zastosuj** — po zakończeniu obchodu, na podstawie zebranej listy.

A najlepszy sposób fazy drugiej to... **komendy** z sekcji 1.3. Dzięki temu gość `Anonimizuj` **nie zmienia** niczego, tylko **produkuje komendy**, które idą na stos. I można je cofnąć. Zauważ, jak te dwa wzorce się domykają: Visitor mówi **gdzie** zmienić, Command mówi **jak** zmienić (i jak cofnąć).

### 1.5. Prawo budowlane jako przepis, czyli Specification

Zamiast mówić „ten budynek jest zbyt wysoki, ale tylko wtedy, gdy stoi w strefie A i nie ma pozwolenia, chyba że jest zabytkiem", piszesz **przepis**:

```text
( Wysokosc > 30m  I  Strefa == "A"  I  NIE MaPozwolenia )
    LUB  ( Zabytkowy  I  Strefa == "A" )
```

Ten przepis da się:
- **sprawdzić dla konkretnego budynku** (`jest_spelniona(budynek)`),
- **złożyć z prostszych przepisów** (`and_`, `or_`, `not_`),
- **zapisać na tabliczce** w formie zrozumiałej dla kogoś innego (bo to jest osobny obiekt, nie warunek ukryty w kodzie).

I tu jest rzecz, dla której ten wzorzec jest w naszym kursie tak cenny. **Ten sam przepis może być użyty do DWÓCH różnych rzeczy:**

1. **do filtrowania danych w Pythonie** — przed zapisem do arkusza (`jest_spelniona(wiersz)`),
2. **do wygenerowania reguły formatowania warunkowego w Excelu** (moduły 13–14).

Czyli definiujesz „duża sprzedaż" **raz** — i ta jedna definicja filtruje rekordy **i** koloruje je na czerwono. Bez tego masz dwa warunki: jeden w Pythonie, drugi w formule Excela, i **z czasem się rozjeżdżają** (ktoś zmieni próg w jednym, nie w drugim).

Ale trzeba być uczciwym: **nie każdy przepis da się zapisać na tabliczce.** Prawo budowlane wyraża się w formułach Excela tylko do pewnego stopnia. Jeżeli przepis brzmi „klient jest na mojej liście VIP z bazy CRM", to **nie ma odpowiednika w Excelu** — i wtedy mówisz wprost: filtruj w Pythonie, koloru nie będzie. To jest lepsze niż **udawanie**, że kolor działa, i generowanie raportu, w którym 200 komórek udaje sprawdzone.

### 1.6. Tablica świetlna na hali, czyli Observer

Na hali wisi **tablica świetlna**. Linia produkcyjna **publikuje** dwa zdarzenia: „produkcja 500 sztuk" i „awaria prasy nr 3". Tablica wyświetla postęp. **Nikt nie dzwoni do tablicy**, żeby jej powiedzieć o postępie — linia po prostu pracuje i wysyła sygnał, a tablica sama wie, co z nim zrobić.

I teraz **uczciwa strona tej analogii**: tablica ma sens, gdy jest **kilku odbiorców** (tablica + księgowość + system jakości). W małym warsztacie, gdzie pracujesz sam, tablica jest **zbędnym kablem** — wystarczy, że sam wiesz, ile zrobiłeś. W kodzie: w projekcie z jednym odbiorcą postępu (`log.info` w konsoli) Observer to nadmiar. Pojawia się, gdy:
- postęp ma trafić **do dwóch różnych miejsc** (log + pasek w GUI),
- ostrzeżenia mają **zbierać się** i trafiać do raportu końcowego,
- o zdarzeniu „utracono sparkline'y" ma dowiedzieć się **ktoś inny** niż ten, kto zapisuje.

### 1.7. Podaj dalej, czyli Chain of Responsibility

Wiersz danych (jeden rekord od klienta) przechodzi przez **stanowiska**:

```text
Oczyszczenie → Walidacja obecnosci → Normalizacja regionu → Konwersja typow → Arkusz
```

Każde stanowisko może:
- **poprawić** wiersz i podać dalej,
- **odrzucić** wiersz (i to jest koniec drogi — wiersz nie trafia do arkusza),
- **przepuścić** bez zmian.

I teraz różnica wobec Decoratora z modułu 24, bo to jest najczęstsze pomieszanie:

| | **Decorator** | **Chain of Responsibility** |
|---|---|---|
| Co robi | **owija** funkcję, dodaje warstwę | **procesuje** obiekt i może go zatrzymać |
| Czy może przerwać? | **Nie.** Zawsze woła wnętrze | **Tak.** Zwraca nic i łańcuch się kończy |
| Kto o czym wie? | Nie wie o innych warstwach | Zna **następne** ogniwo |
| Analogia | opakowanie prezentu | stanowiska na taśmie |

**Praktyczna rada, którą zapamiętaj:** jeżeli łańcuch **nigdy nie zatrzymuje** wiersza, to nie potrzebujesz łańcucha. Wystarczy lista funkcji:

```python
for krok in kroki:
    wiersz = krok(wiersz)
```

Łańcuch ma sens **wtedy**, gdy któryś krok **decyduje o losie** obiektu. To ta sama uczciwość, co przy Strategy: **wzorzec musi mieć uzasadnienie, nie nazwę.**

### 1.8. Jedna decyzja na wszystkie siedem wzorców

Popatrz, co wszystkie te wzorce mają wspólnego:

- **Strategy** — wymienny sposób pracy na tym samym obiekcie,
- **Template Method** — stała rama, wymienna treść,
- **Command** — zmiana jako obiekt,
- **Visitor** — stała trasa, wymienna lista kontrolna,
- **Specification** — warunek jako obiekt,
- **Observer** — zdarzenie jako obiekt,
- **Chain of Responsibility** — krok przetwarzania jako obiekt.

To jest **jeden** pomysł: **wynieś decyzję i zachowanie z kodu do obiektu.** W module 24 wynosiliśmy **strukturę** (fasada, adapter, model). Tu wynosimy **zachowanie**. I dlatego wzorce behawioralne są trudniejsze: w module 24 miałeś model raportu — namacalny, z drzewem sekcji. Tu masz **zachowanie**, którego nie da się obejrzeć — trzeba je **przetestować**, żeby uwierzyć, że działa (i dlatego cały ten moduł kończy się testami, nie klasami).

## 2. Teoria

### 2.1. Strategy — definicja, kontekst i granica między dwiema osiami

**Intuicja.** Wymienna końcówka. Ten sam obiekt, różny sposób pracy.

**Analogia.** Wkrętarka (sekcja 1.1).

**Definicja.** **Strategy** to rodzina wymiennych algorytmów o wspólnym interfejsie, których wybór następuje **w czasie wykonania** i **poza** algorytmem (w kontekście albo w kodzie wołającym). Kontekst zna tylko kontrakt.

Trzy role:

- **Strategia** (`StrategiaFormatowania`) — kontrakt. `zastosuj(widok)`.
- **Strategie konkretne** (`Zebrowe`, `Finansowe`, `Warunkowe`) — implementacje.
- **Kontekst** — obiekt, który **używa** strategii, ale jej **nie zna**. W naszym przypadku to `BazowyRaport` albo `RaportExcelRozszerzony`, który ma pole `formatowanie: StrategiaFormatowania`.

#### Kiedy Strategy wygrywa

- **Warianty są ortogonalne do danych.** Ten sam raport może być gęsty, schowkowy albo finansowy. Dane się nie zmieniają, format się zmienia.
- **Wariantów jest więcej niż dwa.** Przy dwóch `if` jest lepszy. Przy pięciu — już nie.
- **Wybór powtarza się w wielu miejscach.** To najsilniejszy sygnał: jeśli ten sam `if` kopiujesz do trzech funkcji, masz Strategię.
- **Warianty mają różne zależności.** `Warunkowe` potrzebuje specyfikacji (moduł 25, sekcja 2.5), `Finansowe` nie potrzebuje nic. Strategy pozwala każdej końcówce mieć **własne** zależności.

#### Kiedy Strategy to nadmiar

- **Jeden wariant.** „Może kiedyś dodam drugi" to nie argument. To antywzorzec numer 1 z tego modułu.
- **Warianty różnią się jednym parametrem.** Nie potrzebujesz `StrategiaKoloruCzerwonego` i `StrategiaKoloruNiebieskiego` — potrzebujesz strategii przyjmującej kolor. **Klasa na każdy kolor to nie wzorzec, to mnożenie bytów.** Test: jeżeli dwie strategie mają identyczne ciała, a różnią się jednym polem, to jedna klasa z polem.
- **Wybór jest znany w czasie pisania kodu.** Jeżeli wiesz, że raport `A` zawsze jest na zielono, a `B` zawsze na szaro — wersja z konfiguracji (moduł 23, `RaportSpec`) jest prostsza. Strategy jest po to, żeby wybór był **dynamiczny i w jednym miejscu**, nie po to, żeby istniał.

#### Najważniejsza rzecz w tym wzorcu: dwie osie to nie jedna strategia

Nasz problem brzmi: **5 typów raportów × 3 warianty eksportu**. Kuszące i błędne rozwiązanie: **jedna** strategia `StrategiaRaportu` z piętnastoma podklasami (`SprzedazXlsx`, `SprzedazBajty`, `SprzedazMarkdown`, `MagazynXlsx`, ...). Piętnaście klas i piętnaście miejsc do edycji, gdy dojdzie szósty typ.

Rozwiązanie poprawne: **dwie niezależne osie**:

- **oś pierwsza** — *co* i *jak* wygląda: `StrategiaFormatowania` (5 wariantów),
- **oś druga** — *dokąd* to idzie: `StrategiaEksportu` (3 warianty).

Razem: **8 klas**, a nie 15. Dodanie szóstego typu raportu: **+1 klasa**. Dodanie czwartego kanału wyjścia: **+1 klasa**. To jest matematyka, którą trzeba zobaczyć raz, żeby zrozumieć, po co ten wzorzec istnieje.

I zwróć uwagę, co powstało: **dwie strategie użyte razem, jako dwie niezależne osie, to dokładnie Bridge** — wzorzec zapowiedziany w module 24, w sekcji 1.6. Most: abstrakcja (typ raportu) i implementacja (kanał wyjścia) przesuwają się niezależnie, a nie są ze sobą przyklejone.

### 2.2. Template Method — szkielet, haczyki i ich dwa rodzaje

**Intuicja.** Stała rama, wymienna treść.

**Analogia.** Formularz urzędowy (sekcja 1.2).

**Definicja.** **Template Method** to metoda w klasie bazowej, która definiuje **kolejność kroków** algorytmu, wywołując po drodze metody, które podklasa **może albo musi** zaimplementować.

```python
class BazowyRaport(ABC):
    @final                                  # ← mypy: nie nadpisuj tego!
    def generuj(self, cel) -> Path:
        self._waliduj()                     # krok stale
        self._przygotuj()                   # krok stale
        naglowki = self._naglowek()         # HACZYK WYMAGANY
        ...
```

#### Haczyk wymagany vs opcjonalny — i pułapka, która z niego wynika

```python
    @abstractmethod
    def _naglowek(self) -> list[str]: ...        # WYMAGANY: ABC wymusi
    @abstractmethod
    def _wiersze(self) -> Iterable: ...          # WYMAGANY

    def _podsumowanie(self) -> None:
        """Haczyk OPCJONALNY - domyslnie raport nie ma podsumowania."""
        return None

    def _formatowanie(self) -> None:
        """Haczyk OPCJONALNY."""
        return None
```

**I teraz pułapka numer 2 z tego modułu.** Haczyk opcjonalny z domyślnym `return None` jest **bezpieczny technicznie i niebezpieczny organizacyjnie**: jeżeli ktoś napisze nowy raport i **zapomni** o podsumowaniu, nic się nie stanie. Raport powstanie — **bez podsumowania**. Nikt nie zauważy tego w kodzie, bo kompiluje się i działa. Zauważy odbiorca, któremu brakuje kluczowej liczby.

Trzy sposoby obrony, w kolejności od najprostszego:

1. **Haczyk ważny → `@abstractmethod`.** Jeżeli podsumowanie jest wymagane w każdym raporcie, to niech będzie wymagane. Puste podsumowanie to zawsze świadoma decyzja: `def _podsumowanie(self): return None`.
2. **Haczyk opcjonalny → zwraca coś do logu.** `def _podsumowanie(self) -> list[tuple[str, object]]: return []`. Kontekst widzi, że lista jest pusta, i może to **zaraportować** w metrykach (moduł 24). Pusta lista to informacja; `None` to cisza.
3. **Test katalogowy.** Test, który przechodzi po wszystkich podklasach `BazowyRaport` i sprawdza, że ich `_podsumowanie()` coś zwraca. Brzmi jak przesada — do momentu, gdy złapie błąd przed klientem.

#### Zasada Hollywood i `@final`

Podklasa **nie woła** `generuj()`. Podklasa **dostarcza** treść. To odwrócenie sterowania jest istotą wzorca.

Python nie ma wbudowanego `final`, ale ma `typing.final`: adnotacja, którą sprawdza `mypy`, a interpreter **ignoruje**. To znaczy:

```python
from typing import final

class BazowyRaport(ABC):
    @final
    def generuj(self, cel) -> Path:
        ...
```

Jeżeli podklasa nadpisze `generuj`, `mypy` to zgłosi w CI, a **program nadal się uruchomi** i będzie działał źle. **Adnotacja bez egzekwowania w czasie działania jest deklaracją intencji, nie zabezpieczeniem** — i trzeba być tego świadomym. Jeżeli chcesz egzekwować, dopisz test:

```python
def test_podklasy_nie_nadpisuja_generuj():
    for klasa in wszystkie_podklasy(BazowyRaport):
        assert "generuj" not in klasa.__dict__, f"{klasa.__name__} nadpisuje generuj()"
```

### 2.3. Command — obiekt-żądanie, cofanie i transakcja

**Intuicja.** Zapisz żądanie, nie efekt.

**Analogia.** `Ctrl+Z` (sekcja 1.3).

**Definicja.** **Command** to obiekt, który reprezentuje **żądanie** i zawiera wszystko, co potrzebne do jego wykonania **i cofnięcia**.

```python
class Komenda(ABC):
    @property
    @abstractmethod
    def opis(self) -> str: ...
    @abstractmethod
    def wykonaj(self, kontekst) -> None: ...
    @abstractmethod
    def cofnij(self, kontekst) -> None: ...
```

Cztery pożytki, o których mówi zadanie:

| Pożytek | Jak realizowany |
|---|---|
| **Cofanie** | `cofnij()` na każdym obiekcie + `StosKomend` |
| **Log audytowy** | każda komenda ma `opis` i `autor`; stos zapisuje **kto i co** |
| **Makro** | `MakroKomend` = lista komend, która sama jest komendą (Composite!) |
| **Transakcyjność** | stos w pamięci + **jeden** zapis na końcu; brak zapisu = rollback |

#### Command to obiekt, który MUSI mieć stan

W module 24 cały kurs dbał o niemutowalność: `frozen=True` na sekcjach, `WierszRaportu`, `Kolumna`, style. I słusznie — bo niemutowalny obiekt jest bezpieczny, gdy przechodzi przez wiele warstw.

**Command jest wyjątkiem i trzeba to powiedzieć wprost.** Komenda musi **zapamiętać poprzednią wartość**, żeby móc się cofnąć. To jest **jedyny obiekt w tym kursie, który z założenia ma stan mutowalny** — i ma go **dlatego**, bo bez niego nie ma cofania. Analogia: pilot do bramy nie „jest" bramą ani jej stanem — ale **pamięta**, że ją otworzył, bo inaczej nie mógłby jej zamknąć.

I dlatego komenda ma regułę: **stan jest prywatny i zmieniany tylko przez `wykonaj()`**. Nigdy w konstruktorze (to jest pułapka numer 3 z tego modułu):

```python
# ANTYWZORZ: komenda modyfikuje skoroszyt JUZ w konstruktorze
class KomendaUstawKomorke:
    def __init__(self, kontekst, wiersz, kolumna, wartosc):
        self._poprzednia = kontekst.komorka(wiersz, kolumna).value
        kontekst.komorka(wiersz, kolumna).value = wartosc   # ← JUZ ZMIENIL!
```

Dlaczego to jest katastrofa? Bo **tworzenie komendy i jej wykonanie to dwie różne rzeczy** — a tu zostały zlane w jedno. Konsekwencje:

1. **Nie zbudujesz makra bez wykonania.** Komenda ma być obiektem, który można **przygotować, obejrzeć, zalogować i wykonać później**. W antywzorcu każde utworzenie już zmienia plik.
2. **Nie sprawdzisz makra przed wykonaniem.** Makro „usuń kolumnę, dodaj nagłówek" — chcesz obejrzeć listę operacji i powiedzieć „ok, wykonaj". W antywzorcu makro już się wykonało.
3. **Nie wdrożysz potwierdzenia.** `BramaZapisu` z modułu 24 blokuje zapis, ale nie blokuje modyfikacji w pamięci. Jeżeli komenda zmienia w konstruktorze, to „podejrzyj i zatwierdź" przestaje mieć sens — bo zmiany już zaszły.
4. **Log audytowy kłamie.** `wykonaj()` zostanie wywołane i zalogowane — ale zmiana zaszła **wcześniej**, więc kolejność w logu nie odpowiada kolejności zmian.

#### Granica transakcji: `wb.save()`

To jest miejsce, w którym trzeba być precyzyjnym.

**openpyxl nie ma transakcji.** Nie ma `wb.begin()`, `wb.commit()`, `wb.rollback()`. Nie ma nawet pojęcia „brudnego" arkusza. Modyfikacja modelu jest natychmiast widoczna **w pamięci** — i tylko w pamięci, dopóki nie zapiszesz.

Twoja transakcja to więc umowa, którą sam zawierasz:

```text
1. Odczytaj dane (bez modyfikacji oryginalu)
2. Zbuduj model w pamieci
3. Wykonaj komendy na KOPII roboczej (stos komend)
4. Sprawdz wynik (podglad, walidacja, metryki)
5. Zapisz ATOMOWO (mkstemp + os.replace)   <-- TU jest commit
6. W razie bledu: cofnij stos, NIE zapisuj
```

Trzy konsekwencje, które trzeba znać:

1. **Brak zapisu = pełny rollback.** Najprostszy i najskuteczniejszy mechanizm cofania, jaki masz. Jeżeli coś pójdzie źle przed `save()`, **wystarczy nie zapisać**. Plik na dysku jest nietknięty, bo nigdy go nie nadpisałeś (moduł 18).
2. **Zapis atomowy chroni przed połowicznym plikiem.** `mkstemp` + `os.replace` sprawia, że albo plik jest nowy w całości, albo jest stary w całości. Nie ma stanu „w połowie zapisany".
3. **Commit jest nieodwracalny z poziomu stosu.** Po `save()` cofanie komend zmienia **pamięć**, ale nie plik — i nie zmieni go, dopóki nie zapiszesz ponownie. To nie jest błąd, to jest granica transakcji. **Dokumentuj ją w kodzie**:

```python
def zapisz(self, cel) -> Path:
    """COMMIT transakcji. Po tym wywolaniu stos komend NIE cofa pliku -
    cofanie dziala tylko w pamieci, do nastepnego zapisu."""
```

### 2.4. Visitor — double dispatch, dwie fazy i dlaczego modyfikacja psuje iterację

**Intuicja.** Stała trasa, wymienna lista kontrolna.

**Analogia.** Kontroler z checklistą (sekcja 1.4).

**Definicja.** **Visitor** to wzorzec, w którym operacja na strukturze obiektów jest wydzielona **na zewnątrz** tej struktury, a wybór operacji zależy od **dwóch** rzeczy: typu elementu i typu gościa.

#### Double dispatch (podwójne wywołanie zwrotne) — po co to w ogóle

W Pythonie nie ma przeciążania metod po typie, więc nie masz klasycznego problemu „jak wybrać metodę dla dwóch typów". Ale masz **problem projektowy**, który double dispatch rozwiązuje: **gdzie umieścić operację**.

Naiwne rozwiązanie: metody w klasach elementów.

```python
class Arkusz:
    def policz_puste_kolumny(self): ...
    def wykryj_formuly(self): ...
    def zbierz_metryki(self): ...
    def anonimizuj(self): ...
```

Problem? **Każda nowa analiza = zmiana klasy `Arkusz`.** A arkuszów jest w openpyxl kilkanaście typów (`Worksheet`, `Chartsheet`, `ReadOnlyWorksheet`), więc rośnie to w dwóch wymiarach jednocześnie. To jest dokładnie złamanie zasady otwarte-zamknięte z modułu 22: każda nowa analiza wymaga modyfikacji istniejącego kodu.

Rozwiązanie Visitor: **analiza jest na zewnątrz** (w gościu), a struktura udostępnia tylko **jedną** metodę „przyjmij gościa" — albo w naszej wersji: **driver** chodzi po strukturze i woła metody gościa. Nowa analiza = **nowa klasa gościa**, zero zmian w `Arkusz`.

```python
class OdwiedzSkoroszyt(ABC):
    """Bazowy gosc. Metody odwiedz_* NIE MODYFIKUJA - zbieraja."""
    def __init__(self) -> None:
        self._znaleziska: list[Znalezisko] = []

    def odwiedz_arkusz(self, nazwa: str, ws) -> None: ...
    def odwiedz_komorke(self, arkusz: str, wiersz: int, kolumna: str, komorka) -> None: ...
    def odwiedz_tabele(self, arkusz: str, tabela) -> None: ...

    @final
    def przemieć(self, wb) -> list[Znalezisko]:
        """SZKIELET przejscia (to tez Template Method!)."""
        for nazwa in wb.sheetnames:
            ws = wb[nazwa]
            self.odwiedz_arkusz(nazwa, ws)
            for wiersz in ws.iter_rows():
                for komorka in wiersz:
                    if komorka.value is not None:
                        self.odwiedz_komorke(nazwa, komorka.row,
                                             get_column_letter(komorka.column), komorka)
            for tabela in ws.tables.values():
                self.odwiedz_tabele(nazwa, tabela)
        return self.wynik()
```

Zwróć uwagę: **`przemieć()` też jest metodą szablonową.** To pokazuje, że wzorce się składają — Visitor dla „co robić" i Template Method dla „jak chodzić". I że nie jest to wybór „albo-albo", tylko warstwy.

#### Dwie fazy: zbieraj, potem zastosuj

To jest najważniejsza część tego wzorca w kontekście Excela. Wyobraź sobie gościa `Anonimizuj`, który w metodzie `odwiedz_komorke` od razu zamienia nazwisko na pseudonim:

```python
# ANTYWZORZ - modyfikuje w trakcie iteracji
def odwiedz_komorke(self, arkusz, wiersz, kolumna, komorka):
    if self._czy_dane_osobowe(komorka.value):
        komorka.value = self._pseudonim(komorka.value)      # ← zmiana W TRAKCIE
```

Co się dzieje? Trzy problemy naraz:

1. **`iter_rows()` zwraca komórki na bieżąco** i `ws.max_row`/`max_column` są wyliczane na podstawie bieżącej zawartości. Zmiana wartości nie zmienia wymiarów, ale **zmiana typu** (np. z liczby na tekst) może sprawić, że Excel zobaczy inny plik niż Twój model — a przy **usuwaniu** lub **wstawianiu** wierszy w trakcie iteracji **zgubisz komórki** (moduł 07 ostrzegał: to jak zmiana listy podczas iterowania po niej).
2. **Nie policzysz tego, co znalazłeś.** „Zanonimizowano 412 komórek" — a ile ich było? Musisz liczyć w locie, co miesza raportowanie z przetwarzaniem.
3. **Nie cofniesz tego.** Zmiana jest efektem, nie żądaniem. Jeżeli chcesz wrócić do oryginału — musisz mieć kopię pliku.

Dlatego **Visitor w tym kursie działa w dwóch fazach**:

```text
FAZA 1 (obchod):  odwiedz_* nic nie zmieniaja; zbieraja Znaleziska
                  albo PRODUKUJA KOMENDY (jeszcze nie wykonane)
FAZA 2 (zastosowanie): komendy ida na StosKomend i wykonuja sie PO obchodzie
```

Zysk jest konkretny: pełny raport „co znaleziono" **przed** jakąkolwiek zmianą, możliwość **potwierdzenia**, możliwość **cofnięcia**, i — co najważniejsze — **brak modyfikacji struktury, po której chodzisz**. To ta sama lekcja, co Composite z modułu 24: **model to nie efekt**.

### 2.5. Specification — kompozycja i most do formatowania warunkowego

**Intuicja.** Warunek jako obiekt, który da się składać.

**Analogia.** Prawo budowlane jako przepis (sekcja 1.5).

**Definicja.** **Specification** to obiekt reprezentujący predykat (warunek logiczny), który można **łączyć** z innymi predykatami i **sprawdzać** dla kandydata.

```python
class Specyfikacja(ABC):
    @abstractmethod
    def nazwa(self) -> str: ...
    @abstractmethod
    def jest_spelniona(self, wiersz: Mapping[str, object]) -> bool: ...
    @abstractmethod
    def na_formule(self, adres: str) -> str | None: ...

    def __and__(self, inna): return I(self, inna)          # a & b
    def __or__(self, inna):  return Lub(self, inna)        # a | b
    def __invert__(self):    return Nie(self)              # ~a
```

Uczciwa uwaga do operatorów: `~` (tylda) działa, bo `__invert__` da się przeciążyć. **`not` nie działa** — `not` jest operatorem języka, nie metodą, i nie da się go nadpisać. Jeżeli ktoś napisze `if not spec:` na obiekcie specyfikacji, dostanie `False`, bo **każdy obiekt jest prawdziwy** — i to nie jest błąd Pythona, tylko konsekwencja tego, że `not` sprawdza „prawdziwość", a nie logikę. Dlatego w kodzie użyj **`~spec`** i **nigdy** `not spec`.

#### Most do modułów 13–14 — i jego uczciwe granice

Reguły formatowania warunkowego w openpyxl (`FormulaRule`) przyjmują **formułę Excela**, która nie zaczyna się od `=` (pułapka z modułu 13!). Specyfikacja może dostarczyć tę formułę:

```python
spec = WiekszeNiz("cena", 100) & Rowne("region", "PL")
spec.jest_spelniona({"cena": 150, "region": "PL"})   # True  (filtr w Pythonie)
spec.na_formule("$D2")                               # "AND($D2>100,$E2=\"PL\")" (kolor)
```

I teraz **trzy granice, które trzeba znać i powiedzieć wprost**:

1. **Nie każda specyfikacja ma odpowiednik w Excelu.** Warunek „klient jest na liście VIP z bazy" **nie da się** zapisać formułą. `na_formule()` zwraca `None`, generator reguł **loguje ostrzeżenie** i **nie tworzy reguły** — dane są filtrowane, koloru nie ma. **Nie udawaj, że jest.** Raport, w którym 40 wierszy jest kolorowanych „na oko", a 12 nie, jest gorszy niż raport bez kolorów.
2. **Adres w formule musi być zgodny z adresem zakresu.** Reguła na zakresie `"F2:F100"` potrzebuje formuły o **punkcie widzenia `F2`** — czyli odwołań **kolumnowo absolutnych, wierszowo względnych**: `$F2`, `$E2`. Napisanie `$F$2` sprawi, że wszystkie wiersze będą sprawdzać wiersz **drugi**. To najczęstsza pomyłka tego mostu i powtarza pułapkę z modułu 13 w praktyce.
3. **Nie wstrzykuj tekstu użytkownika do formuły bez escapowania.** Jeżeli specyfikacja brzmi „region zawiera X", a `X` pochodzi od użytkownika, to `X` z cudzysłowem **rozerwie formułę** — a w skrajnym przypadku wstrzyknie formułę (moduł 21, formula injection). Podwój cudzysłowy (`"` → `""`), a **dane z zewnątrz odrzucaj, jeżeli zawierają znaki sterujące**.

#### Rozjazd: Python a Excel

To jest drobiazg, który jednak realnie psuje raporty:

```python
# Python: 150.0 > 100.0 -> True
# Excel:  =AND($D2>100,...)  gdzie w komorce jest 100.0000000001 -> False
```

Różnice w reprezentacji `float` (moduł 05: `Decimal` vs `float`, precyzja kwot) sprawiają, że ten sam próg może dać **inny wynik** w Pythonie i w Excelu. Zabezpieczenia:

1. **Zaokrąglenie po obu stronach** — w formule `ROUND($D2,2)>100`, w Pythonie `round(wartosc, 2) > 100`.
2. **Progi nie leżą na granicy** — jeżeli wiesz, że wartości to liczby z dwoma miejscami po przecinku, próg 100,00 jest bezpieczny; próg 100,005 nie.
3. **Test spójności** — test, który dla losowej próbki danych sprawdza, że `jest_spelniona(wiersz)` daje ten sam wynik co formuła. To jest realne, bo wzór na formułę da się policzyć… ręcznie w Pythonie (masz przecież **tę samą** specyfikację). W praktyce: `assert spec.jest_spelniona(w) == ocena_formuly_przez_python(spec, w)`. Brzmi jak testowanie testu — i nim jest, ale łapie rozjazd.

### 2.6. Observer — kiedy zarabia na siebie, a kiedy jest kablem do niczego

**Intuicja.** Zdarzenie jako obiekt, wielu odbiorców.

**Analogia.** Tablica świetlna na hali (sekcja 1.6).

**Definicja.** **Observer** (albo **Event bus**) to mechanizm, w którym **publisher** publikuje zdarzenie, a **subskrybenci** reagują — bez wiedzy o sobie nawzajem.

Minimalna implementacja ma **cztery** decyzje projektowe, o których trzeba pomyśleć:

| Decyzja | Opcja | Nasz wybór i uzasadnienie |
|---|---|---|
| Co jest zdarzeniem? | `str` + `dict` / klasa / enum | klasa `Zdarzenie(typ, payload)` — typy jako stałe, payload jako `dict` |
| Czy można odsubskrybować? | brak / tak / tak, zwracając funkcję | `subskrybuj()` zwraca **funkcję anulującą** — najprostszy idiomatyczny Python |
| Co z wyjątkiem w obsłudze? | przerwać / połknąć / zebrać | **zebrać** i wystawić w `autobus.bledy` — nie połknąć (pułapka 2 z modułu 24!) |
| Czy zdarzenia są kolejkowane? | natychmiast / kolejka | natychmiast; kolejka to już inny wzorzec (i inny problem) |

Ten trzeci punkt jest najważniejszy i łączy się z modułem 24: **obsługa zdarzenia, która rzuca, nie może zatrzymać generowania raportu** — ale **nie może też zniknąć**. Rozwiązanie: złap wyjątek, zapisz go w `autobus.bledy` z nazwą handlera, **kontynuuj publikację** do pozostałych odbiorców, a na końcu raportu sprawdź `bledy` i **dołącz ostrzeżenie do raportu**. Wtedy błąd nie jest ani cichy, ani zabójczy.

**Kiedy Observer to nadmiar:** w projekcie z jednym konsumentem postępu. `logging.info()` jest wtedy prostsze, testowalne i bez własnej infrastruktury. Observer ma sens, gdy odbiorców jest **dwóch lub więcej**, gdy zdarzenia muszą **zbierać się** (ostrzeżenia na koniec raportu) albo gdy o utracie funkcji pliku ma wiedzieć **ktoś inny** niż zapisujący.

### 2.7. Chain of Responsibility — łańcuch vs pipeline vs decorator

**Intuicja.** Stanowiska na taśmie; każde może poprawić, przepuścić albo odrzucić.

**Analogia.** Podaj dalej (sekcja 1.7).

**Definicja.** **Chain of Responsibility** to łańcuch obiektów-przetwórców, w którym każde ogniwo **decyduje**, czy przetworzyć obiekt, czy przekazać go dalej — i **może zatrzymać** przetwarzanie.

Trzy konstrukcje, które łatwo pomylić:

| | **Pipeline** (lista funkcji) | **Decorator** | **Chain of Responsibility** |
|---|---|---|---|
| Kolejność | stała | stała (zagnieżdżenie) | stała, ale **przerwanie możliwe** |
| Może przerwać? | nie | nie | **tak** |
| Może przekierować? | nie | nie | tak (np. do „kosza odrzuconych") |
| Kiedy używać | kroki zawsze się wykonują | dodajesz efekt wokół | kroki **decydują o losie** |

**Praktyczna rada:** jeżeli Twój łańcuch nigdy nie przerwie, **przepisz go na listę funkcji**. To jest ten sam antywzorzec co „Strategy, która ma jeden wariant" — wzorzec bez uzasadnienia jest kosztem bez zysku. Ale jeżeli przetwarzanie ma **odrzucać** wiersze (walidacja), **przekierowywać** je (do arkusza „Odrzucone") i **logować**, gdzie odpadły — łańcuch jest właściwym narzędziem.

I jeszcze jedno: **łańcuch musi mieć ostatnie ogniwo**, które decyduje, co zrobić z przepuszczonym obiektem. Bez tego masz cichy `None` na końcu (pułapka 9 z tego modułu) — czyli kolejny niemy tryb awarii.

### 2.8. Tabela decyzyjna: problem → wzorzec → mini-przykład

| Problem w kodzie | Wzorzec | Mini-przykład |
|---|---|---|
| Ten sam `if/elif` w trzech funkcjach | **Strategy** | `STRATEGIE[spec.format]` |
| Warianty różnią się jednym parametrem | **konfiguracja**, nie wzorzec | `StrategiaGesta(kolor="EEEEEE")` |
| 5 typów × 3 kanały wyjścia | **Strategy × 2 osie (Bridge)** | `Formatowanie` + `Eksport` osobno |
| Stała kolejność kroków, różna treść | **Template Method** | `generuj()` + `_naglowek()`, `_wiersze()` |
| Trzeba cofnąć ostatnią zmianę | **Command + StosKomend** | `stos.cofnij()` |
| Trzeba wiedzieć, kto zmienił C5 | **Command + log** | `stos.audyt()` |
| Wiele zmian jako jedna operacja | **Makro = Composite komend** | `MakroKomend("korekta", [k1, k2])` |
| Ta sama analiza na wielu arkuszach | **Visitor** | `ZbierzMetryki().przemieć(wb)` |
| Analiza ma też coś zmienić | **Visitor (faza 1) + Command (faza 2)** | `Anonimizuj` produkuje komendy |
| Warunek filtrowania **i** kolorowania | **Specification** | `spec & ~STARY` → filtr + `FormulaRule` |
| Progi w Pythonie i w Excelu się rozjeżdżają | **Specification + test spójności** | `assert spec.jest_spelniona(w)` |
| Długie zadanie, trzeba pokazać postęp | **Observer** | `autobus.publikuj(Postep(...))` |
| Ostrzeżenia zbierane na koniec | **Observer + kolektor** | `ZbieraczOstrzezen.lista` |
| Wiersz przechodzi przez kroki, może odpaść | **Chain of Responsibility** | `Oczysc → Waliduj → Normalizuj → Konwertuj` |
| Kroki zawsze się wykonują | **lista funkcji**, nie CoR | `for krok in kroki: wiersz = krok(wiersz)` |

## 3. Przykłady krok po kroku

Sześć przykładów. Ten sam raport sprzedaży, ewolucja od modułu 24. Wszystkie korzystają z plików `examples_24_*.py` z poprzedniego modułu — leżą w tym samym katalogu, więc importy działają.

Na początku jeden plik z **rozszerzeniem fasady**, bo strategie formatowania potrzebują metod, których fasada z modułu 24 nie miała (moduł 24, sekcja 2.1, „opcja A: rozszerz fasadę").

### Przykład 0 — `examples_25_fasada.py`: rozszerzenie fasady i widok arkusza

**Problem.** Fasada z modułu 24 umiała `naglowek`, `wiersze`, `para`, `tytul`, `tabela`, `szerokosci`, `wykres_slupkowy`, `zapisz`, `do_bajtow`. Strategie potrzebują jeszcze: ustawienia formatu dla kolumny, wstawienia reguły warunkowej, ustawienia wyrównania, ustawień druku i **liczby wierszy danych**. Musimy je dodać — ale **bez wypuszczania `ws` na zewnątrz**.

```python
"""Modul 25, przyklad 0: rozszerzenie fasady + widok arkusza dla strategii.

Importujemy fasade i wyjatki z modulu 24. Ten plik dodaje TYLKO to,
czego wymagaja wzorce behawioralne.
"""

from __future__ import annotations

from collections.abc import Sequence

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from examples_24_01_fasada import BladProtokolu, BladRaportu, RaportExcel


class RaportExcelRozszerzony(RaportExcel):
    """Fasada + operacje potrzebne strategiom."

    UWAGA: nadal obowiazuje test jakosci fasady (modul 24, sekcja 2.1):
    zadna metoda PUBLICZNA nie przyjmuje ani nie zwraca typu openpyxl.
    Regula ("FormulaRule") jest naszym typem, nie openpyxl - i dobrze.
    """

    # ------------------------------------------------------------------
    # ODCZYT STANU - strategie musza wiedziec, na czym pracuja
    # ------------------------------------------------------------------
    def liczba_wierszy_danych(self, klucz: str) -> int:
        """Liczba wierszy danych (BEZ naglowka). Z wlasnego stanu fasady."""
        if klucz not in self._pierwszy_danych:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        return max(0, self._ostatni[klucz] - self._pierwszy_danych[klucz] + 1)

    def kolumny(self, klucz: str) -> tuple[str, ...]:
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        return self._kolumny[klucz]

    def indeks_kolumny(self, klucz: str, nazwa: str) -> str:
        """Nazwa kolumny -> litera ('Cena' -> 'E'). Litera NIE wychodzi na zewnatrz."""
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        kolumny = self._kolumny[klucz]
        if nazwa not in kolumny:
            raise BladRaportu(
                f"Nie ma kolumny {nazwa!r}. Dostepne: {list(kolumny)}"
            )
        return get_column_letter(kolumny.index(nazwa) + 1)

    # ------------------------------------------------------------------
    # FORMATOWANIE
    # ------------------------------------------------------------------
    def format_kolumny(
        self, klucz: str, kolumna: str, number_format: str
    ) -> int:
        """Nadaje format liczbowy calej kolumnie danych. Zwraca liczbe komorek."""
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        ws = self._ws(klucz)
        indeks = self._kolumny[klucz].index(kolumna) + 1
        pierwszy = self._pierwszy_danych[klucz]
        licznik = 0
        for wiersz in range(pierwszy, self._ostatni[klucz] + 1):
            ws.cell(row=wiersz, column=indeks).number_format = number_format
            licznik += 1
        return licznik

    def wyrównaj_kolumny(
        self, klucz: str, kolumny_prawo: Sequence[str], *, naglowek: str = "center"
    ) -> None:
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        ws = self._ws(klucz)
        for nazwa in kolumny_prawo:
            indeks = self._kolumny[klucz].index(nazwa) + 1
            for wiersz in range(self._pierwszy_danych[klucz],
                                self._ostatni[klucz] + 1):
                cell = ws.cell(row=wiersz, column=indeks)
                cell.alignment = Alignment(horizontal="right")

    def paski_wierszy(self, klucz: str, kolor: str = "F2F2F2",
                      co_ile: int = 1) -> int:
        """Naprzemienne tlo wierszy danych. Zwraca liczbe pokolorowanych."""
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        ws = self._ws(klucz)
        n = len(self._kolumny[klucz])
        licznik = 0
        for i, wiersz in enumerate(range(self._pierwszy_danych[klucz],
                                        self._ostatni[klucz] + 1)):
            if i % (co_ile * 2) < co_ile:
                for kol in range(1, n + 1):
                    ws.cell(row=wiersz, column=kol).fill = PatternFill(
                        fill_type="solid", fgColor=kolor
                    )
                licznik += 1
        return licznik

    def regula_warunkowa(self, klucz: str, kolumna: str, regula) -> str:
        """Dodaje regule warunkowa do kolumny danych. `regula` to nanosnik.

        Nanosnik trzyma GOTOWY obiekt Rule openpyxl - strategia go nie buduje,
        zeby nie znala openpyxl. Buduje go fabryka regul (przyklad 5).
        """
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        ws = self._ws(klucz)
        indeks = self._kolumny[klucz].index(kolumna) + 1
        litera = get_column_letter(indeks)
        zakres = (f"{litera}{self._pierwszy_danych[klucz]}:"
                  f"{litera}{self._ostatni[klucz]}")
        ws.conditional_formatting.add(zakres, regula)
        return zakres

    def naglowek_kolor(self, klucz: str, kolumna: str, kolor: str) -> None:
        """Zmienia kolor czcionki w KOMORCE NAGLOWKA (nie w danych)."""
        if klucz not in self._kolumny:
            raise BladProtokolu(f"Arkusz {klucz!r} nie ma naglowka")
        ws = self._ws(klucz)
        indeks = self._kolumny[klucz].index(kolumna) + 1
        istnieje = ws.cell(row=1, column=indeks)
        # style sa NIEMUTOWALNE (modul 09) - kopiujemy i podmieniamy
        from copy import copy

        nowy = copy(istnieje.font)
        nowy.color = kolor
        istnieje.font = nowy

    # ------------------------------------------------------------------
    # DRUK
    # ------------------------------------------------------------------
    def druk_poziomo(
        self, klucz: str, kolumny: str, *, powtorz_naglowek: bool = True
    ) -> None:
        ws = self._ws(klucz)
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.print_area = kolumny
        if powtorz_naglowek:
            ws.print_title_rows = "1:1"
        ws.sheet_properties.pageSetUpPr.fitToPage = True    # ← TRZY ustawienia razem
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    def metadane(self) -> tuple[str, str]:
        """(tytul, autor) - DTO, bez openpyxl."""
        return (self._wb.properties.title or "", self._wb.properties.creator or "")
```

```python
# ----------------------------------------------------------------------
# WIDOK ARKUSZA - "port wejsciowy" dla strategii formatowania
# ----------------------------------------------------------------------
from abc import ABC, abstractmethod


class WidokArkusza(ABC):
    """Co strategia formatowania moze zrobic z arkuszem. NIC wiecej.

    To jest ANTY-przeciek: strategia nie dostaje `ws` ani obiektow stylow.
    Dostaje operacje w slownictwie raportu (kolumna po nazwie, format liczbowy).
    """

    @abstractmethod
    def liczba_wierszy(self) -> int: ...
    @abstractmethod
    def kolumny(self) -> tuple[str, ...]: ...
    @abstractmethod
    def format_kolumny(self, kolumna: str, number_format: str) -> int: ...
    @abstractmethod
    def wyrownaj_prawo(self, kolumny: Sequence[str]) -> None: ...
    @abstractmethod
    def paski(self, co_ile: int, kolor: str) -> int: ...
    @abstractmethod
    def regula(self, kolumna: str, regula) -> str: ...
    @abstractmethod
    def naglowek_kolor(self, kolumna: str, kolor: str) -> None: ...


class WidokFasady(WidokArkusza):
    """Adapter: WidokArkusza -> metody fasady. Zero openpyxl na wierzchu."""

    def __init__(self, raport: RaportExcelRozszerzony, klucz: str) -> None:
        self._raport = raport
        self._klucz = klucz

    def liczba_wierszy(self) -> int:
        return self._raport.liczba_wierszy_danych(self._klucz)

    def kolumny(self) -> tuple[str, ...]:
        return self._raport.kolumny(self._klucz)

    def format_kolumny(self, kolumna: str, number_format: str) -> int:
        return self._raport.format_kolumny(self._klucz, kolumna, number_format)

    def wyrownaj_prawo(self, kolumny: Sequence[str]) -> None:
        self._raport.wyrównaj_kolumny(self._klucz, kolumny)

    def paski(self, co_ile: int = 1, kolor: str = "F2F2F2") -> int:
        return self._raport.paski_wierszy(self._klucz, kolor, co_ile)

    def regula(self, kolumna: str, regula) -> str:
        return self._raport.regula_warunkowa(self._klucz, kolumna, regula)

    def naglowek_kolor(self, kolumna: str, kolor: str) -> None:
        self._raport.naglowek_kolor(self._klucz, kolumna, kolor)
```

**Co się dzieje w pamięci.** `WidokFasady` to **cienki adapter** (moduł 24, sekcja 2.2): trzyma referencję do fasady i klucz arkusza, a każde wywołanie przekazuje dalej. Nie ma własnego stanu poza tymi dwiema referencjami. Jego istnienie kupuje jedną, ale fundamentalną rzecz: **strategia ma w sygnaturach wyłącznie typy nasze, nie openpyxl.** Strategia nie może „przypadkiem" sięgnąć po `ws`, bo pole `ws` nie istnieje w `WidokArkusza` — a to jest **egzekwowalna** wersja testu poziomu 1 z modułu 24. Nie musisz `grep`-ować, żeby wiedzieć, że jest czysto: **nie ma czym przeciec**.

**Co trafi do pliku.** Jeszcze nic — to jest plik infrastruktury, używany przez kolejne przykłady.

### Przykład 1 — Strategy: trzy formatowania i trzy kanały wyjścia

**Problem.** Zostawiliśmy w module 24 rozgałęzienie `if odbiorca == "zarzad": ...`. A teraz dochodzą jeszcze trzy style wizualne i trzy sposoby wyjścia.

**Rozwiązanie.** Dwie niezależne osie: `StrategiaFormatowania` (jak wygląda) i `StrategiaEksportu` (dokąd trafia).

```python
"""Modul 25, przyklad 1: Strategy na dwoch osiach."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

from examples_24_01_fasada import RaportExcel
from examples_24_05_composite import CelMarkdown, GrupaSekcji
from examples_25_fasada import RaportExcelRozszerzony, WidokArkusza, WidokFasady


# ======================================================================
# OS 1: JAK WYGLADA - StrategiaFormatowania
# ======================================================================
class StrategiaFormatowania(ABC):
    """Kontrakt. Kazda strategia robi DOKLADNIE jedna rzecz: formatuje."""

    nazwa: str = "?"

    @abstractmethod
    def zastosuj(self, widok: WidokArkusza) -> dict[str, object]:
        """Formatuje i zwraca METRYKI (co zrobila). Bez metryk nie ma audytu."""


class Zebrowe(StrategiaFormatowania):
    """Naprzemienne tlo. Klasyk dla dlugich list."""

    nazwa = "zebrowe"

    def __init__(self, kolor: str = "F2F2F5") -> None:
        self._kolor = kolor

    def zastosuj(self, widok: WidokArkusza) -> dict[str, object]:
        pomalowane = widok.paski(co_ile=1, kolor=self._kolor)
        return {"strategia": self.nazwa, "pomalowane_wiersze": pomalowane}


class Finansowe(StrategiaFormatowania):
    """Formaty walutowe, wyrownanie do prawej, kursywa dla walut obcych."""

    nazwa = "finansowe"

    def __init__(self, kolumna_kwoty: str, kolumny_liczbowe=()) -> None:
        self._kwota = kolumna_kwoty
        self._liczbowe = tuple(kolumny_liczbowe)

    def zastosuj(self, widok: WidokArkusza) -> dict[str, object]:
        sformatowane = widok.format_kolumny(self._kwota, '#,##0.00 "zl";(#,##0.00)')
        for kolumna in self._liczbowe:
            widok.format_kolumny(kolumna, "#,##0")
        widok.wyrownaj_prawo((self._kwota, *self._liczbowe))
        widok.naglowek_kolor(self._kwota, "1F4E78")
        return {
            "strategia": self.nazwa,
            "sformatowane_komorki": sformatowane + len(self._liczbowe),
            "waluta": "PLN",
        }


class Warunkowe(StrategiaFormatowania):
    """Kolory zalezne od wartosci. Most do modulow 13-14 (patrz przyklad 5)."""

    nazwa = "warunkowe"

    def __init__(self, reguly: tuple[tuple[str, object], ...] = ()) -> None:
        """reguly: ((nazwa_kolumny, obiekt_reguly), ...)"""
        self._reguly = reguly

    def zastosuj(self, widok: WidokArkusza) -> dict[str, object]:
        zakresy: list[str] = []
        for kolumna, regula in self._reguly:
            zakresy.append(widok.regula(kolumna, regula))
        return {"strategia": self.nazwa, "zakresy_regul": tuple(zakresy)}
```

```python
# ======================================================================
# OS 2: DOKAD TRAFIA - StrategiaEksportu
# ======================================================================
class WynikEksportu25:
    """DTO wyniku. Zero openpyxl (to samo co WynikEksportu z modulu 24)."""

    def __init__(self, cel: object, bajtow: int = 0, kanał: str = "") -> None:
        self.cel = cel
        self.bajtow = bajtow
        self.kanał = kanał

    def __repr__(self) -> str:
        return f"<WynikEksportu25 {self.kanał} bajtow={self.bajtow}>"


class StrategiaEksportu(ABC):
    @property
    @abstractmethod
    def kanał(self) -> str: ...

    @abstractmethod
    def wyślij(self, raport: RaportExcel, drzewo: GrupaSekcji, cel) -> WynikEksportu: ...


class EksportXlsx(StrategiaEksportu):
    """Plik na dysku - zapis atomowy z modulu 24."""

    @property
    def kanał(self) -> str:
        return "xlsx"

    def wyślij(self, raport, drzewo, cel) -> WynikEksportu:
        sciezka = raport.zapisz(cel)
        return WynikEksportu(cel=sciezka, bajtow=len(raport.do_bajtow()), kanał=self.kanał)


class EksportBajtow(StrategiaEksportu):
    """Do pamieci - dla HTTP (StreamingResponse) albo kolejki."""

    @property
    def kanał(self) -> str:
        return "bytes"

    def wyślij(self, raport, drzewo, cel) -> WynikEksportu:
        dane = raport.do_bajtow()
        return WynikEksportu(cel=dane, bajtow=len(dane), kanał=self.kanał)


class EksportMarkdown(StrategiaEksportu):
    """Podglad tekstowy - ten sam model, inne medium (Composite z modulu 24)."""

    @property
    def kanał(self) -> str:
        return "markdown"

    def __init__(self, max_wierszy_tabeli: int = 20) -> None:
        self._max = max_wierszy_tabeli

    def wyślij(self, raport, drzewo, cel) -> WynikEksportu:
        cel_md = CelMarkdown(max_wierszy_tabeli=self._max)   # sprawdz w swojej wersji
        drzewo.renderuj(cel_md)
        tekst = cel_md.tekst()
        sciezka = Path(cel).with_suffix(".md")
        sciezka.parent.mkdir(parents=True, exist_ok=True)
        sciezka.write_text(tekst, encoding="utf-8")
        return WynikEksportu(cel=sciezka, bajtow=len(tekst.encode("utf-8")),
                             kanał=self.kanał)


# ======================================================================
# REJESTR STRATEGII - jedno miejsce decyzji (Registry z modulu 23)
# ======================================================================
STRATEGIE_FORMATOWANIA: dict[str, StrategiaFormatowania] = {
    "zebrowe": Zebrowe(),
    "finansowe": Finansowe("Cena", ("Ilosc",)),
}

STRATEGIE_EKSPORTU: dict[str, StrategiaEksportu] = {
    "xlsx": EksportXlsx(),
    "bytes": EksportBajtow(),
    "markdown": EksportMarkdown(),
}

# UWAGA na wspoldzielenie: strategie w rejestrze sa BEZSTANOWE albo
# skonfigurowane RAZ. Jezeli strategia ma stan zmienny, tworz ja per uzycie.
```

```python
# ======================================================================
# KONTEKST: raport, ktory UZYWA strategii, ale jej nie zna
# ======================================================================
class RaportKatalogowy:
    """Kontekst. Zna TYLKO nazwy strategii, nie ich klasy."""

    def __init__(
        self,
        tytul: str,
        arkusz: str,
        *,
        format_nazwa: str = "zebrowe",
        eksport_nazwa: str = "xlsx",
    ) -> None:
        self._raport = RaportExcelRozszerzony(tytul, autor="dzial BI")
        self._arkusz = arkusz
        self._format = STRATEGIE_FORMATOWANIA[format_nazwa]      # ← JEDEN wybor
        self._eksport = STRATEGIE_EKSPORTU[eksport_nazwa]       # ← JEDEN wybor
        self.metryki: dict[str, object] = {}

    def zbuduj(self, naglowki, wiersze, formaty=None) -> "RaportKatalogowy":
        self._raport.arkusz(self._arkusz, tytul=self._arkusz.capitalize(), siatka=False)
        self._raport.naglowek(self._arkusz, naglowki)
        self._raport.wiersze(self._arkusz, wiersze, formaty=formaty)
        return self

    def sformatuj(self) -> "RaportKatalogowy":
        widok = WidokFasady(self._raport, self._arkusz)
        self.metryki.update(self._format.zastosuj(widok))        # ← delegacja
        return self

    def wyslij(self, drzewo: GrupaSekcji, cel) -> WynikEksportu:
        return self._eksport.wyślij(self._raport, drzewo, cel)   # ← delegacja
```

```python
def main() -> None:
    from examples_24_03_dekoratory import DZIENNIK, METRYKI

    dane = [
        ("Alfa sp. z o.o.", "PL", 10, 25.0),
        ("Beta S.A.",       "DE",  4, 41.5),
        ("Gamma sp. k.",    "PL", 22, 25.0),
        ("Delta sp. z o.o.", "CZ", 7, 63.0),
    ]

    print("=== 5 typow raportow x 3 kanaly = 15 kombinacji, 8 klas ===")
    for tytul, format_nazwa in (
        ("Sprzedaz (zebrowe)", "zebrowe"),
        ("Sprzedaz (finansowe)", "finansowe"),
    ):
        for kanał in ("xlsx", "bytes", "markdown"):
            kontekst = RaportKatalogowy(
                tytul, "dane", format_nazwa=format_nazwa, eksport_nazwa=kanał
            )
            kontekst.zbuduj(
                ["Klient", "Region", "Ilosc", "Cena"], dane,
                formaty=[None, None, "#,##0", '#,##0.00 "zl"'],
            ).sformatuj()

            # to samo drzewo sekcji z modulu 24 - dla markdown i dla reszty
            from examples_24_05_composite import zbuduj_raport

            drzewo = zbuduj_raport(dane)
            wynik = kontekst.wyslij(drzewo, "output/25_01_raport")
            print(f"{format_nazwa:11} x {kanał:9} -> {wynik!r} | {kontekst.metryki}")

    print("metryki czasu:", METRYKI[-1] if METRYKI else None)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** W rejestrze `STRATEGIE_FORMATOWANIA` leżą **trzy instancje strategii** (dwie w słowniku plus `Warunkowe` tworzone w przykładzie 5), utworzone **raz**. To jest Flyweight z modułu 23: strategie są **bezstanowe** (albo skonfigurowane raz), więc można je współdzielić bezpiecznie. Kontekst `RaportKatalogowy` trzyma dwie **referencje** — nie kopie. Zysk pamięciowy jest niewielki przy dwóch instancjach, ale **zysk organizacyjny jest ogromny**: strategia jest **jedna**, więc zmiana w `Finansowe` działa we wszystkich raportach, natychmiast, bez szukania po plikach.

**Uwaga o współdzieleniu, którą trzeba powiedzieć wprost:** jeżeli strategia miałaby **stan zmienny** (np. licznik wywołań albo bufor), to współdzielenie w rejestrze byłoby **błędem** — dwa raporty w jednym procesie mieszałyby stan. Dlatego `EksportMarkdown` ma **parametr konstruktora** (`max_wierszy_tabeli`), a nie pole zmieniane w trakcie. Reguła: **strategia albo jest bezstanowa, albo jej stan jest ustalony przy tworzeniu.** Jeżeli potrzebujesz stanu per raport — twórz strategię per raport, nie w rejestrze.

**Co trafi do pliku.** Dla `format_nazwa="finansowe"`, `eksport_nazwa="xlsx"` — arkusz `Dane` z czterema wierszami, gdzie kolumna `Cena` ma format `#,##0.00 "zl";(#,##0.00)` (z nawiasem dla wartości ujemnych!), kolumna `Ilosc` format `#,##0`, obie wyrównane do prawej, a nagłówek `Cena` w kolorze `1F4E78`. Dla `zebrowe` — naprzemienne tło `F2F2F5` na wierszach. Dla `bytes` — plik **nie powstaje**, liczy się `len(raport.do_bajtow())`. Dla `markdown` — `output/25_01_raport.md`.

**Trzy rzeczy do przemyślenia:**

1. **Kombinacji było 15, klas jest 8.** Pięć typów raportów to w naszym przypadku **nie pięć klas**, a pięć wywołań z różnymi `format_nazwa` — bo format jest strategią, a nie typem raportu. To pokazuje, jak Strategy **rozbija jeden wymiar na dwa** i redukuje liczbę klas z iloczynu do sumy.
2. **Metryki wracają ze strategii.** `{"strategia": "zebrowe", "pomalowane_wiersze": 4}`. Dlaczego to jest w kontrakcie, a nie opcjonalne? Bo bez tego strategia jest **niema** — a niema strategia łamie zasadę z modułu 24: **błąd i efekt muszą być widoczne**. Metryka `pomalowane_wiersze: 0` natychmiast mówi, że coś nie zadziałało (np. arkusz nie miał danych). Bez metryki zobaczysz raport bez pasków i będziesz się zastanawiał dlaczego.
3. **Kontekst nie zna klas strategii.** `RaportKatalogowy` przyjmuje **nazwy**, nie obiekty. To znaczy, że wstawienie własnej strategii w aplikacji to jedna linia: `STRATEGIE_FORMATOWANIA["moja"] = MojaStrategia()`. **I to jest otwarte-zamknięte** z modułu 22, wreszcie naprawdę: nowe zachowanie bez modyfikacji istniejącego kodu.

### Przykład 2 — Template Method: `BazowyRaport` i dwa raporty konkretne

**Problem.** Trzy raporty mają identyczną kolejność kroków (przygotuj, nagłówek, dane, podsumowanie, formatowanie, wykresy, zapis) i różnią się treścią. W wersji bez wzorca: kopiuj-wklej szkieletu trzy razy, a poprawka kolejności w trzech miejscach.

**Rozwiązanie.** `BazowyRaport` z szkieletem `generuj()` i haczykami.

```python
"""Modul 25, przyklad 2: Template Method - BazowyRaport."""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import final

from examples_24_01_fasada import BladRaportu, RaportExcel
from examples_25_fasada import RaportExcelRozszerzony
from examples_25_strategie import STRATEGIE_EKSPORTU, STRATEGIE_FORMATOWANIA, WidokFasady

log = logging.getLogger("raporty")


class BazowyRaport(ABC):
    """SZKIELET raportu. Podklasa wypelnia TYLKO haczyki.

    Zasada Hollywood: podklasa NIE wola generuj(). Podklasa DOSTARCZA tresc.
    """

    #: nazwa arkusza w skoroszycie
    arkusz: str = "raport"

    def __init__(self, tytul: str, *, format_nazwa: str = "zebrowe",
                 cel: str | Path = "output/raport.xlsx") -> None:
        self._tytul = tytul
        self._cel = Path(cel)
        self._raport = RaportExcelRozszerzony(tytul, autor="dzial BI")
        self._format = STRATEGIE_FORMATOWANIA[format_nazwa]
        self.metryki: dict[str, object] = {}

    # ==================================================================
    # SZKIELET - @final: mypy tego pilnuje, podklasa NIE nadpisuje
    # ==================================================================
    @final
    def generuj(self) -> Path:
        """Szkielet. Kolejnosc jest NARZUCONA i nie podlega negocjacji."""
        self._waliduj()                       # krok staly
        self._przygotuj()                     # krok staly

        naglowki = self._naglowek()           # HACZYK WYMAGANY
        if not naglowki:
            raise BladRaportu(f"{type(self).__name__}._naglowek() zwrocil puste")

        wiersze = self._wiersze()             # HACZYK WYMAGANY
        ile = self._raport.wiersze(
            self.arkusz, wiersze, formaty=self._formaty()
        )
        self.metryki["wierszy"] = ile

        self._podsumowanie()                  # HACZYK OPCJONALNY
        self._formatowanie()                  # HACZYK OPCJONALNY
        self._wykresy()                       # HACZYK OPCJONALNY
        self._druk()                          # HACZYK OPCJONALNY

        return self._zapisz()                 # krok staly

    # ==================================================================
    # KROKI STALE
    # ==================================================================
    def _waliduj(self) -> None:
        """Sprawdza niezmienniki raportu. Rzuca, nie loguje i nie idzie dalej."""
        if not self._tytul.strip():
            raise BladRaportu("Raport bez tytulu - nie generujemy")

    def _przygotuj(self) -> None:
        self._raport.arkusz(self.arkusz, tytul=self.arkusz.capitalize(),
                            siatka=False, freeze_wiersz=1)
        self._raport.naglowek(self.arkusz, self._naglowek())
        self._raport.szerokosci(self.arkusz, self._szerokosci())

    def _zapisz(self) -> Path:
        cel = self._raport.zapisz(self._cel)
        log.info("raport %s -> %s (%s wierszy)", type(self).__name__, cel,
                 self.metryki.get("wierszy"))
        return cel

    # ==================================================================
    # HACZYKI WYMAGANE - ABC wymusi ich implementacje
    # ==================================================================
    @abstractmethod
    def _naglowek(self) -> list[str]:
        """Etykiety kolumn."""

    @abstractmethod
    def _wiersze(self) -> Iterable[Sequence[object]]:
        """Wiersze danych."""

    # ==================================================================
    # HACZYKI OPCJONALNE - domyslnie nic nie robia
    # ==================================================================
    def _formaty(self) -> Sequence[str | None] | None:
        return None

    def _szerokosci(self) -> Sequence[float]:
        return [18.0] * len(self._naglowek())

    def _podsumowanie(self) -> None:
        """Domyslnie: BRAK podsumowania. Swiadoma decyzja, nie zapomnienie."""
        return None

    def _formatowanie(self) -> None:
        widok = WidokFasady(self._raport, self.arkusz)
        self.metryki.update(self._format.zastosuj(widok))

    def _wykresy(self) -> None:
        return None

    def _druk(self) -> None:
        ostatnia = len(self._naglowek())
        from openpyxl.utils import get_column_letter

        self._raport.druk_poziomo(
            self.arkusz, f"A1:{get_column_letter(ostatnia)}{self.metryki.get('wierszy', 1) + 1}"
        )
```

```python
# ======================================================================
# RAPORT KONKRETNY 1 - sprzedaz
# ======================================================================
class RaportSprzedazy(BazowyRaport):
    arkusz = "sprzedaz"

    def __init__(self, dane, **kw) -> None:
        super().__init__("Raport sprzedazy", **kw)
        self._dane = tuple(dane)

    def _naglowek(self) -> list[str]:
        return ["Klient", "Region", "Ilosc", "Cena"]

    def _wiersze(self):
        return self._dane

    def _formaty(self):
        return [None, None, "#,##0", '#,##0.00 "zl"']

    def _szerokosci(self):
        return [24.0, 8.0, 8.0, 12.0]

    def _podsumowanie(self) -> None:
        """HACZYK NADPISANY - bez tego raport nie mialby sumy."""
        suma = sum(w[2] * w[3] for w in self._dane)
        self._raport.para(self.arkusz, "Pozycji", len(self._dane), "#,##0")
        self._raport.para(self.arkusz, "Wartosc lacznie", round(suma, 2),
                          '#,##0.00 "zl"')
        self.metryki["suma"] = round(suma, 2)

    def _wykresy(self) -> None:
        self._raport.wykres_slupkowy(
            self.arkusz, "Sprzedaz wg klienta",
            kolumna_kategorii="Klient", kolumny_wartosci=["Cena"],
            pozycja=(2, 6),
        )
        self.metryki["wykresy"] = 1


# ======================================================================
# RAPORT KONKRETNY 2 - stan magazynu (BEZ wykresu, INNE podsumowanie)
# ======================================================================
class RaportStanuMagazynu(BazowyRaport):
    arkusz = "magazyn"

    def __init__(self, dane, prog: float = 10.0, **kw) -> None:
        super().__init__("Stan magazynu", **kw)
        self._dane = tuple(dane)
        self._prog = prog

    def _naglowek(self) -> list[str]:
        return ["Symbol", "Nazwa", "Stan", "Minimum", "Czy ponizej"]

    def _wiersze(self):
        for symbol, nazwa, stan, minimum in self._dane:
            yield (symbol, nazwa, stan, minimum,
                   "TAK" if stan < minimum else "NIE")

    def _formaty(self):
        return [None, None, "#,##0", "#,##0", None]

    def _szerokosci(self):
        return [12.0, 30.0, 10.0, 10.0, 12.0]

    def _podsumowanie(self) -> None:
        braki = [w for w in self._dane if w[2] < w[3]]
        self._raport.para(self.arkusz, "Pozycji ogolem", len(self._dane), "#,##0")
        self._raport.para(self.arkusz, "Ponizej minimum", len(braki), "#,##0")
        self.metryki["braki"] = len(braki)

    # _wykresy() NIE nadpisane -> domyslna pusta implementacja.
    # To JEST swiadome: raport magazynowy nie ma wykresu.


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    sprzedaz = RaportSprzedazy(
        [("Alfa sp. z o.o.", "PL", 10, 25.0),
         ("Beta S.A.", "DE", 4, 41.5),
         ("Gamma sp. k.", "PL", 22, 25.0)],
        format_nazwa="finansowe", cel="output/25_02_sprzedaz.xlsx",
    )
    magazyn = RaportStanuMagazynu(
        [("A-100", "Kabel 3x1.5", 120, 50),
         ("B-200", "Zlaczka", 8, 20),
         ("C-300", "Puszka", 4, 30)],
        format_nazwa="zebrowe", cel="output/25_02_magazyn.xlsx",
    )

    for raport in (sprzedaz, magazyn):
        cel = raport.generuj()
        print(f"{type(raport).__name__:22} -> {cel} | {raport.metryki}")


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `RaportSprzedazy` ma trzy haczyki wymagane, cztery opcjonalne nadpisane i **jedną** metodę, której nie wolno nadpisać (`generuj`). Podklasa **nie zawiera** ani jednej linii obsługi openpyxl — całe `ws` jest poza nią, w fasadzie. To jest skumulowany efekt modułów 24 i 25: **podklasa jest naprawdę cienka**, a jej kod to opis treści raportu, nie mechaniki.

Zwróć uwagę na `RaportStanuMagazynu._wiersze()` — **generator**. Prezentacja „Czy poniżej" (`TAK`/`NIE`) jest wyliczana w locie, a nie trzymana w pamięci. To moduł 19 w praktyce: tam, gdzie da się streamować, streamuj.

**Co trafi do pliku.** Dwa pliki. W `25_02_sprzedaz.xlsx` — arkusz `Sprzedaz` z zamrożonym pierwszym wierszem, czterema kolumnami, formatem walutowym z nawiasami dla ujemnych, dwoma wierszami KPI **wstawionymi pod danymi** (bo `para` dopisuje na końcu) i wykresem słupkowym od `F2`. W `25_02_magazyn.xlsx` — arkusz `Magazyn` z pięcioma kolumnami, naprzemiennym tłem i podsumowaniem mającym **dwie** pary KPI; wykresu **nie ma**, bo haczyk nie został nadpisany — i to jest w porządku.

**Trzy rzeczy do przemyślenia:**

1. **Kolejność kroków jest egzekwowana jednym miejscem.** Zmiana w `generuj()` (np. dodanie `_wykresy()` **po** `_podsumowanie()`) działa dla **wszystkich** raportów naraz. W wersji kopiuj-wklej zmieniłbyś to w trzech plikach — i zapomniałbyś w czwartym.
2. **`_podsumowanie` w `RaportStanuMagazynu` liczy braki samodzielnie.** Zwróć uwagę, że to jest logika biznesowa („stan poniżej minimum") i leży w podklasie. **To jest granica, o której mówi moduł 26:** w wersji produkcyjnej taką logikę wyciągnąłbyś do domeny (`def braki(dane)`), a podklasa tylko by ją wołała. Tu zostawiam ją w podklasie, bo to moduł o zachowaniu, nie o architekturze — ale zaznaczam, że to jest **dług do spłacenia**.
3. **`_formatowanie()` jest haczykiem opcjonalnym, ale ma domyślną implementację, która coś robi.** To ciekawy przypadek: domyślnie **stosuje strategię formatowania**, więc podklasa nie musi o tym pamiętać. To jest właściwy sposób użycia haczyka opcjonalnego: nie „pusty z natury", ale **„sensowny z natury"**. Jeżeli haczyk ma sensowną domyślną wartość — daj ją. Jeżeli nie ma — rozważ, czy to haczyk wymagany (pułapka 2).

### Przykład 3 — Command: cofanie, audyt, makro, transakcja

**Problem.** Raport trzeba poprawić po wygenerowaniu: ktoś zmienia trzy liczby, dodaje wiersz nagłówkowy sekcji, formatuje kolumnę. Jeżeli coś pójdzie źle, chcemy cofnąć. I chcemy wiedzieć, **kto** to zrobił.

**Rozwiązanie.** Komendy na stosie.

```python
"""Modul 25, przyklad 3: Command - cofanie, audyt, makro, transakcja."""

from __future__ import annotations

from abc import ABC, abstractmethod
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from examples_24_01_fasada import BladRaportu
from examples_25_fasada import RaportExcelRozszerzony
from openpyxl.utils import get_column_letter


# ======================================================================
# KONTEKST - jedyny obiekt, ktory komendy maja prawo dotknac
# ======================================================================
class Kontekst:
    """Owija fasade i daje komendom MINIMALNY interfejs.

    Komendy NIE dostaja `ws` ani nie znaja openpyxl. Dostaja operacje
    w slownictwie raportu: komorka(arkusz, wiersz, kolumna_nazwa).
    """

    def __init__(self, raport: RaportExcelRozszerzony, arkusz: str) -> None:
        self._raport = raport
        self._arkusz = arkusz

    @property
    def arkusz(self) -> str:
        return self._arkusz

    def odczytaj(self, wiersz: int, kolumna: str) -> object:
        """Odczyt wartosci - do zapamietania poprzedniej wartosci."""
        ws = self._raport._ws(self._arkusz)          # wewnatrz infrastruktury!
        indeks = self._raport.kolumny(self._arkusz).index(kolumna) + 1
        return ws.cell(row=wiersz, column=indeks).value

    def ustaw(self, wiersz: int, kolumna: str, wartosc: Any) -> None:
        ws = self._raport._ws(self._arkusz)
        indeks = self._raport.kolumny(self._arkusz).index(kolumna) + 1
        ws.cell(row=wiersz, column=indeks, value=wartosc)

    def usun(self, wiersz: int, kolumna: str) -> None:
        self.ustaw(wiersz, kolumna, None)

    def wstaw_wiersze(self, przed: int, ile: int) -> None:
        self._raport._ws(self._arkusz).insert_rows(przed, amount=ile)

    def usun_wiersze(self, od: int, ile: int) -> None:
        self._raport._ws(self._arkusz).delete_rows(od, amount=ile)

    def styl(self, wiersz: int, kolumna: str) -> dict:
        """Migawka stylu - style sa NIEMUTOWALNE, wiec referencje sa bezpieczne."""
        ws = self._raport._ws(self._arkusz)
        indeks = self._raport.kolumny(self._arkusz).index(kolumna) + 1
        cell = ws.cell(row=wiersz, column=indeks)
        return {"font": cell.font, "fill": cell.fill,
                "number_format": cell.number_format,
                "alignment": cell.alignment}

    def zastosuj_styl(self, wiersz: int, kolumna: str, migawka: dict) -> None:
        ws = self._raport._ws(self._arkusz)
        indeks = self._raport.kolumny(self._arkusz).index(kolumna) + 1
        cell = ws.cell(row=wiersz, column=indeks)
        cell.font = migawka["font"]              # ← referencja, nie kopia!
        cell.fill = migawka["fill"]
        cell.number_format = migawka["number_format"]
        cell.alignment = migawka["alignment"]


# ======================================================================
# KOMENDA - obiekt-zadanie z odwrotnoscia
# ======================================================================
class Komenda(ABC):
    """UWAGA: komenda JEST MUTOWALNA - musi pamietac swoja odwrotnosc.

    To jedyny taki obiekt w tym kursie. Stan zmienia TYLKO wykonaj().
    """

    @property
    @abstractmethod
    def opis(self) -> str: ...

    @abstractmethod
    def wykonaj(self, kontekst: Kontekst) -> None: ...

    @abstractmethod
    def cofnij(self, kontekst: Kontekst) -> None: ...


@dataclass
class KomendaUstawKomorke(Komenda):
    """Ustawia wartosc. Pamięta, CO BYLO, zeby cofnac."""

    arkusz: str
    wiersz: int
    kolumna: str
    wartosc: Any
    autor: str = "system"
    _poprzednia: object = field(default=None, init=False, repr=False)
    _bylo: bool = field(default=False, init=False, repr=False)

    @property
    def opis(self) -> str:
        return f"{self.arkusz}!{self.kolumna}{self.wiersz} = {self.wartosc!r}"

    def wykonaj(self, kontekst: Kontekst) -> None:
        migawka = kontekst.odczytaj(self.wiersz, self.kolumna)
        # rozrozniamy "bylo None" od "byla wartosc None" - to jest sedno
        self._bylo = migawka is not None
        self._poprzednia = migawka
        kontekst.ustaw(self.wiersz, self.kolumna, self.wartosc)

    def cofnij(self, kontekst: Kontekst) -> None:
        if self._bylo:
            kontekst.ustaw(self.wiersz, self.kolumna, self._poprzednia)
        else:
            kontekst.usun(self.wiersz, self.kolumna)      # byla PUSTA komorka


@dataclass
class KomendaDodajWiersz(Komenda):
    """Wstawia puste wiersze przed danym wierszem.

    Uczciwe ostrzezenie (modul 07): openpyxl NIE przesuwa referencji formuł.
    Po wstawieniu wierszy formuly w pozostalych komorkach moga wskazywac
    w zle miejsce. Cofanie przywraca LICZBE wierszy, ale nie naprawi
    formul, ktore Excel przeliczy sam.
    """

    arkusz: str
    przed: int
    ile: int = 1
    autor: str = "system"

    @property
    def opis(self) -> str:
        return f"INSERT {self.ile} wierszy przed {self.arkusz}!{self.przed}"

    def wykonaj(self, kontekst: Kontekst) -> None:
        kontekst.wstaw_wiersze(self.przed, self.ile)

    def cofnij(self, kontekst: Kontekst) -> None:
        kontekst.usun_wiersze(self.przed, self.ile)


@dataclass
class KomendaSformatuj(Komenda):
    """Nadaje format liczbowy. Cofanie = przywrocenie migawki stylu."""

    arkusz: str
    wiersz: int
    kolumna: str
    number_format: str
    autor: str = "system"
    _migawka: dict | None = field(default=None, init=False, repr=False)

    @property
    def opis(self) -> str:
        return (f"FORMAT {self.arkusz}!{self.kolumna}{self.wiersz} "
                f"-> {self.number_format!r}")

    def wykonaj(self, kontekst: Kontekst) -> None:
        self._migawka = kontekst.styl(self.wiersz, self.kolumna)
        migawka = dict(self._migawka)
        migawka["number_format"] = self.number_format
        kontekst.zastosuj_styl(self.wiersz, self.kolumna, migawka)

    def cofnij(self, kontekst: Kontekst) -> None:
        if self._migawka is not None:
            kontekst.zastosuj_styl(self.wiersz, self.kolumna, self._migawka)
```

```python
# ======================================================================
# MAKRO - Composite komend (modul 24!)
# ======================================================================
@dataclass
class MakroKomend(Komenda):
    """Grupa komend, ktora sama jest komenda. Cofa sie w ODWROTNEJ kolejnosci."""

    nazwa: str
    komendy: list[Komenda] = field(default_factory=list)
    autor: str = "system"

    @property
    def opis(self) -> str:
        return f"MAKRO {self.nazwa!r} ({len(self.komendy)} komend)"

    def dodaj(self, komenda: Komenda) -> "MakroKomend":
        if not isinstance(komenda, MakroKomend):
            self.komendy.append(komenda)
        return self

    def wykonaj(self, kontekst: Kontekst) -> None:
        for komenda in self.komendy:
            komenda.wykonaj(kontekst)

    def cofnij(self, kontekst: Kontekst) -> None:
        for komenda in reversed(self.komendy):       # ← ODWROTNA kolejnosc!
            komenda.cofnij(kontekst)


# ======================================================================
# STOS KOMEND - cofanie, ponawianie, audyt
# ======================================================================
class StosKomend:
    def __init__(self) -> None:
        self._wykonane: list[Komenda] = []
        self._cofniete: list[Komenda] = []
        self._audyt: list[dict] = []

    def wykonaj(self, komenda: Komenda, kontekst: Kontekst) -> None:
        komenda.wykonaj(kontekst)
        self._wykonane.append(komenda)
        self._cofniete.clear()          # nowa zmiana kasuje historie "do przodu"
        self._audyt.append({
            "czas": datetime.now().isoformat(timespec="seconds"),
            "operacja": "WYKONAJ",
            "opis": komenda.opis,
            "autor": getattr(komenda, "autor", "system"),
        })

    def cofnij(self, kontekst: Kontekst, ile: int = 1) -> list[str]:
        cofniete: list[str] = []
        for _ in range(min(ile, len(self._wykonane))):
            komenda = self._wykonane.pop()
            komenda.cofnij(kontekst)
            self._cofniete.append(komenda)
            cofniete.append(komenda.opis)
            self._audyt.append({
                "czas": datetime.now().isoformat(timespec="seconds"),
                "operacja": "COFNIJ",
                "opis": komenda.opis,
                "autor": getattr(komenda, "autor", "system"),
            })
        return cofniete

    def ponow(self, kontekst: Kontekst, ile: int = 1) -> list[str]:
        ponowione: list[str] = []
        for _ in range(min(ile, len(self._cofniete))):
            komenda = self._cofniete.pop()
            komenda.wykonaj(kontekst)
            self._wykonane.append(komenda)
            ponowione.append(komenda.opis)
        return ponowione

    def znacznik(self) -> int:
        """Punkt kontrolny transakcji. Do niego mozna wrocic."""
        return len(self._wykonane)

    def do_znacznika(self, kontekst: Kontekst, znacznik: int) -> list[str]:
        """Cofa wszystko do punktu kontrolnego."""
        return self.cofnij(kontekst, ile=len(self._wykonane) - znacznik)

    def audyt(self) -> tuple[dict, ...]:
        return tuple(self._audyt)

    def stan(self) -> str:
        return f"wykonane={len(self._wykonane)} cofniete={len(self._cofniete)}"
```

```python
def main() -> None:
    from examples_24_05_composite import zbuduj_raport
    from examples_25_templatemethod import RaportSprzedazy

    dane = [("Alfa sp. z o.o.", "PL", 10, 25.0),
            ("Beta S.A.", "DE", 4, 41.5),
            ("Gamma sp. k.", "PL", 22, 25.0)]

    raport = RaportSprzedazy(dane, format_nazwa="finansowe",
                             cel="output/25_03_baza.xlsx")
    raport.generuj()

    kontekst = Kontekst(raport._raport, raport.arkusz)
    stos = StosKomend()

    print("--- 1. pojedyncze korekty ---")
    stos.wykonaj(KomendaUstawKomorke(
        "sprzedaz", 2, "Cena", 30.0, autor="anna.kowalska"), kontekst)
    stos.wykonaj(KomendaSformatuj(
        "sprzedaz", 2, "Ilosc", "#,##0.00", autor="anna.kowalska"), kontekst)
    print("Cena w wierszu 2:", kontekst.odczytaj(2, "Cena"))
    print("Stan stosu:", stos.stan())

    print("--- 2. makro: naglowek sekcji + formaty ---")
    znacznik = stos.znacznik()
    makro = MakroKomend("naglowek raportu", autor="piotr.nowak")
    makro.dodaj(KomendaDodajWiersz("sprzedaz", 1, ile=1, autor="piotr.nowak"))
    makro.dodaj(KomendaUstawKomorke("sprzedaz", 1, "Klient",
                                    "SPRZEDAZ - STYCZEN 2026", autor="piotr.nowak"))
    for kolumna in ("Klient", "Region", "Ilosc", "Cena"):
        makro.dodaj(KomendaSformatuj("sprzedaz", 1, kolumna, "General",
                                     autor="piotr.nowak"))
    stos.wykonaj(makro, kontekst)
    print("Stan stosu:", stos.stan())

    print("--- 3. cofniecie CALEGO makra (rollback do znacznika) ---")
    cofniete = stos.do_znacznika(kontekst, znacznik)
    print("Cofnieto:", cofniete)
    print("Cena w wierszu 2 po rollbacku:", kontekst.odczytaj(2, "Cena"))

    print("--- 4. ponowienie korekty ---")
    stos.ponow(kontekst, ile=1)
    print("Cena w wierszu 2:", kontekst.odczytaj(2, "Cena"))

    print("--- 5. audyt ---")
    for wpis in stos.audyt():
        print(f"  {wpis['czas']} {wpis['operacja']:8} {wpis['autor']:15} {wpis['opis']}")

    print("--- 6. COMMIT: dopiero teraz plik zmienia sie na dysku ---")
    cel = raport._raport.zapisz("output/25_03_po_komendach.xlsx")
    print("Zapisano:", cel)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Trzy rzeczy warte nazwania:

1. **`KomendaUstawKomorke` rozróżnia „było `None`" od „była wartość".** Pole `_bylo` istnieje po to, żeby cofanie **usunęło** komórkę, jeśli przedtem była pusta, a **przywróciło wartość**, jeśli nie była. Gdybyśmy zapamiętali tylko `_poprzednia = None`, to cofnięcie komórki, która wcześniej zawierała `0` albo `""`, wpisałoby `None`. To jest dokładnie ta różnica z modułu 05 („puste krzesło na sali"), przeniesiona na poziom mechaniki cofania.

2. **`KomendaSformatuj` zapamiętuje migawkę z **referencjami** do obiektów stylu, a nie ich kopiami.** I to jest **bezpieczne**, bo obiekty stylów są **niemutowalne** (moduł 09). Gdyby były mutowalne, ta migawka byłaby bezwartościowa — bo zmiana oryginału zmieniłaby też zapamiętaną migawkę. **Niemutowalność stylów, która była źródłem frustracji w module 09, w module 25 okazuje się błogosławieństwem.**

3. **`MakroKomend.cofnij()` idzie w odwrotnej kolejności.** To nie kosmetyka — to **warunek poprawności**. Jeżeli makro najpierw wstawia wiersz, a potem wpisuje wartości do tego wiersza, to cofanie musi najpierw usunąć wartości, a potem wiersz. Odwrotna kolejność (usunięcie wiersza przed przywróceniem wartości) zostawiłaby wartości w **innym** wierszu — bo numery wierszy się przesunęły.

**Co trafi do pliku.** `output/25_03_baza.xlsx` (stan po `generuj()`) i `output/25_03_po_komendach.xlsx` (stan po komendach — z ceną 30,0 w wierszu 2, formatem `#,##0.00` w kolumnie `Ilosc`, oraz makrem cofniętym, czyli **bez** dodanego wiersza nagłówka). To jest konkretna, sprawdzalna różnica — otwórz oba pliki obok siebie.

**Trzy rzeczy do przemyślenia:**

1. **`Kontekst` sięga do `raport._ws(...)`.** Zwróć uwagę: **tak, robi to** — i jest to **jedyne** miejsce w tym module, gdzie prywatny akcesor fasady jest użyty, i to z uzasadnieniem: `Kontekst` **jest** częścią infrastruktury (leży w warstwie obsługi openpyxl), a nie warstwą aplikacji. Komendy widzą już tylko `Kontekst` z metodami `odczytaj`/`ustaw`, więc **na poziomie komend openpyxl nie ma**. Gdyby komendy dostawały `ws` bezpośrednio, mielibyśmy przeciek (pułapka 1 z modułu 24). **Jeżeli refaktoryzowałbyś to dalej, `Kontekst` dostałby własne metody w fasadzie** i `_ws` zniknęłoby również stąd.
2. **Stos komend i zapis to dwie różne rzeczy.** `stos.cofnij()` zmienia **pamięć**. `raport._raport.zapisz()` zmienia **dysk**. W naszym `main()` komendy są wykonane, makro cofnięte, korekta ponowiona — i **dopiero** na końcu zapis. Gdybyśmy zapisali wcześniej, cofanie makra nie zmieniłoby pliku bez **ponownego** zapisu. To jest granica transakcji z sekcji 2.3 i ona **musi być** w komentarzu/docstringu, bo jest nieintuicyjna.
3. **Audyt odnotowuje `autor` z komendy.** Komenda ma pole `autor: str` z domyślnym `"system"`. W prawdziwym systemie byłby to obiekt użytkownika, a nie string — ale **string wystarcza, żeby pytanie „kto zmienił tę liczbę" miało odpowiedź**. I to jest cały sens audytu: nie musi być piękny, musi być **kompletny i niepodrabialny w praktyce** (bo każda zmiana przechodzi przez `StosKomend`).

### Przykład 4 — Visitor: czterej goście i dwie fazy

**Problem.** Chcemy cztery analizy tego samego skoroszytu: znaleźć puste kolumny, wykryć formuły, zebrać metryki i zanonimizować dane osobowe. Trzy pierwsze tylko czytają; czwarta **zmienia**. Bez rozdzielenia faz czwarta popsuje iterację.

**Rozwiązanie.** Bazowy gość z szablonowym przejściem + czterej goście konkretni + dwufazowość.

```python
"""Modul 25, przyklad 4: Visitor - przemiatanie skoroszytu."""

from __future__ import annotations

import hashlib
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import final

from openpyxl.utils import get_column_letter

from examples_25_command import (
    Komenda, KomendaSformatuj, KomendaUstawKomorke, Kontekst, StosKomend,
)


@dataclass(frozen=True)
class Znalezisko:
    """Rekord znaleziska. Niemutowalny - bo to wynik, nie stan."""

    typ: str
    arkusz: str
    wiersz: int
    kolumna: str
    szczegol: str

    def __str__(self) -> str:
        return f"[{self.typ}] {self.arkusz}!{self.kolumna}{self.wiersz}: {self.szczegol}"


class OdwiedzSkoroszyt(ABC):
    """Bazowy gosc. Metody odwiedz_* ZBIERAJA - nie modyfikuja.

    UWAGA: to jest jednoczesnie Template Method - `przemieć()` to szkielet,
    a `odwiedz_*` to haczyki. Wzorce sie skladaja.
    """

    def __init__(self) -> None:
        self._znaleziska: list[Znalezisko] = []
        self._komendy: list[Komenda] = []      # faza 2 - do zastosowania pozniej
        self.metryki: dict[str, object] = {}

    # ------------------------------------------------------------------
    # HACZYKI - domyslnie nic nie robia, bo wiekszosc gosci nie potrzebuje
    # wszystkich trzech poziomow
    # ------------------------------------------------------------------
    def odwiedz_arkusz(self, nazwa: str, ws) -> None:
        return None

    def odwiedz_komorke(self, arkusz: str, wiersz: int, kolumna: str, komorka) -> None:
        return None

    def odwiedz_tabele(self, arkusz: str, tabela) -> None:
        return None

    # ------------------------------------------------------------------
    # SZKIELET PRZEJSCIA
    # ------------------------------------------------------------------
    @final
    def przemieć(self, wb) -> "OdwiedzSkoroszyt":
        self.poczatek(wb)
        for nazwa in wb.sheetnames:
            ws = wb[nazwa]
            self.odwiedz_arkusz(nazwa, ws)
            for wiersz in ws.iter_rows():
                for komorka in wiersz:
                    if komorka.value is not None:
                        self.odwiedz_komorke(
                            nazwa, komorka.row,
                            get_column_letter(komorka.column), komorka,
                        )
            for tabela in ws.tables.values():
                self.odwiedz_tabele(nazwa, tabela)
        self.koniec(wb)
        return self

    # haczyki cyklu zycia - opcjonalne
    def poczatek(self, wb) -> None:
        return None

    def koniec(self, wb) -> None:
        return None

    # ------------------------------------------------------------------
    # WYNIKI
    # ------------------------------------------------------------------
    def wynik(self) -> tuple[Znalezisko, ...]:
        return tuple(self._znaleziska)

    def komendy(self) -> tuple[Komenda, ...]:
        """Faza 2: komendy do wykonania PO obchodzie."""
        return tuple(self._komendy)

    def _dodaj(self, typ: str, arkusz: str, wiersz: int, kolumna: str,
               szczegol: str) -> None:
        self._znaleziska.append(Znalezisko(typ, arkusz, wiersz, kolumna, szczegol))
```

```python
# ======================================================================
# GOSC 1: SzukajPustychKolumn - tylko czyta
# ======================================================================
class SzukajPustychKolumn(OdwiedzSkoroszyt):
    """Kolumnа jest pusta, jesli w zakresie danych nie ma ani jednej wartosci.

    Uczciwe ostrzezenie: openpyxl nie ma pojecia "zakres danych". Bierzemy
    maximum z arkusza (ws.max_row, ws.max_column) i odnotowujemy ROZBIEZNOSC
    miedzy wymiarem deklarowanym a faktycznym (modul 07).
    """

    def __init__(self, wiersz_naglowka: int = 1) -> None:
        super().__init__()
        self._naglowek = wiersz_naglowka
        self._obecne: dict[str, set[int]] = {}

    def odwiedz_arkusz(self, nazwa: str, ws) -> None:
        self._obecne[nazwa] = set()

    def odwiedz_komorke(self, arkusz, wiersz, kolumna, komorka) -> None:
        if wiersz > self._naglowek:
            self._obecne[arkusz].add(komorka.column)

    def koniec(self, wb) -> None:
        for nazwa in wb.sheetnames:
            ws = wb[nazwa]
            obecne = self._obecne.get(nazwa, set())
            for indeks in range(1, (ws.max_column or 0) + 1):
                if indeks not in obecne:
                    self._dodaj("pusta_kolumna", nazwa, 0,
                                get_column_letter(indeks),
                                f"kolumna {indeks} bez danych ponizej naglowka")
        self.metryki["pustych_kolumn"] = len(self._znaleziska)


# ======================================================================
# GOSC 2: WykryjFormuly - tylko czyta
# ======================================================================
class WykryjFormuly(OdwiedzSkoroszyt):
    """Formuly sa komorkami o data_type == 'f' albo wartosci zaczynajacej sie od '='."""

    def __init__(self, przyklad_limit: int = 5) -> None:
        super().__init__()
        self._limit = przyklad_limit

    def odwiedz_komorke(self, arkusz, wiersz, kolumna, komorka) -> None:
        value = komorka.value
        jest_formula = (
            komorka.data_type == "f"
            or (isinstance(value, str) and value.startswith("="))
        )
        if jest_formula:
            self._dodaj("formula", arkusz, wiersz, kolumna, str(value)[:80])

    def koniec(self, wb) -> None:
        self.metryki["formul"] = len(self._znaleziska)
        self.metryki["przyklady"] = [str(z) for z in self._znaleziska[: self._limit]]


# ======================================================================
# GOSC 3: ZbierzMetryki - tylko czyta
# ======================================================================
class ZbierzMetryki(OdwiedzSkoroszyt):
    """Liczby o skoroszycie. Punkt wyjscia dla testu regresji (modul 20)."""

    def __init__(self) -> None:
        super().__init__()
        self._arkuszy = 0
        self._komorek = 0
        self._formul = 0
        self._tabel = 0
        self._rozbieznosci: list[str] = []

    def odwiedz_arkusz(self, nazwa: str, ws) -> None:
        self._arkuszy += 1
        deklarowany = (ws.max_row or 0, ws.max_column or 0)

    def odwiedz_komorke(self, arkusz, wiersz, kolumna, komorka) -> None:
        self._komorek += 1
        if komorka.data_type == "f":
            self._formul += 1

    def odwiedz_tabele(self, arkusz: str, tabela) -> None:
        self._tabel += 1
        self._dodaj("tabela", arkusz, 0, "-",
                    f"{tabela.displayName} @ {tabela.ref}")

    def koniec(self, wb) -> None:
        self.metryki.update({
            "arkuszy": self._arkuszy,
            "komorek_z_danymi": self._komorek,
            "formul": self._formul,
            "tabel": self._tabel,
            "arkusze": tuple(wb.sheetnames),
        })


# ======================================================================
# GOSC 4: Anonimizuj - ZMIENIA, ale NIE TERAZ. Produkuje komendy.
# ======================================================================
class Anonimizuj(OdwiedzSkoroszyt):
    """Gość, ktory ZMIENIA. Dlatego NIE zmienia tu i teraz - produkuje komendy.

    To jest dokladna realizacja zasady z sekcji 1.4: kontroler z checklista
    nic nie przestawia w trakcie obchodu.
    """

    def __init__(self, kolumny_wrazliwe: tuple[str, ...] = ("Klient",),
                 sol: str = "raport-2026") -> None:
        super().__init__()
        self._kolumny = kolumny_wrazliwe
        self._sol = sol
        self._mapa: dict[str, str] = {}

    def _pseudonim(self, wartosc: str) -> str:
        """Deterministyczny pseudonim - ta sama wartosc = ten sam pseudonim."""
        if wartosc not in self._mapa:
            skrot = hashlib.sha256((self._sol + wartosc).encode("utf-8"))
            self._mapa[wartosc] = "OS-" + skrot.hexdigest()[:8].upper()
        return self._mapa[wartosc]

    def odwiedz_arkusz(self, nazwa: str, ws) -> None:
        # zapamietujemy naglowki, zeby wiedziec, ktore kolumny sa wrazliwe
        naglowki = [c.value for c in ws[1] if c.value is not None]
        self._kolumny_arkusza = {
            get_column_letter(i + 1): str(v) for i, v in enumerate(naglowki)
        }

    def odwiedz_komorke(self, arkusz, wiersz, kolumna, komorka) -> None:
        if wiersz == 1:
            return
        nazwa_kolumny = getattr(self, "_kolumny_arkusza", {}).get(kolumna)
        if nazwa_kolumny not in self._kolumny:
            return
        wartosc = komorka.value
        if not isinstance(wartosc, str) or not wartosc.strip():
            return
        # FAZA 1: tylko ZBIERAMY komende, nic nie zmieniamy
        self._komendy.append(
            KomendaUstawKomorke(arkusz, wiersz, nazwa_kolumny,
                                self._pseudonim(wartosc), autor="anonimizacja")
        )
        self._dodaj("dane_osobowe", arkusz, wiersz, kolumna,
                    f"{nazwa_kolumny}: {wartosc!r} -> pseudonim")

    def koniec(self, wb) -> None:
        self.metryki["do_anonimizacji"] = len(self._komendy)
        self.metryki["unikalnych_wartosci"] = len(self._mapa)


# ======================================================================
# ZASTOSOWANIE GOŚCIA - faza 2. Komendy ida na stos.
# ======================================================================
def zastosuj_komendy_goscia(
    gosc: OdwiedzSkoroszyt, kontekst: Kontekst, *, autor: str = "system"
) -> tuple[list[str], StosKomend]:
    stos = StosKomend()
    opisy: list[str] = []
    for komenda in gosc.komendy():
        stos.wykonaj(komenda, kontekst)
        opisy.append(komenda.opis)
    return opisy, stos
```

```python
def main() -> None:
    from examples_25_templatemethod import RaportSprzedazy

    dane = [("Alfa sp. z o.o.", "PL", 10, 25.0),
            ("Beta S.A.", "DE", 4, 41.5),
            ("Alfa sp. z o.o.", "PL", 7, 63.0)]
    raport = RaportSprzedazy(dane, format_nazwa="finansowe",
                             cel="output/25_04_baza.xlsx")
    raport.generuj()

    wb = raport._raport._wb                 # w TESCIE / infrastrukturze to OK

    print("=== 1. Czytajace goscie ===")
    for gosc in (SzukajPustychKolumn(), WykryjFormuly(), ZbierzMetryki()):
        gosc.przemieć(wb)
        print(f"{type(gosc).__name__:22} metryki={gosc.metryki}")
        for z in gosc.wynik()[:3]:
            print("   ", z)

    print("=== 2. Gosc ZMIENIAJACY - tylko zbiera ===")
    kontekst = Kontekst(raport._raport, raport.arkusz)
    print("PRZED:", kontekst.odczytaj(2, "Klient"))

    anon = Anonimizuj().przemieć(wb)
    print("Po obchodzie, bez zastosowania:", kontekst.odczytaj(2, "Klient"))
    print("Metryki:", anon.metryki)
    print("Komend zebranych:", len(anon.komendy()))

    print("=== 3. Zastosowanie dopiero TERAZ (faza 2) ===")
    opisy, stos = zastosuj_komendy_goscia(anon, kontekst)
    print("PO zastosowaniu:", kontekst.odczytaj(2, "Klient"))

    print("=== 4. Cofniecie anonimizacji ===")
    stos.cofnij(kontekst, ile=len(opisy))
    print("PO cofnieciu:", kontekst.odczytaj(2, "Klient"))

    print("=== 5. Znowu anonimizacja i COMMIT ===")
    for komenda in anon.komendy():
        stos.wykonaj(komenda, kontekst)
    print("Zapisano:", raport._raport.zapisz("output/25_04_anonimowy.xlsx"))


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** `Anonimizuj._pseudonim` używa **deterministycznego** skrótu z solą: ta sama wartość zawsze daje ten sam pseudonim. To jest kluczowa decyzja projektowa — dzięki niej `Alfa sp. z o.o.` w wierszu 2 i w wierszu 4 dostaje **ten sam** pseudonim `OS-A1B2C3D4`. Bez tego anonimizacja **zniszczyłaby analizę** (nie dałoby się policzyć, ile pozycji ma ten sam klient). Z kolei **sól** sprawia, że pseudonim jest stabilny w obrębie raportu, ale **nie da się go odtworzyć** bez znajomości soli, więc nie da się go porównać z innym raportem — co jest dokładnie tym, czego chcesz od anonimizacji.

**I teraz najważniejsze:** po fazie 1 (`przemieć`) `odczytaj(2, "Klient")` **zwraca oryginalne nazwisko**. Dopiero `zastosuj_komendy_goscia` zmienia wartość. W kodzie `main` to jest widoczne w trzech `print`-ach: PRZED → po obchodzie (bez zmian!) → PO zastosowaniu. **To jest cała różnica między bezpiecznym a niebezpiecznym gościem** i można ją **zobaczyć na ekranie**.

**Co trafi do pliku.** `output/25_04_baza.xlsx` (oryginał, bez zmian — `zastosuj_komendy_goscia` zmieniło tylko pamięć) i `output/25_04_anonimowy.xlsx` (z pseudonimami `OS-XXXXXXXX` w kolumnie `Klient`). Otwórz oba — różnica ma być widoczna **właśnie w tej jednej kolumnie**.

**Trzy rzeczy do przemyślenia:**

1. **`OdwiedzSkoroszyt.przemieć()` jest oznaczony `@final`.** To znaczy (dla `mypy`), że gość nie nadpisuje przejścia. I to jest **właściwe**: gość ma być **reakcją**, nie zmieniać trasy. Gdyby któryś gość potrzebował innego przejścia, to sygnał, że to **nie jest** gość — to jest inna analiza z innym przejściem.
2. **`Anonimizuj.odwiedz_arkusz` nadpisuje pole `_kolumny_arkusza` w trakcie przejścia.** To jest mały stan pomocniczy, który **nie modyfikuje arkusza** — więc jest bezpieczny. Granica jest precyzyjna: **gość może zmieniać swój stan, nie może zmieniać obiektu odwiedzanego.** To jest różnica, którą trzeba umieć nazwać, bo inaczej „Visitor, który modyfikuje" jest nieunikniony.
3. **Czytający goście nie potrzebują `Kontekst`.** Uruchomienie `SzukajPustychKolumn().przemieć(wb)` nie wymaga żadnej infrastruktury zapisu. **To jest sygnał jakości tego wzorca**: im więcej gości **czysto czytających**, tym lepiej podzielona analiza. A gdy gość zmienia — produkuje komendy, które są **osobnym obiektem**, testowalnym bez uruchamiania gościa.

### Przykład 5 — Specification: filtr i kolor z jednej definicji

**Problem.** Zostawmy wreszcie tamten `if odbiorca == "zarzad"` z modułu 24. A do tego: chcemy filtrować dane **i** kolorować je warunkowo, z **jednej** definicji progu.

**Rozwiązanie.** Specyfikacje z kompozycją + generator reguł Excela.

```python
"""Modul 25, przyklad 5: Specification - filtr + kolor z jednej definicji."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill

from examples_25_fasada import RaportExcelRozszerzony


class BladSpecyfikacji(Exception):
    """Specyfikacja nie da sie przelozyc na formule Excela."""


# ======================================================================
# BAZA + KOMPOZYCJA
# ======================================================================
class Specyfikacja(ABC):
    @abstractmethod
    def nazwa(self) -> str: ...

    @abstractmethod
    def jest_spelniona(self, wiersz: Mapping[str, object]) -> bool: ...

    @abstractmethod
    def na_formule(self, adres: str) -> str | None:
        """Formuła Excela BEZ znaku '=' (pułapka z modułu 13!).

        Zwraca None, jezeli reguła nie ma odpowiednika w Excelu.
        """

    # --- kompozycja (bez 'not' - tego operatora nie da sie nadpisac!) ---
    def __and__(self, inna: "Specyfikacja") -> "Specyfikacja":
        return I(self, inna)

    def __or__(self, inna: "Specyfikacja") -> "Specyfikacja":
        return Lub(self, inna)

    def __invert__(self) -> "Specyfikacja":
        return Nie(self)

    def __repr__(self) -> str:
        return f"<{self.nazwa()}>"


@dataclass(frozen=True)
class I(Specyfikacja):
    lewa: Specyfikacja
    prawa: Specyfikacja

    def nazwa(self) -> str:
        return f"({self.lewa.nazwa()} ORAZ {self.prawa.nazwa()})"

    def jest_spelniona(self, wiersz) -> bool:
        return self.lewa.jest_spelniona(wiersz) and self.prawa.jest_spelniona(wiersz)

    def na_formule(self, adres: str) -> str | None:
        a = self.lewa.na_formule(adres)
        b = self.prawa.na_formule(adres)
        if a is None or b is None:
            return None                     # ← cala koniunkcja nieprzetlumaczalna
        return f"AND({a},{b})"


@dataclass(frozen=True)
class Lub(Specyfikacja):
    lewa: Specyfikacja
    prawa: Specyfikacja

    def nazwa(self) -> str:
        return f"({self.lewa.nazwa()} LUB {self.prawa.nazwa()})"

    def jest_spelniona(self, wiersz) -> bool:
        return self.lewa.jest_spelniona(wiersz) or self.prawa.jest_spelniona(wiersz)

    def na_formule(self, adres: str) -> str | None:
        a = self.lewa.na_formule(adres)
        b = self.prawa.na_formule(adres)
        if a is None or b is None:
            return None
        return f"OR({a},{b})"


@dataclass(frozen=True)
class Nie(Specyfikacja):
    wewnetrzna: Specyfikacja

    def nazwa(self) -> str:
        return f"NIE {self.wewnetrzna.nazwa()}"

    def jest_spelniona(self, wiersz) -> bool:
        return not self.wewnetrzna.jest_spelniona(wiersz)

    def na_formule(self, adres: str) -> str | None:
        wewnetrzna = self.wewnetrzna.na_formule(adres)
        if wewnetrzna is None:
            return None
        return f"NOT({wewnetrzna})"
```

```python
# ======================================================================
# SPECYFIKACJE ATOMOWE
# ======================================================================
def _adres_kolumny(adres: str, kolumna: str, kolumny: Mapping[str, str]) -> str | None:
    """Zamienia '$D2' + nazwa kolumny na konkretny adres kolumny."""
    litera = kolumny.get(kolumna)
    if litera is None:
        return None
    # adres przychodzi jako wzorzec: '$X2' -> podmieniamy X na wlasciwa litere
    return f"${litera}{adres.lstrip('$')[1:]}"


@dataclass(frozen=True)
class WiekszeNiz(Specyfikacja):
    kolumna: str
    prog: float

    def nazwa(self) -> str:
        return f"{self.kolumna} > {self.prog}"

    def jest_spelniona(self, wiersz) -> bool:
        wartosc = wiersz.get(self.kolumna)
        return isinstance(wartosc, (int, float)) and round(wartosc, 2) > self.prog

    def na_formule(self, adres: str) -> str | None:
        return f"ROUND({adres},2)>{self.prog}"     # ROUND: spojnosc z Pythonem!


@dataclass(frozen=True)
class Rowne(Specyfikacja):
    kolumna: str
    wzorzec: object

    def nazwa(self) -> str:
        return f"{self.kolumna} == {self.wzorzec!r}"

    def jest_spelniona(self, wiersz) -> bool:
        return wiersz.get(self.kolumna) == self.wzorzec

    def na_formule(self, adres: str) -> str | None:
        if isinstance(self.wzorzec, str):
            # ESCAPOWANIE (modul 21!): cudzyslow w tekscie rozerwalby formule
            if '"' in self.wzorzec or "\n" in self.wzorzec:
                raise BladSpecyfikacji(
                    f"Tekst {self.wzorzec!r} zawiera znaki niedozwolone w formule. "
                    "Odrzucamy - nie wstrzykujemy tresci do Excela."
                )
            return f'{adres}="{self.wzorzec}"'
        if isinstance(self.wzorzec, bool):
            return f"{adres}={str(self.wzorzec).upper()}"
        return f"{adres}={self.wzorzec}"


@dataclass(frozen=True)
class Zawiera(Specyfikacja):
    kolumna: str
    fragment: str

    def nazwa(self) -> str:
        return f"{self.kolumna} zawiera {self.fragment!r}"

    def jest_spelniona(self, wiersz) -> bool:
        wartosc = wiersz.get(self.kolumna)
        return isinstance(wartosc, str) and self.fragment.lower() in wartosc.lower()

    def na_formule(self, adres: str) -> str | None:
        if '"' in self.fragment or "\n" in self.fragment:
            raise BladSpecyfikacji(f"Fragment {self.fragment!r} nie nadaje sie do formuly")
        return f'ISNUMBER(SEARCH("{self.fragment}",{adres}))'


@dataclass(frozen=True)
class Puste(Specyfikacja):
    kolumna: str

    def nazwa(self) -> str:
        return f"{self.kolumna} puste"

    def jest_spelniona(self, wiersz) -> bool:
        return wiersz.get(self.kolumna) in (None, "")

    def na_formule(self, adres: str) -> str | None:
        return f"ISBLANK({adres})"


@dataclass(frozen=True)
class WgranejLiścieVIP(Specyfikacja):
    """PRZYKLAD specyfikacji NIEprzetlumaczalnej na formule Excela.

    Wersja realistyczna: lista VIP przychodzi z bazy/CRM i NIE moze
    trafic do arkusza (dane osobowe). Wiec filtr dziala, kolor - nie.
    """

    kolumna: str
    vip: frozenset[str]

    def nazwa(self) -> str:
        return f"{self.kolumna} w liscie VIP ({len(self.vip)} osob)"

    def jest_spelniona(self, wiersz) -> bool:
        return wiersz.get(self.kolumna) in self.vip

    def na_formule(self, adres: str) -> str | None:
        return None                      # ← UCZCIWIE: nie da sie
```

```python
# ======================================================================
# MOST: SPECYFIKACJA -> REGULA FORMATOWANIA WARUNKOWEGO
# ======================================================================
@dataclass(frozen=True)
class RolaKoloru:
    """Nanosnik reguly. Strategia go nie buduje - dostaje gotowy obiekt.

    MULTIPLEXER: to jest jedyne miejsce, ktore zna FormulaRule.
    """

    tlo: str = "FCE4E4"
    kolor_tekstu: str = "9C0006"
    pogrubienie: bool = False

    def zbuduj(self, formula: str) -> FormulaRule:
        return FormulaRule(
            formula=[formula],          # ← lista, BEZ znaku '='
            fill=PatternFill(fill_type="solid", bgColor=self.tlo),
            font=Font(color=self.kolor_tekstu, bold=self.pogrubienie),
        )


def mapa_kolumn(raport: RaportExcelRozszerzony, arkusz: str) -> dict[str, str]:
    """Nazwa kolumny -> litera. Wyliczane z naglowka, nigdy nie na sztywno."""
    from openpyxl.utils import get_column_letter

    return {
        nazwa: get_column_letter(i + 1)
        for i, nazwa in enumerate(raport.kolumny(arkusz))
    }


def zbuduj_reguly(
    raport: RaportExcelRozszerzony,
    arkusz: str,
    reguly: tuple[tuple[Specyfikacja, RolaKoloru], ...],
    *,
    wiersz_danych: int = 2,
) -> tuple[dict[str, str], tuple[str, ...]]:
    """Buduje reguly dla wszystkich specyfikacji, ktore DA SIE przetlumaczyc.

    Zwraca (metryki, ostrzezenia). Ostrzezenia sa WAZNE: informuja, ze
    czesc regul pominieto z powodu braku odpowiednika w Excelu.
    """
    kolumny = mapa_kolumn(raport, arkusz)
    ostrzezenia: list[str] = []
    zastosowane: list[str] = []

    for spec, rola in reguly:
        kolumna_glowna = getattr(spec, "kolumna", next(iter(kolumny)))
        wzorzec = f"${kolumny[kolumna_glowna]}{wiersz_danych}"
        formula = spec.na_formule(wzorzec)
        if formula is None:
            ostrzezenia.append(
                f"Regula '{spec.nazwa()}' nie ma odpowiednika w Excelu - "
                "dane sa filtrowane w Pythonie, koloru NIE bedzie."
            )
            continue
        regula = rola.zbuduj(formula)
        zakres = raport.regula_warunkowa(arkusz, "Cena", regula)
        zastosowane.append(f"{spec.nazwa()} -> {zakres}")

    return (
        {"regul_zastosowanych": len(zastosowane), "regul_pominietych": len(ostrzezenia)},
        tuple(ostrzezenia),
    )


def filtruj(dane: list[Mapping[str, object]], spec: Specyfikacja
            ) -> list[Mapping[str, object]]:
    """Ten SAM obiekt specyfikacji filtruje dane w Pythonie."""
    return [w for w in dane if spec.jest_spelniona(w)]
```

```python
def main() -> None:
    # DWA ZASTOSOWANIA JEDNEJ DEFINICJI:
    prog = WiekszeNiz("Cena", 100.0)
    pilne = Rowne("Region", "PL")
    spec = prog & pilne

    dane = [
        {"Klient": "Alfa", "Region": "PL", "Cena": 150.0},
        {"Klient": "Beta", "Region": "DE", "Cena": 150.0},
        {"Klient": "Gamma", "Region": "PL", "Cena": 40.0},
        {"Klient": "Delta", "Region": "PL", "Cena": 180.0},
    ]

    print("=== 1. FILTR (Python) ===")
    for w in filtruj(dane, spec):
        print("  ", w)
    print("  nazwa reguly:", spec.nazwa())

    print("=== 2. KOLOR (Excel) ===")
    print("  formula:", spec.na_formule("$D2"))
    print("  formula negacji:", (~prog).na_formule("$D2"))

    print("=== 3. SPECYFIKACJA NIEPRZETLUMACZALNA ===")
    vip = WgranejLiścieVIP("Klient", frozenset({"Alfa", "Delta"}))
    mieszana = spec & vip
    print("  jest_spelniona(Alfa):", mieszana.jest_spelniona(dane[0]))
    try:
        print("  na_formule:", mieszana.na_formule("$D2"))
    except BladSpecyfikacji as blad:
        print("  blad:", blad)

    print("=== 4. WSTRZYKIWANIE TEKSTU DO FORUMULY (modul 21) ===")
    try:
        zly = Rowne("Region", 'PL" OR "1"="1')
        print(zly.na_formule("$B2"))
    except BladSpecyfikacji as blad:
        print("  odrzucone:", blad)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Specyfikacje są `frozen=True` i **tworzą drzewo**. `prog & pilne` nie tworzy listy warunków — tworzy **obiekt `I`**, który trzyma **dwie referencje** do `WiekszeNiz` i `Rowne`. To znaczy, że drzewo specyfikacji można:
- **przekazać** do filtra (Python),
- **przekazać** do generatora reguł (Excel),
- **zalogować** (`nazwa()` daje czytelny opis),
- **przetestować** bez pliku i bez openpyxl.

I najważniejsze: **`zbuduj_reguly` nie zgaduje.** Jeżeli `na_formule` zwróci `None` (jak w `WgranejLiścieVIP`), reguła **nie powstaje**, a ostrzeżenie **wraca do wołającego**. To jest ta sama filozofia, co we wszystkich poprzednich modułach: **brak efektu musi być widoczny jako brak, a nie jako cichy brak efektu.**

**Co trafi do pliku.** Jeżeli `zbuduj_reguly` zostanie wywołane na raporcie, to w `xl/worksheets/sheet1.xml` pojawi się sekcja `conditionalFormatting` z zakresem np. `E2:E4`, a w niej `<cfRule type="expression">` z formułą `AND(ROUND($E2,2)>100.0,$B2="PL")`. **I zwróć uwagę na `$E2`:** kolumna absolutna, wiersz względny. To znaczy, że reguła **przesuwa się w dół** po wierszach, ale zostaje w kolumnie `E` — czyli działa tak, jak oczekujesz. Gdybyśmy wpisali `$E$2`, wszystkie wiersze sprawdzałyby wiersz 2.

**Trzy rzeczy do przemyślenia:**

1. **Jedna definicja, dwa światy.** `spec` filtruje dane w Pythonie i koloruje je w Excelu. Jeżeli zmienisz próg na 150, zmieniasz go **raz**. To jest most, o którym mówiłem w sekcji 2.5, i on realnie eliminuje klasę błędów „w Pythonie próg 100, w formule 150".
2. **`ROUND(...,2)` w formule i `round(...,2)` w Pythonie.** To jest ta jedna celowa linia obrony przed rozjazdem `float`. Bez niej formuła Excela i filtr w Pythonie mogłyby dać różne wyniki dla wartości typu `100.0000000001`. A `Round` w Excelu i `round()` w Pythonie mają tę samą semantykę dla liczb dodatnich — i to jest **sprawdzone**, a nie założone.
3. **`BladSpecyfikacji` przy cudzysłowie.** Specyfikacja `Rowne("Region", 'PL" OR "1"="1')` **rzuca wyjątek**, a nie generuje formułę. **Nie da się** wstrzyknąć formuły przez specyfikację, bo wstrzykiwacz **musi przejść przez escapowanie, a escapowanie albo odmawia, albo ucieka**. To jest jedyny bezpieczny sposób (moduł 21): nie „filtruj podejrzanych znaków", ale **odmów, gdy dane nie pasują do kontraktu**.

### Przykład 6 — Observer i Chain of Responsibility: postęp, ostrzeżenia i łańcuch wiersza

**Problem.** (1) Długi raport ma pokazywać postęp **i** zbierać ostrzeżenia (w tym „utrata funkcji pliku" z modułu 18). (2) Wiersze z wejścia przechodzą przez kroki: oczyszczenie, walidacja, normalizacja, konwersja typów — i mogą **odpadać**.

**Rozwiązanie.** Autobus zdarzeń + łańcuch ogniw.

```python
"""Modul 25, przyklad 6: Observer + Chain of Responsibility."""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any


# ======================================================================
# OBSERVER: autobus zdarzen
# ======================================================================
@dataclass(frozen=True)
class Zdarzenie:
    typ: str
    payload: dict


POSTEP = "postep"
OSTRZEZENIE = "ostrzezenie"
UTRATA_FUNKCJI = "utrata_funkcji"


class Autobus:
    """Minimalny event bus. Pierwszy argument handlera to zdarzenie."""

    def __init__(self) -> None:
        self._handlery: dict[str, list[Callable[[Zdarzenie], None]]] = {}
        self.bledy: list[str] = []          # wyjatki handlerow - ZEBRANE, nie polkniete

    def subskrybuj(self, typ: str, handler: Callable[[Zdarzenie], None]
                   ) -> Callable[[], None]:
        self._handlery.setdefault(typ, []).append(handler)

        def anuluj() -> None:
            self._handlery[typ].remove(handler)

        return anuluj                        # ← zwracamy anulowanie subskrypcji

    def publikuj(self, zdarzenie: Zdarzenie) -> None:
        for handler in tuple(self._handlery.get(zdarzenie.typ, ())):
            try:
                handler(zdarzenie)
            except Exception as blad:       # ← nie polykamy: ZBIERAMY
                self.bledy.append(
                    f"handler {getattr(handler, '__name__', handler)!r} "
                    f"przy {zdarzenie.typ!r}: {type(blad).__name__}: {blad}"
                )


# --- subskrybenci -----------------------------------------------------
class PostepNaKonsole:
    def __call__(self, zdarzenie: Zdarzenie) -> None:
        p = zdarzenie.payload
        print(f"  [postep] {p.get('etap', '?')}: {p.get('gotowe', 0)}/{p.get('razem', 0)}")


class ZbieraczOstrzezen:
    def __init__(self) -> None:
        self.ostrzezenia: list[str] = []

    def __call__(self, zdarzenie: Zdarzenie) -> None:
        self.ostrzezenia.append(str(zdarzenie.payload.get("tresc", "")))


class AlertUtratyFunkcji:
    """Reaguje na to, czego openpyxl NIE umie zapisac (modul 18)."""

    def __init__(self) -> None:
        self.utracone: list[str] = []

    def __call__(self, zdarzenie: Zdarzenie) -> None:
        element = str(zdarzenie.payload.get("element", ""))
        self.utracone.append(element)
        print(f"  [ALERT] plik straci: {element}")
```

```python
# ======================================================================
# CHAIN OF RESPONSIBILITY: kazde ogniwo moze odrzucic wiersz
# ======================================================================
@dataclass
class WynikOgniwa:
    """Albo przekazujemy dalej, albo odrzucamy z POWODEM."""
    wiersz: dict[str, Any] | None
    odrzucony_przez: str | None = None
    powod: str | None = None


class Ogniwo(ABC):
    """Ogniwo lancucha. Zna TYLKO nastepne ogniwo."""

    def __init__(self) -> None:
        self._nastepne: "Ogniwo | None" = None

    @property
    @abstractmethod
    def nazwa(self) -> str: ...

    @abstractmethod
    def przetworz(self, wiersz: dict[str, Any]) -> WynikOgniwa: ...

    def ustaw_nastepne(self, nastepne: "Ogniwo") -> "Ogniwo":
        self._nastepne = nastepne
        return nastepne                     # ← lancuch: a.ustaw(b).ustaw(c)

    def obsłuż(self, wiersz: dict[str, Any]) -> WynikOgniwa:
        wynik = self.przetworz(dict(wiersz))
        if wynik.wiersz is None:
            return wynik                    # ← STOP: odrzucony, dalej nie idzie
        if self._nastepne is None:
            return wynik
        return self._nastepne.obsłuż(wynik.wiersz)


class OczyscTekst(Ogniwo):
    nazwa = "oczysc"

    def przetworz(self, wiersz):
        for klucz, wartosc in wiersz.items():
            if isinstance(wartosc, str):
                wiersz[klucz] = " ".join(wartosc.split())
        return WynikOgniwa(wiersz)


class WalidujObecnosc(Ogniwo):
    nazwa = "waliduj"

    def __init__(self, wymagane: tuple[str, ...]) -> None:
        super().__init__()
        self._wymagane = wymagane

    def przetworz(self, wiersz):
        brakujace = [k for k in self._wymagane if not wiersz.get(k)]
        if brakujace:
            return WynikOgniwa(None, self.nazwa, f"brak pol: {brakujace}")
        return WynikOgniwa(wiersz)


class NormalizujRegion(Ogniwo):
    nazwa = "normalizuj"

    def __init__(self, mapowanie: dict[str, str]) -> None:
        super().__init__()
        self._mapowanie = mapowanie

    def przetworz(self, wiersz):
        region = str(wiersz.get("Region", "")).upper()
        wiersz["Region"] = self._mapowanie.get(region, region)
        return WynikOgniwa(wiersz)


class KonwersjaTypow(Ogniwo):
    nazwa = "konwersja"

    def przetworz(self, wiersz):
        try:
            wiersz["Ilosc"] = int(str(wiersz["Ilosc"]).replace(" ", ""))
            wiersz["Cena"] = float(str(wiersz["Cena"]).replace(",", ".").strip())
        except (KeyError, ValueError) as blad:
            return WynikOgniwa(None, self.nazwa, f"zla liczba: {blad}")
        if wiersz["Ilosc"] < 0 or wiersz["Cena"] < 0:
            return WynikOgniwa(None, self.nazwa, "wartosci ujemne")
        return WynikOgniwa(wiersz)


def zbuduj_lancuch() -> Ogniwo:
    """Konstrukcja lancucha - i sprawdzenie, ze ma ostatnie ogniwo."""
    oczysc = OczyscTekst()
    waliduj = WalidujObecnosc(("Klient", "Region", "Ilosc", "Cena"))
    normalizuj = NormalizujRegion({"PL": "PL", "POLSKA": "PL", "POL": "PL"})
    konwertuj = KonwersjaTypow()

    oczysc.ustaw_nastepne(waliduj).ustaw_nastepne(normalizuj).ustaw_nastepne(konwertuj)
    return oczysc


def przepusc(wiersze: list[dict], autobus: Autobus
             ) -> tuple[list[dict], list[tuple[str, str]]]:
    lancuch = zbuduj_lancuch()
    dobre: list[dict] = []
    odrzucone: list[tuple[str, str]] = []
    razem = len(wiersze)
    for i, wiersz in enumerate(wiersze, start=1):
        wynik = lancuch.obsłuż(wiersz)
        if wynik.wiersz is None:
            odrzucone.append((wynik.odrzucony_przez or "?", wynik.powod or ""))
        else:
            dobre.append(wynik.wiersz)
        if i % 2 == 0 or i == razem:
            autobus.publikuj(Zdarzenie(POSTEP, {
                "etap": "obrobka wierszy", "gotowe": i, "razem": razem}))
    return dobre, odrzucone
```

```python
def main() -> None:
    autobus = Autobus()
    zbieracz = ZbieraczOstrzezen()
    alert = AlertUtratyFunkcji()

    autobus.subskrybuj(POSTEP, PostepNaKonsole())
    autobus.subskrybuj(OSTRZEZENIE, zbieracz)
    autobus.subskrybuj(UTRATA_FUNKCJI, alert)

    def awaryjny(zdarzenie):
        raise RuntimeError("celowo psuje sie w handlerze")

    anuluj_awaryjny = autobus.subskrybuj(POSTEP, awaryjny)

    print("=== 1. Przetwarzanie z postepem ===")
    wiersze = [
        {"Klient": "  Alfa  sp. z o.o. ", "Region": "polska", "Ilosc": "1 0",
         "Cena": "25,00"},
        {"Klient": "", "Region": "PL", "Ilosc": "4", "Cena": "41.50"},
        {"Klient": "Gamma", "Region": "PL", "Ilosc": "22", "Cena": "abc"},
        {"Klient": "Delta", "Region": "DE", "Ilosc": "7", "Cena": "63,00"},
    ]
    dobre, odrzucone = przepusc(wiersze, autobus)
    print("Dobre:", dobre)
    print("Odrzucone:", odrzucone)

    print("=== 2. Zdarzenia zbierane, bledy handlerow ZEBRANE (nie polkniete) ===")
    autobus.publikuj(Zdarzenie(UTRATA_FUNKCJI, {"element": "sparkline grupa 1"}))
    autobus.publikuj(Zdarzenie(OSTRZEZENIE, {"tresc": "region 'PL' zmapowany z 'polska'"}))
    print("bledy handlerow:", autobus.bledy)
    print("ostrzezenia:", zbieracz.ostrzezenia)
    print("utracone funkcje:", alert.utracone)

    print("=== 3. Odsubskrybowanie awaryjnego handlera ===")
    anuluj_awaryjny()
    autobus.bledy.clear()
    autobus.publikuj(Zdarzenie(POSTEP, {"etap": "test", "gotowe": 1, "razem": 1}))
    print("bledy po odsubskrybowaniu:", autobus.bledy)


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Trzy rzeczy:

1. **`Autobus.publikuj` iteruje po `tuple(...)` kopii listy handlerów.** To nie kosmetyka: gdyby handler **odsubskrybował sam siebie** w trakcie publikacji (a to się zdarza — „zgłoś raz i przestań"), iteracja po żywej liście **zgubiłaby następnego handlera**. Kopia chroni przed tym samym problemem, o którym mówiło „nie modyfikuj listy w trakcie iteracji" (moduł 07).
2. **Wyjątek handlera trafia do `autobus.bledy`, a publikacja **trwa dalej**.** Sprawdź w wyniku: `awaryjny` rzuca, pozostali odbiorcy **i tak** dostają zdarzenie, a generowanie raportu **nie przerywa się**. Trzy zasady naraz: nie przerwij, nie połknij, zgłoś.
3. **`Ogniwo.obsłuż` pracuje na kopii (`dict(wiersz)`).** To znaczy, że oryginalny słownik nie jest modyfikowany — więc jeżeli wiersz zostanie odrzucony w 3. ogniwie, masz nadal **oryginał** do zapisania w arkuszu „Odrzucone". Analogia z sekcji 1.7: stanowisko może odrzucić produkt, ale **nie niszczy go**.

**Co trafi do pliku.** Ten przykład **celowo nic nie zapisuje** — i to jest jego wartość. Pokazuje, że warstwy Observer i CoR są **niezależne od Excela** i można je przetestować bez pliku. W `main` widać wszystko na ekranie: postęp, listę odrzuconych z **powodem i nazwą ogniwa**, zebrane ostrzeżenia i zebrane błędy handlerów.

**Trzy rzeczy do przemyślenia:**

1. **Odrzucenie ma powód i nazwę ogniwa.** `("konwersja", "zla liczba: could not convert string to float: 'abc'")`. To jest ta sama zasada, co komunikaty adapterów w module 24: **błąd ma mówić, co się stało i gdzie**. Bez tego `None` w wyniku byłby niemy.
2. **`autobus.bledy` jest publiczne, `_handlery` prywatne.** Konsekwentnie: to, co odbiorca ma czytać, jest publiczne. I to jest miejsce, w którym **Observer różni się od Decoratora z modułu 24**: Decorator **przerywa** (rzuca dalej) albo nie ma sensu; Observer **kontynuuje i zbiera** — bo jego zadaniem jest **informowanie**, nie przerywanie.
3. **Łańcuch ma jawnie zbudowane ostatnie ogniwo.** `KonwersjaTypow` nie ma następnego i zwraca wynik. Bez takiego ogniwa mielibyśmy **cichy `None`** na końcu łańcucha dla wierszy, które przeszły wszystkie kroki — dokładnie pułapka numer 9 z tego modułu.

## 4. Anatomia API

### Część A — konstrukcje Pythona używane w tym module

| Konstrukcja | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `abc.ABC` + `@abstractmethod` | klasa abstrakcyjna | — | **Wymusza** haczyki wymagane. Podstawa `StrategiaFormatowania`, `Komenda`, `Ogniwo` |
| `typing.final` | „nie nadpisuj" | — | Sprawdzane przez **mypy, nie przez interpreter**. Bez testu to deklaracja intencji |
| `__and__`, `__or__`, `__invert__` | operatory na obiektach | — | Dają `a & b`, `a \| b`, `~a`. **`not a` nie da się nadpisać** |
| `__call__` | obiekt wywoływalny | — | Subskrybenci w `Autobus` są obiektami, nie funkcjami (mają własny stan) |
| `functools.partial` | zamrożenie argumentów | — | Alternatywa dla klas-subskrybentów, gdy handler nie ma stanu |
| `copy.copy` / `copy.copy(obiekt)` | płytka kopia | — | Na **niemutowalnym** stylu: bezpieczne i wystarczające (moduł 09) |
| `dataclass(frozen=True)` | niemutowalny wynik | — | `Znalezisko`, specyfikacje, `Zdarzenie`, `RolaKoloru` |
| `dataclass` (mutowalny) | obiekt **ze stanem** | `field(init=False)` | **Tylko** komendy. Stan zmienia wyłącznie `wykonaj()` |
| `field(default_factory=list)` | pusta lista per instancja | — | Bez tego **wszystkie** makra dzielą jedną listę (klasyk!) |
| `reversed()` | odwrotna iteracja | — | **Warunek poprawności** cofania makra |
| `enumerate(iterable, start=1)` | licznik od 1 | — | Excel indeksuje od 1 (moduł 03) |
| generator (`yield`) | leniwe wiersze | — | `RaportStanuMagazynu._wiersze()`, `KonwersjaTypow` (moduł 19) |
| `collections.abc.Mapping` | typ dla „słownika" | — | `jest_spelniona(wiersz: Mapping)` — działa i słownik, i `WierszRaportu` jako `dict` |
| `frozenset` | niemutowalny zbiór | — | Lista VIP — bezpieczna w `frozen` specyfikacji |
| `hashlib.sha256` | skrót | `sol + wartosc` | Deterministyczny pseudonim: ta sama wartość → ten sam pseudonim |
| `tuple(...)` przed iteracją | kopia listy | — | `Autobus.publikuj` — chroni przed odsubskrybowaniem w trakcie |
| `dict(wiersz)` | płytka kopia | — | `Ogniwo.obsłuż` — oryginał zostaje dla arkusza „Odrzucone" |
| `str.isinstance`, `round(x, 2)` | normalizacja | — | Spójność Python ↔ Excel (specyfikacje) |
| `raise` własnego wyjątku | kontrakt warstwy | — | `BladSpecyfikacji` — wołający **nie łapie** openpyxl |
| `setattr`/`getattr` na klasie | dekorator klasy | — | `@audytowany` z modułu 24 |
| `Callable[[Zdarzenie], None]` | typ handlera | — | Czytelny kontrakt subskrypcji |

### Część B — `openpyxl` w kontekście tego modułu

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `FormulaRule(formula=[...], fill=, font=)` | reguła warunkowa z formuły | `formula` — **lista stringów bez `=`** | Jedno z **nielicznych** miejsc w tym module, gdzie widzimy typ openpyxl — i tylko w `RolaKoloru` |
| `ws.conditional_formatting.add(sqref, rule)` | dodanie reguły | `sqref` jak `"E2:E100"` | Zakres i `FormulaRule` są **osobno** — brak spójności to cichy błąd |
| `PatternFill(fill_type="solid", bgColor=...)` | tło reguły | `bgColor` (w `dxf`!), nie `fgColor` | W `FormulaRule` liczy się `bgColor` — **odwrotnie niż w stylu komórki** (pułapka 8) |
| `Font(color=, bold=)` | czcionka reguły | — | Wchodzi do `dxf`, nie do `fonts` |
| `ws.insert_rows(idx, amount=)` | wstawienie wierszy | `idx` **1-based** | **Nie przesuwa referencji formuł** (moduł 07) |
| `ws.delete_rows(idx, amount=)` | usunięcie wierszy | `idx` 1-based | Cofnięcie `insert_rows` — **liczbę wierszy** przywróci, formuł nie |
| `cell.data_type` | typ komórki | `"n"`, `"s"`, `"d"`, `"b"`, `"f"`, `"e"` | **`"f"` = formuła.** Podstawa `WykryjFormuly` |
| `cell.number_format` | format liczbowy | — | `KomendaSformatuj` zapamiętuje poprzedni |
| `cell.font` / `.fill` / `.alignment` | obiekty stylu | — | **Niemutowalne** → bezpieczne w migawce cofania (moduł 09) |
| `ws.iter_rows()` | iteracja komórek | — | W `OdwiedzSkoroszyt.przemieć`. **Nie modyfikuj w trakcie!** (pułapka 4) |
| `ws.max_row` / `ws.max_column` | wymiary | — | **Mogą kłamać** (moduł 07). `SzukajPustychKolumn` bierze to pod uwagę |
| `get_column_letter(n)` | indeks → litera | — | Jedno miejsce zamiany; używane w `Kontekst` i gościach |
| `ws.tables` | słownik tabel | klucz = `displayName` | `OdwiedzSkoroszyt` woła `odwiedz_tabele` dla każdej |
| `ws.sheetnames` (to `wb`) | lista arkuszy | — | Trasa przejścia gościa |
| `ws[1]` | pierwszy wiersz | — | Odczyt nagłówków w `Anonimizuj.odwiedz_arkusz` |
| `wb.properties.title` / `.creator` | metadane | — | `RaportExcelRozszerzony.metadane()` — DTO bez openpyxl |
| `ws.page_setup.orientation` | orientacja | `"landscape"` | W `druk_poziomo` |
| `ws.print_title_rows` | powtarzanie nagłówka | `"1:1"` | **Format zakresu, nie liczby** (pułapka z modułu 11) |
| `ws.sheet_properties.pageSetUpPr.fitToPage` | włączenie dopasowania | — | **Bez tego `fitToWidth` nie działa** — trzy ustawienia razem |
| `ws.conditional_formatting` | reguły arkusza | — | Do diagnostyki: `for cf in ws.conditional_formatting` |
| `wb.save(path)` | **COMMIT** | — | Po tym cofanie komend zmienia tylko pamięć, nie plik |
| `load_workbook(...)` | odczyt | `read_only`, `data_only` | Śledzenie utraty funkcji — moduł 18 |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — trzy strategie formatowania i przełącznik.**

Zaimplementuj **trzy** strategie formatowania na tej samej fasadzie:

```python
class Gesta(StrategiaFormatowania):     # male czcionki, waskie kolumny, brak paskow
class Powietrzna(StrategiaFormatowania): # duze wiersze, szerokie kolumny
class Audytowa(StrategiaFormatowania):   # wszystko widoczne, formaty jawne, zero kolorow
```

**Wymagania:**

1. Każda strategia **podaje własną nazwę** i zwraca **metryki**. Bez metryk nie zaliczam.
2. Żadna strategia **nie może** wystąpić w sygnaturze typ openpyxl. Sprawdź to `grep`-em:
   ```bash
   grep -n "def \|ws\|Workbook\|Font\|FormulaRule" examples/25_cw1.py | grep -v "^.*#" | grep -E "ws|Font|FormulaRule"
   ```
   Wynik: pusty (poza importami wewnątrz fasady i `FormulaRule` w `RolaKoloru`).
3. **Przełącznik** ma być **jednym** słownikiem: `STRATEGIE["gesta"]`. Żadnego `if/elif` przy wyborze.
4. **Test „identyczne wejście, różne strategie"**: ten sam raport wygenerowany trzema strategiami musi mieć **identyczne wartości** we wszystkich komórkach danych. Różni się **tylko** formatowanie.

**Pytania do `output/25_cw1_wnioski.md`:**

1. **Które z trzech strategii dzielą kod?** Jeżeli `Gesta` i `Powietrzna` mają identyczne ciało, różniące się dwiema liczbami — **popraw je na jedną klasę z parametrem**. Zapisz, ile linii to zaoszczędziło.
2. **Co się stanie, gdy ktoś wywoła `STRATEGIE["nieistnieje"]`?** `KeyError` jest w porządku — ale czy chcesz lepszy komunikat? Napisz funkcję `wybierz(nazwa)`, która rzuca `BladStrategii` z listą dostępnych. To ta sama zasada, co w fabryce adaptera (moduł 24).
3. **Czy `Audytowa` potrzebuje `Warunkowe`?** Uzasadnij: kolor w raporcie audytowym to **informacja** czy **ozdoba**? Ta odpowiedź jest ważniejsza, niż wygląda — w wielu regulacjach kolor nie może być nośnikiem informacji, bo traci ją po wydruku na czarno-białej drukarce.

### 🟡 Warsztat

**Zadanie 2 — `BazowyRaport` + dwa raporty konkretne.**

Weź `BazowyRaport` z przykładu 2 i zbuduj na nim **dwa nowe** raporty:

```python
class RaportNaleznosci(BazowyRaport):     # przeterminowane naleznosci
class RaportKPIzarzadu(BazowyRaport):     # same KPI, bez tabeli
```

**Wymagania:**

1. **`RaportNaleznosci`** — kolumny: `Klient`, `Kwota`, `Termin`, `Dni po terminie`. Format daty `dd.mm.yyyy`. **Reguła warunkowa** (`Warunkowe`) kolorująca wiersze po terminie — czyli musisz **połączyć** Template Method z Strategy i ze Specification (przykład 5). Wykres: brak.
2. **`RaportKPIzarzadu`** — **bez** `_wiersze()` w sensie tabeli! Zamiast danych ma same pary KPI. Ale `BazowyRaport` **wymaga** `_naglowek()` i `_wiersze()`. Jak to rozwiązać? **Trzy opcje** — wybierz i uzasadnij:
   - (a) `_naglowek()` zwraca `["KPI", "Wartosc"]`, a `_wiersze()` zwraca pary jako wiersze,
   - (b) **nowy** bazowy `BazowyRaportKPI` bez `_wiersze()` — i wtedy masz **dwa szkielety**,
   - (c) `_wiersze()` zwraca pusty iterator, a dane idą w `_podsumowanie()`.
3. **`_podsumowanie` w obu raportach** — w `RaportNaleznosci` suma kwot i liczba przeterminowanych, w `KPIzarzadu` cztery pary.
4. **`_druk()`** — oba raporty mają iść na A4 poziomo z powtarzanym nagłówkiem.
5. **Test, który nie zna openpyxl:** napisz `CelTestowy` implementujący `CelRaportu` (moduł 24) — czyli liczy wywołania `podnaglowek`, `akapit`, `tabela`, `para`. Dla `KPIzarzadu` sprawdź, że **nie ma wywołania `tabela`**, a dla `Naleznosci` — że **jest**.

**Pomiar (obowiązkowy).** W `output/25_cw2_pomiar.md` porównaj:

| Raport | Linie kodu podklasy | Linie szkieletu (współdzielone) |
|---|---|---|
| Wersja bez Template Method (napisz ją!) | ? | 0 |
| Wersja z `BazowyRaport` | ? | ? |

To ostatnie ćwiczenie — **napisz wersję bez wzorca** — jest najważniejsze. Dopiero wtedy zobaczysz **ile kodu powtarza się** w szkieletach dwóch raportów. Bez tego pomiaru Template Method jest wiarą, nie decyzją.

**Pytania do `output/25_cw2_wnioski.md`:**

1. **Gdzie w Twoim `RaportNaleznosci` leży logika („dni po terminie")?** W `_wiersze()`? A powinna? **Odpowiedź:** dziś może zostać, ale to **dług do spłaty w module 26** — logika biznesowa należy do domeny, podklasa ma tylko **prezentować**.
2. **`KPIzarzadu` bez tabeli — czy to nadużycie wzorca?** Jeżeli `_wiersze()` zwraca pusty iterator, to po co w ogóle jest? Uzasadnij **uczciwie**, czy nie lepiej mieć dwa szkielety (opcja b).
3. **Czy `_podsumowanie` powinno być `@abstractmethod`?** Sprawdź: w Twoich dwóch raportach jest nadpisane. W `RaportStanuMagazynu` też. **Istnieje jakakolwiek podklasa, która NIE nadpisuje `_podsumowanie`?** Jeżeli nie — **zmień je na `@abstractmethod`**. To jest pułapka numer 2 i sprawdzasz ją na własnym kodzie.

### 🔴 Wyzwanie

**Zadanie 3 — `StosKomend` z cofaniem i logiem audytowym.**

To jest największe ćwiczenie tego modułu. Rozbudujesz `StosKomend` do wersji **produkcyjnej** i zrobisz na nim coś, czego nie było w żadnym module: **transakcję z potwierdzeniem**.

**Krok 1 — rozbuduj komendy.**

Dodaj `KomendaUsunWiersz` (usuwa wiersz danych; `cofnij` musi **odtworzyć** całą zawartość wiersza — czyli musi zapamiętać **wszystkie** komórki wiersza) i `KomendaScalKomorki` (scala zakres; `cofnij` rozscala — ale **uczciwie udokumentuj**, że wartości z komórek innych niż lewa-górna **giną** i cofanie ich **nie odtworzy**, moduł 11).

**Krok 2 — `StosKomend` z transakcjami nazwanymi.**

```python
class StosKomend:
    def rozpocznij_transakcje(self, nazwa: str) -> str: ...
    def zatwierdz_transakcje(self, nazwa: str) -> tuple[str, ...]: ...
    def cofnij_transakcje(self, kontekst, nazwa: str) -> tuple[str, ...]: ...
    def transakcje(self) -> tuple[dict, ...]: ...
```

Transakcja to **zakres stosu** z nazwą. `cofnij_transakcje` cofa **wszystkie** komendy z tego zakresu (od ostatniej do pierwszej).

**Krok 3 — potwierdzenie przed commitem (brama z modułu 24).**

Napisz funkcję:

```python
def transakcja_z_potwierdzeniem(
    kontekst: Kontekst,
    stos: StosKomend,
    nazwa: str,
    komendy: list[Komenda],
    *,
    podglad: Callable[[StosKomend], str],
    potwierdzenie: Callable[[str], bool],
    zapis: Callable[[], Path],
) -> dict:
    """Wykonuje komendy, pokazuje podglad, pyta o potwierdzenie, zapisuje ALBO cofa."""
```

Wymagania:
- jeżeli `potwierdzenie` zwróci `False` → **cofnij transakcję**, **nie zapisuj**, zwróć `{"zapisano": False, "cofnieto": N}`,
- jeżeli `zapis()` rzuci wyjątek → **cofnij transakcję**, **nie połknij wyjątku**, dodaj do wyniku `"blad"`,
- jeżeli wszystko OK → `{"zapisano": True, "cel": path, "komend": N}`.

**Krok 4 — log audytowy, który przetrwa.**

Dodaj do `StosKomend` metodę `eksportuj_audyt(sciezka: Path) -> Path`, która zapisuje dziennik **jako plik JSON** (nie jako arkusz — bo dziennik nie może zależeć od tego, że zapis Excela się udał). Wymagania:
- każdy wpis ma `czas`, `operacja`, `opis`, `autor`, `transakcja`,
- plik jest zapisywany **atomowo** (`mkstemp` + `os.replace`, moduł 24),
- **test**: po `eksportuj_audyt` i ponownym wczytaniu JSON-a liczba wpisów jest zgodna.

**Krok 5 — dwufazowa anonimizacja jako transakcja.**

Połącz gościa z przykładu 4 z transakcją z kroku 3:

```python
def anonimizuj_z_potwierdzeniem(wb, kontekst, stos, *, potwierdzenie, zapis):
    gosc = Anonimizuj().przemieć(wb)           # FAZA 1: nic nie zmienione
    ...
```

Wymagania:
- **faza 1 nie może nic zmienić** — test: po `przemieć` wartości w arkuszu są **oryginalne**,
- podgląd pokazuje **liczbę zmian i przykład** (`OS-XXXXXXXX`),
- odmowa → **arkusz wygląda jak przed uruchomieniem**, test to sprawdza komórka po komórce.

**Kryteria oceny (rubryka):**

| Kryterium | Waga | Na co patrzę |
|---|---|---|
| Cofanie `KomendaUstawKomorke` rozróżnia `None` od wartości | ★★★ | wpisanie `0` i cofnięcie daje `0`, nie puste |
| Komenda nie zmienia nic w konstruktorze | ★★★ | `KomendaUstawKomorke(...)` bez `wykonaj` → arkusz bez zmian |
| `MakroKomend.cofnij` w odwrotnej kolejności | ★★★ | makro „wstaw wiersz + wpisz w niego" cofa się poprawnie |
| Transakcja: cofnij zakres, nie całość | ★★☆ | dwie transakcje → cofnięcie drugiej nie rusza pierwszej |
| Odmowa potwierdzenia = brak zmian w arkuszu | ★★★ | test komórka po komórce |
| Zapis przy błędzie `save()` **nie zostawia** pliku | ★★☆ | brak pliku tymczasowego (moduł 24) |
| Audyt w JSON z `transakcja` | ★★☆ | 10 wpisów → 10 rekordów z polem `transakcja` |
| Dwufazowość: faza 1 nic nie zmienia | ★★★ | test po `przemieć` |

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe decyzje i fragmenty</strong></summary>

**Decyzja 1: cofanie usunięcia wiersza wymaga zapamiętania wzorca stylów, nie tylko wartości.**

```python
@dataclass
class KomendaUsunWiersz(Komenda):
    """Usuwa wiersz danych. Cofanie odtwarza wartosci I STYLE z migawki.

    UWAGA: nie odtwarza formul tak, jak byly - openpyxl zapisuje formuly jako
    tekst. Odtworzymy je wiernie, ale Excel przeliczy je od nowa przy otwarciu
    (modul 06: openpyxl nie ma silnika obliczen). To jest akceptowalne,
    ale MUSI byc udokumentowane.
    """

    arkusz: str
    wiersz: int
    kolumny: tuple[str, ...]
    autor: str = "system"
    _migawka: list[dict] = field(default_factory=list, init=False, repr=False)

    @property
    def opis(self) -> str:
        return f"DELETE {self.arkusz}!wiersz {self.wiersz}"

    def wykonaj(self, kontekst: Kontekst) -> None:
        self._migawka = [
            {
                "kolumna": k,
                "wartosc": kontekst.odczytaj(self.wiersz, k),
                "styl": kontekst.styl(self.wiersz, k),
            }
            for k in self.kolumny
        ]
        kontekst.usun_wiersze(self.wiersz, 1)

    def cofnij(self, kontekst: Kontekst) -> None:
        kontekst.wstaw_wiersze(self.wiersz, 1)
        for migawka in self._migawka:
            kontekst.ustaw(self.wiersz, migawka["kolumna"], migawka["wartosc"])
            kontekst.zastosuj_styl(self.wiersz, migawka["kolumna"], migawka["styl"])
```

**Decyzja 2: `KomendaScalKomorki` — cofanie NIE odtwarza danych, i to jest w kontrakcie.**

```python
@dataclass
class KomendaScalKomorki(Komenda):
    """Scala zakres. COFANIE ROZSCALA, ale NIE ODTWORZY wartosci
    z komorek innych niz lewa-gorna - sa tracone bezpowrotnie (modul 11).

    Dlatego w `wykonaj` sprawdzamy, czy komorki sa puste, i jezeli nie sa
    i nie ma jawnego `force=True` - RZUCAMY. Nie dopuszczamy cichej utraty.
    """

    arkusz: str
    zakres: str                     # np. "A1:D1"
    force: bool = False
    autor: str = "system"
    _scalone: bool = field(default=False, init=False, repr=False)

    @property
    def opis(self) -> str:
        return f"MERGE {self.arkusz}!{self.zakres}"

    def wykonaj(self, kontekst: Kontekst) -> None:
        if not self.force:
            niepuste = kontekst.policz_niepuste_poza_lewa(self.zakres)
            if niepuste:
                raise BladRaportu(
                    f"Scalenie {self.zakres} usunie dane w {niepuste} komorkach. "
                    "Uzyj force=True, jezeli wiesz, co robisz."
                )
        kontekst.scal(self.zakres)
        self._scalone = True

    def cofnij(self, kontekst: Kontekst) -> None:
        if self._scalone:
            kontekst.rozscal(self.zakres)
```

To jest **jedyna obrona** przed cichą utratą danych w scalaniu: **sprawdź przed, nie po**. Ta sama filozofia co `BladProtokolu` z modułu 24 i `BladWalidacjiDanych` z modułu 24 — błąd wybucha **w miejscu, w którym można go zrozumieć**.

**Decyzja 3: transakcja jako znacznik na stosie.**

```python
class StosKomend:
    def __init__(self) -> None:
        self._wykonane: list[Komenda] = []
        self._cofniete: list[Komenda] = []
        self._audyt: list[dict] = []
        self._transakcje: list[dict] = []      # {"nazwa", "od", "do"}
        self._otwarte: dict[str, int] = {}

    def rozpocznij_transakcje(self, nazwa: str) -> str:
        if nazwa in self._otwarte:
            raise ValueError(f"Transakcja {nazwa!r} juz otwarta")
        self._otwarte[nazwa] = len(self._wykonane)
        return nazwa

    def zatwierdz_transakcje(self, nazwa: str) -> tuple[str, ...]:
        od = self._otwarte.pop(nazwa, None)
        if od is None:
            raise ValueError(f"Brak otwartej transakcji {nazwa!r}")
        opis = tuple(k.opis for k in self._wykonane[od:])
        self._transakcje.append({"nazwa": nazwa, "od": od,
                                 "do": len(self._wykonane), "stan": "zatwierdzona"})
        return opis

    def cofnij_transakcje(self, kontekst: Kontekst, nazwa: str) -> tuple[str, ...]:
        od = self._otwarte.pop(nazwa, None)
        if od is None:
            # cofamy ZAMKNIETA transakcje, jezeli jest na gorze stosu
            rekord = next((t for t in reversed(self._transakcje)
                           if t["nazwa"] == nazwa and t["stan"] == "zatwierdzona"), None)
            if rekord is None:
                raise ValueError(f"Nie ma transakcji {nazwa!r}")
            if rekord["do"] != len(self._wykonane):
                raise ValueError(
                    f"Transakcja {nazwa!r} nie jest na gorze stosu - "
                    "nie mozna jej cofnac bez cofniecia pozniejszych zmian."
                )
            od = rekord["od"]
        ile = len(self._wykonane) - od
        cofniete = self.cofnij(kontekst, ile=ile)
        self._transakcje.append({"nazwa": nazwa, "od": od,
                                 "do": od, "stan": "cofnieta"})
        return tuple(cofniete)
```

**Decyzja 4: brama potwierdzenia, która NIE połyka wyjątku zapisu.**

```python
def transakcja_z_potwierdzeniem(
    kontekst: Kontekst,
    stos: StosKomend,
    nazwa: str,
    komendy: list[Komenda],
    *,
    podglad,
    potwierdzenie,
    zapis,
) -> dict:
    """Wykonaj -> pokaz -> zapytaj -> zapisz ALBO cofnij.

    UWAGA: jezeli `zapis` rzuci, NIE polykamy wyjatku - dopisujemy go
    do wyniku i podnosimy dalej. Nieme tryby awarii sa zakazane (modul 24).
    """
    stos.rozpocznij_transakcje(nazwa)
    for komenda in komendy:
        stos.wykonaj(komenda, kontekst)

    opis = podglad(stos)
    if not potwierdzenie(opis):
        cofniete = stos.cofnij_transakcje(kontekst, nazwa)
        return {"zapisano": False, "cofnieto": len(cofniete),
                "opis": opis, "transakcja": nazwa}

    try:
        cel = zapis()                                   # COMMIT
    except Exception as blad:
        cofniete = stos.cofnij_transakcje(kontekst, nazwa)
        raise BladRaportu(
            f"Zapis nie udal sie ({type(blad).__name__}: {blad}); "
            f"cofnieto {len(cofniete)} komend. Plik na dysku NIE zmieniony."
        ) from blad

    stos.zatwierdz_transakcje(nazwa)
    return {"zapisano": True, "cel": str(cel), "komend": len(komendy),
            "transakcja": nazwa}
```

**Dlaczego `raise ... from blad`:** zrzut stosu pokaże **oba** wyjątki — oryginalny (np. `PermissionError` z `os.replace`) i nasz (`BladRaportu`). To jest różnica między „nie wiem, co się stało" a „wiem, że zapis się nie udał, bo plik był zablokowany, i cofnąłem 3 komendy". Moduł 24, pułapka 2: **dekorator/warstwa ma dodawać informację, nie odbierać.**

**Decyzja 5: audyt do JSON, nie do arkusza.**

```python
    def eksportuj_audyt(self, sciezka) -> Path:
        """Audyt leci do JSON, NIE do arkusza.

        Powod: dziennik nie moze zalezec od tego, czy zapis Excela sie udal.
        Gdyby byl arkuszem, to przy bledzie zapisu stracilibysmy DOWOD,
        ze cos probowano zmienic.
        """
        import json
        import os
        import tempfile

        sciezka = Path(sciezka)
        sciezka.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(suffix=".json", dir=sciezka.parent)
        os.close(fd)
        try:
            Path(tmp).write_text(
                json.dumps(self._audyt, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            os.replace(tmp, sciezka)
        except BaseException:
            Path(tmp).unlink(missing_ok=True)
            raise
        return sciezka
```

**Czego tu świadomie nie robię:**

- **Nie robię `powtórz()` dla `KomendaUsunWiersz`.** Byłoby możliwe (migawka jest zapamiętana), ale to rzadki przypadek i mnoży powierzchnię błędu. Zostawiam `ponow()` tylko dla komend **idempotentnych z natury** (`KomendaUstawKomorke`, `KomendaSformatuj`) i **dokumentuję ograniczenie**. Lepiej mieć jawnie ograniczony stos niż stos, który „czasem" działa.
- **Nie synchronizuję stosu z zawartością arkusza.** Jeżeli ktoś zmieni komórkę **poza** `StosKomend`, stos o tym nie wie i cofanie może przywrócić „nieaktualną" wartość. Rozwiązanie produkcyjne to **wersjonowanie arkusza** albo **blokada zapisu poza stosem** — ale to jest temat na oddzielny wzorzec (i moduł 26).
- **Nie zapisuję audytu w tym samym pliku co raport.** Osobny plik JSON. Gdyby dziennik był arkuszem w tym samym skoroszycie, to **przy nieudanym zapisie raportu nie byłoby dziennika** — a dziennik jest właśnie od tego, żeby istniał, gdy coś poszło źle.

</details>

## 6. Typowe błędy i pułapki

**1. Objaw: „mam pięć typów raportów i wszędzie `if typ == 'sprzedaz': ... elif typ == 'magazyn': ...`, a teraz doszedł szósty i musiałem poprawić cztery pliki."**
→ *Przyczyna:* **`if/elif` zamiast strategii.** Rozgałęzienie rozsypane po kodzie. Ten sam warunek powtarza się w generatorze, w formatowaniu i w zapisie — i za każdym razem trzeba go pamiętać.
→ *Naprawa:* **strategia + rejestr.** Jedna linia wyboru (`STRATEGIE[nazwa]`), a warianty jako klasy. Uwaga na **odwrotność**: jeżeli masz `if` **w jednym miejscu** i **dwa** warianty, to `if` jest lepszy (patrz pułapka 2). Kryterium: **rozgałęzienie w więcej niż jednym miejscu** albo **więcej niż dwa warianty**.

**2. Objaw: „stworzyłem `StrategiaFormatowania`, `StrategiaFormatowaniaZebra`, `StrategiaFormatowaniaKolor`, `StrategiaFormatowaniaBezPasków` — i każda ma jedno pole różnicy."**
→ *Przyczyna:* **strategia dla jednego przypadku / klasa na każdy parametr.** Wzorzec użyty **bez uzasadnienia**: każda strategia ma identyczne ciało, różni się jedną liczbą. To jest mnożenie bytów (i złamanie zasady z sekcji 2.1: „warianty różnią się jednym parametrem → konfiguracja, nie wzorzec").
→ *Naprawa:* **jedna klasa + parametr konstruktora.** `Zebrowe(kolor="F2F2F5", co_ile=1)`. Kryterium mechaniczne: **jeżeli dwa ciała metod są identyczne poza literałami, to jedna klasa.** Sprawdź to `diff`-em — dosłownie porównaj dwa pliki albo dwie definicje. Jeżeli różnią się dwiema liniami, **zwijaj**.

**3. Objaw: „nowy raport wygenerował się poprawnie, ale odbiorca mówi, że brakuje podsumowania — a ja nie mam pojęcia, że to pominąłem."**
→ *Przyczyna:* **Template Method z haczykiem, który nic nie robi i milczy.** `_podsumowanie()` ma domyślne `return None`, a podklasa go nie nadpisała. Nic nie wybuchło, bo haczyk jest opcjonalny — a mimo to **raport jest niekompletny**.
→ *Naprawa:* **trzy poziomy obrony z sekcji 2.2.** (1) Jeżeli podsumowanie jest **zawsze** wymagane — `@abstractmethod` i puste podsumowanie wymaga świadomej decyzji `return None`. (2) Jeżeli **czasem** wymagane — haczyk ma **zwracać coś do metryk** (np. `return []`), a pusta lista to informacja dla kontekstu, nie cisza. (3) **Test katalogowy** przechodzący po wszystkich podklasach. Diagnostyka: policz, ile podklas `BazowyRaport` **nie** nadpisuje `_podsumowanie`. Jeżeli zero — **zmień haczyk na `@abstractmethod`**.

**4. Objaw: „komenda `KomendaUstawKomorke(...)` — sam **konstruktor** zmienił wartość w arkuszu, a `wykonaj()` zmienił ją drugi raz."**
→ *Przyczyna:* **Command modyfikujący skoroszyt przed `wykonaj()`.** Komenda zapamiętuje poprzednią wartość **w konstruktorze** i **od razu** ustawia nową. To zlewa „utwórz" z „wykonaj" i niszczy wszystkie cztery pożytki wzorca (sekcja 2.3): makra bez wykonania, podglądu przed wykonaniem, bramy potwierdzenia i prawdziwego logu audytu.
→ *Naprawa:* **konstruktor tylko przyjmuje dane.** Zero dostępu do kontekstu w `__init__`. Stan (poprzednia wartość, migawka stylu) zapisywany **wyłącznie** w `wykonaj()`. Test: utwórz komendę i **nie wykonuj** — arkusz musi być bez zmian. To jedno zdanie testu łapie cały błąd.

**5. Objaw: „`jest_spelniona()` działa, ale wersja dla Excela odpowiada raz `True`, raz `False` przy tych samych danych — a czasem specyfikacja wykonuje zapytanie do bazy i generowanie raportu trwa dziesięć minut."**
→ *Przyczyna:* **Specification wykonująca wejście/wyjście podczas ewaluacji.** Dwa objawy, jedna przyczyna: specyfikacja, która **nie jest predykatem**, a **zapytaniem**. Woła bazę, czyta plik, sięga po usługę sieciową — więc (1) wynik zależy od **stanu świata**, nie od danych, i (2) koszt ewaluacji jest nieprzewidywalny.
→ *Naprawa:* **specyfikacja jest czystą funkcją danych.** Jeżeli potrzebuje danych z zewnątrz — **wstrzyknij je w konstruktorze** (`WgranejLiścieVIP("Klient", frozenset(...))`), pobierz je **raz** i **zamroź**. Diagnostyka: wpisz `print`/`log.debug` na początku `jest_spelniona` i policz wywołania. Jeżeli liczba wywołań ≠ liczba wierszy — coś jest nie tak. Drugi test: **determinizm** — dwa wywołania na tym samym wierszu muszą dać ten sam wynik.

**6. Objaw: „uruchomiłem `Anonimizuj().przemieć(wb)` i **od razu** zmieniło mi dane, a przy okazji zgubiło trzy komórki."**
→ *Przyczyna:* **Visitor, który modyfikuje to, co odwiedza.** `odwiedz_komorke` zmienia wartość „na miejscu". Konsekwencje: brak fazy podglądu, brak cofania, brak raportu „co znaleziono" — oraz **ryzyko gubienia komórek przy wstawianiu/usuwaniu w całej strukturze** (moduł 07).
→ *Naprawa:* **dwie fazy.** `odwiedz_*` **zbiera** (`Znalezisko`) albo **produkuje komendy**, ale **nic nie zmienia**. Zmiana dzieje się w kroku drugim — najlepiej przez `StosKomend` (i wtedy jest cofalna). Kryterium mechaniczne: w metodach `odwiedz_*` **nie może** wystąpić przypisanie do `komorka.value`, `ws.cell(...).value` ani `insert/delete`. Dodatkowo: **gość nie może zmieniać obiektu odwiedzanego, ale może zmieniać swój własny stan** (`self._znaleziska`, `self._mapa`).

**7. Objaw: „`Anonimizuj` zamienił nazwiska na pseudonimy, ale teraz w raporcie **każdy klient jest jednym klientem** — nie policzę już, ile pozycji ma Alfa."**
→ *Przyczyna:* **niedeterministyczna anonimizacja.** Jeżeli pseudonim zależy od `uuid4()`, indeksu wiersza albo losowej soli generowanej przy każdym uruchomieniu, to **ta sama wartość daje różne pseudonimy**. Relacja „ten sam klient" ginie.
→ *Naprawa:* **skrót z solą, deterministyczny.** `hashlib.sha256((sol + wartosc).encode())` — ta sama wartość zawsze ten sam pseudonim. Sól ma być **stała w obrębie raportu** (żeby relacje były zachowane) i **niejawna poza nim** (żeby nie dało się odtworzyć z innego raportu). I jeszcze jedno, bardzo praktyczne: **zapisz w metrykach liczbę unikalnych wartości** (`unikalnych_wartosci: 3`, `do_anonimizacji: 42`). Ta para liczb natychmiast pokazuje, czy anonimizacja **skleiła** różne wartości (gdy liczba pseudonimów jest podejrzanie mała).

**8. Objaw: „dodałem `FormulaRule` do zakresu i kolorów w ogóle nie ma — albo cała kolumna jest czerwona, choć miał być jeden wiersz."**
→ *Przyczyna:* **trzy typowe pomyłki reguł warunkowych.** (a) Formuła w `FormulaRule` zaczyna się od `=`, co psuje regułę (moduł 13) — ma być **bez** `=`. (b) Adres w formule nie odpowiada zakresowi: `$E$2` w zakresie `E2:E100` koloruje **wszystko** albo **nic**, bo wszystkie wiersze patrzą na wiersz 2. Trzeba użyć **`$E2`** (kolumna absolutna, wiersz względny). (c) W `PatternFill` na potrzeby `dxf` liczy się **`bgColor`**, nie `fgColor` — odwrotnie niż w stylu komórki (moduł 09).
→ *Naprawa:* **formuła bez `=`; adres zgodny z zakresem i „kotwicą kotwiczenia"; `bgColor` w regule.** Diagnostyka: po wczytaniu pliku przejrzyj `ws.conditional_formatting` — zobaczysz `cf.sqref` i `cf.rules[0].formula`. Test: wygeneruj plik z **jednym** wierszem spełniającym warunek i jednym nie, otwórz i **sprawdź okiem** — to jest jedyny test, który wykrywa błąd kotwiczenia.

**9. Objaw: „łańcuch przetworzył 200 wierszy, ale do arkusza trafiło 0 — i nie ma żadnego błędu."**
→ *Przyczyna:* **łańcuch bez ostatniego ogniwa** albo ogniwo, które zwraca `None` dla przepuszczonego obiektu. Bez finalnego „to jest koniec, zwróć wiersz" każdy wiersz **kończy** w `None`, a kod nie odróżnia „przepuszczony" od „odrzucony".
→ *Naprawa:* **jawne ostatnie ogniwo** (które nie ma następnego i **zwraca** wynik) **plus wynik, który rozróżnia przepuszczenie od odrzucenia.** W naszym szkicu to `WynikOgniwa(wiersz, odrzucony_przez, powod)`. Nigdy `None`, bo `None` jest dwuznaczne: „nie ma czym przekazać" i „odrzucone". Diagnostyka: policz osobno „weszło", „wyszło", „odrzucone". Jeżeli `weszło > wyszło + odrzucone`, masz znikające wiersze.

**10. Objaw: „handler postępu rzucił wyjątek i generowanie raportu **przerwało się** w połowie — a raport i tak nie powstał, bo nie było `save()`."**
→ *Przyczyna:* **Observer z handlerem, który przerywa publikację.** `Autobus.publikuj` iteruje po handlerach i **nie łapie** wyjątków. Jeden zepsuty subskrybent (np. aktualizacja paska postępu bez działającego GUI) **zatrzymuje całe przetwarzanie**.
→ *Naprawa:* **handler nie może przerwać publikacji — ale jego błąd nie może zniknąć.** Złap wyjątek, **zapisz go** w `autobus.bledy` (z nazwą handlera!), **kontynuuj** do pozostałych, a na końcu raportu sprawdź `bledy` i **dołącz ostrzeżenie**. To jest różnica wobec Decoratora z modułu 24: Decorator **musi** przepuszczać wyjątki dalej; Observer **musi** je zebrać i iść dalej. Różne role — różne reguły.

**11. Objaw: „opisałem pięć typów raportów, trzy strategie formatowania, dwa eksporty — i mam **trzydzieści** klas, w których się gubię."**
→ *Przyczyna:* **strategia zamiast dwóch osi.** Ktoś zbudował **jedną** strategię na iloczyn: `SprzedazXlsx`, `SprzedazBytes`, `MagazynXlsx`… i tak dalej. To jest **mnożenie** zamiast **dodawania**, czyli dokładne przeciwieństwo korzyści z sekcji 2.1.
→ *Naprawa:* **rozpoznaj osie i rozdziel.** Jeżeli dwa wymiary zmieniają się **niezależnie** (typ danych i kanał wyjścia), to są **dwie** strategie, nie jedna. Liczba klas spada z iloczynu do sumy. Diagnostyka: nazwy klas zawierają **dwa** wymiary naraz (`SprzedazXlsx`)? To sygnał. Rozbij na `Sprzedaz` (dane/wygląd) i `Xlsx` (kanał). To jest Bridge — i to jest ta sama lekcja, co „trzy osie wymiarów" z modułu 23.

**12. Objaw: „`not spec` zwraca `False`, chociaż specyfikacja mówi, że warunek **nie** jest spełniony. A operator `|` działa. O co chodzi?"**
→ *Przyczyna:* **`not` nie da się przeciążyć.** `and`, `or`, `not`, `is`, `in` to operatory **języka**, nie metody — więc `__and__`/`__or__` dają `&`/`|`, ale `not` **nie ma** odpowiednika. `not spec` sprawdza **prawdziwość obiektu**, a każdy obiekt jest prawdziwy → `not spec` to zawsze `False`. To cicha katastrofa, bo kod **działa** i **kłamie**.
→ *Naprawa:* **używaj `~spec` i `&`/`|`** — nigdy `not`, `and`, `or` na obiektach specyfikacji. Zabezpieczenie: zdefiniuj `__bool__`, które **rzuca wyjątek**:

```python
    def __bool__(self) -> bool:
        raise TypeError(
            "Nie uzywaj specyfikacji w warunkach Pythona (not/and/or). "
            "Uzyj operatorow: & | ~ . Przyklad: spec_a & ~spec_b"
        )
```

To zamienia **cichy błąd logiczny** w **natychmiastowy błąd wykonania** — i jest to jedna z najcenniejszych linijek w całym module. Wypróbuj: `if not spec:` → `TypeError` z instrukcją.

**13. Objaw: „`_podsumowanie()` w moim nowym raporcie nic nie robi — sprawdziłem, że jest zdefiniowane, ale efektu nie ma."**
→ *Przyczyna:* **Template Method z haczykiem wywoływanym, ale w złej kolejności** albo **haczykiem nadpisanym przez pomyłkę inną nazwą**. Częstszy przypadek drugi: `def _podsumowania(self)` (literówka), więc Python tworzy **nową** metodę, a `generuj()` woła **starą, pustą**. Nic nie wybucha, bo nic nie wybucha.
→ *Naprawa:* **`@final` na szablonie plus jawne wywołanie w `generuj()`, plus test, który sprawdza efekt.** Test: wygeneruj raport i sprawdź, że w arkuszu jest wiersz podsumowania (np. komórka z etykietą „Wartosc lacznie"). To nie test „czy metoda istnieje", a „czy efekt jest" — i tylko ten drugi coś znaczy. Diagnostyka: `print(type(raport).__dict__.keys())` — literówka będzie **widoczna** jako dodatkowa metoda.

**14. Objaw: „wpisałem w komendę `KomendaUstawKomorke("sprzedaz", 2, "Cena", 0)` i cofnąłem — komórka zrobiła się **pusta** zamiast wrócić do `0`."**
→ *Przyczyna:* **rozróżnienie „było puste" od „było zero" oparte na prawdziwości wartości.** Jeżeli cofanie robi `if self._poprzednia: kontekst.ustaw(...)`, to `0`, `""`, `False` i `None` **wszystkie** idą do gałęzi „było puste". To ta sama pułapka, co moduł 05: **puste krzesło to nie zero osób, a zero osób to nie brak krzesła.**
→ *Naprawa:* **osobne pole `_bylo: bool`**, ustawiane na podstawie **`is not None`**, nigdy na podstawie prawdziwości wartości. Zapisz to w komendzie jako komentarz, bo to jest **specyficzna** pułapka tego wzorca. Test: ustaw `0`, `False`, `""` i `None`, cofnij każdą i sprawdź, że stan przed/po jest **identyczny** — porównaj typy **i** wartości, nie tylko „prawdziwość".

## 7. Podsumowanie — model mentalny w 5 punktów

1. **Strategy nie usuwa decyzji — przenosi ją w jedno miejsce.** Wkrętarka z wymiennymi końcówkami: jedna maszyna, jedna skrzynka, **jeden** moment wyboru. Pięć typów × trzy kanały = **osiem** klas, nie piętnaście, jeżeli rozpoznasz **dwie niezależne osie** (a użycie dwóch strategii razem to Bridge — dokładnie ten, który zapowiedział moduł 24). I pamiętaj o odwrotności: przy **dwóch** wariantach w **jednym** miejscu lepszy jest `if`. Wzorzec, który nie ma uzasadnienia, jest kosztem.

2. **Template Method to formularz urzędowy: szkielet nadrukowany, treść Twoja.** `generuj()` narzuca kolejność, a podklasa dostarcza `_naglowek()`, `_wiersze()`, `_podsumowanie()`. Rozróżnienie na haczyki **wymagane** (`@abstractmethod`) i **opcjonalne** (domyślnie coś robią) jest jedyną obroną przed raportem, w którym „nikt nie zauważył, że brakuje podsumowania". Haczyk, który **milczy**, jest gorszy niż brak haczyka — bo brak haczyka nie kompiluje się, a milczący haczyk **działa źle i wygląda dobrze**.

3. **Command to obiekt-żądanie, nie efekt — i dlatego ma cofanie, audyt, makra i transakcję.** `Ctrl+Z` działa, bo każdy klawisz był obiektem. Komenda jest **jedynym** obiektem w tym kursie, który **musi** mieć stan (bo pamięta swoją odwrotność) — i dlatego jej stan zmienia **wyłącznie** `wykonaj()`, nigdy konstruktor. Twoją transakcją jest: stos komend w pamięci + **jeden** atomowy zapis na końcu. Brak zapisu = pełny rollback, i to jest najprostszy mechanizm cofania, jaki masz.

4. **Visitor: ta sama trasa, różna lista kontrolna — i nic nie przestawiaj w trakcie obchodu.** Kontroler chodzi po hali i **notuje**; przestawianie maszyn w trakcie obchodu niszczy obchód. Dlatego gość **zbiera** znaleziska albo **produkuje komendy**, a zmiany idą w **fazie drugiej** — najlepiej przez `StosKomend`, żeby dały się cofnąć. Gość może zmieniać **swój stan**, ale **nigdy** obiekt odwiedzany. Trzy pierwsze nasze goście (puste kolumny, formuły, metryki) nie zmieniają niczego — i to jest znak, że analiza jest **dobrze podzielona**.

5. **Specification to warunek jako obiekt — i najpiękniejszy most tego kursu: jedna definicja filtruje dane w Pythonie i koloruje je w Excelu.** Ale bez złudzeń: nie każdy przepis da się zapisać na tabliczce. Jeżeli `na_formule()` zwraca `None`, **mówisz wprost**, że koloru nie będzie — bo raport, w którym połowa reguł jest udawana, jest gorszy od raportu bez kolorów. Uważaj na **`not`** (nie działa — użyj `~`), na **kotwiczenie adresu** (`$E2`, nie `$E$2`), na **escapowanie** tekstu w formule (moduł 21) i na **rozjazd `float`** (`ROUND` po obu stronach). Observer i Chain of Responsibility domykają moduł: pierwszy mówi **kto ma wiedzieć**, drugi **kto decyduje** — i oba są nadmiarem, gdy nie mają uzasadnienia.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. STRATEGY - wymienna koncowka; decyzja w JEDNYM miejscu
# ==================================================================
class StrategiaFormatowania(ABC):
    nazwa: str = "?"
    @abstractmethod
    def zastosuj(self, widok: WidokArkusza) -> dict[str, object]: ...

class Zebrowe(StrategiaFormatowania):
    nazwa = "zebrowe"
    def __init__(self, kolor="F2F2F5"): self._kolor = kolor
    def zastosuj(self, widok):
        return {"strategia": self.nazwa, "pomalowane": widok.paski(1, self._kolor)}

STRATEGIE = {"zebrowe": Zebrowe(), "finansowe": Finansowe("Cena")}
strategia = STRATEGIE[spec.format]           # ← JEDEN wybor, zero if/elif

# DWA wymiary = DWIE strategie (Bridge), nie iloczyn klas:
#   5 typow x 3 kanaly => 8 klas, nie 15
# Odwrotnie: 2 warianty w 1 miejscu => zwykly `if` jest LEPSZY

# Strategia nie moze dostac `ws`. Dostaje WIDOK (port wejsciowy):
class WidokArkusza(ABC):
    def liczba_wierszy(self) -> int: ...
    def kolumny(self) -> tuple[str, ...]: ...
    def format_kolumny(self, kolumna: str, number_format: str) -> int: ...
    def paski(self, co_ile: int, kolor: str) -> int: ...
    def regula(self, kolumna: str, regula) -> str: ...


# ==================================================================
# 2. TEMPLATE METHOD - formularz urzedowy
# ==================================================================
class BazowyRaport(ABC):
    @final                       # mypy: nie nadpisuj! (interpreter to ignoruje)
    def generuj(self) -> Path:
        self._waliduj()          # krok staly
        self._przygotuj()        # krok staly
        self._naglowek()         # HACZYK WYMAGANY
        self._wiersze()          # HACZYK WYMAGANY
        self._podsumowanie()     # HACZYK OPCJONALNY
        self._formatowanie()     # HACZYK OPCJONALNY (ma sensowna domyslna!)
        self._wykresy()          # HACZYK OPCJONALNY
        return self._zapisz()    # krok staly (COMMIT)

    @abstractmethod
    def _naglowek(self) -> list[str]: ...
    @abstractmethod
    def _wiersze(self) -> Iterable[Sequence[object]]: ...

    def _podsumowanie(self) -> None:
        """Domyslnie BRAK. Ale jesli podsumowanie jest ZAWSZE wymagane ->
        @abstractmethod. Haczyk, ktory milczy, jest gorszy niz brak haczyka."""
        return None

# Test katalogowy: ktore podklasy NIE nadpisuja _podsumowanie?
#   jezeli ZADNA -> zmien na @abstractmethod.
# Test na literowke: type(raport).__dict__.keys() -> widac _podsumowania


# ==================================================================
# 3. COMMAND - zadanie jako obiekt; cofanie + audyt + makro + transakcja
# ==================================================================
class Komenda(ABC):
    @property
    @abstractmethod
    def opis(self) -> str: ...
    @abstractmethod
    def wykonaj(self, kontekst) -> None: ...
    @abstractmethod
    def cofnij(self, kontekst) -> None: ...

@dataclass
class KomendaUstawKomorke(Komenda):
    arkusz: str; wiersz: int; kolumna: str; wartosc: Any; autor: str = "system"
    _poprzednia: object = field(default=None, init=False, repr=False)
    _bylo: bool = field(default=False, init=False, repr=False)   # ← KLUCZ!

    def wykonaj(self, kontekst):
        poprzednia = kontekst.odczytaj(self.wiersz, self.kolumna)
        self._bylo = poprzednia is not None          # NIE "if poprzednia"!
        self._poprzednia = poprzednia
        kontekst.ustaw(self.wiersz, self.kolumna, self.wartosc)

    def cofnij(self, kontekst):
        if self._bylo:
            kontekst.ustaw(self.wiersz, self.kolumna, self._poprzednia)
        else:
            kontekst.usun(self.wiersz, self.kolumna)

@dataclass
class MakroKomend(Komenda):
    nazwa: str
    komendy: list[Komenda] = field(default_factory=list)   # ← default_factory!
    def wykonaj(self, kontekst):
        for k in self.komendy: k.wykonaj(kontekst)
    def cofnij(self, kontekst):
        for k in reversed(self.komendy):   # ← ODWROTNA kolejnosc = warunek poprawnosci
            k.cofnij(kontekst)

class StosKomend:
    def wykonaj(self, komenda, kontekst): ...   # + audyt
    def cofnij(self, kontekst, ile=1): ...      # + audyt
    def ponow(self, kontekst, ile=1): ...
    def znacznik(self) -> int: ...
    def do_znacznika(self, kontekst, znacznik): ...
    def audyt(self) -> tuple[dict, ...]: ...
    # transakcje:
    def rozpocznij_transakcje(self, nazwa) -> str: ...
    def zatwierdz_transakcje(self, nazwa) -> tuple[str, ...]: ...
    def cofnij_transakcje(self, kontekst, nazwa) -> tuple[str, ...]: ...
    def eksportuj_audyt(self, sciezka) -> Path: ...   # JSON, NIE arkusz

# REGULY:
# - komenda NIE zmienia nic w konstruktorze (tylko wykonaj())
# - komenda NIE trzyma `ws` (trzyma klucz arkusza + wspolrzedne)
# - COMMIT = wb.save(). Brak zapisu = pelny rollback.
# - po save() cofanie zmienia PAMEC, nie plik - udokumentuj to!
# - migawka stylu = REFERENCJE (style sa niemutowalne -> bezpieczne!)


# ==================================================================
# 4. VISITOR - stala trasa, wymienna lista kontrolna; DWIE FAZY
# ==================================================================
class OdwiedzSkoroszyt(ABC):
    def __init__(self):
        self._znaleziska: list[Znalezisko] = []
        self._komendy: list[Komenda] = []          # ← faza 2

    def odwiedz_arkusz(self, nazwa, ws) -> None: return None
    def odwiedz_komorke(self, arkusz, wiersz, kolumna, komorka) -> None: return None
    def odwiedz_tabele(self, arkusz, tabela) -> None: return None

    @final
    def przemieć(self, wb):                        # ← to TEZ Template Method
        for nazwa in wb.sheetnames:
            ws = wb[nazwa]
            self.odwiedz_arkusz(nazwa, ws)
            for wiersz in ws.iter_rows():
                for komorka in wiersz:
                    if komorka.value is not None:
                        self.odwiedz_komorke(nazwa, komorka.row,
                                             get_column_letter(komorka.column), komorka)
            for tabela in ws.tables.values():
                self.odwiedz_tabele(nazwa, tabela)
        return self

    def wynik(self) -> tuple[Znalezisko, ...]: return tuple(self._znaleziska)
    def komendy(self) -> tuple[Komenda, ...]: return tuple(self._komendy)

# Goscie: SzukajPustychKolumn / WykryjFormuly / ZbierzMetryki (CZYTAJA)
#         Anonimizuj (ZMIENIA -> produkuje KOMENDY, nie zmienia w trakcie!)
# FAZA 2: for k in gosc.konendy(): stos.wykonaj(k, kontekst)
# ZASADA: gosc moze zmieniac SWOJ stan, NIGDY obiekt odwiedzany.

# Formuly: komorka.data_type == "f" albo str(value).startswith("=")


# ==================================================================
# 5. SPECIFICATION - warunek jako obiekt; filtr + kolor z JEDNEJ definicji
# ==================================================================
class Specyfikacja(ABC):
    def jest_spelniona(self, wiersz: Mapping) -> bool: ...
    def na_formule(self, adres: str) -> str | None: ...   # BEZ znaku '='
    def __and__(self, inna): return I(self, inna)
    def __or__(self, inna):  return Lub(self, inna)
    def __invert__(self):    return Nie(self)             # `~`, NIE `not`!
    def __bool__(self):
        raise TypeError("Uzyj & | ~ , nie not/and/or")     # ← cichy blad -> glosny

spec = WiekszeNiz("Cena", 100.0) & Rowne("Region", "PL")
filtruj(dane, spec)                     # Python
spec.na_formule("$E2")                  # Excel: "AND(ROUND($E2,2)>100.0,$B2=\"PL\")"

# KOTWICZENIE: zakres "E2:E100" -> adres "$E2" (kolumna ABS, wiersz WZGL.)
#   "$E$2" = wszystkie wiersze patrza na wiersz 2  -> ZLE
# SPOJNOSC: ROUND(x,2) w Excelu i round(x,2) w Pythonie
# ESCAPOWANIE: tekst z '"' -> BladSpecyfikacji (modul 21!), nie wstrzykuj
# NIEPRZETLUMACZALNE: na_formule() -> None (np. lista VIP z CRM)
#   => filtr dziala, koloru NIE MA - powiedz to wprost, loguj ostrzezenie

class RolaKoloru:                        # jedyne miejsce znajace openpyxl
    def zbuduj(self, formula: str) -> FormulaRule:
        return FormulaRule(formula=[formula],          # ← lista, BEZ '='
                           fill=PatternFill(fill_type="solid", bgColor=self.tlo),
                           font=Font(color=self.kolor_tekstu, bold=self.pogrubienie))
        # UWAGA: w dxf liczy sie bgColor, w stylu komorki fgColor (odwrotnie!)


# ==================================================================
# 6. OBSERVER - zdarzenie jako obiekt; blad handlera ZEBRANY, nie polkniety
# ==================================================================
@dataclass(frozen=True)
class Zdarzenie:
    typ: str
    payload: dict

class Autobus:
    def __init__(self):
        self._handlery: dict[str, list[Callable]] = {}
        self.bledy: list[str] = []

    def subskrybuj(self, typ, handler) -> Callable[[], None]:
        self._handlery.setdefault(typ, []).append(handler)
        return lambda: self._handlery[typ].remove(handler)

    def publikuj(self, zdarzenie):
        for handler in tuple(self._handlery.get(zdarzenie.typ, ())):  # ← tuple!
            try:
                handler(zdarzenie)
            except Exception as blad:        # ← nie polykamy: ZBIERAMY
                self.bledy.append(f"{handler!r}: {type(blad).__name__}: {blad}")

# Roznica wobec Decoratora: Decorator PRZEPUSZCZA wyjatek dalej.
# Observer ZBIERA wyjatek i IDZIE DALEJ. Rozne role - rozne reguly.
# Typy zdarzen: POSTEP, OSTRZEZENIE, UTRATA_FUNKCJI (sparkline'y - modul 18)


# ==================================================================
# 7. CHAIN OF RESPONSIBILITY - ogniwo moze ODRZUCIC
# ==================================================================
@dataclass
class WynikOgniwa:
    wiersz: dict | None
    odrzucony_przez: str | None = None
    powod: str | None = None

class Ogniwo(ABC):
    def ustaw_nastepne(self, nastepne) -> "Ogniwo":
        self._nastepne = nastepne
        return nastepne                      # a.ustaw(b).ustaw(c)

    def obsłuż(self, wiersz):
        wynik = self.przetworz(dict(wiersz))     # ← kopia! oryginal zostaje
        if wynik.wiersz is None:
            return wynik                     # STOP - odrzucony, POWOD w wyniku
        if self._nastepne is None:
            return wynik                     # ← OSTATNIE ogniwo MUSI istniec!
        return self._nastepne.obsłuż(wynik.wiersz)

# Lancuch: OczyscTekst -> WalidujObecnosc -> NormalizujRegion -> KonwersjaTypow
# JEZELI lancuch NIGDY nie przerywa -> uzyj listy funkcji:
#     for krok in kroki: wiersz = krok(wiersz)
# Roznica wobec Decoratora: CoR MOZE zatrzymac; Decorator NIE.


# ==================================================================
# 8. TABELA DECYZYJNA + PULAPKI
# ==================================================================
# ten sam if w 3 funkcjach          -> Strategy + rejestr
# warianty roznia sie 1 parametrem  -> konfiguracja, NIE wzorzec
# 5 typow x 3 kanaly                -> DWIE strategie (Bridge), 8 klas
# stala kolejnosc, rozna tresc      -> Template Method
# trzeba cofnac / wiedziec kto      -> Command + StosKomend
# wiele zmian jako jedna operacja   -> MakroKomend (Composite komend)
# ta sama analiza, wiele arkuszy    -> Visitor
# analiza ma tez zmieniac           -> Visitor (faza 1) + Command (faza 2)
# warunek filtruje I koloruje       -> Specification
# dlugie zadanie, postep            -> Observer
# wiersz moze odpadac po drodze     -> Chain of Responsibility
# kroki zawsze sie wykonuja         -> lista funkcji, nie CoR

# PULAPKI:
# - not spec -> TypeError z __bool__; uzywaj ~ | &
# - FormulaRule: formula BEZ '=', lista; adres $E2 (nie $E$2); bgColor (nie fgColor)
# - Template Method: haczyk bez @abstractmethod + milczy = raport niekompletny
# - Command w konstruktorze -> zlewa "utworz" z "wykonaj"; stan TYLKO w wykonaj()
# - cofanie: `if poprzednia` zamiast `is not None` -> 0/False/"" staja sie puste
# - MakroKomend.cofnij: reversed() - inaczej cofasz w zlej kolejnosci
# - field(default_factory=list) - inaczej wszystkie makra dziela jedna liste
# - Visitor modyfikujacy w trakcie iteracji -> gubione komorki; DWIE FAZY
# - anonimizacja niedeterministyczna -> jedna osoba = wiele pseudonimow
# - CoR bez ostatniego ogniwa -> cichy None (nie odrozniasz przepuszczony/odrzucony)
# - Autobus bez try/except -> jeden zepsuty handler zabija caly raport
# - Autobus bez tuple(...) -> odsubskrybowanie w trakcie gubi kolejnego handlera
# - insert_rows NIE przesuwa formul (modul 07) - udokumentuj ograniczenie cofania
# - merge_cells gubi wartosci poza lewa-gorna - sprawdz PRZED (force=True)
# - Strategy z jednym wariantem / dwiema identycznymi metodami -> zwin do 1 klasy
# - Specification z I/O w jest_spelniona -> niedeterminizm i nieprzewidywalny czas
```

## 9. Co dalej

Zatrzymaj się i spójrz, ile **zachowania** przeniosłeś do obiektów. To jest najtrudniejszy moduł kursu, bo wzorce behawioralne nie mają **kształtu** — nie zobaczysz ich na diagramie tak jak drzewa sekcji z modułu 24. Trzeba je **przetestować**, żeby uwierzyć.

Masz **Strategy** z dwiema osiami, rejestrem i widokiem arkusza, który **nie pozwala** strategii sięgnąć po `ws` — bo tej metody tam nie ma. Masz **Template Method** z haczykiem wymaganym (ABC wymusi) i opcjonalnym (domyślnie coś robi), z `@final` na szablonie i z uczciwą wiedzą, że `@final` to deklaracja dla `mypy`, nie dla interpretera. Masz **Command** z cofaniem, ponawianiem, makrem (Composite), transakcjami i logiem, który leci do **JSON-a, nie do arkusza** — bo dziennik musi istnieć, gdy zapis raportu się nie uda. Masz **Visitor** w dwóch fazach, gdzie gość zmieniający dane **produkuje komendy**, a nie zmienia w trakcie obchodu. Masz **Specification**, która jest **jedną** definicją dla filtra w Pythonie i koloru w Excelu — z jawnym `None`, gdy Excel tego nie umie, i `__bool__`, który zamienia cichy błąd `not spec` w głośny `TypeError`. I masz **Observer** oraz **Chain of Responsibility** z uczciwą granicą: pierwszy zbiera błędy handlerów, ale nie przerywa; drugi wymaga ostatniego ogniwa, żeby nie mylić „przepuszczone" z „odrzucone".

W module 26 wychodzimy **ponad** wzorce — do architektury. I to, czego dziś nie dokończyliśmy, jest tam **głównym tematem**:

1. **Logika biznesowa w podklasach raportu.** W `RaportStanuMagazynu` liczymy „stan poniżej minimum" **wewnątrz** `_podsumowanie()`. W `RaportNaleznosci` liczymy „dni po terminie" w `_wiersze()`. **To jest dług do spłaty** i moduł 26 pokaże, jak go spłacić: te reguły idą do **domeny**, a podklasa tylko je woła. Kryterium: `grep -c "openpyxl"` w warstwie domenowej musi dać **zero**.
2. **`Kontekst` z przykładu 3 sięga do `raport._ws(...)`.** To jest **jedno** miejsce w całym module, gdzie prywatny akcesor fasady jest używany — i jest to miejsce uzasadnione (Kontekst **jest** infrastrukturą). Moduł 26 pokaże, jak **zamknąć** i tę ostatnią furtkę: `Kontekst` dostanie własne metody w fasadzie, a `_ws` zniknie nawet z niego.

A oprócz tego moduł 26 da Ci to, czego brakuje całej części IV — **odpowiedź na pytanie „gdzie to postawić"**:

- **Ports & Adapters.** Port `EksporterRaportu` (`Protocol`) i adaptery: `EksporterXlsxOpenpyxl`, `EksporterCsv`, `EksporterFake` do testów. Zobaczysz, że **cały `BazowyRaport` z tego modułu da się przetestować bez openpyxl** — a to jest ostateczny dowód, że warstwy są dobrze rozdzielone.
- **Repository i Unit of Work dla skoroszytu.** Dyscyplina wokół pytania „**kiedy** plik jest przepisywany" (bo openpyxl nie edytuje w miejscu — moduł 03, moduł 18). Transakcje ze stosu komend staną się elementem architektury, nie ciekawostką.
- **Konfiguracja deklaratywna.** `RaportSpec` w YAML-u opisujący kolumny, formaty, warunki i tytuły — i **zasada graniczna**: konfiguracja opisuje **co**, kod opisuje **jak**. Zobaczysz, kiedy deklaratywność zamienia się w **nowy język programowania** i dlaczego trzeba umieć się zatrzymać.
- **Wersjonowanie, idempotencja i determinizm.** Arkusz `_meta`, znaczniki czasu, sortowanie stabilne, brak `set` w kolejności danych. To wszystko jest w Twoim kodzie z modułu 25 — wystarczy nazwać i uporządkować.
- **Wydajność i bezpieczeństwo jako decyzje architektoniczne.** Raport synchroniczny w `request/response` vs zadanie w tle vs plik na dysku; `StreamingResponse` z `do_bajtow()` (masz go!), sanitacja na wejściu do domeny, nie przy zapisie.

I ostatnia rzecz, którą moduł 25 wnosi do modułu 26 i którą chcę, żebyś zapamiętał ponad wszystko, bo dotyczy **każdego** wzorca z tego modułu:

> **Wzorzec jest decyzją, nie ozdobą.** Wkrętarka z wymiennymi końcówkami nie jest lepsza od zwykłego wkrętaka, gdy wkręcasz **jedną** śrubę. Formularz urzędowy nie pomaga, gdy nie ma urzędu. `Ctrl+Z` jest bez sensu, gdy edytor nie edytuje. Strategy z jednym wariantem, Template Method z milczącym haczykiem, Command bez `cofnij`, Visitor który nie ma nic do odwiedzenia, Specification, która robi zapytanie do bazy — to nie wzorce, to **koszty bez zysku**. I dlatego w tym module powtarzają się dwie rzeczy: **tabela decyzyjna** (żeby wiedzieć, **kiedy**) i **testy** (żeby wiedzieć, **czy działa**).

Zanim pójdziesz do modułu 26, zrób jedną rzecz na **swoim** kodzie — nie na ćwiczeniach:

Weź funkcję, która zawiera najwięcej `if`-ów, i policz je. Potem policz, **ile z nich porównuje ten sam zestaw wartości w różnych miejscach**. Jeżeli odpowiedź jest większa niż jeden — masz kandydata na strategię. Jeżeli nie — **nie rób nic**. Moduł 25 nie jest po to, żeby zamienić trzy `if`-y na trzydzieści linii klas. Jest po to, żebyś wiedział, **kiedy** trzydzieści linii klas jest tańsze niż trzeci `if` — a to wiesz tylko wtedy, gdy policzysz.