# Moduł 22 — Wzorce: dlaczego abstrakcja ma sens (i kiedy nie ma)

> **Część:** IV — Wzorce projektowe i architektura · **Poziom:** ⭐⭐ · **Wymaga:** modułów 00–21

## 0. W tym module nauczysz się

- **Rozpoznasz sześć sygnałów ostrzegawczych** w kodzie automatyzującym Excela — zobaczysz je w prawdziwym skrypcie i nauczysz się je nazywać, bo bez nazwy nie da się o nich rozmawiać z zespołem.
- **Zrozumiesz, dlaczego skrypt „na kolanie" z 15 linii zamienia się w 900 linii, których nikt nie chce dotknąć** — i że to nie jest wina autora, a brakującej struktury.
- **Opanujesz tabelę czterech poziomów abstrakcji** i — co ważniejsze — **kryteria decyzji „kiedy przejść wyżej"**, oparte na *następnej zmianie*, a nie na wyobrażeniach o przyszłości.
- **Zrozumiesz YAGNI i regułę trzech** — ten moduł **nie namawia Cię do abstrakcji**. Namawia Cię do abstrakcji *wtedy, gdy jest na nią dowód*.
- **Poznasz fundament wszystkich wzorców z modułów 23–26: podział na trzy światy** — dane, prezentację i zapis — oraz **trzy pytania**, które rozstrzygają, do którego świata należy sporny fragment kodu.
- **Zobaczysz, jak zasady SOLID przekładają się na konkretny plik `.xlsx`** — zwłaszcza SRP („jeden arkusz = jedno zadanie") i OCP („nowy raport bez dotykania starego kodu").
- **Zobaczysz ten sam raport sprzedaży na czterech poziomach abstrakcji** — od 60-liniowego skryptu z usterkami do mini-architektury z portem i adapterem — i zmierzysz, ile każdy poziom kosztuje.

## 1. Intuicja i analogia

### 1.1. Historia pewnego skryptu

To zdarzyło się naprawdę, w niejednej firmie. Wersje mogą się różnić, finał — nie.

**Wersja 1 (15 linii).** Ania ma przygotować raport sprzedaży do Excela. Pisze skrypt w piętnaście linii: `Workbook()`, pętla po liście słowników, `save()`. Działa. Ania jest zadowolona, szef jest zadowolony. **Raport, który zajmował pół godziny w Excelu, powstaje w pół sekundy.**

**Wersja 2 (34 linie).** „Aniu, a dodasz kolumnę z regionem?" Dodaje. Cztery linie. Wciąż łatwe.

**Wersja 3 (71 linii).** „A można nagłówek na żółto, tak jak w firmowych raportach?" Dodaje `Font`, `PatternFill`, `Alignment`. Kolor wpisuje jako `"FFC000"` — bo skąd ma wiedzieć, że to nazwa firmowego żółtego? Wpisuje tak, jak pokazał tutorial.

**Wersja 4 (128 linii).** Ten sam raport idzie do działu niemieckiego. „Tylko zmień nagłówek na `Kunde` i walutę na euro." Ania kopiuje cztery linie, zmienia tekst. Robi się drugi skrypt. Potem trzeci, dla kontrolingu — „no i jeszcze sumę na dole".

**Wersja 5 (210 linii).** Marek z zespołu dokłada obsługę pliku wejściowego z CSV. Żeby „nie psuć tego, co działa", wkleja fragment Ani i go modyfikuje. W kodzie są teraz dwie kopie pętli formatującej, w których kolor `"FFC000"` występuje już w ośmiu miejscach.

**Wersja 6 (330 linii).** Ktoś zauważa, że raport ma sumę w kolumnie `7`, choć tabela kończy się na `F`. Poprawia, ale nie wie, że to samo `7` występuje w jeszcze trzech miejscach — w jednym z nich **akurat** `7` jest poprawne (bo to indeks innej rzeczy). Więc poprawia dwa, psuje jedno. Nikt tego nie widzi przez trzy tygodnie, bo kwota i tak wygląda „jakoś dziwnie tylko w styczniu".

**Wersja 7 (480 linii).** Dochodzi logika biznesowa: „jeżeli cena powyżej 30, to na czerwono". Powstaje `if` **w środku pętli zapisującej do komórki**. Teraz nie da się zmienić progu bez czytania kodu, który zapisuje komórki. A przetestować progu nie da się **w ogóle**, bo nie ma funkcji — jest skrypt, który przy uruchomieniu od razu pisze plik.

**Wersja 8 (650 linii).** Kolega z modułu 21 dopisuje bezpieczne wpisywanie danych od użytkownika: `cell.data_type = "s"` tam, gdzie trzeba. Wpisuje to w dziewięciu miejscach. Jest zadowolony. **Dwa tygodnie później ktoś dodaje dziesiąty endpoint i nie wie, że trzeba** — bo nigdzie nie ma jednego miejsca, w którym zapisuje się dane.

**Wersja 9 (900 linii).** Gdybyś zapytał Marka i Anię, gdzie w tym pliku zmienia się nagłówek kolumny „Klient", odpowiedzą: „yyy… chyba w pętli, ale nie wiem, czy w tej pierwszej". I to jest moment, w którym refaktoryzacja przestaje być przyjemnością, a staje się projektem na dwa dni.

| Wersja | Linie | Co doszło | Co to kosztowało |
|---|---|---|---|
| 1 | 15 | raport działa | nic |
| 3 | 71 | kolory | 4 literały bez nazw |
| 4 | 128 | drugi odbiorca | duplikacja całego pliku |
| 6 | 330 | fix sumy | magiczne `7` w 4 miejscach |
| 7 | 480 | reguła biznesowa | kod niemożliwy do przetestowania |
| 8 | 650 | bezpieczeństwo | `data_type = "s"` w 9 miejscach |
| 9 | 900 | — | **nikt nie potrafi zmienić nagłówka** |

Zwróć uwagę na kształt tej historii. **Żadna z tych decyzji nie była głupia.** Każda z nich była **rozsądna w momencie, w którym zapadała**. Kolor „na sztywno" w wersji 3 jest najszybszym możliwym rozwiązaniem i gdyby skrypt został przy jednym raporcie, byłby **poprawny**. Magiczne `7` w wersji 6 naprawiło czyjś błąd w trzy minuty. To nie jest opowieść o niekompetencji. **To opowieść o braku struktury, która nie boli, dopóki system jest mały.**

### 1.2. Dlaczego nikt nie potrafi zmienić nagłówka

Zadaj sobie pytanie, co właściwie jest tu problemem. Przecież zmiana nagłówka to **zmiana jednego stringa**. Technicznie: trzy sekundy.

Problem nie jest techniczny. Problem polega na tym, że żeby wykonać tę zmianę **bezpiecznie**, musisz najpierw odpowiedzieć na pięć pytań:

1. **Czy ten nagłówek występuje w jednym miejscu, czy w sześciu?** Musisz przeczytać cały plik, żeby to wiedzieć.
2. **Czy zmiana nagłówka wpływa na coś jeszcze?** Może od tego tekstu zależy szerokość kolumny, która jest policzona „na oko" i wyrównana do długości napisu?
3. **Czy jak to zmienię, to nie zepsuję raportu dla działu niemieckiego**, który ma własną kopię tego kodu?
4. **Czy to na pewno bezpieczne?** Nie masz testów, więc „na pewno" znaczy „mam nadzieję".
5. **Czy dostanę feedback, jeśli coś zepsuję?** Nie — bo nie ma testów, a błąd w raporcie wykryje dział księgowości po dwóch tygodniach.

Koszt zmiany nagłówka to nie trzy sekundy. Koszt to **czytanie 900 linii plus niepewność**. I to jest dokładnie to, o czym mówił moduł 20:

> **Kod, którego nie umiesz przetestować, to kod, którego nie umiesz zmienić. Kod, którego nie umiesz zmienić, to kod, który umiera — bo nikt nie chce go dotknąć.**

### 1.3. Rusztowanie i oznaczone pudełka

Masz teraz dwie rzeczy do zrozumienia i potrzebujesz do nich dwóch obrazów. Zacznę od pierwszego.

**Skrypt „na kolanie" to prowizoryczne rusztowanie.** Wyobraź sobie remont łazienki. Wystawiasz rusztowanie: cztery deski, dwie drabiny, przywiązane sznurkiem. **Na jeden dzień to jest genialne rozwiązanie.** Przechodzisz na drugą stronę, malujesz sufit, schodzisz, rozbierasz. Koszt: dwadzieścia minut. Gdybyś miał budować „profesjonalne rusztowanie modułowe" na jednodniową robotę, byłbyś głupcem.

Ale teraz wyobraź sobie, że to rusztowanie **stoi rok**. Deski są już zmurszałe, sznurek przetarł się w trzech miejscach, ktoś dostawił jeszcze jedną deskę „na chwilę" i przybił ją gwoździem, a na drugiej postawił wiadro z farbą, które kapie od maja. I przychodzi ktoś i mówi: „tylko przesuń tę deskę o pół metra". A Ty nawet nie wiesz, **które elementy są ze sobą połączone**, bo nikt tego nie rysował.

Twoje 900 linii to to rusztowanie po roku. I teraz najważniejsze zdanie tego modułu:

> **Wzorce projektowe to nie biurokracja. To oznaczone pudełka na narzędzia.**

Wyobraź sobie warsztat, w którym wszystkie narzędzia leżą w jednej wielkiej skrzyni. Szukasz klucza nasadowego 13 — **przesypujesz wszystko i nasłuchujesz**, bo metaliczny dźwięk podpowie, że coś znalazłeś. To działa! Znajdziesz ten klucz. Ale ile to zajęło? I co się stanie, kiedy przyjdzie klient i poprosi Cię o klucz, którego **w skrzyni nie ma**, a Ty nie wiesz, czy go nie ma, czy źle szukasz?

Warsztat z oznaczonymi pudełkami ma inne właściwości:
- **nie szukasz nasłuchując** — patrzysz na etykietę,
- **widzisz, czego nie ma** — puste miejsce ma etykietę, więc puste miejsce to **informacja**,
- **możesz podać pudełko komuś innemu** — bo nie musisz tłumaczyć, jak szukać,
- **możesz wymienić narzędzia w pudełku** — nie ruszając reszty warsztatu.

I to jest cała definicja wzorca, podana bez żargonu:

> **Wzorzec to nazwane, wypróbowane pudełko na konkretny rodzaj problemu — z etykietą, którą rozumie cały zespół.**

Bo widzisz, w tej historii z 900 liniami jest jeszcze jeden element, którego nie omówiliśmy. Marek i Ania **wiedzą**, że kod jest zły. Ale nie potrafią o tym rozmawiać. Bo jak nazwiesz problem? „Kod jest taki… no wiesz… rozjechany"? Nazwanie problemu to pierwszy krok do jego rozwiązania — i to jest główny powód, dla którego ten kurs wchodzi w abstrakcje. **Nazwa pozwala myśleć.**

### 1.4. Czego ten moduł NIE mówi

Zanim pójdziemy dalej, muszę Cię zatrzymać, bo to jest uczciwe i chroni Cię przed realnym niebezpieczeństwem.

> **Ten moduł nie uczy żadnego wzorca.**

Nie zobaczysz tu Buildera ani Strategy. Nie zobaczysz `Facade`, `Adapter`, `Command`. Nie napiszesz ani jednej klasy abstrakcyjnej „bo tak wypada". To jest **świadome**.

Powód jest prosty: **wzorce są narzędziem, a ten moduł ma Cię nauczyć, kiedy po nie sięgnąć.** Jeżeli nauczę Cię Buildera w tym modułce, użyjesz go do raportu z dziesięcioma wierszami i trzema kolumnami — a to jest dokładnie ten sam błąd, co niesienie aparatu fotograficznego na zakupy, żeby „było profesjonalnie". **Zrobiłbyś więcej szkody niż pożytek**, bo nadmiarowa abstrakcja jest realnym problemem, nie hipotetycznym.

W tym module nauczysz się więc **diagnozy**, nie terapii:
- jak **rozpoznać**, że struktura zaczyna się sypać,
- jak **zmierzyć**, ile Cię to kosztuje,
- jak **zdecydować**, czy już warto podnieść poziom abstrakcji, czy jeszcze nie,
- i jak to zrobić **tak, żeby nie zepsuć działającego raportu**.

## 2. Teoria

### 2.1. Sześć sygnałów ostrzegawczych

Masz listę. Wypisz ją sobie na kartce i przyklej nad biurkiem, bo te sześć rzeczy będzie się powtarzać w każdym projekcie automatyzacji Excela, jaki kiedykolwiek napiszesz.

**Sygnał 1 — powtórzone importy i powtórzone bloki formatowania.**

Dokładnie to z historii Ani: `from openpyxl.styles import Font` pojawia się w trzydziestu miejscach. Nie trzydzieści razy w jednym pliku (choć i tak bywa), a w **każdej funkcji, która coś formatuje** — bo ktoś kopiuje funkcję i zapomina, że import jest na górze.

```python
def naglowek_sprzedazy(ws):
    from openpyxl.styles import Font          # ← sygnał 1
    ...

def naglowek_zakupow(ws):
    from openpyxl.styles import Font          # ← sygnał 1
    ...
```

**Dlaczego to problem:** import wewnątrz funkcji nie jest sam w sobie błędem (czasem jest celowy — leniwe ładowanie). Problemem jest to, że jest **kopią**. Kopia importu znaczy: „skopiowałem całą funkcję". A kopia funkcji znaczy: **dwie wersje tej samej logiki, które zaczną się rozjeżdżać.**

**Sygnał 2 — literały bez nazw: kolory i progi.**

`"FFC000"` w czterdziestu miejscach. `30` jako próg alertu. `14` jako rozmiar czcionki. `0.23` jako VAT.

**Dlaczego to problem:** literał nie niesie **znaczenia**. `"FFC000"` to nie „firmowy żółty" — to ciąg sześciu znaków. Kiedy dział marketingu zmieni identyfikację wizualną, Ty nie masz czego wyszukać. Szukasz `"FFC000"`, ale **nie wiesz, czy każdy `"FFC000"` to firmowy żółty** — może być czyimś przypadkowym wyborem. I masz czterdzieści decyzji do podjęcia zamiast jednej.

Z progiem `30` jest gorzej. Bo próg to **reguła biznesowa**:
- w wersji 7 jest `30` w warunku kolorowania,
- w wersji 9 ktoś dodaje podsumowanie „ile transakcji powyżej 30" i wpisuje `30` drugi raz,
- miesiąc później dział sprzedaży mówi: „podnieśliśmy próg do 35" i ktoś zmienia jeden z nich.

**Efekt: kolor mówi jedno, a podsumowanie drugie.** I to jest błąd, którego w ogóle nie widać w kodzie — widać go w Excelu, u klienta.

**Sygnał 3 — magiczne liczby w indeksach.**

`ws.cell(row=4, column=7, ...)`. Skąd `4`? Skąd `7`?

To jest ten sam problem co `"FFC000"`, ale **groźniejszy**, bo dotyczy współrzędnych, a współrzędne mają zwyczaj się przesuwać. W historii Ani kolumna `7` była wynikiem błędu — tabela kończyła się na `F` (czyli 6), a suma wylądowała **za tabelą**. I nikt tego nie zauważył, bo w Excelu po prostu „suma wisi z boku" i wygląda prawie normalnie.

`openpyxl` daje Ci narzędzie do nazwania współrzędnej:

```python
from openpyxl.utils import get_column_letter, column_index_from_string

get_column_letter(7)                 # 'G'  — teraz WIESZ, ze to G
column_index_from_string("G")        # 7    — i odwrotnie
```

Ale **prawdziwym** rozwiązaniem nie jest `get_column_letter(7)`, tylko **brak liczby 7 w kodzie**. Kolumna ma wynikać z pozycji w definicji tabeli (zobaczysz to w przykładzie 2).

**Sygnał 4 — logika biznesowa przemieszana z zapisem.**

To najkosztowniejszy z sześciu sygnałów i najtrudniejszy do naprawienia później.

```python
for rekord in DANE:
    ws.cell(row=wiersz, column=1, value=rekord["data"])
    ws.cell(row=wiersz, column=6, value=rekord["cena"])

    if rekord["cena"] > 30:                        # ← reguła biznesowa
        ws.cell(row=wiersz, column=6).font = Font(color="FF0000")   # ← prezentacja
                                                                    # ↑ razem z ZAPISEM
    wiersz += 1
```

Popatrz, ile decyzji jest w tych pięciu liniach: **jaka jest reguła** (`> 30`), **jak ją pokazać** (czerwona czcionka), i **gdzie to trafi** (`column=6`). Trzy różne światy w jednym `if`-ie.

**Konsekwencja praktyczna:** nie da się odpowiedzieć na pytanie „które transakcje są alertami?" bez **wygenerowania pliku i obejrzenia go**. A przecież to pytanie **nie jest o plik** — to pytanie o dane.

**Sygnał 5 — brak możliwości testowania.**

Ten sygnał jest cichy, bo **nie widać go w kodzie**. Widać go dopiero, gdy próbujesz napisać test:

```python
# Jak przetestowac, czy regula "> 30" dziala poprawnie?
# Nie da sie - kod jest na poziomie modulu (top-level), wiec import
# tego pliku OD RAZU generuje raport i pisze plik na dysku.
import nasz_skrypt        # ← i juz masz plik w output/
```

W module 20 zbudowałeś narzędzie do testowania bez plików (`BytesIO`). Ale żeby z niego skorzystać, musisz mieć **co wywołać** — funkcję, która przyjmuje dane i zwraca skoroszyt. Kod, który jest wyłącznie ciągiem instrukcji na poziomie modułu, jest **z definicji** nietestowalny.

**Sygnał 6 — brak możliwości zmiany formatu bez czytania całości.**

To jest suma pozostałych pięciu i jednocześnie **jedyny sygnał, który zauważy Twój szef**.

Usiądź i wyobraź sobie, że chcesz zmienić kolor nagłówka. Musisz:
- znaleźć wszystkie miejsca, w których powstaje nagłówek (pięć funkcji? sześć kopii?),
- sprawdzić, czy któryś nie ma dodatkowego warunku (`if region == "DE"`),
- sprawdzić, czy zmiana koloru nie wpłynie na coś innego (np. na styl nazwany, od którego zależą inne komórki),
- nie mieć pewności, że czegoś nie przegapiłeś.

Jeśli **jakakolwiek** zmiana wyglądu wymaga przeczytania całości, to znaczy, że **prezentacja nie jest oddzielona od reszty**. I to jest dokładnie to, co naprawia podział na trzy światy — ale nie od razu o tym.

**Tabela zbiorcza:** do czego prowadzi każdy sygnał i gdzie kurs go naprawia.

| # | Sygnał | Co Cię kosztuje | Gdzie to naprawiamy |
|---|---|---|---|
| 1 | powtórzone importy i bloki | zmiana koloru = 30 edycji; jedną pominiesz | 23 (Registry, Flyweight) |
| 2 | literały bez nazw | brak znaczenia; rozjazd wartości w dwóch miejscach | 22 (konfiguracja), 23 |
| 3 | magiczne indeksy | off-by-one; kolumna poza tabelą | 22 (`Kolumna`, brak indeksów) |
| 4 | logika w zapisie | reguła nie do przetestowania; próg w dwóch miejscach | 22 (światy), 25 (Specification) |
| 5 | brak testowalności | brak odwagi do zmian (moduł 20) | 20 (`BytesIO`), 22 (funkcje) |
| 6 | brak lokalnej zmiany | 900 linii czytane, żeby zmienić kolor | 22 (światy), 24 (Facade) |

### 2.2. Trzy poziomy abstrakcji (i jedna tabela, którą warto wydrukować)

To jest najważniejsza tabela tego modułu i prawdopodobnie najważniejsza tabela całej Części IV. Wróć do niej za miesiąc, kiedy będziesz się zastanawiać, „czy już robić z tego klasy".

| Poziom | Kiedy wystarcza (sygnał) | Co dostajesz | Co płacisz | Kiedy przejść wyżej |
|---|---|---|---|---|
| **0 — skrypt liniowy** | jednorazowy raport, 1 plik, 1 odbiorca, uruchomisz go raz albo dwa | szybkość, zero ceremonii, zero pojęć do nauczenia | zmiana czegokolwiek = czytanie całości; brak testów; jeden błąd = jeden dzień | gdy uruchomisz go **drugi raz** albo pojawi się **drugi odbiorca** |
| **1 — funkcje + konfiguracja** | powtarzalny raport; zmieniają się kolumny, kolory, progi; chcesz mieć testy | trzy światy rozdzielone na funkcje; konfiguracja w jednym miejscu; funkcje czyste testowalne bez plików | trzeba pilnować granic (inaczej funkcje znowu się sklejają) | gdy masz **trzy warianty** tego samego raportu albo **2+ odbiorców** |
| **2 — klasy, specyfikacje, strategie** | wiele raportów tym samym kodem; wymagania „to samo, ale trochę inaczej" | nowy raport **bez zmiany istniejącego kodu** (OCP); wymienne warianty; konfiguracja jako obiekt | abstrakcja, której trzeba się nauczyć; **testy stają się obowiązkowe**, bo inaczej abstrakcja szkodzi | gdy pojawia się **integracja**, **wiele formatów wyjścia** albo **wymóg audytu** |
| **3 — warstwy, porty i adaptery** | produkt, zespół, wersjonowanie, audyt, wymiana formatu wyjścia (Excel → PDF → CSV) | testy **całej logiki bez plików**; wymiana `openpyxl` nie dotyka logiki biznesowej; konfiguracja deklaratywna | wysoki koszt wejścia; potrzebna dyscyplina zespołowa; **przesada tu boli najbardziej** | to jest szczyt dla raportów; dalej idzie się tylko po to, żeby obsłużyć *prawdziwy* system, nie „na wszelki wypadek" |

**Cztery zdania komentarza do tej tabeli**, bo tabela bez komentarza staje się dogmatem.

1. **Poziomy nie są stopniami wtajemniczenia.** Nie ma nic „nieprofesjonalnego" w poziomie 0. Poziom 0 to **jedyne właściwe rozwiązanie** dla zadania z tabeli: jednorazowy raport dla jednego odbiorcy. Zbudowanie na to poziomu 2 to nie rozwój, a **marnotrawstwo**.

2. **Przejście jest wywołane zdarzeniem, nie ambicją.** Każda komórka w kolumnie „Kiedy przejść wyżej" opisuje **fakt, który się stał** („drugi raz", „trzy warianty", „integracja"), a nie przypuszczenie („może kiedyś będzie skalować"). To jest cała istota YAGNI z sekcji 2.3.

3. **Kolumna „co płacisz" jest prawdziwa.** Poziom 2 i 3 wymagają **testów**. Jeżeli nie masz testów, a podniesiesz poziom abstrakcji, **będziesz w gorszej sytuacji niż na poziomie 0** — bo teraz nie wiesz, gdzie jest błąd, a kod jest rozbity na więcej plików.

4. **Można i trzeba schodzić.** Jeżeli raport, który napisałeś na poziomie 2, okaże się jednorazowy (odbiorca zniknął, raport nie jest już potrzebny), to **usunięcie abstrakcji jest poprawną decyzją**. Nie ma tu „dorobku", który trzeba chronić.

**Diagram decyzji.** Trzy pytania rozstrzygające:

```text
Czy ten skrypt uruchomi sie wiecej niz 2 razy?    nie → POZIOM 0
        │ tak
        v
Czy zmienialy sie juz kolumny / kolory / progi?   tak → POZIOM 1 (natychmiast)
lub: czy bedziesz chcial testy?
        │
        v
Czy istnieja 3+ warianty tego raportu,            tak → POZIOM 2
2+ odbiorcow, albo "to samo ale inaczej"?
        │
        v
Czy raport musi dawac sie wyslac w innym formacie, tak → POZIOM 3
albo istnieje integracja / audyt / zespol?
        │
        v
POZOSTAN NA POZIOMIE, KTORY MASZ. Nie awansuj "na zapas".
```

### 2.3. YAGNI, reguła trzech i asymetria kosztów

**YAGNI** — *You Aren't Gonna Need It*, „nie będziesz tego potrzebować". Zasada brzmi: **nie buduj funkcjonalności dla wymagań, których nie masz**. Brzmi banalnie, ale w praktyce jest najczęściej łamana z dwóch powodów: (a) bo „profesjonalny kod tak wygląda", (b) bo przewidywanie przyszłości jest przyjemne i daje poczucie kontroli.

Muszę być tu bardzo precyzyjny, bo YAGNI jest często nadużywane jako wymówka na bałagan:

> **YAGNI nie znaczy „nigdy nie abstrahuj". YAGNI znaczy „abstrahuj, gdy masz dowód, a nie hipotezę".**

Dowodem jest **konkretny, istniejący drugi przypadek**. I stąd druga zasada, która jest praktycznym narzędziem:

> **Reguła trzech:** pierwszy raz — napisz wprost. Drugi raz — **zauważ** duplikację i zapisz ją sobie. Trzeci raz — **wtedy** wyciągnij wspólny kod.

Dlaczego trzeci, a nie drugi? Bo przy drugim przypadku jeszcze nie wiesz, **co jest istotą wspólną, a co przypadkiem**. Trzeci przypadek pokazuje, co się naprawdę powtarza. Wyciągnięcie wspólnego kodu z dwóch przypadków to klasyczny błąd: tworzysz abstrakcję na podstawie jednego podobieństwa, a trzeci przypadek ją łamie — i wtedy musisz dodać parametr, potem `if`, potem drugi parametr… i po miesiącu masz „framework", o którym mówi pitfall z końca tego modułu.

**Teraz asymetria, o której nikt nie mówi, a która wyjaśnia, dlaczego ludzie tak chętnie przesadzają.**

Koszt **braku** abstrakcji jest **odroczony i niewidoczny**:
- nie pojawia się w dniu, w którym popełniasz błąd,
- pojawia się pół roku później, przy zmianie wymagań,
- jest rozłożony na wiele małych irytacji,
- **nikt go nigdy nie policzy**, bo nie ma momentu, w którym „coś się zepsuło".

Koszt **nadmiarowej** abstrakcji jest **natychmiastowy i widoczny**:
- czujesz go już w momencie pisania („po co ten Builder do trzech kolumn?"),
- więc **reagujesz i go usuniesz**.

Z tej asymetrii wynikają dwie praktyczne reguły, które wyznaczają balans:

**Reguła A — wybieraj najniższy poziom, który przetrwa Twoją *następną* zmianę.**

Nie „przetrwa skalowanie do 50 raportów", nie „przetrwa wejście zespołu", a **następną zmianę**, którą faktycznie przewidujesz (a nie wymyślasz). Jeśli wiesz, że za tydzień dojdzie kolumna — poziom 1 wystarczy. Jeśli wiesz, że w tym kwartale dojdzie drugi odbiorca — poziom 1 też wystarczy.

**Reguła B — abstrakcja musi być opłacona testem.**

To reguła najważniejsza i najczęściej łamana:

> **Abstrakcja bez testów jest gorsza niż duplikacja.**

Dlaczego? Bo duplikacja jest **widoczna**. Kiedy dwie kopie się rozjeżdżają, widzisz dwie różne wartości i wiesz, że coś jest nie tak. Abstrakcja **ukrywa** duplikację: skoro masz jedną funkcję, „przecież wszystko jest spójne" — i nie sprawdzasz. Jeśli ta jedna funkcja jest źle napisana, **błąd jest zmultiplikowany**, a Ty nie masz testu, który by krzyknął.

Dlatego kolejność jest zawsze taka sama:

1. Podziel kod na **funkcje czyste** (poziom 1).
2. Napisz **testy** na te funkcje (moduł 20 — `BytesIO`, brak plików).
3. Dopiero **potem** podnoś poziom abstrakcji.

Odwrotna kolejność — najpierw wzorce, potem testy — kończy się abstrakcją, której nikt nie umie zmienić, bo nikt nie wie, co ona właściwie robi.

I ostatnia rzecz z tej asymetrii, bo warto ją zobaczyć razem:

| | Brak abstrakcji | Nadmiar abstrakcji |
|---|---|---|
| Kiedy boli? | rok później | od razu |
| Jak wygląda? | „nie da się zmienić nagłówka" | „ten Builder do 3 kolumn to przesada" |
| Kto zauważy? | następna osoba, która dotknie kodu | autor |
| Czy da się naprawić? | refaktoryzacja, godziny | usunięcie kodu, minuty |
| Główny powód | brak nazw i struktur | „profesjonalnie wygląda" |

### 2.4. Fundament: trzy światy

Teraz rzecz, do której zmierza cała Część IV. Wszystkie wzorce z modułów 23–26 są **sposobami na utrzymanie tego jednego podziału**. Jeżeli zrozumiesz tę sekcję, kolejne moduły będą tylko wariantami tego samego motywu.

**Analogia: budowa domu.**

- **Architekt** decyduje, co ma być w budynku: ile pokoi, jak duże, gdzie łazienka. Rysuje **projekt** — abstrakcyjny plan, który mówi, *co* ma być.
- **Ekipa budowlana** decyduje, jak to wygląda i działa: gdzie postawić ścianę, jaką farbą pomalować, jak zamontować okno. To jest **wykończenie**.
- **Dostawca materiałów** dowozi konkretne cegły, konkretną farbę, konkretne okna. To jest **realizacja** — dostarczenie fizycznych rzeczy na plac budowy.

Zwróć uwagę na kluczową właściwość tego podziału:

> **Możesz zmienić dostawcę materiałów, nie zmieniając projektu domu.**

Jeżeli architekt powie „zmieniamy cegłę z czerwonej na kremową", to **projekt się nie zmienia**. Zmienia się tylko to, co przyjeżdża na plac. I na odwrót: jeżeli dostawca zbankrutuje i znajdziesz nowego, architekt nie musi poprawiać rysunku.

To jest cały sens. A teraz przełóżmy to na raport.

#### Świat 1: DANE — „co ma być w raporcie"

To architekt. Mieszka tu:
- model domenowy (`WierszSprzedazy` jako `dataclass`),
- reguły biznesowe („transakcja wymaga uwagi, gdy cena > prog"),
- czyszczenie i walidacja danych z modułu 21,
- agregacje („suma per region").

**Czego tu nie ma:** `import openpyxl`. Ani jednego. W tym świecie nie istnieje pojęcie komórki, koloru ani arkusza. Istnieje sprzedaż, kwota, region, data.

**Test tego świata:** `grep -r "openpyxl" src/domena/` musi zwrócić **zero wyników**. Jeżeli zwróci cokolwiek — architekt właśnie zaczął murować, i to jest początek końca podziału (zobaczysz to twardo w module 26).

#### Świat 2: PREZENTACJA — „jak to wygląda"

To ekipa wykończeniowa. Mieszka tu:
- kolory, czcionki, obramowania,
- szerokości kolumn, zamrażanie okien, ustawienia wydruku,
- decyzja, czy alert pokazać **czerwoną czcionką**, czy **regułą formatowania warunkowego** (moduł 13–14),
- układ: gdzie nagłówek, gdzie podsumowanie, gdzie wykres (moduł 15).

**Czego tu nie ma:** reguł. Świat prezentacji **nie decyduje**, co jest alertem — dostaje tę decyzję **gotową**. To jest najważniejsze zdanie tej sekcji, więc powtórzę je w innym porządku: **prezentacja nie pyta „czy cena > 30?", tylko „czy mam pokazać ten wiersz jako alert?".**

**Dlaczego to ma znaczenie:** gdyby prezentacja sama sprawdzała próg `30`, to próg żyłby w dwóch miejscach — raz w regule, raz w rysowaniu. A wtedy „podnieśliśmy próg do 35" jest zmianą w dwóch plikach i **jedno z nich zostanie pominięte**. To jest dokładnie pułapka nr 2 z sekcji 2.1, tylko przeniesiona na wyższy poziom.

#### Świat 3: ZAPIS — „jak to trafia do pliku `.xlsx`"

To dostawca materiałów. Mieszka tu:
- `Workbook`, `load_workbook`, `save`,
- `cell.value`, `number_format`, `data_type` (w tym neutralizacja z modułu 21),
- `Font`, `PatternFill` (jako **wykonanie** decyzji, nie jako decyzja),
- `ws.append`, `iter_rows`, tabele, wykresy.

**Czego tu nie ma:** żadnych decyzji. Ani `if cena > 30`, ani `if region == "DE"`, ani `if raport == "regionalny"`. Ten świat **wykonuje polecenia**. Jeżeli tu jest `if`, to znaczy, że brakuje pola w świecie danych albo w świecie prezentacji.

#### Trzy światy razem

```text
   ŚWIAT DANYCH                ŚWIAT PREZENTACJI           ŚWIAT ZAPISU
  (architekt)                 (ekipa wykonczeniowa)       (dostawca materialow)
 ┌───────────────┐   dane    ┌──────────────────┐  widok  ┌──────────────────┐
 │ WierszSprzed. │ ────────> │ "alert = czerwony"│ ──────> │ openpyxl: cell = │
 │ regula: > prog│  + decyzje│ szerokosci, tytul │  + plik │ ..., save()      │
 │ agregacje     │           │ formaty liczb     │         │                  │
 └───────────────┘           └──────────────────┘         └──────────────────┘
    NIE importuje              ZNA wyglad,               WYKONUJE,
    openpyxl                   NIE liczy regul           NIE decyduje
```

**Kierunek zależności jest jednokierunkowy:** zapis zna prezentację, prezentacja zna dane, **dane nie znają nikogo**. To znaczy, gdy w module 26 wymienisz `openpyxl` na `xlsxwriter` (albo dopiszesz eksport do PDF), **świat danych nie zostanie nawet dotknięty**.

Sprawdź, jak ta zasada działa w praktyce na Twoim kodzie z modułu 21. Pamiętasz `normalizuj()`? Funkcja zwracała `Komorka(wartosc, jako_tekst)` — **bez openpyxl, bez I/O**. To był świat danych. A `zapisz_komorke()` — który tylko *wykonywał* decyzję — to świat zapisu. **Ten podział już zbudowałeś, nie nazywając go.** Teraz ma nazwę.

### 2.5. Trzy pytania, które rozstrzygają, do którego świata należy kod

W teorii wszystko jest jasne. W praktyce stajesz przed fragmentem kodu i nie wiesz, gdzie go położyć. Na przykład: „ustawiam czerwony kolor, gdy cena > 30" — dane czy prezentacja?

Zamiast zgadywać, zadaj trzy pytania. Są one **operacyjne** — dają rozstrzygnięcie, a nie wątpliwość.

**Pytanie 1: Czy ta informacja zmieniłaby się, gdybyśmy wysyłali ten sam raport do PDF-a (albo na ekran, albo na papier)?**

- „Cena powyżej progu to transakcja wymagająca uwagi" — **nie zmieniłaby się**. Reguła jest o sprzedaży, nie o formacie. → **ŚWIAT DANYCH.**
- „Alert pokazujemy czerwoną czcionką" — **zmieniłaby się** (w PDF-ie może to być pogrubienie albo znaczek). → **ŚWIAT PREZENTACJI.**
- „Wartość ląduje w komórce `F12`" — **zniknęłaby całkowicie** (PDF nie ma komórek). → **ŚWIAT ZAPISU.**

**Pytanie 2: Czy ta informacja zmieniłaby się, gdyby dane przyszły z innego źródła (baza danych zamiast CSV)?**

- „Wiersz ma pola: data, klient, region, produkt, ilość, cena" — **nie zmieniłaby się**, to kształt danych. → **DANE.**
- „Kolumny w kolejności: Data, Klient, Region…" — **zmieniłaby się**, to układ prezentacji. → **PREZENTACJA.**

**Pytanie 3: Czy ta informacja zniknęłaby, gdyby openpyxl przestał istnieć?**

- `Font(bold=True)`, `ws.save()`, `cell.number_format` — **zniknęłaby**. → **ZAPIS.**
- „Motyw firmowy ma kolor `FFC000`" — **nie zniknęłaby** (firma nadal ma kolor). → **PREZENTACJA.**
- „Suma sprzedaży to iloczyn ilości i ceny" — **nie zniknęłaby nigdy**. → **DANE.**

**Zastosujmy to do spornego przypadku z sekcji 2.4.** „Ustawiam czerwony kolor, gdy cena > 30":

| Pytanie | Odpowiedź | Wniosek |
|---|---|---|
| Zmieni się przy PDF? | „cena > 30" — nie; „czerwony" — tak | **rozłącz na dwie rzeczy** |
| Zmieni się przy innym źródle danych? | reguła — nie | reguła → DANE |
| Zniknie bez openpyxl? | `Font(color=...)` — tak | kolor → ZAPIS |

I mamy odpowiedź, która zgadza się z intuicją z sekcji 2.4: **reguła idzie do danych, kolor do prezentacji, a przypisanie koloru do komórki — do zapisu.** Trzy fragmenty, jeden `if` w oryginale. To jest dokładnie ten `if` z sygnału 4.

### 2.6. SOLID po excelowemu

SOLID to pięć zasad projektowania obiektowego. Zwykle wykłada się je na przykładzie sklepu internetowego, co dla osoby automatyzującej Excela jest zupełnie bezużyteczne. Przełożę je więc na raport.

**S — Single Responsibility Principle (jedna odpowiedzialność).**

W wersji excelowej brzmi to tak:

> **Jeden arkusz = jedno zadanie. Jedna funkcja = jeden świat.**

Arkusz `Dane` nie robi podsumowań. Arkusz `Podsumowanie` nie zawiera surowych danych. Arkusz `Wykresy` nie trzyma danych źródłowych (bo wykres ma je brać z `Dane`, nie kopiować).

A funkcja `naglowek_sprzedazy` **nie powinna**: liczyć sum, decydować o kolorach na podstawie danych, ani zapisywać pliku. Jeżeli funkcja ma w nazwie „i" („przygotuj dane i zapisz"), to prawdopodobnie **złamanie SRP**.

Jak sprawdzisz SRP w praktyce? Zadaj pytanie: **„ile powodów ma ta funkcja, żeby się zmienić?"** Jeżeli zmiana koloru, zmiana progu i zmiana ścieżki pliku mogą wymusić edycję tej samej funkcji — to są **trzy powody**. Trzy odpowiedzialności.

**O — Open/Closed Principle (otwarte na rozszerzenia, zamknięte na modyfikacje).**

To najważniejsza i najtrudniejsza zasada dla kogoś, kto automatyzuje raporty, bo dotyka codziennego bólu:

> **Nowy typ raportu nie może wymagać edycji kodu istniejącego raportu.**

Wersja **zła** (i to jest wersja, w której żyje 90% skryptów):

```python
def generuj(raport, dane):
    if raport == "sprzedaz":
        kolumny = ["Data", "Klient", "Cena"]
    elif raport == "zakupy":                       # ← dodales elif
        kolumny = ["Data", "Dostawca", "Kwota"]
    elif raport == "regionalny":                   # ← dodales elif
        kolumny = ["Region", "Wartosc"]
    # za rok: 11 elif-ow i nikt nie wie, ktora galaz jest zywotna
```

Za każdym razem, gdy dochodzi raport, **modyfikujesz działający kod**. A modyfikacja działającego kodu bez testów to ryzyko: przy okazji dodania raportu regionalnego możesz **zepsuć** raport sprzedaży. I dowiesz się o tym od klienta.

Wersja **dobra**:

```python
SPEC_SPRZEDAZ = Specyfikacja(kolumny=..., regula_alertu=...)
SPEC_REGIONALNY = Specyfikacja(kolumny=..., regula_alertu=...)

raport = RaportXlsx(SPEC_REGIONALNY)     # ← nowy raport = nowy OBIEKT, nie nowy if
raport.generuj(dane)
```

Nie zmieniłeś **ani jednej linii** w klasie `RaportXlsx`. Dodałeś **dane** opisujące nowy raport. To jest OCP.

Masz tu **mierzalne kryterium**, które wykorzystasz w ćwiczeniu 🔴:

> **Ile linii dopisałeś do pliku istniejącego raportu? Jeżeli więcej niż zero — złamałeś OCP.** (Wyjątkiem jest poprawienie błędu w tym raporcie — to nie jest rozszerzanie.)

**L — Liskov Substitution (podstawialność).**

W exceli brzmi: jeżeli masz `BazowyRaport` z metodą `generuj()`, to każdy raport konkretny musi dać się użyć **w tym samym miejscu** i **nie wymagać niczego dodatkowego**. Jeżeli `RaportWykresowy.generuj()` wymaga, żeby najpierw **ręcznie** wywołać `_przygotuj_dane()`, to złamałeś LSP — bo nie jest zamiennikiem, tylko ma dodatkowy warunek. Objaw: „ten działa, ale tego musisz wywołać inaczej".

**I — Interface Segregation (segregacja interfejsów).**

Nie pakuj ośmiu metod w jeden port. Jeżeli port ma nazywać się `EksporterRaportu`, to niech ma **jedną** metodę `eksportuj(...)`. Port, który wymaga implementacji `otworz()`, `zamknij()`, `zwaliduj()`, `zapisz()`, `loguj()`, `wyslij_mail()`, `skonfiguruj()` i `eksportuj()` — to nie port, to **kontrakt na cały system**. Nikt go nie zaimplementuje w testach i nikt go nie wymieni.

**D — Dependency Inversion (odwrócenie zależności).**

`RaportXlsx` **nie powinien** tworzyć `Workbook()` w środku „na sztywno". Powinien **dostać** obiekt, który umie zapisać. Wtedy w testach dostaje wersję, która nic nie pisze do dysku (zobaczysz to w przykładzie 4). To jest bezpośrednie przygotowanie do modułu 26.

### 2.7. Mapowanie wzorców na problemy openpyxl

Ta tabela jest **zapowiedzią** modułów 23–25. Nie musisz jej teraz rozumieć od deski do deski — ale przeczytaj ją uważnie, bo pokazuje, że wzorce nie są „dla programistów aplikacji", tylko rozwiązują **konkretne, nazwane problemy z tego kursu**.

| Wzorzec | Problem w openpyxl, który rozwiązuje | Moduł |
|---|---|---|
| **Factory Method** | tworzenie arkuszy przez `create_sheet` + `title` + `index` + `tabColor` w 20 miejscach | 23 |
| **Builder** | raport składany w 6 krokach, gdzie kolejność i kompletność mają znaczenie | 23 |
| **Prototype** | szablon skoroszytu jako pieczątka: nadruk gotowy, dopisujesz treść, wzorzec nigdy nie jest nadpisywany | 23 |
| **Singleton / Registry** | rejestr stylów: `Font` tworzony raz, używany w tysiącach komórek | 23 |
| **Flyweight** | współdzielenie obiektów stylu — openpyxl deduplikuje w `styles.xml`, ale Python nadal alokuje | 23 |
| **Facade** | jedno wejście `RaportExcel`, które ukrywa `Workbook`, tabele, wykresy i zapis | 24 |
| **Adapter** | model domenowy → wiersz arkusza; ten sam model z CSV i z JSON | 24 |
| **Decorator** | walidacja, logowanie i pomiar czasu wokół zapisu, bez dotykania zapisu | 24 |
| **Proxy** | leniwe ładowanie skoroszytu; blokada zapisu bez zatwierdzenia | 24 |
| **Composite** | sekcje raportu, które same mogą zawierać podsekcje | 24 |
| **Strategy** | wymienne formatowanie i warianty eksportu bez `if`-ów | 25 |
| **Template Method** | `BazowyRaport.generuj()` — szkielet narzucony, treść Twoja | 25 |
| **Command** | operacje edycyjne z `undo()` i logiem audytowym | 25 |
| **Visitor** | przemiatanie arkusza: szukaj pustych kolumn, formuł, danych poufnych | 25 |
| **Specification** | reguły deklaratywne: filtrują dane **i** generują formatowanie warunkowe z tych samych obiektów | 25 |
| **Observer** | postęp, ostrzeżenia o utracie funkcji pliku przy zapisie | 25 |

Zwróć uwagę na dwie rzeczy w tej tabeli.

Po pierwsze: **każdy wzorzec odpowiada na problem, który już widziałeś w tym kursie**. „Singleton/Registry" rozwiązuje sygnał 2 z sekcji 2.1 (kolory w 40 miejscach) i to, o czym mówił moduł 09 (rejestr stylów). „Adapter" rozwiązuje pytanie z sekcji 2.5 („czy ta informacja zmieni się przy innym źródle danych?"). To nie są wzorce „z zewnątrz" — to **odpowiedzi na Twoje konkretne pytania**.

Po drugie: **Specification** zasługuje na szczególną uwagę, bo rozwiązuje **dwa problemy jednym narzędziem**. Ta sama reguła (`cena > prog`), wyrażona jako obiekt, może:
- **filtrować dane** („pokaż mi tylko transakcje wymagające uwagi") — świat danych,
- **generować regułę formatowania warunkowego** (moduł 13) — świat prezentacji.

Czyli **reguła żyje w jednym miejscu i jest wykorzystana w dwóch światach, bez duplikacji**. To jest OCP w najczystszej postaci i główny powód, dla którego ten wzorzec zamyka Część IV.

## 3. Przykłady krok po kroku

Cztery przykłady, ten sam raport sprzedaży, cztery poziomy abstrakcji. **To jest ten sam przypadek domenowy, który będzie towarzyszył Ci w modułach 23–27** — dzięki temu zobaczysz **ewolucję kodu**, a nie cztery rozłączne ciekawostki.

### Przykład 1 — skrypt „na kolanie" (poziom 0)

Ten plik jest **celowo zły**. Czytaj go nie po to, żeby się nauczyć, a po to, żeby **rozpoznać**. Każdy z sześciu sygnałów z sekcji 2.1 jest tu zaznaczony komentarzem.

```python
"""Modul 22, przyklad 1: skrypt 'na kolanie' - POZIOM 0.

Ten plik jest celowo zly. Sluzy do cwiczenia rozpoznawania sygnalow
ostrzegawczych z sekcji 2.1. NIE przepisuj go od razu na wzorce -
najpierw sprawdz, CZY WARTO (sekcja 2.3, YAGNI).

Uruchom: python examples/22_poziom0.py
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles import Font, PatternFill          # ← SYGNAL 1: duplikat


DANE = [
    {"data": date(2026, 1, 5),  "klient": "Alfa sp. z o.o.",
     "region": "PL", "produkt": "Widzet A", "ilosc": 10, "cena": 25.0},
    {"data": date(2026, 1, 9),  "klient": "Beta S.A.",
     "region": "DE", "produkt": "Widzet B", "ilosc": 4,  "cena": 41.5},
    {"data": date(2026, 1, 17), "klient": "Gamma sp. k.",
     "region": "PL", "produkt": "Widzet A", "ilosc": 22, "cena": 25.0},
    {"data": date(2026, 1, 23), "klient": "Alfa sp. z o.o.",
     "region": "PL", "produkt": "Gadzet C", "ilosc": 7,  "cena": 63.0},
]


wb = Workbook()
ws = wb.active
ws.title = "Raport"

# --- tytul ---
ws["A1"] = "Raport sprzedazy - styczen 2026"
ws["A1"].font = Font(bold=True, size=16)               # ← SYGNAL 2: literaly bez nazw
ws.merge_cells("A1:F1")

# --- naglowki ---
naglowki = ["Data", "Klient", "Region", "Produkt", "Ilosc", "Cena"]
for i, tekst in enumerate(naglowki, start=1):
    komorka = ws.cell(row=3, column=i, value=tekst)
    komorka.font = Font(bold=True, color="FFFFFF")     # ← SYGNAL 2: "FFFFFF"
    komorka.fill = PatternFill(fill_type="solid", fgColor="FFC000")   # ← SYGNAL 2
    komorka.alignment = Alignment(horizontal="center", vertical="center")

# --- dane ---
wiersz = 4                                             # ← SYGNAL 3: magiczne 4
for rekord in DANE:
    ws.cell(row=wiersz, column=1, value=rekord["data"])
    ws.cell(row=wiersz, column=2, value=rekord["klient"])
    ws.cell(row=wiersz, column=3, value=rekord["region"])
    ws.cell(row=wiersz, column=4, value=rekord["produkt"])
    ws.cell(row=wiersz, column=5, value=rekord["ilosc"])
    ws.cell(row=wiersz, column=6, value=rekord["cena"])

    # ↓ SYGNAL 4: regula biznesowa + prezentacja + zapis w jednym if-ie
    if rekord["cena"] > 30:                            # ← SYGNAL 2: prog bez nazwy
        ws.cell(row=wiersz, column=6).font = Font(color="FF0000")

    ws.cell(row=wiersz, column=1).number_format = "dd.mm.yyyy"
    wiersz += 1

# --- podsumowanie ---
suma = sum(r["ilosc"] * r["cena"] for r in DANE)
ws.cell(row=wiersz + 1, column=6, value="SUMA:").font = Font(bold=True)
ws.cell(row=wiersz + 1, column=7, value=suma)          # ← SYGNAL 3: magiczne 7 (G!)
                                                       #   tabela konczy sie na F!

# --- szerokosci ---
ws.column_dimensions["A"].width = 12                   # ← SYGNAL 6: zmiana szerokosci
ws.column_dimensions["B"].width = 24                   #   wymaga czytania calosci
ws.column_dimensions["C"].width = 8
ws.column_dimensions["D"].width = 16
ws.column_dimensions["E"].width = 8
ws.column_dimensions["F"].width = 10

# ← SYGNAL 5: zapis na poziomie modulu = nie da sie zaimportowac bez efektu
wb.save("output/22_poziom0.xlsx")
```

**Co się dzieje w pamięci.** `DANE` to lista słowników. Każdy `rekord["cena"]` to odczyt z `dict` — **Python nie ma pojęcia, że to „cena"**. Ma klucz `"cena"`. Jeśli w jednym rekordzie klucz ma literówkę (`"cenna"`), dowiesz się o tym **dopiero w tej iteracji**, w której ten rekord się pojawi — czyli być może po wgraniu pliku od klienta w listopadzie. Zero typowania, zero podpowiedzi w edytorze.

`ws.cell(row=3, column=1, value=...)` tworzy obiekt komórki. Zwróć uwagę, że **nie przypisujesz tego obiektu do zmiennej** i to jest w porządku — openpyxl trzyma komórki w wewnętrznej strukturze arkusza, więc dostęp przez `ws.cell(...)` to odczyt z tej struktury. Ale to znaczy, że `ws.cell(row=wiersz, column=6).font = Font(...)` w `if`-ie **modyfikuje komórkę, którą utworzyłeś cztery linie wyżej** — i tylko dlatego, że zgadza się numer wiersza. Jeżeli ktoś przesunie jedną z tych linii, `if` zacznie kolorować **inną komórkę**. Nic tego nie złapie.

**Co trafi do pliku.** Do `xl/worksheets/sheet1.xml` trafią cztery wiersze danych z `s="3"` (indeks stylu) w kolumnie `F` — dla rekordów Beta i Alfa. Do `xl/styles.xml` trafi kilka `font`/`fill`/`cellXfs`, a openpyxl **zdeduplikuje** je: dziesięć identycznych `Font(color="FF0000")` w Pythonie da **jeden** wpis w pliku (bo `IndexedList` w `styles` sprawdza, czy taki styl już istnieje). To dobra wiadomość dla rozmiaru pliku — i zła dla Ciebie, bo **nie zobaczysz w pliku, że w kodzie masz duplikację**. Plik będzie wyglądał tak samo, jakby kod był czysty.

I jeszcze jedno, bardzo pouczające: **`suma` trafi do kolumny `G`**, czyli za tabelę. Arkusz ma filtry/nagłówki w `A`–`F`. W Excelu wygląda to tak, że sumka „wisi z boku", wyrównana do prawej, trochę na uboczu. To nie wygląda jak błąd. To wygląda jak „taka konwencja". I dlatego nikt tego nie zgłosi.

**Trzy rzeczy do przemyślenia:**

1. **Ten kod jest poprawny.** Wygeneruje raport, który wygląda dobrze. **Gdyby to był jedyny raport w historii firmy, nie należałoby go zmieniać.** To jest test YAGNI w praktyce: jeżeli wersja 0 jest ostatnią wersją, refaktoryzacja jest stratą czasu.
2. **Sygnały nie są rozproszone — są skupione.** Zwróć uwagę, że pięć z sześciu sygnałów siedzi w promieniu dwudziestu linii. To znaczy, że w prawdziwych plikach **sygnały tworzą skupiska** — i dlatego da się je wykrywać automatycznie (zobaczysz narzędzie w ściągawce).
3. **Kolumna `G` to najlepsza ilustracja sygnału 3.** Bo tu widać mechanizm: ktoś pomylił się o jeden (liczył kolumnę `F` jako `7`, bo zapomniał, że pierwsza to `1`, a nie `0`). To jest dokładnie to off-by-one z modułu 03: **Excel numeruje od 1, Python od 0.** Wartość „7" nie mówi nic o tym, że chodziło o „za tabelą". `get_column_letter(7)` dałby `"G"` — i nadal nie rozwiązuje problemu, bo problemem nie jest nazwa, tylko to, że **kolumna G nie należy do tabeli**.

### Przykład 2 — funkcje + konfiguracja (poziom 1)

Ten sam raport, ten sam wynik. Trzy światy rozdzielone, konfiguracja w jednym miejscu, zero magicznych liczb.

```python
"""Modul 22, przyklad 2: POZIOM 1 - funkcje + konfiguracja.

Ten sam raport sprzedazy co w przykladzie 1, ale podzielony na trzy swiaty.
Roznice mierzalne:
  - literaly "FFC000" / "FF0000": 4 -> 0 (sa w Motyw)
  - magiczne indeksy kolumn:      11 -> 0 (wynikaja z KOLUMNY)
  - reguly biznesowe w zapisie:   1 -> 0
  - funkcje mozliwe do testu:     0 -> 4

Uruchom: python examples/22_poziom1.py
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

if TYPE_CHECKING:                      # import TYLKO dla typow, nie w runtime
    from openpyxl.worksheet.worksheet import Worksheet


# =====================================================================
# KONFIGURACJA - jedno miejsce, w ktorym zyja kolory, progi i kolumny
# =====================================================================
@dataclass(frozen=True)
class Kolumna:
    """Opis JEDNEJ kolumny. Z tego opisu wynika jej indeks - nigdy recznie."""
    naglowek: str
    pole: str                          # nazwa pola w WierszSprzedazy
    szerokosc: float
    format_liczby: str | None = None


@dataclass(frozen=True)
class Motyw:
    """Motyw firmowy. `frozen=True`, wiec bezpiecznie dzielimy jeden obiekt."""
    tlo_naglowka: str = "FFC000"
    tekst_naglowka: str = "FFFFFF"
    tekst_alertu: str = "FF0000"
    rozmiar_tytulu: int = 16


@dataclass(frozen=True)
class UstawieniaRaportu:
    tytul: str
    kolumny: tuple[Kolumna, ...]
    motyw: Motyw = Motyw()             # OK: Motyw jest niemutowalny
    prog_alertu: float = 30.0
    wiersz_tytulu: int = 1
    wiersz_naglowkow: int = 3
    wiersz_danych: int = 4


KOLUMNY = (
    Kolumna("Data",    "data",    12, "dd.mm.yyyy"),
    Kolumna("Klient",  "klient",  24),
    Kolumna("Region",  "region",   8),
    Kolumna("Produkt", "produkt", 16),
    Kolumna("Ilosc",   "ilosc",    8, "#,##0"),
    Kolumna("Cena",    "cena",    10, '#,##0.00 "zl"'),
)

USTAWIENIA = UstawieniaRaportu(
    tytul="Raport sprzedazy - styczen 2026",
    kolumny=KOLUMNY,
    prog_alertu=30.0,
)


# =====================================================================
# SWIAT DANYCH - nie importuje openpyxl i nie zna kolorow
# =====================================================================
@dataclass(frozen=True)
class WierszSprzedazy:
    data: date
    klient: str
    region: str
    produkt: str
    ilosc: int
    cena: float


def wczytaj_wiersze() -> list[WierszSprzedazy]:
    """Zrodlo danych. W prawdziwym zyciu: CSV, baza, API."""
    return [
        WierszSprzedazy(date(2026, 1, 5),  "Alfa sp. z o.o.", "PL", "Widzet A", 10, 25.0),
        WierszSprzedazy(date(2026, 1, 9),  "Beta S.A.",       "DE", "Widzet B",  4, 41.5),
        WierszSprzedazy(date(2026, 1, 17), "Gamma sp. k.",     "PL", "Widzet A", 22, 25.0),
        WierszSprzedazy(date(2026, 1, 23), "Alfa sp. z o.o.", "PL", "Gadzet C",  7, 63.0),
    ]


def wymaga_alertu(wiersz: WierszSprzedazy, prog: float) -> bool:
    """REGUŁA BIZNESOWA jako funkcja CZYSTA.

    Czysta = bez I/O, bez openpyxl, bez efektow ubocznych.
    Taka funkcje testujesz w mikrosekundach (modul 20, sekcja 2.10).
    """
    return wiersz.cena > prog


@dataclass(frozen=True)
class WierszWidoku:
    """Projekcja: gotowe WARTOSCI w kolejnosci kolumn + gotowa DECYZJA.

    Kluczowe: to nie jest wiersz danych. To wiersz PRZYGOTOWANY DO POKAZANIA.
    Prezentacja nie bedzie juz niczego liczyla - dostanie wynik.
    """
    wartosci: tuple[object, ...]
    alert: bool


def zrob_projekcje(
    wiersze: list[WierszSprzedazy],
    kolumny: tuple[Kolumna, ...],
    prog_alertu: float,
) -> list[WierszWidoku]:
    """Dane -> widok. Decyzja o alercie zapada TU, w swiecie danych."""
    return [
        WierszWidoku(
            wartosci=tuple(getattr(w, k.pole) for k in kolumny),
            alert=wymaga_alertu(w, prog_alertu),
        )
        for w in wiersze
    ]


# =====================================================================
# SWIAT ZAPISU - tylko openpyxl, ZADNYCH decyzji biznesowych
# =====================================================================
def zbuduj_arkusz(
    wb: Workbook,
    tytul_arkusza: str,
    ustawienia: UstawieniaRaportu,
    widok: list[WierszWidoku],
) -> Worksheet:
    ws = wb.create_sheet(title=tytul_arkusza)

    ws.cell(row=ustawienia.wiersz_tytulu, column=1, value=ustawienia.tytul)

    for i, kolumna in enumerate(ustawienia.kolumny, start=1):
        ws.cell(row=ustawienia.wiersz_naglowkow, column=i, value=kolumna.naglowek)

    wiersz = ustawienia.wiersz_danych
    for wp in widok:
        for i, wartosc in enumerate(wp.wartosci, start=1):
            c = ws.cell(row=wiersz, column=i, value=wartosc)
            fmt = ustawienia.kolumny[i - 1].format_liczby
            if fmt:
                c.number_format = fmt
        wiersz += 1

    return ws


# =====================================================================
# SWIAT PREZENTACJI - wykonuje decyzje, nie podejmuje ich
# =====================================================================
def sformatuj_arkusz(
    ws: Worksheet,
    ustawienia: UstawieniaRaportu,
    widok: list[WierszWidoku],
) -> None:
    m = ustawienia.motyw
    n_kolumn = len(ustawienia.kolumny)

    # tytul: piszemy go raz, potem scalanie (wartosc w lewej-gornej przezyje)
    ws.cell(row=ustawienia.wiersz_tytulu, column=1).font = Font(
        bold=True, size=m.rozmiar_tytulu
    )
    ws.merge_cells(
        start_row=ustawienia.wiersz_tytulu, start_column=1,
        end_row=ustawienia.wiersz_tytulu, end_column=n_kolumn,
    )

    # naglowki
    for i in range(1, n_kolumn + 1):
        c = ws.cell(row=ustawienia.wiersz_naglowkow, column=i)
        c.font = Font(bold=True, color=m.tekst_naglowka)
        c.fill = PatternFill(fill_type="solid", fgColor=m.tlo_naglowka)
        c.alignment = Alignment(horizontal="center", vertical="center")

    # alerty: WYKONUJEMY gotowa decyzje z WierszWidoku.alert
    wiersz = ustawienia.wiersz_danych
    for wp in widok:
        if wp.alert:                                  # to nie reguła - to wynik reguly
            for i in range(1, n_kolumn + 1):
                ws.cell(row=wiersz, column=i).font = Font(color=m.tekst_alertu)
        wiersz += 1

    # szerokosci i widok wynikaja z konfiguracji
    for i, kolumna in enumerate(ustawienia.kolumny, start=1):
        ws.column_dimensions[get_column_letter(i)].width = kolumna.szerokosc
    ws.freeze_panes = f"A{ustawienia.wiersz_danych}"


# =====================================================================
# I/O + zlozenie
# =====================================================================
def zapisz(wb: Workbook, sciezka: Path) -> Path:
    sciezka.parent.mkdir(parents=True, exist_ok=True)
    wb.save(sciezka)
    return sciezka


def main() -> None:
    wiersze = wczytaj_wiersze()
    widok = zrob_projekcje(wiersze, USTAWIENIA.kolumny, USTAWIENIA.prog_alertu)

    wb = Workbook()
    del wb["Sheet"]                       # domyslny arkusz nie jest nam potrzebny
    ws = zbuduj_arkusz(wb, "Sprzedaz", USTAWIENIA, widok)
    sformatuj_arkusz(ws, USTAWIENIA, widok)
    wb.properties.creator = "Automat raportowy"

    zapisz(wb, Path("output/22_poziom1.xlsx"))


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Zauważ dwie rzeczy, bo obie są istotne.

Po pierwsze: `USTAWIENIA` i `KOLUMNY` to **niemutowalne** obiekty (`frozen=True`) tworzone raz, przy imporcie modułu. `motyw=Motyw()` jako wartość domyślna oznacza, że **wszystkie** instancje `UstawieniaRaportu` dzielą **ten sam** obiekt `Motyw` — i jest to bezpieczne **właśnie dlatego**, że jest `frozen`. Gdyby `Motyw` był zwykłym `dataclass` (mutowalnym), ta sama konstrukcja byłaby pułapką: zmiana `motyw.tlo_naglowka` w jednym miejscu zmieniłaby ją **wszędzie**, we wszystkich raportach w procesie. Zapamiętaj to: **`frozen=True` nie jest ozdobą — jest warunkiem poprawności wspólnego obiektu.**

Po drugie: `zrob_projekcje` tworzy listę `WierszWidoku`, w której **każdy wiersz ma już rozstrzygnięty alert**. To znaczy, że `wymaga_alertu` jest wywołane **raz na wiersz**, a nie raz na komórkę. W oryginalnym skrypcie warunek też był raz na wiersz — ale tu jest różnica: wynik jest **danymi**, więc możesz go **wypisać, policzyć i przetestować** bez pliku:

```python
alberty = [wp.wartosci[1] for wp in widok if wp.alert]
print(alberty)      # ['Beta S.A.', 'Alfa sp. z o.o.']
```

**To jest cała różnica między „kodem, który działa" a „kodem, który da się sprawdzić".**

**Co trafi do pliku.** Wynikowy `.xlsx` jest **identyczny z przykładem 1** w warstwie widocznej — te same kolory, te same szerokości, te same alerty, ten sam zamrożony widok. W warstwie technicznej jest **jedna różnica**: suma nie wisi już w kolumnie `G` (bo jej nie ma — zostanie dodana dopiero, gdy specyfikacja będzie wiedzieć, które kolumny sumować). To celowe.

Zwróć jeszcze uwagę na rzecz, której w pliku **nie widać**, a która jest największą korzyścią: **`"FFC000"` występuje w całym projekcie dokładnie raz.** Zmiana firmowego żółtego to zmiana jednej linii. Zmiana progu alertu to zmiana jednej linii i **jednego testu**.

**Trzy rzeczy do przemyślenia:**

1. **`getattr(w, k.pole)` jest sprytne, ale ma termin ważności.** Działa, dopóki każda kolumna to **pole** obiektu. Przestaje działać, gdy kolumna jest **obliczona** — na przykład „Wartość = ilość × cena" albo „Udział %". **To jest dokładnie ten moment, w którym poziom 1 przestaje wystarczać.** Zapamiętaj to, bo to jest trigger przejścia do przykładu 3 — i dobry przykład reguły „przejście jest wywołane zdarzeniem, nie ambicją" z sekcji 2.2.
2. **`TYPE_CHECKING` to nie ozdoba.** Import `Worksheet` z `openpyxl.worksheet.worksheet` **wykonuje się** (openpyxl nie jest wolny), a jest potrzebny **wyłącznie dla adnotacji typów**. Dlatego jest schowany pod `TYPE_CHECKING` — w runtime nie istnieje, w edytorze i w `mypy` tak. To jest jednocześnie odpowiedź na pytanie „gdzie postawić granicę, żeby typy nie przeciekały": **adnotujesz, ale nie importujesz**. Uwaga: typy w openpyxl są **częściowe**, więc `mypy` może wymagać `# type: ignore` (moduł 20) — sprawdź w swojej konfiguracji.
3. **Nazwy pól w `Kolumna.pole` tworzą ciche sprzężenie.** `Kolumna("Klient", "klient", 24)` — jeżeli zmienisz nazwę pola w `WierszSprzedazy` na `nabywca`, a zapomnisz zmienić string `"klient"`, dostaniesz `AttributeError` w połowie pętli. **Widoczne, ale późno.** To jest świadomy koszt tego rozwiązania, który zniknie w przykładzie 3 (gdzie zamiast stringa jest funkcja) i zupełnie w module 24 (gdzie jest adapter).

### Przykład 3 — specyfikacja + klasa raportu (poziom 2)

Trigger, który uruchomia poziom 2, jest **konkretny** i wynika z przykładu 2: firma potrzebuje **drugiego raportu** — regionalnego. Ten raport:
- ma **inne kolumny** (Region, Liczba transakcji, Ilość, Wartość),
- **agreguje** wiersze (jedna linia na region),
- ma **inną kolumnę obliczaną** (Wartość = suma `ilosc * cena`),
- ma **inną regułę alertu** (alert, gdy wartość regionu **spadła poniżej** progu — choć w tym uproszczonym przykładzie: gdy wartość jest **niewielka**).

I teraz jest pytanie, na które odpowiada OCP: **ile linii kodu dopiszesz do istniejącego raportu sprzedaży?**

Odpowiedź, do której zmierzamy: **zero**.

```python
"""Modul 22, przyklad 3: POZIOM 2 - specyfikacja + klasa raportu.

Cel: DRUGI raport (regionalny) BEZ ZMIANY kodu pierwszego (OCP).

Uruchom: python examples/22_poziom2.py
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING, Callable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet


# =====================================================================
# SWIAT DANYCH
# =====================================================================
@dataclass(frozen=True)
class WierszSprzedazy:
    data: date
    klient: str
    region: str
    produkt: str
    ilosc: int
    cena: float

    @property
    def wartosc(self) -> float:
        return self.ilosc * self.cena


@dataclass(frozen=True)
class WierszRegionu:
    """INNY ksztalt danych - wynik agregacji. Ten sam swiat."""
    region: str
    transakcje: int
    ilosc: int
    wartosc: float


def wczytaj_wiersze() -> list[WierszSprzedazy]:
    return [
        WierszSprzedazy(date(2026, 1, 5),  "Alfa sp. z o.o.", "PL", "Widzet A", 10, 25.0),
        WierszSprzedazy(date(2026, 1, 9),  "Beta S.A.",       "DE", "Widzet B",  4, 41.5),
        WierszSprzedazy(date(2026, 1, 17), "Gamma sp. k.",     "PL", "Widzet A", 22, 25.0),
        WierszSprzedazy(date(2026, 1, 23), "Alfa sp. z o.o.", "PL", "Gadzet C",  7, 63.0),
    ]


def agreguj_regiony(wiersze: Sequence[WierszSprzedazy]) -> list[WierszRegionu]:
    """Dane -> inne dane. Czysta logika: zero openpyxl, zero I/O."""
    grupy: dict[str, list[WierszSprzedazy]] = {}
    for w in wiersze:
        grupy.setdefault(w.region, []).append(w)

    return [
        WierszRegionu(
            region=region,
            transakcje=len(pozycje),
            ilosc=sum(p.ilosc for p in pozycje),
            wartosc=sum(p.wartosc for p in pozycje),
        )
        for region, pozycje in sorted(grupy.items())
    ]


# =====================================================================
# SPECYFIKACJA - opis raportu jako DANE, nie jako kod
# =====================================================================
@dataclass(frozen=True)
class Kolumna:
    """Kolumna opisana FUNKCJA, nie nazwa pola.

    Dlaczego nie `pole: str` jak w przykladzie 2: bo kolumna moze byc
    OBLICZONA (np. "Wartosc" = ilosc * cena, "% udzialu"). `getattr`
    tego nie potrafi. Cena: kolumny biora `object` (patrz uWAGA nizej).
    """
    naglowek: str
    szerokosc: float
    wartosc: Callable[[object], object]
    format_liczby: str | None = None


@dataclass(frozen=True)
class Motyw:
    tlo_naglowka: str = "FFC000"
    tekst_naglowka: str = "FFFFFF"
    tekst_alertu: str = "FF0000"
    rozmiar_tytulu: int = 16


@dataclass(frozen=True)
class WierszWidoku:
    wartosci: tuple[object, ...]
    alert: bool


@dataclass(frozen=True)
class Specyfikacja:
    """CALY raport opisany danymi.

    UWAGA (swiadomy kompromis poziomu 2): `Kolumna.wartosc` i
    `regula_alertu` przyjmuja `object`, bo rozne raporty maja rozne
    wiersze (WierszSprzedazy, WierszRegionu, ...). To rozwiazanie
    TRACI ostrosc typow. Generyk / protokol naprawi to na poziomie 3
    (modul 24). Na poziomie 2 nie robimy tego swiadomie - patrz YAGNI.
    """
    tytul: str
    kolumny: tuple[Kolumna, ...]
    regula_alertu: Callable[[object], bool]
    motyw: Motyw = Motyw()
    wiersz_tytulu: int = 1
    wiersz_naglowkow: int = 3
    wiersz_danych: int = 4
    tytul_arkusza: str = "Raport"


# --- nazwane progi (koniec z magia 30/400 w kodzie) -------------------
PROG_CENY = 30.0
PROG_WARTOSCI_REGIONU = 400.0

SPEC_SPRZEDAZ = Specyfikacja(
    tytul="Raport sprzedazy - styczen 2026",
    kolumny=(
        Kolumna("Data",    12, lambda w: w.data,   "dd.mm.yyyy"),
        Kolumna("Klient",  24, lambda w: w.klient),
        Kolumna("Region",   8, lambda w: w.region),
        Kolumna("Produkt", 16, lambda w: w.produkt),
        Kolumna("Ilosc",    8, lambda w: w.ilosc,  "#,##0"),
        Kolumna("Cena",    10, lambda w: w.cena,   '#,##0.00 "zl"'),
    ),
    regula_alertu=lambda w: w.cena > PROG_CENY,
)

SPEC_REGIONALNY = Specyfikacja(
    tytul="Sprzedaz wg regionu - styczen 2026",
    kolumny=(
        Kolumna("Region",             10, lambda w: w.region),
        Kolumna("Liczba transakcji",  18, lambda w: w.transakcje, "#,##0"),
        Kolumna("Ilosc",              10, lambda w: w.ilosc,      "#,##0"),
        Kolumna("Wartosc",            14, lambda w: w.wartosc,    '#,##0.00 "zl"'),
    ),
    # INNA regula, ten sam mechanizm:
    regula_alertu=lambda w: w.wartosc < PROG_WARTOSCI_REGIONU,
    tytul_arkusza="Regiony",
)


# =====================================================================
# KLASA RAPORTU - jedyne miejsce, w ktorym spotykaja sie trzy swiaty
# =====================================================================
class RaportXlsx:
    """Uruchomienie: RaportXlsx(SPEC).generuj(wiersze) -> Workbook.

    Trzy swiaty sa tu tylko POSKLADANE, nie pomieszane:
      _projekcja  -> swiat danych (uzywa reguly ze specyfikacji)
      _zapisz_*   -> swiat zapisu (tylko openpyxl)
      _sformatuj  -> swiat prezentacji (wykonuje decyzje)
    """

    def __init__(self, specyfikacja: Specyfikacja) -> None:
        self.spec = specyfikacja

    # --- publiczne API -------------------------------------------------
    def generuj(self, wiersze: Sequence[object]) -> Workbook:
        widok = self._projekcja(wiersze)

        wb = Workbook()
        del wb["Sheet"]
        ws = wb.create_sheet(title=self.spec.tytul_arkusza)

        self._zapisz_tytul(ws)
        self._zapisz_naglowki(ws)
        self._zapisz_dane(ws, widok)
        self._sformatuj(ws, widok)

        wb.properties.title = self.spec.tytul
        wb.properties.creator = "Automat raportowy"
        return wb

    # --- swiat danych: projekcja --------------------------------------
    def _projekcja(self, wiersze: Sequence[object]) -> list[WierszWidoku]:
        return [
            WierszWidoku(
                wartosci=tuple(k.wartosc(w) for k in self.spec.kolumny),
                alert=self.spec.regula_alertu(w),
            )
            for w in wiersze
        ]

    # --- swiat zapisu --------------------------------------------------
    def _zapisz_tytul(self, ws: Worksheet) -> None:
        ws.cell(row=self.spec.wiersz_tytulu, column=1, value=self.spec.tytul)

    def _zapisz_naglowki(self, ws: Worksheet) -> None:
        for i, k in enumerate(self.spec.kolumny, start=1):
            ws.cell(row=self.spec.wiersz_naglowkow, column=i, value=k.naglowek)

    def _zapisz_dane(self, ws: Worksheet, widok: list[WierszWidoku]) -> None:
        wiersz = self.spec.wiersz_danych
        for wp in widok:
            for i, wartosc in enumerate(wp.wartosci, start=1):
                c = ws.cell(row=wiersz, column=i, value=wartosc)
                fmt = self.spec.kolumny[i - 1].format_liczby
                if fmt:
                    c.number_format = fmt
            wiersz += 1

    # --- swiat prezentacji ---------------------------------------------
    def _sformatuj(self, ws: Worksheet, widok: list[WierszWidoku]) -> None:
        m = self.spec.motyw
        n = len(self.spec.kolumny)

        ws.cell(row=self.spec.wiersz_tytulu, column=1).font = Font(
            bold=True, size=m.rozmiar_tytulu
        )
        ws.merge_cells(
            start_row=self.spec.wiersz_tytulu, start_column=1,
            end_row=self.spec.wiersz_tytulu, end_column=n,
        )

        for i in range(1, n + 1):
            c = ws.cell(row=self.spec.wiersz_naglowkow, column=i)
            c.font = Font(bold=True, color=m.tekst_naglowka)
            c.fill = PatternFill(fill_type="solid", fgColor=m.tlo_naglowka)
            c.alignment = Alignment(horizontal="center", vertical="center")

        wiersz = self.spec.wiersz_danych
        for wp in widok:
            if wp.alert:
                for i in range(1, n + 1):
                    ws.cell(row=wiersz, column=i).font = Font(color=m.tekst_alertu)
            wiersz += 1

        for i, k in enumerate(self.spec.kolumny, start=1):
            ws.column_dimensions[get_column_letter(i)].width = k.szerokosc
        ws.freeze_panes = f"A{self.spec.wiersz_danych}"


# =====================================================================
# UZYCIE - dwa raporty, jedna klasa, ZERO zmian w RaportXlsx
# =====================================================================
def zapisz(wb: Workbook, sciezka: Path) -> Path:
    sciezka.parent.mkdir(parents=True, exist_ok=True)
    wb.save(sciezka)
    return sciezka


def main() -> None:
    wiersze = wczytaj_wiersze()

    # Raport 1 - jak w przykladzie 2
    zapisz(
        RaportXlsx(SPEC_SPRZEDAZ).generuj(wiersze),
        Path("output/22_poziom2_sprzedaz.xlsx"),
    )

    # Raport 2 - NOWY, bez dotykania klasy RaportXlsx
    regiony = agreguj_regiony(wiersze)          # agregacja nalezy do swiata danych
    zapisz(
        RaportXlsx(SPEC_REGIONALNY).generuj(regiony),
        Path("output/22_poziom2_regiony.xlsx"),
    )


if __name__ == "__main__":
    main()
```

**Co się dzieje w pamięci.** Kluczowa zmiana w stosunku do przykładu 2: **kolumna to funkcja, nie nazwa pola**. `Kolumna("Wartosc", 14, lambda w: w.wartosc, ...)` — funkcja dostaje wiersz i zwraca wartość. Dzięki temu kolumna może być **obliczona** (a nie tylko odczytana), co było niemożliwe w `getattr(w, k.pole)`.

I teraz bardzo ważne, bo tu jest **cena** tego rozwiązania. `Wartosc` przyjmuje `WierszRegionu`, a `Klient` przyjmuje `WierszSprzedazy` — a typy w `Specyfikacja` mówią `Callable[[object], object]`. **Typy przestały być ostre.** Jeżeli napiszesz `Kolumna("Klient", 24, lambda w: w.klient)` i przypadkiem podasz `WierszRegionu`, dostaniesz `AttributeError` **w trakcie działania**, a nie od `mypy`. To jest prawdziwy koszt i zapisałem go w docstringu wprost, bo **ukrywanie kosztu abstrakcji jest gorsze niż sam koszt**.

To jest też doskonała ilustracja reguły z sekcji 2.3: **widzisz koszt od razu** („typy się rozjechały"), więc możesz go zaadresować, kiedy będzie na to dowód. Naprawa (generyk albo protokół) to poziom 3, czyli moduł 24. **Nie robimy tego teraz, bo nie mamy dowodu, że potrzebujemy więcej niż dwa raporty.**

Zauważ jeszcze jedno: **dwa raporty używają dwóch różnych typów wiersza** (`WierszSprzedazy` i `WierszRegionu`), ale **jednej klasy generującej**. Agregacja (`agreguj_regiony`) **nie jest** w klasie — jest poza nią, w świecie danych. To znaczy, że możesz jej użyć gdzie indziej (np. w mailu, w podsumowaniu na stronie) i możesz ją **przetestować bez openpyxl**.

**Co trafi do pliku.** Dwa pliki. W `22_poziom2_sprzedaz.xlsx` — to samo, co w przykładzie 2. W `22_poziom2_regiony.xlsx` — nowy układ: cztery kolumny, dwa wiersze (PL, DE), inne szerokości, inne formaty, inny tytuł, inny arkusz (`Regiony`), inne alerty (PL ma wartość 1490, czyli powyżej 400 → bez alertu; DE ma 166 → **alert**). Zwróć uwagę: **ta sama funkcja `_sformatuj`** obsłużyła dwa różne raporty, bo **czyta wszystko ze specyfikacji**.

**Trzy rzeczy do przemyślenia:**

1. **Ile linii dopisałeś do `RaportXlsx`, żeby powstał raport regionalny?** Zero. **To jest OCP zmierzone, a nie zadeklarowane.** I to jest dokładnie to kryterium, które wykorzystasz w ćwiczeniu 🔴: nie „czy czuję, że to OCP", a „ile linii dopisałem do istniejącego pliku?".
2. **Specyfikacja jest teraz `frozen=True`, ale zawiera `Callable`.** Czy to problem? Nie — funkcje są niemutowalne w praktyce (nie „modyfikujesz" lambdy). Ale gdyby w specyfikacji była **lista** kolumn (`list` zamiast `tuple`), `frozen=True` **nie ochroni** jej zawartości: `spec.kolumny.append(...)` by się udało. Dlatego `kolumny` to `tuple`, a nie `list` — i to nie jest detal stylistyczny, tylko **warunek** poprawności niemutowalnego obiektu.
3. **`PROG_CENY` i `PROG_WARTOSCI_REGIONU` są nazwane** i występują raz. To domyka sygnał 2 z sekcji 2.1. Ale zwróć uwagę na subtelność: **reguła jest w lambdzie, próg w stałej.** Dla dwóch raportów to jest właściwy poziom. Dla pięciu raportów i trzech rodzajów alertów stanie się to nieczytelne i wtedy wzorzec **Specification** (moduł 25) złoży regułę z klocków: `CenaWiekszaNiz(30) | WartoscMniejszaNiz(400)`. Znowu: **przejście wywołane zdarzeniem.**

### Przykład 4 — zajawka poziomu 3: port i adapter

Ostatni przykład jest **krótki i celowo niekompletny**. Pokaże Ci **jedną rzecz**: jak wygląda świat, w którym logika aplikacji **nie wie, że istnieje openpyxl**. Pełną wersję (transakcyjność, konfiguracja deklaratywna, `Repository`) zobaczysz w module 26.

```python
"""Modul 22, przyklad 4: ZAJAWKA POZIOMU 3 - port i adaptery.

To nie jest pelna architektura (ta jest w module 26). To pokazanie
JEDNEJ rzeczy: logika aplikacji moze nie wiedziec o openpyxl.

Uruchom: python examples/22_poziom3_zajawka.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Protocol, Sequence

from openpyxl import Workbook       # ⚠️ openpyxl wolno importowac TYLKO w adapterze


# =====================================================================
# PORT - obietnica: "umiem wyeksportowac raport". Nic wiecej.
# =====================================================================
class EksporterRaportu(Protocol):
    """Port (modul 26 rozwinie to w pelny ports-and-adapters).

    `Protocol` = sposob na "typ strukturalny": cokolwiek, co MA te metode,
    pasuje. Nie trzeba dziedziczyc. To jak okreslenie "cokolwiek, co ma
    wtyczke i gwint" - nie musisz znac marki zarowki.
    """
    def eksportuj(self, wiersze: Sequence[object], cel: Path) -> None: ...


# =====================================================================
# ADAPTER 1: realna realizacja przez openpyxl
# =====================================================================
class EksporterXlsx:
    """ADAPTER wie o domenie (zna pola .klient, .cena) i o openpyxl."""

    def eksportuj(self, wiersze: Sequence[object], cel: Path) -> None:
        wb = Workbook()
        del wb["Sheet"]
        ws = wb.create_sheet("Dane")
        ws.append(["Klient", "Ilosc", "Cena", "Wartosc"])

        wiersz = 2
        for i, w in enumerate(wiersze, start=2):
            ws.cell(row=i, column=1, value=w.klient)
            ws.cell(row=i, column=2, value=w.ilosc)
            ws.cell(row=i, column=3, value=w.cena)
            ws.cell(row=i, column=4, value=w.wartosc)

        ws.cell(row=wiersz + len(wiersze) + 1, column=4,
                value=sum(w.wartosc for w in wiersze)).number_format = '#,##0.00 "zl"'

        cel.parent.mkdir(parents=True, exist_ok=True)
        wb.save(cel)


# =====================================================================
# ADAPTER 2: realizacja dla testow - ZERO I/O, ZERO plikow
# =====================================================================
class EksporterDoTestow:
    """Nazywany tez 'fake'. Zapisuje WYWOLANIA, nie dane."""

    def __init__(self) -> None:
        self.wywolania: list[tuple[int, Path]] = []

    def eksportuj(self, wiersze: Sequence[object], cel: Path) -> None:
        self.wywolania.append((len(wiersze), cel))


# =====================================================================
# LOGIKA APLIKACJI - nie wie, ze istnieje openpyxl
# =====================================================================
def wygeneruj_raport(
    eksporter: EksporterRaportu,
    wiersze: Sequence[object],
    cel: Path,
) -> None:
    """Ta funkcja dziala identycznie dla XLSX, CSV, PDF i dla fake'a."""
    eksporter.eksportuj(wiersze, cel)


if __name__ == "__main__":
    from examples._dane import wczytaj_wiersze      # te same dane co wyzej

    wiersze = wczytaj_wiersze()

    # W produkcyjnym kodzie:
    wygeneruj_raport(EksporterXlsx(), wiersze, Path("output/22_poziom3.xlsx"))

    # A w TESTACH (modul 20):
    fake = EksporterDoTestow()
    wygeneruj_raport(fake, wiersze, Path("cokolwiek.xlsx"))
    assert fake.wywolania == [(4, Path("cokolwiek.xlsx"))]
    print("Test przeszedl. Zadnego pliku nie zapisano.")
```

**Co się dzieje w pamięci.** `wygeneruj_raport` dostaje **obiekt**, nie ścieżkę ani flagę. Jeżeli przekażesz `EksporterXlsx()`, w pamięci powstanie `Workbook` i skoroszyt trafi na dysk. Jeżeli przekażesz `EksporterDoTestow()`, powstanie **lista wywołań** — a dysk nie zostanie dotknięty.

**To jest cała istota poziomu 3 i jednocześnie odpowiedź na pytanie z ćwiczenia 🔴 z modułu 20** („jak przetestować generator bez Excela?"): **nie testujesz generatora — testujesz funkcję aplikacyjną z innym adapterem.**

Zwróć uwagę, że `Protocol` nie wymaga dziedziczenia. `EksporterDoTestow` **nie jest** podklasą `EksporterRaportu` — i mimo to jest w porządku, bo ma metodę `eksportuj` o zgodnym kształcie. Analogia: kiedy mówisz „potrzebuję czegoś, co ma wtyczkę i gwint E27", nie wymagasz, żeby żarówka miała na sobie nazwę Twojej firmy. **Interfejs to kształt, nie pochodzenie.** (To są tzw. typy strukturalne; `mypy` to sprawdzi, a Python w runtime nie — dlatego testy z modułu 20 są tu nadal potrzebne.)

**Co trafi do pliku.** W trybie produkcyjnym: jeden plik `.xlsx`. W trybie testowym: **nic**. Zero plików, zero katalogów, zero `BytesIO`. I to jest różnica, którą poczujesz w CI: testy poziomu 3 są **równie szybkie jak testy zwykłej logiki**, bo nie dotykają ani openpyxl, ani dysku. Porównaj to z testem, który musi zbudować skoroszyt, zapisać go i odczytać — jest 10–50× wolniejszy (moduł 20, `--durations`).

**Trzy rzeczy do przemyślenia:**

1. **`EksporterXlsx` ma w środku `if`-y? Nie ma. Ma pętle i przypisania.** To jest kryterium z sekcji 2.4: adapter nie podejmuje decyzji, a **wykonuje** je. Gdyby tu pojawił się `if w.cena > 30`, byłby to sygnał, że reguła wróciła do świata zapisu.
2. **Port ma jedną metodę.** To jest zasada **I** z SOLID (segregacja interfejsów) w praktyce. Port z ośmioma metodami nie dałby się zaimplementować w 5 liniach — a `EksporterDoTestow` ma **5 linii**, i bardzo dobrze. **Port ma być tak mały, żeby fałszywa implementacja była trywialna.**
3. **Zwróć uwagę na ten jeden, brzydki `ws.append` z listą nagłówków.** W tej zajawce pominąłem specyfikację, żeby kod był krótki. **Na poziomie 3 specyfikacja nie znika — ona zostaje** (moduł 26). Zajawka pokazuje tylko *wymianę warstwy zapisu*, nie całą architekturę. To celowe: **nie chcę, żebyś wyszedł z tego modułu z wrażeniem, że poziom 3 to „wszystko od nowa".** To ten sam kod plus jedna granica.

## 4. Anatomia API

Ten moduł jest nietypowy, więc i tabela jest nietypowa. Jego tematem nie jest API `openpyxl`, a **struktura kodu** — więc pierwsza część tabeli opisuje konstrukcje Pythona, których użyjesz do zbudowania trzech światów, a druga część przypomina te wywołania `openpyxl`, które w tej strukturze mają znaczenie.

### Część A — konstrukcje Pythona (to są Twoje narzędzia w tym module)

| Konstrukcja | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `@dataclass` | generuje `__init__`, `__repr__`, `__eq__` z pól klasy | `frozen`, `eq`, `order`, `slots` | **Podstawa modelu domenowego.** Bez tego piszesz `__init__` ręcznie i ktoś zapomni jednego pola |
| `@dataclass(frozen=True)` | obiekt **niemutowalny** | — | Warunek bezpiecznego **dzielenia** obiektu (np. `Motyw`) między raportami; `TypeError` przy przypisaniu |
| `field(default_factory=...)` | domyślna wartość tworzona **per instancja** | `default_factory`, `repr`, `compare` | Musisz użyć dla `list`/`dict`; dla `frozen` dataclass można zwykle podać `default=ObiektNiemutowalny()` |
| `@property` | pole wyliczane (np. `wartosc`) | — | Nie jest polem! Nie pojawi się w `dataclasses.fields()` ani w `asdict()` |
| `dataclasses.asdict(obj)` | dataclass → `dict` | `dict_factory` | Wygodne dla eksportu do CSV/JSON; **rekurencyjne** — uważaj na głębokie struktury |
| `typing.Protocol` | „typ strukturalny": pasuje wszystko, co **ma** metodę | — | Nie wymaga dziedziczenia. Idealny na **port** (moduł 26) |
| `typing.TYPE_CHECKING` | `True` tylko dla narzędzi typów | — | W runtime `False` → import się **nie wykona**. Standard dla importów typów z openpyxl |
| `typing.Callable[[A], B]` | adnotacja funkcji przekazywanej jako wartość | — | Podstawa wzorca **Strategy** (moduł 25) i „kolumny jako funkcji" |
| `typing.Sequence[T]` | „coś, co można indeksować i iterować" | — | **Preferuj nad `list[T]`** w sygnaturach — nie zmuszasz wołającego do konkretnej kolekcji |
| `abc.ABC` + `@abstractmethod` | klasa bazowa, której nie da się utworzyć bez implementacji | — | Zapowiedź `BazowyRaport` (moduł 25). **`Protocol` jest dziś często lepszy** — mniej dziedziczenia |
| `enum.Enum` | nazwane stałe | — | Dla kategorii (np. typ alertu). Literówka → `AttributeError` zamiast cichej nowej wartości |
| `class Kategoria(str, Enum)` | enum, który jest jednocześnie stringiem | — | Działa w `f-string` i `json.dumps`. Uwaga: `StrEnum` to Python **3.11+** — dla 3.9 użyj `(str, Enum)` |
| `getattr(obj, "nazwa")` | pole po **nazwie** | `default` | Podstawa adaptera w przykładzie 2. **Cicha pułapka: brak literówek nie złapie `mypy`** |
| `functools.partial` | funkcja z „wstępnie ustawionymi" argumentami | — | Mini-Builder: `wypisz_naglowek = partial(naglowek, motyw=MOTYW)` |
| `pathlib.Path` | ścieżki jako obiekty | — | `Path(__file__).parent` zamiast „zależnie od CWD". `mkdir(parents=True, exist_ok=True)` (moduł 04) |
| `Path.parent.mkdir(...)` | utworzenie katalogu na plik | `parents`, `exist_ok` | Bez tego `save()` padnie przy pierwszym uruchomieniu na czystym środowisku |

### Część B — wywołania `openpyxl`, które mają znaczenie w tej strukturze

| Metoda / klasa | Co robi | Parametry | Uwagi |
|---|---|---|---|
| `Workbook()` | nowy skoroszyt z domyślnym arkuszem | — | `del wb["Sheet"]`, gdy tworzysz własne arkusze — inaczej zostanie pusty „Sheet" w raporcie |
| `Workbook(write_only=True)` | tryb strumieniowy | — | **Świat zapisu.** Nie da się łączyć z drugim przebiegiem kolorowania — formatuj w tym samym przebiegu (moduł 19) |
| `wb.create_sheet(title=..., index=...)` | nowy arkusz | `title`, `index` | Nazwa musi być unikalna i ≤31 znaków (moduł 08) |
| `load_workbook(..., data_only=, keep_links=)` | wczytanie | `read_only`, `data_only`, `keep_vba`, `keep_links`, `rich_text` | Adapter wejściowy. `data_only=True` to próg bezpieczeństwa (moduł 21) |
| `ws.cell(row=, column=, value=)` | komórka po współrzędnych | `row`, `column`, `value` | **Świat zapisu.** `value` przypisuje i **ustala `data_type`** — neutralizację robisz po (moduł 21) |
| `cell.number_format` | kod formatu liczby | — | **Prezentacja.** Ale ustawiany w świecie zapisu — to jedyne miejsce, gdzie te dwa światy się stykają (moduł 10) |
| `cell.data_type = "s"` | wymuszenie tekstu | — | **Zawsze po** przypisaniu wartości (moduł 21) |
| `Font` / `PatternFill` / `Alignment` / `Border` | obiekty stylu | `bold`, `color`, `fill_type`, `fgColor`, `horizontal`, … | **Niemutowalne** (moduł 09). Twórz raz, używaj wielokrotnie — openpyxl deduplikuje w `styles.xml` |
| `wb.add_named_style(NamedStyle(...))` | styl nazwany skoroszytu | `name`, `font`, `fill`, `border`, `alignment`, `number_format`, `protection` | Potem `cell.style = "Naglowek"`. **Krok do rejestru stylów** (moduł 23). Sprawdź w dokumentacji swojej wersji |
| `PatternFill(fill_type="solid", fgColor=...)` | wypełnienie | `fill_type`, `fgColor`, `bgColor` | Dla `solid` **kolor tła to `fgColor`** — odwrotnie niż podpowiada intuicja (moduł 09) |
| `ws.merge_cells(start_row=, start_column=, end_row=, end_column=)` | scalanie | cztery współrzędne | **Prezentacja.** Wartości poza lewą-górną są tracone (moduł 11) |
| `ws.column_dimensions[get_column_letter(i)].width` | szerokość kolumny | — | Wynikaj z konfiguracji, nie z ręcznie wpisanej litery |
| `ws.freeze_panes = f"A{wiersz}"` | zamrożenie okien | — | „Przyklejony nagłówek" (moduł 11) |
| `openpyxl.utils.get_column_letter(i)` | 1 → `"A"` | — | Nazwij współrzędną zamiast wpisywać literę |
| `ws.iter_rows(values_only=True)` | iteracja po wierszach | `min_row`, `max_row`, `min_col`, `max_col`, `values_only` | **`values_only=True` to naturalna granica między światem zapisu a danymi** — dostajesz wartości, nie komórki |
| `wb.properties.title` / `.creator` | metadane | — | Ślad audytowy (moduł 04). Ustawiaj **w jednym miejscu**, nie w pięciu |

## 5. Ćwiczenia

### 🟢 Rozgrzewka

**Zadanie 1 — znajdź w swoim skrypcie 5 „magic numbers".**

Weź **swój** skrypt (ten prawdziwy, nie z kursu — albo `examples/22_poziom0.py`, jeżeli nie masz własnego). Twoim zadaniem nie jest poprawić go, tylko go **zdiagnozować**.

1. Skopiuj plik do `output/22_cw1_skrypt.py` (**nie modyfikuj oryginału** — zasada z modułu 18).
2. Wypisz tabelę **minimum pięciu** magicznych wartości z tego pliku. Dla każdej podaj:

| Wartość | W którym miejscu (opis, nie linia) | Jaką nazwę byś jej nadał | Do którego świata należy |
|---|---|---|---|

Przykład wpisu: `30` | warunek kolorowania w pętli danych | `PROG_ALERTU` | **dane** | `"FFC000"` | wypełnienie nagłówka | `TLO_NAGLOWKA` (albo pole `Motyw`) | **prezentacja** | `column=7` | zapis sumy | — (liczba nie powinna istnieć) | **zapis**.

3. Odpowiedz na trzy pytania:

- **Która z tych wartości występuje więcej niż raz?** (Użyj `grep` albo wyszukiwania w edytorze — `grep -n "30" plik.py`.)
- **Czy któraś z nich występuje dwa razy w dwóch różnych znaczeniach?** To jest najgroźniejszy przypadek z sekcji 2.1 (sygnał 2 — rozjazd wartości).
- **Czy któraś z nich to ta sama liczba, która przypadkiem „się zgadza"?** Na przykład: próg alertu `30` i rozmiar czcionki `30`? Albo indeks kolumny `7` i tydzień `7`?

W pliku `output/22_cw1_wnioski.md` odpowiedz jeszcze na jedno pytanie, najważniejsze:

> **Czy którykolwiek z tych magic numbers dałoby się wykryć testem?** Jeżeli nie — napisz, dlaczego nie (i przeczytaj jeszcze raz sekcję 2.3 o regule B: *abstrakcja musi być opłacona testem*).

**Wskazówka:** użyj narzędzia ze ściągawki (sekcja 8) — `wykryj_sygnaly()` policzy Ci heurystycznie: literały kolorów, duplikaty importów, magiczne indeksy, progi w warunkach, funkcje na poziomie modułu. Porównaj liczby dla `22_poziom0.py`, `22_poziom1.py` i `22_poziom2.py` i **wklej wynik do wniosków** — to jest Twoja pierwsza metryka jakości kodu w tym kursie.

### 🟡 Warsztat

**Zadanie 2 — podziel skrypt na trzy funkcje (dane / prezentacja / zapis).**

To jest rozbudowany przykład 2, ale z **kryteriami odbioru**, które musisz spełnić. Pracuj na własnym pliku z `output/22_cw1_skrypt.py`.

**Wymagania (każde jest sprawdzalne):**

1. **Trzy funkcje, każda w innym świecie.** Nazwy: dowolne, ale sygnatury **muszą** wyglądać tak (albo bardzo podobnie — uzasadnij odstępstwo):

```python
def przygotuj_dane(sciezka: Path) -> list[WierszSprzedazy]:        # ŚWIAT DANYCH
def zbuduj_arkusz(wb, dane, ustawienia) -> Worksheet:              # ŚWIAT ZAPISU
def sformatuj_arkusz(ws, dane, ustawienia) -> None:                # ŚWIAT PREZENTACJI
```

2. **Zero literałów kolorów i progów poza konfiguracją.** Kryterium: `grep -n 'FFC000\|FF0000\|> 30' src/` (poza plikiem konfiguracji) **musi zwrócić zero wyników**.
3. **Zero magicznych indeksów kolumn.** Kryterium: `grep -n 'column=[0-9]' src/` — jedyne dopuszczalne trafienia to `column=1` w zapisie tytułu (i musisz to uzasadnić w komentarzu) albo wyliczenie indeksu z konfiguracji.
4. **Żadna funkcja nie dłuższa niż 25 linii.** Przekroczenie = zła dekompozycja.
5. **Zero `import openpyxl` w funkcji `przygotuj_dane`** i w niczym, co ona wywołuje. Kryterium: usuń tymczasowo wszystkie importy openpyxl z pliku — `przygotuj_dane` musi się zaimportować i wywołać bez błędu (test w `python -c`).
6. **Pięć testów** (moduł 20, `pytest`):
   - `test_przygotuj_dane_zwraca_wiersze` — sprawdza liczbę i typ,
   - `test_wymaga_alertu_powyzej_progu` i `test_wymaga_alertu_rowne_progowi` (**`>` czy `>=` — rozstrzygnij świadomie!**),
   - `test_arkusz_ma_naglowek` (odczyt z `BytesIO`),
   - `test_liczba_formul_w_skoroszycie` — użyj `policz_formuly()` z modułu 21 i **assertuj dokładną liczbę**,
   - `test_prog_zmieniony_zmienia_tylko_alerty` — zmień próg, wygeneruj, sprawdź, że **liczba komórek z kolorem** się zmieniła, a **wartości pozostały identyczne**.

Ten ostatni test jest najważniejszy, bo **udowadnia, że reguła jest oddzielona od danych**. Gdyby była wtopiona w zapis, tego testu nie dałoby się napisać.

W pliku `output/22_cw2_wnioski.md` odpowiedz:

1. **Ile miejsc zmieniłeś, żeby zmienić kolor nagłówka?** (Policz pliki i linie.) Ile byś zmienił w wersji 0?
2. **Które magic numbers dały się zastąpić konfiguracją, a których nie dało się?** Dlaczego? (Podpowiedź: niektóre liczby — jak `column=1` w tytule — są „współrzędną strukturalną", a nie konfiguracją. Nie każdy magic number jest zły.)
3. **Czy `>` czy `>=`?** Jak to rozstrzygnąłeś i **kto** powinien to rozstrzygnąć (Ty czy dział sprzedaży)? Co się dzieje, gdy transakcja ma dokładnie `30.00`?
4. **Co się stało z długością pliku?** Porównaj liczbę linii przed i po. Wyjaśnij, dlaczego podział na funkcje **może wydłużyć** plik — i dlaczego to nie znaczy, że jest gorzej.
5. **Napisz `wykryj_sygnaly()` na obu wersjach** (przed i po) i wklej wyniki. Które wskaźniki spadły do zera? Które **nie** (i dlaczego — podpowiedź: `number_format` nadal występuje, ale teraz pochodzi z konfiguracji)?

### 🔴 Wyzwanie

**Zadanie 3 — przepisz skrypt na `dataclass` + funkcję zapisu, dodaj drugi raport bez zmiany pierwszego.**

To jest zadanie, które sprawdza, czy zrozumiałeś **OCP** — i jest zaprojektowane tak, żebyś się na nim potknął, jeśli tylko „poczujesz", że rozumiesz.

**Krok 1 — model domenowy.**

Przepisz dane z `dict` na `dataclass(frozen=True)`. Wymagania:
- `WierszSprzedazy` z polami: `data`, `klient`, `region`, `produkt`, `ilosc`, `cena`,
- pole obliczane `wartosc` jako `@property`,
- **typy** na wszystkich polach (`date`, `str`, `int`, `float`),
- funkcja `wczytaj_wiersze()` (nie zmieniaj źródła danych w tym kroku; jeżeli masz plik CSV, wczytaj go i **nie zapisuj z powrotem**).

**Krok 2 — funkcja zapisu.**

Napisz `zapisz_raport(wiersze, specyfikacja, cel: Path) -> Path`, która:
- składa trzy światy w jednym miejscu,
- **nie zawiera ani jednego `if`** dotyczącego raportu („jeśli sprzedaż, to…") — kryterium jest twarde,
- używa `Path.parent.mkdir(parents=True, exist_ok=True)`,
- ustawia `wb.properties.title` na `specyfikacja.tytul`,
- zapisuje **atomowo** (moduł 18/21: `mkstemp` + `os.replace`, **bez `mktemp`**),
- ma `try/finally` i domyka skoroszyt (przy `read_only` — `wb.close()`).

**Krok 3 — drugi raport. To jest właściwy test.**

Twoja firma chce **„Raport alertów"**: tylko wiersze spełniające regułę, kolumny `Klient`, `Produkt`, `Cena`, tytuł „Transakcje wymagające uwagi", inny motyw (np. czerwony nagłówek), **inna reguła alertu** (`cena > PROG` zostaje, ale alert pokazujesz na **całym wierszu**).

I teraz kryterium:

> **Ile linii dopisałeś do pliku, w którym jest `zapisz_raport` i `Specyfikacja`? Wpisz tę liczbę do wniosków.**

Jeżeli dodałeś `if nazwa_raportu == "alerty"` albo `elif`, albo nowy parametr typu `rodzaj_raportu` — **złamałeś OCP** i Twój `zapisz_raport` właśnie zaczął się zamieniać w funkcję `generuj(raport, dane)` z sekcji 2.6. To jest **dobry** wynik tego ćwiczenia, jeżeli to zauważysz: **właśnie po to jest to zadanie.** Napisz wtedy w uwagach, gdzie dokładnie się potknąłeś.

**Krok 4 — testy (minimum 8).**

1. Test, że `wczytaj_wiersze()` zwraca `list[WierszSprzedazy]` i poprawne typy.
2. Test, że `wartosc` liczy się poprawnie (i że jest `float`, nie `Decimal`-em ani stringiem).
3. Test „golden": wygeneruj raport i sprawdź **odcisk struktury** — nazwy arkuszy, wymiary, liczbę komórek z danymi, **liczbę formuł**, liczbę wpisów stylów. Nie porównuj bajt w bajt (moduł 20!).
4. Test, że **oba raporty** powstają z tego samego `zapisz_raport`.
5. Test, że raport alertów ma **mniej wierszy** niż pełny, i że **wszystkie** jego wiersze spełniają regułę.
6. Test, że zmiana `PROG` w specyfikacji zmienia **tylko** alerty (nie wartości).
7. Test, że pliki tymczasowe nie zostają po zapisie (sprawdź zawartość katalogu — moduł 21).
8. Test, że funkcja `zapisz_raport` **nie zawiera w źródle** wystąpień `if` związanych z raportem. **Tak, taki test da się napisać** — i to jest komentarz-wykładnia OCP:

```python
import inspect

def test_zapisz_raport_nie_rozgalezia_sie_po_typie_raportu():
    """Testujemy STRUKTURE, nie zachowanie. To obrona przed regresja OCP."""
    zrodlo = inspect.getsource(zapisz_raport)
    zabronione = ("== \"sprzedaz\"", "== 'sprzedaz'", "elif", "rodzaj_raportu")
    for wzorzec in zabronione:
        assert wzorzec not in zrodlo, (
            f"W zapisz_raport pojawilo sie rozgalezienie: {wzorzec!r}. "
            "Prawdopodobnie zlamales OCP - nowy raport powinien byc DANYMI."
        )
```

W pliku `output/22_cw3_wnioski.md` odpowiedz:

1. **Ile linii dopisałeś do `zapisz_raport`, żeby powstał raport alertów?** To jest Twoja liczba OCP.
2. **Czym różni się `Specyfikacja` od klasy `RaportXlsx`?** Odpowiedz w trzech zdaniach: co jest **danymi** (i może być w YAML), a co jest **kodem** (i musi być w Pythonie). Gdzie postawiłbyś granicę?
3. **Co się stało z agregacją?** Jeżeli raport alertów wymaga filtra, a raport regionalny agregacji — gdzie te operacje mieszkają? Dlaczego **nie** w `zapisz_raport`?
4. **Ile masz testów i ile trwają?** Uruchom `pytest --durations=5`. Ile z nich nie dotyka `openpyxl`? Ile **nie dotyka dysku**? Wyjaśnij, jak ten drugi wskaźnik wpłynie na Twoje CI (moduł 20).
5. **Czy Twój model domenowy importuje `openpyxl`?** Sprawdź `grep -n "openpyxl" src/domena.py` (albo jakkolwiek nazwałeś plik z modelami). **Musi być zero.** Jeżeli nie jest — napisz, co importowało i dlaczego to był błąd.
6. **Nazwij to, co zbudowałeś**, używając słownictwa z sekcji 2.4 i 2.7: które elementy Twojego kodu to **świat danych**, które **prezentacji**, a które **zapisu**? Czy jest jakiś fragment, którego nie umiesz jednoznacznie zakwalifikować? (Jeżeli tak — to najciekawsza część wniosków. Napisz, które z trzech pytań z sekcji 2.5 było najtrudniejsze i **dlaczego**.)

<details>
<summary><strong>Szkic rozwiązania zadania 3 — kluczowe decyzje i fragmenty</strong></summary>

**Decyzja 1: `frozen=True` na modelu domenowym.**

```python
@dataclass(frozen=True)
class WierszSprzedazy:
    data: date
    klient: str
    region: str
    produkt: str
    ilosc: int
    cena: float

    @property
    def wartosc(self) -> float:
        return self.ilosc * self.cena
```

**Dlaczego `frozen`:** bo wiersz danych **nie zmienia się w trakcie budowania raportu**. Jeżeli w kodzie jest `wiersz.cena = 0` (żeby „ukryć" jakąś transakcję), to jest to **błąd logiczny**, a nie operacja. `frozen=True` zamienia go w `FrozenInstanceError` **w momencie wykonania** — a więc w testach, nie u klienta. To jest reguła B w praktyce: **niemutowalność jest tańsza niż testowanie mutacji.**

**Pułapka, na którą się natkniesz:** `frozen=True` nie chroni przed mutacją **pola mutowalnego**. Gdybyś miał `produkty: list[str]`, to `wiersz.produkty.append(...)` zadziała. Dlatego w modelu domenowym **preferuj `tuple` nad `list`** i `Mapping`/`dict` tylko tam, gdzie naprawdę trzeba.

**Pułapka druga:** `@property` **nie jest polem**. Nie pojawi się w `dataclasses.fields()`, w `asdict()` ani w `repr()`. Jeżeli używasz `asdict()` do eksportu, `wartosc` trzeba policzyć ręcznie. Zapamiętaj: **`@property` to wygoda, nie dane.**

**Decyzja 2: `Specyfikacja` jako `frozen=True` dataclass z `tuple`, a nie z `list`.**

```python
@dataclass(frozen=True)
class Specyfikacja:
    tytul: str
    kolumny: tuple[Kolumna, ...]        # ⬅ TUPLE, nie list!
    regula_alertu: Callable[[object], bool]
    motyw: Motyw = Motyw()
```

**Dlaczego `tuple`, a nie `list`:** `frozen=True` blokuje **przypisanie** (`spec.kolumny = (...)` → `FrozenInstanceError`), ale **nie blokuje mutacji zawartości**. `spec.kolumny.append(...)` zadziała, jeżeli `kolumny` to lista. Sam `frozen=True` nie daje więc prawdziwej niemutowalności — **musisz użyć typu niemutowalnego.** To jest klasyczna pułapka i dotyczy każdego „niemutowalnego" obiektu w Pythonie.

**Decyzja 3: drugi raport = nowa `Specyfikacja` + funkcja filtrująca, ZERO zmian w `zapisz_raport`.**

```python
# --- świat danych: selekcja nalezy do danych ---------------------------
def wybierz_alerty(
    wiersze: Sequence[WierszSprzedazy], prog: float
) -> list[WierszSprzedazy]:
    """Filtr jest CZYSTĄ funkcją - testowalna bez pliku."""
    return [w for w in wiersze if w.cena > prog]


SPEC_ALERTY = Specyfikacja(
    tytul="Transakcje wymagajace uwagi",
    kolumny=(
        Kolumna("Klient",  24, lambda w: w.klient),
        Kolumna("Produkt", 16, lambda w: w.produkt),
        Kolumna("Cena",    10, lambda w: w.cena, '#,##0.00 "zl"'),
    ),
    regula_alertu=lambda _w: True,     # w tym raporcie KAZDY wiersz jest alertem
    motyw=Motyw(tlo_naglowka="C00000", tekst_naglowka="FFFFFF"),
    tytul_arkusza="Alerty",
)

def main() -> None:
    wiersze = wczytaj_wiersze()

    zapisz_raport(wiersze, SPEC_SPRZEDAZ, Path("output/22_cw3_sprzedaz.xlsx"))

    # Nowy raport = nowe DANE (specyfikacja + selekcja). Zero nowego kodu!
    zapisz_raport(
        wybierz_alerty(wiersze, PROG_CENY),
        SPEC_ALERTY,
        Path("output/22_cw3_alerty.xlsx"),
    )
```

**Zwróć uwagę na `regula_alertu=lambda _w: True`.** To jest miejsce, w którym łatwo się potknąć i wpisać do `Specyfikacja` pole typu `pokazuj_alerty_dla_wszystkich: bool`. **Nie rób tego.** `lambda _w: True` znaczy „każdy wiersz jest alertem" i jest **regułą** — a bool byłby **flagą**, czyli nowym `if`-em czekającym na pojawienie się w `zapisz_raport`. **Flaga w konfiguracji to zakamuflowany `if` w kodzie.** Jeżeli kiedykolwiek zobaczysz `if spec.pokazuj_alarmy:`, to znaczy, że jedna flaga już przeciekła przez granicę świata.

**Decyzja 4: gdzie mieszka selekcja i agregacja (i dlaczego tam).**

| Operacja | Gdzie | Dlaczego tam |
|---|---|---|
| `wybierz_alerty()` | świat danych | To reguła biznesowa: **które transakcje wymagają uwagi**. Nie ma nic wspólnego z plikiem |
| `agreguj_regiony()` | świat danych | To przekształcenie danych w inne dane |
| `spec.kolumny` (kolejność i formaty) | granica danych/prezentacji | To decyzja **jak pokazać** — ale opisana jako **dane**, bo różni się per raport |
| `regula_alertu` | świat danych | Zwraca **bool opisujący dane**, nie kolor |
| `wp.alert` → `Font(color=...)` | granica prezentacji/zapisu | Wykonanie decyzji w komórce |

Zauważ, że nie ma tu żadnego punktu, w którym **liczba** jest decyzją zapisu. To jest test jakości podziału: **jeżeli nie umiesz znaleźć miejsca dla jakiejś operacji w tej tabeli, to znaczy, że w Twoim modelu brakuje pojęcia** — a nie że operacja „przecieka".

**Decyzja 5: `zapisz_raport` — pełna sygnatura i atomowy zapis.**

```python
def zapisz_raport(
    wiersze: Sequence[object],
    spec: Specyfikacja,
    cel: Path,
) -> Path:
    """Składa trzy światy i zapisuje. ZERO rozgalezien po typie raportu."""
    widok = [
        WierszWidoku(
            wartosci=tuple(k.wartosc(w) for k in spec.kolumny),
            alert=spec.regula_alertu(w),
        )
        for w in wiersze
    ]

    wb = Workbook()
    del wb["Sheet"]
    ws = wb.create_sheet(title=spec.tytul_arkusza)

    # --- zapis wartosci (swiat zapisu) ---
    ws.cell(row=spec.wiersz_tytulu, column=1, value=spec.tytul)
    for i, k in enumerate(spec.kolumny, start=1):
        ws.cell(row=spec.wiersz_naglowkow, column=i, value=k.naglowek)

    wiersz = spec.wiersz_danych
    for wp in widok:
        for i, wartosc in enumerate(wp.wartosci, start=1):
            c = ws.cell(row=wiersz, column=i, value=wartosc)
            fmt = spec.kolumny[i - 1].format_liczby
            if fmt:
                c.number_format = fmt
        wiersz += 1

    # --- prezentacja ---
    m = spec.motyw
    n = len(spec.kolumny)
    ws.cell(row=spec.wiersz_tytulu, column=1).font = Font(
        bold=True, size=m.rozmiar_tytulu
    )
    ws.merge_cells(
        start_row=spec.wiersz_tytulu, start_column=1,
        end_row=spec.wiersz_tytulu, end_column=n,
    )
    for i in range(1, n + 1):
        c = ws.cell(row=spec.wiersz_naglowkow, column=i)
        c.font = Font(bold=True, color=m.tekst_naglowka)
        c.fill = PatternFill(fill_type="solid", fgColor=m.tlo_naglowka)
        c.alignment = Alignment(horizontal="center", vertical="center")

    wiersz = spec.wiersz_danych
    for wp in widok:
        if wp.alert:
            for i in range(1, n + 1):
                ws.cell(row=wiersz, column=i).font = Font(color=m.tekst_alertu)
        wiersz += 1

    for i, k in enumerate(spec.kolumny, start=1):
        ws.column_dimensions[get_column_letter(i)].width = k.szerokosc
    ws.freeze_panes = f"A{spec.wiersz_danych}"

    wb.properties.title = spec.tytul
    wb.properties.creator = "Automat raportowy"

    # --- zapis ATOMOWY (moduly 18 i 21) ---
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
```

**Ile tu jest `if`-ów?** Cztery: `if fmt`, `if wp.alert` (dwa razy). **Ani jednego o typie raportu.** Wszystkie dotyczą **treści** obiektu, nie jego tożsamości — i to jest kryterium.

**Test `inspect.getsource` warto rozszerzyć** o jedno: sprawdź też, że w źródle nie ma `rodzaj`, `typ_raportu`, `nazwa_raportu`. Boje się, że pierwszy „poprawiający" ten kod doda `if spec.tytul == "Transakcje wymagajace uwagi"`. **Test na strukturę to test przeciwko przyszłej głupocie — i to jest w porządku, bo przyszłą głupotę popełnia zwykle ktoś zmęczony o 23:00.**

**Czego tu świadomie NIE ma:**

- **Nie ma podsumowania na dole.** Świadomie: wymagałoby wiedzy „które kolumny sumować", czyli kolejnego pola w `Specyfikacja` (`agregat: Callable | None`). **To jest właściwy kierunek, ale nie mam na niego jeszcze dowodu** (dwa raporty, oba bez sumy w specyfikacji). Zgodnie z sekcją 2.3, czekam na trzeci raport. Gdyby dziś ktoś dodał sumę „na sztywno" w `zapisz_raport`, złamałby OCP i wrócił do `if typ == ...`.
- **Nie ma generyka na typie wiersza** (`RaportXlsx[T]`). Dałoby się, ale: (a) `mypy` dla `Protocol` + `TypeVar` w `Callable[[T], object]` bywa kłopotliwy, (b) koszt zrozumienia rośnie, (c) mamy dwa raporty. **To jest poziom 3 z tabeli z sekcji 2.2** i pojawi się w module 24 — wtedy, gdy będzie na niego dowód.
- **Nie ma konfiguracji w YAML.** W module 26 zobaczysz, jak to zrobić i **jakie jest z tym realne ryzyko** (konfiguracja, która staje się językiem programowania). Teraz specyfikacja jest w Pythonie i to jest właściwe, bo w Pythonie masz **typy i podpowiedzi edytora**.

</details>

## 6. Typowe błędy i pułapki

**1. Objaw: „rozbudowałem strukturę, ale nadal nie umiem niczego zmienić — teraz tylko nie umiem w trzech plikach zamiast w jednym."**
→ *Przyczyna:* abstrakcja **bez testów**. Podzieliłeś kod na funkcje i klasy, ale nie napisałeś ani jednego testu. To znaczy, że przy każdej zmianie nadal nie masz informacji zwrotnej — tylko teraz błąd może być w trzech plikach. To jest reguła B z sekcji 2.3 złamana w najgorszym możliwym kierunku: **zainwestowałeś w strukturę, nie kupując tego, co ta struktura miała dać.**
→ *Naprawa:* **cofnij się i napisz testy.** Zacznij od **funkcji czystych** (reguła alertu, agregacja, filtry) — one nie wymagają nawet `BytesIO`, wystarczy `assert`. Dopiero potem testy na `BytesIO` (moduł 20). I dopiero **potem** wróć do podnoszenia poziomu.

**2. Objaw: „dziesięciowierszowy raport wymaga klasy, buildera, protokołu, trzech plików i pliku YAML."**
→ *Przyczyna:* **wzorce dla samych wzorców.** Nie jesteś na poziomie 2 — jesteś na poziomie 0, ale przeczytałeś o wzorcach i chcesz ich użyć. Poziom abstrakcji wybrałeś na podstawie „co wypada", a nie „co przetrwa następną zmianę" (sekcja 2.2). Wzorce są pudełkami na narzędzia — użycie trzech pudełek na jeden klucz to nie porządek, a bałagan.
→ *Naprawa:* **zadaj trzy pytania z sekcji 2.2.** Czy uruchomi się więcej niż dwa razy? Czy zmieniały się kolumny? Czy są trzy warianty? Jeżeli wszystkie odpowiedzi brzmią „nie" — **napisz skrypt liniowy i usuń resztę.** Usunięcie abstrakcji jest poprawną decyzją, nie porażką.

**3. Objaw: „usunąłem duplikację kodu, a teraz muszę dodać trzeci parametr do wspólnej funkcji, żeby obsłużyła drugi przypadek."**
→ *Przyczyna:* **abstrakcja wyciągnięta z dwóch przypadków.** Złamałeś regułę trzech (sekcja 2.3). Przy dwóch przypadkach widzisz **podobieństwo**, ale nie wiesz, co jest istotą, a co przypadkiem. Trzeci przypadek pokaże Ci prawdę — i wtedy wspólna funkcja pęka, a Ty zaczynasz dodawać parametry i `if`-y, które **odtwarzają** oryginalny problem.
→ *Naprawa:* **rozduplikuj z powrotem.** Tak, to boli, ale to jest tańsze niż „framework" z trzema flagami. Potem poczekaj na trzeci przypadek — tym razem wiesz już, co ma być wspólne. Sygnał ostrzegawczy: **jeśli funkcja ma więcej niż dwa parametry konfiguracyjne, prawdopodobnie obsługuje dwa różne przypadki.**

**4. Objaw: „przepisałem skrypt na `dataclass` i funkcje, i raport wygląda inaczej niż przedtem — nie wiem, czy to błąd, czy nie."
→ *Przyczyna:* **refaktoryzacja i zmiana zachowania wykonane w jednym kroku, bez testu „golden".** Refaktoryzacja z definicji **nie zmienia zachowania** — ale w tej definicji jest słowo „zachowanie", a Ty go nie zmierzyłeś, więc nie wiesz, czy refaktoryzacja się udała (moduł 20, sekcja 3).
→ *Naprawa:* **przed refaktoryzacją** wygeneruj plik referencyjny i zapisz jego **odcisk struktury** (nazwy arkuszy, wymiary, liczba komórek z danymi, liczba formuł, liczba reguł formatowania warunkowego, ustawienia wydruku). Po refaktoryzacji wygeneruj drugi i **porównaj odciski.** Jeżeli są identyczne — refaktoryzacja jest poprawna. Jeżeli nie — **wiesz, co się zmieniło**, zamiast zgadywać „chyba jest dobrze". Nigdy nie porównuj bajt w bajt: ZIP ma nieoznaczone metadane.

**5. Objaw: „wszystko jest w klasie `RaportManager`, ale ma 47 metod i nikt nie wie, które są publiczne."**
→ *Przyczyna:* **klasa jako worek na kod.** Zamiast trzech światów masz jeden obiekt, do którego wrzuciłeś wszystko: konfigurację, reguły, zapis, formatowanie, I/O, wysyłkę maila. Objaw: klasa ma w nazwie `Manager`, `Helper`, `Utils` albo `Processor`.
→ *Naprawa:* **klasa ma reprezentować POJĘCIE, nie paczkę funkcji.** `RaportXlsx` reprezentuje pojęcie „oszablonowany raport". Reguła alertu **nie jest** metodą tej klasy — jest **funkcją w świecie danych**. Rozbij worek na trzy obiekty i sprawdź liczby: jeśli klasa ma więcej niż ~8 metod publicznych, to prawdopodobnie ma więcej niż jedną odpowiedzialność (zasada **S** z sekcji 2.6).

**6. Objaw: „dodałem do konfiguracji pole `czy_wysylac_maila` i teraz mój `zapisz_raport` ma `if spec.czy_wysylac_maila:`."**
→ *Przyczyna:* **flaga w konfiguracji to zakamuflowany `if` w kodzie.** Uważałeś, że przeniesienie decyzji do konfiguracji to „czystość" — ale przeniosłeś tylko jej **warunek**, zostawiając **rozgałęzienie** w logice. Konfiguracja opisuje **co**, kod opisuje **jak**.
→ *Naprawa:* **flaga → zachowanie.** Zamiast `if spec.czy_wysylac_maila` zrób `EksporterEmail` jako **drugi adapter** (przykład 4) i wybieraj go **poza** funkcją zapisu. Wtedy `masz_wysylac = True` żyje w miejscu składania systemu, a nie w środku algorytmu. To jest dokładnie różnica między OCP a „konfiguracją, która udaje OCP".

**7. Objaw: „napisałem `Specyfikacja` w YAML, ale żeby dodać warunek `i co najwyżej 3 regiony`, musiałem dodać `python:` z kodem w YAML-u."
→ *Przyczyna:* **konfiguracja deklaratywna, która urosła do języka programowania.** To klasyczna pułapka poziomu 3 i **główny powód, dla którego w tym module użyłem Pythona, a nie YAML** — i dlaczego moduł 26 będzie tego **równie mocno ostrzegał**, jak zachwalał.
→ *Naprawa:* **granica jest jasna: konfiguracja opisuje CO, kod opisuje JAK.** Formaty liczb, nagłówki, kolory, progi — tak, w konfiguracji. Operatory (`>`, `and`, `or`), pętle i warunki — **nie**, w kodzie, jako **nazwane funkcje** w świecie danych. Jeżeli w konfiguracji musisz napisać `and`, to znaczy, że powinieneś napisać funkcję `wybierz_alerty()`. Test: **jeżeli konfiguracja wymaga debuggera, to nie jest konfiguracja.**

**8. Objaw: „mój adapter do CSV robi `w.get("klietn", "")` i raport wyszedł z pustą kolumną klientów, a nikt tego nie zauważył."
→ *Przyczyna:* **`.get()` z domyślną wartością cicho połyka literówkę.** Brak klucza w `dict` normalnie dałby `KeyError` — **widoczny i szybki**. Ale domyślna `""` zamienia błąd w **pustą kolumnę**, którą klient zobaczy po tygodniu. To jest ten sam mechanizm co `max_row` zawyżony przez formatowanie (moduł 07): **narzędzie, które „naprawia" problem, ukrywa go.**
→ *Naprawa:* **zamień `dict` na `dataclass`.** `WierszSprzedazy(..., klient=...)` — literówka w nazwie pola to `TypeError` **przy tworzeniu obiektu**, czyli natychmiast, w momencie wczytania, zanim powstanie raport. A jeżeli musisz zostać przy `dict` — używaj `w["klient"]` (bez domyślnej) i **wcześnie waliduj** braki jednym przebiegiem: `brakujace = {"klient", "cena"} - set(probka.keys())`.

**9. Objaw: „mój `@dataclass(frozen=True)` da się zmodyfikować — `spec.kolumny.append(...)` działa."**
→ *Przyczyna:* **`frozen=True` blokuje przypisanie, nie mutację zawartości.** `spec.kolumny = x` podniesie `FrozenInstanceError`, ale `spec.kolumny.append(x)` — nie. `frozen` chroni **referencję**, nie **obiekt, na który referencja wskazuje**.
→ *Naprawa:* **używaj typów niemutowalnych w polach.** `tuple[Kolumna, ...]` zamiast `list`, `Mapping` zamiast `dict`, `frozenset` zamiast `set`. To nie jest detal stylistyczny: **bez tego niemutowalność jest deklaracją bez pokrycia.** I pamiętaj, że `tuple` chroni **jeden poziom** — jeżeli `Kolumna` byłaby mutowalna, to i tak możesz ją zmienić przez referencję. Dlatego **`frozen=True` na wszystkich obiektach w łańcuchu** (tu: `Kolumna`, `Motyw`, `Specyfikacja`) jest warunkiem spójności.

**10. Objaw: „wyniosłem regułę alertu do specyfikacji, ale teraz to samo `cena > 30` występuje też w podsumowaniu, i po zmianie progu jedno mówi `35`, a drugie `30`."**
→ *Przyczyna:* **duplikacja reguły na dwóch poziomach abstrakcji.** Wyciągnąłeś stałą (`PROG_CENY`) i myślałeś, że to wystarczy — ale **reguła jako całość** (`cena > 30`) jest nadal w dwóch miejscach: raz w lambdzie, raz w warunku podsumowania. To jest dokładnie sygnał 2 z sekcji 2.1 w nowym przebraniu.
→ *Naprawa:* **wyciągnij CAŁĄ regułę, nie jej parametr.** Zrób funkcję `wymaga_uwagi(w, prog)` i **wywołuj ją z obu miejsc** — także z podsumowania. Docelowo (moduł 25) reguła stanie się **obiektem** (`Specification`), który potrafi i **filtrować**, i **wygenerować regułę formatowania warunkowego** — jedno źródło prawdy dla dwóch światów. Test kontrolny: `grep -n "cena >" src/` musi zwracać **jedno** trafienie.

**11. Objaw: „zrobiłem trzy pliki zamiast jednego i teraz nie wiem, gdzie jest błąd — cały projekt jest mniej czytelny."**
→ *Przyczyna:* **podział bez granic.** Rozbiłeś kod na trzy pliki według linijek („tu przepisałem pierwszą połowę"), a nie według **światów**. Pliki mają nazwy `utils.py`, `helpers.py`, `main.py` — a nie `domena.py`, `prezentacja.py`, `zapis.py`. Objaw: nie umiesz odpowiedzieć na pytanie „do którego pliku należy ta funkcja?" bez czytania kodu.
→ *Naprawa:* **granice wyznacza pytanie, nie linijka.** Dla każdej funkcji zadaj **trzy pytania z sekcji 2.5**. Jeżeli nie umiesz zakwalifikować — funkcja prawdopodobnie robi dwie rzeczy (złamanie SRP) i trzeba ją rozciąć. A jeśli nadal nie umiesz, **przenieś ją do świata danych** — domyślny kierunek jest zawsze „do danych", bo dane nie zależą od nikogo.

**12. Objaw: „mam funkcję zapisującą 80 linii — w jednej funkcji, bo inaczej nie da się jej przetestować."**
→ *Przyczyna:* **odwrócenie zależności logicznej.** Testowanie nie wymaga jednej funkcji — wymaga **granicy, na której można wstrzyknąć fałszywy obiekt**. Jeżeli funkcja tworzy `Workbook()` w środku i zapisuje na dysk, to **na sztywno** związała się z openpyxl — i dlatego nie da się jej przetestować inaczej niż plikiem.
→ *Naprawa:* **wstrzyknij eksportera** (przykład 4). Podziel 80 linii na `buduj_widok` (świat danych, **w pełni testowalny**), `wypelnij_arkusz` (świat zapisu, testowalny przez `BytesIO`) i `zapisz_raport` (I/O, testowalny przez fake). Kluczowe: **najtrudniejsza logika wychodzi z funkcji zapisu do funkcji czystej**, a wtedy testy są szybkie i nie potrzebują dysku.

**13. Objaw: „wybrałem poziom abstrakcji raz i trzymam się go, choć projekt się zmienił — przecież architektury się nie zmienia."**
→ *Przyczyna:* **traktowanie poziomu abstrakcji jako decyzji jednorazowej.** A to jest decyzja **zmienna w czasie**: poziom 0 → 1 → 2 → 3 przy wzroście, ale również 2 → 1 przy spadku (raport przestał być potrzebny, zespół się rozpadł, wymagania stopniały do jednego odbiorcy).
→ *Naprawa:* **rewiduj poziom przy zmianie wymagań.** W praktyce: raz na kwartał zadaj pytanie „czy wszystkie komponenty mojej struktury są jeszcze używane?". Usuwanie abstrakcji jest **tak samo inżynierskie** jak jej dodawanie. A konkretny sygnał: **jeśli jakiś element struktury nie ma testu albo nikt go nigdy nie rozszerzał od pół roku, to kandydat do usunięcia.**

## 7. Podsumowanie — model mentalny w 5 punktach

1. **Skrypt „na kolanie" to rusztowanie. Świetne na jeden dzień, śmiertelne po roku.** Historia z 900 liniami nie jest opowieścią o głupocie, a o **braku nazw i struktur**: żadna z dziewięciu decyzji nie była głupia w momencie, w którym zapadała. Problem polegał na tym, że w kodzie nie było ani nazwy `PROG_ALERTU`, ani nazwy `TLO_NAGLOWKA`, ani podziału na funkcje. **Nazwa pozwala myśleć — a bez myślenia nie da się zmienić nagłówka bez czytania 900 linii.** Wzorce to oznaczone pudełka na narzędzia: nie szukasz nasłuchując, widzisz, czego nie ma, i możesz podać pudełko komuś innemu.

2. **Poziom abstrakcji wybierasz na podstawie NASTĘPNEJ zmiany — nie na podstawie wyobrażeń.** Poziom 0 wystarcza na jednorazowy raport, 1 na powtarzalny, 2 na trzy warianty, 3 na integrację i audyt. **Przejście jest wywołane zdarzeniem, nie ambicją.** YAGNI nie znaczy „nigdy nie abstrahuj" — znaczy „abstrahuj, gdy masz dowód". A dowodem jest konkretny, istniejący przypadek: **reguła trzech** — pierwszy raz napisz wprost, drugi raz zauważ, trzeci raz wyciągnij wspólny kod.

3. **Abstrakcja musi być opłacona testem. Abstrakcja bez testów jest gorsza niż duplikacja.** Duplikacja jest **widoczna** (dwie różne wartości krzyczą), abstrakcja **ukrywa** duplikację (skoro jedna funkcja, to „przecież spójne") i **mnoży błąd**, jeżeli jest źle napisana. Dlatego kolejność jest zawsze taka sama: funkcje czyste → testy → abstrakcja. Odwrotna kolejność daje strukturę, której nikt nie umie zmienić, bo nikt nie wie, co ona robi.

4. **Trzy światy to fundament wszystkiego, co dalej: dane, prezentacja, zapis.** Architekt decyduje, *co* ma być; ekipa wykończeniowa decyduje, *jak to wygląda*; dostawca materiałów *dowozi konkretne rzeczy*. **Możesz zmienić dostawcę, nie zmieniając projektu domu.** Reguła operacyjna: `import openpyxl` wolno **tylko** w świecie zapisu. Trzy pytania rozstrzygające przynależność: czy zmieni się przy PDF? czy zmieni się przy innym źródle danych? czy zniknie bez openpyxl? I najważniejsze: **prezentacja nie pyta „czy cena > 30?", tylko „czy mam pokazać alert?"** — decyzja zapada w świecie danych i przychodzi **gotowa**.

5. **OCP jest mierzalne: „ile linii dopisałeś do istniejącego kodu?".** Nowy raport to **nowe dane** (`Specyfikacja`), a nie nowy `elif` w `generuj(raport, dane)`. Robisz raport regionalny i dopisujesz **zero** linii do klasy raportu podstawowego → OCP działa. Dopisujesz `if spec.tytul == ...` albo parametr `rodzaj_raportu` → OCP złamane, a Twoja funkcja zaczęła się zamieniać w to, przed czym ten moduł ostrzega. I pamiętaj o zakamuflowanych `if`-ach: **flaga w konfiguracji (`czy_wysylac_maila`) to `if` w kodzie, tylko przebrany za czystość.** Flaga → osobny adapter.

## 8. Ściągawka modułu

```python
# ==================================================================
# 1. NARZEDZIE: HEURYSTYCZNY SKANER SYGNALOW OSTRZEGAWCZYCH
#    Uzyj go na WŁASNYM kodzie, nie na cudzym. Wynik to metryka.
# ==================================================================
import re
from pathlib import Path

WZORCE = {
    "kolor_hex":              re.compile(r'"[0-9A-Fa-f]{6}"'),
    "import_styles":          re.compile(r"^from openpyxl\.styles import", re.M),
    "magiczny_indeks_kolumny": re.compile(r"column=\d+"),
    "magiczny_indeks_wiersza": re.compile(r"row=\d+"),
    "prog_w_warunku":         re.compile(r"[<>]=?\s*\d+(?:\.\d+)?"),
    "zapis_na_poziomie_modulu": re.compile(r"^wb\.save\(", re.M),
}


def wykryj_sygnaly(sciezka) -> dict:
    """Heurystyka. Male liczby = mniej sygnalow. Zero nie istnieje."""
    tekst = Path(sciezka).read_text(encoding="utf-8")
    liczby = {n: len(w.findall(tekst)) for n, w in WZORCE.items()}
    liczby["def"] = len(re.findall(r"^def ", tekst, re.M))
    liczby["class"] = len(re.findall(r"^class ", tekst, re.M))
    return liczby


# ==================================================================
# 2. TABELA POZIOMOW - skrot do wydrukowania
# ==================================================================
# 0: 1 plik, top-level        | jednorazowy raport         | 0 testow
# 1: funkcje + konfiguracja   | powtarzalny, zmienne kolumny| testy czyste
# 2: specyfikacja + klasa     | 3+ warianty, OCP           | testy OBOWIAZKOWE
# 3: porty + adaptery         | integracja, audyt, formaty | testy bez plikow
#
# PRZEJSCIE WYZWALANE ZDARZENIEM:
#   drugi odbiorca / drugie uruchomienie  -> 1
#   3 warianty albo "to samo, ale inaczej" -> 2
#   inny format wyjscia / integracja / audyt -> 3


# ==================================================================
# 3. TRZY PYTANIA - do ktorego swiata nalezy ten kod?
# ==================================================================
# Q1: Zmieniloby sie, gdybysmy wysylali raport do PDF?
#     tak  -> PREZENTACJA (albo ZAPIS, jesli zniknie calkiem)
# Q2: Zmieniloby sie, gdyby dane przyszly z bazy zamiast CSV?
#     tak  -> PREZENTACJA albo DANE (zalezy czy to uklad, czy regula)
# Q3: Znikneloby, gdyby openpyxl przestal istniec?
#     tak  -> ZAPIS
#     nie  -> DANE albo PREZENTACJA
#
# DOMYSLNY KIERUNEK W RAZIE WATPLIWOSCI: przenies do SWIATA DANYCH.
# Powod: dane nie zaleza od nikogo, wiec tam zawsze bezpiecznie.


# ==================================================================
# 4. REGULY OPERACYJNE (sprawdzalne grepem!)
# ==================================================================
# SWIAT DANYCH nie importuje openpyxl:
#     grep -rn "openpyxl" src/domena/     -> 0 wynikow
# REGULA wystepuje w JEDNYM miejscu:
#     grep -rn "cena >" src/               -> 1 wynik
# LITERALY (kolory, progi, rozmiary) tylko w konfiguracji:
#     grep -rn 'FFC000\|FF0000\|> 30' src/ -> tylko config.py
# SWIAT ZAPISU nie ma rozgalezien po typie raportu:
#     grep -n 'elif\|rodzaj_raportu\|typ_raportu' src/zapis.py -> 0


# ==================================================================
# 5. TESTOWALNOSC: trzy poziomy szybkosci testow
# ==================================================================
# (a) FUNKCJE CZYSTE (dane)  - bez openpyxl, bez I/O, mikrosekundy
#     assert wymaga_alertu(WierszSprzedazy(..., cena=41.5), 30.0) is True
#
# (b) ZAPIS DO BytesIO (modul 20) - milisekundy, zero plikow
#     wb = zbior_raport(...); buf = io.BytesIO(); wb.save(buf)
#     wb2 = load_workbook(io.BytesIO(buf.getvalue()))
#
# (c) ADAPTER FAKE (przyklad 4) - milisekundy, ZERO openpyxl
#     class EksporterDoTestow:
#         def __init__(self): self.wywolania = []
#         def eksportuj(self, wiersze, cel):
#             self.wywolania.append((len(wiersze), cel))


# ==================================================================
# 6. OCP - TEST STRUKTURY, NIE ZACHOWANIA
# ==================================================================
import inspect

def test_zapisz_raport_nie_rozgalezia_sie_po_typie_raportu():
    zrodlo = inspect.getsource(zapisz_raport)
    zabronione = ("elif", "rodzaj_raportu", "typ_raportu", "nazwa_raportu")
    for w in zabronione:
        assert w not in zrodlo, f"Zlamane OCP: {w!r} w zapisz_raport"


# ==================================================================
# 7. SZKIELET: TRZY SWIATY W JEDNYM PLIKU
# ==================================================================
# ── KONFIGURACJA (dane, ktore moga pochodzic z YAML) ──────────────
# @dataclass(frozen=True) class Kolumna(...)      # nazwa, szerokosc, fn, fmt
# @dataclass(frozen=True) class Motyw(...)        # kolory, rozmiary
# @dataclass(frozen=True) class Specyfikacja(...) # tytul, kolumny, regula
#
# ── SWIAT DANYCH (bez openpyxl!) ─────────────────────────────────
# @dataclass(frozen=True) class Wiersz...         # model (tuple, nie list!)
# def wczytaj_...(...) -> list[Wiersz]            # zrodlo
# def regula_...(w, prog) -> bool                 # funkcja CZYSTA
# def agreguj_...(wiersze) -> list[...]           # przeksztalcenie
# def zrob_projekcje(...) -> list[WierszWidoku]   # wartosci + decyzje
#
# ── SWIAT ZAPISU (tylko openpyxl, zero decyzji) ──────────────────
# def wypelnij_arkusz(wb, tytul, spec, widok) -> Worksheet
# def zapisz_atomowo(wb, cel: Path) -> Path       # mkstemp + os.replace
#
# ── SWIAT PREZENTACJI (wykonuje, nie decyduje) ───────────────────
# def sformatuj_arkusz(ws, spec, widok) -> None   # czyta wp.alert
#
# ── ZLOZENIE (jedyne miejsce, gdzie swiaty sie stykaja) ──────────
# def zapisz_raport(wiersze, spec, cel) -> Path   # ZERO if-ow o typie


# ==================================================================
# 8. PULAPKI DO ZAPAMIETANIA
# ==================================================================
# - frozen=True chroni PRZYPISANIE, nie MUTACJE zawartosci
#   -> w polach uzywaj tuple / frozenset / Mapping, nie list / set / dict
# - @property NIE jest polem: nie ma go w fields(), asdict(), repr()
# - dataslass default= tylko dla obiektow NIEMUTOWALNYCH
# - getattr(obj, "nazwa") - literowka wyjdzie dopiero w runtime
# - dict.get("klucz", "") cicho POŁYKA literowke (widoczna pusta kolumna)
# - Flaga w konfiguracji (czy_wysylac_maila) = zakamuflowany if w kodzie
# - Konfiguracja, ktora wymaga debuggera, nie jest konfiguracja
# - if po typie raportu = zlamane OCP (nowy raport ma byc DANYMI)
# - Funkcja z 3+ parametrami konfiguracyjnymi obsluguje 2 przypadki
# - Refaktoryzacja bez testu "golden" = nie wiesz, czy sie udala
# - Podzial bez granic (utils.py/helpers.py) to nie podzial
# - Usuniecie abstrakcji jest tak samo inzynierskie jak jej dodanie
```

## 9. Co dalej

Zatrzymaj się i policz, co masz po tym module. Nie napisałeś ani jednego wzorca. **I to jest w porządku** — bo masz coś ważniejszego: masz **kryteria**.

Masz sześć nazw na to, co czułeś, ale nie umiałeś powiedzieć. Masz tabelę czterech poziomów i — co najważniejsze — **kolumnę „kiedy przejść wyżej" opartą na zdarzeniach, a nie na ambicji**. Masz trzy światy i trzy pytania, które rozstrzygają sporne przypadki bez zgadywania. Masz mierzalne kryterium OCP („ile linii dopisałem do istniejącego kodu?") i mierzalne kryterium jakości (`grep -rn "openpyxl" src/domena/`). I masz jedną zasadę, która przez całą Część IV będzie Twoim bezpiecznikiem: **abstrakcja bez testów jest gorsza niż duplikacja.**

Zauważ jeszcze jedną rzecz, bo jest subtelna i ważna. **Ten moduł był o strukturze, a wróciłeś w nim trzy razy do modułu 21.** Bo reguła „biała lista zamiast czarnej" z sekcji 2.2 tamtego modułu to nic innego jak **OCP**: zamiast utrzymywać rosnącą listę rzeczy do usunięcia, deklarujesz to, co ma zostać. A zdanie „sanitacja na wejściu do domeny, nie w momencie zapisu" to zapowiedź **granicy architektonicznej**, którą zobaczysz w module 26. **Bezpieczeństwo i struktura to ten sam temat, opowiedziany z dwóch stron.**

W module 23 zaczyna się właściwa część o wzorcach — i robimy to **na dokładnie tym samym przykładzie**, który widziałeś tutaj. Zobaczysz cztery wzorce konstrukcyjne, każdy z **konkretnym sygnałem z sekcji 2.1 jako uzasadnieniem**:

- **Factory** rozwiązuje sygnał 1: `create_sheet` + `title` + `index` + `tabColor` w dwudziestu miejscach zamienia się w **jedną receptę**.
- **Builder** rozwiązuje problem złożoności z przykładu 3: raport składany w sześciu krokach, gdzie **kolejność i kompletność** mają znaczenie. Zobaczysz też, czym Builder **nie jest** (bo to najczęściej nadużywany wzorzec w automatyzacji raportów).
- **Prototype** rozwiązuje problem szablonu skoroszytu: pieczątka firmowa, nadruk gotowy, dopisujesz treść — i wzorzec **nigdy nie jest nadpisywany**. Wrócimy tu do modułu 18, bo szablon z sparkline'ami albo makrami to ta sama pułapka w nowym przebraniu.
- **Singleton / Registry** rozwiązuje sygnał 2 i to, o czym mówił moduł 09: `"FFC000"` w czterdziestu miejscach albo `Font` tworzony tysiąc razy. Zobaczysz **rejestr stylów** — i jego pułapkę: stan globalny utrudnia testy, więc rejestr trzeba **wstrzykiwać**, a nie importować.

I zobaczysz tam wzorzec, który jest kłopotliwy: **Builder, który i tak pisze do arkusza w trakcie budowy — i przez to traci cały sens.** To najczęstszy błąd przy pierwszym zetknięciu z tym wzorcem i warto go zobaczyć **przed** tym, jak go popełnisz.

Ale zanim tam pójdziemy, zrób jedno. Uruchom skaner z sekcji 8 na **swoim najważniejszym skrypcie**. Nie na tym z kursu — na tym, który naprawdę utrzymujesz. Zapisz liczby. To jest Twój punkt odniesienia i jedyna uczciwa miara tego, ile Cię jeszcze czeka. I pamiętaj, co mówi sekcja 1.4: **jeżeli liczby są małe i raport jest jednorazowy, to nie masz co refaktoryzować.** Wróć do tego kursu za miesiąc i porównaj — to będzie jedyny wskaźnik, który naprawdę się liczy.