# Word AI Secure Live — sesja dokumentu v2

Stan: 2026-09-18. Fork `jpierzchala/word-ai`, baza `1edacba40e89073b0e03ec6dcf0fcc380dbcb33f`.
Zmiany lokalne na `main`, do review przed commitem. Nie wykonano push ani publikacji obrazu.

## Codzienna praca

1. Otwórz i zapisz dokument Word. Dodaj **Word AI Secure Live** z katalogu Shared Folder.
2. Kliknij **Udostępnij dokument AI**. Nie zaznaczaj fragmentów ani nie kopiuj tokenów.
3. Powiedz Codexowi lub Claude Code, co zmienić. Globalny skill `word-live` łączy się
   z tym dokumentem i prowadzi odczyt → podgląd → zapis → sprawdzenie.
4. **Cofnij dostęp** kończy sesję. Przeładowanie panelu, zmiana tożsamości dokumentu
   i restart mostka wymagają ponownego udostępnienia. Sam zapis w pamięci Worda
   nie oznacza zapisu pliku ani ukończonej synchronizacji SharePoint.

Zgoda obejmuje odczyt udostępnionego dokumentu i zlecone zmiany przez czas sesji,
bez potwierdzania każdej zmiany w panelu. Użytkownik wyraźnie wybrał ten model
2026-09-16: „Tak — zgoda na sesję dokumentu”. Dostęp nie zastępuje zlecenia redakcji.
Tekst odczytany przez MCP trafia do klienta i jego usługi modelowej.

## Zakres

Wymagany Word Microsoft 365 z **WordApi 1.6**. Snapshot odczytuje akapity body
(w tym komórki tabel), nagłówków, stopek, przypisów dolnych i końcowych. Kształty
oraz pola tekstowe nie są enumerowane. Nie ma dostępu do innych otwartych dokumentów,
bibliotek ani katalogów dysku bez odrębnego udostępnienia ich sesji.

Operacje: celowana zamiana unikalnego tekstu, zamiana prostego akapitu o jednolitym
formatowaniu, wstawienie akapitu przed/po i nadanie stylu Normal/Title/Subtitle/Heading1–9.
Zamiana celowana zachowuje format poza dopasowaniem; nowy tekst dziedziczy format
początku dopasowania. Pola, komentarze, rewizje, zakładki i obiekty w akapicie blokują
nadpisanie jego tekstu. Obrazy inline można wstawiać, podmieniać w miejscu albo usuwać;
kształty pływające nie są obsługiwane. Od 2026-09-23 także operacje strukturalne:
usunięcie akapitu, formatowanie fragmentu (pogrubienie/kursywa/podkreślenie), listy
(numeracja od 1, typ poziomu, poziom elementu), dodanie wiersza tabeli i przeniesienie
sekcji nagłówka (opis niżej). Nie ma edycji kolumn i scalania tabel, pól, definicji stylów
ani automatycznego wyłączania śledzenia zmian. Włączone śledzenie blokuje zapis.

## Instalacja i klienci

`compose.secure.yaml` uruchamia `word-ai-secure:local`, kontener `word-ai-secure`.
Do odtworzenia instalacji:

```powershell
& 'C:\BI Git\word-ai\scripts\Start-SecureWordAi.ps1' -Build -TrustCertificate
```

Certyfikat localhost (ważny sześć miesięcy) jest zaufany w CurrentUser/Root.
Oryginalne sekrety z 2026-09-16 powstały z Codexa (MSIX), więc fizycznie leżą
w `%LOCALAPPDATA%\Packages\OpenAI.Codex_…\LocalCache\Local\WordAiSecure\secrets`.
2026-09-23 skrypt uruchomiony z powłoki spoza MSIX nie znalazł ich w prawdziwym
`%LOCALAPPDATA%`, wygenerował nowy certyfikat i wgrał go do wolumenu; Word zablokował
panel („isn't signed by a valid security certificate”). Przywrócono oryginały (odcisk
`A38612E9…`) i skopiowano je także do prawdziwego `%LOCALAPPDATA%\WordAiSecure\secrets`,
żeby obie ścieżki dawały ten sam zaufany certyfikat. Przed `-Build` z innej powłoki
sprawdź, że widzi ona istniejące sekrety.
Sekrety TLS są poza repo/OneDrive w `%LOCALAPPDATA%/WordAiSecure/secrets`, chronione
ACL użytkownika i SYSTEM, i w prywatnym wolumenie Docker. Historyczny pairing-token
może pozostać w tym katalogu, lecz v2 go nie czyta. Nie używaj starego skryptu kopiowania tokena.

Oba klienty mają użytkownikowy serwer `word_ai_secure`:
`C:/ProgramData/anaconda3/python.exe C:/BI Git/word-ai/scripts/start_secure_mcp.py`.
Launcher uruchamia Rancher Desktop w tle na żądanie, czeka na Docker i uruchamia
wyłącznie istniejący kontener o właściwych etykietach Compose. Nie buduje go automatycznie.
Codex ma timeout startowy 210 s. Claude Code ma `MCP_TIMEOUT=210000` w osobistych
settings.json; ten timeout dotyczy również startu innych MCP, bez zmiany ich uprawnień.

Codex `enabled_tools` zawiera trzynaście narzędzi v2 (podmiana i usuwanie obrazu od
2026-09-18, przeniesienie sekcji i wiersz tabeli od 2026-09-23); `word_session_apply.approval_mode`
wynosi `approve` zgodnie z autoryzacją sesyjną. Claude ma allow dla dokładnie tych
trzynastu nazw `mcp__word_ai_secure__...`; inne pozwolenia i ustawienia modeli zachowano.
Gdy przebudowa zamknie transport, odśwież połączenie MCP w kliencie. Zapis konfiguracji
nie jest dowodem przeładowania narzędzi w już trwającej rozmowie.

Canonical skill: `C:/BI Git/bi-shared/.claude/skills/word-live/SKILL.md`.
Globalne junctiony `~/.claude/skills/word-live` i `~/.agents/skills/word-live` wskazują
ten sam katalog. Aktywacja: `/word-live` w Claude Code, `$word-live` lub naturalna
prośba o pracę w otwartym Wordzie w Codexie. Koszt always-on: nazwa/opis skilla
oraz trzynaście schematów MCP; pełna instrukcja jest czytana na żądanie. Bez nowego modelu,
harmonogramu ani zmian zespołowych ustawień runtime.

### Jednorazowy katalog dla zwykłych dokumentów

```powershell
& 'C:\BI Git\word-ai\scripts\Install-SecureWordAiCatalog.ps1'
```

Uruchom skrypt ze zwykłej sesji swojego użytkownika. Przygotowuje folder
`%LOCALAPPDATA%\WordAiSecure\catalog` z jednym plikiem `manifest.xml`, następnie
ustala fizyczną ścieżkę przez `GetFinalPathNameByHandle`. Codex jako aplikacja MSIX
może przekierować logiczny AppData do `Packages/<rodzina-pakietu>/LocalCache/Local`;
usługi Windows nie widzą tej samej nakładki ścieżek co proces aplikacji.

Instalator wykorzystuje istniejący udział dysku, np. `C$`, i sprawdza odczyt
dokładnie tego samego manifestu przez bieżącego użytkownika (SHA-256).
Nie tworzy udziałów, nie zmienia ich uprawnień, usług ani zapory i nie wymaga UAC.
Jeśli istniejący udział nie jest dostępny, kończy się błędem bez rozszerzania dostępu.
Katalog Office wskazuje tylko podfolder z manifestem, nie katalog sekretów.

Skrypt zapisuje katalog w rzeczywistym profilu użytkownika:
`HKEY_USERS/<SID bieżącego użytkownika>/Software/Microsoft/Office/16.0/WEF/TrustedCatalogs`.
Używa `StdRegProv` z uprawnieniami wywołującego, ponieważ zapis HKCU z MSIX może
trafić do prywatnej nakładki niewidocznej dla Worda. Ustawia `Flags=1`
(„Show in Menu”) i sprawdza wartości po zapisie. Nie zapisuje nic w profilu innej
osoby ani w `HKEY_LOCAL_MACHINE`.

Po zakończeniu instalacji uruchom Word ponownie i wybierz:
**Home → Add-ins → Advanced → Shared Folder → Word AI Secure Live → Add**.
Dodatek można wtedy wstawić do bieżącego dokumentu, bez kopiowania go do pliku
testowego. Katalog wskazuje fizyczny podfolder `WordAiSecure/catalog` przez
istniejący udział lokalnego komputera. Opcja `-PrepareOnly` sprawdza manifest,
dostęp i rzeczywisty profil bez rejestracji katalogu. Zaktualizowany manifest o innej treści wymaga świadomego
przeglądu istniejącej kopii; instalator jej automatycznie nie nadpisuje.

Weryfikacja 2026-09-16: w zwykłym nowym `Document1` galeria **SHARED FOLDER**
pokazała **Word AI Secure Live**. Użytkownik dodał dodatek i potwierdził działający
panel zrzutem ekranu. Pierwotne próby utworzenia osobnego udziału nie powiodły się
(`New-SmbShare` i `NetShareAdd`: błąd 2), ponieważ używały logicznej ścieżki MSIX.
Usunięto nieskuteczny etap administratora. Żaden dodatkowy udział nie powstał.

Jest to metoda lokalnego pilota na Windows. Microsoft klasyfikuje katalog folderowy
jako sposób testowania dodatków, nie wdrożenie produkcyjne organizacji:
[instrukcja Microsoftu](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

## Obrazy z lokalnych plików

`word_session_preview_image` przyjmuje `image_path`, `session_id`, `paragraph_id`,
`expected_sha256`, `position` (before/after), `width_pt`, `caption` i `alt_text`.
Następnie używa się istniejących apply/status. Obraz trafia do nowego akapitu jako
inline; opcjonalny zwykły podpis jest w kolejnym akapicie. Nie ma opływania tekstem,
automatycznej numeracji podpisu.

Snapshot zwraca dla każdego akapitu `inline_picture_count` oraz listę
`inline_pictures` z zerowym `image_index`, wymiarami i tekstem alternatywnym.
`word_session_preview_replace_image` podmienia wskazany obraz inline w tym samym
miejscu. `width_pt=0` zachowuje dotychczasową szerokość; wartość 24–500 ustawia nową.
`word_session_preview_delete_image` usuwa tylko wskazany obraz. Obie operacje wymagają
aktualnego `expected_sha256` akapitu. Podpis w sąsiednim akapicie pozostaje bez zmian,
więc jego usunięcie lub korekta wymaga osobnej, jawnej operacji tekstowej.

Obsługiwane PNG/JPEG: do 2 MiB, do 8192 px na bok i 20 MP. Szerokość 24–500 pt,
wysokość proporcjonalna do 700 pt; 72 pt = 1 cal. Podpis do 500 znaków, alt do 1000,
pojedyncze linie. Dobierz rozmiar do marginesów/komórki. Nietypowe pliki, np. JPEG
z orientacją EXIF zmieniającą wymiary dekodowania, mogą wymagać eksportu do PNG.

Nowy osobisty adapter `scripts/image_mcp_adapter.py` odczytuje dokładnie wskazany
regularny plik z lokalnego dysku. Odrzuca URL, UNC, mapowany dysk sieciowy, ADS,
dowiązania i reparse points. Nie przegląda katalogów, nie montuje ich w Dockerze.
Wąski adapter zamienia ścieżkę na Base64 poza kontekstem modelu i przekazuje dane
przez lokalny STDIO. Narzędzia zwracają tylko metadane/sha256, nie bajty obrazu.
Kontener ponownie sprawdza format/rozmiar/piksele, a panel dekoduje obraz i sprawdza hash.
Obraz pozostaje w RAM, obowiązuje TTL podglądu i globalny budżet zakolejkowanych obrazów.
Istniejąca zgoda sesyjna dotyczy zapisu w Wordzie; agent nadal potrzebuje zlecenia
użycia konkretnego pliku. Osobne obejrzenie pliku przez agenta może przekazać jego
zawartość do modelu. Ta ścieżka zmienia opis wcześniejszego profilu bez odczytów plików
na hoście; izolacja systemu plików kontenera pozostaje bez zmian.

Podgląd wiąże konkretne bajty, miejsce, hash akapitu, wymiary, alt i podpis. Zmiana
pliku po podglądzie nie podmienia zawartości operacji. Word może częściowo wykonać
batch przed błędem; stan unknown blokuje ponawianie, nie obiecuje rollbacku.

Testy rozszerzenia: 31 Python i 43 JavaScript PASS. Rzeczywisty hostowy adapter
udostępnił schemat image_path bez Base64 i przesłał PNG 480×240. Word wstawił PNG
przed akapitem KEEP LAST w rozmiarze 288×144 pt z podpisem i alt; JPEG za tym akapitem
w rozmiarze 144×72 pt bez podpisu. Obie operacje zwróciły succeeded po odczycie właściwości.
Próba kontroli wizualnej i zapisu pliku została przerwana przez użytkownika klawiszem
Escape. Nie deklaruje się weryfikacji osadzenia tych obrazów w zapisanym DOCX.

Odbiór podmiany 2026-09-18: nowy PNG (1 290 580 bajtów) zastąpił Figure 2
w Executive Report - Part 1. Pobrany po zapisie DOCX z SharePoint zawierał
`word/media/image2.png` o identycznym SHA-256 jak plik wejściowy, szerokości 450 pt
i wysokości 337,5 pt. Samo apply zgłosiło `unknown` po skutecznym zapisie, ponieważ
kontrola po zapisie odrzuciła eksport XML przez stary limit 2 mln znaków. Nie ponowiono
mutacji. Limit eksportu akapitu podniesiono do 8 Mi znaków (Base64 obrazu i zależności
formatowania); limit wejściowego PNG/JPEG pozostaje 2 MiB. Dodano test regresji
podmiany z eksportem Flat OPC obrazu 2 MiB oraz test granicy rozmiaru i DTD.
Błąd parsera snapshotu jest teraz jawną luką pokrycia. Odbiór usuwania w Wordzie
pozostaje otwarty.

Diagnostyka sesji: dwa restarty podczas wdrażania obsługi obrazów unieważniły sesje
oraz istniejący transport MCP. Panel otwarty przed aktualizacją nadal wykonywał
stary JavaScript do przeładowania. Świeży launcher połączył się poprawnie.
Oczekujący pairing wygasa po 300 s, a połączona sesja po 900 s bez kontaktu panelu;
45 s ukrywa nieaktywną sesję na liście, ale jej nie usuwa. Test zegara potwierdził,
że regularny poll utrzymuje połączoną sesję ponad 20 minut. To nie wyklucza innych
przyczyn zawieszenia panelu; nie zmieniano limitów zgody ani retencji.

## Operacje strukturalne (2026-09-23)

Wszystkie nowe operacje idą tym samym trybem: podgląd → hash → zapis → sprawdzenie.
Podgląd nie zmienia dokumentu. Oprócz hasha akapitu zwraca `guard_sha256`, czyli odcisk
otoczenia (sąsiedzi, lista, tabela albo cała sekcja). Mostek przyjmuje podgląd tylko
z tym odciskiem i przekazuje go do zapisu. Zapis wylicza wszystko od nowa i odmawia,
gdy odcisk się różni. Po zapisie panel sprawdza wynik; niezgodność po wysłanej mutacji
daje `unknown` (Ctrl+Z w Wordzie cofa zmianę; nie ponawiać automatycznie). Wyjątek: gdy
Word odrzuci zapis operacji na liście, a świeży odczyt po błędzie pokaże stan identyczny
z podglądem (ten sam odcisk, numery wszystkich elementów list i układ akapitów), wynik to
`failed` z `result.unchanged=true`; sesja nie jest blokowana.

Przez `word_session_preview` (pole `text` jest parametrem, jak przy `set_style`):

- `delete_paragraph` (`text=""`, `find=""`): usuwa jeden akapit. Blokują: ostatni akapit
  obszaru (dokumentu, komórki, nagłówka, przypisu), podział sekcji, kontrolka zawartości,
  komentarze, rewizje, pole sięgające poza akapit, akapit rozdzielający dwie tabele oraz
  obiekty pływające, wykresy i obiekty osadzone (snapshot ich nie pokazuje) oraz akapit
  z numerem przypisu (`footnoteRef`). Pola są sprawdzane w kolejności dokumentu: blokuje też
  akapit, w którym jedno pole się kończy, a drugie zaczyna.
  Podgląd ujawnia pola, zakładki, przypisy i obrazy, które znikną razem z akapitem.
  Kontrola po zapisie porównuje sekwencję akapitów obszaru nadrzędnego.
- `format_text`: `find` = unikalny fragment (puste = cały akapit), `text` = lista po przecinku
  z `bold`, `italic`, `underline`, `no_bold`, `no_italic`, `no_underline`. Zmienia tylko
  wskazane właściwości zakresu (podkreślenie Single/None), bez przepisywania tekstu. Typowy
  lead-in: `replace_paragraph`, potem `format_text` na „Lead-in:”. Blokują rewizje
  i kontrolki zawartości w akapicie.
  `^` w `find` jest escapowane jako `^^` (Word interpretuje kody ^ także bez wildcardów;
  poprawka dotyczy również `replace_text`).
- `list_level` (`text="0"…"8"`): poziom tylko tego elementu.
- `list_type` (`text="bullet"|"number"`): typ poziomu elementu w całej liście
  (`setLevelBullet` Solid albo `setLevelNumbering` Arabic „1.”). Podgląd podaje liczbę
  elementów objętych zmianą; kontrola sprawdza, że inne listy się nie zmieniły.
- `list_restart` (`text=""`): numeracja od 1 od tego elementu; podgląd podaje `method`.
  Poziom 0 (także pierwszy element listy): `start_override`, czyli to samo co „Uruchom ponownie
  od 1” w Wordzie (szczegóły w sekcji „Restart numeracji listy”). Pierwszy element głębszego
  poziomu: `set_starting_number` (`setLevelStartingNumber(level, 1)`), tylko gdy jego numer
  wynika z wartości początkowej poziomu (inaczej odmowa); środek listy na głębszym poziomie jest
  odrzucany. `separateList()` nie jest używane. Podgląd odmawia z powodem, gdy
  restartu nie da się wykonać bezpiecznie (np. element kończy sekcję Worda, leży w tabeli, ma
  śledzoną zmianę właściwości). Kontrola wymaga odczytu wartości numeru (`listFormat.listValue`,
  WordApiDesktop 1.3); bez niego podgląd odmawia, zamiast kończyć poprawny zapis stanem
  `unknown`. Każda operacja na liście sprawdza też, że numeracja innych list się nie zmieniła.
  Numeracja nagłówków (styl Heading) jest blokowana: służy do tego `set_style`.
  Operacje na listach obsługują listy w treści głównej (nie w nagłówkach i przypisach).

Nowe narzędzia:

- `word_session_preview_table_row(session_id, paragraph_id, expected_sha256, position, cells)`:
  wiersz przed/po wierszu zawierającym wskazany akapit komórki (`TableRow.insertRows`).
  `cells` to tablica jednowierszowych tekstów (1–63, do 4000 znaków, łącznie 20000),
  dokładnie tyle, ile komórek ma wiersz. Blokują: tabela zagnieżdżona, scalone lub
  nieregularne komórki (`vMerge`, `gridSpan`, `isUniform=false`), kontrolki zawartości
  i śledzone zmiany w tabeli, wiersz nagłówka jako wzorzec. Nowy wiersz dziedziczy formatowanie wiersza odniesienia. Kontrola: liczba
  wierszy +1, bez zmiany liczby wierszy nagłówka, wartości nowego i wszystkich innych wierszy.
- `word_session_preview_move_section(session_id, paragraph_id, expected_sha256,
  target_paragraph_id, target_expected_sha256, position)`: przenosi nagłówek (Heading1–9
  albo styl własny z poziomem konspektu 1–9) z całą treścią do następnego nagłówka
  tego samego lub wyższego poziomu: podsekcje, tabele, obrazy, pola, przypisy.
  Cel leży poza sekcją, w treści głównej, poza tabelą; `position=after` nie może wskazywać
  miejsca przed tabelą ani za ostatnim akapitem dokumentu. Podgląd podaje pierwszy
  i ostatni akapit, liczbę akapitów, podsekcji, tabel, obrazów, pól, przypisów i zakładek,
  nagłówek kończący sekcję oraz akapit, przed którym trafi treść.

Przeniesienie to jeden `word_session_apply`: `getOoxml()` sekcji → `insertOoxml()` przed
celem → `delete()` oryginału. Batch Office.js nie jest transakcją, więc zapis ma dwa kolejne
batche: pierwszy wstawia kopię i w tej samej chwili ponownie eksportuje źródło; drugi usuwa
oryginał dopiero, gdy ten eksport ma ten sam odcisk co podgląd, a śledzony (`track()`)
zakres źródłowy nadal zawiera ten sam tekst. W przeciwnym razie oryginał zostaje (duplikat
zamiast utraty treści, stan `unknown`). Wynik sprawdza pełne porównanie sekwencji
akapitów (tekst, styl, poziom tabeli), liczby tabel, obrazów i sekcji oraz zakładek.
Eksport i odcisk dużej sekcji trwają sekundy, więc tuż przed zapisem panel jeszcze raz
lekko odczytuje całą treść, obrazy sekcji i cel. Zmiana tekstu, układu, obrazów lub celu
w tym czasie przerywa zapis (stale), zamiast nadpisać ją starszą kopią. Każdy kolejny
batch (sprzątanie styku, przywracanie zakładek, korekta numeracji) ponownie sprawdza
cofnięcie dostępu i tożsamość dokumentu.
Paczka ma usunięte `sectPr` sekcji źródłowej i pusty akapit-wartownik na końcu: Word z założenia
scala ostatni akapit wstawianego OOXML z akapitem docelowym (OfficeDev/office-js #2411,
#2914). Puste akapity, które Word zostawia na styku (wartownik oraz niejawny akapit za
tabelą zamykającą sekcję), są usuwane po sprawdzeniu, że to jedyna różnica (maks. 2).
Zakładki (w tym ukryte `_Toc`/`_Ref`) są wstawiane pod nazwami tymczasowymi i po
usunięciu oryginału przywracane: Word pomija wstawianą zakładkę, której nazwa jeszcze
istnieje, a bez tego odsyłacze REF pokazywały „Error! Reference source not found.”.
Zakładka zaczynająca się dokładnie w miejscu wstawienia (typowo `_Toc` docelowego
nagłówka) rozszerza się w Wordzie na wstawioną kopię; po przeniesieniu panel przywraca jej
pierwotny zakres i sprawdza tekst (`destination_bookmarks` w podglądzie).
Blokują: komentarze, rewizje, kontrolki zawartości, podział sekcji w środku, pole lub
zakładka przekraczające granicę sekcji, zakładka obejmująca miejsce docelowe, miejsce
docelowe wewnątrz wyniku pola (np. tuż za spisem treści; kolejne F9 skasowałoby sekcję),
zakładki o nazwach, których Office.js nie odtworzy (np. `getting-started` z Pandoca),
przypisy z zakładkami, komentarzami, rewizjami lub kontrolkami, sekcja kończąca się
niepustym ostatnim akapitem dokumentu (najpierw dodaj pusty akapit na końcu). Przeniesione akapity dostają nowe
`paragraph_id`; po zapisie ponów snapshot. Pola (numery podpisów, spis treści, odsyłacze)
aktualizuje się w Wordzie (Ctrl+A, F9).

Operacje strukturalne czytają i sprawdzają całą treść dokumentu, więc mają limit wykonania
60 s (jak snapshot), a zapis musi ruszyć w 20 s od begin. Zamiana tekstu zachowuje 15 s i 5 s.

Snapshot zwraca teraz `list_item` (`level`, `list_string`) dla elementów list, bo
numer nie jest częścią tekstu akapitu. Odcisk akapitu i sekcji pomija
`w:lastRenderedPageBreak` (pamięć podręczna podziału stron) oraz `wp14:anchorId`/`wp14:editId`
obrazów: walidacja COM wykazała, że Word losuje je przy każdym eksporcie tego samego,
niezmienionego zakresu, co dawało fałszywe „stale”.

Weryfikacja 2026-09-23: 38 testów Python i 78 JavaScript PASS, w tym model Office.js
przepuszczający podgląd i zapis wszystkich nowych operacji przez prawdziwy executor panelu
(z modelem rozszerzania zakładek jak w Wordzie). `scripts/check_secure_move_com.py`
na prawdziwym Wordzie (COM, osobna niewidoczna instancja, dokument syntetyczny) przeniósł
sekcję w górę, w dół, na koniec i z końca dokumentu, także sekcję zakończoną tabelą
i zwykłym pustym akapitem (6 przypadków). Zachował kolejność, style, tabelę, obraz,
przypis, sekcje, zakładki `_Ref`/`bm_test`/`_Toc` i wyniki pól REF/SEQ; dwa osobne eksporty
tego samego zakresu dały ten sam odcisk. COM używa tej samej ścieżki Flat OPC co
`getOoxml`/`insertOoxml`, ale nie jest testem dodatku Office.js. Przegląd kodu przez czterech
niezależnych recenzentów z adwersarialną weryfikacją znalazł m.in. cel w wyniku pola TOC,
rozszerzanie zakładek w celu, zakładki w przypisach i nazwy zakładek z `-`; wszystkie
potwierdzone zgłoszenia poprawiono i pokryto testami.

Odbiór w prawdziwym dodatku Word Office.js na syntetycznym DOCX potwierdził
usunięcie akapitu, formatowanie fragmentu, restart numeracji, zmianę typu i poziomu
listy oraz trzy przeniesienia sekcji (także z obrazem i tabelą). Po poprawce odczytu
w nowej paczce `Word.run` dodanie wiersza zwróciło `succeeded` i oczekiwaną liczbę
wierszy; wstawienie pustego akapitu również zwróciło `succeeded`. Zapisany DOCX
potwierdził obie zmiany. Spis treści pozostał bez odświeżenia pól.

## Restart numeracji listy (2026-09-24)

Zgłoszenie z 2026-09-23 (Word M365, dokument z SharePoint): `list_restart` na pierwszym
elemencie 9.1 przeszedł podgląd (`separate_list`, `current_value=8`), a zapis skończył się
„Microsoft Word: This command is not available.” (`Paragraph.separateList`), stanem
`unknown` i blokadą sesji, choć dokument się nie zmienił. Ten sam błąd dawało COM
(`SeparateList`, `ApplyListTemplateWithLevel`); `SeparateList` na pierwszym elemencie 9.2
działało, a `ExecuteMso("NumberingRestart")` działało na 9.1.

Przyczyna, odtworzona `scripts/check_secure_list_restart_com.py` (osobna niewidoczna instancja
Worda, dokumenty syntetyczne): Word numeruje według definicji listy (`w:abstractNum`), a nie
według jej wystąpienia (`w:num`). Dokumenty składane z części mają kilka wystąpień tej samej
definicji; numeracja ciągnie się przez nie (1–7, 8–11, 12–14). `SeparateList` odmawia dokładnie
na pierwszym w dokumencie akapicie wystąpienia, które kontynuuje wcześniejsze. Nie ma znaczenia
nagłówek, tabela, odległość, styl List Number ani numeracja bezpośrednia. Tam, gdzie działa,
dzieli tylko akapity tego samego wystąpienia, więc przy kilku wystąpieniach dalsze bloki
numeruje inaczej niż „od 1”. Przewidzenie odmowy wymagałoby eksportu wszystkich wcześniejszych
elementów listy, dlatego dodatek nie używa już `separateList`.

Metoda `start_override` robi to co „Uruchom ponownie od 1”: Word dodaje jedno wystąpienie
`w:num` tej samej definicji z `w:lvlOverride/w:startOverride=1` i zmienia `w:numPr` tylko tego
akapitu; dalsze elementy liczą 2, 3… Panel eksportuje akapit zakresem Whole (eksport Content nie
zawiera `w:pPr`; Office.js dokleja do eksportu Whole pusty akapit, który panel pomija po
sprawdzeniu tekstu). Buduje paczkę z pustym akapitem niosącym właściwości elementu z nowym
`w:numPr`, pustym wartownikiem i sklonowanym `w:num` ze `startOverride`, i wstawia ją
`getRange("Content").insertOoxml(paczka, "End")`. Office.js zachowuje znacznik akapitu docelowego
dla ostatniego wstawionego akapitu (odbiór na żywo), więc bez wartownika właściwości przepadają.
Z wartownikiem tekst elementu kończy się znacznikiem z paczki w nowym akapicie, a dawny znacznik
i dawne `paragraph_id` zostają na pustym akapicie tuż za nim. Panel sprawdza dokładnie ten
kształt (reszta układu bez zmian) i w drugim batchu, po ponownym sprawdzeniu dostępu, usuwa ten
pusty akapit; w tym samym batchu przed usunięciem odczytuje jego tekst i obrazy, więc treść wpisana
w międzyczasie daje błąd weryfikacji i `unknown` zamiast cichej utraty. Tekst, przebiegi, zakładki, komentarze i pola elementu nie są przepisywane, ale
element ma **nowe `paragraph_id`**; wynik podaje je w `paragraph_id` (dawne w
`previous_paragraph_id`). Definicję z paczki Word łączy z istniejącą po `w:nsid`.

Import OOXML (inaczej niż polecenie Worda) nie dodaje wystąpienia identycznego z istniejącym,
tylko podpina akapit pod istniejące: bez obsługi drugi restart w tej samej liście nic by nie
zmienił, a restart przed istniejącym restartem by go usunął (COM, przegląd kodu). Nowe
wystąpienie dostaje więc neutralny `lvlOverride` innego poziomu. W listach wielopoziomowych
jest to zdefiniowany poziom ze startem równym jego własnemu (widoczny w eksporcie). W listach
jednopoziomowych (np. List Number) jest to poziom niezdefiniowany z unikalną wartością startu:
Word go zachowuje i uwzględnia przy imporcie, ale nie eksportuje; taki poziom nie ma elementów,
więc numeracja się nie zmienia (COM oraz Office.js na żywo: kolejny restart w liście z już
istniejącym restartem dostał nowe wystąpienie). Odcisk podglądu obejmuje właściwości akapitu,
jego treść, definicję listy i wszystkie jej wystąpienia z eksportu Whole.

Podgląd odmawia z powodem, gdy: element kończy sekcję Worda (`w:sectPr`), ma śledzoną zmianę
w `w:pPr`, leży w kontrolce zawartości albo w tabeli (eksport Whole ostatniego akapitu komórki
obejmuje cały wiersz, a wstawienie w komórce nie zostało sprawdzone), eksport ma nieoczekiwaną
strukturę lub nie odpowiada tekstowi akapitu (powód zawiera liczby: akapity, przebiegi, długość
tekstu, bez treści), definicja listy nie ma `w:nsid`, poziom z OOXML różni się od poziomu listy,
albo element ma już numer 1. Odmowa proponuje „Uruchom ponownie od 1” w Wordzie lub uzgodniony
fallback COM. Podgląd odmawia też, gdy wystąpienie listy elementu ma własną wartość początkową
(np. z „Ustaw wartość numeracji”), a dalsze akapity używają tego wystąpienia: Word stosuje ją przy
pierwszym akapicie wystąpienia, więc po przeniesieniu elementu przeszłaby na następny element.

Office.js grupuje elementy w listy według wystąpienia, a Word liczy według definicji, więc restart
przesuwa także późniejsze listy Office.js tej samej definicji (np. restart Requirement 3 przesuwa
9.1 i 9.2). Panel eksportuje więc jednym zakresem akapity od elementu do ostatniego elementu listy
w dokumencie, przypisuje akapity numerowane do elementów list z Office.js (liczba musi się
zgadzać) i wyznacza przebieg: dalsze elementy tej samej definicji (po `w:nsid`) i poziomu, których
numery kontynuują numer elementu, do pierwszego płytszego elementu albo istniejącego restartu.
`renumbered_item_count` obejmuje cały przebieg; etykiety elementów podrzędnych z numerem nadrzędnym
(np. 9.1) zmienią się razem z nim. Przebieg i numery jego elementów wchodzą do odcisku podglądu.

Paczka niesie `w:pPr` elementu, którego nie widzi hash zakresu Content. Dlatego tuż przed
zapisem panel ponownie eksportuje zakres Whole i odmawia (stale), gdy odcisk się różni,
a w tej samej paczce co `insertOoxml` eksportuje go jeszcze raz: zmiana w chwili zapisu daje
błąd weryfikacji „właściwości akapitu zmieniły się w chwili zapisu” i `unknown`.

Kontrola po zapisie: element (pod nowym ID) ma numer 1 i ten sam tekst; wcześniejsze elementy
bez zmian; elementy przebiegu (w dowolnej liście Office.js) przesunięte o `current_value - 1`;
wszystkie pozostałe dalsze elementy (inne definicje, głębsze poziomy, elementy za istniejącym
restartem) bez zmian; układ
akapitów (ID i tekst) jak przed zapisem poza nowym ID elementu; eksport Whole elementu z tą samą
treścią i właściwościami poza `w:numPr`, które wskazuje nowe wystąpienie tej samej definicji
(`w:nsid`) ze `startOverride=1`, różne od każdego wcześniejszego (gdy odróżnia je tylko ukryty
override, rozstrzyga kontrola numerów); numeracja innych list bez zmian.

Klasyfikacja błędu (wszystkie operacje na listach): gdy zapis rzuci wyjątek, panel robi świeży
odczyt. Stan identyczny z podglądem (ten sam `guard_sha256` z nowym hashem akapitu, te same
teksty, poziomy i numery wszystkich elementów list, przy `list_restart` także ich wartości
`listValue`, ten sam układ akapitów) daje
`failed` z `unchanged=true` i komunikatem „Word odrzucił zmianę listy… dokument się nie
zmienił”. Mostek nie blokuje wtedy sesji. Gdy odczyt się nie uda albo cokolwiek się różni,
zostaje `unknown`; błąd po pierwszym batchu (np. przy usuwaniu pustego akapitu) jest zawsze
`unknown`. Mostek uznaje tylko dosłowne `unchanged: true`.

Przegląd kodu (trzy rundy, recenzenci z weryfikacją adwersarialną) znalazł m.in. scalanie
identycznych wystąpień przy imporcie, przebieg między listami Office.js, przekazanie własnej
wartości początkowej wystąpienia, nieskuteczne `setLevelStartingNumber` dla kontynuowanych numerów
i wyścig przy usuwaniu pustego akapitu; wszystkie potwierdzone zgłoszenia poprawiono i pokryto
testami.

Walidacja COM (`scripts/check_secure_list_restart_com.py`): macierz dostępności `SeparateList`
dla 7 układów (numeracja ze stylu, bezpośrednia, wielopoziomowa, drugie wystąpienie jedno-
i wielopoziomowe, przeplot, akapit z komentarzem/zakładką/polem) oraz sekwencje restartów
wykonywane tak jak w dodatku. Każda numeruje zgodnie z podglądem i nie zmienia innych akapitów,
stylów, zakładek, komentarzy, pól ani definicji list; w listach jednopoziomowych wynik jest
identyczny z `NumberingRestart`. W listach wielopoziomowych z kilkoma wystąpieniami własne
„Uruchom ponownie od 1” Worda potrafi skopiować definicję listy; dalsze bloki liczą wtedy według
starej (np. po a7 = 3 następuje b1 = 3). Metoda panelu zachowuje definicję, więc dalsze elementy
liczą od restartu, zgodnie z podglądem.

Odbiór na żywo (Word M365, dodatek Office.js, 2026-09-24) pokazał trzy rzeczy, których COM nie
ujawnił: Office.js grupuje elementy w listy według wystąpienia `w:num` (w dokumencie testowym
„Plan step A” był pierwszym elementem swojej listy, choć miał numer 8; w zgłoszeniu z 2026-09-23
Office.js pokazał 14 elementów w jednej liście, więc grupowanie zależy od budowy dokumentu), stara
ścieżka `set_starting_number` nie zmienia takiego numeru (zapis zakończony `unknown`, dokument
bez zmian), a eksport Whole i wstawianie OOXML działają inaczej niż w COM (dodatkowy pusty akapit;
właściwości ostatniego wstawionego akapitu odrzucone). Stąd: `start_override` dla każdego
elementu poziomu 0, tolerancja eksportu i przepływ z wartownikiem. Warianty sprawdzono w panelu
przez WebView2 DevTools na kopii dokumentu testowego (Word uruchomiony z lokalnym portem
debugowania tylko na czas testu, potem zamknięty).

Końcowy odbiór na żywo (świeży dokument z `make_list_restart_live_test_doc.py`, sterowanie przez
MCP): „Closing item X” odrzucony w podglądzie („akapit kończy sekcję Worda”); restart Requirement 3
(`renumbered_item_count=14`) dał 1, 2, 1–5 oraz przesunął dalsze listy tej samej definicji (9.1:
6–9, 9.2: 10–12, 9.4: 13–14); restart „Plan step A” dał 1–4 (9.2: 5–7, 9.4: 8–9); restart „Risk A”
dał 1–3 (9.4: 4–5); niezależna lista przez cały czas 1–3; ponowny restart „Plan step A” odrzucony
jako no-op. Każdy zapis zakończył się `succeeded` z nowym `paragraph_id`. W odzyskanym dokumencie
(„AutoRecovered”), w którym Word odrzucał każdy zapis (także `insertText`), wynik to `failed`
z `unchanged=true` bez blokady sesji. Testy: 113 JavaScript i 39 Python PASS; skrypt COM 0 porażek.

Dokument do odbioru na żywo tworzy `scripts/make_list_restart_live_test_doc.py`: lista 1–7,
przerwa z nagłówkami i tabelą, 9.1 (8–11) jako drugie wystąpienie tej samej definicji, 9.2
(12–14), niezależna lista 1–3 oraz 9.4 (15–16), gdzie pierwszy element kończy sekcję. Skrypt
sprawdza, że Word odmawia `SeparateList` na pierwszym elemencie 9.1.

## Kontrakt i zabezpieczenia

- Kliknięcie panelu tworzy metadata-only pairing, ważny 300 s. MCP przyjmuje dokładny
  `pairing_id` i adres dokumentu; panel odbiera sekret sesji przez tajny bilet.
  Same metadane pairing nie dają prawa odczytu Office API. Brak sekretów w localStorage,
  argumentach procesu, opisach MCP i wynikach listy.
- `word_session_list`, `word_session_connect`, `word_session_snapshot`, `word_session_read`,
  `word_session_preview`, `word_session_preview_image`, `word_session_preview_replace_image`,
  `word_session_preview_delete_image`, `word_session_preview_move_section`, `word_session_preview_table_row`,
  `word_session_apply`, `word_session_command_status`, `word_session_cancel` (13 narzędzi).
  Snapshot/read/preview/apply są asynchroniczne. Sprawdzaj wynik po command_id.
- Snapshot ma strony 1–50 akapitów, wyszukiwanie i jawny next_start. Read ma strony
  1–20000 znaków. Zapis wymaga ID akapitu odkrytego w tej samej sesji i aktualnego hasha.
  `uniqueLocalId` obowiązuje w lokalnej sesji Worda, nie jest trwałym ID w pliku.
- Preview jest jednorazowy, ważny 120 s i związany z treścią oraz formatowaniem akapitu,
  stylami/fontami/motywem eksportu. Techniczne rsid/proofErr pomijane. Przed zapisem
  ponowny odczyt i porównanie hasha. Word używa natywnych metod tekstu/akapitu,
  bez przebudowy pliku i bez wstawiania całego eksportu OOXML.
- Jedna komenda na sesję. Kolejka 90 s; wykonanie snapshot i operacji strukturalnych do 60 s,
  pozostałe do 15 s. Zapis musi zostać wysłany w ciągu 5 s od begin (strukturalny: 20 s).
  Stabilny request_id zapobiega duplikatom.
  Stan unknown blokuje dalsze komendy; nie ponawiać zapisu automatycznie.
- Cofnięcie dostępu blokuje zapis jeszcze niewysłany do Worda. Już wysłany może się
  zakończyć. Nie ma obiecanego rollbacku ani transakcyjnego compare-and-swap:
  nie edytować równocześnie tego samego akapitu z agentem.
- Dane w RAM, szczegóły komend usuwane po 600 s od utworzenia, sesja po 900 s bez
  aktywności; nieaktywna przez 45 s znika z listy aktywnych. Rozłączenie/restart usuwa stan.
  Maksymalnie 8 sesji, 8 ticketów, 1000 komend na sesję; po wyczerpaniu trzeba połączyć ponownie.
- Codex/Claude → STDIO docker exec → prywatny socket Unix. Word → HTTPS localhost.
  Brak HTTP /mcp, proxy URL, Graph, narzędzi offline i montowania katalogów dokumentów.
  Host/Origin są sprawdzane dokładnie, bez CORS. Wszystkie endpointy danych wymagają capability;
  wyjątek /office/pair przyjmuje tylko metadane i wymaga dokładnego Origin.
- Kontener user10001, rootfs read-only, cap_drop ALL, no-new-privileges, limit RAM256MiB,
  CPU0.5, procesów64, loopback127.0.0.1:3100, socket w prywatnym tmpfs, sekrety read-only.
  Administrator hosta i użytkownik kontrolujący Docker są zaufani. Legacy serwer upstream,
  .NET i proxy Node nie są kopiowane do obrazu. Office.js pochodzi z CDN Microsoftu.

## Audyt i odbiór v2

Zmiana zgód jest świadomą korektą pierwszej wersji audytu: zamiast wymagać zgody na
każdy zapis, granicą jest jawnie udostępniona sesja dokumentu oraz zakres zlecenia.
Pozostają izolacja dokumentów, uwierzytelnione odczyty, retencja, preview/hash,
idempotencja i blokada nieznanego wyniku. Nie jest to usunięcie uwierzytelniania.

- 25 testów Python (22 stan/HTTP + 3 launcher): PASS.
- 31 testów JavaScript: PASS; mock Office.js oraz rzeczywista logika XML/hasha/operacji.
- `scripts/check_secure_container.py`: PASS na rzeczywistym TLS i STDIO, osiem narzędzi,
  dwustronne parowanie, odkrycie akapitu, podgląd, idempotencja, cancellation i revoke.
- Claude Code `mcp get word_ai_secure`: User scope, Connected.
- Globalny skill: quick_validate PASS; oba junctiony do jednego źródła.

### Potwierdzony test Word Office.js, 2026-09-16

Dodatek v2 dodano z Shared Folder do zwykłego dokumentu bez osadzonych kontrolek.
Jedno kliknięcie Udostępnij dokument AI wystarczyło do parowania i kolejnych zapisów.
Przeszły: replace_paragraph, replace_text w mieszanych runach, replace_text w komórce,
insert_after, set_style Heading1 oraz kolejna zamiana na końcowym obrazie. Niezależny
snapshot potwierdził wszystkie zmiany. XML zapisanego DOCX potwierdził sześć dokładnych
akapitów, zachowane pogrubienie zmienionego słowa, tabelę 1×1, Heading1, niezmieniony
akapit KEEP LAST i zero kontrolek. Plik testowy: `word-ai-validation/document-session-v2-clean.docx` (poza repozytorium).

Test wykazał, że Word może przenieść uniqueLocalId na akapit wstawiony po istniejącym.
Po zmianach struktury trzeba ponowić snapshot i rozpoznać cel; hash nadal broni przed
zapisem na nieaktualnym akapicie. Test pustych nagłówków/stopek wykazał GeneralException
przy Range.getOoxml; snapshot zachowuje dostępny tekst i jawnie zwraca coverage_errors,
complete=false, hash=null oraz read_only. Brak przypisów zwracał ItemNotFound i także
jest raportowany jawnie. Nie sprawdzono edycji niepustych nagłówków/stopek/przypisów.

Po restarcie Codexa osiem narzędzi v2 oraz globalny skill zostały odkryte w tej samej
rozmowie. Bezpośrednie word_session_list i snapshot/status zwróciły właściwy dokument
oraz succeeded dla sześciu akapitów. Kliknięcie Cofnij dostęp usunęło sesję; bezpośrednia
lista MCP jest pusta. Świeży klient Claude Code wcześniej potwierdził User scope/Connected;
nie uruchamiano dodatkowego modelowego zadania Claude.

25 testów Python, 31 JavaScript, walidacja skilla i parytetu konfiguracji, parser
3 skryptów PowerShell oraz 7 JSON: PASS. Semantic KB: 691 dokumentów, 0 błędów,
374 zastane ostrzeżenia; indeks 9637 krawędzi i 0 brakujących odnośników. Zaktualizowano
rejestr rozwiązania, źródło, pakiet wpływu i dwa indeksy, zachowując proposed.

Nie potwierdzono współedycji, synchronizacji SharePoint ani awarii hosta podczas zapisu.
Nie wykonano pełnego skanu CVE systemu bazowego. Nie deklaruje się bezpieczeństwa
każdego możliwego dokumentu ani zgody organizacyjnej na jego przetwarzanie.

## Historyczne testy v1 — nie są instrukcją obsługi v2

Poniższe zapisy dokumentują wcześniejszy profil z zaznaczanymi kontrolkami i osobną
zgodą na zapis. Bieżący kontrakt opisano powyżej.

### Próba na żywym Wordzie, 2026-09-16

Po restarcie sesji Codexa Rancher był wyłączony, więc serwer MCP nie został
załadowany przez klienta. Po uruchomieniu Ranchera potwierdzono działający kontener,
zaufany certyfikat Windows, HTTPS oraz sześć narzędzi przez rzeczywisty STDIO.
Narzędzia nie zostały automatycznie dodane do już rozpoczętej rozmowy. Początkowo
wymagano ręcznego uruchomienia Ranchera; później zastąpiono to skryptem startowym
opisanym niżej.

Przygotowano osobny, syntetyczny dokument. Parametr `--document` oficjalnego
`office-addin-dev-settings sideload` tylko kopiuje podany plik; nie osadza w nim
panelu. Do testu użyto więc kopii oficjalnego `WordDocumentWithTaskPane.docx`
z osadzonym identyfikatorem tego dodatku i jednym testowym zakresem.
Po pełnym restarcie Word rozpoznał rejestrację dodatku.

Test wykrył błąd w serwerze: żądanie `taskpane.html?_host_Info=...` dostawało 404.
Dodano regresję (najpierw FAIL), poprawiono obsługę query dla statycznych plików,
uzyskano 20/20 testów i przebudowano kontener. Panel Office.js następnie się
uruchomił i przycisk „Pokaż oznaczone zakresy” wyświetlił kontrolkę
`WORD-AI:live-test` z otwartego dokumentu. To potwierdza ładowanie panelu i lokalny
odczyt listy kontrolek przez Office.js.

Po zgłoszeniu 401 użytkownik wyraźnie zlecił wykonanie parowania przez Computer Use.
Token hosta i kontenera był zgodny. Po wpisaniu go bezpośrednio do pola i wybraniu
testowej kontrolki parowanie się powiodło; nie ustalono, jaką wartość wprowadzono
we wcześniejszej nieudanej próbie.

Żywy test wykrył również zmienność surowego eksportu OOXML: niezmieniony tekst
dawał różne hashe i poprawnie działająca bramka stale-preview blokowała zapis.
Odcisk obejmuje teraz kanoniczną strukturę wybranego akapitu (wraz z tekstem,
właściwościami i atrybutami) oraz eksportowane zależności stylów, fontów i motywu.
Pomija opakowanie transportowe, deklaracje prefiksów i techniczne identyfikatory
rewizji Worda. Testy potwierdzają stabilność reprezentacji oraz zmianę odcisku
po zmianie tekstu, bezpośredniego formatowania lub definicji stylu.

Po tej poprawce dwa żywe odczyty i podgląd zwróciły zgodny odcisk. Kliknięcie
„Odrzuć” dało `cancelled`; kolejny odczyt potwierdził `The test is ready.`.
Następnie po nowym podglądzie i kliknięciu zgody polecenie zakończyło się
`succeeded`, a niezależny odczyt zwrócił `The MCP live test passed.`.
Zapisany DOCX zweryfikowano dodatkowo przez odczyt XML: oba akapity poza zakresem
pozostały identyczne, zachowano jedną kontrolkę i jej tag `WORD-AI:live-test`.
Kliknięcia zgody/odmowy w tym syntetycznym teście wykonał agent na zlecenie użytkownika.

Tymczasowe logowanie Office wyłączono. Nie zmieniono ustawień Trust Center
i nie edytowano dokumentów biznesowych użytkownika. Zmiana obrazu zerwała istniejący
transport MCP w rozmowie (`Transport closed`); test zakończono świeżym połączeniem
STDIO z tym samym serwerem. Po zakończeniu przebudów klient Codex wymaga przeładowania
połączenia lub sesji; Word i Rancher mogą pozostać otwarte.

### Automatyczny start zależności MCP, 2026-09-16

Po kolejnym restarcie Codexa ponownie nie działał Docker, więc samo `docker exec`
nie mogło uruchomić MCP. Dodano `scripts/start_secure_mcp.py`, który uruchamia
Ranchera w tle na żądanie, czeka do 180 sekund na Docker i prywatny socket oraz
uruchamia wyłącznie istniejący kontener z oczekiwanymi etykietami Compose.
Skrypt nie buduje obrazu, nie tworzy kontenera i nie zmienia certyfikatów ani
sekretów. Nie dodano zadania harmonogramu ani globalnego autostartu Ranchera.

Rzeczywisty test przy początkowo wyłączonym Dockerze przeszedł w **86,1 s**:
`initialize`, `notifications/initialized`, sześć narzędzi i `word_session_list`.
STDOUT zawierał wyłącznie poprawne ramki JSON-RPC. Timeout startowy Codexa ustawiono
na 210 sekund. Po restarcie kontenera lista sesji jest pusta zgodnie z zasadą
nietrwałego stanu; dokument trzeba ponownie połączyć w panelu Worda.

Kolejna próba ujawniła zależność życia Ranchera od procesu MCP: log Codexa pokazał
udaną inicjalizację, po czym anulowanie klienta; zniknęło również połączenie Dockera.
Uruchamianie Ranchera zmieniono na `CREATE_BREAKAWAY_FROM_JOB`, z ukrytym oknem
i nadal na koncie użytkownika. Mechanizm korzysta z uprawnienia Windows do
oddzielenia procesu; nie zmienia limitów job object ani ustawień bezpieczeństwa.
Klient MCP i `docker exec` pozostają zwykłymi procesami zarządzanymi przez Codexa.

Po poprawce `initialize` z wyłączonego środowiska przeszedł w **95,9 s**, a Docker
pozostał dostępny po zakończeniu procesu startowego. Następnie odświeżono wyłącznie
`word_ai_secure` przez krótkie przełączenie `enabled` w konfiguracji (stan końcowy
`true`). Codex zarejestrował `status=ready`, udostępnił sześć narzędzi w tej samej
rozmowie, a bezpośrednie wywołanie `word_session_list` zwróciło `{"sessions": []}`.
Nie był potrzebny kolejny restart Codexa. Pełnego restartu aplikacji po tej ostatniej
poprawce nie testowano; potwierdzono zakończenie launchera i odnowienie klienta MCP.
Zachowanie flagi jest opisane w [dokumentacji Windows Job Objects](https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects).

## Odłączenie

`docker stop word-ai-secure` zatrzymuje mostek i usuwa stan z pamięci.
Usunięcie konfiguracji klienta: `codex mcp remove word_ai_secure` lub
`claude mcp remove word_ai_secure -s user`. Pełne odinstalowanie wymaga osobnego
zlecenia dotyczącego dokładnego kontenera, obrazu, wolumenu, manifestu i certyfikatu;
nie używać `docker system prune` ani usuwać innych certyfikatów.

## Źródła

- [Word Paragraph / uniqueLocalId](https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph?view=word-js-preview).
- [Word Document / getParagraphByUniqueLocalId](https://learn.microsoft.com/en-us/javascript/api/word/word.document?view=word-js-preview).
- [Codex MCP](https://learn.chatgpt.com/docs/extend/mcp?surface=cli).
- [Codex konfiguracja zgód](https://learn.chatgpt.com/docs/config-file/config-reference).
- [Claude Code MCP](https://code.claude.com/docs/en/mcp), [skills](https://code.claude.com/docs/en/skills).

- [Word InlinePicture: natywne wstawianie i właściwości](https://learn.microsoft.com/en-us/javascript/api/word/word.inlinepicture?view=word-js-preview).

### Stan klienta po rozszerzeniu o obrazy

Po restarcie Codexa lokalny launcher zgłosił WinError 5 (Access is denied) podczas
zimnego uruchamiania zależności; inicjalizacja MCP w tej rozmowie nie zakończyła się.
Po uruchomieniu Ranchera świeży launcher odpowiedział poprawnie i wystawił dziewięć
narzędzi z image_path; Claude Code potwierdził Connected. Przełączenie enabled w pliku
nie odświeżyło narzędzi aktywnego kroku tej rozmowy. Sam zapis konfiguracji nie jest
dowodem gotowości klienta Codex; potrzebne odświeżenie MCP przy działającym Rancherze.
Nie zmieniano uprawnień Windows ani reguł job object w celu obejścia odmowy.

Po dodaniu podmiany i usuwania obrazu profil wystawia 11 narzędzi. Nowy schemat klienta
jest widoczny dopiero po przebudowie/restarcie kontenera i odświeżeniu połączenia MCP.
Operacje strukturalne z 2026-09-23 dodają dwa narzędzia (13 łącznie):
`word_session_preview_move_section` i `word_session_preview_table_row`. Oba trzeba
dopisać do Codex `enabled_tools` oraz do allow Claude (`mcp__word_ai_secure__...`);
pozostałe operacje korzystają z istniejącego `word_session_preview`.

Walidacja rozszerzenia KB: 691 dokumentów, 0 błędów, 374 zastane ostrzeżenia.
Zmieniono istniejący rejestr rozwiązania, źródło i pakiet wpływu, bez zmian modeli BI.
