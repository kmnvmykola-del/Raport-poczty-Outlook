# Instrukcja raportowania liczby wiadomości w Outlooku i Microsoft 365

**Wersja dokumentacji:** 3.1  
**Data dokumentacji:** 14 października 2026 r.  
**Status:** materiał do oceny przez pracodawcę i dział IT przed wdrożeniem.

## Szybki przewodnik użytkownika

**Raport poczty Outlook** to program do liczenia wiadomości i prowadzenia dziennego raportu w Excelu. Obsługuje się go przez menu w oknie terminalu. Do codziennego używania nie trzeba wpisywać poleceń PowerShell ani znać programowania. Wdrożenie w pracy należy wcześniej uzgodnić z pracodawcą i IT zgodnie z punktem 3.

### Pierwsze uruchomienie i codzienna obsługa

1. Rozpakować cały ZIP do jednego folderu. Nie uruchamiać programu bezpośrednio z archiwum. Przy pierwszym uruchomieniu może być potrzebny Internet do instalacji wymaganych modułów; pozostawić okno otwarte.
2. Zamknąć poprzedni `Raport-poczty.xlsx` oraz pliki CSV, jeżeli są otwarte w Excelu.
3. Dwukrotnie kliknąć **`Uruchom-raport.cmd`**. W menu wpisać numer i nacisnąć Enter. Dla służbowej skrzynki Microsoft 365 wybrać **`1`**.
4. Dokończyć ewentualne logowanie Microsoft. Sprawdzić adres skrzynki wyświetlony w terminalu. W menu potwierdzenia konta wybrać **`1 — Użyj tej skrzynki`**; jeżeli adres jest niewłaściwy, wybrać **`2 — Zmień konto i zaloguj się ponownie`**. Przy niezgodności z kontem wymaganym w ustawieniach zatwierdzenie jest zablokowane.
5. Poczekać na zakończenie liczenia i komunikat **„Skoroszyt Excel zapisany:”** z lokalizacją wyniku. Sam opis uprawnień `Mail.Read` i `User.Read` nie oznacza zakończenia — program może jeszcze oczekiwać na logowanie.
6. Gdy pojawi się „Naciśnij dowolny klawisz, aby zamknąć okno”, nacisnąć klawisz. Ten komunikat pojawia się również po błędzie lub zakończeniu bez pomiaru; poprawny zapis potwierdza komunikat z kroku 5. W razie błędu przeczytać jego treść i punkt 10 przed zamknięciem okna.
7. W folderze programu otworzyć **`raporty` → `Raport-poczty.xlsx` → arkusz `Dziennik`**. Sprawdzić dzisiejszą datę, konto i liczniki. Następnego dnia powtórzyć kroki 2–7; program sam kontynuuje tabelę.

Duża skrzynka może wymagać kilku minut lub dłużej. Nie zamykać terminalu podczas liczenia lub zapisu. Zachowane logowanie może przyspieszyć kolejne uruchomienia, ale Microsoft lub zasady organizacji mogą ponownie wymagać uwierzytelnienia.

### Co robi każdy numer menu

Poniższa tabela opisuje główne menu po uruchomieniu `Uruchom-raport.cmd`.

| Wybór | Co robi program | Co należy zrobić |
| --- | --- | --- |
| **`1` — Microsoft 365** | Łączy się z podstawową skrzynką służbową lub szkolną przez serwer Microsoft. Outlook nie musi być uruchomiony. | Zalogować się na konto służbowe, sprawdzić adres i zatwierdzić skrzynkę. Wynik trafi do `raporty`. |
| **`2` — Outlook.com / Hotmail / Live** | Otwiera dodatkowe menu logowania do osobistej skrzynki Microsoft. | W dodatkowym menu wybrać `1` albo sam Enter, aby użyć okna logowania. Następnie zalogować się i potwierdzić adres. Wynik trafi do `raporty`. |
| **`3` — Outlook klasyczny** | Korzysta z klasycznego Outlooka dla Windows i pozwala wskazać jedną skrzynkę lub plik danych z jego profilu. | Otworzyć i zsynchronizować klasyczny Outlook, następnie wybrać numer magazynu z listy. Nowy Outlook nie obsługuje tego trybu; przy niedostępnym interfejsie program wyjaśni przyczynę i wróci do głównego menu. |
| **`4` — Test bez logowania** | Tworzy raport z danych przykładowych: **10 wszystkich, 3 nieprzeczytane, 7 przeczytanych wiadomości w 3 folderach**. Nie odczytuje rzeczywistej poczty. | Otworzyć wynik w **`raporty-testowe/Raport-poczty.xlsx`**. Do pierwszej instalacji ImportExcel może być potrzebny Internet. |
| **`0` — Zakończ** | Kończy uruchomienie bez nowego pomiaru. | Nacisnąć klawisz, gdy pojawi się polecenie zamknięcia okna. |

**Numery dotyczą aktualnie wyświetlanego menu.** Przykładowo `2` w menu głównym wybiera konto osobiste, `2` w menu metody logowania wybiera kod urządzenia, a `2` przy potwierdzaniu konta służy do jego zmiany. Wpisanie nieprawidłowego numeru powoduje ponowne wyświetlenie prośby o wybór.

Przykładowa kolejność dla skrzynki służbowej: **`Uruchom-raport.cmd` → `1` w menu głównym → logowanie, jeśli wymagane → sprawdzenie adresu → `1` przy potwierdzeniu konta → zapis raportu**. Dla konta osobistego: **`Uruchom-raport.cmd` → `2` w menu głównym → `1` lub Enter w menu logowania → logowanie → sprawdzenie adresu → `1` przy potwierdzeniu konta → zapis raportu**.

Wybór trybu obowiązuje tylko podczas bieżącego uruchomienia. Nie zmienia pliku ustawień. Instrukcja logowania kodem urządzenia i konfiguracji stałej znajduje się w punkcie 5.

### Do czego służy wynik w Excelu

| Element | Jak go rozumieć |
| --- | --- |
| Arkusz **`Dziennik`** | Jedna pozycja na dzień, konto i źródło danych. Liczniki pokazują stan całej skrzynki w chwili pomiaru, a nie liczbę wiadomości otrzymanych tego dnia. |
| Ponowny pomiar tego samego dnia | Aktualizuje ten sam wiersz najpóźniejszym wynikiem. Kolejny dzień dodaje nowy wiersz. Różne konta lub źródła danych mają osobne wiersze. |
| Status **`Pomiar`** | Dostępne są rzeczywiste zapisane wartości z tego dnia. Zero oznacza wynik pomiaru równy zero. |
| Status **`Brak pomiaru`** | Dzień pomiędzy zapisanymi datami, dla którego nie ma wyniku. Puste liczniki nie oznaczają zera. Program nie odtwarza dawnego stanu nieprzeczytanych wiadomości z dzisiejszej skrzynki. |
| Kolumna **`Uwagi`** | Miejsce na własną notatkę. Zapisać i zamknąć plik po edycji; program zachowa notatkę przy następnym pomiarze. |
| Arkusz **`Foldery`** | Szczegóły ostatniego pomiaru jednej skrzynki; po kolejnym pomiarze zestawienie zostanie odświeżone. |

Program zachowuje także dodatkowe arkusze użytkownika. Nie zmieniać nazw ani układu arkuszy i tabel programu. Do codziennej pracy wystarczy XLSX; szczegóły działania historii i pomocniczych CSV opisuje punkt 7. Sam program nie uruchamia się codziennie — pomiar wykonuje się ręcznie albo przez harmonogram opisany w punkcie 9.

### Jakie pliki są potrzebne

Pakiet zawiera wyłącznie sześć plików potrzebnych do uruchamiania, konfiguracji, tworzenia raportów i obsługi programu:

| Plik | Przeznaczenie |
| --- | --- |
| **`Uruchom-raport.cmd`** | Główny przycisk uruchamiania: wybór trybu, liczenie i zapis raportów. |
| `Pokaz-skrzynki.cmd` | Pomocnicze sprawdzenie konta lub listy magazynów. Wyświetla to samo menu trybów, ale **nie tworzy raportów**. Dla Microsoft 365 pokazuje jedną zalogowaną skrzynkę, dla Outlooka klasycznego — dostępne magazyny, a w teście — nazwę `DANE_TESTOWE`. |
| `Policz-wiadomosci.ps1` | Kod programu wykonywany przez pliki `.cmd`. |
| `ustawienia.json` | Stała konfiguracja, w szczególności konto i tryb dla uruchomień bez menu. |
| `Szablon-raportu.xlsx` | Wzór techniczny używany przy tworzeniu nowego raportu. Pozostawić go obok programu; nie wpisywać do niego pomiarów. |
| `Instrukcja.md` | Niniejszy przewodnik oraz szczegółowa dokumentacja dla pracodawcy i IT. |

Foldery `raporty` i `raporty-testowe` powstają dopiero podczas pracy. Archiwum dystrybucyjne nie zawiera prywatnych wyników, historii ani plików używanych do testowania kodu. **Zachowywać kopie zapasowe własnego `Raport-poczty.xlsx` i `historia.csv`** — są potrzebne do zachowania wyników i uwag. Sześciu plików pakietu nie usuwać przy porządkowaniu folderu.

## 1. Cel i wybór trybu

Program `Policz-wiadomosci.ps1` przygotowuje raport liczby wszystkich, nieprzeczytanych i przeczytanych wiadomości w jednej skrzynce. Głównym wynikiem jest plik Excel `Raport-poczty.xlsx` z dziennikiem pomiarów. Przy kolejnym uruchomieniu program uzupełnia istniejący skoroszyt. Instalacja aplikacji Excel nie jest wymagana do liczenia ani zapisu pliku XLSX.

| Tryb | Zastosowanie | Źródło danych |
| --- | --- | --- |
| `Microsoft365` — konto służbowe, domyślnie | Służbowa skrzynka Microsoft 365 / Exchange Online; dowolna wersja Outlooka lub Outlook w przeglądarce | Serwer przez Microsoft Graph; Outlook nie musi być zainstalowany |
| `Microsoft365` — konto osobiste | Skrzynka Outlook.com, Hotmail lub Live przypisana do osobistego konta Microsoft | Serwer przez Microsoft Graph; w menu należy wybrać konto osobiste |
| `OutlookKlasyczny` | Skrzynka lub plik danych skonfigurowany w klasycznym Outlooku dla Windows | Dane dostępne przez Outlook na komputerze |
| `Demonstracja` | Sprawdzenie uruchamiania programu i tworzenia raportów bez dostępu do poczty | Dane przykładowe zapisane w kodzie |

**Dla służbowej skrzynki Microsoft 365 wybrać w menu opcję 1. Do sprawdzenia programu bez logowania wybrać opcję 4.** Wersja Outlooka nie wpływa na liczenie przez serwer. Tryb serwerowy wymaga Internetu i dostępu do Microsoft Graph. Obsługa konta osobistego dotyczy skrzynki Outlook.com, Hotmail lub Live; nie obejmuje Gmaila tylko dlatego, że został dodany do Outlooka. Tryb serwerowy nie obsługuje serwera Exchange utrzymywanego wyłącznie w siedzibie organizacji ani plików PST. Takie dane można liczyć w trybie klasycznym, jeżeli są dostępne w klasycznym Outlooku.

Pliki `Uruchom-raport.cmd` i `Pokaz-skrzynki.cmd` wyświetlają menu w terminalu. Wybrany tryb obowiązuje tylko podczas bieżącego uruchomienia. Program nie zapisuje tego wyboru w `ustawienia.json`. Bez przełącznika `-Menu` skrypt korzysta z ustawień zapisanych w pliku; dotyczy to również harmonogramu.

## 2. Zakres i ograniczenia

Program uwzględnia widoczne zwykłe foldery i podfoldery, między innymi skrzynkę odbiorczą, elementy wysłane, wersje robocze, spam i elementy usunięte. Wiadomości w folderze usuniętych są nadal liczone. Pomija foldery wyszukiwania, foldery ukryte oraz systemowe foldery synchronizacji i diagnostyki.

Liczniki wiadomości obejmują klasę `IPM.Note` i klasy zaczynające się od `IPM.Note.`. Zaproszenia na spotkania, kontakty, zadania oraz raporty dostarczenia i odczytu nie wchodzą do tej liczby. Pomocnicze liczniki wszystkich elementów mogą obejmować także inne rodzaje elementów.

W trybie serwerowym klasa elementu jest odczytywana z właściwości MAPI o identyfikatorze `0x001A`. Program porównuje liczbową wartość identyfikatora i rozpoznaje równoważne zapisy, np. `String 0x001A` i `String 0x1a`. Nie przypisuje domyślnej klasy, jeżeli serwer nie zwrócił jej wartości. [Microsoft: właściwość klasy wiadomości](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property), [Microsoft: format identyfikatorów właściwości rozszerzonych](https://learn.microsoft.com/en-us/graph/api/resources/extended-properties-overview?view=graph-rest-1.0).

W trybie serwerowym raport dotyczy **podstawowej skrzynki zalogowanego użytkownika**. Nie obejmuje dodatkowych skrzynek współdzielonych, osobnego archiwum online ani plików PST. Zwykły folder „Archiwum” w podstawowej skrzynce jest uwzględniany. Microsoft odróżnia go od archiwum online. [Microsoft: foldery pocztowe](https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0).

W trybie klasycznym wybierany jest jeden magazyn z listy Outlooka. Inne magazyny nie są dodawane. Niepełna synchronizacja lub ograniczony okres poczty offline mogą zmniejszyć wynik. Tryb serwerowy Microsoft 365 usuwa zależność od zakresu poczty pobranej lokalnie.

Foldery są odczytywane kolejno, a czas raportu oznacza zakończenie liczenia. Nie jest to jednoczesna migawka skrzynki. Podczas pomiaru należy ograniczyć przenoszenie, usuwanie i oznaczanie wiadomości.

## 3. Wymagania i uzgodnienie z IT

Wymagane są Windows, Windows PowerShell 5.1 i prawo zapisu w zatwierdzonej lokalizacji pakietu. Program uruchamia się jako zwykły użytkownik, bez uprawnień administratora.

Tryb serwerowy wymaga dodatkowo:

- służbowego konta z podstawową skrzynką Exchange Online albo osobistego konta Microsoft ze skrzynką Outlook.com, Hotmail lub Live;
- dostępu do usług logowania Microsoft, Microsoft Graph i — przy instalacji — PowerShell Gallery;
- .NET Framework 4.7.2 lub nowszego oraz sprawnego PowerShellGet i dostawcy NuGet;
- modułu `Microsoft.Graph.Authentication` w wersji `2.41.1`.

Jeżeli modułu Microsoft Graph brakuje, program próbuje zainstalować ten moduł z PowerShell Gallery dla bieżącego użytkownika. Jeśli brakuje dostawcy NuGet, również próbuje go zainstalować dla bieżącego użytkownika. Instalację i wymagane składniki musi dopuszczać polityka organizacji. [Microsoft: wymagania i instalacja SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0).

Zapis XLSX we wszystkich trybach wymaga dodatkowo modułu **ImportExcel 7.8.10**. Jest to moduł społecznościowy, odrębny od oficjalnego SDK Microsoft. Program instaluje go z PowerShell Gallery dla bieżącego użytkownika, jeżeli go brakuje. IT powinien zatwierdzić także tę zależność. Moduł tworzy pliki bez uruchamiania aplikacji Excel. [ImportExcel: pakiet i wersja](https://www.powershellgallery.com/packages/ImportExcel/7.8.10), [kod projektu](https://github.com/dfinke/ImportExcel).

Tryb demonstracyjny nie wymaga Outlooka, konta Microsoft, interfejsu COM ani modułu Microsoft Graph. Do zapisu XLSX potrzebuje ImportExcel; po wcześniejszej instalacji tego modułu działa również bez Internetu. Sprawdza działanie programu na danych przykładowych; nie potwierdza dostępu do rzeczywistej skrzynki.

Pracodawca i IT powinni zaakceptować zakres raportu, kod, moduł, uprawnienia API, logowanie, lokalizację zapisu, okres przechowywania i ewentualny harmonogram. Wymóg zgody administratora należy przekazać IT.

Pliki `.cmd` używają `-ExecutionPolicy Bypass` tylko dla uruchamianego procesu. IT może wymagać podpisanego skryptu lub innej metody uruchamiania. [Microsoft: zasady wykonywania](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1).

## 4. Dane i uprawnienia

Tryb serwerowy korzysta z oficjalnego modułu Microsoft i delegowanego uprawnienia `Mail.Read`. Uprawnienie pozwala aplikacji czytać pocztę zalogowanego użytkownika, w tym treść. **Jego zakres jest szerszy niż dane pobierane przez ten program.** Program odczytuje identyfikatory, stan przeczytania, klasę elementu, właściwości folderów i liczniki. Nie pobiera tematów, treści ani załączników do raportu. Nie żąda uprawnień do wysyłania ani zmieniania poczty. [Microsoft: zakres Mail.Read](https://learn.microsoft.com/en-us/graph/permissions-reference#mailread).

Program żąda także delegowanego uprawnienia `User.Read`, aby rozpoznać własne konto użytkownika po zalogowaniu. Odczytuje z profilu identyfikator, nazwę wyświetlaną, adres poczty i identyfikator logowania. Dane te służą do pokazania właściwej skrzynki i sprawdzenia zgodności z ustawieniami. [Microsoft: zakres User.Read](https://learn.microsoft.com/en-us/graph/permissions-reference#userread).

Program nie przechowuje hasła, nie oznacza wiadomości jako przeczytane, nie wysyła ani nie usuwa wiadomości. Microsoft Graph wymaga komunikacji HTTPS z usługami Microsoft. Raporty są zapisywane obok programu i nie są przez niego przesyłane pocztą.

Moduł Microsoft przechowuje kontekst logowania i tokeny dla bieżącego użytkownika Windows, aby następne uruchomienia mogły korzystać z wcześniejszego logowania. Po pomiarze skrypt nie wylogowuje użytkownika. Wybranie polecenia zmiany konta w menu powoduje jawne wylogowanie przez `Disconnect-MgGraph` i ponowne logowanie. Tokeny nie trafiają do ustawień ani CSV. [Microsoft: kontekst uwierzytelniania](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0).

Po aktualizacji programu należy raz zalogować się ręcznie i zatwierdzić również `User.Read` przed sprawdzeniem harmonogramu. Zmiana konta w menu może wymagać ponownego logowania do konta służbowego przed następnym pomiarem zaplanowanym.

Raport zawiera nazwę konta, ścieżki folderów, czas i liczby. Dane te mogą mieć charakter służbowy; ich ochronę i odbiorców określa pracodawca.

## 5. Menu i logowanie do właściwej skrzynki

1. Rozpakować pakiet do stałej lokalizacji zatwierdzonej przez IT, np. `C:\Raport-poczty-Outlook`. Zachować razem sześć plików, w tym `Szablon-raportu.xlsx`. Szablon jest potrzebny przy tworzeniu nowego skoroszytu.
2. Uruchomić `Pokaz-skrzynki.cmd` albo `Uruchom-raport.cmd`. W terminalu wpisać numer wybranej opcji i nacisnąć Enter.

| Numer | Wybór | Kiedy użyć |
| --- | --- | --- |
| `1` | Konto służbowe Microsoft 365 | Skrzynka pracodawcy w Exchange Online |
| `2` | Konto osobiste Outlook.com / Hotmail / Live | Własna skrzynka do próby logowania i pomiaru |
| `3` | Outlook klasyczny | Pomiar danych dostępnych przez klasyczny Outlook dla Windows |
| `4` | Test demonstracyjny bez logowania | Weryfikacja raportów na danych przykładowych; po instalacji ImportExcel także bez Internetu |
| `0` | Zakończenie | Zamknięcie programu bez wykonywania pomiaru |

3. Opcja 1 używa metody logowania zapisanej w ustawieniach. Po wyborze opcji 2 program pyta o metodę logowania do konta osobistego: `1` lub Enter oznacza okno logowania Microsoft, `2` oznacza kod urządzenia, a `0` kończy uruchomienie. Zalecana jest metoda `1`, aby nie wpisywać ręcznie kodu urządzenia. Następnie zaczekać na ewentualną instalację modułu i dokończyć logowanie w oknie Microsoft lub przeglądarce. Gdy okno nie jest widoczne, sprawdzić pasek zadań. Numer menu wybiera rodzaj konta i metodę; konkretny adres wybiera się podczas logowania Microsoft. Wykonać wymagane uwierzytelnienie. Microsoft może nadal wymagać własnego kodu MFA lub innej weryfikacji konta.
4. Sprawdzić adres konta pokazany przez program. Wybrać `1`, aby użyć tej skrzynki, `2`, aby wylogować bieżące konto i zalogować inne, albo `0`, aby zakończyć uruchomienie. Gdy służbowa `NazwaSkrzynki` jest wpisana w ustawieniach, niewłaściwego konta nie można zatwierdzić. `Pokaz-skrzynki.cmd` w trybie serwerowym pokazuje jedną zalogowaną skrzynkę, nie wszystkie konta dodane do Outlooka.
5. Do pomiaru uruchomić `Uruchom-raport.cmd` i ponownie wybrać właściwą opcję. Zamknąć `Raport-poczty.xlsx` i wcześniejsze CSV w Excelu. Poczekać na „Skoroszyt Excel zapisany:” i zakończenie programu. Duża skrzynka może wymagać kilku minut lub dłużej.
6. Otworzyć `raporty/Raport-poczty.xlsx` dla prawdziwej skrzynki albo `raporty-testowe/Raport-poczty.xlsx` dla demonstracji. Na arkuszu `Dziennik` sprawdzić datę, skrzynkę, czas i źródło pomiaru. Wykonać kontrolę z punktu 8.

Komunikat zaczynający się od „Mail.Read umożliwia odczyt poczty” opisuje wymagane uprawnienia; nie oznacza błędu. Przy logowaniu kodem urządzenia terminal pokazuje komunikat Microsoft zawierający kod. Program czeka na zakończenie logowania w przeglądarce, dlatego okno terminalu należy pozostawić otwarte. Jeżeli sesja jest już zapisana, nowy kod może nie być potrzebny i program przejdzie do potwierdzenia konta.

Jeżeli wybrano metodę `2` w menu logowania do konta osobistego, otworzyć `https://microsoft.com/devicelogin` i wpisać wyłącznie kod z bieżącego komunikatu Microsoft, np. z fragmentu „enter the code …”. Nie wpisywać numeru menu, adresu e-mail ani kodu z wcześniejszego uruchomienia. Program nie tworzy własnego kodu: otrzymuje go z usługi Microsoft. Kod ma ograniczony czas ważności. Pozostawić terminal otwarty aż do końca logowania; osobiste konto Microsoft może wymagać ponownego zalogowania. Jeżeli kod jest odrzucany, zamknąć stare uruchomienie i rozpocząć nowe albo wybrać metodę `1` — okno logowania Microsoft. [Microsoft: logowanie kodem urządzenia](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code).

**Komunikat Microsoft „Nie można użyć konta osobistego” oznacza wybór logowania dla kont służbowych lub szkolnych.** Dla adresu Outlook.com, Hotmail lub Live należy wrócić do programu i wybrać opcję 2. Opcja 1 korzysta z logowania organizacyjnego, a opcja 2 z logowania do osobistych kont Microsoft. Zmiana rodzaju logowania nie zastępuje uprawnień ani zgody administratora. W organizacji korzystającej z zatwierdzonego identyfikatora dzierżawy obowiązuje konfiguracja wskazana przez IT.

Wybór konta osobistego w menu nie zmienia służbowych ustawień zapisanych w pliku. Na potrzeby tego uruchomienia program wybiera logowanie `consumers` i metodę wskazaną w dodatkowym menu (`Przegladarka` albo `KodUrzadzenia`). Nie używa służbowej wartości `NazwaSkrzynki` jako ograniczenia dla konta prywatnego ani służbowego identyfikatora aplikacji. Korzysta z domyślnej aplikacji modułu Microsoft Graph; do tego wyboru nie trzeba tworzyć własnej rejestracji aplikacji. Test demonstracyjny można wykonać niezależnie od logowania Microsoft. [Microsoft: metody logowania i parametr TenantId](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph?view=graph-powershell-1.0).

### Ustawienia stałe i konto służbowe

Otworzyć `ustawienia.json` w Notatniku. Dla zaplanowanego pomiaru służbowego pozostawić `TrybPolaczenia` jako `Microsoft365`. Zaleca się wpisać dokładny służbowy identyfikator logowania w `NazwaSkrzynki`, aby zapobiec pomiarowi niewłaściwego konta. Program sprawdza zgodność tego ustawienia z kontem logowania i zatrzymuje się przy rozbieżności.

Przykład ustawień:

```json
{
  "TrybPolaczenia": "Microsoft365",
  "NazwaSkrzynki": "uzytkownik@firma.pl",
  "MetodaLogowania": "Przegladarka",
  "IdDzierzawy": "organizations",
  "IdAplikacji": ""
}
```

Przykładowy identyfikator zastąpić właściwym. Pusta `NazwaSkrzynki` oznacza pomiar konta użytego do logowania, a nie konta wybranego w Outlooku. Pozostałe ustawienia pozostawić bez zmian, chyba że IT dostarczy inną konfigurację. Puste `IdAplikacji` korzysta z domyślnej aplikacji SDK. Własna rejestracja aplikacji nie jest standardowym krokiem użytkownika; IT może wskazać zatwierdzony identyfikator aplikacji typu public client. Menu nie zapisuje zmian w tej konfiguracji.

Jeżeli domyślne logowanie nie działa, po uzgodnieniu z IT można ustawić `MetodaLogowania` na `KodUrzadzenia`. Program wyświetli `https://microsoft.com/devicelogin` i kod. Należy użyć kodu wygenerowanego przez własne uruchomienie i właściwego konta. Organizacja może blokować tę metodę; zmiana ustawienia nie zastępuje zgody administratora.

## 6. Tryb klasycznego Outlooka

Otworzyć klasyczny Outlook z menu „Plik”, zakończyć logowanie i zaczekać na synchronizację. Uruchomić `Pokaz-skrzynki.cmd` i wybrać opcję 3, aby sprawdzić dostępne magazyny. Następnie uruchomić `Uruchom-raport.cmd` i wybrać opcję 3. W menu program wyświetla ponumerowaną listę magazynów i pozwala wskazać jeden przez wpisanie numeru. Wybór dotyczy tylko bieżącego uruchomienia; przy uruchamianiu bez menu nadal obowiązuje `NazwaSkrzynki` z ustawień.

Jeżeli interfejs klasycznego Outlooka nie jest zarejestrowany, opcja 3 wyświetli wyjaśnienie i ponownie pokaże menu. Dla nowego Outlooka wybrać opcję 1 lub 2, zgodnie z rodzajem konta. Sam wybór numeru 3 nie instaluje klasycznego Outlooka. Uruchomienie bez menu z ustawieniem `OutlookKlasyczny` nadal kończy się błędem, gdy wymagany interfejs jest niedostępny.

Dla uruchamiania bez menu ustawić `TrybPolaczenia` na `OutlookKlasyczny` i wpisać dokładną nazwę magazynu w `NazwaSkrzynki`. Nazwa magazynu nie musi być adresem e-mail. Bez menu pusta wartość pozwala na automatyczny wybór tylko przy jednym magazynie.

Należy używać konta Windows z właściwym profilem Outlooka. Skrypt może otworzyć Outlook i nie zamyka aplikacji po pomiarze. Nowy Outlook nie udostępnia interfejsu COM używanego w tym trybie.

Dla całej historii dostępnej lokalnie IT powinien sprawdzić okres poczty offline i synchronizację. [Microsoft: ograniczona synchronizacja Outlooka](https://learn.microsoft.com/en-us/troubleshoot/outlook/user-interface/only-subset-items-synchronized).

## 7. Raporty i Excel

Folder `raporty` powstaje automatycznie obok skryptu dla pomiarów rzeczywistej poczty. Test demonstracyjny zapisuje te same cztery rodzaje plików w osobnym folderze `raporty-testowe`. **Raporty przykładowe nie zastępują raportów służbowych.**

| Plik | Zawartość |
| --- | --- |
| `Raport-poczty.xlsx` | Główny plik Excel: dziennik dzienny i ostatnie zestawienie folderów |
| `podsumowanie.csv` | Ostatni pomyślny pomiar skrzynki |
| `foldery.csv` | Ostatni pomyślny pomiar według folderów |
| `historia.csv` | Kolejne pomyślne pomiary dopisywane jako wiersze |

### Kontynuacja tabeli Excel i uzupełnianie dat

Arkusz `Dziennik` zawiera prawdziwą tabelę Excel z filtrowaniem i zamrożonym nagłówkiem. Wiersz odpowiada jednej dacie, skrzynce i źródłu danych. Rozdzielenie źródeł zapobiega mieszaniu wyniku serwerowego z lokalnymi danymi klasycznego Outlooka. Dzień wynika z daty zapisanej w czasie pomiaru.

- Pierwsze uruchomienie tworzy skoroszyt z szablonu i odczytuje wcześniejsze pomiary z `historia.csv`.
- Dla każdego dnia program wybiera najpóźniejszy zapisany pomiar. Powtórny pomiar tego samego dnia aktualizuje istniejący wiersz; nie sumuje liczników i nie tworzy drugiego wiersza dla tej samej skrzynki i źródła.
- Program porównuje zawartość istniejącego XLSX, historię CSV i bieżący pomiar. Zachowuje wcześniejsze dni także wtedy, gdy historia CSV została utracona lub skrócona.
- Między pierwszym zapisanym dniem a datą bieżącego pomiaru dopisuje brakujące daty ze statusem **`Brak pomiaru`**. Liczniki i czas pozostają puste. Rzeczywisty pomiar z wynikiem zero jest zapisywany jako zero ze statusem `Pomiar`.
- Jeżeli później w historii CSV pojawi się pomiar z brakującego dnia, program uzupełni ten wiersz rzeczywistymi wartościami.
- W kolumnie `Uwagi` można wpisać wyjaśnienie lub notatkę. Program zachowuje ją przy kolejnych uruchomieniach, także po uzupełnieniu brakującego pomiaru. Zachowuje również dodatkowe arkusze użytkownika.

Liczniki pokazują stan skrzynki w chwili pomiaru. Nie oznaczają liczby wiadomości otrzymanych danego dnia. **Bez wcześniej zapisanego pomiaru nie można odtworzyć historycznej liczby nieprzeczytanych wiadomości z dzisiejszej skrzynki.** Program nie wpisuje takich wartości na podstawie przypuszczeń. Aby uzyskiwać rzeczywiste wyniki za każdy dzień, należy wykonywać pomiar codziennie lub ustawić harmonogram opisany w punkcie 9.

Przykład: zapisano pomiary 10 i 12 października, a następny wykonano 14 października. Tabela pokaże daty 10–14 października. Dni 11 i 13 otrzymają status `Brak pomiaru` i puste liczniki. Jeżeli 14 października program zostanie uruchomiony ponownie, zmieni wynik i czas tylko w dzisiejszym wierszu.

Arkusz `Foldery` zawiera szczegóły **ostatniego pomiaru jednej skrzynki**. Jest odświeżany przy zapisie; nie jest historią folderów za wszystkie dni. Podczas odczytu pliku nie zmieniać nazw tabel, nagłówków ani układu dwóch arkuszy programu. Własne obliczenia można umieścić w dodatkowym arkuszu, a notatki w `Uwagi`.

Przed kolejnym pomiarem zamknąć skoroszyt w Excelu. Program nie steruje otwartym oknem Excel. Przy blokadzie pliku zatrzymuje zapis i zachowuje poprzednie raporty. Po pomiarze otworzyć plik ponownie. Zachowywać kopie zapasowe `Raport-poczty.xlsx` i `historia.csv`: utracony XLSX można odtworzyć z historii, ale usunięcie obu plików uniemożliwia odzyskanie wcześniejszych pomiarów i uwag. Same pomiary CSV nie zawierają ręcznych uwag z XLSX.

### Pomocnicze pliki CSV

| Kolumna | Znaczenie |
| --- | --- |
| `DataPomiaru` | Czas zakończenia pomiaru ze strefą czasową komputera |
| `Skrzynka` | Konto Microsoft, magazyn Outlooka albo `DANE_TESTOWE` |
| `WszystkieWiadomosci` | Wszystkie wiadomości e-mail w przyjętym zakresie |
| `NieprzeczytaneWiadomosci` | Nieprzeczytane wiadomości e-mail |
| `PrzeczytaneWiadomosci` | Wszystkie wiadomości minus nieprzeczytane |
| `WszystkieElementyOutlook` | Pomocniczy licznik wszystkich elementów folderów |
| `NieprzeczytaneElementyOutlook` | Pomocniczy licznik nieprzeczytanych elementów |
| `LiczbaFolderow` | Liczba uwzględnionych folderów |
| `ZrodloDanych` | `Microsoft365Serwer`, `OutlookKlasycznyDaneLokalne` albo `TrybDemonstracyjny` |
| `SciezkaFolderu` | Ścieżka folderu w zestawieniu folderów |

Do raportowania poczty służą trzy kolumny kończące się na `Wiadomosci`. Nazwy pomocniczych kolumn z wyrazem `Outlook` są jednakowe we wszystkich trybach; w trybie serwerowym pochodzą z serwera, a w demonstracji z danych przykładowych.

CSV używa UTF-8 z BOM i separatora `;`. Jeśli Excel umieści cały wiersz w jednej komórce, wybrać „Dane” → „Z tekstu/CSV”, UTF-8 i średnik. Przed kolejnym pomiarem zamknąć CSV. Otwarta tabela nie odświeża się automatycznie.

## 8. Sprawdzenie programu i pierwszego pomiaru

### Test bez konta, po przygotowaniu modułu ImportExcel

1. Uruchomić `Uruchom-raport.cmd`, wpisać `4` i nacisnąć Enter. Program powinien wskazać tryb demonstracyjny i zapisać pliki w `raporty-testowe`.
2. Otworzyć `raporty-testowe/Raport-poczty.xlsx`. W arkuszu `Dziennik` dla bieżącej daty i skrzynki `DANE_TESTOWE` powinny być: `Wszystkie` = `10`, `Nieprzeczytane` = `3`, `Przeczytane` = `7`, `Foldery` = `3`, `Status` = `Pomiar` i źródło `TrybDemonstracyjny`.
3. Sprawdzić arkusz `Foldery`. Trzy foldery zawierają odpowiednio 6 wiadomości, w tym 2 nieprzeczytane; 3 wiadomości, wszystkie przeczytane; oraz 1 wiadomość nieprzeczytaną. Suma wynosi 10, 3 i 7.
4. Dopisać próbny tekst w `Uwagi`, zapisać i zamknąć XLSX. Powtórzyć test tego samego dnia. Dziennik powinien nadal mieć jeden wiersz dla tego dnia, z nowszym czasem i zachowaną uwagą; `historia.csv` powinna otrzymać kolejny pomiar. Dane w folderze `raporty` nie powinny ulec zmianie.

`Pokaz-skrzynki.cmd` z opcją 4 wyświetla informację o przykładowej skrzynce `DANE_TESTOWE`; samo wyświetlenie tej listy nie tworzy raportów. Dane demonstracyjne są sztuczne i nie opisują żadnej rzeczywistej skrzynki.

Test bez menu można uruchomić bezpośrednio w terminalu. W poleceniu podać rzeczywistą lokalizację skryptu:

```text
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "C:\Raport-poczty-Outlook\Policz-wiadomosci.ps1" -TestDemonstracyjny
```

### Sprawdzenie na własnej lub służbowej poczcie

Do próby na własnym Outlook.com wybrać opcję 2, a następnie 1 w menu metody logowania. Do zatwierdzonej próby na służbowym Microsoft 365 wybrać opcję 1 w menu głównym. Test na danych przykładowych nie potwierdza logowania Microsoft Graph, uprawnień ani zgodności wyniku z pocztą. W ramach kontroli jakości wykonano dwa kolejne pomiary rzeczywistej osobistej skrzynki Outlook.com. Wynik wymaga osobnego potwierdzenia dla skrzynki służbowej pracodawcy.

1. Sprawdzić konto, świeże `DataPomiaru` i właściwe źródło danych.
2. Wybrać dwa lub trzy niewielkie foldery ze zwykłymi wiadomościami. Policzyć wszystkie i nieprzeczytane w Outlooku lub Outlooku w przeglądarce; wyłączyć filtrowanie i grupowanie w konwersacje. Porównać z wierszami `foldery.csv`.
3. Dla folderu z 10 zwykłymi wiadomościami, w tym 3 nieprzeczytanymi, oczekiwać 10, 3 i 7 w kolumnach wiadomości. Nie trzeba tworzyć ani zmieniać poczty do kontroli.
4. Sprawdzić obecność wymaganych folderów i podfolderów oraz uzgodniony zakres archiwum.
5. Zamknąć CSV i powtórzyć pomiar. Sprawdzić nowy czas i dodatkowy wiersz w historii. Przy niezmienionej skrzynce liczby powinny pozostać zgodne.

Sumy wiadomości z `foldery.csv` powinny odpowiadać podsumowaniu. Różnice z Outlookiem wymagają wyjaśnienia: filtrowanie, rodzaje elementów lub zmiany podczas liczenia. Wyznaczona osoba powinna potwierdzić zakres i wyniki przed użyciem raportu w pracy.

### Wynik kontroli jakości pakietu

Kontrolę wykonano 14 października 2026 r. w Windows PowerShell 5.1, z modułami Microsoft.Graph.Authentication 2.41.1 i ImportExcel 7.8.10. Mechanizmy dotychczasowego raportu objęto 69 sprawdzeniami automatycznymi. Nowy zapis XLSX sprawdzono dodatkowo w 23 przypadkach, a utworzone pliki odczytano niezależnym czytnikiem i sprawdzono ich wygląd.

| Zakres | Wynik i sposób sprawdzenia |
| --- | --- |
| Menu, rodzaje kont i metody logowania | Sprawdzono wybór, anulowanie, błędne numery, zmianę konta i odrzucenie niezgodnego konta; interfejs logowania podstawiono w testach automatycznych |
| Podliczanie przez Microsoft Graph | Sprawdzono hierarchię i kolejne strony folderów, puste foldery, pomijanie folderów technicznych, różne zapisy identyfikatora klasy oraz odróżnianie poczty od zaproszeń i raportów; użyto przygotowanych odpowiedzi API |
| Błędy i spójność danych | Sprawdzono brak wymaganych pól, sprzeczne duplikaty, zapętlenie stron, niedozwolone adresy, zmianę liczników, odmowę dostępu oraz ponawianie HTTP 429 |
| CSV i historia | Sprawdzono wartości, sumy, UTF-8 z BOM, polskie znaki, średniki i cudzysłowy, dopisywanie historii, ochronę tekstu w Excelu oraz przywracanie poprzednich plików przy blokadzie zapisu |
| Dzienny XLSX | Sprawdzono wybór ostatniego pomiaru dnia, aktualizację bez duplikatu, brakujące daty i puste liczniki, uzupełnianie z odnalezionej historii, prawdziwe zera, kilka skrzynek, rok przestępny, zachowanie uwag i dodatkowych arkuszy, odtwarzanie utraconego XLSX oraz wycofanie zmian przy blokadzie czwartego pliku |
| Rzeczywiste uruchamianie | Uruchomiono oba pliki CMD oraz osobny proces PowerShell w ścieżce ze spacjami i polskimi znakami; sprawdzono demonstrację i blokadę równoległego pomiaru |
| Rzeczywista osobista skrzynka Outlook.com | Wykonano dwa kolejne odczyty z serwera i dodatkowo dwa pełne uruchomienia przez zwykły plik `Uruchom-raport.cmd`. Sprawdzono zapis CSV, zgodność sum i dopisywanie historii; nie pobierano tematów, treści ani załączników |
| Rzeczywisty XLSX po aktualizacji | W dwóch kolejnych uruchomieniach zwykłego CMD utworzono i zaktualizowano skoroszyt rzeczywistej skrzynki. Potwierdzono jeden wiersz dzienny, brak duplikatu, zgodność z najnowszym pomiarem CSV, odczyt dat i liczb przez niezależny czytnik oraz zestawienie ośmiu folderów |
| Klasyczny Outlook | Sprawdzono algorytm na obiektach testowych i poprawną reakcję na rzeczywisty brak zarejestrowanego interfejsu COM. Integracja z działającym klasycznym Outlookiem wymaga osobnego sprawdzenia na komputerze, na którym jest dostępna |

Interfejs automatyzacji aplikacji Excel jest niedostępny na komputerze testowym; sprawdzono format i odczyt CSV oraz XLSX, w tym typy dat, liczniki, tabele i wygląd obu arkuszy, bez otwierania w samej aplikacji Excel. Nie wykonano pomiaru rzeczywistej skrzynki pracodawcy ani testu Harmonogramu zadań z jego zasadami logowania. Te elementy należy potwierdzić podczas uzgodnionego pilotażu w organizacji. Wyniki kontroli pakietu nie zastępują akceptacji pracodawcy i IT.

## 9. Codzienny harmonogram

Po akceptacji i pilotażu utworzyć zadanie w Harmonogramie zadań Windows:

1. Wybrać „Utwórz zadanie”, zwykłe konto użytkownika i **„Uruchom tylko wtedy, gdy użytkownik jest zalogowany”**. Nie zaznaczać najwyższych uprawnień.
2. Dodać codzienny wyzwalacz.
3. Jako program podać `powershell.exe`, jako argumenty polecenie poniżej, a w „Rozpocznij w” rzeczywistą lokalizację pakietu bez cudzysłowów.
4. Wybrać „Nie uruchamiaj nowego wystąpienia” dla już działającego zadania. Wykonać próbny start i sprawdzić czas raportu.

```text
-NoProfile -STA -ExecutionPolicy Bypass -File "C:\Raport-poczty-Outlook\Policz-wiadomosci.ps1"
```

W harmonogramie uruchamiać bezpośrednio skrypt, **bez przełącznika `-Menu`**. Nie wskazywać pliku `Uruchom-raport.cmd`, ponieważ czeka on na wybór w terminalu. Harmonogram używa `ustawienia.json`; wcześniejszy wybór konta osobistego lub demonstracji w menu nie zmienia zaplanowanego pomiaru służbowego.

Komputer musi być włączony, niewstrzymany, a użytkownik zalogowany. Microsoft 365 wymaga Internetu i wcześniejszego udanego logowania pod tym samym kontem Windows. Zapisana sesja może umożliwić następne pomiary bez udziału użytkownika, ale polityka logowania, MFA lub wygaśnięcie sesji mogą wymagać ponownego logowania. **Pakiet nie gwarantuje całkowicie bezobsługowego działania.** Takie wdrożenie wymaga osobnej konfiguracji IT. Należy kontrolować aktualność raportów.

## 10. Obsługa błędów

| Problem | Działanie |
| --- | --- |
| `80040154` / `Class not registered` | Interfejs COM trybu klasycznego jest niedostępny. Dla Exchange Online użyć `Microsoft365`. Komunikat sam nie rozstrzyga, czy Outlooka brakuje, czy wymaga naprawy |
| Błąd instalacji modułu | Sprawdzić Internet; przekazać IT komunikat dotyczący Gallery, NuGet, PowerShellGet lub .NET |
| Brakuje szablonu XLSX | Rozpakować także `Szablon-raportu.xlsx` do tego samego folderu co skrypt. Nie uruchamiać samego pliku PS1 skopiowanego z niepełnego pakietu |
| Skoroszyt ma inny układ lub duplikat dnia | Zachować kopię pliku i wyjaśnić zmianę tabeli z osobą odpowiedzialną za wsparcie. Program nie nadpisuje nierozpoznanej tabeli ani nie usuwa samodzielnie niejednoznacznych wierszy |
| Dzień ma status `Brak pomiaru` | Dla tej daty nie znaleziono wcześniejszego pomiaru w XLSX ani historii CSV. Dostarczenie zachowanej historii pozwoli uzupełnić dzień przy kolejnym pomiarze; same dzisiejsze dane poczty nie odtwarzają historycznego stanu |
| Microsoft nie pozwala użyć konta osobistego | Dla skrzynki Outlook.com, Hotmail lub Live wybrać opcję 2 w menu. Opcja 1 jest przeznaczona dla kont służbowych i szkolnych |
| Nieudane logowanie konta osobistego | Wybrać opcję 2 w menu głównym, a następnie 1 w menu logowania. Dokończyć logowanie we właściwym oknie Microsoft. Zachować dokładny komunikat Microsoft dla osoby odpowiedzialnej za wsparcie. Do sprawdzenia CSV można niezależnie użyć opcji 4 |
| Kod urządzenia jest nieprawidłowy lub nieważny | Zamknąć stare uruchomienie i uruchomić program ponownie. Wybrać 2 w menu głównym, następnie 1 w menu logowania, aby skorzystać z okna Microsoft. Przy wyborze metody 2 wpisać aktualny kod z tego samego otwartego uruchomienia, bez spacji; numer menu nie jest kodem |
| Po komunikacie o uprawnieniach program czeka, a kod nie jest widoczny | Sprawdzić wybraną metodę: okno Microsoft może być ukryte na pasku zadań, a zapisana sesja nie wymaga nowego kodu. Przy metodzie kodu urządzenia użyć aktualnego skryptu z pakietu; starsza wersja mogła ukrywać komunikat Microsoft. Pozostawić terminal otwarty |
| Zgoda administratora / dostęp zabroniony | IT ocenia aplikację, `Mail.Read`, `User.Read` i politykę logowania |
| Starszy komunikat „Brakuje identyfikatora, stanu przeczytania lub jednoznacznej klasy wiadomości” | Zaktualizować skrypt z bieżącego pakietu i ponowić pomiar. Starsza wersja mogła nie rozpoznawać skróconego zapisu identyfikatora klasy, np. `String 0x1a`. Nowa wersja rozróżnia brak identyfikatora, nieprawidłowy stan przeczytania i brak lub niejednoznaczność klasy |
| Microsoft Graph nie zwrócił klasy wiadomości | Pomiar nie zapisuje częściowego wyniku. Przekazać dokładny komunikat osobie odpowiedzialnej za wsparcie; nie zastępować nieznanej klasy zwykłym e-mailem |
| Niewłaściwe konto | Sprawdzić rodzaj konta w menu i adres pokazany po logowaniu; w menu potwierdzenia konta wybrać `2`, aby zmienić konto. Dla stałego pomiaru porównać także `NazwaSkrzynki` |
| Nie znaleziono magazynu klasycznego | Odczytać listę i poprawić nazwę magazynu |
| Błąd JSON | Sprawdzić nazwy ustawień, cudzysłowy, przecinki i nawiasy |
| Blokada zapisu | Zamknąć `Raport-poczty.xlsx` i CSV; sprawdzić prawo zapisu i brak równoległego pomiaru |
| Wynik zbyt mały | Sprawdzić zakres i archiwum; w trybie klasycznym także synchronizację |

Błąd odczytu przerywa pomiar, aby częściowy wynik nie został zapisany jako prawidłowy. Poprzedni raport może pozostać; zawsze sprawdzać `DataPomiaru`. Przy błędzie zapisu program próbuje przywrócić poprzednie pliki. Jeżeli pozostawi kopie `.bak`, zachować je i przekazać komunikat osobie odpowiedzialnej za wsparcie.
