#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$PlikUstawien = '',
    [switch]$PokazSkrzynki,
    [switch]$Menu,
    [switch]$TestDemonstracyjny
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# W PowerShell 5.1 katalog skryptu może być pusty podczas wyliczania parametrów.
# Ustal lokalizację dopiero po wejściu do skryptu, niezależnie od katalogu roboczego.
$SciezkaSkryptu = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($SciezkaSkryptu)) { $SciezkaSkryptu = $PSCommandPath }
if ([string]::IsNullOrWhiteSpace($SciezkaSkryptu)) {
    throw 'Nie można ustalić lokalizacji programu. Uruchom zapisany plik Policz-wiadomosci.ps1.'
}
$KatalogProgramu = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($SciezkaSkryptu))
if ([string]::IsNullOrWhiteSpace($PlikUstawien)) {
    $PlikUstawien = Join-Path $KatalogProgramu 'ustawienia.json'
}

function Zwolnij-referencjeCom {
    param($Wartosc)
    if ($null -ne $Wartosc -and [Runtime.InteropServices.Marshal]::IsComObject($Wartosc)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Wartosc)
    }
}

function Sprawdz-brakWlasciwosciMapi {
    param([Exception]$Wyjatek)
    while ($null -ne $Wyjatek) {
        if ($Wyjatek.HResult -eq -2147221233) { return $true } # MAPI_E_NOT_FOUND
        $Wyjatek = $Wyjatek.InnerException
    }
    return $false
}

function Pobierz-ustawienie {
    param($Ustawienia, [string]$Nazwa, $DomyslnaWartosc)
    $wlasciwosc = $Ustawienia.PSObject.Properties[$Nazwa]
    if ($null -eq $wlasciwosc) { return $DomyslnaWartosc }
    return $wlasciwosc.Value
}

function Wybierz-metodeLogowaniaOsobistego {
    while ($true) {
        Write-Host ''
        Write-Host 'LOGOWANIE DO KONTA OSOBISTEGO'
        Write-Host '1 - Okno logowania Microsoft (zalecane; bez ręcznego wpisywania kodu urządzenia)'
        Write-Host '2 - Kod urządzenia na stronie microsoft.com/devicelogin'
        Write-Host '0 - Zakończ'
        $wyborMetody = Read-Host 'Wpisz numer albo naciśnij Enter, aby wybrać 1'
        switch ($wyborMetody) {
            { $_ -in @('', '1') } { return 'Przegladarka' }
            '2' { return 'KodUrzadzenia' }
            '0' { return $null }
            default { Write-Host 'Nieprawidłowy wybór. Wpisz 0, 1 albo 2.' }
        }
    }
}

function Wybierz-trybUruchomienia {
    param($Ustawienia)
    # Wybór w menu dotyczy tylko tego uruchomienia; plik ustawień pozostaje bez zmian.
    $kopia = [pscustomobject][ordered]@{
        TrybPolaczenia = [string](Pobierz-ustawienie $Ustawienia 'TrybPolaczenia' 'Microsoft365')
        NazwaSkrzynki = [string](Pobierz-ustawienie $Ustawienia 'NazwaSkrzynki' '')
        MetodaLogowania = [string](Pobierz-ustawienie $Ustawienia 'MetodaLogowania' 'Przegladarka')
        IdDzierzawy = [string](Pobierz-ustawienie $Ustawienia 'IdDzierzawy' 'organizations')
        IdAplikacji = [string](Pobierz-ustawienie $Ustawienia 'IdAplikacji' '')
    }
    while ($true) {
        Write-Host ''
        Write-Host 'WYBÓR SKRZYNKI LUB TESTU'
        Write-Host '1 - Microsoft 365: konto służbowe lub szkolne'
        Write-Host '2 - Outlook.com / Hotmail / Live: konto osobiste'
        Write-Host '3 - Outlook klasyczny: skrzynki i pliki danych na komputerze'
        Write-Host '4 - Test bez logowania: dane przykładowe'
        Write-Host '0 - Zakończ'
        $wybor = Read-Host 'Wpisz numer'
        switch ($wybor) {
            '1' {
                $kopia.TrybPolaczenia = 'Microsoft365'
                if ($kopia.IdDzierzawy -in @('', 'common', 'consumers')) { $kopia.IdDzierzawy = 'organizations' }
                return $kopia
            }
            '2' {
                $metodaOsobista = Wybierz-metodeLogowaniaOsobistego
                if ($null -eq $metodaOsobista) { return $null }
                $kopia.TrybPolaczenia = 'Microsoft365'
                $kopia.IdDzierzawy = 'consumers'
                $kopia.MetodaLogowania = $metodaOsobista
                $kopia.IdAplikacji = ''
                $kopia.NazwaSkrzynki = ''
                Write-Host 'Wybrano konto osobiste. Wybrana metoda logowania dotyczy tylko tego uruchomienia; ustawienia firmowe nie zostaną zapisane ani zastąpione.'
                return $kopia
            }
            '3' {
                try {
                    Sprawdz-klasycznyOutlook
                    $kopia.TrybPolaczenia = 'OutlookKlasyczny'
                    return $kopia
                } catch {
                    Write-Host ''
                    Write-Host 'Opcja 3 wymaga klasycznego Outlooka dla Windows. Jego interfejs jest niedostępny na tym komputerze.' -ForegroundColor Yellow
                    Write-Host 'Nowy Outlook nie obsługuje opcji 3. Wybierz 1 dla służbowego Microsoft 365, 2 dla osobistego Outlook.com albo 4 do testu bez logowania.'
                }
            }
            '4' { $kopia.TrybPolaczenia = 'Demonstracja'; return $kopia }
            '0' { return $null }
            default { Write-Host 'Nieprawidłowy wybór. Wpisz 0, 1, 2, 3 albo 4.' }
        }
    }
}

function Pobierz-daneDemonstracyjne {
    # Stałe dane przykładowe, bez odczytywania skrzynki i bez połączenia z Internetem.
    foreach ($folder in @(
        @{ Nazwa = 'Skrzynka odbiorcza'; Wszystkie = 6; Nieprzeczytane = 2 },
        @{ Nazwa = 'Elementy wysłane'; Wszystkie = 3; Nieprzeczytane = 0 },
        @{ Nazwa = 'Wersje robocze'; Wszystkie = 1; Nieprzeczytane = 1 }
    )) {
        [pscustomobject][ordered]@{
            SciezkaFolderu = '\\DANE_TESTOWE\' + $folder.Nazwa
            WszystkieWiadomosci = [long]$folder.Wszystkie
            NieprzeczytaneWiadomosci = [long]$folder.Nieprzeczytane
            PrzeczytaneWiadomosci = [long]($folder.Wszystkie - $folder.Nieprzeczytane)
            WszystkieElementyOutlook = [long]$folder.Wszystkie
            NieprzeczytaneElementyOutlook = [long]$folder.Nieprzeczytane
        }
    }
}

function Wybierz-skrzynkeZListy {
    param([object[]]$Skrzynki)
    while ($true) {
        Write-Host 'Wybierz skrzynkę lub plik danych Outlooka:'
        for ($i = 0; $i -lt $Skrzynki.Count; $i++) {
            Write-Host ('{0} - {1}' -f ($i + 1), $Skrzynki[$i].Nazwa)
        }
        Write-Host '0 - Zakończ'
        $wybor = Read-Host 'Wpisz numer skrzynki'
        if ($wybor -eq '0') { return $null }
        [int]$numer = 0
        if ([int]::TryParse($wybor, [ref]$numer) -and $numer -ge 1 -and $numer -le $Skrzynki.Count) {
            return $Skrzynki[$numer - 1]
        }
        Write-Host 'Nieprawidłowy numer skrzynki.'
    }
}

function Przygotuj-modulMicrosoftGraph {
    $wersjaModulu = [version]'2.41.1'
    $dostepne = @(Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Where-Object { $_.Version -eq $wersjaModulu })
    if ($dostepne.Count -eq 0) {
        Write-Host 'Pierwsze uruchomienie: instaluję moduł Microsoft.Graph.Authentication 2.41.1 dla bieżącego użytkownika.'
        Write-Host 'Instalacja wymaga Internetu i dopuszczenia przez zasady organizacji.'
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $repozytorium = Get-PSRepository -Name PSGallery -ErrorAction Stop
            if ($repozytorium.SourceLocation.TrimEnd('/') -ne 'https://www.powershellgallery.com/api/v2') {
                throw 'Źródło PSGallery różni się od oficjalnego repozytorium Microsoft.'
            }
            $dostawcy = @(Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' })
            if ($dostawcy.Count -eq 0) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            }
            Install-Module -Name Microsoft.Graph.Authentication -RequiredVersion $wersjaModulu -Scope CurrentUser -Repository PSGallery -Force -ErrorAction Stop
        } catch {
            throw ('Nie udało się zainstalować modułu Microsoft Graph. Dział IT może wykonać: Install-Module Microsoft.Graph.Authentication -RequiredVersion 2.41.1 -Scope CurrentUser -Repository PSGallery -Force. Szczegóły: ' + $_.Exception.Message)
        }
    }
    try {
        Import-Module Microsoft.Graph.Authentication -RequiredVersion $wersjaModulu -ErrorAction Stop
    } catch {
        throw ('Nie można wczytać modułu Microsoft Graph. Wymagane są Windows PowerShell 5.1 z .NET Framework 4.7.2 lub nowszym albo PowerShell 7. Szczegóły: ' + $_.Exception.Message)
    }
}

function Polacz-zMicrosoft365 {
    param($Ustawienia, [switch]$WybierzKonto)
    Przygotuj-modulMicrosoftGraph
    $metoda = [string](Pobierz-ustawienie $Ustawienia 'MetodaLogowania' 'Przegladarka')
    $dzierzawa = [string](Pobierz-ustawienie $Ustawienia 'IdDzierzawy' 'organizations')
    $aplikacja = [string](Pobierz-ustawienie $Ustawienia 'IdAplikacji' '')
    if ($metoda -notin @('Przegladarka', 'KodUrzadzenia')) {
        throw 'MetodaLogowania musi mieć wartość Przegladarka albo KodUrzadzenia.'
    }
    if ([string]::IsNullOrWhiteSpace($dzierzawa)) { throw 'IdDzierzawy nie może być puste.' }
    $parametry = @{
        Scopes = @('Mail.Read', 'User.Read')
        TenantId = $dzierzawa
        ContextScope = 'CurrentUser'
        NoWelcome = $true
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($aplikacja)) { $parametry.ClientId = $aplikacja }
    if ($metoda -eq 'KodUrzadzenia') { $parametry.UseDeviceCode = $true }
    while ($true) {
        if ($dzierzawa -eq 'consumers') { Write-Host 'Logowanie do konta osobistego Outlook.com / Hotmail / Live.' }
        else { Write-Host 'Logowanie do konta służbowego lub szkolnego Microsoft 365.' }
        Write-Host 'Mail.Read umożliwia odczyt poczty; program pobiera właściwości potrzebne do liczenia. User.Read służy do sprawdzenia własnego profilu i adresu konta.'
        if ($metoda -eq 'KodUrzadzenia') {
            Write-Host 'Jeżeli wymagane jest logowanie, poniżej pojawi się kod Microsoft. Otwórz https://microsoft.com/devicelogin, wpisz ten kod i dokończ logowanie w przeglądarce. To okno pozostaw otwarte.'
            Write-Host 'Wpisz wyłącznie aktualny kod z komunikatu Microsoft, bez spacji. Numer 2 w menu nie jest kodem logowania. Kod ze starego uruchomienia może być nieważny.'
            Write-Host 'Jeżeli strona odrzuca kod, zamknij to uruchomienie i rozpocznij ponownie. W menu konta osobistego możesz zamiast kodu wybrać okno logowania Microsoft.'
        } else {
            Write-Host 'Dokończ logowanie w oknie Microsoft lub w przeglądarce. Jeżeli okno jest ukryte, sprawdź pasek zadań. To okno terminalu pozostaw otwarte.'
        }
        try {
            # SDK przekazuje kod urządzenia do strumienia wyników. Wyświetl go od razu,
            # bez dodawania komunikatu logowania do zwracanego adresu skrzynki.
            Connect-MgGraph @parametry | Out-Host
        } catch {
            throw ('Logowanie Microsoft nie powiodło się. Wybierz właściwy rodzaj konta: służbowe Microsoft 365 albo osobiste Outlook.com. Organizacja może wymagać zgody administratora na Mail.Read i User.Read. Szczegóły: ' + $_.Exception.Message)
        }
        # Osobiste tokeny mogą być nieprzezroczyste; SDK nie zawsze odczytuje z nich nazwę konta.
        # Tożsamość potwierdzamy odpowiedzią Microsoft Graph, a nie samym Get-MgContext.Account.
        try {
            $profil = Odwolaj-zadanieGraph 'https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName'
        } catch { throw ('Nie można potwierdzić adresu zalogowanego konta. Wymagane jest User.Read. Szczegóły: ' + $_.Exception.Message) }
        $adres = [string](Pobierz-wlasciwoscGraph $profil 'mail')
        $identyfikatorLogowania = [string](Pobierz-wlasciwoscGraph $profil 'userPrincipalName')
        $idProfilu = [string](Pobierz-wlasciwoscGraph $profil 'id')
        if ([string]::IsNullOrWhiteSpace($idProfilu) -or ([string]::IsNullOrWhiteSpace($adres) -and [string]::IsNullOrWhiteSpace($identyfikatorLogowania))) {
            throw 'Microsoft Graph nie zwrócił identyfikatora i adresu konta. Raport nie został zapisany.'
        }
        $konto = $adres
        if ([string]::IsNullOrWhiteSpace($konto)) { $konto = $identyfikatorLogowania }
        $spodziewanaNazwa = [string](Pobierz-ustawienie $Ustawienia 'NazwaSkrzynki' '')
        $zgodnaNazwa = [string]::IsNullOrWhiteSpace($spodziewanaNazwa) -or
            [string]::Equals($adres, $spodziewanaNazwa, [StringComparison]::OrdinalIgnoreCase) -or
            [string]::Equals($identyfikatorLogowania, $spodziewanaNazwa, [StringComparison]::OrdinalIgnoreCase)
        if ($WybierzKonto) {
            while ($true) {
                Write-Host ''
                Write-Host ('Zalogowana skrzynka: ' + $konto)
                if ($zgodnaNazwa) { Write-Host '1 - Użyj tej skrzynki' }
                else { Write-Host ("Ustawienia wskazują '$spodziewanaNazwa'. To inne konto; nie będzie liczone.") }
                Write-Host '2 - Zmień konto i zaloguj się ponownie'
                Write-Host '0 - Zakończ'
                $wyborKonta = Read-Host 'Wpisz numer'
                if ($wyborKonta -eq '0') { return $null }
                if ($wyborKonta -eq '1' -and $zgodnaNazwa) { break }
                if ($wyborKonta -eq '2') {
                    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
                    catch { throw ('Nie udało się wylogować poprzedniego konta. Szczegóły: ' + $_.Exception.Message) }
                    break
                }
                Write-Host 'Nieprawidłowy wybór.'
            }
            if ($wyborKonta -eq '2') { continue }
        } elseif (-not $zgodnaNazwa) {
            throw ("Zalogowana skrzynka '$konto' różni się od NazwaSkrzynki '$spodziewanaNazwa'. Uruchom program z menu i wybierz zmianę konta.")
        }
        # Zachowaj zaakceptowany identyfikator logowania także wtedy, gdy główny adres jest innym aliasem.
        if (-not [string]::IsNullOrWhiteSpace($spodziewanaNazwa)) { return $spodziewanaNazwa }
        return $konto
    }
}

function Sprawdz-klasycznyOutlook {
    if ($null -eq [Type]::GetTypeFromProgID('Outlook.Application')) {
        throw 'Interfejs klasycznego Outlooka nie jest zarejestrowany. Nowy Outlook i Outlook w przeglądarce nie obsługują tego trybu. Dla Microsoft 365 ustaw TrybPolaczenia = Microsoft365. Jeśli potrzebny jest tryb OutlookKlasyczny, dział IT powinien sprawdzić instalację klasycznego Outlooka.'
    }
}

function Pobierz-wlasciwoscGraph {
    param($Obiekt, [string]$Nazwa)
    if ($null -eq $Obiekt) { return $null }
    if ($Obiekt -is [Collections.IDictionary]) {
        if ($Obiekt.Contains($Nazwa)) { return ,$Obiekt[$Nazwa] }
        return $null
    }
    $wlasciwosc = $Obiekt.PSObject.Properties[$Nazwa]
    if ($null -ne $wlasciwosc) { return ,$wlasciwosc.Value }
    return $null
}

function Pobierz-informacjeBleduGraph {
    param([Management.Automation.ErrorRecord]$Blad)
    [int]$status = 0
    $kod = ''
    [int]$opoznienie = 0
    $wyjatek = $Blad.Exception
    while ($null -ne $wyjatek) {
        $odpowiedz = Pobierz-wlasciwoscGraph $wyjatek 'Response'
        foreach ($wartoscStatusu in @(
            (Pobierz-wlasciwoscGraph $odpowiedz 'StatusCode'),
            (Pobierz-wlasciwoscGraph $wyjatek 'ResponseStatusCode'),
            (Pobierz-wlasciwoscGraph $wyjatek 'StatusCode')
        )) {
            if ($null -ne $wartoscStatusu) {
                try { $status = [int]$wartoscStatusu } catch { }
                if ($status -ge 100 -and $status -le 599) { break }
            }
        }
        $naglowki = Pobierz-wlasciwoscGraph $odpowiedz 'Headers'
        if ($null -eq $naglowki) { $naglowki = Pobierz-wlasciwoscGraph $wyjatek 'ResponseHeaders' }
        $ponowienie = Pobierz-wlasciwoscGraph $naglowki 'RetryAfter'
        $sekundy = Pobierz-wlasciwoscGraph (Pobierz-wlasciwoscGraph $ponowienie 'Delta') 'TotalSeconds'
        $data = Pobierz-wlasciwoscGraph $ponowienie 'Date'
        if ($null -ne $sekundy) { $opoznienie = [int][Math]::Ceiling([double]$sekundy) }
        elseif ($null -ne $data) { $opoznienie = [int][Math]::Ceiling((([DateTimeOffset]$data) - [DateTimeOffset]::UtcNow).TotalSeconds) }
        if ($opoznienie -le 0 -and $null -ne $naglowki) {
            $wartoscNaglowka = Pobierz-wlasciwoscGraph $naglowki 'Retry-After'
            if ($null -eq $wartoscNaglowka -and $null -ne $naglowki.PSObject.Methods['TryGetValues']) {
                $wartosciNaglowka = $null
                try {
                    if ($naglowki.TryGetValues('Retry-After', [ref]$wartosciNaglowka)) {
                        $wartoscNaglowka = @($wartosciNaglowka)[0]
                    }
                } catch { }
            }
            [int]$liczbaSekund = 0
            if ([int]::TryParse([string](@($wartoscNaglowka)[0]), [ref]$liczbaSekund)) { $opoznienie = $liczbaSekund }
            elseif ($null -ne $wartoscNaglowka) {
                [DateTimeOffset]$dataNaglowka = [DateTimeOffset]::MinValue
                if ([DateTimeOffset]::TryParse([string](@($wartoscNaglowka)[0]), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dataNaglowka)) {
                    $opoznienie = [int][Math]::Ceiling(($dataNaglowka - [DateTimeOffset]::UtcNow).TotalSeconds)
                }
            }
        }
        $wyjatek = $wyjatek.InnerException
    }
    $tresc = [string]$Blad.Exception.Message
    $opisyDoSprawdzenia = @($tresc)
    if ($null -ne $Blad.ErrorDetails) {
        $szczegoly = [string]$Blad.ErrorDetails.Message
        $opisyDoSprawdzenia += $szczegoly
        $tresc += "`n" + $szczegoly
    }
    foreach ($opisTekstowy in $opisyDoSprawdzenia) {
        $trescJson = $opisTekstowy
        # SDK może dołączyć wiersz HTTP i nagłówki przed właściwym obiektem JSON błędu.
        if ($opisTekstowy -match '(?s)(\{\s*"error"\s*:.*\})\s*$') { $trescJson = $Matches[1] }
        try {
            $opis = $trescJson | ConvertFrom-Json -ErrorAction Stop
            $kod = [string](Pobierz-wlasciwoscGraph (Pobierz-wlasciwoscGraph $opis 'error') 'code')
        } catch { }
        if ($kod -ne '') { break }
    }
    if ($status -eq 0 -and $tresc -match '(?im)(?:\bStatus(?:Code)?\s*[:=]\s*|\bHTTP(?:/\d(?:\.\d)?)?\s+)([1-5]\d{2})\b') { $status = [int]$Matches[1] }
    if ($kod -eq '' -and $tresc -match '(?im)^\s*(?:Code|ErrorCode)\s*[:=]\s*([A-Za-z][A-Za-z0-9_.]+)\s*$') { $kod = $Matches[1] }
    if ($opoznienie -le 0 -and $tresc -match '(?im)^\s*Retry-After\s*:\s*(\d+)\s*$') { $opoznienie = [int]$Matches[1] }
    return [pscustomobject]@{ Status = $status; Kod = $kod; Opoznienie = [Math]::Max(0, $opoznienie) }
}

function Odwolaj-zadanieGraph {
    param([Parameter(Mandatory)][string]$Adres)
    [Uri]$adresSprawdzony = $null
    if (-not [Uri]::TryCreate($Adres, [UriKind]::Absolute, [ref]$adresSprawdzony) -or
        $adresSprawdzony.Scheme -ne 'https' -or
        -not [string]::Equals($adresSprawdzony.Host, 'graph.microsoft.com', [StringComparison]::OrdinalIgnoreCase) -or
        $adresSprawdzony.Port -ne 443 -or $adresSprawdzony.UserInfo -ne '' -or $adresSprawdzony.Fragment -ne '' -or
        ($adresSprawdzony.AbsolutePath -ne '/v1.0/me' -and -not $adresSprawdzony.AbsolutePath.StartsWith('/v1.0/me/', [StringComparison]::Ordinal))) {
        throw 'Odrzucono nieprawidłowy adres Microsoft Graph. Dozwolone są wyłącznie adresy HTTPS graph.microsoft.com/v1.0/me.'
    }
    # Najwyżej cztery ponowienia; łączny czas oczekiwania nie przekracza 120 sekund na żądanie.
    for ($proba = 0; $proba -le 4; $proba++) {
        try { return Invoke-MgGraphRequest -Method GET -Uri $Adres -OutputType PSObject -ErrorAction Stop }
        catch {
            $opisBledu = Pobierz-informacjeBleduGraph $_
            if ($opisBledu.Status -notin @(429, 500, 502, 503, 504) -or $proba -eq 4) { throw }
            $oczekiwanie = $opisBledu.Opoznienie
            if ($oczekiwanie -le 0) { $oczekiwanie = [int][Math]::Pow(2, $proba + 1) }
            # Nie ponawiaj wcześniej niż zalecił serwer, gdy wymagane oczekiwanie przekracza limit programu.
            if ($oczekiwanie -gt 30) { throw 'Serwer Microsoft 365 wymaga dłuższej przerwy. Poczekaj kilka minut i uruchom raport ponownie.' }
            Write-Warning ("Microsoft 365 chwilowo nie odpowiada (HTTP {0}). Ponowna próba za {1} s." -f $opisBledu.Status, $oczekiwanie)
            Start-Sleep -Seconds $oczekiwanie
        }
    }
}

function Pobierz-liczbeGraph {
    param($Obiekt, [string]$Nazwa, [string]$Sciezka)
    $wartosc = Pobierz-wlasciwoscGraph $Obiekt $Nazwa
    [long]$liczba = 0
    if ($null -eq $wartosc -or $wartosc -is [bool] -or
        -not [long]::TryParse([string]$wartosc, [Globalization.NumberStyles]::Integer, [Globalization.CultureInfo]::InvariantCulture, [ref]$liczba) -or $liczba -lt 0) {
        throw "Nieprawidłowa liczba '$Nazwa' w folderze '$Sciezka'. Raport nie został zapisany."
    }
    return $liczba
}

function Pobierz-stronyGraph {
    param([string]$Adres)
    $odwiedzoneStrony = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    while (-not [string]::IsNullOrWhiteSpace($Adres)) {
        if (-not $odwiedzoneStrony.Add($Adres)) { throw 'Microsoft Graph zwrócił zapętloną listę stron. Raport nie został zapisany.' }
        $strona = Odwolaj-zadanieGraph $Adres
        $elementy = Pobierz-wlasciwoscGraph $strona 'value'
        if ($null -eq $elementy -or $elementy -is [string] -or $elementy -isnot [Collections.IEnumerable]) {
            throw 'Microsoft Graph zwrócił niepełną listę elementów. Raport nie został zapisany.'
        }
        foreach ($element in $elementy) { $element }
        $nastepnaStrona = Pobierz-wlasciwoscGraph $strona '@odata.nextLink'
        if ($null -ne $nastepnaStrona -and ($nastepnaStrona -isnot [string] -or [string]::IsNullOrWhiteSpace($nastepnaStrona))) {
            throw 'Microsoft Graph zwrócił nieprawidłowy adres kolejnej strony. Raport nie został zapisany.'
        }
        # Używaj całego adresu kolejnej strony bez modyfikowania parametrów Microsoft Graph.
        $Adres = [string]$nastepnaStrona
    }
}

function Sprawdz-identyfikatorKlasyGraph {
    param($IdWlasciwosci)
    if ($IdWlasciwosci -isnot [string]) { return $false }
    # Graph może zwrócić np. String 0x1a zamiast żądanego String 0x001A.
    # Porównuj typ i liczbowy identyfikator właściwości MAPI, nie zapis z zerami.
    $opcje = [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant
    $dopasowanie = [regex]::Match($IdWlasciwosci.Trim(), '^String\s+0x([0-9a-f]{1,4})$', $opcje)
    if (-not $dopasowanie.Success) { return $false }
    return ([Convert]::ToInt32($dopasowanie.Groups[1].Value, 16) -eq 0x001A)
}

function Pobierz-daneFolderowMicrosoft365 {
    param([Parameter(Mandatory)][string]$Skrzynka)
    if ([string]::IsNullOrWhiteSpace($Skrzynka)) { throw 'Nie podano nazwy zalogowanej skrzynki Microsoft 365.' }
    $adresPodstawowy = 'https://graph.microsoft.com/v1.0/me/mailFolders'
    $parametryFolderow = '?$select=id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden&$top=100&includeHiddenFolders=true'
    $pominieteFoldery = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $odwiedzoneFoldery = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $wiersze = New-Object 'System.Collections.Generic.List[object]'
    # Pomijaj techniczne foldery i foldery wyszukiwania, aby nie dodawać diagnostyki ani duplikatów.
    foreach ($nazwaSystemowa in @('syncissues', 'conflicts', 'localfailures', 'serverfailures', 'searchfolders')) {
        try {
            $folderSystemowy = Odwolaj-zadanieGraph ($adresPodstawowy + '/' + $nazwaSystemowa + '?$select=id')
            $idSystemowe = [string](Pobierz-wlasciwoscGraph $folderSystemowy 'id')
            if ([string]::IsNullOrWhiteSpace($idSystemowe)) { throw 'Microsoft Graph nie zwrócił identyfikatora folderu systemowego.' }
            [void]$pominieteFoldery.Add($idSystemowe)
        } catch {
            $opisBledu = Pobierz-informacjeBleduGraph $_
            if ($opisBledu.Status -ne 404 -or $opisBledu.Kod -notin @('ErrorItemNotFound', 'ErrorFolderNotFound')) { throw }
        }
    }

    function Przejdz-folderMicrosoft365 {
        param($Folder, [string]$SciezkaRodzica)
        $idFolderu = [string](Pobierz-wlasciwoscGraph $Folder 'id')
        $nazwaFolderu = Pobierz-wlasciwoscGraph $Folder 'displayName'
        if ([string]::IsNullOrWhiteSpace($idFolderu) -or $nazwaFolderu -isnot [string]) { throw 'Microsoft Graph nie zwrócił identyfikatora lub nazwy folderu.' }
        if (-not $odwiedzoneFoldery.Add($idFolderu)) { return }
        if ($pominieteFoldery.Contains($idFolderu)) { return }
        $ukryty = Pobierz-wlasciwoscGraph $Folder 'isHidden'
        if ($ukryty -isnot [bool]) { throw "Nie można ustalić widoczności folderu '$nazwaFolderu'." }
        if ($ukryty -or (Pobierz-wlasciwoscGraph $Folder '@odata.type') -eq '#microsoft.graph.mailSearchFolder') { return }
        $sciezka = $SciezkaRodzica + '\' + $nazwaFolderu
        try {
            $wszystkieElementy = Pobierz-liczbeGraph $Folder 'totalItemCount' $sciezka
            $nieprzeczytaneElementy = Pobierz-liczbeGraph $Folder 'unreadItemCount' $sciezka
            $liczbaPodfolderow = Pobierz-liczbeGraph $Folder 'childFolderCount' $sciezka
            $adresFolderu = $adresPodstawowy + '/' + [Uri]::EscapeDataString($idFolderu)
            $rozwiniecie = [Uri]::EscapeDataString('singleValueExtendedProperties($filter=id eq ''String 0x001A'')')
            $adresWiadomosci = $adresFolderu + '/messages?$select=id,isRead&$top=500&$expand=' + $rozwiniecie
            [long]$wszystkieWiadomosci = 0
            [long]$nieprzeczytaneWiadomosci = 0
            $odwiedzoneWiadomosci = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
            foreach ($wiadomosc in (Pobierz-stronyGraph $adresWiadomosci)) {
                $idWiadomosci = [string](Pobierz-wlasciwoscGraph $wiadomosc 'id')
                $przeczytana = Pobierz-wlasciwoscGraph $wiadomosc 'isRead'
                $wlasciwosciRozszerzone = Pobierz-wlasciwoscGraph $wiadomosc 'singleValueExtendedProperties'
                $klasyWiadomosci = @($wlasciwosciRozszerzone | Where-Object {
                    Sprawdz-identyfikatorKlasyGraph (Pobierz-wlasciwoscGraph $_ 'id')
                })
                if ([string]::IsNullOrWhiteSpace($idWiadomosci)) { throw 'Brakuje identyfikatora wiadomości. Nie można przygotować pełnego raportu.' }
                if ($przeczytana -isnot [bool]) { throw 'Brakuje prawidłowego stanu przeczytania wiadomości. Nie można przygotować pełnego raportu.' }
                if ($klasyWiadomosci.Count -eq 0) { throw 'Microsoft Graph nie zwrócił klasy wiadomości (właściwość String 0x001A). Nie można przygotować pełnego raportu.' }
                if ($klasyWiadomosci.Count -ne 1) { throw 'Microsoft Graph zwrócił więcej niż jedną klasę wiadomości. Nie można przygotować pełnego raportu.' }
                $klasa = Pobierz-wlasciwoscGraph $klasyWiadomosci[0] 'value'
                if ($klasa -isnot [string] -or [string]::IsNullOrWhiteSpace($klasa)) { throw 'Microsoft Graph zwrócił pustą klasę wiadomości. Nie można przygotować pełnego raportu.' }
                $stan = $przeczytana.ToString() + '|' + $klasa.ToUpperInvariant()
                if ($odwiedzoneWiadomosci.ContainsKey($idWiadomosci)) {
                    if ($odwiedzoneWiadomosci[$idWiadomosci] -ne $stan) { throw 'Wiadomość zmieniła się podczas liczenia. Uruchom program ponownie.' }
                    continue
                }
                $odwiedzoneWiadomosci.Add($idWiadomosci, $stan)
                if ([string]::Equals($klasa, 'IPM.Note', [StringComparison]::OrdinalIgnoreCase) -or $klasa.StartsWith('IPM.Note.', [StringComparison]::OrdinalIgnoreCase)) {
                    $wszystkieWiadomosci++
                    if (-not $przeczytana) { $nieprzeczytaneWiadomosci++ }
                }
            }
            # Zweryfikuj liczniki folderu po pobraniu wszystkich stron wiadomości.
            $folderPoPomiarze = Odwolaj-zadanieGraph ($adresFolderu + '?$select=id,totalItemCount,unreadItemCount,childFolderCount')
            $wszystkieElementyPo = Pobierz-liczbeGraph $folderPoPomiarze 'totalItemCount' $sciezka
            $nieprzeczytaneElementyPo = Pobierz-liczbeGraph $folderPoPomiarze 'unreadItemCount' $sciezka
            $podfolderyPo = Pobierz-liczbeGraph $folderPoPomiarze 'childFolderCount' $sciezka
            if ($wszystkieElementyPo -ne $wszystkieElementy -or $nieprzeczytaneElementyPo -ne $nieprzeczytaneElementy -or $podfolderyPo -ne $liczbaPodfolderow -or
                $nieprzeczytaneElementy -gt $wszystkieElementy -or $wszystkieWiadomosci -gt $wszystkieElementy -or $nieprzeczytaneWiadomosci -gt $nieprzeczytaneElementy) {
                throw 'Zawartość folderu zmieniła się podczas liczenia lub liczniki są niespójne. Poczekaj na synchronizację i uruchom program ponownie.'
            }
            $wiersze.Add([pscustomobject][ordered]@{
                SciezkaFolderu = $sciezka
                WszystkieWiadomosci = $wszystkieWiadomosci
                NieprzeczytaneWiadomosci = $nieprzeczytaneWiadomosci
                PrzeczytaneWiadomosci = $wszystkieWiadomosci - $nieprzeczytaneWiadomosci
                WszystkieElementyOutlook = $wszystkieElementy
                NieprzeczytaneElementyOutlook = $nieprzeczytaneElementy
            })
            if ($liczbaPodfolderow -gt 0) {
                foreach ($podfolder in (Pobierz-stronyGraph ($adresFolderu + '/childFolders' + $parametryFolderow))) {
                    Przejdz-folderMicrosoft365 $podfolder $sciezka
                }
            }
        } catch { throw "Nie udało się w pełni policzyć folderu '$sciezka': $($_.Exception.Message)" }
    }

    foreach ($folder in (Pobierz-stronyGraph ($adresPodstawowy + $parametryFolderow))) {
        Przejdz-folderMicrosoft365 $folder ('\\' + $Skrzynka)
    }
    if ($wiersze.Count -eq 0) { throw 'Nie znaleziono żadnego dostępnego folderu poczty Microsoft 365. Raport nie został zapisany.' }
    return $wiersze.ToArray()
}

function Wybierz-skrzynke {
    param([object[]]$Skrzynki, [string]$NazwaSkrzynki)
    if ($Skrzynki.Count -eq 0) { throw 'W profilu Outlook nie ma dostępnych skrzynek ani plików danych.' }
    if ([string]::IsNullOrWhiteSpace($NazwaSkrzynki)) {
        if ($Skrzynki.Count -eq 1) { return $Skrzynki[0] }
        throw 'Znaleziono kilka skrzynek lub plików danych. Uruchom Pokaz-skrzynki.cmd i wpisz NazwaSkrzynki w ustawienia.json.'
    }
    $dopasowania = @($Skrzynki | Where-Object {
        [string]::Equals($_.Nazwa, $NazwaSkrzynki, [StringComparison]::OrdinalIgnoreCase)
    })
    if ($dopasowania.Count -ne 1) {
        throw "Nazwa '$NazwaSkrzynki' nie wskazuje dokładnie jednej skrzynki. Skopiuj dokładną, unikalną nazwę z Pokaz-skrzynki.cmd."
    }
    return $dopasowania[0]
}

function Pobierz-daneFolderow {
    param(
        [Parameter(Mandatory)]$FolderGlowny,
        [string[]]$PominieteIdFolderow = @()
    )
    $wiersze = New-Object 'System.Collections.Generic.List[object]'
    $odwiedzone = New-Object 'System.Collections.Generic.HashSet[string]'
    $filtrWiadomosci = '@SQL=("http://schemas.microsoft.com/mapi/proptag/0x001A001F" = ''IPM.Note'' OR "http://schemas.microsoft.com/mapi/proptag/0x001A001F" LIKE ''IPM.Note.%'')'

    function Przejdz-folder {
        param($Folder)
        $dostepDoWlasciwosci = $null
        $elementy = $null
        $wiadomosci = $null
        $nieprzeczytane = $null
        $podfoldery = $null
        $sciezka = '(nieznany folder)'
        try {
            $sciezka = [string]$Folder.FolderPath
            $id = [string]$Folder.EntryID
            if (-not $odwiedzone.Add($id)) { return }
            if ($PominieteIdFolderow -contains $id) { return }
            $dostepDoWlasciwosci = $Folder.PropertyAccessor
            # Foldery wyszukiwania zawierają odnośniki do wiadomości, więc pomijamy je w sumie.
            $typFolderu = [int]$dostepDoWlasciwosci.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003')
            if ($typFolderu -eq 2) { return }
            $ukryty = $false
            try {
                $ukryty = [bool]$dostepDoWlasciwosci.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x10F4000B')
            } catch {
                if (-not (Sprawdz-brakWlasciwosciMapi $_.Exception)) { throw }
                # Niektórzy dostawcy nie udostępniają PR_ATTR_HIDDEN dla zwykłych widocznych folderów.
            }
            if ($ukryty) { return }

            # 0 = olMailItem. Przechodzimy także inne foldery, aby znaleźć podfoldery poczty.
            if ([int]$Folder.DefaultItemType -eq 0) {
                $elementy = $Folder.Items
                $wiadomosci = $elementy.Restrict($filtrWiadomosci)
                $nieprzeczytane = $wiadomosci.Restrict('[UnRead] = true')
                $wszystkieWiadomosci = [long]$wiadomosci.Count
                $nieprzeczytaneWiadomosci = [long]$nieprzeczytane.Count
                $wszystkieElementy = [long]$elementy.Count
                $nieprzeczytaneElementy = [long]$Folder.UnReadItemCount
                if ($nieprzeczytaneWiadomosci -gt $wszystkieWiadomosci -or $nieprzeczytaneElementy -gt $wszystkieElementy -or $wszystkieWiadomosci -gt $wszystkieElementy) {
                    throw 'Zawartość zmieniła się podczas liczenia. Poczekaj na synchronizację i uruchom program ponownie.'
                }
                $wiersze.Add([pscustomobject][ordered]@{
                    SciezkaFolderu = $sciezka
                    WszystkieWiadomosci = $wszystkieWiadomosci
                    NieprzeczytaneWiadomosci = $nieprzeczytaneWiadomosci
                    PrzeczytaneWiadomosci = $wszystkieWiadomosci - $nieprzeczytaneWiadomosci
                    WszystkieElementyOutlook = $wszystkieElementy
                    NieprzeczytaneElementyOutlook = $nieprzeczytaneElementy
                })
            }
            $podfoldery = $Folder.Folders
            for ($i = 1; $i -le $podfoldery.Count; $i++) {
                $podfolder = $null
                try {
                    $podfolder = $podfoldery.Item($i)
                    Przejdz-folder $podfolder
                } finally { Zwolnij-referencjeCom $podfolder }
            }
        } catch {
            throw "Nie udało się w pełni policzyć folderu '$sciezka': $($_.Exception.Message)"
        } finally {
            Zwolnij-referencjeCom $nieprzeczytane
            Zwolnij-referencjeCom $wiadomosci
            Zwolnij-referencjeCom $elementy
            Zwolnij-referencjeCom $podfoldery
            Zwolnij-referencjeCom $dostepDoWlasciwosci
        }
    }

    Przejdz-folder $FolderGlowny
    if ($wiersze.Count -eq 0) { throw 'Nie znaleziono żadnego dostępnego folderu poczty. Raport nie został zapisany.' }
    return $wiersze.ToArray()
}

function Utworz-podsumowanie {
    param([object[]]$Wiersze, [string]$Skrzynka, [string]$DataPomiaru, [string]$ZrodloDanych = 'OutlookKlasycznyDaneLokalne')
    [long]$wszystkieWiadomosci = 0; [long]$nieprzeczytaneWiadomosci = 0
    [long]$wszystkieElementy = 0; [long]$nieprzeczytaneElementy = 0
    foreach ($wiersz in $Wiersze) {
        $wszystkieWiadomosci += $wiersz.WszystkieWiadomosci
        $nieprzeczytaneWiadomosci += $wiersz.NieprzeczytaneWiadomosci
        $wszystkieElementy += $wiersz.WszystkieElementyOutlook
        $nieprzeczytaneElementy += $wiersz.NieprzeczytaneElementyOutlook
    }
    return [pscustomobject][ordered]@{
        DataPomiaru = $DataPomiaru
        Skrzynka = $Skrzynka
        WszystkieWiadomosci = $wszystkieWiadomosci
        NieprzeczytaneWiadomosci = $nieprzeczytaneWiadomosci
        PrzeczytaneWiadomosci = $wszystkieWiadomosci - $nieprzeczytaneWiadomosci
        WszystkieElementyOutlook = $wszystkieElementy
        NieprzeczytaneElementyOutlook = $nieprzeczytaneElementy
        LiczbaFolderow = $Wiersze.Count
        ZrodloDanych = $ZrodloDanych
    }
}

function Zabezpiecz-tekstDlaExcela {
    param([object[]]$Wiersze)
    foreach ($wiersz in $Wiersze) {
        $bezpiecznyWiersz = [ordered]@{}
        foreach ($wlasciwosc in $wiersz.PSObject.Properties) {
            $wartosc = $wlasciwosc.Value
            if ($wartosc -is [string] -and $wartosc -match '^[\s]*[=+@-]') {
                $wartosc = "'" + $wartosc # Traktuj nazwy folderów i skrzynek jako tekst w Excelu.
            }
            $bezpiecznyWiersz[$wlasciwosc.Name] = $wartosc
        }
        [pscustomobject]$bezpiecznyWiersz
    }
}

function Zapisz-csvUtf8 {
    param([object[]]$Wiersze, [string]$Sciezka)
    $wierszeCsv = @(Zabezpiecz-tekstDlaExcela $Wiersze | ConvertTo-Csv -Delimiter ';' -NoTypeInformation)
    $kodowanie = New-Object System.Text.UTF8Encoding($true)
    [IO.File]::WriteAllLines($Sciezka, [string[]]$wierszeCsv, $kodowanie)
}

function Przygotuj-modulExcel {
    $wersja = [version]'7.8.10'
    if (@(Get-Module -ListAvailable -Name ImportExcel | Where-Object { $_.Version -eq $wersja }).Count -eq 0) {
        Write-Host 'Pierwsze uruchomienie Excel: instaluję moduł ImportExcel 7.8.10 dla bieżącego użytkownika.'
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $repozytorium = Get-PSRepository -Name PSGallery -ErrorAction Stop
            if ($repozytorium.SourceLocation.TrimEnd('/') -ne 'https://www.powershellgallery.com/api/v2') { throw 'Źródło PSGallery różni się od oficjalnego repozytorium.' }
            if (@(Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' }).Count -eq 0) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            }
            Install-Module -Name ImportExcel -RequiredVersion $wersja -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
        } catch { throw ('Nie udało się przygotować zapisu XLSX. Dział IT może wykonać: Install-Module ImportExcel -RequiredVersion 7.8.10 -Scope CurrentUser -Repository PSGallery -Force. Szczegóły: ' + $_.Exception.Message) }
    }
    try { Import-Module ImportExcel -RequiredVersion $wersja -ErrorAction Stop }
    catch { throw ('Nie można wczytać modułu ImportExcel. Szczegóły: ' + $_.Exception.Message) }
}

function Pobierz-kluczDnia {
    param($Wiersz)
    return (@($Wiersz.Data.ToString('yyyy-MM-dd'), $Wiersz.Skrzynka, $Wiersz.ZrodloDanych) | ConvertTo-Json -Compress)
}

function Konwertuj-pomiarDzienny {
    param($Pomiar)
    [DateTimeOffset]$czas = [DateTimeOffset]::MinValue
    if ([string]$Pomiar.DataPomiaru -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$' -or
        -not [DateTimeOffset]::TryParse([string]$Pomiar.DataPomiaru, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$czas) -or $czas.Year -lt 1900) {
        throw 'Nieprawidłowy czas pomiaru w historii. Istniejący skoroszyt nie został zmieniony.'
    }
    $skrzynka = [string]$Pomiar.Skrzynka
    $zrodlo = [string]$Pomiar.ZrodloDanych
    if ([string]::IsNullOrWhiteSpace($skrzynka) -or $zrodlo -notin @('Microsoft365Serwer', 'OutlookKlasycznyDaneLokalne', 'TrybDemonstracyjny')) { throw 'Brak prawidłowej skrzynki lub źródła danych w historii.' }
    $liczby = @{}
    foreach ($nazwa in @('WszystkieWiadomosci','NieprzeczytaneWiadomosci','PrzeczytaneWiadomosci','LiczbaFolderow')) {
        [long]$liczba = 0
        if (-not [long]::TryParse([string]$Pomiar.$nazwa, [Globalization.NumberStyles]::Integer, [Globalization.CultureInfo]::InvariantCulture, [ref]$liczba) -or $liczba -lt 0) { throw ('Nieprawidłowa wartość ' + $nazwa + ' w historii.') }
        $liczby[$nazwa] = $liczba
    }
    if ($liczby.NieprzeczytaneWiadomosci -gt $liczby.WszystkieWiadomosci -or $liczby.PrzeczytaneWiadomosci -ne $liczby.WszystkieWiadomosci - $liczby.NieprzeczytaneWiadomosci) { throw 'Niespójne liczniki w historii. Skoroszyt nie został zmieniony.' }
    return [pscustomobject]@{
        Data=$czas.Date; Skrzynka=$skrzynka; ZrodloDanych=$zrodlo; Czas=$czas
        WszystkieWiadomosci=$liczby.WszystkieWiadomosci; NieprzeczytaneWiadomosci=$liczby.NieprzeczytaneWiadomosci
        PrzeczytaneWiadomosci=$liczby.PrzeczytaneWiadomosci; LiczbaFolderow=$liczby.LiczbaFolderow
        Status='Pomiar'; Uwagi=''; FormulaUwagi=''; Style=$null; Wysokosc=23.0
    }
}

function Pobierz-dziennikZeSkoroszytu {
    param($Arkusz)
    $naglowki = @('Data','Skrzynka','Wszystkie','Nieprzeczytane','Przeczytane','Foldery','Czas pomiaru','Źródło danych','Status','Uwagi')
    if ($null -eq $Arkusz -or $null -eq $Arkusz.Tables['TabelaDziennika']) { throw 'Istniejący XLSX nie zawiera tabeli TabelaDziennika. Zachowano dotychczasowy plik.' }
    for ($kolumna=1; $kolumna -le $naglowki.Count; $kolumna++) {
        if ([string]$Arkusz.Cells[6,$kolumna].Value -cne $naglowki[$kolumna-1]) { throw 'Układ tabeli Dziennik różni się od oczekiwanego. Nie zmieniaj jej nagłówków.' }
    }
    $tabela = $Arkusz.Tables['TabelaDziennika']
    if ($tabela.Address.Start.Row -ne 6 -or $tabela.Address.Start.Column -ne 1 -or $tabela.Address.End.Column -ne 10 -or $tabela.ShowTotal) { throw 'Nieobsługiwany układ tabeli Dziennik. Istniejący plik został zachowany.' }
    for ($nr=7; $nr -le $tabela.Address.End.Row; $nr++) {
        $wartoscDaty = $Arkusz.Cells[$nr,1].Value
        if ($null -eq $wartoscDaty -and [string]::IsNullOrWhiteSpace([string]($Arkusz.Cells[$nr,2].Value))) { continue }
        if ($wartoscDaty -is [DateTime]) { $data=$wartoscDaty.Date }
        elseif ($wartoscDaty -is [double] -or $wartoscDaty -is [int] -or $wartoscDaty -is [long]) { $data=[DateTime]::FromOADate([double]$wartoscDaty).Date }
        else { throw 'Kolumna Data musi zawierać daty Excel, a nie tekst. Istniejący plik został zachowany.' }
        if ($data.Year -lt 1900) { throw 'Nieprawidłowa data w skoroszycie.' }
        $stan = [string]$Arkusz.Cells[$nr,9].Value
        if ($stan -eq 'Pomiar') {
            $rekord = Konwertuj-pomiarDzienny ([pscustomobject]@{
                DataPomiaru=$Arkusz.Cells[$nr,7].Value; Skrzynka=$Arkusz.Cells[$nr,2].Value; ZrodloDanych=$Arkusz.Cells[$nr,8].Value
                WszystkieWiadomosci=$Arkusz.Cells[$nr,3].Value; NieprzeczytaneWiadomosci=$Arkusz.Cells[$nr,4].Value
                PrzeczytaneWiadomosci=$Arkusz.Cells[$nr,5].Value; LiczbaFolderow=$Arkusz.Cells[$nr,6].Value
            })
            if ($rekord.Data -ne $data) { throw 'Data w tabeli różni się od daty pomiaru. Istniejący plik został zachowany.' }
        } elseif ($stan -eq 'Brak pomiaru') {
            if (@(3..7 | Where-Object { $null -ne $Arkusz.Cells[$nr,$_].Value -and [string]$Arkusz.Cells[$nr,$_].Value -ne '' }).Count -gt 0) { throw 'Dzień bez pomiaru powinien mieć puste liczniki i czas.' }
            $rekord=[pscustomobject]@{Data=$data;Skrzynka=[string]$Arkusz.Cells[$nr,2].Value;ZrodloDanych=[string]$Arkusz.Cells[$nr,8].Value;Czas=$null;WszystkieWiadomosci=$null;NieprzeczytaneWiadomosci=$null;PrzeczytaneWiadomosci=$null;LiczbaFolderow=$null;Status='Brak pomiaru';Uwagi='';FormulaUwagi='';Style=$null;Wysokosc=23.0}
            if ([string]::IsNullOrWhiteSpace($rekord.Skrzynka) -or $rekord.ZrodloDanych -notin @('Microsoft365Serwer','OutlookKlasycznyDaneLokalne','TrybDemonstracyjny')) { throw 'Nieprawidłowa skrzynka lub źródło dnia bez pomiaru.' }
        } else { throw 'Nieprawidłowy status w tabeli Dziennik.' }
        $rekord.Uwagi=$Arkusz.Cells[$nr,10].Value
        $rekord.FormulaUwagi=[string]$Arkusz.Cells[$nr,10].Formula
        $rekord.Style=@(1..10 | ForEach-Object { $Arkusz.Cells[$nr,$_].StyleID })
        $rekord.Wysokosc=$Arkusz.Row($nr).Height
        $rekord
    }
}

function Polacz-pomiaryDzienne {
    param([object[]]$Dotychczasowe, [object[]]$Historia, $AktualnyPomiar)
    $wedlugDnia=New-Object 'Collections.Generic.Dictionary[string,object]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($wiersz in $Dotychczasowe) {
        $klucz=Pobierz-kluczDnia $wiersz
        if ($wedlugDnia.ContainsKey($klucz)) { throw 'W istniejącym skoroszycie są dwa wiersze dla tej samej daty, skrzynki i źródła. Wyjaśnij duplikat przed uruchomieniem.' }
        $wedlugDnia.Add($klucz,$wiersz)
    }
    $biezacy=Konwertuj-pomiarDzienny $AktualnyPomiar
    $znaneNazwy=@($Dotychczasowe | ForEach-Object { $_.Skrzynka })+@($AktualnyPomiar.Skrzynka)
    foreach ($pomiar in @($Historia)+@($AktualnyPomiar)) {
        # CSV zabezpiecza niektóre nazwy apostrofem. Rozpoznaj zapis znanej skrzynki,
        # aby nie utworzyć drugiej tożsamości po wczytaniu własnego pliku CSV.
        $kopiaPomiaru=$pomiar | Select-Object *
        foreach ($znanaNazwa in $znaneNazwy) {
            if ($znanaNazwa -match '^[\s]*[=+@-]' -and [string]::Equals([string]$kopiaPomiaru.Skrzynka, ("'"+$znanaNazwa), [StringComparison]::OrdinalIgnoreCase)) { $kopiaPomiaru.Skrzynka=$znanaNazwa; break }
        }
        $nowy=Konwertuj-pomiarDzienny $kopiaPomiaru
        $klucz=Pobierz-kluczDnia $nowy
        if ($wedlugDnia.ContainsKey($klucz)) {
            $stary=$wedlugDnia[$klucz]
            $nowy.Uwagi=$stary.Uwagi; $nowy.FormulaUwagi=$stary.FormulaUwagi; $nowy.Style=$stary.Style; $nowy.Wysokosc=$stary.Wysokosc
            if ($stary.Status -eq 'Pomiar') {
                if ($nowy.Czas -lt $stary.Czas) { continue }
                if ($nowy.Czas -eq $stary.Czas -and ($nowy.WszystkieWiadomosci -ne $stary.WszystkieWiadomosci -or $nowy.NieprzeczytaneWiadomosci -ne $stary.NieprzeczytaneWiadomosci -or $nowy.LiczbaFolderow -ne $stary.LiczbaFolderow)) { throw 'Ten sam czas pomiaru ma różne wyniki. Wyjaśnij niespójność historii.' }
            }
        }
        $wedlugDnia[$klucz]=$nowy
    }
    $grupy=@($wedlugDnia.Values | Group-Object Skrzynka,ZrodloDanych)
    [long]$liczbaDocelowa=0
    foreach ($grupa in $grupy) {
        $uporzadkowane=@($grupa.Group | Sort-Object Data)
        $pierwszy=$uporzadkowane[0]; $koniec=$biezacy.Data
        if ($uporzadkowane[-1].Data -gt $koniec) { $koniec=$uporzadkowane[-1].Data }
        $liczbaDocelowa += [long](($koniec-$pierwszy.Data).TotalDays+1)
        if ($liczbaDocelowa -gt 1048570) { throw 'Zakres dat przekracza limit wierszy arkusza Excel.' }
        for ($dzien=$pierwszy.Data; $dzien -le $koniec; $dzien=$dzien.AddDays(1)) {
            $pusty=[pscustomobject]@{Data=$dzien;Skrzynka=$pierwszy.Skrzynka;ZrodloDanych=$pierwszy.ZrodloDanych;Czas=$null;WszystkieWiadomosci=$null;NieprzeczytaneWiadomosci=$null;PrzeczytaneWiadomosci=$null;LiczbaFolderow=$null;Status='Brak pomiaru';Uwagi='';FormulaUwagi='';Style=$null;Wysokosc=23.0}
            $klucz=Pobierz-kluczDnia $pusty
            if (-not $wedlugDnia.ContainsKey($klucz)) { $wedlugDnia.Add($klucz,$pusty) }
            if ($dzien -eq $koniec) { break }
        }
    }
    return @($wedlugDnia.Values | Sort-Object Data,Skrzynka,ZrodloDanych)
}

function Ustaw-zakresTabeliExcel {
    param($Tabela, [string]$Adres)
    $Tabela.TableXml.DocumentElement.SetAttribute('ref',$Adres)
    $filtr=$Tabela.TableXml.SelectSingleNode('/*[local-name()="table"]/*[local-name()="autoFilter"]')
    if ($null -ne $filtr) { $filtr.SetAttribute('ref',$Adres) }
}

function Zapisz-skoroszytEtapowy {
    param($Podsumowanie,[object[]]$Wiersze,[string]$PlikHistorii,[string]$IstniejacyPlik,[string]$PlikEtapowy)
    Przygotuj-modulExcel
    $pakiet=$null
    try {
        if ([IO.File]::Exists($IstniejacyPlik)) {
            try { $kontrola=[IO.File]::Open($IstniejacyPlik,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None); $kontrola.Dispose() }
            catch { throw 'Zamknij Raport-poczty.xlsx w Excelu i sprawdź prawo zapisu. Poprzednie raporty zostały zachowane.' }
            [IO.File]::Copy($IstniejacyPlik,$PlikEtapowy)
        } else {
            $szablon=Join-Path $KatalogProgramu 'Szablon-raportu.xlsx'
            if (-not [IO.File]::Exists($szablon)) { throw 'Brakuje pliku Szablon-raportu.xlsx. Rozpakuj cały pakiet programu.' }
            [IO.File]::Copy($szablon,$PlikEtapowy)
        }
        $pakiet=New-Object OfficeOpenXml.ExcelPackage -ArgumentList ([IO.FileInfo]$PlikEtapowy)
        $dziennik=$pakiet.Workbook.Worksheets['Dziennik']
        $dotychczasowe=@(Pobierz-dziennikZeSkoroszytu $dziennik)
        $historia=@(Import-Csv -LiteralPath $PlikHistorii -Delimiter ';' -Encoding UTF8)
        $dni=@(Polacz-pomiaryDzienne $dotychczasowe $historia $Podsumowanie)
        $koniec=[Math]::Max(7,6+$dni.Count)
        $koniecStary=$dziennik.Tables['TabelaDziennika'].Address.End.Row
        $dziennik.Cells[7,1,[Math]::Max($koniec,$koniecStary),10].Clear()
        $nr=7
        foreach ($dzien in $dni) {
            $dziennik.Cells[$nr,1].Value=$dzien.Data
            $dziennik.Cells[$nr,2].Value=$dzien.Skrzynka
            $dziennik.Cells[$nr,3].Value=$dzien.WszystkieWiadomosci
            $dziennik.Cells[$nr,4].Value=$dzien.NieprzeczytaneWiadomosci
            $dziennik.Cells[$nr,5].Value=$dzien.PrzeczytaneWiadomosci
            $dziennik.Cells[$nr,6].Value=$dzien.LiczbaFolderow
            if ($null -ne $dzien.Czas) { $dziennik.Cells[$nr,7].Value=$dzien.Czas.ToString('o') }
            $dziennik.Cells[$nr,8].Value=$dzien.ZrodloDanych
            $dziennik.Cells[$nr,9].Value=$dzien.Status
            $dziennik.Cells[$nr,10].Value=$dzien.Uwagi
            if (-not [string]::IsNullOrWhiteSpace($dzien.FormulaUwagi)) { $dziennik.Cells[$nr,10].Formula=$dzien.FormulaUwagi }
            if ($null -ne $dzien.Style) { for ($kol=1;$kol -le 10;$kol++) { $dziennik.Cells[$nr,$kol].StyleID=$dzien.Style[$kol-1] } }
            else {
                $dziennik.Cells[$nr,1,$nr,10].Style.Font.Name='Arial'; $dziennik.Cells[$nr,1,$nr,10].Style.Font.Size=11
                $dziennik.Cells[$nr,1,$nr,10].Style.VerticalAlignment='Center'
                $dziennik.Cells[$nr,10].Style.Fill.PatternType='Solid'
                $dziennik.Cells[$nr,10].Style.Fill.BackgroundColor.SetColor([Drawing.Color]::FromArgb(255,249,232))
            }
            $dziennik.Cells[$nr,1].Style.Numberformat.Format='dd.mm.yyyy'
            $dziennik.Cells[$nr,3,$nr,6].Style.Numberformat.Format='#,##0'
            $dziennik.Row($nr).Height=$dzien.Wysokosc
            $nr++
        }
        Ustaw-zakresTabeliExcel $dziennik.Tables['TabelaDziennika'] ('A6:J'+$koniec)
        foreach ($regula in $dziennik.ConditionalFormatting) {
            if ($regula.Address.Address -match '^I7(?::I\d+)?$') { $regula.Address=New-Object OfficeOpenXml.ExcelAddress ('I7:I'+$koniec) }
        }
        $foldery=$pakiet.Workbook.Worksheets['Foldery']
        if ($null -eq $foldery -or $null -eq $foldery.Tables['TabelaFolderow']) { throw 'Brakuje tabeli Foldery w skoroszycie. Zachowano dotychczasowy plik.' }
        $ostatni=[Math]::Max(7,6+$Wiersze.Count)
        $foldery.Cells[7,1,[Math]::Max($ostatni,$foldery.Tables['TabelaFolderow'].Address.End.Row),8].Clear()
        $nr=7
        foreach ($wiersz in $Wiersze) {
            $foldery.Cells[$nr,1].Value=$Podsumowanie.DataPomiaru
            $foldery.Cells[$nr,2].Value=$Podsumowanie.Skrzynka
            $foldery.Cells[$nr,3].Value=$wiersz.SciezkaFolderu
            $foldery.Cells[$nr,4].Value=$wiersz.WszystkieWiadomosci
            $foldery.Cells[$nr,5].Value=$wiersz.NieprzeczytaneWiadomosci
            $foldery.Cells[$nr,6].Value=$wiersz.PrzeczytaneWiadomosci
            $foldery.Cells[$nr,7].Value=$wiersz.WszystkieElementyOutlook
            $foldery.Cells[$nr,8].Value=$wiersz.NieprzeczytaneElementyOutlook
            $foldery.Row($nr).Height=23
            $nr++
        }
        $foldery.Cells[7,1,$ostatni,8].Style.Font.Name='Arial'
        $foldery.Cells[7,1,$ostatni,8].Style.Font.Size=11
        $foldery.Cells[7,4,$ostatni,8].Style.Numberformat.Format='#,##0'
        Ustaw-zakresTabeliExcel $foldery.Tables['TabelaFolderow'] ('A6:H'+$ostatni)
        $pakiet.Workbook.FullCalcOnLoad=$true
        $pakiet.Save()
    } finally { if ($null -ne $pakiet) { $pakiet.Dispose() } }
}

function Zapisz-raporty {
    param($Podsumowanie, [object[]]$Wiersze, [string]$KatalogRaportow)
    [void][IO.Directory]::CreateDirectory($KatalogRaportow)
    $idUruchomienia = [guid]::NewGuid().ToString('N')
    $nazwy = @('podsumowanie.csv', 'foldery.csv', 'historia.csv', 'Raport-poczty.xlsx')
    $plikiTymczasowe = @(); $kopieZapasowe = @(); $zapisanePliki = @()
    $przywracanieNieudane = $false
    try {
        foreach ($nazwa in $nazwy) {
            $rozszerzenie = '.tmp'
            if ($nazwa -eq 'Raport-poczty.xlsx') { $rozszerzenie = '.tmp.xlsx' }
            $plikiTymczasowe += Join-Path $KatalogRaportow ".$idUruchomienia.$nazwa$rozszerzenie"
            $kopieZapasowe += Join-Path $KatalogRaportow ".$idUruchomienia.$nazwa.bak"
        }
        Zapisz-csvUtf8 @($Podsumowanie) $plikiTymczasowe[0]
        $raportFolderow = @($Wiersze | ForEach-Object {
            [pscustomobject][ordered]@{
                DataPomiaru = $Podsumowanie.DataPomiaru
                Skrzynka = $Podsumowanie.Skrzynka
                SciezkaFolderu = $_.SciezkaFolderu
                WszystkieWiadomosci = $_.WszystkieWiadomosci
                NieprzeczytaneWiadomosci = $_.NieprzeczytaneWiadomosci
                PrzeczytaneWiadomosci = $_.PrzeczytaneWiadomosci
                WszystkieElementyOutlook = $_.WszystkieElementyOutlook
                NieprzeczytaneElementyOutlook = $_.NieprzeczytaneElementyOutlook
            }
        })
        Zapisz-csvUtf8 $raportFolderow $plikiTymczasowe[1]
        $sciezkaHistorii = Join-Path $KatalogRaportow $nazwy[2]
        if ([IO.File]::Exists($sciezkaHistorii)) {
            # Kopiuj historię bez importowania, zachowując dane, BOM i nagłówek.
            $dotychczasowaTresc = [IO.File]::ReadAllText($sciezkaHistorii, [Text.Encoding]::UTF8)
            $noweWiersze = @(Zabezpiecz-tekstDlaExcela @($Podsumowanie) | ConvertTo-Csv -Delimiter ';' -NoTypeInformation)
            $pierwszyWiersz = ($dotychczasowaTresc -split '\r?\n', 2)[0]
            if ($pierwszyWiersz -ne $noweWiersze[0]) { throw 'Kolumny w historia.csv różnią się od oczekiwanych. Przenieś dotychczasową historię do innego pliku.' }
            if (-not $dotychczasowaTresc.EndsWith("`n")) { $dotychczasowaTresc += "`r`n" }
            [IO.File]::WriteAllText($plikiTymczasowe[2], $dotychczasowaTresc + $noweWiersze[1] + "`r`n", (New-Object Text.UTF8Encoding($true)))
        } else { Zapisz-csvUtf8 @($Podsumowanie) $plikiTymczasowe[2] }

        Zapisz-skoroszytEtapowy $Podsumowanie $Wiersze $plikiTymczasowe[2] (Join-Path $KatalogRaportow 'Raport-poczty.xlsx') $plikiTymczasowe[3]
        for ($i = 0; $i -lt $nazwy.Count; $i++) {
            $plikDocelowy = Join-Path $KatalogRaportow $nazwy[$i]
            $istnial = [IO.File]::Exists($plikDocelowy)
            if ($istnial) { [IO.File]::Replace($plikiTymczasowe[$i], $plikDocelowy, $kopieZapasowe[$i]) }
            else { [IO.File]::Move($plikiTymczasowe[$i], $plikDocelowy) }
            $zapisanePliki += [pscustomobject]@{ Indeks = $i; Docelowy = $plikDocelowy; Istnial = $istnial }
        }
        Write-Host ('Skoroszyt Excel zapisany: ' + (Join-Path $KatalogRaportow 'Raport-poczty.xlsx'))
    } catch {
        $pierwotnyBlad = $_
        $bledyPrzywracania = @()
        foreach ($plik in @($zapisanePliki | Sort-Object Indeks -Descending)) {
            try {
                if ($plik.Istnial) {
                    # PS 5.1 zamienia pusty argument tekstowy null na nieprawidłową pustą ścieżkę.
                    # Zapisz zastępowany nowy plik w ścieżce tymczasowej, a potem go usuń.
                    [IO.File]::Replace($kopieZapasowe[$plik.Indeks], $plik.Docelowy, $plikiTymczasowe[$plik.Indeks])
                }
                else { [IO.File]::Delete($plik.Docelowy) }
            } catch { $bledyPrzywracania += $_.Exception.Message }
        }
        if ($bledyPrzywracania.Count -gt 0) {
            $przywracanieNieudane = $true
            $zachowaneKopie = @($kopieZapasowe | Where-Object { [IO.File]::Exists($_) })
            throw ("Błąd zapisu i przywracania CSV. Zachowane kopie zapasowe: " + ($zachowaneKopie -join ', ') + ". Pierwotny błąd: " + $pierwotnyBlad.Exception.Message + ". Błędy przywracania: " + ($bledyPrzywracania -join '; '))
        }
        throw $pierwotnyBlad
    } finally {
        foreach ($sciezka in $plikiTymczasowe) {
            try { if ([IO.File]::Exists($sciezka)) { [IO.File]::Delete($sciezka) } }
            catch { Write-Warning ("Nie udało się usunąć pliku tymczasowego: " + $sciezka) }
        }
        # Po zapisie lub udanym przywróceniu usuń tylko kopie z tego uruchomienia.
        if (-not $przywracanieNieudane) {
            foreach ($sciezka in $kopieZapasowe) {
                try { if ([IO.File]::Exists($sciezka)) { [IO.File]::Delete($sciezka) } }
                catch { Write-Warning ("Nie udało się usunąć kopii zapasowej: " + $sciezka) }
            }
        }
    }
}

function Uruchom-pomiar {
    $outlook = $null; $sesja = $null; $magazyny = $null; $magazyn = $null; $folderGlowny = $null; $blokada = $null
    try {
        $ustawienia = Get-Content -LiteralPath $PlikUstawien -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $ustawienia -or -not ($ustawienia.PSObject.Properties.Name -contains 'NazwaSkrzynki') -or $ustawienia.NazwaSkrzynki -isnot [string]) {
            throw 'Plik ustawienia.json musi zawierać pole NazwaSkrzynki z wartością tekstową. Dopuszczalna jest pusta nazwa.'
        }
        if ($TestDemonstracyjny) {
            $ustawienia = [pscustomobject]@{ TrybPolaczenia = 'Demonstracja'; NazwaSkrzynki = '' }
        } elseif ($Menu) {
            $ustawienia = Wybierz-trybUruchomienia $ustawienia
            if ($null -eq $ustawienia) { Write-Host 'Zakończono bez wykonywania pomiaru.'; return }
        }
        $nazwaSkrzynki = $ustawienia.NazwaSkrzynki
        $tryb = [string](Pobierz-ustawienie $ustawienia 'TrybPolaczenia' 'Microsoft365')
        if ($tryb -notin @('Microsoft365', 'OutlookKlasyczny', 'Demonstracja')) {
            throw 'TrybPolaczenia musi mieć wartość Microsoft365, OutlookKlasyczny albo Demonstracja.'
        }
        if ($tryb -eq 'Microsoft365') {
            $konto = Polacz-zMicrosoft365 $ustawienia -WybierzKonto:$Menu
            if ([string]::IsNullOrWhiteSpace($konto)) { Write-Host 'Zakończono bez wykonywania pomiaru.'; return }
            if ($PokazSkrzynki) {
                Write-Host 'Skrzynka podstawowa zalogowanego konta Microsoft. Tę nazwę można wpisać w NazwaSkrzynki:'
                Write-Host ('  ' + $konto)
                return
            }
            if (-not [string]::IsNullOrWhiteSpace($nazwaSkrzynki) -and -not [string]::Equals($konto, $nazwaSkrzynki, [StringComparison]::OrdinalIgnoreCase)) {
                throw ("Zalogowane konto '$konto' różni się od NazwaSkrzynki '$nazwaSkrzynki'. Wpisz właściwy adres konta w ustawienia.json lub zaloguj się do właściwego konta Microsoft 365.")
            }
        } elseif ($tryb -eq 'Demonstracja') {
            Write-Host 'TEST DEMONSTRACYJNY: używane są tylko dane przykładowe, bez logowania i odczytu poczty.'
            if ($PokazSkrzynki) { Write-Host 'DANE_TESTOWE'; return }
        } else { Sprawdz-klasycznyOutlook }
        if (-not $PokazSkrzynki) {
            $nazwaKataloguRaportow = 'raporty'
            if ($tryb -eq 'Demonstracja') { $nazwaKataloguRaportow = 'raporty-testowe' }
            $katalogRaportow = Join-Path $KatalogProgramu $nazwaKataloguRaportow
            [void][IO.Directory]::CreateDirectory($katalogRaportow)
            try {
                $blokada = [IO.File]::Open((Join-Path $katalogRaportow '.blokada-uruchomienia'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            } catch { throw ('Inne liczenie jest już uruchomione lub folder ' + $nazwaKataloguRaportow + ' jest niedostępny.') }
        }
        if ($tryb -eq 'Demonstracja') {
            $wiersze = @(Pobierz-daneDemonstracyjne)
            $podsumowanie = Utworz-podsumowanie $wiersze 'DANE_TESTOWE' ([DateTimeOffset]::Now.ToString('o')) 'TrybDemonstracyjny'
            Zapisz-raporty $podsumowanie $wiersze $katalogRaportow
            Write-Host 'TEST: wszystkie wiadomości = 10; nieprzeczytane = 3; przeczytane = 7; foldery = 3.'
            Write-Host ('Raporty testowe zapisane w: ' + $katalogRaportow)
            return
        }
        if ($tryb -eq 'Microsoft365') {
            Write-Host ('Liczenie wiadomości bezpośrednio na serwerze pocztowym Microsoft: ' + $konto)
            $wiersze = @(Pobierz-daneFolderowMicrosoft365 -Skrzynka $konto)
            $podsumowanie = Utworz-podsumowanie $wiersze $konto ([DateTimeOffset]::Now.ToString('o')) 'Microsoft365Serwer'
            Zapisz-raporty $podsumowanie $wiersze $katalogRaportow
            Write-Host ("Wszystkie wiadomości: {0}; nieprzeczytane: {1}; przeczytane: {2}; foldery: {3}." -f $podsumowanie.WszystkieWiadomosci, $podsumowanie.NieprzeczytaneWiadomosci, $podsumowanie.PrzeczytaneWiadomosci, $podsumowanie.LiczbaFolderow)
            Write-Host ('Raporty zapisane w: ' + $katalogRaportow)
            return
        }
        try { $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
        catch { $outlook = New-Object -ComObject Outlook.Application }
        $sesja = $outlook.GetNamespace('MAPI')
        $magazyny = $sesja.Stores
        $skrzynki = @(
            for ($i = 1; $i -le $magazyny.Count; $i++) {
                $magazynKandydata = $null
                try {
                    $magazynKandydata = $magazyny.Item($i)
                    [pscustomobject]@{ Indeks = $i; Nazwa = [string]$magazynKandydata.DisplayName }
                } finally { Zwolnij-referencjeCom $magazynKandydata }
            }
        )
        if ($PokazSkrzynki) {
            Write-Host 'Dostępne skrzynki i pliki danych. Skopiuj jedną nazwę do pola NazwaSkrzynki:'
            foreach ($kandydat in $skrzynki) { Write-Host ("  " + $kandydat.Nazwa) }
            return
        }
        if ($Menu -and $skrzynki.Count -gt 0) {
            $wybranaSkrzynka = Wybierz-skrzynkeZListy $skrzynki
            if ($null -eq $wybranaSkrzynka) { Write-Host 'Zakończono bez wykonywania pomiaru.'; return }
        } else { $wybranaSkrzynka = Wybierz-skrzynke $skrzynki $nazwaSkrzynki }
        $magazyn = $magazyny.Item($wybranaSkrzynka.Indeks)
        $folderGlowny = $magazyn.GetRootFolder()
        $pominiete = @()
        # Foldery problemów synchronizacji i diagnostyki nie są folderami poczty użytkownika.
        foreach ($numerFolderu in @(19, 20, 21, 22)) {
            $folderSystemowy = $null
            try {
                $folderSystemowy = $magazyn.GetDefaultFolder($numerFolderu)
                if ($null -ne $folderSystemowy) { $pominiete += [string]$folderSystemowy.EntryID }
            } catch {
                # Nie każdy typ magazynu udostępnia te opcjonalne foldery domyślne.
                if (-not (Sprawdz-brakWlasciwosciMapi $_.Exception)) { throw }
            } finally { Zwolnij-referencjeCom $folderSystemowy }
        }
        Write-Host ("Liczę wiadomości w skrzynce: " + $wybranaSkrzynka.Nazwa)
        $wiersze = @(Pobierz-daneFolderow $folderGlowny $pominiete)
        $podsumowanie = Utworz-podsumowanie $wiersze $wybranaSkrzynka.Nazwa ([DateTimeOffset]::Now.ToString('o'))
        Zapisz-raporty $podsumowanie $wiersze $katalogRaportow
        Write-Host ("Wszystkie wiadomości: {0}; nieprzeczytane: {1}; przeczytane: {2}; foldery: {3}." -f $podsumowanie.WszystkieWiadomosci, $podsumowanie.NieprzeczytaneWiadomosci, $podsumowanie.PrzeczytaneWiadomosci, $podsumowanie.LiczbaFolderow)
        Write-Host ("Raporty zapisane w: " + $katalogRaportow)
    } finally {
        if ($null -ne $blokada) { $blokada.Dispose() }
        Zwolnij-referencjeCom $folderGlowny
        Zwolnij-referencjeCom $magazyn
        Zwolnij-referencjeCom $magazyny
        Zwolnij-referencjeCom $sesja
        Zwolnij-referencjeCom $outlook
        # Nie wywołuj Outlook.Quit(): sesja użytkownika pozostaje otwarta.
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    try { Uruchom-pomiar }
    catch {
        Write-Error ("Liczenie nie zostało zakończone. Sprawdź plik Instrukcja.md. " + $_.Exception.Message) -ErrorAction Continue
        exit 1
    }
}
