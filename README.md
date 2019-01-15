![screenshot](/uploads/654cd95f47988ac1d6febd0ac5aa856d/screenshot.png)

# Luokanhallinta

## Toteutus
Luokanhallinta käyttää Windowsiin sisäänrakennettua WinRM-protokollaa tietokoneiden etähallintaan. Ohjelmointikielenä on Windows PowerShell, jossa on monia valmiita Windowsin etähallintaa hyödyntäviä komentoja.

## Ominaisuudet
* Tietokoneiden etäkäynnistys ja -sammutus
* Etäkomentojen ajaminen
* Etäohjelmien käynistys (interaktiivisesti paikallisessa sessiossa)
* Tiedostojen kopioiminen (esim. addon synkkaus)

## Asennus
### Aja seuraavat komennot Admin PowerShellissä kaikissa luokan tietokoneissa:
1. Hyväksy ulkoisten skriptien ajaminen
```PowerShell
Set-ExecutionPolicy RemoteSigned
```
2. Ota WinRM käyttöön. Tämä on mahdollista vain, jos käytetty verkko on merkitty Windowsissa yksityiseksi. Tarkastuksen voi kuitenkin ohittaa -SkipNetworkProfileCheck asetuksella.
```PowerShell
Enable-PSRemoting -SkipNetworkProfileCheck
```
3. Aseta WinRM käynistymään heti tietokoneen käynistyksen yhteydessä. (Oletuksena startup type on "Automatic Delayed Start")
```PowerShell
sc.exe config "WinRM" start=auto
```
### Seuraavat koskevat ainoastaan niitä koneita, josta luokanhallintaa tullaan käyttämään (kouluttajan kone):
4. Lisää hallintakoneen TrustedHosts listaan luokan koneet
```PowerShell
Set-Item WSMan:\localhost\Client\TrustedHosts -value "10.132.0.*" # Esim.
```
5. Luo välilyönnein erotettu luokkatiedosto sarakkeilla "Name", "Mac", "Row", "Column".
```
"Name" "Mac" "Column" "Row"
"localhost" "A1:B2:C3:D4:E5:F6" "1" "1"
"VKY00000" "A2:B3:C4:D5:E6:F7" "2" "1"
"10.132.0.1" "A3:B4:C5:D6:E7:F8" "1" "2"
"192.168.0.1" "A4:B6:C6:D7:E8:F9" "2" "2"
```
6. Avaa run.ps1 tiedosto tekstieditorissa. Laita $classFilePath muuttujan arvoksi luomasi luokkatiedoston polku (jos luokkatiedostoa ei määritellä erikseen, ohjelma pyytää käyttäjää määrittelemään sen käynnistyksen yhteydessä). Määritä myös $username ja $password muuttujilla käyttäjätunnukset, joilla etäkomennot ajetaan (jos käyttäjätunnuksia ei määritellä erikseen, ohjelma pyytää käyttäjää määrittelemään ne käynnistyksen yhteydessä). Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallitaviin tietokoneisiin. Lisäksi $addonSyncPath muuttujaan tulee polku, josta addonit synkataan. Seuraavat $addonSyncUsername ja $addonSyncPassword muuttujat voi jättää tyhjäksi jos addon lähdekansioon ei tarvitse käyttäjätunnuksia (eli se on jaettu kaikille).
```PowerShell
$classFilePath = "C:\Users\Uzer\Documents\Luokanhallinta\luokka.csv"
$username = $(whoami.exe) # Gets username of currently logged on user
$password = ""
$addonSyncPath = "\\10.132.0.97\Addons"
$addonSyncUsername = ""
$addonSyncPassword = ""
```
7. Aja run-skripti järjestelmänvalvojana (pikakuvakkeen luonti on järkevä idea):
```PowerShell
.\run.ps1
```
8. Raportoi kuinka mikään ei toimi
```
pls fix
```

## Pikakuvakkeen luonti
Jotta luokanhallinnan saa käynnistettyä järjestelmänvalvojan oikeuksilla, tulee pikakuvake luoda run-skriptin sijaan powershell.exeen. Pikakuvakkeen target kenttään tulee siis kirjoittaa "powershell C:\path\to\run.ps1". Paina pikakuvakkeen ominaisuuksien "Shortcut" välilehdeltä "Advanced" ja varmista että "Run as administrator" on valittuna.

## Huomautuksia
* Käyttäjätunnuksilla, joilla käytetään luokanhallintaa tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin, jotta etäkomentojen ajaminen onnistuu.
* Luokanhallinta pitää käynnistää run.ps1 skriptin kautta.
* Luokanhallinta ei toimi, jos hallittavan tietokoneen salasana on vanhentunut.
* Jos luokanhallinta on ollut käyttämättömänä auki pitemmän ajan, kannattaa se käynnistää uudestaan tai vähintääkin päivittää painamalla F5.
* F-Securen automaattinen päivitys tarvitsee toimiakseen päivitystyökalun (fsdbupdate9.exe), jonka voi ladata F-Securen nettisivuilta (https://www.f-secure.com/en/web/labs_global/database-updates). Sijoita tiedosto F-Securen juurikansioon (C:\Program Files (x86)\F-Secure).
