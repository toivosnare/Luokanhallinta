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
5. Luo välilyönnein erotettu "luokka.csv" tiedosto sarakkeilla "Name", "Mac", "Row", "Column" skriptien kanssa samaan kansioon.
```
"Name" "Mac" "Column" "Row"
"localhost" "A1:B2:C3:D4:E5:F6" "1" "1"
"VKY00000" "A2:B3:C4:D5:E6:F7" "2" "1"
"10.132.0.1" "A3:B4:C5:D6:E7:F8" "1" "2"
"192.168.0.1" "A4:B6:C6:D7:E8:F9" "2" "2"
```
6. Korjaa hallinta.ps1-skriptin asetuket:
    - Esimerkiksi rivillä 512 määritellään addon synkkauksen lähdekansio ja sinne pääsyyn käytettävät käyttäjätunnukset. Jos addonit synkataan Panssariprikaatin nassilta, älä käytä oletus "testi" käyttäjää vaan vaihda ne oman joukko-osastosi tunnuksiin. Jos sellaisia ei ole, pyydä ne Parolasta. Jos lähdekansioon ei tarvitse käyttäjätunnuksia, username ja password parametreja ei tarvitse määritellä.
    ```PowerShell
    @{Name="Synkaa addonit"; Click={Copy-ItemToTarget -target ([Host]::GetActive()) -source "\\10.132.0.97\Addons" -destination "%programfiles%\Bohemia Interactive Simulations\VBS3 3.9.0.FDF EZYQC_FI\mycontent\addons" -username "WORKGROUP\testi" -password "pleasedonotuse" -parameter "/MIR /XO /NJH"}}
    ```
    - Rivillä 531 määritellään käyttäjätunnukset, jolla luokanhallintaa käytetään. Määrittele Init komennon parametreiksi sellaiset käyttäjätunnukset, joilla on järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin. $(whoami) hakee hallintakoneelle kirjautuneen käyttäjänimen automaattisesti, joka on toimiva ratkaisu jos kaikilla luokan koneilla on vastaavan niminen järjestelmänvalvojakäyttäjä.
    ```PowerShell
    @{Name="Vaihda käyttäjä"; Init={Set-ScriptCredential -username $(whoami) -password ""}; Click={Set-ScriptCredential}}
    ```
7. Aja run-skripti järjestelmänvalvojana (pikakuvakkeen luonti on järkevä idea):
```PowerShell
.\run.ps1
```
8. Raportoi kuinka mikään ei toimi
```
pls fix
```

## Pikakuvakkeen luonti ja käynnistysparametrit
Jotta luokanhallinnan saa käynnistettyä järjestelmänvalvojan oikeuksilla, tulee pikakuvake luoda run-skriptin sijaan powershell.exeen. Pikakuvakkeen target kenttään tulee siis kirjoittaa "powershell C:\path\to\run.ps1 -path C:\path\to\class.csv". Target kentän loppuun kannataa lisätä käynnistyparametri "-path", joka tarkoittaa luokkatiedoston sijaintia. Jos luokkatiedostoa ei ole erikseen määritelty, ohjelma yrittää lukea run-skriptin kansiossa olevaa tiedostoa "luokka.csv". Viimeiseksi valitse pikakuvakkeen asetuksien "Shortcut" välilehdeltä "Advanced" ja varmista että "Run as administrator" on valittuna.

## Huomautuksia
* Käyttäjätunnuksilla, joilla käytetään luokanhallintaa tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin, jotta etäkomentojen ajaminen onnistuu.
* Luokanhallinta pitää käynnistää run.ps1 skriptin kautta.
* Luokanhallinta ei toimi, jos hallittavan tietokoneen salasana on vanhentunut.
* Jos luokanhallinta on ollut käyttämättömänä auki pitemmän ajan, kannattaa se käynnistää uudestaan tai vähintääkin päivittää painamalla F5.
