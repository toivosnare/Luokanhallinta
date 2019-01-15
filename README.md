![screenshot2](/uploads/c7cd3deb16e02fd2ba0b3df62560684f/screenshot2.png)

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
1. Hyväksy ulkoisten skriptien ajaminen.
```PowerShell
Set-ExecutionPolicy RemoteSigned
```
2. Ota WinRM käyttöön. Tämä on mahdollista vain, jos käytetty verkko on merkitty Windowsissa yksityiseksi. Tarkastuksen voi kuitenkin ohittaa -SkipNetworkProfileCheck asetuksella.
```PowerShell
Enable-PSRemoting -SkipNetworkProfileCheck
```
3. Aseta WinRM käynistymään heti tietokoneen käynistyksen yhteydessä (oletuksena startup type on "Automatic Delayed Start").
```PowerShell
sc.exe config "WinRM" start=auto
```
### Seuraavat koskevat ainoastaan niitä koneita, josta luokanhallintaa tullaan käyttämään (kouluttajan kone):
4. Lataa [run.ps1](/run.ps1) ja [hallinta.ps1](/hallinta.ps1) samaan kansioon.
5. Lisää hallintakoneen TrustedHosts listaan luokan koneet.
```PowerShell
Set-Item WSMan:\localhost\Client\TrustedHosts -value "10.132.0.*" # Esim.
```
6. Luo välilyönnein erotettu .csv luokkatiedosto sarakkeilla "Name", "Mac", "Row", "Column" ([esimerkki](/luokka.csv)).
7. Avaa [run.ps1](/run.ps1) tekstieditorissa. Laita $classFilePath muuttujan arvoksi luomasi luokkatiedoston polku (jos luokkatiedostoa ei määritellä erikseen, ohjelma pyytää käyttäjää määrittelemään sen käynnistyksen yhteydessä). Määritä myös $username ja $password muuttujilla käyttäjätunnukset, joilla etäkomennot ajetaan (jos käyttäjätunnuksia ei määritellä erikseen, ohjelma pyytää käyttäjää määrittelemään ne käynnistyksen yhteydessä). Käyttäjällä tulee olla järjestelmänvalvojan oikeudet hallitaviin tietokoneisiin. Lisäksi $addonSyncPath muuttujaan tulee polku, josta addonit synkataan. Seuraavat $addonSyncUsername ja $addonSyncPassword muuttujat voi jättää tyhjäksi jos addon lähdekansioon ei tarvitse käyttäjätunnuksia (eli se on jaettu julkisesti kaikille).
8. Luo työpöydän pikakuvake. Jotta luokanhallinnan saa käynnistettyä järjestelmänvalvojan oikeuksilla, tulee pikakuvake luoda [run.ps1](/run.ps1)-skriptin sijaan powershell.exeen. Pikakuvakkeen target kenttään tulee siis kirjoittaa "powershell C:\path\to\run.ps1". Paina pikakuvakkeen ominaisuuksien "Shortcut" välilehdeltä "Advanced" ja varmista että "Run as administrator" on valittuna.

## Huomautuksia
* Käyttäjätunnuksilla, joilla käytetään luokanhallintaa tulee olla järjestelmänvalvojan oikeudet hallittaviin tietokoneisiin, jotta etäkomentojen ajaminen onnistuu.
* Luokanhallinta pitää käynnistää [run.ps1](/run.ps1) skriptin kautta.
* Luokanhallinta ei toimi, jos hallittavan tietokoneen salasana on vanhentunut.
* Jos luokanhallinta on ollut käyttämättömänä auki pitemmän ajan, kannattaa se käynnistää uudestaan tai vähintääkin päivittää painamalla F5.
* F-Securen automaattinen päivitys tarvitsee toimiakseen päivitystyökalun (fsdbupdate9.exe), jonka voi ladata F-Securen [nettisivuilta](https://www.f-secure.com/en/web/labs_global/database-updates). Sijoita tiedosto F-Securen juurikansioon (C:\Program Files (x86)\F-Secure).
