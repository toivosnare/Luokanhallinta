# Luokanhallinta

blablabla

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

3. Aseta WinRM käynistymään heti tietokoneen käynistyksen yhteydessä
```PowerShell
sc.exe config "WinRM" start=auto
```

### Seuraavat koskevat ainoastaan niitä koneita, josta luokanhallintaa tullaan käyttämään:

4. Lisää hallintakoneen TrustedHosts listaan luokan koneet
```PowerShell
Set-Item WSMan:\localhost\Client\TrustedHosts -value "10.132.0.*" # Esim.
```

5. Luo välilyönnein erotettu "luokka.csv" tiedosto sarakkeilla "Nimi", "Mac", "Sarake", "Rivi" skriptien kanssa samaan kansioon
```
Nimi Mac Sarake Rivi
localhost A1:B2:C3:D4:E5:F6 1 1
VKY00000 A2:B3:C4:D5:E6:F7 2 1
VKY00001 A3:B4:C5:D6:E7:F8 1 2
```

6. Aja run-skripti:
```PowerShell
.\run.ps1
```

7. Raportoi kuinka mikään ei toimi
```
pls fix
```
