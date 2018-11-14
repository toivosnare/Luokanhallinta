#### Luokanhallinta ####

blablabla

## Setup ##

Aja seuraavat komennot Admin PowerShellissä kaikissa luokan tietokoneissa:
```PowerShell
Enable-PSRemoting
Set-ExecutionPolicy RemoteSigned
```

Lisäksi jos tietokoneet eivät kuulu domain ympäristöön, lisää hallintakoneen TrustedHosts listaan luokan koneet:
```PowerShell
Set-Item WSMan:\localhost\Client\TrustedHosts -value "10.132.0.*" # Esim.
```

Luo välilyönnein erotettu luokka.csv-tiedosto skriptien kanssa samaan kansioon

Aja run-skripti:
```Batch
powershell .\run.ps1
```
