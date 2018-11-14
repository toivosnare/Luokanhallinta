#### Luokanhallinta

blablabla

## Setup

* Aja seuraavat komennot Admin PowerShellissä kaikissa luokan tietokoneissa:
```PowerShell
Enable-PSRemoting
# Set-ExecutionPolicy RemoteSigned ???
```
* Lisäksi jos tietokoneet eivät kuulu domain ympäristöön, lisää hallintakoneen TrustedHosts listaan luokan koneet:
```PowerShell
# Esim. $lista = 10.132.0.*
Set-Item WSMan:\localhost\Client\TrustedHosts -value "$lista"
```