# Vermaat Handtekening, Outlook invoegtoepassing

Voegt de Nico Vermaat e-mailhandtekening automatisch in bij elke nieuwe mail, en wisselt mee als je het Van-adres verandert.

## Wat hij doet

* **Automatisch invoegen** bij een nieuwe mail, een antwoord of een doorgestuurd bericht.
* **Versie op basis van Van-adres**
  * `@nicovermaat.nl` krijgt **Privé** (naam, privénummer, e-mail, adres)
  * `@vermaat-it.nl` krijgt **Nico Vermaat IT Consultancy** (logo, functie, slogan, zakelijk nummer, website, certificeringen, bedrijfsblok met KVK)
  * Wissel je het Van-adres tijdens het schrijven, dan wisselt de handtekening mee.
* **Knop "Handtekening"** in het lint van een nieuwe mail. Die opent een paneel waarin je:
  * de versie kiest: Privé, Nico Vermaat IT Consultancy of Nico Vermaat IT
  * per versie onderdelen aan of uit zet (woordmerk, functie, slogan, telefoonnummers, adres, website, certificeringen, KVK, bedrijfsnaam onderaan)
  * je gegevens aanpast
  * met **Invoegen** de handtekening in de huidige mail zet
  * met **Opslaan als standaard** je keuzes bewaart voor alle volgende mails. Ze worden per mailbox opgeslagen en werken op elk apparaat waar de invoegtoepassing staat.
* Zet je eigen Outlook handtekening uit in mails waar hij invoegt, zodat je geen dubbele handtekening krijgt. Dit kun je in het paneel uitschakelen.

## Werkt in

* Outlook op het web (Microsoft 365)
* Nieuwe Outlook voor Windows
* Klassiek Outlook voor Windows (Microsoft 365, recente versie)
* Outlook voor Mac (nieuwe Outlook)

Niet in de Outlook app op telefoon of tablet. Daar gebruik je de WhatsApp tekstversie of de handtekening uit de app zelf.

Vereist een Microsoft 365 of Exchange Online mailbox.

## Bestanden

| Bestand | Doel |
|---|---|
| `manifest.xml` | De invoegtoepassing zelf. Dit installeer je in Outlook. |
| `taskpane.html` | Het paneel achter de knop Handtekening |
| `commands.html` | Laadt het automatisch invoegen in Outlook op het web en nieuwe Outlook |
| `launchevent.js` | Automatisch invoegen, ook voor klassiek Outlook |
| `signature.js` | De handtekening zelf, gedeeld door alle onderdelen |
| `assets/` | Logo, iconen voor telefoon en e-mail, knopiconen |
| `set-host.ps1`, `set-host.sh` | Zet je webadres in de bestanden |

## Installeren

### 1. Kies een HTTPS adres
Outlook laadt de invoegtoepassing vanaf een website. Het logo en de iconen in je handtekening komen ook van dit adres, dus het moet blijvend online staan. Voorbeelden:

* `https://www.vermaat-it.nl/handtekening` (een map op je eigen website)
* GitHub Pages of Azure Static Web Apps (beide gratis)

### 2. Zet het adres in de bestanden
Pak de zip uit, open PowerShell in de map en voer uit:

```powershell
.\set-host.ps1 https://www.vermaat-it.nl/handtekening
```

Mag PowerShell geen scripts uitvoeren, gebruik dan eenmalig:
`powershell -ExecutionPolicy Bypass -File .\set-host.ps1 https://www.vermaat-it.nl/handtekening`

### 3. Upload alle bestanden
Zet alle bestanden en de map `assets` op dat adres. Controleer in je browser of `https://www.vermaat-it.nl/handtekening/taskpane.html` opent.

### 4. Installeer in Outlook

**Alleen voor jezelf**
1. Ga naar https://aka.ms/olksideload. Outlook op het web opent met het venster Invoegtoepassingen.
2. Kies **Mijn invoegtoepassingen**, dan onderaan **Aangepaste invoegtoepassing toevoegen** en **Toevoegen uit bestand**.
3. Kies `manifest.xml` en bevestig.

De invoegtoepassing verschijnt daarna ook in nieuwe en klassieke Outlook met hetzelfde account.

**Voor je hele tenant (beheerder)**
1. Microsoft 365 beheercentrum, **Instellingen**, **Geïntegreerde apps**, **Aangepaste apps uploaden**.
2. Kies **Office invoegtoepassing**, upload `manifest.xml` en wijs gebruikers toe.
3. Het kan tot 24 uur duren voordat hij overal zichtbaar is.

### 5. Testen
Open een nieuwe mail vanuit `info@vermaat-it.nl`: de zakelijke handtekening verschijnt. Wissel het Van-adres naar `info@nicovermaat.nl`: de privéversie verschijnt.

## Aanpassen

* Gegevens en vinkjes: via het paneel, knop **Opslaan als standaard**.
* Standaardwaarden voor nieuwe installaties: in `signature.js` onder `DEFAULTS`. Neem na een wijziging `signature.js` ook over in `launchevent.js` (dat bestand is `signature.js` en `events.js` achter elkaar), of vraag Claude het opnieuw te bouwen.
* Nieuwe versie uitrollen: verhoog `<Version>` in `manifest.xml`.

## Problemen oplossen

* **Geen handtekening bij nieuwe mail:** controleer of `commands.html` en `launchevent.js` bereikbaar zijn op je adres en of **Handtekening automatisch invoegen** in het paneel aan staat.
* **Dubbele handtekening:** zet in het paneel **Eigen Outlook handtekening uitzetten** aan, of verwijder de standaardhandtekening in de Outlook instellingen.
* **Logo niet zichtbaar bij ontvanger:** sommige mailprogramma's laden afbeeldingen pas na een klik. Het adres moet online blijven.
