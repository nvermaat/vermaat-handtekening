# Sjablonen

Sjablonen voor de knop Handtekening, tab **Sjablonen**.

## Een sjabloon toevoegen
1. Maak een HTML bestand in deze map, bijvoorbeeld `offerte.html`. Gebruik gewone alinea's: `<p>...</p>`.
2. Zet het in `index.json` met een `id`, `title`, `desc`, `subject` en `file`.
3. Commit en push. Na een minuut staat het sjabloon in Outlook.

## Invulvelden
* `{{Naam}}` wordt een invulveld met de naam Naam.
* `{{Samenvatting:lang}}` wordt een groot invulveld voor meerdere regels.
* Invulvelden mogen ook in het onderwerp staan.
* Een veld dat je leeg laat, blijft als `[Naam]` in de mail staan, zodat je het niet vergeet.
