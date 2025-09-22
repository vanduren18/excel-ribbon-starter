
# Excel Online Ribbon Add-in (Starter)

Deze starter maakt een eigen tab in het lint met de knoppen **Connectie**, **Download**, **Functie** en **Help**.

## Bestanden
- `manifest.xml` – definieert de add-in en de knoppen
- `commands.html` + `commands.js` – handlers voor de knoppen (behalve Help)
- `taskpane.html` + `taskpane.js` – inhoud van het Help-venster (task pane)
- `assets/icon-*.png` – simpele iconen

## Snel aan de slag (Excel Online)
> Excel op het web kan **geen** `http://localhost` laden. Host de bestanden publiek via **HTTPS** (bijv. GitHub Pages).

1. **Uploaden naar GitHub**
   - Maak een nieuw openbare repository, bijv. `excel-ribbon-starter`.
   - Upload alle bestanden uit deze map.

2. **GitHub Pages inschakelen**
   - Ga in de repo naar **Settings → Pages**.
   - Kies **Source: Deploy from a branch** en selecteer de `main` branch, map `/ (root)`.
   - Na enkele seconden krijg je een URL, bijv. `https://<jouw-naam>.github.io/excel-ribbon-starter`.

3. **URLs bijwerken**
   - Open `manifest.xml` en vervang `https://REPLACE_ME.your.host` door jouw Pages-URL.
   - Commmit en push de wijziging (of upload het bestand opnieuw).

4. **Sideloaden in Excel Online**
   - Open Excel Online → **Invoegen** → **Office-invoegtoepassingen** → **Mijn invoegtoepassingen** → **Upload (Custom Add-in)**.
   - Selecteer je **`manifest.xml`**.
   - Je ziet nu een tab **“Mijn Add-in”** met de knoppen.

## Gebruik
- **Connectie**: laat een mini-dialog zien ("Connectietest geslaagd").
- **Download**: plaatst demo-data in A1:C4 en autofit kolommen.
- **Functie**: vermenigvuldigt geselecteerde getallen ×2.
- **Help**: opent het taakvenster met deze uitleg.

## Veelvoorkomende problemen
- **Blanco knoppen / niets gebeurt**: controleer dat alle URLs in `manifest.xml` naar jouw **HTTPS**-host wijzen.
- **CORS of mixed content**: alle pagina's en scripts moeten via **HTTPS** komen.
- **Console-fouten**: open DevTools (F12) in je browser en bekijk de console.
- **Naamhandler mismatch**: de `FunctionName` in `manifest.xml` moet exact overeenkomen met functie-namen in `commands.js`.
- **Cache**: forceer herladen (Ctrl/Cmd+Shift+R) of wijzig de bestandsnaam/versie indien updates niet zichtbaar zijn.

## Alternatieven voor hosting
- **Azure Static Web Apps** of **Azure Storage Static Website**.
- **SharePoint App Catalog / Asset Library** in M365-tenant (voor organisatiebrede distributie).

## Desktop Excel (optioneel)
Op Windows/Mac kun je tijdens ontwikkeling ook lokaal hosten (https://localhost) en sideloaden met de **Yeoman generator** of **Office Add-in Debugger**. Excel **op het web** vereist echter een publiek bereikbare HTTPS-host.

Succes!
