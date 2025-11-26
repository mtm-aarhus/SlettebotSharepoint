# SharePoint Sletningsrobot – Procesbeskrivelse

Denne robot er udviklet til at slette specifikke mapper og deres indhold i SharePoint som en del af oprydning efter afsluttede sager. Den anvender OpenOrchestrators kø-framework og SharePoint Online via Office365 API.

---

## Formål

Robotten sletter mapper relateret til aktindsigtssager i følgende biblioteker:
- `Delte dokumenter/Aktindsigter/{mappenavn}`
- `Delte dokumenter/Dokumentlister/{mappenavn}`

---

## Kødata

Robotten forventer, at hvert køelement indeholder følgende JSON:
```json
{
  "SharepointMappeNavn": "<mappenavn>"
}
