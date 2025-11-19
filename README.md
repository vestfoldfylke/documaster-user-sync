# documaster-user-sync
Script som håndterer brukersynk i Documaster


Flyt
  - Hent alle rader fra Sharepoint-liste som arkivet styrer
    ○ En rad representerer en gruppe i Entra som også er synket over i Documaster
    ○ Kolonner som brukes av Logic App
      § Navn / Tittel (et forståelig navn på gruppa)
      § Personer som har / skal ha tilgang (multiple people picker)
  - For hver rad / gruppe:
    ○ Finn korrekt id for gruppa basert på gruppe-mapping (oppslag på tittel) i Logic appen
    ○ Hent gruppemedlemmer fra Entra (basert på groupId)
    ○ For hver person i "har tilgang" kolonnen på nåværende gruppe:
      § Finnes personen allerede i Entra-gruppen?
        □ Ja: Ok, hopp over
        □ Nei:
          ® Try: Legg til i Entra-gruppen
          ® Catch: brukeren finnes ikke i Entra lenger
            ◊ Fjern brukeren fra Sharepoint-listen (eller legg de til i en liste over de som må fjernes, og bruk denne også i steget under)
        
    ○ For hver person i Entra-gruppen
      § Finnes personen i "har tilgang" kolonnen?
        □ Ja: Ok, hopp over
        □ Nei: Fjern brukeren fra Entra-gruppen
  - Done


# Hvordan gi tilgang til sites.selected

MÅ HA Sites.FullControl.All som delegated (typ Nils som Sharepoint admin)

  https://graph.microsoft.com/v1.0/sites/b0a484ab-cfd4-411f-8b12-cab75ef0cfae/permissions

{
  "roles": ["write"],  // or ["read"] vi trenger write
  "grantedToIdentitiesV2": [
    {
      "application": {
        "id": "<app-client-id>",
        "displayName": "<app-name>"
      }
    }
  ]
}
