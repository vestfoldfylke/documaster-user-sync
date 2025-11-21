# documaster-user-sync
Script som håndterer brukersynk i Documaster

## Flyt
- Hent alle rader i Sharepoint-liste (Documaster tilgang)
  - Hver rad representerer en gruppe i EntraID
  - Kolonner som kreves er `Title` og `Har tilgang` (multiple people picker)
- For hver rad/item:
  - Hent korrekt Entra gruppeid ved oppslag i GROUP_MAPPINGS fra env (bruker item.title som nøkkel, og får tilbake objectId for gruppe i Entra)
  - Hent alle eksisterende gruppemedlemmer fra EntraId
  - Lag en liste over membersToAdd (alle i "Har tilgang"-kolonnen som IKKE finnes i Entra-gruppen)
  - Lag en liste over membersToRemove (alle i Entra-gruppen som IKKE finnes i "Har tilgang"-kolonnen)

  For hvert element i membersToAdd
    - Hent brukerens object id basert på email (som vi får fra Sharepoint)
    - Legg til brukeren i gruppen
  
  For hvert element i membersToRemove (disse har vi id på allerede, fordi de kommer fra Entra-gruppen)
    - Fjern brukeren fra gruppen

  Logg ut resultatet


## Hvordan sette opp scriptet?
1. Klon ned eller deploy koden der det skal kjøres
  - `git clone <repo>`
  - `npm run build`
1. Opprett en .env fil på rot med følgende innhold
```bash
AZURE_TENANT_ID="<din-tenant-id>"
AZURE_CLIENT_ID="<client-id fra app registration>"
AZURE_CLIENT_SECRET="<client-secret fra app registration>"
# AZURE_LOG_LEVEL="verbose" # Hvis du vil ha logger fra azure/identity-pakka

DOCUMASTER_SOURCE_LIST_SITE_ID="<site-id til der documaster-tilgang lista ligger>"
DOCUMASTER_SOURCE_LIST_LIST_ID="<liste-id til documaster-tilgang>"

# GROUP MAPPINGS (SP_ROW_INTERNAL_NAME="sharepoint_title;entra_group_id")
SP_ROW_TEST_GROUP="Test-gruppe;ac5b6d4f-aa1e-464f-8d68-14a725c2f88e"
SP_ROW_ARKIV="Arkivtjenesten;<en guid>"
...
```
1. Lag en app registration i Azure, typ "Documaster-user-sync" ellerno
  - Legg til owner, og beskrivelse på både app og enterprise app
  - Skru på assignment required under properties på enterprise app
  - Legg til API Permissions:
    - Microsoft Graph
      - GroupMember.ReadWrite.All (Application)
      - Sites.Selected (Application)
      - User.ReadBasic.All (Application)
  - Gå til Certificates og secrets og lag deg en secret.

1. Gi app registration tilgang til Sharepoint siten der Documaster-tilgang listen ligger
  - Fordi vi bruker Sites.Selected API permission, [må vi også sette permission på sharepoint-siten](https://learn.microsoft.com/en-us/graph/api/site-post-permissions?view=graph-rest-1.0&tabs=http)
  - Finn en voksen som har rollen Sharepoint-admin (eller er eier av Sharepoint-siten) OG en app med mulighet for å kjøre delegated Sites.FullControl.All (f. eks Graph explorer)
  - Finn site-id ved f. eks å kjøre `GET https://<orgname>.sharepoint.com/sites/<sitename>/_api/site/id`
  - Be den voksne kjøre POST til følgende url, og med følgende payload:
  ```
    POST https://graph.microsoft.com/v1.0/sites/<site-id>/permissions

    {
      "roles": ["write"],
      "grantedToIdentities": [{
        "application": {
            "id": "<app registration id>",
            "displayName": "<app registration name>"
          }
      }]
    }

  ```
  - Få tilbake 200 ok smil (pass på å ikke bli amper)
1. Kjør røkla med `npm run start` eller `node --env-file=<sti til din .env> ./dist/index.js`




