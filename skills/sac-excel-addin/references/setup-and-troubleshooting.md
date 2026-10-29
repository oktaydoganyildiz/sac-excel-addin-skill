# Setup, permissions, security and troubleshooting

Verified against SAP Help (deliverables "Onboarding" `2b73edbc…`, "Security" `0ef628d4…`, "Troubleshooting & FAQ" `4554bec9…`) on 2026-10-29.

## Requirements

- **Clients:** Microsoft 365 for Windows or Mac, desktop or web. Office 2021 for Windows works with a reduced feature set because Microsoft no longer develops it. SAP recommends the Monthly Enterprise Channel. Mobile is not supported.
- **Browser:** latest Edge, Chrome or Safari. Third-party cookies must **not** be blocked. JavaScript, session cookies and pop-ups must be allowed.
- **Network:** the firewall or proxy must let `assets.sapanalytics.cloud`, `appsforoffice.microsoft.com` and `hcs.cloud.sap` through.
- **SAC:** you need an SAC license. Tenants can run on BTP Cloud Foundry or Neo. The tenant URL must follow the form `https://<tenant>.<domain>`.
- **Live connections (created in SAC):**
  - Datasphere: direct or tunnel connection, OAuth 2 with PKCE.
  - S/4HANA Cloud: OAuth 2 SAML Bearer.
  - HANA Cloud: user/password or SAML SSO.
  - S/4HANA and BW 7.50 SP14+ / BW/4HANA: tunnel connection. With user/password, tick *Save this credential for all users on this system*.
  - On Mac, the SAC and Datasphere tenants must be on the same domain.

## Deployment

- **Per user:** `Insert ▸ Office Add-ins ▸ Store` → "SAP Analytics Cloud for Excel". This works only for tenants on the standard domain `hcs.cloud.sap`.
- **Central deployment:** use M365 admin center `Settings ▸ Integrated apps`. It needs Exchange Online mailboxes. Users may have to restart Office. Centrally deployed users must not also load the add-in from the Store.

## Tenants on non-standard domains

Domains such as `analytics.cloud.sap`, `sapanalytics.cloud` and `hanacloudservices.cloud.sap` need a Microsoft 365 admin. That admin works on a Windows machine with the PowerShell module `O365CentralizedAddInDeployment`.

The admin runs, in this order:

1. `Connect-OrganizationAddInService` to sign in as the Microsoft 365 admin.
2. `Set-OrganizationAddInOverrides` with
   - ProductId `8a512e2b-c04e-4c06-9235-11b6cd59f584`;
   - an `-AppDomains` list that always contains `"https://hcs.cloud.sap"` plus every tenant URL, for example `"https://<tenant>.eu10.analytics.cloud.sap"`. Do not put a trailing slash on the URLs.

To roll back, run the same command with only `"https://hcs.cloud.sap"` as AppDomains.

Afterwards, every user who loaded the add-in **before** this change must clear the add-in cache:
- **Excel Web:** in the browser developer tools, clear the Local Storage entry of the Excel Online site.
- **Mac:** close Excel, then in Excel's container folder (`…/Microsoft/Office/16.0/Wef`) delete the entries matching the ProductId.
- **Windows:** `File ▸ Options ▸ Trust Center ▸ Trusted Add-in Catalogs` → tick *clear all previously-started web add-ins cache*, then restart Excel.

This is an admin change to the Microsoft 365 tenant. Explain the steps to the user; do not run them on their behalf. The exact procedure is on the SAP Help page *Tenant Configuration*.

## Trusted origins in SAC

Check them in `System ▸ Administration ▸ App Integration ▸ Trusted Origins`.

- Preconfigured: `https://hcs.cloud.sap`, `https://*.live.com`, `https://*.online.office.de`, `https://*.sharepoint.com`.
- **Add by hand:** `https://*.cloud.microsoft`. It is new and not preconfigured; see What's New 2026.13.

## Permissions

- The SAC object is **Add-in Workbook**, with Create / Read / Update permissions.
  - **Designer** (Create + Read + Update): full analysis and planning, and can save workbooks. Saving needs both Create **and** Update.
  - **Viewer** (Read):
    - opens workbooks;
    - changes the display, number format and variables;
    - runs planning triggers;
    - adds comments when the designer enables them.
    Planning is off by default for viewers.
- **Giving viewers more rights:**
  - filter dialog `More Options ▸ Allow viewers to modify selections`;
  - the table's *Planning Enabled* setting;
  - or Workbook Preferences `Workbook Permissions ▸ All users have designer rights`.
  Without that last setting, users with no permission cannot view the workbook.
- Workbooks created before 2023.25, and tables added from a Story, give designer rights to every user.

## Saving and sharing

- **Saving to SAC:**
  - The workbook becomes an *Add-in Workbook* object: `.xlsx`, or `.xlsm` when it has macros.
  - Macro workbooks open via Download. Excel for the Web disables them.
  - Excel autosave does **not** save back to SAC. A workbook opened from SAC is a local copy.
- **Workbook names:** at least 3 characters, unique, no leading or trailing space.
- **Sharing:**
  - Use Excel *Share* with **Allow editing**. Without it, the recipient sees no ribbon.
  - Do not use Excel co-authoring.
  - Use *Clear Data* before sharing. After *Refresh All*, each user sees only the data they are authorised for.
- **Tenant binding:** a workbook is tied to one tenant once the first data source is added. *Replace Tenant* moves it; every source must also exist on the new tenant.

## Security and data protection

- **How it runs:** an M365 web add-in in WebView2 (Windows) or WebKit2 (Mac). It calls the same backend APIs as SAC stories and Data Analyzer, over SAML/OAuth. Users need authorisations in both M365 and SAC.
- **Logging:** Read Access Logging and the Change Log are kept on the SAC tenant, not in Excel.
- **Your responsibility:** designers may type personal data into workbooks, and it is up to them to classify and protect it. Protect local files with device management and encryption.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Go!` button disabled | Wrong URL, or a trusted origin is missing (see above, including `https://*.cloud.microsoft`), or a non-standard domain is not configured. |
| Builder panel empty after sign-in | Third-party cookies are blocked, or a required domain is blocked by the firewall. |
| Empty browser window at sign-in | Tenant configuration missing or incomplete. Redo it and clear the cache for every user. |
| Mac + non-standard domain (Excel 16.44+) | Web: Safari `Preferences ▸ Privacy` → untick *Prevent cross-site tracking*. Desktop: sign-in not possible. |
| Ribbon buttons missing / add-in outdated | Insert the add-in again: `Insert ▸ Office Add-ins`. |
| Red refresh icon | Sign-in to the tenant failed. Refresh and sign in again. |
| "Missing named range" / corrupted objects | `SAC…` named ranges (`SACTable1`, `SACPlanningTrigger1`) were renamed or deleted, often by deleting a sheet offline. Restore or delete them in **Workbook panel ▸ Components** (2026.05). |
| Add-in crashes on a big table | Memory limit at about 0.5–1 M cells. Switch to high-volume mode. |
| Formula shows an error offline | Set the preference **Mode for Data Formulas ▸ Offline Mode**. GETDATA, SETDATA, GETDATACOMMENT and MEMBERPROPERTY then keep their last value. |
| SETDATA keeps reopening mass data entry | A high-precision delta from server-side rounding. Wrap the formula in `ROUND()`. |
| Comments lost after refresh | Edit comments only in the Comment tab. |
| Macro workbook opens with macros disabled | Use *Download*, unblock the file properties, then *Enable Content*; or add the tenant to Trusted Sites. VBA runs only in Excel desktop (Windows/Mac). |
| Data action trigger does not run | The target table is in mass data entry mode, or a parameter is invalid. Click the check icon in the trigger header. |
| "The last action could not be applied…" | The server rejected a state and the add-in rolled back (2026.19). Review the last input. |

## Keyboard shortcuts

On Windows press `Shift+Ctrl+Alt` plus the key; on Mac press `control+option+command` plus the key.

| Key | Command |
|---|---|
| F5 | Refresh All |
| F6 | Refresh Sheet |
| F10 | Refresh Sheet Functions |
| F8 | Refresh Selected Functions |
| D | Clear Data |
| F7 | Toggle automatic refresh |
| ↑ / ↓ | Open / close the designer panel |
| → / ← | Expand / collapse a node |
| Enter | Process Data |
| Space | Click a cell icon, e.g. the GETDIMENSIONFILTER pencil |

Exception: enabling or disabling a model for planning is `Ctrl+Alt+I` on Windows and `control+option+command+I` on Mac.

## SAP support

- Component **`LOD-ANA-OF-XLA`**.
- Include in the incident:
  - platform (Windows or Mac) and browser;
  - tenant URL and tenant version;
  - repro steps and screenshots;
  - whether the issue affects all users, all models and all workbooks.
- Make sure any shared workbook is not read-only.
- SAP may ask for SAC credentials. Enter them only in the incident's secure area, never in a chat or ticket text.
- A **preview version** of the add-in is available about 3 weeks before each release through SAP Note 3652483 (since 2026.01). Test only on copies of workbooks: the preview converts formulas to the `SAPPreview` namespace and cannot convert them back.
