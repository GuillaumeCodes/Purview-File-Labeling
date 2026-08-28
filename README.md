# Microsoft Purview File Labeling Utility

`Invoke-PurviewFileLabeling.ps1` is a guided PowerShell utility for reviewing and applying Microsoft Purview sensitivity labels to files. It supports two distinct sources: a file-system path (local, UNC, or SharePoint Server content exposed through a mounted path), or a native SharePoint Online document library. Dry run is the default. Apply mode requires an explicit selection and a second confirmation before any file is changed.

## Choosing a source

The first question of a guided run is where the files are. The two paths use different technology, so their requirements differ:

| | File path / SharePoint Server | SharePoint Online library |
| --- | --- | --- |
| Host | Windows PowerShell 5.1 or PowerShell 7 | PowerShell 7.2 or later |
| Module | `PurviewInformationProtection` | `PnP.PowerShell` |
| Read label | `Get-FileStatus` | `Get-PnPFileSensitivityLabel` |
| Apply label | `Set-FileLabel -PreserveFileDetails` | `Add-PnPFileSensitivityLabel`, which calls Graph `assignSensitivityLabel` |
| Sign-in | `Set-Authentication` | Delegated sign-in to survey; certificate-based confidential client to apply |
| Permissions | File-system access plus Rights Management rights | Delegated SharePoint `AllSites.Read` and Graph `Files.Read.All` to survey; application SharePoint `Sites.Read.All` and Graph `Files.ReadWrite.All` to apply |

These choices deliberately do not share a write implementation. The file-path choice writes the label into the file with the Purview Information Protection client. It covers local and UNC content, a OneDrive-synced library, and SharePoint Server content only when that content is exposed to the machine as a file-system path. It does not call Microsoft Graph, create an Azure billing resource, or incur a metered API charge.

The native SharePoint Online choice does not download files, change their labels, and upload them again. It calls [driveItem: assignSensitivityLabel](https://learn.microsoft.com/en-us/graph/api/driveitem-assignsensitivitylabel) through `Add-PnPFileSensitivityLabel`. Microsoft documents this as an advanced, protected, metered API. A successful call returns `202 Accepted`; the operation continues asynchronously, so the read-back verification and the `Not confirmed` outcome distinguish acceptance from completion.

A SharePoint Online library selected by URL is always handled by this native, metered path. The Purview client is used only for files available through a Windows file-system path.

SharePoint Online Apply is offered only when the current connection uses the configured certificate-based confidential client. That application requests SharePoint application permission `Sites.Read.All` and Graph application permission `Files.ReadWrite.All`. Successful writes also require administrator consent and an Azure billing association. A delegated public-client connection can still enumerate files, read labels, and run a dry run, but it cannot call the metered API.

### Cost and client-type limits on the SharePoint Online apply path

The constraints below come from Microsoft's [metered API overview](https://learn.microsoft.com/en-us/graph/metered-api-overview), [metered API list](https://learn.microsoft.com/en-us/graph/metered-api-list), and [setup procedure](https://learn.microsoft.com/en-us/graph/metered-api-setup):

- `assignSensitivityLabel` is **billed by API usage**. One Apply attempt is one metered call; consult Microsoft's metered API list for the current rate. Reading existing labels, enumerating files, and dry runs do not call it.
- The calling application **must be a confidential client**. Microsoft's wording is "web application, web API, or daemon/service", and "public client applications (desktop and mobile applications) aren't supported". A confidential client is one that can keep a certificate or client secret private. A workstation PowerShell session signing in as *you* cannot, which is why the default sign-in is refused; an application signing in with its own certificate can, which is why the setup below works.
- The application must be associated with an active Azure subscription by a `Microsoft.GraphServices/accounts` resource. The subscription must be in the application's tenant, and the operator needs effective Contributor or Owner access on the subscription plus direct ownership of the application registration.
- Metered APIs are available only in the Microsoft global environment. Public clients and Azure managed identities are not supported.
- SharePoint applies labels to [fewer file formats](https://learn.microsoft.com/en-us/purview/sensitivity-labels-sharepoint-onedrive-files#supported-file-types) than the Purview client does on disk: `.docx`, `.docm`, `.xlsx`, `.xlsm`, `.xlsb`, `.pptx`, `.ppsx`, and `.pdf` once the tenant enables PDF support. Legacy formats such as `.doc` and `.xls`, template formats, and Visio drawings are recognized on upload but cannot be labeled in place, so the utility offers this shorter set for the SharePoint source.

Reading labels, enumerating files, and dry runs do not invoke `assignSensitivityLabel` and therefore do not incur its per-call charge; applicable Microsoft 365 and Purview licensing still applies. Before a confidential client is configured, Apply is not offered for the SharePoint Online source. Once the current connection is app-only, Apply becomes selectable and each file that reaches the write step causes one metered API call. The run-mode check does not independently revalidate the billing resource, so a missing prerequisite is reported as a rejected file operation. Dry run remains the default and a second confirmation is still required.

### What "protected" means for this API

[assignSensitivityLabel](https://learn.microsoft.com/en-us/graph/api/driveitem-assignsensitivitylabel) is documented as a **protected API**, meaning it needs validation beyond permissions and admin consent. That validation is the metered-API setup itself: associating the application with an active Azure subscription by creating a `Microsoft.GraphServices/accounts` resource, as described in [Enable metered APIs and services in Microsoft Graph](https://learn.microsoft.com/en-us/graph/metered-api-setup). Nothing else is submitted and nothing is waited on. The prerequisites are an application registration, an Azure subscription, Contributor rights on it, and ownership of the application registration, and this utility automates the association.

There is no unmetered alternative **through Graph**. `assignSensitivityLabel` is the only Microsoft Graph API that writes a sensitivity label to a file in SharePoint or OneDrive, and it is available only in the global cloud.

There is, however, a route that avoids this metered API and its per-call charge altogether. The Purview Information Protection client writes labels to files on disk, and the OneDrive client syncs a SharePoint library to disk, so labeling a **synced copy** achieves the same result: OneDrive uploads each change straight back to SharePoint. When you choose the local source, the utility reads the OneDrive client's own list of mounted libraries and offers each one as a folder to scan, so there is no path to type and no guessing which folder maps to which library. When the SharePoint source refuses to apply, it offers to switch you to that route directly. A Purview auto-labeling policy is another service-side route that avoids this API's per-call charge. Existing product and compliance licensing still applies to both alternatives.

### Temporary artifacts are cleaned up

Certificates generated during the current run are tracked by thumbprint and registration time, and on exit the utility deletes each one **except** the credential the saved client still needs. Startup and exit also sweep older certificates whose subject contains `Purview File Labeling`, while preserving the saved client's thumbprint. PnP's exported copies are written to a temporary folder and deleted immediately. After a user-fixable setup failure, the utility offers to delete the application, remove its certificate, and forget the pending link. A server-side failure inside Microsoft.GraphServices is treated differently: the valid application, certificate, and exact pending link are retained without a rollback prompt. Registering another application would not change a server-side outcome, and keeping these values lets a later attempt target the same resource safely.

### Enabling the SharePoint Online apply path

Main menu option 3, *Enable SharePoint Online metered label writing*, performs the three steps Microsoft requires:

1. **Registers a confidential client.** It creates an Entra application whose credential is a certificate generated into your personal certificate store (`Cert:\CurrentUser\My`), where Windows protects the private key. PnP also writes exported copies of that certificate; those are directed to a temporary folder and **deleted immediately**, rather than being left in Downloads, because the store already holds the only copy that is needed. If another machine ever needs it, export it from the store deliberately. It requests Graph `Files.ReadWrite.All` and SharePoint `Sites.Read.All`, both application permissions, so a **Global Administrator or Privileged Role Administrator** must grant administrator consent. Only the client ID, tenant ID, and certificate thumbprint are remembered in `PURVIEW_FILE_LABELING_CC_*` environment variables; none is secret or can authenticate without the private key. The saved tenant ID is checked against the site's tenant on every run, so the application is only ever used in the tenant it belongs to.
2. **Grants administrator consent and assigns an owner**, from inside the utility. Registering creates the *application*; app-only sign-in additionally needs a **service principal** in the tenant, and Microsoft requires the Azure billing operator to own the application registration. When Azure CLI is available, the utility signs it in once to the application tenant and reuses that same account to add and read back the application owner, create or read the service principal, grant administrator consent, and read back every requested application permission. The same CLI session continues into subscription selection, prerequisite checks, and billing creation. No access token is printed, copied into an environment variable, written to disk, or passed on a command line; `az` obtains and attaches tokens from its own protected cache. Use one account that can grant consent and also has Contributor or Owner on the subscription. Permissions and ownership already present are treated as success, so this step is safe to repeat. If the CLI route cannot perform consent, the utility stops without opening another sign-in and offers the isolated `Microsoft.Graph.Authentication` route as an explicit fallback choice. If you decline, or both routes fail, the utility describes the portal steps instead. It deliberately does **not** link to the `/adminconsent` endpoint, which requires a registered `redirect_uri` that a certificate-only application does not have and would land on an error page.
3. **Links Azure billing.** It creates the `Microsoft.GraphServices/accounts` resource that associates the application with the Azure subscription to be charged, per [Microsoft's metered API setup](https://learn.microsoft.com/en-us/graph/metered-api-setup). The utility signs Azure CLI in when needed, lists subscriptions and resource groups as numbered menus, excludes resource groups whose provisioning state is not `Succeeded`, can create a resource group, rejects a subscription from another tenant, registers `Microsoft.GraphServices`, and verifies both Microsoft prerequisites on the operator: direct ownership of the application and effective Contributor or Owner access at subscription scope, including inherited and group assignments. It invokes Microsoft's documented Azure CLI create command first. The step is idempotent: it checks for the resource before creation and uses read-only verification after either a success or an error. If that GA request fails inside the provider itself, one provider-advertised preview recovery can be selected explicitly as described below. If Azure CLI is unavailable, an isolated Az PowerShell route remains available; its expanded resource read-back must pass the same success contract.

The preferred billing route is Azure CLI. After the utility collects and validates the resource group, billing-resource name, subscription ID, and application client ID, it runs Microsoft's documented command with those values:

```powershell
az graph-services account create --resource-group <resource-group> `
    --resource-name <billing-resource-name> --subscription <subscription-id> `
    --location global --app-id <application-client-id>
```

Before that call, the utility registers the `Microsoft.GraphServices` provider on that subscription with `--wait`, runs the non-mutating preflight, and then installs or upgrades the `graphservices` extension. Each of those must succeed before resource creation is attempted. The installed extension version is also recorded, and is noted as unavailable rather than treated as a failure when it cannot be read. The subscription is passed to every command instead of being selected, so an Azure CLI session that was already open keeps its own default.

If the documented create command reaches Microsoft.GraphServices and comes back with a server-side type-load failure, such as one naming `OpenTelemetry` and `TryCreateLogger`, the utility reports it as a server-side error rather than a subscription, tenant, resource-group, or authorization problem. When that subscription's provider metadata advertises Microsoft's published `2022-09-22-preview` contract, the utility offers one explicit recovery choice. Acceptance sends one direct `az resource create` PUT with that API version, the same subscription, resource group, resource name, and application ID, and `location` still set to `global`. It does not export a token, use a template deployment, or create another application, certificate, group, or billing-resource name.

The preview response is never trusted by exit code alone. The same strict read-back must report type `Microsoft.GraphServices/accounts`, location `Global`, the intended `properties.appId`, a nonempty `properties.billingPlanId`, and `properties.provisioningState` equal to `Succeeded`. If preview is not advertised, is declined, or also fails verification, no further mutation is sent. The utility retains the application, certificate, exact pending link, and full errors, and suppresses automatic retries of the same provider signature. It also makes one read-only Azure Activity Log query for the failed resource and records the ARM correlation and operation IDs. Because Activity Log ingestion can lag, when no event is visible yet the utility prints the exact read-only lookup command to run a few minutes later. Main-menu option 3 can reuse the retained values for a later interactive attempt, which is worth doing because a server-side failure may be transient.

### Billing preflight and verification

The metered setup owns one fail-fast prerequisite sequence. Its application-side checks run before subscription selection, and its Azure-side **non-mutating preflight** runs immediately before `az graph-services account create`. It stops before creation, and does not save a retry, when it finds any of these user-fixable conditions:

- The saved private-key certificate is no longer in `Cert:\CurrentUser\My`, the application no longer exists, or its requested application-permission shape is empty.
- The application service principal is missing, one or more requested application roles are not assigned, or the assignments cannot be verified through the tenant-pinned CLI account. Incomplete consent triggers one offer to repair it; an unverifiable response stops without attempting billing.
- The active Azure cloud is not the global `AzureCloud` environment, the subscription is not `Enabled`, or the subscription tenant differs from the Entra application tenant.
- Azure CLI confirms that the signed-in user is not an owner of the application, or has neither Contributor nor Owner at subscription scope after inherited and group assignments are included. If Azure CLI cannot query either directory relationship, the utility records that it is unverified and lets the documented create command remain authoritative.
- The selected resource group cannot be read or its provisioning state is not `Succeeded`, or `Microsoft.GraphServices` does not finish registering.

Whether the service will accept a given write cannot be proven in advance: Microsoft Graph exposes no query that reports it, and the answer can depend on the individual file. The utility therefore verifies the billing resource it created and reports a refused write as a failed file operation rather than as an Azure billing failure.

Azure subscription ownership and Entra application ownership are separate permissions. If the application has no owner, choose option 3, select the existing confidential client, and run **Grant administrator consent and ensure the billing operator is an app owner**. The utility repairs and verifies ownership without replacing the application or certificate. Alternatively, add the Azure CLI user under the application's **Owners** page, or have a current owner or authorized Entra administrator run the exact `az ad app owner add` command printed by preflight. Then re-link with the existing application, resource group, and billing-resource name; do not register another application.

The preflight deliberately does not call `az deployment group validate`; Microsoft does not require that call, and a provider-side validation defect must not prevent the documented create command from being attempted. After creation, the utility follows Microsoft's documented verification sequence with the selected subscription specified explicitly: `az resource list` locates the resource, then `az resource show` must report type `Microsoft.GraphServices/accounts`, location `Global`, the intended `properties.appId`, a nonempty `properties.billingPlanId`, and `properties.provisioningState` equal to `Succeeded`. A token issued before the association may still need to be refreshed.

A confidential client that is configured but not yet consented to does **not** break ordinary use. Before preferring it, the utility checks whether it is actually usable in that tenant; if it is not, it offers to grant consent there and then, and otherwise falls back to the read-only delegated sign-in so the run continues as a survey. Apply is offered only when the connection the run actually holds is app-only, not merely when an application is configured. That proves the authentication shape, but the billing resource, a token older than it, or the file itself can still cause individual write requests to be refused.

### Sign-ins are reused rather than repeated

Enabling the apply path spans clients with separate OAuth token caches, so the utility reuses authentication wherever Microsoft supports it:

- An open **SharePoint** connection is reused when the site and application match, proven with a live call rather than assumed.
- The PnP registration cmdlet creates and installs the certificate but exposes no supported access-token input, so its registration sign-in cannot be handed to Azure CLI or Microsoft.Graph. Passing a raw bearer token between clients would not create an Azure session and would unnecessarily expose a short-lived credential. This is the one unavoidable client boundary in a new registration.
- After registration, a matching **Azure CLI** user session is reused for Entra application ownership, service-principal creation, administrator consent, subscription discovery, RBAC and provider checks, and billing. A normal new setup therefore needs at most two interactive sign-ins: PnP registration, then one Azure CLI sign-in. Re-linking an existing application can need only the reused CLI session. `-DeviceLogin` applies to both sign-ins. Azure CLI 2.61 and later normally asks for a default subscription during login; the utility disables that selector inside its temporary profile because subscription selection happens later in its own validated menu.
- The **Microsoft.Graph** module is now a fallback for consent rather than a normal third sign-in. When needed, it runs in a separate PowerShell process because `PnP.PowerShell`, `Az.Resources`, and `Microsoft.Graph.Authentication` can load incompatible versions of the same .NET assemblies, and .NET cannot replace an assembly after it has been loaded. Az PowerShell remains isolated for the same reason. The Azure CLI route is already a separate executable, so the main process invokes it directly; every relevant resource command receives `--subscription` explicitly and does not change an existing CLI default subscription.
- The utility only reuses an Azure CLI account when it is a user session whose active tenant matches the application tenant. Otherwise it points `AZURE_CONFIG_DIR` at a new temporary directory and opens one tenant-pinned sign-in there. Every later `az` command inherits that isolated profile. Cleanup restores the previous `AZURE_CONFIG_DIR` value and deletes the temporary profile instead of calling `az logout`, so a pre-existing account, default subscription, and custom CLI profile remain untouched.
- Discovered **sensitivity labels** are cached for the session, so a second run does not re-query them.
- [New-LabelTestSite.ps1](New-LabelTestSite.ps1) hands over its site, library, and source. It also hands over an application that survives cleanup: a supplied application automatically, or a generated one when `-KeepApp` is used. A helper-generated application has delegated SharePoint `AllSites.FullControl` but not Graph `Files.Read.All`, so it cannot read existing file labels unless that Graph permission is added and consented; otherwise the labeling utility offers to register its own read-only application.

### Stale application IDs are removed, not proposed

A remembered application that the tenant no longer has is worse than none, because it is offered as the default. Before proposing one, the utility asks the tenant's device-authorization endpoint whether it is usable, which needs no sign-in and no consent. If the answer is no, it confirms against the directory itself whenever a Graph session is already open, and **clears the remembered value** when absence is confirmed. The one exception is an application registered moments earlier in the same run, which is kept because Entra replication has not caught up yet. Values owned by other tools are reported but never modified.

### How credentials and identifiers are handled

No secret is written to environment variables, logs, or retained certificate-export files. The one intentionally persistent secret is the confidential client's private key in the Windows certificate store:

- The **certificate private key** lives only in `Cert:\CurrentUser\My`, protected by Windows. Exported copies are deleted as soon as registration finishes.
- **No password, client secret, or token** is collected, stored, or logged. The utility never prompts for one.
- The values it remembers are a client ID, a tenant ID, a certificate thumbprint, a site URL, and a library name. These are identifiers rather than secrets, and none can authenticate on its own.
- Interactive user sign-in happens in Microsoft's browser or device-code flow, so no user credential passes through the script. App-only SharePoint sign-in uses the certificate in the Windows store.
- Run logs and the CSV report record site URLs, tenant IDs, and paths, which is why `.gitignore` keeps them out of the repository.

Afterwards the SharePoint source signs in as the application rather than as you, and the run-mode prompt offers Apply. Three limits are worth knowing before you start:

- Metered APIs are **unavailable in national clouds**, including GCC.
- A token issued *before* the billing link exists is still refused, so restart the utility after linking.
- Some **IRM-protected labels cannot be applied app-only** and return `Not Supported`, because SharePoint cannot validate user rights without a user.

Option 3 also lets you re-link billing for an existing application, grant its consent, replace it, or forget it and return to survey-only mode. Replacing never deletes the old application until the replacement has been registered and saved, so a failed registration cannot leave you pointing at an application that no longer exists. Forgetting or replacing one offers to **remove it from Entra ID** as well, since an orphaned application holding `Files.ReadWrite.All` is worth cleaning up: it strips every requested permission and redirect URI, disables sign-in on the service principal, then deletes the application. If your sign-in lacks Graph `Application.ReadWrite.All` the delete is refused, but the strip and disable usually still succeed, so what is left behind can no longer reach data; the exact `Connect-MgGraph` commands to finish the job by hand are printed either way.

You do not have to find option 3 first: when a SharePoint run is forced to dry run because no confidential client exists, it explains why and **offers to run the setup there and then**. That returns you to the main menu afterwards, because registering the application replaces the sign-in the run was using.

If you would rather not use the metered `assignSensitivityLabel` API:

- A **Microsoft Purview auto-labeling policy** for SharePoint and OneDrive, which labels content at rest under the applicable compliance licensing rather than this API's per-call meter.
- **Syncing the library with OneDrive** and running this utility's local source against the synced folder. That path uses the Purview Information Protection client and `Set-FileLabel`, avoids the metered Graph API altogether, and syncs the labeled files back. Applicable Microsoft 365, Purview, and OneDrive licensing still applies.

The SharePoint source needs PowerShell 7.2 or later, and the local source is most reliable in Windows PowerShell 5.1, so the utility does not force a host on you. Where the files are is asked **before** it signs in to anything, so if you pick SharePoint Online in Windows PowerShell it can offer to restart in the newest installed PowerShell 7 without discarding a Microsoft 365 sign-in. It closes its own connections first, hands over in the same window, and ends when that run ends. The restarted run **carries your source choice with it**, so the question is never asked twice. Answer no to stay where you are and use the local source. `-NoRelaunch` suppresses the offer, and the restarted run always carries it so the handover happens at most once.

Tenant label discovery, the priority rules, dry run, logging, and the CSV report are identical for both sources.

### The Entra application for the SharePoint Online source

You do not have to create one yourself. When you choose the SharePoint source, the utility offers to register an application named `PnP PowerShell - Purview File Labeling` with exactly two delegated permissions: SharePoint `AllSites.Read` to enumerate files, and Microsoft Graph `Files.Read.All` to read the label already on each file. Neither one can write. The tenant is taken from SharePoint's own authentication challenge, so the application is always single-tenant in the site's tenant, and no application ID is embedded in the script.

Both permissions are user-consentable, so an administrator is only needed if tenant policy restricts user consent. Accept the consent prompt the browser shows on first sign-in. To confirm afterwards which scopes the session actually obtained, run `Get-PnPAccessToken -Scopes` while connected.

The generated non-secret client ID is remembered per tenant, in `PURVIEW_FILE_LABELING_CLIENT_ID_<tenantid>`, so later runs against the same tenant simply offer it and runs against a different tenant never do. Registration is done through **Microsoft Graph directly**, which reports progress in the console; PnP's own registration command opens a sign-in window this utility cannot see and can appear to hang, so it is kept only as a fallback. Neither permission GUID is hard-coded: each is looked up from the resource that publishes it, so a renamed or reissued permission cannot silently produce a broken registration. The tenant itself is read from SharePoint's own authentication challenge before any sign-in, and it is passed to the sign-in explicitly, so an account cached from a previous tenant is never silently reused. You can also pick an existing application instead. A newly registered application is not visible immediately, so sign-in is retried while Entra replicates it.

This read-only application is what the SharePoint source uses by default. If you complete the confidential client setup from main menu option 3, that certificate-based application takes precedence instead **for the tenant it was registered in**, the sign-in becomes app-only, and this prompt no longer appears. Against any other tenant the utility says so and falls back to the delegated application described here. Forgetting the confidential client also returns you to it.

The application that [New-LabelTestSite.ps1](New-LabelTestSite.ps1) creates is not sufficient for label reading by default: it holds only SharePoint permissions and is deleted at the end of the provisioning run unless `-KeepApp` is used. Keeping it hands its client ID to this utility, but Graph `Files.Read.All` must still be added and consented before it can read existing file labels; otherwise let the labeling utility register its own read-only application. If helper cleanup is refused, its consent, permissions, redirect URIs, and sign-in are stripped first, so what remains is inert. A name collision is not fatal in either script: the helper registers under a timestamped name, and the labeling utility offers to reuse the existing application, register under a new name, or take an application ID you paste in.

Unlike the provisioning helper, this application is kept, because it is what you sign in with. Remove it in the Entra admin center when you no longer need the utility.

## Prerequisites

- Windows with Windows PowerShell 5.1 (`powershell.exe`) and .NET Framework 4.7.2 or later. PowerShell 7 (`pwsh`) also works; because the Purview client ships .NET Framework assemblies, the utility falls back to `Import-Module -UseWindowsPowerShell` when a native import fails. Use Windows PowerShell 5.1 if the module misbehaves. The SharePoint Online source needs PowerShell 7.2 or later.
- The Microsoft Purview Information Protection client and `PurviewInformationProtection` PowerShell module installed for the local and UNC source. The utility cannot install this client automatically because it is distributed as a client installer rather than a PowerShell Gallery module.
- `PnP.PowerShell`, for the SharePoint Online source.
- The `ExchangeOnlineManagement` module, used to connect to Security & Compliance PowerShell and retrieve labels with `Get-Label`.
- Only to apply labels in SharePoint: preferably the **Azure CLI**, which reuses one sign-in for consent, ownership, and billing. `Microsoft.Graph.Authentication` is needed only when its separate consent fallback is explicitly selected, and **`Az.Resources`** remains the billing fallback when Azure CLI is unavailable. None of them is needed to survey. That step also needs an Azure subscription in the same tenant, effective Contributor or Owner rights on it, direct ownership of the application, and a Global Administrator or Privileged Role Administrator to grant administrator consent.
- A Microsoft Entra account with a published sensitivity label policy that includes the labels being applied.
- Purview permissions that allow the admin to run `Get-Label` in Security & Compliance PowerShell.
- Read and write access to the target local folder, UNC share, or SharePoint library and its files.
- Rights Management permissions needed to inspect or replace protection. For protected content outside the operator's normal access, configure the account as a Purview Information Protection super user according to Microsoft guidance.
- An authenticated module session for the local source. Run `Import-Module PurviewInformationProtection` and `Set-Authentication` before an apply run. The utility verifies authentication with `Get-FileStatus` and offers to re-check; it does not collect credentials or secrets.

The required PowerShell Gallery modules do not have to be installed in advance. When one is missing at the point it is needed, the utility says so and offers to install it for your user account, which requires no administrator. Declining just prints the `Install-Module` command instead.

One upgrade needs a decision from you. PnP.PowerShell moved from a Microsoft signing certificate to the .NET Foundation, and PowerShellGet refuses to install across a publisher change: *"a Microsoft-signed module ... conflicts with the new module ... use -SkipPublisherCheck"*. The utility recognises that error, explains that the check exists to catch a package changing hands unexpectedly, and asks before retrying with `-SkipPublisherCheck`. The default is to stop.

Because a loaded assembly cannot be replaced in a running session, the utility **restarts itself** after updating PnP.PowerShell, in the same window, carrying your source choice with it. It refuses to do that twice in a row, so a repeatedly failing update cannot loop.

### More than one PnP.PowerShell version installed

PnP.PowerShell loads .NET assemblies, and .NET cannot unload them. Whichever version a PowerShell session touches **first** therefore wins for the life of that session, and `Import-Module -RequiredVersion` will not replace it. If an older copy is still installed, it can silently disable newer features such as registering a confidential client, while every version check appears to pass.

This is easy to miss, because `Get-Module -ListAvailable` reports only the **newest** version per module; older copies are invisible unless you add `-All`. Both scripts now enumerate with `-All`, warn when more than one version is installed, and `Invoke-PurviewFileLabeling.ps1` offers to uninstall the older ones. That offer only appears in a session that has not yet loaded PnP, since a loaded module cannot be removed.

`Uninstall-Module` only sees modules in the current host's module paths, so a copy installed for the other PowerShell edition, typically under `Documents\WindowsPowerShell\Modules` while you are running PowerShell 7, cannot be uninstalled that way. When that happens the utility recommends renaming the version folder so PowerShell stops discovering it; immediate deletion is the secondary choice. If the folder itself cannot be renamed, which is what happens once a session has loaded its assemblies, the utility renames the module manifest instead. PowerShell finds a module through its manifest, so that retires the version even while its files stay locked. It refuses any path that is not a `PnP.PowerShell` version folder, and it reports afterwards whether more than one version is still discoverable. To clean up by hand:

```powershell
Get-Module -ListAvailable -Name PnP.PowerShell -All | Select-Object Version, ModuleBase
Uninstall-Module PnP.PowerShell -RequiredVersion <old-version> -Force
```

Then open a **new** PowerShell window before running either script again.

Renaming does not delete the old DLLs, and PowerShell only discovers a module in a folder whose name parses as a version number. A folder renamed to `1.12.0.disabled-<timestamp>` is ignored by future PowerShell processes and can be deleted later. If the current process already loaded an assembly from that folder, it must be closed before the newer PnP version is used; no Windows sign-out or reboot is required. If renaming is denied, check the folder permissions or security software rather than assuming an editor lock.

Install and authentication guidance:

1. Install the Microsoft Purview Information Protection client from Microsoft Learn.
2. Install the label-discovery module with `Install-Module ExchangeOnlineManagement -Scope CurrentUser`.
3. Open Windows PowerShell 5.1 (or PowerShell 7) under the identity that will access the files.
4. Run `Import-Module PurviewInformationProtection`.
5. Run `Set-Authentication` and complete the Microsoft Entra sign-in.

The utility never asks for a password, application secret, label GUID, or tenant ID. Microsoft 365 authentication is completed in the browser.

## Label Discovery

The utility retrieves sensitivity labels directly from the customer's tenant. If the current PowerShell session does not already have `Get-Label`, it offers to:

1. Import `ExchangeOnlineManagement`.
2. Open an interactive browser sign-in with `Connect-IPPSSession`.
3. Run `Get-Label` and list enabled labels that can apply to files.
4. Show the friendly label name and tenant priority in a numbered menu. The GUID is not shown in the menu; the selected one is written to the run log.
5. Use the selected GUID automatically with `Set-FileLabel`.

Parent labels or label groups that contain child labels are excluded because they are not valid file-labeling choices. Sublabels are displayed as `Parent \ Child` and the child GUID is used.

Higher tenant priority numbers mean higher sensitivity. The script skips a file when its existing label has equal or higher tenant priority. It also skips an existing label that was not returned by discovery because its priority cannot be established safely.

## Optional validation-only helper

This repo also includes [New-LabelTestSite.ps1](New-LabelTestSite.ps1). It is a small SharePoint Online helper that creates a temporary test site, a document library, nested folders, and synthetic Office files so you can validate how Microsoft Purview sensitivity labels behave in a realistic structure.

It is intentionally separate from the labeling utility and is not required for normal use. If you already have a local folder or a managed SharePoint library to test, you can skip this script entirely. It is only meant to create disposable validation data and should not be treated as a production or operational labeling workflow.

How much content it creates is up to you. The defaults are deliberately small, so a run finishes quickly and a first labeling pass is easy to read: 2 top-level folders, 3 levels deep, and 1 subfolder inside each folder, which is 6 folders. Raise `-TopLevelFolders` (0-10), `-FolderDepth` (1-5), and `-SubfoldersPerFolder` (0-3) for a busier library, or answer the prompts in an interactive run. Folder names are drawn from a per-level pool such as `Finance/Reports/Q1`. Remember that every extra level multiplies the folder count by the number of subfolders, so 5 top-level folders, 5 levels, and 2 subfolders each is already 155 folders. Amounts that would create more than 250 folders or 1000 files are refused, and an interactive run asks again rather than ending.

`-FilesPerFolder` takes a range and defaults to `1-4`, so each folder gets its own randomly drawn number of files rather than a uniform count, and some folders can be left empty when the range starts at 0. A single number pins it: `-FilesPerFolder 3` puts exactly three files in every folder. The library root counts as a folder and gets files too. Files alternate between `.docx` and `.xlsx`.

### Test files that classifiers can detect

Every generated file holds real text rather than a placeholder line, and `-SensitiveFilesPerFolder` decides how many of each folder's files carry fabricated sensitive data. It also takes a range and defaults to `0-2`, drawn per folder and never more than that folder's file count, so the library ends up with a natural mix of clean and sensitive folders:

- `SensitiveDoc-*` files contain invented customer records in the spirit of the public sample sets at [dlptest.com](https://dlptest.com/sample-data/): names, US addresses, dates of birth, Social Security numbers, credit card numbers with expiration dates and verification values, bank account and ABA routing numbers, an IBAN, a passport number, a California driver's license number, and an ITIN.
- `TestDoc-*` files hold plain business text with nothing to match, which is what makes them useful as a control group.

A [sensitive information type](https://learn.microsoft.com/en-us/purview/sit-sensitive-information-type-learn-about) normally needs a primary element **and** supporting evidence within about 300 characters, so each value is written next to a keyword from that type's documented list, such as `Social Security Number (SSN): ...`, `Credit card number (Visa): ...`, `Card expiration date: ...`, `Card verification value (CVV2): ...`, `ABA routing number: ...`, and `US passport number: ...`. Values that carry a checksum pass it: the card numbers pass the Luhn test, the routing number passes the ABA check, and the IBAN passes its mod-97 check. The card numbers are deliberately **not** the reserved test numbers such as `4111 1111 1111 1111`, because the credit card type ignores those by design. Nothing in these files refers to a real person, account, or card, and the phone numbers use the reserved `555-01xx` range.

Classification is not immediate. Allow the service time to crawl the new library before expecting auto-labeling policies, DLP policy matches, or content explorer to show results.

### Handing straight over to the labeling utility

When provisioning finishes, the helper places the site URL and library name in process-scoped `PURVIEW_FILE_LABELING_SITE_URL` and `PURVIEW_FILE_LABELING_LIBRARY` values. If it starts `Invoke-PurviewFileLabeling.ps1` immediately, that child run proposes both as defaults. They are not saved to the user environment and disappear with the console process.

If `Invoke-PurviewFileLabeling.ps1` sits in the same folder, the helper also offers to start it. That launch is deliberately deferred until **after** helper cleanup finishes, so a generated application is removed before the next sign-in unless `-KeepApp` was used. Supplied applications are never removed. Answer no, or pass `-AcceptDefaults`, and nothing is started.

### Working with more than one tenant

Both scripts support multiple tenants and prevent a tenant-specific application from being silently reused in another tenant:

- The tenant is resolved from the SharePoint address itself, before any sign-in, and passed explicitly to the sign-in. Without that, MSAL reuses whichever account was cached last, which produces a confusing "user account does not exist in this tenant" error when you switch.
- Application IDs are remembered under a per-tenant name, so an application registered in one tenant is never proposed for another. A value found in a tenant-agnostic variable such as `AZURE_CLIENT_ID` is still offered, but labeled as possibly belonging to another tenant, and registering a fresh one becomes the default.
- The confidential client records the tenant it was registered in and is skipped, with an explanation, against any other tenant.
- `New-LabelTestSite.ps1` requires the tenant on every standalone run, either interactively or through `-TenantRootUrl`. It never proposes a tenant from an earlier run or an existing PnP session.

Nothing tenant-specific is stored in the repository itself. Site URLs and library names are process-only; tenant-scoped application IDs and the confidential-client configuration remain user-scoped so reusable applications are not abandoned or offered across tenants. Main-menu option 5 discovers and clears every value owned by either script. Both scripts also remove legacy user-scoped tenant, site, and library defaults created by earlier versions. To perform the same all-tenant cleanup manually:

```powershell
foreach ($target in [EnvironmentVariableTarget]::Process, [EnvironmentVariableTarget]::User) {
    [Environment]::GetEnvironmentVariables($target).Keys |
        Where-Object { $_ -match '^(PURVIEW_FILE_LABELING_|LABEL_TEST_SITE_)' } |
        ForEach-Object { [Environment]::SetEnvironmentVariable($_, $null, $target) }
}
```

### Helper behavior and authentication

The helper focuses on validation, not production labeling. It creates a disposable SharePoint site and library so that the label experience can be tested without affecting live content. The supplied tenant is validated and its admin URL, cloud, and tenant ID are discovered automatically, but the tenant URL is not persisted or inferred from a previous connection.

Every prompt with a sensible default proposes it in brackets, and values can be supplied up front with `-LogFolder`, `-SiteName`, `-LibraryName`, `-TopLevelFolders`, `-FolderDepth`, `-SubfoldersPerFolder`, `-FilesPerFolder`, and `-SensitiveFilesPerFolder`. Pass `-AcceptDefaults` to take every proposed value without being asked.

The helper writes every action to a timestamped `New-LabelTestSite-*.log` in the log folder, which defaults to the script directory. The log records the host and Windows user, each resolved value, each script prompt and answer, every site, folder, and file created, every warning and error, and the outcome of application cleanup. Lines recorded before the log file exists are buffered and flushed into it, so the log is complete even when the folder is chosen interactively. If the file cannot be created, the helper reports that the run is logged to the console only; otherwise, the log path is printed again on exit.

The normal interactive path can ask for these script settings; installation, authentication, and recovery can add their own prompts:

- The log folder, skipped with `-LogFolder` or `-AcceptDefaults`.
- The SharePoint root URL. `-TenantRootUrl` supplies it outright; otherwise every standalone run asks for it. The value can be a full SharePoint URL, an admin or OneDrive host, a `<tenant>.onmicrosoft.com` domain, or just the tenant alias; the scheme is optional.
- The application-registration decision, skipped when a usable application is found or when `-RegisterApp` or `-ClientId` supplies the decision.
- The site name and document library name. `-SiteName` and `-LibraryName` set their proposed values; `-AcceptDefaults` accepts them without prompting.
- How much test content to create: top-level folders, folder levels, subfolders per folder, files per folder, and how many of those files carry fabricated sensitive data. `-TopLevelFolders`, `-FolderDepth`, `-SubfoldersPerFolder`, `-FilesPerFolder`, and `-SensitiveFilesPerFolder` set their proposed values; the two file settings accept a number or a range such as `1-4`; `-AcceptDefaults` accepts them without prompting.
- Whether to start `Invoke-PurviewFileLabeling.ps1` after cleanup. `-AcceptDefaults` answers no.

When the host cannot prompt, the helper takes every default and stops immediately if a value it cannot infer is missing, naming the parameter to supply instead of waiting for input. Transient failures such as timeouts, HTTP 429 throttling, and 5xx responses are retried with exponential backoff and honor the server's `Retry-After` header; if the generated site URL is already in use, the next available name is used automatically.

A tenant's SharePoint address does not always match its `onmicrosoft.com` prefix, and some tenants have no SharePoint provisioned. When the address is derived rather than typed, the helper confirms the host exists in DNS before contacting it, reports whether the Microsoft 365 tenant itself exists, and asks again instead of retrying a name that cannot exist. The check is skipped on networks that resolve names only through a proxy.

PnP PowerShell requires an Entra application for interactive authentication. The helper does not contain or fall back to a shared application ID. Tenant and application IDs are always discovered, supplied, or generated at runtime; applications created by the helper are single-tenant registrations in the resolved tenant.

Before opening a SharePoint sign-in, the helper automatically:

1. Validates and normalizes the SharePoint Online root URL and derives the matching admin URL.
2. Reads the anonymous SharePoint Bearer challenge to discover the tenant realm and login authority.
3. Resolves OpenID metadata, selects the matching PnP cloud environment, and rejects a tenant hint or authority that points outside the SharePoint tenant's cloud.
4. Finds a client ID from `-ClientId`, a tenant-specific saved environment variable, `ENTRAID_APP_ID`, `ENTRAID_CLIENT_ID`, or `AZURE_CLIENT_ID`.
5. Calls the resolved tenant's device-authorization endpoint to verify that the app is available to that tenant as a public client and can request SharePoint access. This catches `AADSTS700016` before a browser window opens; supplied multi-tenant apps can also pass this availability check.
6. After sign-in, verifies the connected admin host, client ID when PnP exposes it, tenant ID when available, and access to the SharePoint tenant-admin API.
7. Disconnects the current PnP session before replacing it and again on every exit path. It does not clear persisted login caches or disconnect unrelated Graph, Exchange, or Purview sessions.

When no usable application is found, the helper offers to run `Register-PnPEntraIDAppForInteractiveLogin` (or its legacy PnP name) and requests only delegated SharePoint `AllSites.FullControl`. Automatic registration requires PowerShell 7.2 or later, a compatible PnP.PowerShell version, and an account allowed to create app registrations; tenant policy may require an administrator to grant consent. On Windows PowerShell 5.1, pass an existing compatible app with `-ClientId`. The generated non-secret client ID is saved in a tenant-specific user environment variable for later runs. Use `-RegisterApp` to choose automatic registration without the menu.

### Removing the application the helper creates

Unless `-KeepApp` is used, an application the helper registered is removed again on every exit path, including failed runs. An application supplied with `-ClientId`, or found in an environment variable, is never touched. When `-KeepApp` is used, the helper logs the generated application's name and client ID so it is not forgotten.

Deleting an app registration needs Microsoft Graph `Application.ReadWrite.All`, which the helper deliberately does not request for the app it creates. Cleanup therefore tries, in order:

1. `Remove-PnPEntraIDApp` on the existing PnP session, which succeeds only when that session already has Graph rights.
2. A **separate PowerShell process** that signs in with `Connect-MgGraph -Scopes Application.ReadWrite.All` and disconnects immediately afterwards. A child process is required because PnP.PowerShell and Microsoft.Graph load incompatible versions of `Microsoft.Identity.Client`, so an in-process Graph sign-in fails once PnP has been used. This step needs `Microsoft.Graph.Authentication`, which the helper does not install. On this path the application's access is stripped **before** the delete is attempted: every delegated consent grant is revoked, the service principal's sign-in is disabled, and all requested API permissions and redirect URIs are removed.
3. If neither works, the helper logs exactly what was left behind: the display name, client ID, tenant ID, whether the access strip succeeded, and the commands to finish the removal by hand.

When the application is deleted, or when it survives but its access was stripped, the saved client ID is cleared so no later run reuses it. When nothing could be changed, the saved client ID is deliberately kept, so a later run reuses that one application instead of leaving another orphan behind.

The helper selects a PnP.PowerShell version the host can load: `1.12.0` on Windows PowerShell 5.1, `2.12.0` on PowerShell 7.2 and 7.3, and the current release on PowerShell 7.4 or later. It recognizes the commercial, US Government (`sharepoint.us`), German, and 21Vianet (`sharepoint.cn`) SharePoint hosts, and stops when the SharePoint host and the resolved sign-in authority are not in the same cloud.

Only PowerShell 7.2 and later can register an Entra application, because the PnP release for Windows PowerShell has no registration command. When started from Windows PowerShell, the helper finds the newest installed PowerShell 7, restarts itself there with the same parameters, and returns that run's exit code. Pass `-NoRelaunch` to stay in the current host. If PowerShell 7 is not installed, the run continues and explains that `-ClientId` is required.

Run the full preflight without creating a SharePoint site:

```powershell
.\New-LabelTestSite.ps1 -TenantRootUrl 'https://contoso.sharepoint.com' -PreflightOnly
```

Allow the helper to register the tenant app if needed, then run the same checks:

```powershell
.\New-LabelTestSite.ps1 -TenantRootUrl 'https://contoso.sharepoint.com' -RegisterApp -PreflightOnly
```

After preflight succeeds, omit `-PreflightOnly`. A site name is generated automatically unless `-SiteName` is supplied. Device-code authentication remains available with `-AuthMode DeviceCode`; browser authentication is the default.

A run with every script setting accepted up front uses:

```powershell
.\New-LabelTestSite.ps1 -TenantRootUrl 'https://contoso.sharepoint.com' -RegisterApp -AcceptDefaults
```

This suppresses the helper's setting prompts; Microsoft authentication can still open a browser or device-code flow.

Keep the generated application, write the log elsewhere, and name the site and library explicitly:

```powershell
.\New-LabelTestSite.ps1 -TenantRootUrl 'https://contoso.sharepoint.com' -RegisterApp -KeepApp `
    -SiteName 'Label Test March' -LibraryName 'Label Test Library' -LogFolder 'C:\Logs'
```

Create a larger structure, here 4 top-level folders, 4 levels deep, 2 subfolders inside each folder, and 2 to 6 files per folder of which 1 to 3 carry fabricated sensitive data:

```powershell
.\New-LabelTestSite.ps1 -TenantRootUrl 'https://contoso.sharepoint.com' -RegisterApp -AcceptDefaults `
    -TopLevelFolders 4 -FolderDepth 4 -SubfoldersPerFolder 2 -FilesPerFolder '2-6' -SensitiveFilesPerFolder '1-3'
```

The helper's generated application has only delegated SharePoint access, so it cannot read file labels in `Invoke-PurviewFileLabeling.ps1` without Graph `Files.Read.All`. It is deleted when provisioning ends unless `-KeepApp` is used. A kept or supplied application is handed over, but the Graph permission must be added and consented before it is sufficient; otherwise the labeling utility offers to register its own read-only application.

## Run

From Windows PowerShell 5.1 or PowerShell 7:

```powershell
Set-Location "C:\path\to\the\utility"
.\Invoke-PurviewFileLabeling.ps1
```

Use Windows PowerShell 5.1 for local and UNC folders, because the Purview Information Protection client is most reliable there. For SharePoint Online either start it with `pwsh`, or start it anywhere and accept the restart it offers when you choose that source.

### If the browser insists on a passkey

The embedded browser sometimes offers only a passkey, and on a device enrolled with an Android Work Profile it can fail outright with *"We couldn't sign you in... please use the camera app in that profile"*. That is the sign-in page choosing a credential for you, not a fault in the utility. Two ways past it:

```powershell
.\Invoke-PurviewFileLabeling.ps1 -DeviceLogin
```

`-DeviceLogin` prints a short code and a URL to open in any normal browser, where **Microsoft Authenticator** can be chosen like any other method. The same choice is offered automatically after a failed sign-in, together with an option to retry in the browser forcing a fresh account picker, which discards the cached credential choice. A failed sign-in keeps the site you already entered, so only the sign-in is retried, and choosing a different application does not make you retype the URL.

You are not asked to sign in repeatedly either. Before authenticating, the utility checks whether the connection it already holds serves the same site with the same application, and proves it with a live call rather than trusting a stale object. If it does, that sign-in is reused and no browser window opens at all, which matters when a window can appear behind the console and be missed. Forcing a fresh sign-in always bypasses the reuse.

`New-LabelTestSite.ps1` behaves the same way: a failed sign-in no longer ends the run. It explains the failure, naming the tenant when the application belongs to a different one, and offers to retry with device code, forget the remembered application and register a fresh one for this tenant, enter a different client ID, or stop. `-AuthMode DeviceCode` selects device code from the start.

### Run artifacts may contain tenant information

The labeling CSV records file paths and label outcomes. Run logs can additionally contain site URLs, tenant IDs, and, for the helper, the Windows user and machine name. Both scripts default their artifacts **next to the script**, which is this repository. The `.gitignore` patterns for `Invoke-PurviewFileLabeling-*.log`, `Invoke-PurviewFileLabeling-*.csv`, and `New-LabelTestSite-*.log` reduce the risk of an accidental commit, but do not replace normal review. Choose an output folder outside the repository if you would rather keep the artifacts elsewhere.

No positional parameters are required. **Where the files are is asked once, before the main menu**, so an unreachable source costs no sign-in and the question is not repeated for every run. The menu shows the current choice and option 4 changes it. The guided session then asks, in this order:

- Microsoft 365 interactive sign-in when tenant label discovery is needed.
- The target folder, or for SharePoint Online the site URL, the sign-in application, and then the document library picked from a numbered list of the libraries that site actually has, plus an optional subfolder.
- Whether to include subfolders.
- Which files to scan. The offered sets match what the chosen source can actually label: for a file path, all Office documents, Visio drawings, and PDF including legacy and macro-enabled formats; for SharePoint Online, the shorter list that service supports (`.docx .docm .xlsx .xlsm .xlsb .pptx .ppsx .pdf`). Both sources also offer the current formats only (`.docx .xlsx .pptx .pdf`), or a list you type. A typed list that includes formats the source cannot label is accepted but reported, because on SharePoint each refused write is still a metered call.
- Target sensitivity label.
- Dry run or apply mode.
- Log and CSV report folder, defaulting to the script directory.

Every prompt that has a sensible default proposes it in brackets, so pressing Enter accepts it. You never have to know a library's internal URL: the utility signs in first and lists the libraries for you. A site URL that signs in successfully is reused only within the current process, including by confidential-client setup later in the same session. A later standalone run asks again, preventing a previous tenant from appearing as the default.

This is an interactive utility, so it is meant to be run in a console. If its input is redirected and runs out, it stops with a clear message instead of driving every menu at its default.

Start with a representative dry run and inspect the CSV report. Back up important data and test label behavior, encryption, access, and application compatibility before a production apply run.

## Main menu

After source selection, the main menu controls the session. Guided and batch runs survey by default and apply only after Apply is selected and confirmed. Most completed actions return to this menu; a clean-exit action ends the session.

| Option | What it does |
| --- | --- |
| 1. Guided run | Surveys or applies one label to the matching files in one folder or library, optionally including subfolders. |
| 2. Batch run from a CSV | Surveys or applies a different label to the matching files in each listed folder. See below. |
| 3. Enable SharePoint Online metered label writing | Registers a confidential client, grants its administrator consent, and links Azure billing so the SharePoint Online source can apply labels. Also re-links billing, grants consent for an existing application, replaces it, or forgets it. |
| 4. Change where the files are | Switches between the local/UNC/SharePoint Server file-path source and native SharePoint Online without restarting. |
| 5. Forget settings remembered from previous runs | Lists every value this utility remembers, with its scope, and clears them on confirmation. Also signs out of Microsoft Graph, which is otherwise kept so later runs do not prompt. Variables owned by other tools are listed but never modified. |
| 6. Show current session summary | Totals so far for this session: scanned, labeled, unconfirmed, skipped, failed, and elapsed time. |
| 7. Exit cleanly | Exports the report, prints the summary, and closes every session the utility opened. |

To ignore remembered values for a single run **without deleting anything**, start it with `-Fresh`:

```powershell
.\Invoke-PurviewFileLabeling.ps1 -Fresh
```

That suppresses every remembered application ID, site URL, library, and confidential client, so nothing from an earlier run or a different tenant is proposed as a default. The startup banner says so, and the setting carries across the PowerShell 7 relaunch.

Option 3 enables metered writing but does not label files itself. Per-file `assignSensitivityLabel` charges occur only when Apply is selected in a guided or batch run. The setup is described in full under [Enabling the SharePoint Online apply path](#enabling-the-sharepoint-online-apply-path). Until a usable confidential client is configured, the SharePoint Online source remains survey-only; applicable product licensing still applies.

## Batch runs from a CSV

A guided run applies one label to the contents of one folder or library. Here, *contents* means the matching files selected by the extension filter, optionally including subfolders; the folder object itself is not labeled. When different folders' contents need different labels, the main menu's batch option takes a CSV instead, so one sign-in covers the whole set:

```csv
Folder,Label,Recurse,Extensions
HR/Policies,Confidential,true,".docx,.pdf"
Public,General \ Anyone (unrestricted),false,
Archive/2019,Highly Confidential,true,
,Confidential,false,
```

Reading those four rows in order: `HR/Policies` and everything beneath it gets `Confidential`, but only Word documents and PDFs are touched. `Public` gets a parented label and is not recursed, so its subfolders are left alone. `Archive/2019` and everything beneath it gets `Highly Confidential` using the extension set you chose at the prompt, because its `Extensions` cell is empty. The last row has an empty `Folder`, which means the root itself, so it labels the loose files sitting directly in the root without descending into any of the folders above.

- `Folder` is relative to the library or root folder you pick after choosing the source. Leave it empty, or use `.`, to mean the root itself. Forward and backward slashes both work.
- `Label` accepts the display name shown in the label menu, the child name on its own for a parented label such as `Anyone (unrestricted)`, or the label GUID.
- `Recurse` is optional and defaults to `true`. It accepts `true`/`false`, `yes`/`no`, or `1`/`0`.
- `Extensions` is optional and falls back to the extension set you choose before the file is read.

The whole plan is validated before anything runs. A row naming a label that does not exist in the tenant, or an unreadable `Recurse` value, is dropped and reported, and the remaining rows still run. If nothing usable is left, the file is rejected. The resolved plan is printed for review, and apply mode still requires its own confirmation. All folders share one log and one CSV report.

Rows are processed in the order they appear, and overlapping folders are not detected. A recursive row combined with a row for one of its own subfolders therefore visits those files twice and reports them twice. In apply mode the priority rule still prevents the second visit from lowering a label that the first one raised.

## Outputs and Safety

Each guided or batch operation creates timestamped `.log` and `.csv` files after its output folder is selected. Merely opening the utility does not create them, and option 3 creates no standalone log when it is used before any guided or batch operation. The CSV records every processed file, previous label, requested new label, outcome, and details. The closing summary reports totals scanned, labeled, not confirmed, skipped, failed, and elapsed time.

For the file-path source, the utility uses `Get-FileStatus` per file and calls `Set-FileLabel -PreserveFileDetails` only after apply mode and explicit confirmation. For SharePoint Online it reads labels with `Get-PnPFileSensitivityLabel`; Apply calls `Add-PnPFileSensitivityLabel`, which submits the metered `assignSensitivityLabel` request. Note the near-identical `Get-PnPFileSensitivityLabelInfo`: despite the more descriptive name it is a SharePoint tenant-admin CSOM call that fails with `Attempted to perform an unauthorized operation` for anyone who is not a SharePoint Administrator, so this utility deliberately does not use it. Failures are collected and reported instead of terminating the whole run. Operator menus allow retry, changed input, skip, main menu, or clean exit. Every session it opens, to Security &amp; Compliance PowerShell or to SharePoint, is closed on exit.

Startup environment and dependency checks occur before an output folder is chosen, so their messages are not retroactively written to a later run log. When a guided or batch operation has initialized the log, a subsequent Azure CLI billing attempt records the executable and version, `graphservices` extension version, active cloud, subscription state and tenant, application and application tenant, resource-group location and provisioning state, provider registration state, advertised `Microsoft.GraphServices/accounts` API versions, authorization-check outcomes, full Azure errors, and tracking or correlation identifiers when available. A server failure also triggers one read-only Activity Log query for the exact billing resource; matching event timestamps, operation IDs, correlation IDs, status, and substatus are recorded without repeating the create request. These detailed billing diagnostics are log-only so the console stays readable. No password, token, client secret, or certificate material is logged.

PowerShell execution policy may block local scripts. Follow the customer's approved policy; do not weaken organization-wide policy solely to run this utility.

## As-Is Disclaimer

This sample is provided **as is**, without warranty of any kind. The customer is responsible for validating discovered labels and priorities, permissions, retention and encryption effects, backups, change controls, and compliance requirements before production use. Applying a sensitivity label can encrypt content and change who can open it. Microsoft and the script author are not responsible for data loss, loss of access, or unintended policy effects arising from its use.
