# Message Notes - GitHub Pages edition

Version 1.0.2.0. This package serves the sidebar online; it does not require a local web server, Node.js, or localhost certificates on user computers. It has not been published to your GitHub account or installed in your tenant by this package's author.

## Publish the website

1. Create a GitHub repository, for example `message-notes`. On GitHub Free, the usual Pages route requires a public repository. Private-repository Pages availability depends on your plan. The published website is generally public even when its repository is private. Obtain any required organizational approval first.
2. Extract this ZIP. Upload **the contents of `site/`** to the repository root, including the `assets` folder. Do not upload the ZIP, this entire package, the local manifest, local certificates, or private office data. The repository root must contain `index.html`, `taskpane.html`, `manifest.template.xml`, and the JavaScript/CSS files, with `assets/` beside them. Use GitHub's **Add file > Upload files**, drag the files and assets folder, and commit. A `.nojekyll` file is included for uploads that preserve hidden files; this plain-file site also works without it.
3. Open **Settings > Pages**. Set **Source: Deploy from a branch**, **Branch: main**, **Folder: /(root)**, then save. Use the branch actually containing the upload if it is not named main. Inspect GitHub's deployment status rather than assuming the website is already live.
4. Open the published HTTPS website address shown by GitHub Pages, normally `https://YOUR-USERNAME.github.io/message-notes/`. This is not a `github.com/...` repository URL or a `raw.githubusercontent.com/...` URL. Keep HTTPS enforced.
5. The home page checks its own files. Click **Download Outlook manifest**. This creates `Message-Notes-GitHub.xml` with your actual website address in every required place. No editing or build command is required. **Do not install `manifest.template.xml`.**

## Install into the mailbox

Installing into the mailbox through Outlook on the web or Exchange Online is not a second, separate desktop installation. A supported Outlook desktop client using that same mailbox should receive the add-in. Hosting and mailbox registration are separate: publishing the website does not update an already registered localhost manifest.

### Use the PowerShell method that already worked

In the same PowerShell session used for Exchange administration:

```powershell
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline
& 'C:\path\to\Message-Notes-GitHub\Install-Hosted.ps1'
```

Replace the example path with the extracted script path. Do not change machine-wide execution policy just to run this package. Follow your organization policy for downloaded scripts. The installer asks for the intended mailbox address and the full path to the XML downloaded from the website.

The script validates the file and checks the hosted sidebar before making changes. It targets only ID `7c6010d9-3b66-45f7-86ca-296e175b2b86` in the selected mailbox. If the old installation exists, type `REPLACE` to remove that registration and install the hosted version. It explicitly enables the add-in and prints its registration. It does not use `-OrganizationApp`, change tenant security settings, or edit email bodies.

This is not an atomic update. If removal succeeds but reinstallation fails, the script reports that state. Retain this package and the downloaded hosted XML. A copy of the original local manifest is included for manual rollback. The original localhost server would still be required for that old version.

The add-in ID and `messageNote` property name are unchanged; hosting migration is not designed to migrate or erase note data. Preserve a copy of important notes before changing installations. This script's Exchange commands have been checked against Microsoft documentation but have not been executed in your live tenant.

### Browser installation alternative

Sign in to Outlook on the web as the intended user. Open `https://aka.ms/olksideload`. In **My add-ins > Custom Addins**, remove the old Message Notes registration if present, then choose **Add a custom add-in > Add from File** and select the newly downloaded `Message-Notes-GitHub.xml`. Refresh Outlook on the web afterward. The old Add from URL option is not the current installation route, even though the application itself is hosted online.

If your organization deployed this through Integrated apps, have the administrator update that deployment instead of trying to remove it as a user. For an office rollout, deploy the final hosted manifest through Microsoft 365 admin center and assign a pilot group first.

## Installed but missing

Open an ordinary, unencrypted **received email in the exact mailbox where you installed it**. Check **Apps on the message** or **More apps on the message ribbon**. Do not only check the left navigation app list. This is a read-mode message add-in, not a compose-mode or standalone app. Shared/delegated mailbox support is not enabled in this manifest.

In a connected Exchange Online PowerShell session:

```powershell
$Mailbox = Read-Host 'Mailbox email address'
$AppId = '7c6010d9-3b66-45f7-86ca-296e175b2b86'
Get-App -Mailbox $Mailbox |
    Where-Object { $_.AppId -eq $AppId } |
    Format-List DisplayName, Enabled, AppVersion, AppId
```

If it is present and `Enabled` is false:

```powershell
Enable-App -Mailbox $Mailbox -Identity $AppId
```

Get-App -Mailbox does not provide a complete inventory of apps deployed from Integrated apps. For such deployments, check user assignments in the admin portal. If the mailbox registration is enabled but the button is absent on an ordinary received email, compare Outlook on the web and new Outlook rather than assuming a hosting error. A missing website does not necessarily hide an installed ribbon button.

Refresh web Outlook and fully close/reopen new Outlook after an installation change. The same mailbox must be active in both. Do not disable security policies or bypass certificate warnings to troubleshoot it.

## Pilot checks

1. On GitHub, deployment completes and the HTTPS site loads. The home-page asset checks pass and a manifest downloads.
2. The intended mailbox shows Message Notes version 1.0.2.0 enabled.
3. Open the sidebar from a normal received message in Outlook on the web; pin it.
4. Save a non-sensitive note and wait for Saved. Select another message, then return. Reopen the pane and confirm the saved text remains.
5. Open the **same email in the same mailbox** in new Outlook and confirm the same note. Test changing the note and clearing a sample note.
6. Test any needed message move, archive, copy, protected-message, or multi-device behavior before using important correspondence. Shared/delegated and mobile use are not included in this build.

The sidebar does not autosave. Unsaved drafts remain only while the pane is open. It has a roughly 2,200-character limit, subject to Outlook's 2,500-character serialized custom-property object limit.

## Security and privacy

Only static program files are published. In this code, actual notes are saved to mailbox-item custom properties through Office.js, not uploaded to GitHub or stored in a notes database. GitHub receives ordinary requests for website assets, and Microsoft Office.js loads from Microsoft's CDN. Simply visiting the public website does not grant mailbox access.

Do not upload emails, note exports, tokens, passwords, private certificates, or organization exports. Protect repository write access because changed JavaScript will execute with the add-in's permissions when a user opens it. Mailbox metadata is not a secret vault; do not store passwords in notes.

## Contents and testing

- `site/`: the only folder whose contents should be published. The hosted manifest uses a single current Outlook Message Read command surface and no separate command runtime.
- `Install-Hosted.ps1`: single-mailbox replacement/enablement helper; run locally in an Exchange Online PowerShell session.
- `manifest.original-local.xml`: original localhost manifest, not for GitHub deployment.
- `tests/`, `verification/`: local automated checks and their limits; not required on the website.
- `package.json`: optional development/test commands, not a runtime requirement.

## Primary references

- Microsoft sideloading and desktop availability: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing
- Outlook Apps button: https://support.microsoft.com/en-us/outlook/getstarted/use-add-ins-in-outlook
- Get-App: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-app
- Enable-App: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/enable-app
- New-App: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-app
- Remove-App: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/remove-app
- Microsoft per-message metadata: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in
- Microsoft manifest validation: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest
- GitHub Pages publishing: https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site
- GitHub Pages HTTPS: https://docs.github.com/en/pages/getting-started-with-github-pages/securing-your-github-pages-site-with-https
