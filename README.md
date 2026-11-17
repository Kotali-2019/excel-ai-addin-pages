# AI Assistant for Excel — build output

Generated files only. Do not edit by hand; they are produced by the release pipeline.

The source repository is private. This repository exists solely to serve the built add-in
through GitHub Pages.

## Use the add-in from this site

The add-in itself always loads from the site above; the only file you install locally is the
**manifest**, because desktop Excel reads add-in manifests from a file path.

### 1. Download the manifest

Open PowerShell and run:

```powershell
$dir = "$env:LOCALAPPDATA\excel-ai-addin"
New-Item -ItemType Directory -Force $dir | Out-Null
Invoke-WebRequest https://kotali-2019.github.io/excel-ai-addin-pages/manifest.xml -OutFile "$dir\manifest.xml"
```

Keep this folder. The registration points at the file, so deleting it removes the add-in from Excel.

### 2. Register it with Excel

```powershell
$x   = [xml](Get-Content "$dir\manifest.xml")
$id  = $x.SelectSingleNode('//*[local-name()="Id"]').InnerText.Trim()
$key = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name $id -PropertyType String -Value "$dir\manifest.xml" -Force | Out-Null
New-ItemProperty -Path $key -Name RefreshAddins -PropertyType DWord -Value 1 -Force | Out-Null
```

No Node.js, no build, no admin rights — this writes to `HKCU` only.

### 3. Restart Excel completely

Close **all** Excel windows and confirm `EXCEL.EXE` is gone in Task Manager (Office reads the add-in
list only at startup). Then open Excel and a workbook.

### 4. Open the pane

**Home** ribbon → group **AI Assistant** → **Open AI Assistant**.

Check that it loaded from this site: **Settings → Diagnostics & support** shows version `0.4.0`;
`=AIASSIST.AIPING()` in a cell returns `pong`.

### 5. Configure your own endpoint

The add-in ships with no server and no key. In **Settings**, enter the base URL of any
OpenAI-compatible endpoint (for example a local server at `http://localhost:11434/v1`), pick a model,
then press **Test connection** and **Save**. Settings stay in this browser's local storage; nothing is
sent anywhere until you press Send with a scope selected.

### Removing it

```powershell
Remove-ItemProperty 'HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer' -Name '3b7f9a8c-4c2e-4a66-9a52-3b0d41a37b46'
```

then restart Excel. To switch this machine back to another origin, download that origin's
`manifest.xml` over the same path and restart Excel — registering a different manifest replaces the
entry because the add-in id is the same.

## Notes

- GitHub Pages caches files for up to 10 minutes, so a freshly published build can take a few minutes
  to appear. Restarting Excel earlier will not force a refresh.
- Excel on the web does not use this registration: use **Insert → Add-ins → Manage My Add-ins →
  Upload** and pick the downloaded `manifest.xml`.
- License: MIT. See the source repository for documentation.
