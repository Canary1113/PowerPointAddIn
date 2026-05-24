# Release Checklist

Use this checklist before publishing a GitHub Release.

## Version

- Create or reuse a version tag, for example `v2026.05.24`.
- Use a release title such as `PPT Assistant 2026-05-24`.

## Build

- Build the add-in in Release configuration.
- Verify that `PowerPointAddIn/bin/Release` contains:
  - `PowerPointAddIn.vsto`
  - `PowerPointAddIn.dll`
  - `PowerPointAddIn.dll.manifest`
  - `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`
  - `Assets/ThemeStyles.pptx`

## Installer Package

- Create a package folder named like `PPTAssistantInstall-YYYYMMDD`.
- Include the Release output under `PowerPointAddIn/`.
- Include:
  - `Install-PPTAssistant.cmd`
  - `Install-PPTAssistant.ps1`
  - `Uninstall-PPTAssistant.cmd`
  - `Uninstall-PPTAssistant.ps1`
  - `README-INSTALL.txt`
- Compress that folder as `PPTAssistantInstall-YYYYMMDD.zip`.
- Exclude `.vs`, `bin`, `obj`, `RuntimeTemp`, `TestResults`, source-only files, and all `.pfx` files from the release asset.
- Open the zip and confirm it does not contain `PowerPointAddIn_TemporaryKey.pfx`.

## GitHub Release

- Open `https://github.com/Canary1113/PowerPointAddIn/releases/new`.
- Select the release tag.
- Upload the installable zip as the release asset.
- Publish the release.

Recommended release notes:

```text
PPT Assistant for PowerPoint

Install:
1. Download PPTAssistantInstall-YYYYMMDD.zip.
2. Unzip it.
3. Close PowerPoint.
4. Run Install-PPTAssistant.cmd.
5. Reopen PowerPoint and use the PPT Assistant ribbon.

Features:
- Background Blur for selected pictures
- Transparent, White, and Black Glass effects
- Recover for generated Glass effects
- Nav Bar Animation
- Theme templates with Color customization

Requirements:
- Microsoft PowerPoint for Windows
- .NET Framework 4.7.2 or newer
- Microsoft Visual Studio Tools for Office Runtime
```

## Signing

- Never commit or upload `.pfx` private key files.
- Use a local or CI secret certificate for signed release builds.
- Rotate the certificate if a `.pfx` was ever uploaded publicly.
