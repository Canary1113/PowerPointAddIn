# Release Checklist

Use this checklist before publishing a GitHub Release.

## Package

- Build the add-in in Visual Studio.
- Create the release archive from source and required assets only.
- Exclude `.vs`, `bin`, `obj`, `RuntimeTemp`, `TestResults`, and all `.pfx` files.
- Open the archive and confirm it does not contain `PowerPointAddIn_TemporaryKey.pfx`.

## Signing

- Never commit or upload `.pfx` private key files.
- Use a local or CI secret certificate for signed release builds.
- Rotate the certificate if a `.pfx` was ever uploaded publicly.

## GitHub Release Notes

Recommended text:

```text
PPT Assistant for PowerPoint

Features:
- Background Blur for selected pictures
- Transparent, White, and Black Glass effects
- Recover for generated Glass effects
- Nav Bar Animation
- Theme templates with Color customization

How to use:
Install the add-in, open PowerPoint, and use the PPT Assistant ribbon. Object-based commands require selecting the target object first.
```
