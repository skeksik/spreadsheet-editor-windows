# 0003: Dependencies and licences

Status: accepted (2026-11-03)

## Decision
- Every NuGet package is declared once in `Directory.Packages.props` (central package management) with an exact version (`[x.y.z]`).
  Projects reference packages without a version.
- Accepted licences for anything the app ships: MIT, Apache-2.0, BSD-2-Clause, BSD-3-Clause, and Microsoft's licences for the .NET
  runtime, the Windows SDK and the Windows App SDK. Anything else (LGPL included) needs the owner's approval first.
- Adding a package, a tool or a GitHub Action, or bumping any version (the .NET SDK in `global.json` included) needs the owner's
  approval. Standing exception: Microsoft's own Windows platform packages (the Windows SDK, the Windows App SDK and its component
  packages, the Windows Community Toolkit) are approved in advance; the PR names each package, version and licence.
- GitHub Actions are pinned by commit SHA with the tag in a comment.
- The licence texts of shipped packages go into the published app's `licences\` folder once there is a release build (docs/roadmap.md).

## Packages at the start
| Package | Version | Licence | Why |
|---|---|---|---|
| Microsoft.WindowsAppSDK.WinUI, .InteractiveExperiences, .Foundation, .Base | 2.3.9, 2.1.9, 2.3.12, 2.0.4 (Windows App SDK 2.5.1) | Microsoft | WinUI 3 (ADR 0001) |
| DocumentFormat.OpenXml | 3.5.1 | MIT | `.xlsx` (ADR 0002) |
| MSTest.TestFramework, MSTest.TestAdapter | 4.5.2 | MIT | tests only |
| Microsoft.NET.Test.Sdk | 18.11.0 | MIT | tests only |
