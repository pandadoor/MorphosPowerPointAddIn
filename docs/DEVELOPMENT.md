# Developing Morphos

This document is the practical map for building, testing, and extending the add-in.

## Stack

- C#
- .NET Framework 4.8
- VSTO / Microsoft PowerPoint interop
- WPF task pane UI
- Open XML SDK 3.0.2
- PowerShell and Python test harnesses

## Core workflows

### Build and load

```powershell
.\build-and-run.ps1
```

### Run targeted verification

```powershell
.\tests\Verify-FontReplacementTargets.ps1
.\tests\Verify-MissingFontReplacementTargets.ps1
.\tests\Verify-StaleFontReplacementTargetCache.ps1
.\tests\Verify-PowerPointInteractiveReplace.ps1
.\tests\Verify-ServiceFontReplace.ps1
```

## Repository layout

| Path | Responsibility |
| --- | --- |
| `ThisAddIn.cs` | PowerPoint host lifecycle, task-pane ownership, refresh orchestration |
| `Ribbon/` | Ribbon button and tab integration |
| `UI/` | Main task pane shell and layout behavior |
| `Dialogs/` | Replace Font and Replace Color dialogs |
| `ViewModels/` | Pane state, selection state, node trees, async flows |
| `Services/` | Scan, cache, replace, Open XML fallback/fast-path behavior |
| `Utilities/` | Interop helpers, font lookup, retry helpers, selection helpers |
| `Models/` | Scan and replacement DTOs |
| `packaging/` | Enterprise installer scripts, WiX source, and build automation |
| `tests/` | PowerShell and Python verification harnesses |

## Architecture notes

### Add-in host

`ThisAddIn.cs` owns:

- PowerPoint event wiring
- warm refresh scheduling
- task-pane recovery after mutations
- visibility/ribbon synchronization

### Presentation service

`Services/PowerPointPresentationService.cs` is the operational core. It handles:

- active-presentation scanning
- replacement target generation
- font replacement
- color replacement
- validation after mutation
- fallback between live COM edits and package-based edits

### Cache layer

`Services/FontScanSessionCache.cs` keeps presentation-scoped state such as:

- cached font results
- cached color results
- replacement target inventory
- theme metadata

The replacement target cache is versioned so old or invalid target shapes are rebuilt before being used.

### Performance helpers

- `OpenXmlScanService` scans package content faster than walking every COM shape repeatedly.
- `OpenXmlFontReplacer` and `OpenXmlColorReplacer` handle safe package mutations where appropriate.
- `ComFontAccessorCache` reduces repeated dynamic COM property access overhead.
- `OfficeBusyMessageFilter` and retry wrappers smooth transient Office busy/rejected-call behavior.

## Testing notes

### Live PowerPoint checks

The repo uses real PowerPoint automation where it matters:

- add-in load verification
- task-pane auto-scan
- replace-dialog opening
- live service replacement

### UI harness

`tests/morphos_ui_harness.py` drives the installed add-in from the user's point of view. It is used by the PowerShell verification wrappers for interactive checks.

### Replacement-target safety

The font replacement flow is protected by multiple checks:

- installed Windows font inventory
- theme token support
- cache sanitization/versioning
- live service verification scripts

## Common extension points

### Add a new deck analysis

Start in `Services/PowerPointPresentationService.cs`, then add corresponding view-model and UI representation only after the service output is stable.

### Change replace behavior

Update:

- `Services/PowerPointPresentationService.cs`
- `Utilities/FontReplacementTargetBuilder.cs`
- `Dialogs/ReplaceFontsDialog.xaml(.cs)` or `Dialogs/ReplaceColorsDialog.xaml(.cs)`
- related verification scripts in `tests/`

### Adjust task-pane layout

Main task-pane behavior lives in:

- `UI/FontsUserControl.xaml`
- `UI/FontsUserControl.xaml.cs`
- `UI/FontsTaskPaneHost.cs`

## Development expectations

- Keep the add-in usable in a narrow task pane.
- Prefer compact, signal-first UI over heavy copy.
- Preserve the PowerPoint-first workflow: scan, inspect, replace, verify.
- Verify changes against live PowerPoint automation before claiming the flow is fixed.
