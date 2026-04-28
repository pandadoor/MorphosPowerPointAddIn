This guide covers how to install Morphos, either from a pre-built release or from source.

## Enterprise Release Installation (Recommended)

For most users, installing from a pre-built release is the most efficient method.

1.  Download the latest `Morphos_vX.X.X.zip` from the [Releases](https://github.com/pandadoor/MorphosPowerPointAddIn/releases) page.
2.  Extract the contents to a local folder.
3.  Right-click `install.ps1` and select **Run with PowerShell**.
    *   To install for all users, run from an elevated (Administrator) PowerShell prompt: `.\install.ps1 -AllUsers`
4.  Launch PowerPoint and locate the **Morphos** tab.

## Local Source Installation

## Requirements

- Windows 10 or Windows 11
- Microsoft PowerPoint desktop
- .NET Framework 4.8
- Microsoft Visual Studio 2022 or Visual Studio 2022 Build Tools with MSBuild
- Visual Studio Tools for Office runtime

For the current project configuration, x64 PowerPoint is the safest match.

## Local source install

### 1. Clone the repository

```powershell
git clone https://github.com/pandadoor/MorphosPowerPointAddIn.git
cd MorphosPowerPointAddIn
```

### 2. Build and register the add-in

```powershell
.\build-and-run.ps1
```

What the script does:

- closes running PowerPoint instances
- cleans old `bin` and `obj` outputs
- rebuilds `MorphosPowerPointAddIn.csproj`
- rewrites the PowerPoint add-in registry entry to the local VSTO manifest
- appends `|vstolocal` to the manifest registration so PowerPoint loads the latest local build
- restores `LoadBehavior=3`
- starts PowerPoint and verifies the COM add-in is connected

### 3. Open Morphos in PowerPoint

1. Open a presentation in PowerPoint.
2. Open the `Morphos` tab on the ribbon.
3. Click `Open Inspector`.

## Useful script options

```powershell
.\build-and-run.ps1 -NoStart
.\build-and-run.ps1 -SkipLoadVerification
.\build-and-run.ps1 -Configuration Release
```

- `-NoStart`: rebuild and register without launching PowerPoint
- `-SkipLoadVerification`: skip the COM add-in connection check
- `-Configuration Release`: build from `bin\x64\Release`

## If PowerPoint disables the add-in

Run:

```powershell
.\build-and-run.ps1
```

The script reasserts the local manifest registration and resets the add-in load behavior.

If you still do not see Morphos:

1. Close PowerPoint completely.
2. Run `.\build-and-run.ps1` again.
3. Open `File -> Options -> Add-ins` in PowerPoint and confirm `MorphosPowerPointAddIn` is present under COM Add-ins.

## Certificate and trust notes

This repository includes `MorphosPowerPointAddIn.cer`. Depending on your Office trust state, Windows or Office may prompt for trust when loading a locally built VSTO add-in.

If trust prompts block loading on your machine:

1. Open `MorphosPowerPointAddIn.cer`.
2. Import it into the appropriate trusted certificate stores for your environment.
3. Rerun `.\build-and-run.ps1`.

Keep certificate trust decisions aligned with your local security policy.

## Uninstall / reset

To remove a locally registered Morphos debug build:

1. Close PowerPoint.
2. Remove the registry key:

```text
HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\MorphosPowerPointAddIn
```

3. Delete local `bin` and `obj` folders if you want a full cleanup.

## Troubleshooting

### PowerPoint opens but Morphos is not connected

- Make sure PowerPoint is the desktop application, not the web app.
- Rerun `.\build-and-run.ps1`.
- Verify the build succeeded and the script reported `Connected: True`.

### Build cleanup fails because files are locked

`build-and-run.ps1` already retries locked cleanup paths. If a lock persists, close PowerPoint and any test runners or shells that may have loaded the add-in assembly, then rerun the script.

### The pane opens but looks empty

Open a presentation first, then open the Morphos pane. The add-in now schedules warm refreshes on presentation open and window activation, so an initial scan should happen automatically.
