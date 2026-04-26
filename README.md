<p align="center">
  <img src="docs/assets/morphos-banner.png" alt="Morphos Banner" width="100%">
</p>

<p align="center">
  <img src="docs/assets/morphos-icon.png" alt="Morphos Icon" width="120">
</p>

<h1 align="center">Morphos</h1>

<p align="center">
  <strong>Integrated Presentation Inspection and Cleanup for Microsoft PowerPoint</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows-0078D4?style=flat-square" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-A86B27?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Build-VSTO-5C2D91?style=flat-square" alt="Build">
  <img src="https://img.shields.io/badge/PowerPoint-2016+-D83B01?style=flat-square" alt="PowerPoint">
</p>

---

## Overview

Morphos is a professional-grade VSTO (Visual Studio Tools for Office) add-in for Microsoft PowerPoint. It provides a centralized workspace within the PowerPoint environment to identify and resolve inconsistencies in fonts and colors. By consolidating fragmented cleanup tasks into a responsive task pane, Morphos enhances productivity and ensures presentation integrity.

## Core Functionality

### Automated Presentation Auditing
The system performs real-time scans of the active presentation to surface critical issues that impact deck quality.

*   **Font Inventory Management**: Detects missing, non-installed, and embedded fonts.
*   **Color Usage Analysis**: Identifies direct RGB color usage and highlights theme inconsistencies.
*   **Slide-Level Drilldown**: Enables direct navigation to the specific slides containing identified issues.

### Precision Cleanup Tools
Morphos facilitates safe and efficient bulk modifications without requiring the user to navigate through multiple native dialogs.

*   **Intelligent Font Replacement**: Validates replacement targets against system fonts and active theme tokens.
*   **Theme Alignment**: Simplifies the conversion of ad-hoc colors to standardized theme-linked colors.
*   **Safe Execution**: Implements validation layers to ensure modifications are technically sound before execution.

### Performance Optimization
Designed for professional use cases involving large-scale presentations.

*   **Open XML Integration**: Utilizes low-level package inspection for high-performance scanning.
*   **Smart Caching**: Employs the `FontScanSessionCache` to minimize redundant processing during repeated operations.
*   **Stability Framework**: Built-in retry mechanisms and Office interop filters manage transient application states reliably.

---

## Technical Architecture

The project follows a layered architecture to ensure maintainability and high performance:

*   **Host Layer**: Managed via `ThisAddIn.cs` for application lifecycle and UI integration.
*   **Logic Layer**: Orchestrated by `PowerPointPresentationService.cs` for scanning and mutation.
*   **Data Access**: Optimized through Open XML services for direct package-level interaction.
*   **Interface**: Built with WPF to provide a modern, responsive user experience within the task pane.

---

## Installation and Deployment

### 1. Prerequisites
Ensure the environment meets the following requirements:
*   Microsoft PowerPoint 2016 or newer (Windows)
*   .NET Framework 4.8
*   Visual Studio Build Tools (for local compilation)

### 2. Automated Setup
The provided automation script handles the end-to-end deployment process, including compilation and registry registration:

```powershell
.\build-and-run.ps1
```

### 3. Usage
Once installed, the add-in is accessible via the **Morphos** tab on the PowerPoint Ribbon. Click **Open Inspector** to launch the task pane.

---

## Documentation and Licensing

*   **Licensing**: This software is released under the [MIT License](LICENSE).
*   **Setup Details**: Refer to [docs/INSTALLATION.md](docs/INSTALLATION.md) for detailed configuration.
*   **Development**: Technical specifications are available in [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

---

<p align="center">
  <em>Engineering efficiency into every presentation.</em>
</p>
