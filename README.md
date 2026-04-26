<p align="center">
  <img src="docs/assets/morphos-icon.png" alt="Morphos Logo" width="120">
</p>

<h1 align="center">Morphos</h1>

<p align="center">
  <strong>Enterprise-Grade Presentation Asset Management and Quality Assurance for Microsoft PowerPoint</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows-0078D4?style=flat-square" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-A86B27?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Stack-VSTO%20%7C%20.NET%204.8%20%7C%20WPF-5C2D91?style=flat-square" alt="Stack">
  <img src="https://img.shields.io/badge/Compliance-Open%20XML%20SDK-D83B01?style=flat-square" alt="Compliance">
</p>

---

## Technical Abstract

Morphos is a high-performance VSTO (Visual Studio Tools for Office) solution engineered to optimize the presentation refinement lifecycle within Microsoft PowerPoint. It leverages low-level Open XML package inspection and a robust asynchronous processing architecture to provide real-time auditing and bulk remediation of document assets, specifically targeting font and color consistency.

## Architectural Overview

The system is built upon a decoupled, service-oriented architecture designed to mitigate the inherent latency of COM-based Office interop while maintaining strict thread safety and UI responsiveness.

### Core Architecture Layers

*   **Application Lifecycle Management**: The `ThisAddIn` controller orchestrates the integration with the PowerPoint host, managing multi-window task pane synchronization and event-driven state transitions.
*   **Asynchronous Service Layer**: The `PowerPointPresentationService` handles the heavy lifting of presentation scanning and mutation, utilizing background task scheduling to prevent UI thread blocking.
*   **Asset Inspection & Mutation**: A hybrid approach combines the **Office Object Model (COM)** for live document interaction and the **Open XML SDK** for high-speed package-level auditing and bulk modifications.
*   **Presentation State Management**: Implemented via **WPF and the MVVM pattern**, the user interface provides a responsive, state-aware workspace with real-time progress reporting and error handling.

## System Lifecycle and Integration

Morphos integrates deeply with the Microsoft PowerPoint application lifecycle through a series of deterministic event handlers:

*   **Initialization**: On startup, the add-in initializes its internal services and registers global application hooks (`PresentationOpen`, `WindowActivate`).
*   **Dynamic Task Pane Synchronization**: The system monitors window activation states to ensure the Morphos task pane remains synchronized with the active document context, creating or recovering pane instances as required.
*   **Automated Auditing**: Proactive scanning is triggered by document activation and visibility changes, ensuring the user is always presented with the most current asset inventory.

## Performance Optimization Strategies

To meet the demands of enterprise-scale presentations, Morphos implements several advanced performance optimizations:

*   **Open XML Fast Paths**: Utilizing direct package inspection allows for significantly faster scanning of large presentations compared to traditional COM-based iteration.
*   **Session-Scoped Caching**: The `FontScanSessionCache` minimizes redundant IO and processing by maintaining snapshots of presentation metadata, which are only invalidated upon material document changes.
*   **COM Accessor Optimization**: The `ComFontAccessorCache` reduces the overhead of repeated COM property lookups during intensive font audits.
*   **Retry and Message Filtering**: Implements custom `IOleMessageFilter` logic to gracefully handle Office "Busy" or "Rejected" states, ensuring operational reliability during concurrent PowerPoint activities.

---

## Deployment and Configuration

### System Requirements

| Component | Specification |
| :--- | :--- |
| Operating System | Windows 10 / 11 (x64) |
| Host Application | Microsoft PowerPoint 2016, 2019, 2021, or Microsoft 365 |
| Runtime Environment | .NET Framework 4.8 |
| Development Framework | VSTO (Visual Studio Tools for Office) |

### Installation Procedure

The project includes a comprehensive PowerShell automation script for end-to-end deployment, covering compilation, manifest generation, and registry configuration:

```powershell
# Execute the automated build and deployment sequence
.\build-and-run.ps1
```

---

## Documentation and Compliance

*   **Licensing**: Distributed under the [MIT License](LICENSE).
*   **Security**: Implements safe mutation patterns with validation layers to prevent document corruption.
*   **Compliance**: Adheres to Microsoft Office UI Guidelines and Open XML File Format standards.

---

<p align="center">
  <em>Optimizing digital communication through engineering excellence.</em>
</p>
