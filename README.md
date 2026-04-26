<p align="center">
  <img src="docs/assets/morphos-banner.png" alt="Morphos Banner" width="100%">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows-blue.svg" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-gold.svg" alt="License">
  <img src="https://img.shields.io/badge/Build-VSTO-brown.svg" alt="Build">
  <img src="https://img.shields.io/badge/PowerPoint-2016+-orange.svg" alt="PowerPoint">
</p>

---

# Morphos

### PowerPoint cleanup, without leaving PowerPoint.

**Morphos** is a professional-grade Windows VSTO add-in for Microsoft PowerPoint designed to streamline the deck cleanup process. It scans your active presentation, identifies critical font and color inconsistencies, and provides a focused workspace to repair them—all from a compact, responsive task pane.

## ✨ Why Morphos?

PowerPoint cleanup is traditionally a fragmented process involving multiple dialogs and manual slide inspection. Morphos consolidates this workflow into four intuitive steps:
1.  **Scan**: Automatically audit your deck's internal structure.
2.  **Inspect**: Visualize hidden issues like missing fonts or non-theme colors.
3.  **Replace**: Execute precise, bulk repairs safely.
4.  **Verify**: Confirm your deck is professional and consistent.

---

## 🚀 Key Capabilities

### 🔎 Intelligent Font Inventory
*   **Audit**: View usage counts, installation status, and embedding states.
*   **Drilldown**: Jump directly to specific slides containing the target font.
*   **Safe Replacement**: Validate replacement targets against real system fonts and active theme tokens.

### 🎨 Precise Color Management
*   **Inventory**: Group direct RGB color usage into actionable cleanup tables.
*   **Theme Alignment**: Identify colors that *should* be theme-linked and convert them in bulk.
*   **Visual Context**: Navigate slide-by-slide to see exactly where colors are being used.

### ⚡ Optimized for Performance
*   **Open XML Speed**: Uses low-level package scanning to avoid COM-heavy overhead on large decks.
*   **Smart Caching**: Reuses presentation snapshots (`FontScanSessionCache`) to keep repeated scans lightning-fast.
*   **Robust Interop**: Built-in retry helpers and busy-state filters to ensure stability during intensive PowerPoint operations.

---

## 🛠 Architecture at a Glance

Morphos is built with a clean, layered architecture that separates Office interop from business logic and UI:

- **Core Hosting**: `ThisAddIn.cs` manages the lifecycle and task-pane integration.
- **Service Layer**: `PowerPointPresentationService.cs` orchestrates scans and mutations.
- **Fast Paths**: `OpenXml*.cs` provides direct package inspection for high-volume decks.
- **Modern UI**: WPF-based task pane with responsive layouts and MVVM state management.

---

## 🏁 Quick Start

### 1. Clone the Repository
```powershell
git clone https://github.com/pandadoor/MorphosPowerPointAddIn.git
cd MorphosPowerPointAddIn
```

### 2. Build and Install
Run the automation script to handle everything from compilation to registry registration:
```powershell
.\build-and-run.ps1
```
*The script rebuilds the add-in, configures the local VSTO manifest, and launches PowerPoint automatically.*

### 3. Open the Inspector
1.  Open any PowerPoint presentation.
2.  Navigate to the **Morphos** tab on the Ribbon.
3.  Click **Open Inspector** to begin your scan.

---

## 📄 Licensing & Documentation

*   **License**: This project is licensed under the [MIT License](LICENSE).
*   **Setup Guide**: See [docs/INSTALLATION.md](docs/INSTALLATION.md) for environment requirements.
*   **Developer Notes**: See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for build and testing details.

---

<p align="center">
  <i>Built for precision. Designed for efficiency.</i>
</p>
