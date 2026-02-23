# STAJ Implementation Reporting Tool

## ⚖️ Proprietary Notice
**Copyright © 2024-2026 Judiciary of Kenya. All Rights Reserved.**

This software and its source code are the exclusive property of the **Judiciary of Kenya**. Unauthorized copying, distribution, or modification of this code, via any medium, is strictly prohibited. This tool is intended for internal use by authorized personnel only.

---

## 📋 Project Overview
The **STAJ Reporting Tool** is a Python-based automation utility designed to streamline the monitoring and evaluation of the **Social Transformation through Access to Justice (STAJ)** initiative. 

The tool extracts performance data from complex Excel workbooks and transforms it into high-fidelity, executive-level PDF reports featuring dynamic visualizations and institutional branding.

### Core Capabilities
* **Data Extraction:** Automatically parses Excel structures to identify targets and quarterly achievements.
* **Intelligent Metrics:** Calculates overall performance averages with a 100% cap on individual indicators to ensure statistical integrity.
* **Dynamic Visualizations:** * **Gauge Dials:** Real-time generation of Plotly-based gauges showing overall progress.
    * **Progress Bars:** Color-coded status bars (Red → Gold → Green) for individual Outcomes and Lead Units.
* **macOS Integration:** Native AppleScript hooks for a seamless "Point-and-Click" file selection and save experience.

---

## 🚀 Installation & Setup

### 1. Prerequisites
Ensure you are running **Python 3.9** or higher on a macOS environment.

### 2. Required Libraries
Install the necessary dependencies via terminal:
```bash
pip install openpyxl reportlab plotly pandas kaleido numpy
