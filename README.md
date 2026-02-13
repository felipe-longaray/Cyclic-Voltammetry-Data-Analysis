# Cyclic Voltammetry (CV) Peak Analysis Pipeline

[![R](https://img.shields.io/badge/R-276DC3?logo=r&logoColor=white)](https://www.r-project.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

This repository contains an automated R pipeline for processing Cyclic Voltammetry (CV) data. It is designed to extract key thermodynamic and kinetic parameters from raw multicycle instrument exports (such as those from PalmSens PSTrace) without relying on manual peak picking.

## Electrochemical Context
In analyzing reversible and quasi-reversible faradaic processes, extracting accurate peak parameters is critical. This script automatically evaluates the voltammogram to identify:

1. **Anodic Peak Current ($I_{pa}$):** The maximum oxidation current.
2. **Cathodic Peak Current ($I_{pc}$):** The minimum reduction current.
3. **Peak Potential Separation ($\Delta E_p$):** Calculated as $|E_{pa} - E_{pc}|$.

$\Delta E_p$ is a crucial diagnostic criterion. For an ideal, fully reversible single-electron transfer process at $25^\circ C$, the theoretical separation is $\approx 59 \text{ mV}$. Deviations from this value, or varying $\Delta E_p$ across different experimental conditions, can indicate quasi-reversibility, uncompensated resistance, or coupled chemical reactions (EC mechanisms).

## Key Features
* **Automated Peak Picking:** Programmatically scans alternating Potential/Current columns to locate local maxima and minima ($E_{pa}$ and $E_{pc}$).
* **Statistical Rigor:** Automatically calculates global ANOVA across different experimental conditions (e.g., Positive, Negative, Ultrapure Water arrays) to determine statistical significance of peak shifts.
* **Publication-Grade Plotting:** Uses `ggplot2` to generate boxplots with integrated p-values and proper standard mathematical notation (e.g., $\Delta E_p$, $I_{pa}$) on axis labels.
* **Excel Consolidation:** Exports raw extracted peak metrics and aggregated statistical summaries into clearly defined sheets for easy sharing with collaborators.

## Dependencies
* `readxl`, `openxlsx` (I/O handling)
* `dplyr` (Data manipulation)
* `ggplot2`, `ggpubr` (Visualization and statistical testing)

## Usage
1. Place the exported CV `.xlsx` file into a local `/data` directory.
2. Run `analyze_cv.R` in R or RStudio.
3. The script will output the compiled graphs and save a consolidated `CV_Consolidated_Results.xlsx` data file.

## Author
**Felipe Longaray Kadel**
*Materials Engineering Undergraduate - UFRGS*
felipe.longaray@ufrgs.br
