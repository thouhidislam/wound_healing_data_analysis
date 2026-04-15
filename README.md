# Wound Healing Data Analysis (R)

This repository contains an R-based workflow for analyzing wound healing data using:

- One-way ANOVA
- Tukey's post hoc test
- Mean ± SEM calculation
- Percentage wound contraction
- Publication-ready tables (PLOS format)

## Features
- Automated statistical analysis
- Prism-style significance annotation
- Excel and Word export
- Fully reproducible pipeline

## Requirements
- R (>= 4.2)
- tidyverse
- rstatix
- flextable
- officer
- rio

## Usage
1. Place your dataset in `data/`
2. Update the dataset name in `analysis.R`
3. Run the script:

```r
source("scripts/analysis.R")
