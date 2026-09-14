# Literature Retrieval Quality Assessor (LRQA)

A specialized tool for systematic review literature retrieval, supporting overlap quantification, quality evaluation and batch screening of multi-strategy results.

## Features
- **Overlap Quantification**: Calculates pairwise literature overlap via the Sørensen-Dice coefficient, with overlap matrix and heatmap visualization.
- **Journal Tier System**: Built-in T0-T4 ranking, supports custom rank import, with Bayesian shrinkage for small samples.
- **4D Quality Scoring**: Evaluates retrieval strategies by core capture rate, journal quality, unique contribution and redundancy penalty.
- **Flexible Filtering**: Filter by occurrence frequency and journal tier, export results with complete original metadata.
- **Full Deduplication**: Merge all input sets and export a deduplicated full table in one click.
- **Bilingual GUI**: Chinese / English dual language interface.

## Installation
```bash
pip install pandas ttkbootstrap matplotlib openpyxl xlrd

