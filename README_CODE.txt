REPRODUCIBILITY CODE

Purpose
-------
generate_figures.py regenerates all five manuscript figures and the supplementary
clustering figure directly from Supplementary_Data_S1_Sediment_LCA_Evidence.xlsx.
It does not read the copyrighted article PDFs and does not depend on intermediate
CSV files.

The generated figures contain no internal figure or panel titles. Multi-panel
figures use only the labels (A), (B), (C), and (D) at the upper-left of each
panel; complete descriptions are supplied through the manuscript captions.

Recommended environment
-----------------------
Python 3.10 or later. Install the required packages with:

    python -m pip install -r requirements.txt

Run
---
Place this directory next to the evidence workbook, then run:

    python generate_figures.py \
      --workbook ../Supplementary_Data_S1_Sediment_LCA_Evidence.xlsx \
      --output-dir generated_figures

Optional argument:

    --dpi 420

Outputs
-------
The output directory will contain PNG and PDF versions of:

    Figure_01_PRISMA
    Figure_02_Evidence_trend
    Figure_03_Methodological_choices
    Figure_04_Interpretation_context_circularity
    Figure_05_Integrated_evidence_and_outcomes
    Figure_S1_Clustering_audit

It also contains figure_manifest.csv and a derived_tables directory. The derived
tables are recalculated from the workbook and provide an audit trail for the
counts displayed in the figures.

Analytical scope
----------------
- The technology-contaminant matrix reports study-level evidence coverage and is
  not a ranking of technology performance.
- Environmental outcomes are counts of explicit study-level directional
  statements. LCIA magnitudes are not pooled.
- The six-item completeness score describes methodological reporting; it is not a
  validity or risk-of-bias weight.
- The three-cluster dendrogram is exploratory. The script recalculates the
  clustering under Hamming, missing-excluded and Jaccard specifications and
  reports their adjusted Rand indices.

Expected workbook structure
---------------------------
The script requires these sheets and their existing column names:

    Screening
    Article Evidence
    Contribution Analysis
    Outcome Evidence

The supplied Supplementary Dataset S1 contains the complete required structure.

Reproducibility note
--------------------
Eligibility decisions, evidence coding and scientific interpretation were made
and verified by the researcher. Code-generation assistance and automated
consistency checks do not replace source-level verification.
