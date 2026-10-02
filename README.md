# GSAT-Scoring

Reproducibility repository for the study evaluating AI model performance on Taiwan College Entrance Examination Center (CEEC) examinations from 2015 to 2024.

This repository contains the analysis-ready AI scoring datasets, CEEC human-score distributions used for AI–human percentile comparison, official examination papers, scoring/benchmark tables, and the Python scripts used for the revised statistical analyses.

## Repository structure

```text
GSAT-Scoring/
├─ README.md
├─ analysis/
│  ├─ empirical_percentile_sensitivity_analysis.py
│  └─ paired_blocked_statistical_analysis.py
│
└─ supplementary-materials/
   ├─ CEEC_Official_Exam_Papers_2015_2024/
   │  ├─ AST_Mathematics_A_2015_2024/
   │  ├─ GSAT_Chinese_2015_2024/
   │  ├─ GSAT_English_2015_2024/
   │  ├─ GSAT_Mathematics_2015_2024/
   │  ├─ GSAT_Natural_Science_2015_2024/
   │  ├─ GSAT_Social_Studies_2015_2024/
   │  └─ README.md
   │
   ├─ analysis-ready-data/
   │  ├─ GSAT_Chinese.xlsx
   │  ├─ GSAT_English.xlsx
   │  ├─ GSAT_Mathematics.xlsx
   │  ├─ GSAT_Science.xlsx
   │  ├─ GSAT_Social_Studies.xlsx
   │  ├─ MathA.xlsx
   │  ├─ Human Baseline_GSAT_2015_2024.xlsx
   │  ├─ Human Baseline_MathA_2014_2024.xlsx
   │  └─ README.md
   │
   └─ scoring-tables/
      ├─ GSAT_Chinese_AI_Benchmark_2015-2024_Data.xlsx
      ├─ GSAT_English_AI_Benchmark_2015-2024_Data.xlsx
      ├─ GSAT_Mathematics_AI_Benchmark_2015-2024_Data.xlsx
      ├─ GSAT_ Social Studies_AI_Benchmark_2015-2024_Data.xlsx
      ├─ GSAT_Science_AI_Benchmark_2015-2024_Data.xlsx
      ├─ MathA_AI_Benchmark_2015-2024_Data.xlsx
      └─ README.md
```

## Analysis scripts

### 1. Empirical percentile and sensitivity analysis

`analysis/empirical_percentile_sensitivity_analysis.py`

This script:

- aggregates AI item-level scores to annual Subject × Model scores;
- harmonizes AI scores and CEEC human distributions to the same comparison scale before percentile mapping;
- calculates the primary human-percentile estimates from CEEC empirical cumulative counts using within-band interpolation;
- applies the mid-rank convention for single-score bands;
- calculates the normal-approximation percentile as a sensitivity analysis;
- performs QC checks for score-scale harmonization, annual completeness, human-distribution totals, interval matching, and subject-year mapping;
- generates manuscript-ready Table 8 and Table 9 outputs.

### 2. Paired / blocked statistical analysis

`analysis/paired_blocked_statistical_analysis.py`

This script:

- calculates annual objective-choice accuracy (%);
- treats examination year (2015–2024) as the matched block;
- applies the Friedman test across the five AI models within each subject;
- reports Kendall's W as the Friedman effect-size measure;
- performs two-sided paired Wilcoxon signed-rank post-hoc comparisons;
- applies Holm correction across the 10 model-pair comparisons within each subject;
- exports the statistical tables and QC outputs used in the revised analysis.

## Fixed score-scale definitions

The study distinguishes **AI Accuracy** from the **human-percentile comparison score**. AI and human scores are harmonized to the same comparison scale before percentile estimation.

### GSAT Chinese

- **2015–2017**
  - AI objectively scoreable maximum = 54
  - AI Accuracy = AI raw score / 54 × 100
  - AI percentile comparison score = AI raw score / 54 × 100
  - CEEC human distribution uses the full 108-point scale
  - Human comparison boundary = official human raw boundary / 108 × 100
- **2018–2024**
  - AI and human scores are already on the 0–100 scale; no additional conversion is applied.

### GSAT English

- **2015–2024**
  - AI objectively scoreable maximum = 72
  - AI Accuracy = AI raw score / 72 × 100
  - AI percentile comparison score = AI raw score / 72 × 100
  - CEEC human distributions remain on the official 0–100 scale.

### Other GSAT subjects

For GSAT Mathematics, Social Studies, and Natural Science, no artificial scale conversion is applied when AI annual scores and CEEC human distributions are already on the same official scale.

For the GSAT Mathematics human baseline:

- 2015–2021: Math
- 2022–2024: MathA
- MathB is not included in the formal study analysis.

### AST Mathematics A

- **2015–2021:** AI and human distributions are compared on the raw-score scale.
- **2022–2024:** AI raw scores are converted to the official CEEC 0–60 level scale before comparison with the official human level distribution.

Official annual intervals:

- 2022: 1.43700
- 2023: 1.36100
- 2024: 1.56867

## Primary percentile method

The primary AI–human percentile analysis uses CEEC official empirical cumulative counts:

```text
Percentile =
[N below band + within-band fraction × N in band]
/ Total N × 100
```

For a continuous score band:

```text
within-band fraction =
(AI comparison score - lower boundary)
/
(upper boundary - lower boundary)
```

For a single-score band, the mid-rank convention is used:

```text
within-band fraction = 0.5
```

The normal approximation is retained only as a sensitivity analysis.

## Requirements

Install the required Python packages:

```bash
pip install pandas numpy scipy statsmodels openpyxl python-docx
```

## Reproducing the analyses

Download or clone the complete repository. The scripts use repository-relative paths and do not require computer-specific absolute paths.

From the repository root, run:

```bash
python analysis/empirical_percentile_sensitivity_analysis.py
python analysis/paired_blocked_statistical_analysis.py
```

Each script reads the required files directly from:

```text
supplementary-materials/analysis-ready-data/
```

The scripts create an `output/` directory automatically when executed and write their generated Excel, Word, CSV, and log files there.

The two analysis scripts are independently executable; neither requires a preceding Jupyter Notebook step.

## Data organization

- `supplementary-materials/analysis-ready-data/` contains the standardized AI analysis datasets and the CEEC human baseline distributions required by the executable analysis scripts.
- `supplementary-materials/scoring-tables/` contains the benchmark/scoring workbooks used for AI score calculation and auditability.
- `supplementary-materials/CEEC_Official_Exam_Papers_2015_2024/` contains the official CEEC examination papers used in the study.

## Reproducibility note

The public GitHub version uses the same score-scale conversion rules, AI score calculations, human-score harmonization, empirical percentile method, within-band interpolation, single-score mid-rank convention, subject-year mapping, human baseline mapping, AST Mathematics A level conversion, sensitivity-analysis method, and Table 8/Table 9 calculation logic as the revised study analysis. Repository-relative file paths are used only to make the analysis portable across computers.
