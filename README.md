# Monitoring and Processing Tool for Alcohol Concentration (MAPTAC)

Python toolkit for processing Skyn biosensor (TAC) data and running cohort-level analyses. The repo is organized around **study-specific scripts** that call shared processing and analysis modules, with **SLURM batch scripts** for running large jobs on the cluster.

> **Naming:** This project is now **MAPTAC** (Monitoring and Processing Tool for Alcohol Concentration). It was formerly **Skyn Data Manager**; the **SDM** acronym is retained for legacy purposes and still appears in paths (`App/SDM/`), the conda env (`sdm-env`), and older scripts.

**Reference script:** [`App/SDM/Scripts/Test/default_analysis.py`](App/SDM/Scripts/Test/default_analysis.py) — minimal end-to-end example for local testing.

---

## Repository layout

```
skyn_data_manager/
├── App/SDM/                  # Core library
│   ├── Run/                  # Cohort orchestration (process_many, process_single)
│   ├── Skyn_Processors/      # Per-subject skynDataset pipeline
│   ├── Analysis/             # Cohort-level stats (curves, days, events)
│   ├── Configuration/        # File I/O, timestamps, helpers
│   ├── Signal_Processing/    # Imputation, curve detection
│   ├── Visualization/        # Plot and workbook generation
│   └── Scripts/              # Study-specific entry-point scripts
│       ├── Test/             # Minimal end-to-end example
│       ├── [PROJECT NAME 1]/
│       ├── [PROJECT NAME 2]/
│       └── ...
├── Inputs/
│   ├── Skyn_Data_RAW/        # Raw Skyn CSV exports, by cohort
│   ├── Skyn_Data_PROCESSED/  # Saved per-subject processors (.sdp)
│   └── Metadata/             # Events, study dates, annotations
├── Results/                  # Dated Excel workbooks and outputs, by cohort
├── batch_scripts/            # SLURM sbatch wrappers
└── batch_out/                # Job stdout/stderr (gitignored)
```

Study scripts live under `App/SDM/Scripts/{COHORT}/`. Each cohort typically has:

- **`process/`** — run signal processing, curve/day identification, and (where applicable) event matching
- **`analysis/`** — load processed data and export cohort-level workbooks or figures
- **`{cohort}_settings.py`** — cohort overrides on top of shared defaults

---

## How a cohort run works

Most studies follow a **two-phase** pattern:

| Phase | What it does | Where state lives |
|-------|----------------|-------------------|
| **1. Process** | Loop subjects, run `process_and_analyze_data()`, save one pickle per subject | `Inputs/Skyn_Data_PROCESSED/{COHORT}/` |
| **2. Analyze** | Load those pickles, aggregate stats, export Excel workbooks | `Results/{COHORT}/` |

The central entry point for phase 1 is `App/SDM/Run/process_many.py` → `process_and_analyze_data()`. Phase 2 uses classes such as `curveFeatures`, `dayFeatures`, or `curveFeaturesWithEvents` in `App/SDM/Analysis/`.

Heavy runs are often split so you do not repeat expensive steps:

1. **`process_raw_data.py`** — gaps/non-wear, smoothing, imputation (writes `.sdp` files)
2. **`process_curves_and_events.py`** — load prior saves, identify curves/days, match events (if any)
3. **`analysis/*.py`** — cohort statistics and exports

The **Test** cohort is the simplest full example in one file: [App/SDM/Scripts/Test/default_analysis.py](App/SDM/Scripts/Test/default_analysis.py). It reprocesses the nine raw files in `Inputs/Skyn_Data_RAW/TestData` and writes `Results/Test`.

---

## Day-level drinking detection

Day-level drinking is an XGBoost model, not a curve-overlap rule. After day features and curve slots are attached, `dayFeatures.attach_curves_and_predict_drinking_days` scores every social day (06:00–06:00). The stored model is `App/SDM/Trained_Models/xgb_drinking_day_v0805.json` (XGBoost 2.1.4). Scoring code is `App/SDM/Machine_Learning/drinking_day_xgb.py`. Training details, the 150 inputs, and leave-one-subject-out metrics are in [App/SDM/Trained_Models/drinking_day_xgb_README.md](App/SDM/Trained_Models/drinking_day_xgb_README.md).

The training labels are human reviews of TAC plots (`annotation_drinking_day`), not morning self-report and not a laboratory alcohol assay. On the 2,361 labeled device-on days used for training, leave-one-subject-out accuracy was 0.956 (sensitivity 0.934, specificity 0.964, ROC AUC 0.990).

| Column | Meaning |
|--------|---------|
| `pred_drinking_day` | 1 when the predicted probability is at least 0.5 |
| `proba_drinking_day` | Predicted probability of a drinking day |
| `prediction_reliable` | 1 only when none of the flags below are 1 |

`prediction_reliable` is 0 when any of these is true. The cutoffs are arguments of `attach_curves_and_predict_drinking_days` (the Test script sets the defaults below):

| Argument | Default in the Test script | Flag |
|----------|----------------------------|------|
| `flag_non_wear_hours` | 6 | Hours not worn are at least this value. Not worn = 24 − `device_worn_duration` (powered-off or missing time plus on-device non-wear). |
| `flag_very_negative_hours` | 1 | `extreme_negative_duration` (TAC &lt; −15) is at least this many hours |
| `flag_artifact_hours` | 1 | `jump_duration` + `plummet_duration` is at least this many hours |
| `flag_proba_low`, `flag_proba_high` | 0.05, 0.95 | `proba_drinking_day` falls inside this inclusive interval |

Days with no TAC are still scored. A powered-off day is usually predicted non-drinking and is marked unreliable. Curve-level `DRINKING_PRED` (from `identify_drinking_curves`) is a separate label and is not copied onto the day table. Day workbooks split plots with `split_plots_by='drinking_xgb'`.

---

## Settings

Settings are layered:

1. **Defaults** — `App/SDM/Run/default_settings/` (`default_curve_settings.py`, `default_flag_settings.py`, `default_smooth_impute_settings.py`)
2. **Cohort file** — e.g. `cohort_example_settings.py`, `ace_settings.py` (imports defaults, overrides threshold, day start hour, event columns, etc.)
3. **Script call** — arguments passed to `process_and_analyze_data()` control which pipeline steps run

Common arguments when calling `process_and_analyze_data()`:

| Argument | Typical use |
|----------|-------------|
| `use_prior_save` | `True` to reload existing `.sdp` instead of reprocessing from raw |
| `adjust_for_gaps_and_non_wear` | First-pass raw processing |
| `smooth_and_impute` | First-pass raw processing |
| `identify_curves` / `analyze_days` | Curve and day feature extraction |
| `match_events_to_curves` | Studies with self-report event CSVs (ACE, ARC) |
| `filter_by_study_dates` | Trim to per-subject study windows |
| `subids_to_process` | Restrict to specific subject IDs |

When adding a new cohort, copy an existing `{cohort}_settings.py` and a `process/` script from the example files.

---

## Writing a new script

1. **Create a cohort folder** under `App/SDM/Scripts/MyCohort/` with `process/` and `analysis/` as needed.
2. **Add `{cohort}_settings.py`** — import from `default_settings`, override only what differs.
3. **Process script** — set paths, import settings, call `process_and_analyze_data()` with the right arguments.
4. **Analysis script** — point at `Inputs/Skyn_Data_PROCESSED/MyCohort`, run the appropriate `Analysis` class, write dated output under `Results/MyCohort/`.

**Paths:** Prefer resolving the repo root from the script file (see `default_analysis.py`) so the same script works locally and on the cluster. Many existing scripts use a hardcoded `/users/ndidier/SDM/skyn_data_manager` path; update that to your clone when developing elsewhere.

**Outputs:** Workbooks are usually named `{COHORT}_{description}_{MM.DD.YYYY}.xlsx`.

**Run locally first** on a small subset (`Test` data or `subids_to_process`) before submitting a full cohort job.

---

## Running on the cluster (batch_scripts)

Jobs are submitted with **SLURM** (`sbatch`) from the repo root. Scripts in `batch_scripts/` follow a common pattern:

```bash
cd /path/to/skyn_data_manager
source ~/miniconda3/etc/profile.d/conda.sh
conda activate sdm-env
python App/SDM/Scripts/...
conda deactivate
```

`batch_scripts/sdm_test_run.sh` is the small drinking-day smoke test. It activates `sdm-env` and runs `App/SDM/Scripts/Test/default_analysis.py`, which scores days with the 08.05 model and writes `Results/Test`. Submit it from the repo root with `sbatch batch_scripts/sdm_test_run.sh`. Other cohort shell scripts are local to this cluster account and are not in git.

### Naming

| Pattern | Example | Purpose |
|---------|---------|---------|
| `sdm_test_run.sh` | — | Small smoke test (`Test/default_analysis.py`) |
| `sdm_run_{cohort}.sh` | `sdm_run_mycohort.sh` | Main cohort job |
| `sdm_submit_{pipeline}.sh` | `sdm_submit_mycohort_pipeline.sh` | Chain jobs with dependencies |
| `sdm_run_{task}.sh` | `sdm_run_mycohort_xgb.sh` | Follow-on analysis (ML, plots) |

Logs go to `batch_out/{JobName}-{JOB_ID}.out` and `.err`.

### Submitting

```bash
# From repo root
sbatch batch_scripts/sdm_test_run.sh

# Cohort job (edit the .sh file to uncomment the python lines you need)
sbatch batch_scripts/sdm_run_{cohort}.sh
```

Pipeline wrappers can use `sbatch --dependency=afterok:...` so a follow-on job starts only if the first succeeds. Pass variables with `--export=ALL,VAR=value` (for example subject-range or input-path overrides defined in the wrapper).

### Resource defaults (adjust per job)

| Script type | Time | Memory | Notes |
|-------------|------|--------|--------|
| Test | ~1 h | 4G | 1 node |
| Cohort process/analyze | ~30 h | 60G | Often 4 nodes |

Update `#SBATCH` lines and the `python` command(s) in the shell script before submitting. Comment/uncomment steps to match the pipeline stage you are running—batch files often keep earlier steps commented when only re-running analysis.

Monitor: `squeue -u $USER` · Inspect failures: `batch_out/*.err`

---

## Data conventions

| Location | Contents |
|----------|----------|
| `Inputs/Skyn_Data_RAW/{COHORT}/` | Raw Skyn CSV files |
| `Inputs/Skyn_Data_PROCESSED/{COHORT}/` | `{subid}_{dataset}_skyn_data_processed.sdp` |
| `Inputs/Metadata/{COHORT}/` | Events, study-day lists, annotations |
| `Results/{COHORT}/` | Analysis exports |

Per-run processing also writes under `Results/{COHORT}/{MM.DD.YYYY}/` (combined day/curve tables, plots).

---

## Further reading

| Topic | Location |
|-------|----------|
| End-to-end script (single file) | [App/SDM/Scripts/Test/default_analysis.py](App/SDM/Scripts/Test/default_analysis.py) |
| Drinking-day model | [App/SDM/Trained_Models/drinking_day_xgb_README.md](App/SDM/Trained_Models/drinking_day_xgb_README.md) |
| Process / analyze skeleton | [App/SDM/Scripts/CohortExample/](App/SDM/Scripts/CohortExample/) |
| Cohort settings template | [App/SDM/Scripts/CohortExample/cohort_example_settings.py](App/SDM/Scripts/CohortExample/cohort_example_settings.py) |

---

## Development setup

### Environment and tools

- **Conda** — environment management
- **VS Code** — recommended editor
- **Git** — version control
- **Python 3.8.8**

### Install

```bash
conda create --name sdm-env python=3.8.8
conda activate sdm-env
conda install --file requirements.txt
conda install -c conda-forge kneed=0.7.0 scikit-learn=1.3.0
```

### Quick start

```bash
# Local smoke test, from the repo root.
# The script adds the repo root to the import path itself.
python App/SDM/Scripts/Test/default_analysis.py

# Or on the cluster (same script; 1 hour / 4G request)
sbatch batch_scripts/sdm_test_run.sh
```
