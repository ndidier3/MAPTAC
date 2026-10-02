"""
Phase 2 — Analyze (CohortExample skeleton)

Load processed .sdp pickles from Inputs/Skyn_Data_PROCESSED/{COHORT}/,
aggregate curve- and day-level stats, and export dated Excel workbooks
under Results/{COHORT}/.

Run after process/process_data.py (or an equivalent process step).
"""

from datetime import datetime
from pathlib import Path

from App.SDM.Analysis.curveFeatures import curveFeatures
from App.SDM.Analysis.dayFeatures import dayFeatures
from App.SDM.Scripts.CohortExample.cohort_example_settings import (
    curve_attrs,
    smooth_and_impute_attrs,
)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
script_dir = Path(__file__).resolve().parent
project_root = script_dir.parents[4]  # .../Scripts/CohortExample/analysis → repo root

cohort_name = 'CohortExample'
processed_data_folder = project_root / 'Inputs' / 'Skyn_Data_PROCESSED' / cohort_name
results_dir = project_root / 'Results' / cohort_name
results_dir.mkdir(parents=True, exist_ok=True)

today = datetime.today().strftime('%m.%d.%Y')

print(f'Project root:  {project_root}')
print(f'Processed dir: {processed_data_folder}')
print(f'Results dir:   {results_dir}')

# ---------------------------------------------------------------------------
# Curve-level analysis
# ---------------------------------------------------------------------------
print('\nRunning curve-level analysis...')
curves = curveFeatures(
    str(processed_data_folder),
    smooth_and_impute_attrs=smooth_and_impute_attrs,
    curve_attrs=curve_attrs,
)

curves.apply_flag_analysis()
curves.run_stats()
# Optional drinking-curve labeling (quality / shape criteria):
# curves.identify_drinking_curves(min_drinking_duration_minutes=60, require_valid=True)

curve_workbook = results_dir / f'{cohort_name}_curve_stats_{today}.xlsx'
curves.export_workbook_curves(
    str(curve_workbook),
    # split_plots_by='drinking_pred',  # if identify_drinking_curves() was run
    include_plots=False,
)

print(f'Curve workbook: {curve_workbook}')

# ---------------------------------------------------------------------------
# Day-level analysis
# ---------------------------------------------------------------------------
print('\nRunning day-level analysis...')
days = dayFeatures(str(processed_data_folder))
days.compute_low_quality_stats()

# Optional: attach curve-based drinking-day labels
# days.attach_curves_and_predict_drinking_days(curves.curve_features)

day_workbook = results_dir / f'{cohort_name}_day_stats_{today}.xlsx'
days.export_workbook_days(
    str(day_workbook),
    include_nonwear_plots=True,
    include_signal_processing_plots=True,
)

print(f'Day workbook:   {day_workbook}')
print(f'\n{cohort_name} analysis complete.')
