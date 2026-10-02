"""
Test Analysis Script with Drinking Detection

This script processes SKYN data and performs curve-level and day-level analysis.

Day-level drinking prediction is the 08.05 XGBoost model (pred_drinking_day),
scored when curve overlap is attached. Reliability flags are written alongside it.

Curve-level DRINKING_PRED (quality and shape rules) is still computed for curve
workbook splits. Day-level drinking is pred_drinking_day only.
"""

import sys
from pathlib import Path

# Repo root must be importable as ``App`` when this file is run as a script.
_script_dir = Path(__file__).resolve().parent
_project_root = _script_dir.parent.parent.parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

import pandas as pd
import os
from App.SDM.Run.process_many import *
from App.SDM.Analysis.curveFeatures import curveFeatures
from App.SDM.Analysis.dayFeatures import dayFeatures
from App.SDM.Scripts.Test.test_settings import (
    smooth_and_impute_attrs,
    curve_attrs,
    day_attrs,
    gaps_and_non_wear_attrs,
    event_attrs
)

# gaps_and_non_wear_attrs['export_excel'] = True
smooth_and_impute_attrs['export_excel'] = False

# Dynamic path resolution - works regardless of where project is cloned
script_dir = _script_dir
project_root = _project_root

# Alternative method using os.path (more compatible with older Python versions)
# script_dir = os.path.dirname(os.path.abspath(__file__))
# project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(script_dir))))

print(f"Script directory: {script_dir}")
print(f"Project root: {project_root}")

# Path settings
data_input_folder = project_root / 'Inputs' / 'Skyn_Data_RAW' / 'TestData'
processed_data_folder = project_root / 'Inputs' / 'Skyn_Data_PROCESSED' / 'TestData'
cohort_name = 'Test'

# Process and analyze data with enhanced settings matching ARC analysis approach
process_and_analyze_data(
    project_root, data_input_folder, cohort_name,
    use_prior_save=False,  # Load previously processed data if available
    adjust_for_gaps_and_non_wear=True,  # Remove gaps and non-wear periods
    smooth_and_impute=True,  # Smooth signals and impute missing values
    analyze_days=True,  # Perform day-level analysis
    compute_curve_threshold=True,  # Auto-compute curve detection threshold
    identify_curves=True,  # Detect drinking curves in the data
    include_raw_curves=False,  # Compute curve features made from non-corrected data
    match_events_to_curves=False,  # Match drinking events to detected curves
    gaps_and_non_wear_attrs=gaps_and_non_wear_attrs,
    smooth_and_impute_attrs=smooth_and_impute_attrs,
    day_attrs=day_attrs,
    curve_attrs=curve_attrs,
    event_attrs=event_attrs
)

from datetime import datetime
today = datetime.today().strftime('%m.%d.%Y')

# Initialize curveFeatures with processed data and settings
curves = curveFeatures(processed_data_folder,
                        smooth_and_impute_attrs=smooth_and_impute_attrs,
                        curve_attrs=curve_attrs)
output_dir = project_root / 'Results' / cohort_name

# Identify drinking curves using quality and shape criteria
print(f"\nIdentifying drinking curves...")
curves.identify_drinking_curves()

# Run statistics
curves.run_stats()
curves.count_curve_flags()
curves.compute_imputation_stats()

# Export curve workbook with drinking prediction splits
curves.export_workbook_curves(
    str(project_root / 'Results' / cohort_name / f'{cohort_name}_curve_stats_{today}.xlsx'),
    split_plots_by='drinking_pred'  # Split plots by drinking prediction
)

# Day-level analysis using dayFeatures
print(f"\nRunning day-level analysis...")
day_features_calculator = dayFeatures(processed_data_folder)
day_features_calculator.compute_low_quality_stats()  # Computes stats and adds to day_stat_frames

# Attach curve slots and score the 08.05 drinking-day model
# (pred_drinking_day, proba_drinking_day, reliability flags).
flag_non_wear_hours = 6
flag_very_negative_hours = 1
flag_artifact_hours = 1
flag_proba_low = 0.05
flag_proba_high = 0.95

day_features_calculator.attach_curves_and_predict_drinking_days(
    curves.curve_features,
    flag_non_wear_hours=flag_non_wear_hours,
    flag_very_negative_hours=flag_very_negative_hours,
    flag_artifact_hours=flag_artifact_hours,
    flag_proba_low=flag_proba_low,
    flag_proba_high=flag_proba_high,
)

# Export day workbook with XGB drinking-day splits
day_features_calculator.export_workbook_days(
    str(project_root / 'Results' / cohort_name / f'{cohort_name}_day_stats_{today}.xlsx'),
    split_plots_by='drinking_xgb',
    include_nonwear_plots=True,
    include_signal_processing_plots=True
)

print(f"\nTest analysis complete!")
print(f"Results exported to: {project_root / 'Results' / cohort_name}")
print(f"- {cohort_name}_curve_stats_{today}.xlsx (curve-level stats with 'Drinking Curves' and 'Non-Drinking Curves' tabs)")
print(f"- {cohort_name}_day_stats_{today}.xlsx (day-level stats with 'Drinking Days (XGB)' and 'Non-Drinking Days (XGB)' tabs)")
print(f"\nDay-level drinking prediction:")
print(f"  - pred_drinking_day / proba_drinking_day from the 08.05 XGBoost model")
print(
    f"  - prediction_reliable is 0 when non-wear (24h minus wear time) is {flag_non_wear_hours:g}+ hours,"
)
print(
    f"    very negative TAC is {flag_very_negative_hours:g}+ hours, artifacts are {flag_artifact_hours:g}+ hours, "
    f"or probability is in [{flag_proba_low:g}, {flag_proba_high:g}]"
)
