# 08.05 drinking-day model

This model labels each social day as drinking or not drinking. The stored file is `xgb_drinking_day_v0805.json`. Load it with XGBoost 2.1.4 (`XGBClassifier.load_model`). Feature order is `xgb_drinking_day_v0805_features.json` (150 columns). MAPTAC scores days through `App/SDM/Machine_Learning/drinking_day_xgb.py`, which `dayFeatures.attach_curves_and_predict_drinking_days` calls by default.

## What the labels are

The training target is `annotation_drinking_day`: a human review of TAC plots (1 = drinking, 0 = drinking not detected). It is not a morning self-report and not a laboratory alcohol assay. Agreement with self-report is not what this model was trained to match. Morning self-report workbooks compare against `pred_drinking_day`. A non-drinking day is unknown in that comparison when `prediction_reliable` is 0.

## Training


| Item             | Value                                                                                                     |
| ---------------- | --------------------------------------------------------------------------------------------------------- |
| Source run       | 08.05.2026 day-curve attachment, SubIDs 5001–5085                                                         |
| Cohort           | 80 participants (SubIDs 5001–5084)                                                                        |
| Labeled days     | 2,361 device-on days                                                                                      |
| Social day       | 06:00 to 06:00 the next calendar day                                                                      |
| Curve slots      | Up to three overlapping curves per day, ranked by high-quality duration (slot 1 = most high-quality time) |
| Outer validation | Leave-one-subject-out on `SubID`                                                                          |
| Inner search     | 5-fold `GroupKFold` on `SubID`, scored by ROC AUC                                                         |
| Final fit        | Those hyperparameters, refit on all labeled days                                                          |


Hyperparameters: `n_estimators=100`, `max_depth=6`, `learning_rate=0.1`, `subsample=0.9`, `colsample_bytree=1.0`, `min_child_weight=1`, `reg_lambda=1.0`.

Leave-one-subject-out on the 2,361 labeled days:


| Metric           | Value                                                                            |
| ---------------- | -------------------------------------------------------------------------------- |
| Accuracy         | 0.956                                                                            |
| Sensitivity      | 0.934                                                                            |
| Specificity      | 0.964                                                                            |
| ROC AUC          | 0.990                                                                            |
| Confusion matrix | 1,653 true negatives, 61 false positives, 43 false negatives, 604 true positives |


Per-subject accuracy ranged from 0.75 to 1.0 (median 0.95). These figures are out-of-fold. The JSON file is the model refit on every labeled day, so scores on the training days are higher than the leave-one-subject-out numbers.

## Features

Inputs are day-level TAC summaries, quality and wear, TAC in the first 6 hours of the social day (`tac_q1_*`) and the remaining 18 hours (`tac_q2_q4_*`), curve-slot statistics for slots 1–3 (including periphery before and after the curve), and medians of qualifying curves and day-level TAC from earlier days in the same burst.

`tac_q1_below_threshold_percent` and `tac_q2_q4_below_threshold_percent` are computed at score time as `1 -` the matching above-threshold column. Missing curve slots stay missing; XGBoost accepts those NaNs. Days with no TAC are still scored.

`day_no` is one of the 150 inputs. The trees split it only at 6, 7, 8, 9, 11, and 14. Any `day_no` of 14 or higher follows the same branch as day 14. A 15th day is scored like day 14, not like a missing `day_no`. The feature is about 0.5% of total gain. The model was trained on 14-day bursts.

## Output columns


| Column                       | Meaning                                                           |
| ---------------------------- | ----------------------------------------------------------------- |
| `pred_drinking_day`          | 1 = drinking, 0 = not drinking                                    |
| `proba_drinking_day`         | Predicted probability of drinking                                 |
| `flag_non_wear`              | 1 when hours not worn are at least 6                              |
| `flag_very_negative_tac`     | 1 when `extreme_negative_duration` (TAC < −15) is at least 1 hour |
| `flag_artifacts`             | 1 when `jump_duration + plummet_duration` is at least 1 hour      |
| `flag_uncertain_probability` | 1 when `proba_drinking_day` is inside [0.05, 0.95]                |
| `prediction_reliable`        | 1 only when all four flags are 0; otherwise 0                     |


Hours not worn are `day_hours - device_worn_duration` (the social day is 24 hours). `device_worn_duration` counts minutes the device was on and classified as worn, so the remainder is powered-off or missing time plus on-device non-wear readings. This flag does not use `non_wear_duration`. If wear duration is missing, wear is treated as 0, so the whole day is not worn and `flag_non_wear` is 1. A powered-off day is still given a drinking prediction (usually 0) and is marked unreliable.

Hour cutoffs use `>=`. Missing very-negative and artifact durations are treated as 0 hours (not flagged). Defaults are arguments of `score_drinking_days` and of `attach_curves_and_predict_drinking_days`: `flag_non_wear_hours`, `flag_very_negative_hours`, `flag_artifact_hours`, `flag_proba_low`, `flag_proba_high`.

## Curve-level labels

`identify_drinking_curves` still writes curve-level `DRINKING_PRED`. That label is not copied onto the day table. Day-level drinking is `pred_drinking_day`. Plot workbooks split days with `split_plots_by='drinking_xgb'`. `add_curve_overlap_detection` is a deprecated alias of `attach_curves_and_predict_drinking_days`.

## Loading the model outside MAPTAC

```python
import json
from xgboost import XGBClassifier

clf = XGBClassifier()
clf.load_model("App/SDM/Trained_Models/xgb_drinking_day_v0805.json")
feature_names = json.load(open("App/SDM/Trained_Models/xgb_drinking_day_v0805_features.json"))
```

Pass a numeric matrix whose columns are `feature_names` in that order. Pin `xgboost==2.1.4`. The model was trained and checked on that 80-participant cohort only.