"""
Score day-level drinking with the 08.05 XGBoost model.

The booster is stored as native XGBoost JSON (not a pickled training-script class).
Feature order matches ``xgb_drinking_day_v0805_features.json``.
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

_SLOT_CURVE_STATS = (
    'duration_CURVE',
    'peak_CURVE',
    'auc_total_CURVE',
    'curve_threshold',
    'rise_rate_CURVE',
    'fall_rate_CURVE',
    'rise_rate_point_to_point_CURVE',
    'fall_rate_point_to_point_CURVE',
    'rise_duration_CURVE',
    'fall_duration_CURVE',
    'below_threshold_percent_CURVE',
    'started_curve_count_CURVE',
)
_SLOT_QUALITY = (
    'device_worn_percent_CURVE',
    'total_low_quality_percent_CURVE',
    'unimputed_low_quality_percent_CURVE',
    'total_gap_percent_CURVE',
    'total_non_wear_percent_CURVE',
    'total_jump_percent_CURVE',
    'total_plummet_percent_CURVE',
    'total_extreme_negative_percent_CURVE',
    'low_quality_imputation_ratio_CURVE',
)
_PERIPHERY_STEMS = (
    'device_worn_percent',
    'total_low_quality_percent',
    'unimputed_low_quality_percent',
    'total_gap_percent',
    'total_non_wear_percent',
    'total_extreme_negative_percent',
    'low_quality_imputation_ratio',
)
_SLOT_PERIPHERY = tuple(
    f'{stem}_{side}'
    for side in ('PERIPHERY_BEFORE', 'PERIPHERY_AFTER')
    for stem in _PERIPHERY_STEMS
)
_SLOT_DERIVED = (
    'start_offset_hours',
    'end_offset_hours',
)
_ALL_SLOT_SUFFIXES = _SLOT_CURVE_STATS + _SLOT_QUALITY + _SLOT_PERIPHERY + _SLOT_DERIVED

FEATURE_COLUMNS = [
    'day_no',
    'new_device',
    'low_quality_percent',
    'device_worn_percent_of_day',
    'above_threshold_high_quality_percent_of_day',
    *[f'tac_{r}_{sfx}' for r in ('q1', 'q2_q4')
      for sfx in ('mean', 'max', 'below_threshold_percent', 'auc',
                  'rise_rate_point_to_point', 'fall_rate_point_to_point',
                  'ascending_duration', 'descending_duration')],
    *[f'curve_{n}_{sfx}' for n in range(1, 4) for sfx in _ALL_SLOT_SUFFIXES],
    'median_prior_peak_CURVE',
    'median_prior_auc_total_CURVE',
    'median_prior_duration_CURVE',
    'median_prior_total_low_quality_percent_CURVE',
    'median_prior_rise_rate_CURVE',
    'median_prior_fall_rate_CURVE',
    'median_prior_rise_rate_point_to_point_CURVE',
    'median_prior_fall_rate_point_to_point_CURVE',
    'median_prior_fall_duration_CURVE',
    'median_prior_rise_duration_CURVE',
    'median_prior_tac_q1_mean',
    'median_prior_tac_q1_max',
    'median_prior_tac_q1_auc',
    'median_prior_tac_q2_q4_mean',
    'median_prior_tac_q2_q4_max',
    'median_prior_tac_q2_q4_auc',
    'median_prior_curve_1_start_offset_hours',
    'median_prior_curve_1_end_offset_hours',
]

# Model-facing below-threshold columns derived at score time as ``1 - above_threshold``.
_BELOW_THRESHOLD_FROM_ABOVE: dict[str, str] = {
    'tac_q1_below_threshold_percent': 'tac_q1_above_threshold_percent',
    'tac_q2_q4_below_threshold_percent': 'tac_q2_q4_above_threshold_percent',
}

_MODEL_DIR = Path(__file__).resolve().parents[1] / 'Trained_Models'
MODEL_JSON = _MODEL_DIR / 'xgb_drinking_day_v0805.json'
FEATURE_NAMES_JSON = _MODEL_DIR / 'xgb_drinking_day_v0805_features.json'

PREDICTION_COLUMNS = (
    'pred_drinking_day',
    'proba_drinking_day',
    'flag_non_wear',
    'flag_very_negative_tac',
    'flag_artifacts',
    'flag_uncertain_probability',
    'prediction_reliable',
)


def load_feature_names(path: Path | None = None) -> list[str]:
    """Feature names in booster order. Must match ``FEATURE_COLUMNS``."""
    path = Path(path) if path is not None else FEATURE_NAMES_JSON
    with path.open() as f:
        names = json.load(f)
    if names != FEATURE_COLUMNS:
        raise ValueError(
            f'Feature list in {path} does not match FEATURE_COLUMNS '
            f'({len(names)} vs {len(FEATURE_COLUMNS)}).'
        )
    if len(names) != 150:
        raise ValueError(f'Expected 150 drinking-day features, found {len(names)}.')
    return list(names)


def load_drinking_day_model(path: Path | None = None):
    """Load the native XGBoost JSON booster as an ``XGBClassifier``."""
    from xgboost import XGBClassifier

    path = Path(path) if path is not None else MODEL_JSON
    feature_names = load_feature_names()
    clf = XGBClassifier()
    clf.load_model(path)
    n_in = getattr(clf, 'n_features_in_', None)
    if n_in is not None and int(n_in) != len(feature_names):
        raise ValueError(
            f'Model at {path} expects {n_in} features; feature file has {len(feature_names)}.'
        )
    booster_names = clf.get_booster().feature_names
    if booster_names is not None and list(booster_names) != feature_names:
        raise ValueError(f'Booster feature names in {path} do not match the feature file.')
    return clf


def build_feature_matrix(
    df: pd.DataFrame,
    feature_cols: list[str] | None = None,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Build ``X`` in model order.

    Below-threshold region features are ``1 - above_threshold``. Columns that are
    absent are filled with NaN (XGBoost accepts missing values). Returns the matrix
    and the names that were not found on ``df`` (derived columns count as found
    when their source column exists).
    """
    feature_cols = list(feature_cols or FEATURE_COLUMNS)
    missing: list[str] = []
    series_by_col: dict[str, pd.Series] = {}

    for col in feature_cols:
        if col in _BELOW_THRESHOLD_FROM_ABOVE:
            src = _BELOW_THRESHOLD_FROM_ABOVE[col]
            if src not in df.columns:
                missing.append(col)
                series_by_col[col] = pd.Series(np.nan, index=df.index, dtype=float)
                continue
            above = pd.to_numeric(df[src], errors='coerce')
            series_by_col[col] = 1.0 - above
            continue
        if col not in df.columns:
            missing.append(col)
            series_by_col[col] = pd.Series(np.nan, index=df.index, dtype=float)
            continue
        series_by_col[col] = pd.to_numeric(df[col], errors='coerce')

    X = pd.DataFrame({c: series_by_col[c] for c in feature_cols}, index=df.index)
    return X, missing


def _positive_class_proba(clf, proba: np.ndarray) -> np.ndarray:
    classes = getattr(clf, 'classes_', None)
    if classes is not None:
        classes = [int(c) for c in list(classes)]
        if 1 in classes:
            return proba[:, classes.index(1)]
    if proba.shape[1] > 1:
        return proba[:, 1]
    return proba[:, 0]


def hours_not_worn(df: pd.DataFrame, day_hours: float = 24.0) -> pd.Series:
    """
    Hours in the social day that were not worn.

    ``device_worn_duration`` counts minutes the device was on and classified as worn.
    The remainder of the social day is powered-off or missing time plus on-device
    non-wear readings. Missing wear duration is treated as 0 hours worn.
    """
    if 'day_hours' in df.columns:
        day_h = pd.to_numeric(df['day_hours'], errors='coerce').fillna(day_hours)
    else:
        day_h = pd.Series(float(day_hours), index=df.index)
    if 'device_worn_duration' in df.columns:
        worn = pd.to_numeric(df['device_worn_duration'], errors='coerce').fillna(0.0)
    else:
        worn = pd.Series(0.0, index=df.index)
    return day_h - worn


def add_prediction_reliability_flags(
    df: pd.DataFrame,
    *,
    flag_non_wear_hours: float = 6.0,
    flag_very_negative_hours: float = 1.0,
    flag_artifact_hours: float = 1.0,
    flag_proba_low: float = 0.05,
    flag_proba_high: float = 0.95,
    proba_col: str = 'proba_drinking_day',
    day_hours: float = 24.0,
) -> pd.DataFrame:
    """
    Flag unreliable day-level drinking predictions.

    ``flag_non_wear`` is 1 when hours not worn (``day_hours - device_worn_duration``)
    are at least ``flag_non_wear_hours``. That remainder includes powered-off or missing
    time and on-device non-wear readings. It does not use ``non_wear_duration``.

    Very-negative and artifact flags use ``>=`` on ``extreme_negative_duration``
    (TAC < -15) and ``jump_duration + plummet_duration``. Missing values on those
    duration columns are treated as 0 (not flagged).

    ``flag_uncertain_probability`` is 1 when ``proba_col`` is inside
    ``[flag_proba_low, flag_proba_high]``. ``prediction_reliable`` is 1 only when every flag is 0.
    """
    not_worn = hours_not_worn(df, day_hours=day_hours)
    df['flag_non_wear'] = (not_worn >= float(flag_non_wear_hours)).astype(int)

    if 'extreme_negative_duration' in df.columns:
        very_neg = pd.to_numeric(df['extreme_negative_duration'], errors='coerce').fillna(0.0)
    else:
        very_neg = pd.Series(0.0, index=df.index)
    df['flag_very_negative_tac'] = (very_neg >= float(flag_very_negative_hours)).astype(int)

    jump = (
        pd.to_numeric(df['jump_duration'], errors='coerce').fillna(0.0)
        if 'jump_duration' in df.columns
        else pd.Series(0.0, index=df.index)
    )
    plummet = (
        pd.to_numeric(df['plummet_duration'], errors='coerce').fillna(0.0)
        if 'plummet_duration' in df.columns
        else pd.Series(0.0, index=df.index)
    )
    df['flag_artifacts'] = ((jump + plummet) >= float(flag_artifact_hours)).astype(int)

    if proba_col not in df.columns:
        uncertain = pd.Series(0, index=df.index)
    else:
        proba = pd.to_numeric(df[proba_col], errors='coerce')
        uncertain = proba.between(float(flag_proba_low), float(flag_proba_high), inclusive='both').fillna(False)
    df['flag_uncertain_probability'] = uncertain.astype(int)

    flagged = (
        (df['flag_non_wear'] == 1)
        | (df['flag_very_negative_tac'] == 1)
        | (df['flag_artifacts'] == 1)
        | (df['flag_uncertain_probability'] == 1)
    )
    df['prediction_reliable'] = (~flagged).astype(int)
    return df


def score_drinking_days(
    df: pd.DataFrame,
    *,
    model=None,
    flag_non_wear_hours: float = 6.0,
    flag_very_negative_hours: float = 1.0,
    flag_artifact_hours: float = 1.0,
    flag_proba_low: float = 0.05,
    flag_proba_high: float = 0.95,
    day_hours: float = 24.0,
) -> tuple[pd.DataFrame, list[str]]:
    """
    Add ``pred_drinking_day``, ``proba_drinking_day``, and reliability flags.

    Days with no TAC are still scored. Missing curve-slot values stay NaN.
    Returns the same dataframe and the list of model features absent from ``df``.
    """
    feature_names = load_feature_names()
    X, missing = build_feature_matrix(df, feature_names)
    clf = model if model is not None else load_drinking_day_model()
    proba = clf.predict_proba(X)
    df['pred_drinking_day'] = clf.predict(X).astype(int)
    df['proba_drinking_day'] = _positive_class_proba(clf, np.asarray(proba))
    add_prediction_reliability_flags(
        df,
        flag_non_wear_hours=flag_non_wear_hours,
        flag_very_negative_hours=flag_very_negative_hours,
        flag_artifact_hours=flag_artifact_hours,
        flag_proba_low=flag_proba_low,
        flag_proba_high=flag_proba_high,
        day_hours=day_hours,
    )
    return df, missing
