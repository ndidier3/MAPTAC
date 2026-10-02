"""
Default curve analysis settings.
These settings are used by cohort analysis scripts.
"""

from App.SDM.Run.default_settings.default_flag_settings import DEFAULT_FLAG_SELECTIONS

DEFAULT_CURVE_ATTRS = {
    'flag_selections': DEFAULT_FLAG_SELECTIONS,
    'curve_threshold': 'auto',
    'default_threshold': 8.0,
    'ensure_default_curve_threshold_applied': False,
    'periphery_buffer_before': 2,
    'periphery_buffer_after': 2,
    'merge_curves_within_duration': 2,
}

DEFAULT_RAW_CURVE_ATTRS = {
    'raw_curve_demarcation_mode': 'independent',  # 'independent' | 'imputed_windows' | 'adjust'
    'raw_tac_column': 'TAC_smoothed_unimputed',
    'curve_threshold': 'auto',
    'default_threshold': 8.0,
    'ensure_default_curve_threshold_applied': False,
    'periphery_buffer_before': 2,
    'periphery_buffer_after': 2,
    'merge_curves_within_duration': 2,
}
